Imports DLL_DataManager
Imports DLL_Export_Service
Imports Microsoft.DirectX.DirectInput
Imports DLL_SetupAndCalibrate
Imports System.Threading
Imports Microsoft.win32
Imports Microsoft.Office.Interop
Imports System.Messaging
Imports Microsoft.VisualBasic.ComClassAttribute
Imports System.Math
Imports System.IO
Imports DLL_Export_Device_Vision

Public Class FormProgramming
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        'InitialPatternParameterSetup()
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
    Friend WithEvents MenuItem1 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem6 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem8 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem11 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem15 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem17 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem19 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem20 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem25 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem29 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem30 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem32 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem33 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem35 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem36 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem62 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem64 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem65 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem67 As System.Windows.Forms.MenuItem
    Friend WithEvents DomainUpDown1 As System.Windows.Forms.DomainUpDown
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents ImageListFiles As System.Windows.Forms.ImageList
    Friend WithEvents imageListElement As System.Windows.Forms.ImageList
    Friend WithEvents ImageListReference As System.Windows.Forms.ImageList
    Friend WithEvents TBFiducialPt As System.Windows.Forms.ToolBarButton
    Friend WithEvents TBHeightPt As System.Windows.Forms.ToolBarButton
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents ImageListYesNo As System.Windows.Forms.ImageList
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents MenuItem76 As System.Windows.Forms.MenuItem
    Friend WithEvents CBExpandSpreadsheet As System.Windows.Forms.CheckBox
    '   Friend WithEvents AxSpreadsheet1 As AxOWC10.AxSpreadsheet
    Friend WithEvents ImageListGeneralTools As System.Windows.Forms.ImageList
    Friend WithEvents ButtonHome As System.Windows.Forms.Button
    Friend WithEvents ButtonVolCal As System.Windows.Forms.Button
    Friend WithEvents ButtonNeedleCal As System.Windows.Forms.Button
    Friend WithEvents ButtonPurge As System.Windows.Forms.Button
    Friend WithEvents ButtonClean As System.Windows.Forms.Button
    Friend WithEvents ImageListMultiField As System.Windows.Forms.ImageList
    Friend WithEvents CBDoorLock As System.Windows.Forms.CheckBox
    Friend WithEvents PanelVision As System.Windows.Forms.Panel
    Friend WithEvents PanelVisionCtrl As System.Windows.Forms.Panel
    Friend WithEvents NeedleContextMenu As System.Windows.Forms.ContextMenu
    Friend WithEvents MenuItem81 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem82 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem83 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem84 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem85 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem86 As System.Windows.Forms.MenuItem
    Friend WithEvents OpenPatternFileDialog As System.Windows.Forms.OpenFileDialog
    Friend WithEvents SavePatternFileDialog As System.Windows.Forms.SaveFileDialog
    Friend WithEvents TextBoxRobotZ As System.Windows.Forms.TextBox
    Friend WithEvents TextBoxRobotY As System.Windows.Forms.TextBox
    Friend WithEvents TextBoxRobotX As System.Windows.Forms.TextBox
    Friend WithEvents ImageListLineArc As System.Windows.Forms.ImageList
    Friend WithEvents MenuItem87 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem88 As System.Windows.Forms.MenuItem
    Friend WithEvents ImageListOper As System.Windows.Forms.ImageList
    Friend WithEvents MainMenuProgramming As System.Windows.Forms.MainMenu
    Friend WithEvents MenuFileNew As System.Windows.Forms.MenuItem
    Friend WithEvents MenuFileOpen As System.Windows.Forms.MenuItem
    Friend WithEvents MenuFileImport As System.Windows.Forms.MenuItem
    Friend WithEvents MenuFileExport As System.Windows.Forms.MenuItem
    Friend WithEvents TBBDot As System.Windows.Forms.ToolBarButton
    Friend WithEvents TBBLine As System.Windows.Forms.ToolBarButton
    Friend WithEvents TBBArc As System.Windows.Forms.ToolBarButton
    Friend WithEvents TBBRectangle As System.Windows.Forms.ToolBarButton
    Friend WithEvents TBBCircle As System.Windows.Forms.ToolBarButton
    Friend WithEvents TBBFilledRectangle As System.Windows.Forms.ToolBarButton
    Friend WithEvents TBBFilledCircle As System.Windows.Forms.ToolBarButton
    Friend WithEvents TBBLink As System.Windows.Forms.ToolBarButton
    Friend WithEvents TBBChipEdge As System.Windows.Forms.ToolBarButton
    Friend WithEvents TBBMove As System.Windows.Forms.ToolBarButton
    Friend WithEvents TBBWait As System.Windows.Forms.ToolBarButton
    Friend WithEvents TBBPurge As System.Windows.Forms.ToolBarButton
    Friend WithEvents TBBClean As System.Windows.Forms.ToolBarButton
    Friend WithEvents TBBQC As System.Windows.Forms.ToolBarButton
    Friend WithEvents TBBSubPattern As System.Windows.Forms.ToolBarButton
    Friend WithEvents TBBArray As System.Windows.Forms.ToolBarButton
    Friend WithEvents TBBGetIO As System.Windows.Forms.ToolBarButton
    Friend WithEvents TBBSetIO As System.Windows.Forms.ToolBarButton
    Friend WithEvents TBBMeasure As System.Windows.Forms.ToolBarButton
    Friend WithEvents MenuFileSave As System.Windows.Forms.MenuItem
    Friend WithEvents MenuFileSaveAs As System.Windows.Forms.MenuItem
    Friend WithEvents MenuFileExit As System.Windows.Forms.MenuItem
    Friend WithEvents TimerForUpdate As System.Windows.Forms.Timer
    Friend WithEvents CheckBoxLockZ As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBoxLockY As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBoxLockX As System.Windows.Forms.CheckBox
    Friend WithEvents MenuEditUndo As System.Windows.Forms.MenuItem
    Friend WithEvents MenuEditRedo As System.Windows.Forms.MenuItem
    Friend WithEvents MenuEditCut As System.Windows.Forms.MenuItem
    Friend WithEvents MenuEditCopy As System.Windows.Forms.MenuItem
    Friend WithEvents MenuEditPaste As System.Windows.Forms.MenuItem
    Friend WithEvents MenuEditDelete As System.Windows.Forms.MenuItem
    Friend WithEvents MenuEditSelectAll As System.Windows.Forms.MenuItem
    Friend WithEvents OptimizePath As System.Windows.Forms.MenuItem
    Friend WithEvents SaveFileDialog1 As System.Windows.Forms.SaveFileDialog
    Friend WithEvents MenuItem3 As System.Windows.Forms.MenuItem
    Friend WithEvents AxSpreadsheetProgramming As AxOWC10.AxSpreadsheet
    Public WithEvents IOCheck As System.Windows.Forms.Timer
    Friend WithEvents LabelMessege As System.Windows.Forms.Label
    Friend WithEvents TBBOffset As System.Windows.Forms.ToolBarButton
    Friend WithEvents TBBDotArray As System.Windows.Forms.ToolBarButton
    Friend WithEvents NeedleMode As System.Windows.Forms.RadioButton
    Friend WithEvents ButtonToggleMode As System.Windows.Forms.Button
    Public WithEvents PanelToBeAdded As System.Windows.Forms.Panel
    Friend WithEvents VisionMode As System.Windows.Forms.RadioButton
    Friend WithEvents ReferenceCommandBlock As System.Windows.Forms.ToolBar
    Friend WithEvents ElementsCommandBlock As System.Windows.Forms.ToolBar
    Friend WithEvents DispensingMode As System.Windows.Forms.ComboBox
    Friend WithEvents TBBVolumeCal As System.Windows.Forms.ToolBarButton
    Friend WithEvents ValueBrightness As System.Windows.Forms.NumericUpDown
    Friend WithEvents ButtonStartFirstStage As System.Windows.Forms.Button
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents RedLight As System.Windows.Forms.PictureBox
    Friend WithEvents AmberLight As System.Windows.Forms.PictureBox
    Friend WithEvents GreenLight As System.Windows.Forms.PictureBox
    Friend WithEvents TowerLightImageList As System.Windows.Forms.ImageList
    Friend WithEvents btPlay As System.Windows.Forms.Button
    Friend WithEvents btPause As System.Windows.Forms.Button
    Friend WithEvents btStop As System.Windows.Forms.Button
    Friend WithEvents btRelease As System.Windows.Forms.Button
    Friend WithEvents btRetrieve As System.Windows.Forms.Button
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents tbOpenedFile As System.Windows.Forms.TextBox
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents cbContinueTest As System.Windows.Forms.CheckBox
    Friend WithEvents btChangeSyringe As System.Windows.Forms.Button
    Friend WithEvents gbConveyor As System.Windows.Forms.GroupBox
    Friend WithEvents btExit As System.Windows.Forms.Button
    Friend WithEvents lbPostName As System.Windows.Forms.Label
    Friend WithEvents cbDisplayIndicator As System.Windows.Forms.CheckBox
    Friend WithEvents btResetVolCalSettings As System.Windows.Forms.Button
    Friend WithEvents tbLsHeight As System.Windows.Forms.TextBox
    Friend WithEvents btEject As System.Windows.Forms.Button
    Friend WithEvents btReset As System.Windows.Forms.Button
    Friend WithEvents btConfirm As System.Windows.Forms.Button
    Friend WithEvents btCancel As System.Windows.Forms.Button
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents Panel3 As System.Windows.Forms.Panel
    Friend WithEvents btStep As System.Windows.Forms.Button
    Friend WithEvents btEdit As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(FormProgramming))
        Me.MainMenuProgramming = New System.Windows.Forms.MainMenu
        Me.MenuItem1 = New System.Windows.Forms.MenuItem
        Me.MenuFileNew = New System.Windows.Forms.MenuItem
        Me.MenuFileOpen = New System.Windows.Forms.MenuItem
        Me.MenuItem19 = New System.Windows.Forms.MenuItem
        Me.MenuFileImport = New System.Windows.Forms.MenuItem
        Me.MenuFileExport = New System.Windows.Forms.MenuItem
        Me.MenuItem20 = New System.Windows.Forms.MenuItem
        Me.MenuFileSave = New System.Windows.Forms.MenuItem
        Me.MenuFileSaveAs = New System.Windows.Forms.MenuItem
        Me.MenuItem6 = New System.Windows.Forms.MenuItem
        Me.MenuFileExit = New System.Windows.Forms.MenuItem
        Me.MenuItem8 = New System.Windows.Forms.MenuItem
        Me.MenuEditUndo = New System.Windows.Forms.MenuItem
        Me.MenuEditRedo = New System.Windows.Forms.MenuItem
        Me.MenuItem15 = New System.Windows.Forms.MenuItem
        Me.MenuEditCut = New System.Windows.Forms.MenuItem
        Me.MenuEditCopy = New System.Windows.Forms.MenuItem
        Me.MenuEditPaste = New System.Windows.Forms.MenuItem
        Me.MenuEditDelete = New System.Windows.Forms.MenuItem
        Me.MenuItem17 = New System.Windows.Forms.MenuItem
        Me.MenuEditSelectAll = New System.Windows.Forms.MenuItem
        Me.MenuItem11 = New System.Windows.Forms.MenuItem
        Me.MenuItem67 = New System.Windows.Forms.MenuItem
        Me.MenuItem87 = New System.Windows.Forms.MenuItem
        Me.MenuItem88 = New System.Windows.Forms.MenuItem
        Me.MenuItem25 = New System.Windows.Forms.MenuItem
        Me.MenuItem33 = New System.Windows.Forms.MenuItem
        Me.MenuItem32 = New System.Windows.Forms.MenuItem
        Me.MenuItem35 = New System.Windows.Forms.MenuItem
        Me.MenuItem29 = New System.Windows.Forms.MenuItem
        Me.MenuItem30 = New System.Windows.Forms.MenuItem
        Me.MenuItem36 = New System.Windows.Forms.MenuItem
        Me.MenuItem62 = New System.Windows.Forms.MenuItem
        Me.MenuItem76 = New System.Windows.Forms.MenuItem
        Me.OptimizePath = New System.Windows.Forms.MenuItem
        Me.MenuItem3 = New System.Windows.Forms.MenuItem
        Me.MenuItem64 = New System.Windows.Forms.MenuItem
        Me.MenuItem65 = New System.Windows.Forms.MenuItem
        Me.PanelVision = New System.Windows.Forms.Panel
        Me.ImageListGeneralTools = New System.Windows.Forms.ImageList(Me.components)
        Me.ButtonPurge = New System.Windows.Forms.Button
        Me.PanelVisionCtrl = New System.Windows.Forms.Panel
        Me.cbDisplayIndicator = New System.Windows.Forms.CheckBox
        Me.ValueBrightness = New System.Windows.Forms.NumericUpDown
        Me.CheckBoxLockZ = New System.Windows.Forms.CheckBox
        Me.TextBoxRobotZ = New System.Windows.Forms.TextBox
        Me.CheckBoxLockY = New System.Windows.Forms.CheckBox
        Me.TextBoxRobotY = New System.Windows.Forms.TextBox
        Me.CheckBoxLockX = New System.Windows.Forms.CheckBox
        Me.TextBoxRobotX = New System.Windows.Forms.TextBox
        Me.lbPostName = New System.Windows.Forms.Label
        Me.DomainUpDown1 = New System.Windows.Forms.DomainUpDown
        Me.Label1 = New System.Windows.Forms.Label
        Me.ButtonVolCal = New System.Windows.Forms.Button
        Me.ImageListFiles = New System.Windows.Forms.ImageList(Me.components)
        Me.imageListElement = New System.Windows.Forms.ImageList(Me.components)
        Me.ImageListReference = New System.Windows.Forms.ImageList(Me.components)
        Me.ReferenceCommandBlock = New System.Windows.Forms.ToolBar
        Me.TBFiducialPt = New System.Windows.Forms.ToolBarButton
        Me.TBHeightPt = New System.Windows.Forms.ToolBarButton
        Me.ElementsCommandBlock = New System.Windows.Forms.ToolBar
        Me.TBBDot = New System.Windows.Forms.ToolBarButton
        Me.TBBLine = New System.Windows.Forms.ToolBarButton
        Me.TBBArc = New System.Windows.Forms.ToolBarButton
        Me.TBBRectangle = New System.Windows.Forms.ToolBarButton
        Me.TBBCircle = New System.Windows.Forms.ToolBarButton
        Me.TBBFilledRectangle = New System.Windows.Forms.ToolBarButton
        Me.TBBFilledCircle = New System.Windows.Forms.ToolBarButton
        Me.TBBLink = New System.Windows.Forms.ToolBarButton
        Me.TBBChipEdge = New System.Windows.Forms.ToolBarButton
        Me.TBBMove = New System.Windows.Forms.ToolBarButton
        Me.TBBWait = New System.Windows.Forms.ToolBarButton
        Me.TBBPurge = New System.Windows.Forms.ToolBarButton
        Me.TBBClean = New System.Windows.Forms.ToolBarButton
        Me.TBBQC = New System.Windows.Forms.ToolBarButton
        Me.TBBSubPattern = New System.Windows.Forms.ToolBarButton
        Me.TBBArray = New System.Windows.Forms.ToolBarButton
        Me.TBBGetIO = New System.Windows.Forms.ToolBarButton
        Me.TBBSetIO = New System.Windows.Forms.ToolBarButton
        Me.TBBOffset = New System.Windows.Forms.ToolBarButton
        Me.TBBMeasure = New System.Windows.Forms.ToolBarButton
        Me.TBBDotArray = New System.Windows.Forms.ToolBarButton
        Me.TBBVolumeCal = New System.Windows.Forms.ToolBarButton
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.Label5 = New System.Windows.Forms.Label
        Me.VisionMode = New System.Windows.Forms.RadioButton
        Me.CBExpandSpreadsheet = New System.Windows.Forms.CheckBox
        Me.NeedleMode = New System.Windows.Forms.RadioButton
        Me.ButtonClean = New System.Windows.Forms.Button
        Me.ButtonNeedleCal = New System.Windows.Forms.Button
        Me.ButtonHome = New System.Windows.Forms.Button
        Me.CBDoorLock = New System.Windows.Forms.CheckBox
        Me.ImageListMultiField = New System.Windows.Forms.ImageList(Me.components)
        Me.AxSpreadsheetProgramming = New AxOWC10.AxSpreadsheet
        Me.LabelMessege = New System.Windows.Forms.Label
        Me.DispensingMode = New System.Windows.Forms.ComboBox
        Me.btStop = New System.Windows.Forms.Button
        Me.btPause = New System.Windows.Forms.Button
        Me.btPlay = New System.Windows.Forms.Button
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.GreenLight = New System.Windows.Forms.PictureBox
        Me.AmberLight = New System.Windows.Forms.PictureBox
        Me.RedLight = New System.Windows.Forms.PictureBox
        Me.ButtonToggleMode = New System.Windows.Forms.Button
        Me.PanelToBeAdded = New System.Windows.Forms.Panel
        Me.ButtonStartFirstStage = New System.Windows.Forms.Button
        Me.btRelease = New System.Windows.Forms.Button
        Me.btRetrieve = New System.Windows.Forms.Button
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.cbContinueTest = New System.Windows.Forms.CheckBox
        Me.gbConveyor = New System.Windows.Forms.GroupBox
        Me.btReset = New System.Windows.Forms.Button
        Me.btEject = New System.Windows.Forms.Button
        Me.tbOpenedFile = New System.Windows.Forms.TextBox
        Me.GroupBox3 = New System.Windows.Forms.GroupBox
        Me.btChangeSyringe = New System.Windows.Forms.Button
        Me.btExit = New System.Windows.Forms.Button
        Me.btResetVolCalSettings = New System.Windows.Forms.Button
        Me.tbLsHeight = New System.Windows.Forms.TextBox
        Me.btConfirm = New System.Windows.Forms.Button
        Me.btCancel = New System.Windows.Forms.Button
        Me.Panel2 = New System.Windows.Forms.Panel
        Me.Panel3 = New System.Windows.Forms.Panel
        Me.btStep = New System.Windows.Forms.Button
        Me.btEdit = New System.Windows.Forms.Button
        Me.ImageListYesNo = New System.Windows.Forms.ImageList(Me.components)
        Me.ImageListOper = New System.Windows.Forms.ImageList(Me.components)
        Me.NeedleContextMenu = New System.Windows.Forms.ContextMenu
        Me.MenuItem81 = New System.Windows.Forms.MenuItem
        Me.MenuItem82 = New System.Windows.Forms.MenuItem
        Me.MenuItem83 = New System.Windows.Forms.MenuItem
        Me.MenuItem84 = New System.Windows.Forms.MenuItem
        Me.MenuItem85 = New System.Windows.Forms.MenuItem
        Me.MenuItem86 = New System.Windows.Forms.MenuItem
        Me.OpenPatternFileDialog = New System.Windows.Forms.OpenFileDialog
        Me.SavePatternFileDialog = New System.Windows.Forms.SaveFileDialog
        Me.ImageListLineArc = New System.Windows.Forms.ImageList(Me.components)
        Me.TimerForUpdate = New System.Windows.Forms.Timer(Me.components)
        Me.SaveFileDialog1 = New System.Windows.Forms.SaveFileDialog
        Me.IOCheck = New System.Windows.Forms.Timer(Me.components)
        Me.TowerLightImageList = New System.Windows.Forms.ImageList(Me.components)
        Me.PanelVisionCtrl.SuspendLayout()
        CType(Me.ValueBrightness, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.AxSpreadsheetProgramming, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel1.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.gbConveyor.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.Panel3.SuspendLayout()
        Me.SuspendLayout()
        '
        'MainMenuProgramming
        '
        Me.MainMenuProgramming.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem1, Me.MenuItem8, Me.MenuItem11, Me.MenuItem25, Me.MenuItem64})
        Me.MainMenuProgramming.RightToLeft = CType(resources.GetObject("MainMenuProgramming.RightToLeft"), System.Windows.Forms.RightToLeft)
        '
        'MenuItem1
        '
        Me.MenuItem1.Enabled = CType(resources.GetObject("MenuItem1.Enabled"), Boolean)
        Me.MenuItem1.Index = 0
        Me.MenuItem1.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuFileNew, Me.MenuFileOpen, Me.MenuItem19, Me.MenuFileImport, Me.MenuFileExport, Me.MenuItem20, Me.MenuFileSave, Me.MenuFileSaveAs, Me.MenuItem6, Me.MenuFileExit})
        Me.MenuItem1.Shortcut = CType(resources.GetObject("MenuItem1.Shortcut"), System.Windows.Forms.Shortcut)
        Me.MenuItem1.ShowShortcut = CType(resources.GetObject("MenuItem1.ShowShortcut"), Boolean)
        Me.MenuItem1.Text = resources.GetString("MenuItem1.Text")
        Me.MenuItem1.Visible = CType(resources.GetObject("MenuItem1.Visible"), Boolean)
        '
        'MenuFileNew
        '
        Me.MenuFileNew.Enabled = CType(resources.GetObject("MenuFileNew.Enabled"), Boolean)
        Me.MenuFileNew.Index = 0
        Me.MenuFileNew.Shortcut = CType(resources.GetObject("MenuFileNew.Shortcut"), System.Windows.Forms.Shortcut)
        Me.MenuFileNew.ShowShortcut = CType(resources.GetObject("MenuFileNew.ShowShortcut"), Boolean)
        Me.MenuFileNew.Text = resources.GetString("MenuFileNew.Text")
        Me.MenuFileNew.Visible = CType(resources.GetObject("MenuFileNew.Visible"), Boolean)
        '
        'MenuFileOpen
        '
        Me.MenuFileOpen.Enabled = CType(resources.GetObject("MenuFileOpen.Enabled"), Boolean)
        Me.MenuFileOpen.Index = 1
        Me.MenuFileOpen.Shortcut = CType(resources.GetObject("MenuFileOpen.Shortcut"), System.Windows.Forms.Shortcut)
        Me.MenuFileOpen.ShowShortcut = CType(resources.GetObject("MenuFileOpen.ShowShortcut"), Boolean)
        Me.MenuFileOpen.Text = resources.GetString("MenuFileOpen.Text")
        Me.MenuFileOpen.Visible = CType(resources.GetObject("MenuFileOpen.Visible"), Boolean)
        '
        'MenuItem19
        '
        Me.MenuItem19.Enabled = CType(resources.GetObject("MenuItem19.Enabled"), Boolean)
        Me.MenuItem19.Index = 2
        Me.MenuItem19.Shortcut = CType(resources.GetObject("MenuItem19.Shortcut"), System.Windows.Forms.Shortcut)
        Me.MenuItem19.ShowShortcut = CType(resources.GetObject("MenuItem19.ShowShortcut"), Boolean)
        Me.MenuItem19.Text = resources.GetString("MenuItem19.Text")
        Me.MenuItem19.Visible = CType(resources.GetObject("MenuItem19.Visible"), Boolean)
        '
        'MenuFileImport
        '
        Me.MenuFileImport.Enabled = CType(resources.GetObject("MenuFileImport.Enabled"), Boolean)
        Me.MenuFileImport.Index = 3
        Me.MenuFileImport.Shortcut = CType(resources.GetObject("MenuFileImport.Shortcut"), System.Windows.Forms.Shortcut)
        Me.MenuFileImport.ShowShortcut = CType(resources.GetObject("MenuFileImport.ShowShortcut"), Boolean)
        Me.MenuFileImport.Text = resources.GetString("MenuFileImport.Text")
        Me.MenuFileImport.Visible = CType(resources.GetObject("MenuFileImport.Visible"), Boolean)
        '
        'MenuFileExport
        '
        Me.MenuFileExport.Enabled = CType(resources.GetObject("MenuFileExport.Enabled"), Boolean)
        Me.MenuFileExport.Index = 4
        Me.MenuFileExport.Shortcut = CType(resources.GetObject("MenuFileExport.Shortcut"), System.Windows.Forms.Shortcut)
        Me.MenuFileExport.ShowShortcut = CType(resources.GetObject("MenuFileExport.ShowShortcut"), Boolean)
        Me.MenuFileExport.Text = resources.GetString("MenuFileExport.Text")
        Me.MenuFileExport.Visible = CType(resources.GetObject("MenuFileExport.Visible"), Boolean)
        '
        'MenuItem20
        '
        Me.MenuItem20.Enabled = CType(resources.GetObject("MenuItem20.Enabled"), Boolean)
        Me.MenuItem20.Index = 5
        Me.MenuItem20.Shortcut = CType(resources.GetObject("MenuItem20.Shortcut"), System.Windows.Forms.Shortcut)
        Me.MenuItem20.ShowShortcut = CType(resources.GetObject("MenuItem20.ShowShortcut"), Boolean)
        Me.MenuItem20.Text = resources.GetString("MenuItem20.Text")
        Me.MenuItem20.Visible = CType(resources.GetObject("MenuItem20.Visible"), Boolean)
        '
        'MenuFileSave
        '
        Me.MenuFileSave.Enabled = CType(resources.GetObject("MenuFileSave.Enabled"), Boolean)
        Me.MenuFileSave.Index = 6
        Me.MenuFileSave.Shortcut = CType(resources.GetObject("MenuFileSave.Shortcut"), System.Windows.Forms.Shortcut)
        Me.MenuFileSave.ShowShortcut = CType(resources.GetObject("MenuFileSave.ShowShortcut"), Boolean)
        Me.MenuFileSave.Text = resources.GetString("MenuFileSave.Text")
        Me.MenuFileSave.Visible = CType(resources.GetObject("MenuFileSave.Visible"), Boolean)
        '
        'MenuFileSaveAs
        '
        Me.MenuFileSaveAs.Enabled = CType(resources.GetObject("MenuFileSaveAs.Enabled"), Boolean)
        Me.MenuFileSaveAs.Index = 7
        Me.MenuFileSaveAs.Shortcut = CType(resources.GetObject("MenuFileSaveAs.Shortcut"), System.Windows.Forms.Shortcut)
        Me.MenuFileSaveAs.ShowShortcut = CType(resources.GetObject("MenuFileSaveAs.ShowShortcut"), Boolean)
        Me.MenuFileSaveAs.Text = resources.GetString("MenuFileSaveAs.Text")
        Me.MenuFileSaveAs.Visible = CType(resources.GetObject("MenuFileSaveAs.Visible"), Boolean)
        '
        'MenuItem6
        '
        Me.MenuItem6.Enabled = CType(resources.GetObject("MenuItem6.Enabled"), Boolean)
        Me.MenuItem6.Index = 8
        Me.MenuItem6.Shortcut = CType(resources.GetObject("MenuItem6.Shortcut"), System.Windows.Forms.Shortcut)
        Me.MenuItem6.ShowShortcut = CType(resources.GetObject("MenuItem6.ShowShortcut"), Boolean)
        Me.MenuItem6.Text = resources.GetString("MenuItem6.Text")
        Me.MenuItem6.Visible = CType(resources.GetObject("MenuItem6.Visible"), Boolean)
        '
        'MenuFileExit
        '
        Me.MenuFileExit.Enabled = CType(resources.GetObject("MenuFileExit.Enabled"), Boolean)
        Me.MenuFileExit.Index = 9
        Me.MenuFileExit.Shortcut = CType(resources.GetObject("MenuFileExit.Shortcut"), System.Windows.Forms.Shortcut)
        Me.MenuFileExit.ShowShortcut = CType(resources.GetObject("MenuFileExit.ShowShortcut"), Boolean)
        Me.MenuFileExit.Text = resources.GetString("MenuFileExit.Text")
        Me.MenuFileExit.Visible = CType(resources.GetObject("MenuFileExit.Visible"), Boolean)
        '
        'MenuItem8
        '
        Me.MenuItem8.Enabled = CType(resources.GetObject("MenuItem8.Enabled"), Boolean)
        Me.MenuItem8.Index = 1
        Me.MenuItem8.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuEditUndo, Me.MenuEditRedo, Me.MenuItem15, Me.MenuEditCut, Me.MenuEditCopy, Me.MenuEditPaste, Me.MenuEditDelete, Me.MenuItem17, Me.MenuEditSelectAll})
        Me.MenuItem8.Shortcut = CType(resources.GetObject("MenuItem8.Shortcut"), System.Windows.Forms.Shortcut)
        Me.MenuItem8.ShowShortcut = CType(resources.GetObject("MenuItem8.ShowShortcut"), Boolean)
        Me.MenuItem8.Text = resources.GetString("MenuItem8.Text")
        Me.MenuItem8.Visible = CType(resources.GetObject("MenuItem8.Visible"), Boolean)
        '
        'MenuEditUndo
        '
        Me.MenuEditUndo.Enabled = CType(resources.GetObject("MenuEditUndo.Enabled"), Boolean)
        Me.MenuEditUndo.Index = 0
        Me.MenuEditUndo.Shortcut = CType(resources.GetObject("MenuEditUndo.Shortcut"), System.Windows.Forms.Shortcut)
        Me.MenuEditUndo.ShowShortcut = CType(resources.GetObject("MenuEditUndo.ShowShortcut"), Boolean)
        Me.MenuEditUndo.Text = resources.GetString("MenuEditUndo.Text")
        Me.MenuEditUndo.Visible = CType(resources.GetObject("MenuEditUndo.Visible"), Boolean)
        '
        'MenuEditRedo
        '
        Me.MenuEditRedo.Enabled = CType(resources.GetObject("MenuEditRedo.Enabled"), Boolean)
        Me.MenuEditRedo.Index = 1
        Me.MenuEditRedo.Shortcut = CType(resources.GetObject("MenuEditRedo.Shortcut"), System.Windows.Forms.Shortcut)
        Me.MenuEditRedo.ShowShortcut = CType(resources.GetObject("MenuEditRedo.ShowShortcut"), Boolean)
        Me.MenuEditRedo.Text = resources.GetString("MenuEditRedo.Text")
        Me.MenuEditRedo.Visible = CType(resources.GetObject("MenuEditRedo.Visible"), Boolean)
        '
        'MenuItem15
        '
        Me.MenuItem15.Enabled = CType(resources.GetObject("MenuItem15.Enabled"), Boolean)
        Me.MenuItem15.Index = 2
        Me.MenuItem15.Shortcut = CType(resources.GetObject("MenuItem15.Shortcut"), System.Windows.Forms.Shortcut)
        Me.MenuItem15.ShowShortcut = CType(resources.GetObject("MenuItem15.ShowShortcut"), Boolean)
        Me.MenuItem15.Text = resources.GetString("MenuItem15.Text")
        Me.MenuItem15.Visible = CType(resources.GetObject("MenuItem15.Visible"), Boolean)
        '
        'MenuEditCut
        '
        Me.MenuEditCut.Enabled = CType(resources.GetObject("MenuEditCut.Enabled"), Boolean)
        Me.MenuEditCut.Index = 3
        Me.MenuEditCut.Shortcut = CType(resources.GetObject("MenuEditCut.Shortcut"), System.Windows.Forms.Shortcut)
        Me.MenuEditCut.ShowShortcut = CType(resources.GetObject("MenuEditCut.ShowShortcut"), Boolean)
        Me.MenuEditCut.Text = resources.GetString("MenuEditCut.Text")
        Me.MenuEditCut.Visible = CType(resources.GetObject("MenuEditCut.Visible"), Boolean)
        '
        'MenuEditCopy
        '
        Me.MenuEditCopy.Enabled = CType(resources.GetObject("MenuEditCopy.Enabled"), Boolean)
        Me.MenuEditCopy.Index = 4
        Me.MenuEditCopy.Shortcut = CType(resources.GetObject("MenuEditCopy.Shortcut"), System.Windows.Forms.Shortcut)
        Me.MenuEditCopy.ShowShortcut = CType(resources.GetObject("MenuEditCopy.ShowShortcut"), Boolean)
        Me.MenuEditCopy.Text = resources.GetString("MenuEditCopy.Text")
        Me.MenuEditCopy.Visible = CType(resources.GetObject("MenuEditCopy.Visible"), Boolean)
        '
        'MenuEditPaste
        '
        Me.MenuEditPaste.Enabled = CType(resources.GetObject("MenuEditPaste.Enabled"), Boolean)
        Me.MenuEditPaste.Index = 5
        Me.MenuEditPaste.Shortcut = CType(resources.GetObject("MenuEditPaste.Shortcut"), System.Windows.Forms.Shortcut)
        Me.MenuEditPaste.ShowShortcut = CType(resources.GetObject("MenuEditPaste.ShowShortcut"), Boolean)
        Me.MenuEditPaste.Text = resources.GetString("MenuEditPaste.Text")
        Me.MenuEditPaste.Visible = CType(resources.GetObject("MenuEditPaste.Visible"), Boolean)
        '
        'MenuEditDelete
        '
        Me.MenuEditDelete.Enabled = CType(resources.GetObject("MenuEditDelete.Enabled"), Boolean)
        Me.MenuEditDelete.Index = 6
        Me.MenuEditDelete.Shortcut = CType(resources.GetObject("MenuEditDelete.Shortcut"), System.Windows.Forms.Shortcut)
        Me.MenuEditDelete.ShowShortcut = CType(resources.GetObject("MenuEditDelete.ShowShortcut"), Boolean)
        Me.MenuEditDelete.Text = resources.GetString("MenuEditDelete.Text")
        Me.MenuEditDelete.Visible = CType(resources.GetObject("MenuEditDelete.Visible"), Boolean)
        '
        'MenuItem17
        '
        Me.MenuItem17.Enabled = CType(resources.GetObject("MenuItem17.Enabled"), Boolean)
        Me.MenuItem17.Index = 7
        Me.MenuItem17.Shortcut = CType(resources.GetObject("MenuItem17.Shortcut"), System.Windows.Forms.Shortcut)
        Me.MenuItem17.ShowShortcut = CType(resources.GetObject("MenuItem17.ShowShortcut"), Boolean)
        Me.MenuItem17.Text = resources.GetString("MenuItem17.Text")
        Me.MenuItem17.Visible = CType(resources.GetObject("MenuItem17.Visible"), Boolean)
        '
        'MenuEditSelectAll
        '
        Me.MenuEditSelectAll.Enabled = CType(resources.GetObject("MenuEditSelectAll.Enabled"), Boolean)
        Me.MenuEditSelectAll.Index = 8
        Me.MenuEditSelectAll.Shortcut = CType(resources.GetObject("MenuEditSelectAll.Shortcut"), System.Windows.Forms.Shortcut)
        Me.MenuEditSelectAll.ShowShortcut = CType(resources.GetObject("MenuEditSelectAll.ShowShortcut"), Boolean)
        Me.MenuEditSelectAll.Text = resources.GetString("MenuEditSelectAll.Text")
        Me.MenuEditSelectAll.Visible = CType(resources.GetObject("MenuEditSelectAll.Visible"), Boolean)
        '
        'MenuItem11
        '
        Me.MenuItem11.Enabled = CType(resources.GetObject("MenuItem11.Enabled"), Boolean)
        Me.MenuItem11.Index = 2
        Me.MenuItem11.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem67, Me.MenuItem87, Me.MenuItem88})
        Me.MenuItem11.Shortcut = CType(resources.GetObject("MenuItem11.Shortcut"), System.Windows.Forms.Shortcut)
        Me.MenuItem11.ShowShortcut = CType(resources.GetObject("MenuItem11.ShowShortcut"), Boolean)
        Me.MenuItem11.Text = resources.GetString("MenuItem11.Text")
        Me.MenuItem11.Visible = CType(resources.GetObject("MenuItem11.Visible"), Boolean)
        '
        'MenuItem67
        '
        Me.MenuItem67.Enabled = CType(resources.GetObject("MenuItem67.Enabled"), Boolean)
        Me.MenuItem67.Index = 0
        Me.MenuItem67.Shortcut = CType(resources.GetObject("MenuItem67.Shortcut"), System.Windows.Forms.Shortcut)
        Me.MenuItem67.ShowShortcut = CType(resources.GetObject("MenuItem67.ShowShortcut"), Boolean)
        Me.MenuItem67.Text = resources.GetString("MenuItem67.Text")
        Me.MenuItem67.Visible = CType(resources.GetObject("MenuItem67.Visible"), Boolean)
        '
        'MenuItem87
        '
        Me.MenuItem87.Enabled = CType(resources.GetObject("MenuItem87.Enabled"), Boolean)
        Me.MenuItem87.Index = 1
        Me.MenuItem87.Shortcut = CType(resources.GetObject("MenuItem87.Shortcut"), System.Windows.Forms.Shortcut)
        Me.MenuItem87.ShowShortcut = CType(resources.GetObject("MenuItem87.ShowShortcut"), Boolean)
        Me.MenuItem87.Text = resources.GetString("MenuItem87.Text")
        Me.MenuItem87.Visible = CType(resources.GetObject("MenuItem87.Visible"), Boolean)
        '
        'MenuItem88
        '
        Me.MenuItem88.Enabled = CType(resources.GetObject("MenuItem88.Enabled"), Boolean)
        Me.MenuItem88.Index = 2
        Me.MenuItem88.Shortcut = CType(resources.GetObject("MenuItem88.Shortcut"), System.Windows.Forms.Shortcut)
        Me.MenuItem88.ShowShortcut = CType(resources.GetObject("MenuItem88.ShowShortcut"), Boolean)
        Me.MenuItem88.Text = resources.GetString("MenuItem88.Text")
        Me.MenuItem88.Visible = CType(resources.GetObject("MenuItem88.Visible"), Boolean)
        '
        'MenuItem25
        '
        Me.MenuItem25.Enabled = CType(resources.GetObject("MenuItem25.Enabled"), Boolean)
        Me.MenuItem25.Index = 3
        Me.MenuItem25.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem33, Me.MenuItem32, Me.MenuItem35, Me.MenuItem29, Me.MenuItem30, Me.MenuItem36, Me.MenuItem62, Me.MenuItem76, Me.OptimizePath, Me.MenuItem3})
        Me.MenuItem25.Shortcut = CType(resources.GetObject("MenuItem25.Shortcut"), System.Windows.Forms.Shortcut)
        Me.MenuItem25.ShowShortcut = CType(resources.GetObject("MenuItem25.ShowShortcut"), Boolean)
        Me.MenuItem25.Text = resources.GetString("MenuItem25.Text")
        Me.MenuItem25.Visible = CType(resources.GetObject("MenuItem25.Visible"), Boolean)
        '
        'MenuItem33
        '
        Me.MenuItem33.Enabled = CType(resources.GetObject("MenuItem33.Enabled"), Boolean)
        Me.MenuItem33.Index = 0
        Me.MenuItem33.Shortcut = CType(resources.GetObject("MenuItem33.Shortcut"), System.Windows.Forms.Shortcut)
        Me.MenuItem33.ShowShortcut = CType(resources.GetObject("MenuItem33.ShowShortcut"), Boolean)
        Me.MenuItem33.Text = resources.GetString("MenuItem33.Text")
        Me.MenuItem33.Visible = CType(resources.GetObject("MenuItem33.Visible"), Boolean)
        '
        'MenuItem32
        '
        Me.MenuItem32.Enabled = CType(resources.GetObject("MenuItem32.Enabled"), Boolean)
        Me.MenuItem32.Index = 1
        Me.MenuItem32.Shortcut = CType(resources.GetObject("MenuItem32.Shortcut"), System.Windows.Forms.Shortcut)
        Me.MenuItem32.ShowShortcut = CType(resources.GetObject("MenuItem32.ShowShortcut"), Boolean)
        Me.MenuItem32.Text = resources.GetString("MenuItem32.Text")
        Me.MenuItem32.Visible = CType(resources.GetObject("MenuItem32.Visible"), Boolean)
        '
        'MenuItem35
        '
        Me.MenuItem35.Enabled = CType(resources.GetObject("MenuItem35.Enabled"), Boolean)
        Me.MenuItem35.Index = 2
        Me.MenuItem35.Shortcut = CType(resources.GetObject("MenuItem35.Shortcut"), System.Windows.Forms.Shortcut)
        Me.MenuItem35.ShowShortcut = CType(resources.GetObject("MenuItem35.ShowShortcut"), Boolean)
        Me.MenuItem35.Text = resources.GetString("MenuItem35.Text")
        Me.MenuItem35.Visible = CType(resources.GetObject("MenuItem35.Visible"), Boolean)
        '
        'MenuItem29
        '
        Me.MenuItem29.Enabled = CType(resources.GetObject("MenuItem29.Enabled"), Boolean)
        Me.MenuItem29.Index = 3
        Me.MenuItem29.Shortcut = CType(resources.GetObject("MenuItem29.Shortcut"), System.Windows.Forms.Shortcut)
        Me.MenuItem29.ShowShortcut = CType(resources.GetObject("MenuItem29.ShowShortcut"), Boolean)
        Me.MenuItem29.Text = resources.GetString("MenuItem29.Text")
        Me.MenuItem29.Visible = CType(resources.GetObject("MenuItem29.Visible"), Boolean)
        '
        'MenuItem30
        '
        Me.MenuItem30.Enabled = CType(resources.GetObject("MenuItem30.Enabled"), Boolean)
        Me.MenuItem30.Index = 4
        Me.MenuItem30.Shortcut = CType(resources.GetObject("MenuItem30.Shortcut"), System.Windows.Forms.Shortcut)
        Me.MenuItem30.ShowShortcut = CType(resources.GetObject("MenuItem30.ShowShortcut"), Boolean)
        Me.MenuItem30.Text = resources.GetString("MenuItem30.Text")
        Me.MenuItem30.Visible = CType(resources.GetObject("MenuItem30.Visible"), Boolean)
        '
        'MenuItem36
        '
        Me.MenuItem36.Enabled = CType(resources.GetObject("MenuItem36.Enabled"), Boolean)
        Me.MenuItem36.Index = 5
        Me.MenuItem36.Shortcut = CType(resources.GetObject("MenuItem36.Shortcut"), System.Windows.Forms.Shortcut)
        Me.MenuItem36.ShowShortcut = CType(resources.GetObject("MenuItem36.ShowShortcut"), Boolean)
        Me.MenuItem36.Text = resources.GetString("MenuItem36.Text")
        Me.MenuItem36.Visible = CType(resources.GetObject("MenuItem36.Visible"), Boolean)
        '
        'MenuItem62
        '
        Me.MenuItem62.Enabled = CType(resources.GetObject("MenuItem62.Enabled"), Boolean)
        Me.MenuItem62.Index = 6
        Me.MenuItem62.Shortcut = CType(resources.GetObject("MenuItem62.Shortcut"), System.Windows.Forms.Shortcut)
        Me.MenuItem62.ShowShortcut = CType(resources.GetObject("MenuItem62.ShowShortcut"), Boolean)
        Me.MenuItem62.Text = resources.GetString("MenuItem62.Text")
        Me.MenuItem62.Visible = CType(resources.GetObject("MenuItem62.Visible"), Boolean)
        '
        'MenuItem76
        '
        Me.MenuItem76.Enabled = CType(resources.GetObject("MenuItem76.Enabled"), Boolean)
        Me.MenuItem76.Index = 7
        Me.MenuItem76.Shortcut = CType(resources.GetObject("MenuItem76.Shortcut"), System.Windows.Forms.Shortcut)
        Me.MenuItem76.ShowShortcut = CType(resources.GetObject("MenuItem76.ShowShortcut"), Boolean)
        Me.MenuItem76.Text = resources.GetString("MenuItem76.Text")
        Me.MenuItem76.Visible = CType(resources.GetObject("MenuItem76.Visible"), Boolean)
        '
        'OptimizePath
        '
        Me.OptimizePath.Enabled = CType(resources.GetObject("OptimizePath.Enabled"), Boolean)
        Me.OptimizePath.Index = 8
        Me.OptimizePath.Shortcut = CType(resources.GetObject("OptimizePath.Shortcut"), System.Windows.Forms.Shortcut)
        Me.OptimizePath.ShowShortcut = CType(resources.GetObject("OptimizePath.ShowShortcut"), Boolean)
        Me.OptimizePath.Text = resources.GetString("OptimizePath.Text")
        Me.OptimizePath.Visible = CType(resources.GetObject("OptimizePath.Visible"), Boolean)
        '
        'MenuItem3
        '
        Me.MenuItem3.Enabled = CType(resources.GetObject("MenuItem3.Enabled"), Boolean)
        Me.MenuItem3.Index = 9
        Me.MenuItem3.Shortcut = CType(resources.GetObject("MenuItem3.Shortcut"), System.Windows.Forms.Shortcut)
        Me.MenuItem3.ShowShortcut = CType(resources.GetObject("MenuItem3.ShowShortcut"), Boolean)
        Me.MenuItem3.Text = resources.GetString("MenuItem3.Text")
        Me.MenuItem3.Visible = CType(resources.GetObject("MenuItem3.Visible"), Boolean)
        '
        'MenuItem64
        '
        Me.MenuItem64.Enabled = CType(resources.GetObject("MenuItem64.Enabled"), Boolean)
        Me.MenuItem64.Index = 4
        Me.MenuItem64.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem65})
        Me.MenuItem64.Shortcut = CType(resources.GetObject("MenuItem64.Shortcut"), System.Windows.Forms.Shortcut)
        Me.MenuItem64.ShowShortcut = CType(resources.GetObject("MenuItem64.ShowShortcut"), Boolean)
        Me.MenuItem64.Text = resources.GetString("MenuItem64.Text")
        Me.MenuItem64.Visible = CType(resources.GetObject("MenuItem64.Visible"), Boolean)
        '
        'MenuItem65
        '
        Me.MenuItem65.Enabled = CType(resources.GetObject("MenuItem65.Enabled"), Boolean)
        Me.MenuItem65.Index = 0
        Me.MenuItem65.Shortcut = CType(resources.GetObject("MenuItem65.Shortcut"), System.Windows.Forms.Shortcut)
        Me.MenuItem65.ShowShortcut = CType(resources.GetObject("MenuItem65.ShowShortcut"), Boolean)
        Me.MenuItem65.Text = resources.GetString("MenuItem65.Text")
        Me.MenuItem65.Visible = CType(resources.GetObject("MenuItem65.Visible"), Boolean)
        '
        'PanelVision
        '
        Me.PanelVision.AccessibleDescription = resources.GetString("PanelVision.AccessibleDescription")
        Me.PanelVision.AccessibleName = resources.GetString("PanelVision.AccessibleName")
        Me.PanelVision.Anchor = CType(resources.GetObject("PanelVision.Anchor"), System.Windows.Forms.AnchorStyles)
        Me.PanelVision.AutoScroll = CType(resources.GetObject("PanelVision.AutoScroll"), Boolean)
        Me.PanelVision.AutoScrollMargin = CType(resources.GetObject("PanelVision.AutoScrollMargin"), System.Drawing.Size)
        Me.PanelVision.AutoScrollMinSize = CType(resources.GetObject("PanelVision.AutoScrollMinSize"), System.Drawing.Size)
        Me.PanelVision.BackColor = System.Drawing.Color.SlateGray
        Me.PanelVision.BackgroundImage = CType(resources.GetObject("PanelVision.BackgroundImage"), System.Drawing.Image)
        Me.PanelVision.Dock = CType(resources.GetObject("PanelVision.Dock"), System.Windows.Forms.DockStyle)
        Me.PanelVision.Enabled = CType(resources.GetObject("PanelVision.Enabled"), Boolean)
        Me.PanelVision.Font = CType(resources.GetObject("PanelVision.Font"), System.Drawing.Font)
        Me.PanelVision.ImeMode = CType(resources.GetObject("PanelVision.ImeMode"), System.Windows.Forms.ImeMode)
        Me.PanelVision.Location = CType(resources.GetObject("PanelVision.Location"), System.Drawing.Point)
        Me.PanelVision.Name = "PanelVision"
        Me.PanelVision.RightToLeft = CType(resources.GetObject("PanelVision.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.PanelVision.Size = CType(resources.GetObject("PanelVision.Size"), System.Drawing.Size)
        Me.PanelVision.TabIndex = CType(resources.GetObject("PanelVision.TabIndex"), Integer)
        Me.PanelVision.Text = resources.GetString("PanelVision.Text")
        Me.ToolTip1.SetToolTip(Me.PanelVision, resources.GetString("PanelVision.ToolTip"))
        Me.PanelVision.Visible = CType(resources.GetObject("PanelVision.Visible"), Boolean)
        '
        'ImageListGeneralTools
        '
        Me.ImageListGeneralTools.ImageSize = CType(resources.GetObject("ImageListGeneralTools.ImageSize"), System.Drawing.Size)
        Me.ImageListGeneralTools.ImageStream = CType(resources.GetObject("ImageListGeneralTools.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.ImageListGeneralTools.TransparentColor = System.Drawing.Color.White
        '
        'ButtonPurge
        '
        Me.ButtonPurge.AccessibleDescription = resources.GetString("ButtonPurge.AccessibleDescription")
        Me.ButtonPurge.AccessibleName = resources.GetString("ButtonPurge.AccessibleName")
        Me.ButtonPurge.Anchor = CType(resources.GetObject("ButtonPurge.Anchor"), System.Windows.Forms.AnchorStyles)
        Me.ButtonPurge.BackColor = System.Drawing.SystemColors.Control
        Me.ButtonPurge.BackgroundImage = CType(resources.GetObject("ButtonPurge.BackgroundImage"), System.Drawing.Image)
        Me.ButtonPurge.Dock = CType(resources.GetObject("ButtonPurge.Dock"), System.Windows.Forms.DockStyle)
        Me.ButtonPurge.Enabled = CType(resources.GetObject("ButtonPurge.Enabled"), Boolean)
        Me.ButtonPurge.FlatStyle = CType(resources.GetObject("ButtonPurge.FlatStyle"), System.Windows.Forms.FlatStyle)
        Me.ButtonPurge.Font = CType(resources.GetObject("ButtonPurge.Font"), System.Drawing.Font)
        Me.ButtonPurge.Image = CType(resources.GetObject("ButtonPurge.Image"), System.Drawing.Image)
        Me.ButtonPurge.ImageAlign = CType(resources.GetObject("ButtonPurge.ImageAlign"), System.Drawing.ContentAlignment)
        Me.ButtonPurge.ImageIndex = CType(resources.GetObject("ButtonPurge.ImageIndex"), Integer)
        Me.ButtonPurge.ImageList = Me.ImageListGeneralTools
        Me.ButtonPurge.ImeMode = CType(resources.GetObject("ButtonPurge.ImeMode"), System.Windows.Forms.ImeMode)
        Me.ButtonPurge.Location = CType(resources.GetObject("ButtonPurge.Location"), System.Drawing.Point)
        Me.ButtonPurge.Name = "ButtonPurge"
        Me.ButtonPurge.RightToLeft = CType(resources.GetObject("ButtonPurge.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.ButtonPurge.Size = CType(resources.GetObject("ButtonPurge.Size"), System.Drawing.Size)
        Me.ButtonPurge.TabIndex = CType(resources.GetObject("ButtonPurge.TabIndex"), Integer)
        Me.ButtonPurge.Text = resources.GetString("ButtonPurge.Text")
        Me.ButtonPurge.TextAlign = CType(resources.GetObject("ButtonPurge.TextAlign"), System.Drawing.ContentAlignment)
        Me.ToolTip1.SetToolTip(Me.ButtonPurge, resources.GetString("ButtonPurge.ToolTip"))
        Me.ButtonPurge.Visible = CType(resources.GetObject("ButtonPurge.Visible"), Boolean)
        '
        'PanelVisionCtrl
        '
        Me.PanelVisionCtrl.AccessibleDescription = resources.GetString("PanelVisionCtrl.AccessibleDescription")
        Me.PanelVisionCtrl.AccessibleName = resources.GetString("PanelVisionCtrl.AccessibleName")
        Me.PanelVisionCtrl.Anchor = CType(resources.GetObject("PanelVisionCtrl.Anchor"), System.Windows.Forms.AnchorStyles)
        Me.PanelVisionCtrl.AutoScroll = CType(resources.GetObject("PanelVisionCtrl.AutoScroll"), Boolean)
        Me.PanelVisionCtrl.AutoScrollMargin = CType(resources.GetObject("PanelVisionCtrl.AutoScrollMargin"), System.Drawing.Size)
        Me.PanelVisionCtrl.AutoScrollMinSize = CType(resources.GetObject("PanelVisionCtrl.AutoScrollMinSize"), System.Drawing.Size)
        Me.PanelVisionCtrl.BackgroundImage = CType(resources.GetObject("PanelVisionCtrl.BackgroundImage"), System.Drawing.Image)
        Me.PanelVisionCtrl.Controls.Add(Me.cbDisplayIndicator)
        Me.PanelVisionCtrl.Controls.Add(Me.ValueBrightness)
        Me.PanelVisionCtrl.Controls.Add(Me.CheckBoxLockZ)
        Me.PanelVisionCtrl.Controls.Add(Me.TextBoxRobotZ)
        Me.PanelVisionCtrl.Controls.Add(Me.CheckBoxLockY)
        Me.PanelVisionCtrl.Controls.Add(Me.TextBoxRobotY)
        Me.PanelVisionCtrl.Controls.Add(Me.CheckBoxLockX)
        Me.PanelVisionCtrl.Controls.Add(Me.TextBoxRobotX)
        Me.PanelVisionCtrl.Controls.Add(Me.lbPostName)
        Me.PanelVisionCtrl.Controls.Add(Me.DomainUpDown1)
        Me.PanelVisionCtrl.Controls.Add(Me.Label1)
        Me.PanelVisionCtrl.Dock = CType(resources.GetObject("PanelVisionCtrl.Dock"), System.Windows.Forms.DockStyle)
        Me.PanelVisionCtrl.Enabled = CType(resources.GetObject("PanelVisionCtrl.Enabled"), Boolean)
        Me.PanelVisionCtrl.Font = CType(resources.GetObject("PanelVisionCtrl.Font"), System.Drawing.Font)
        Me.PanelVisionCtrl.ImeMode = CType(resources.GetObject("PanelVisionCtrl.ImeMode"), System.Windows.Forms.ImeMode)
        Me.PanelVisionCtrl.Location = CType(resources.GetObject("PanelVisionCtrl.Location"), System.Drawing.Point)
        Me.PanelVisionCtrl.Name = "PanelVisionCtrl"
        Me.PanelVisionCtrl.RightToLeft = CType(resources.GetObject("PanelVisionCtrl.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.PanelVisionCtrl.Size = CType(resources.GetObject("PanelVisionCtrl.Size"), System.Drawing.Size)
        Me.PanelVisionCtrl.TabIndex = CType(resources.GetObject("PanelVisionCtrl.TabIndex"), Integer)
        Me.PanelVisionCtrl.Text = resources.GetString("PanelVisionCtrl.Text")
        Me.ToolTip1.SetToolTip(Me.PanelVisionCtrl, resources.GetString("PanelVisionCtrl.ToolTip"))
        Me.PanelVisionCtrl.Visible = CType(resources.GetObject("PanelVisionCtrl.Visible"), Boolean)
        '
        'cbDisplayIndicator
        '
        Me.cbDisplayIndicator.AccessibleDescription = resources.GetString("cbDisplayIndicator.AccessibleDescription")
        Me.cbDisplayIndicator.AccessibleName = resources.GetString("cbDisplayIndicator.AccessibleName")
        Me.cbDisplayIndicator.Anchor = CType(resources.GetObject("cbDisplayIndicator.Anchor"), System.Windows.Forms.AnchorStyles)
        Me.cbDisplayIndicator.Appearance = CType(resources.GetObject("cbDisplayIndicator.Appearance"), System.Windows.Forms.Appearance)
        Me.cbDisplayIndicator.BackgroundImage = CType(resources.GetObject("cbDisplayIndicator.BackgroundImage"), System.Drawing.Image)
        Me.cbDisplayIndicator.CheckAlign = CType(resources.GetObject("cbDisplayIndicator.CheckAlign"), System.Drawing.ContentAlignment)
        Me.cbDisplayIndicator.Checked = True
        Me.cbDisplayIndicator.CheckState = System.Windows.Forms.CheckState.Checked
        Me.cbDisplayIndicator.Dock = CType(resources.GetObject("cbDisplayIndicator.Dock"), System.Windows.Forms.DockStyle)
        Me.cbDisplayIndicator.Enabled = CType(resources.GetObject("cbDisplayIndicator.Enabled"), Boolean)
        Me.cbDisplayIndicator.FlatStyle = CType(resources.GetObject("cbDisplayIndicator.FlatStyle"), System.Windows.Forms.FlatStyle)
        Me.cbDisplayIndicator.Font = CType(resources.GetObject("cbDisplayIndicator.Font"), System.Drawing.Font)
        Me.cbDisplayIndicator.Image = CType(resources.GetObject("cbDisplayIndicator.Image"), System.Drawing.Image)
        Me.cbDisplayIndicator.ImageAlign = CType(resources.GetObject("cbDisplayIndicator.ImageAlign"), System.Drawing.ContentAlignment)
        Me.cbDisplayIndicator.ImageIndex = CType(resources.GetObject("cbDisplayIndicator.ImageIndex"), Integer)
        Me.cbDisplayIndicator.ImeMode = CType(resources.GetObject("cbDisplayIndicator.ImeMode"), System.Windows.Forms.ImeMode)
        Me.cbDisplayIndicator.Location = CType(resources.GetObject("cbDisplayIndicator.Location"), System.Drawing.Point)
        Me.cbDisplayIndicator.Name = "cbDisplayIndicator"
        Me.cbDisplayIndicator.RightToLeft = CType(resources.GetObject("cbDisplayIndicator.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.cbDisplayIndicator.Size = CType(resources.GetObject("cbDisplayIndicator.Size"), System.Drawing.Size)
        Me.cbDisplayIndicator.TabIndex = CType(resources.GetObject("cbDisplayIndicator.TabIndex"), Integer)
        Me.cbDisplayIndicator.Text = resources.GetString("cbDisplayIndicator.Text")
        Me.cbDisplayIndicator.TextAlign = CType(resources.GetObject("cbDisplayIndicator.TextAlign"), System.Drawing.ContentAlignment)
        Me.ToolTip1.SetToolTip(Me.cbDisplayIndicator, resources.GetString("cbDisplayIndicator.ToolTip"))
        Me.cbDisplayIndicator.Visible = CType(resources.GetObject("cbDisplayIndicator.Visible"), Boolean)
        '
        'ValueBrightness
        '
        Me.ValueBrightness.AccessibleDescription = resources.GetString("ValueBrightness.AccessibleDescription")
        Me.ValueBrightness.AccessibleName = resources.GetString("ValueBrightness.AccessibleName")
        Me.ValueBrightness.Anchor = CType(resources.GetObject("ValueBrightness.Anchor"), System.Windows.Forms.AnchorStyles)
        Me.ValueBrightness.Dock = CType(resources.GetObject("ValueBrightness.Dock"), System.Windows.Forms.DockStyle)
        Me.ValueBrightness.Enabled = CType(resources.GetObject("ValueBrightness.Enabled"), Boolean)
        Me.ValueBrightness.Font = CType(resources.GetObject("ValueBrightness.Font"), System.Drawing.Font)
        Me.ValueBrightness.ImeMode = CType(resources.GetObject("ValueBrightness.ImeMode"), System.Windows.Forms.ImeMode)
        Me.ValueBrightness.Location = CType(resources.GetObject("ValueBrightness.Location"), System.Drawing.Point)
        Me.ValueBrightness.Maximum = New Decimal(New Integer() {255, 0, 0, 0})
        Me.ValueBrightness.Name = "ValueBrightness"
        Me.ValueBrightness.RightToLeft = CType(resources.GetObject("ValueBrightness.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.ValueBrightness.Size = CType(resources.GetObject("ValueBrightness.Size"), System.Drawing.Size)
        Me.ValueBrightness.TabIndex = CType(resources.GetObject("ValueBrightness.TabIndex"), Integer)
        Me.ValueBrightness.TextAlign = CType(resources.GetObject("ValueBrightness.TextAlign"), System.Windows.Forms.HorizontalAlignment)
        Me.ValueBrightness.ThousandsSeparator = CType(resources.GetObject("ValueBrightness.ThousandsSeparator"), Boolean)
        Me.ToolTip1.SetToolTip(Me.ValueBrightness, resources.GetString("ValueBrightness.ToolTip"))
        Me.ValueBrightness.UpDownAlign = CType(resources.GetObject("ValueBrightness.UpDownAlign"), System.Windows.Forms.LeftRightAlignment)
        Me.ValueBrightness.Value = New Decimal(New Integer() {10, 0, 0, 0})
        Me.ValueBrightness.Visible = CType(resources.GetObject("ValueBrightness.Visible"), Boolean)
        '
        'CheckBoxLockZ
        '
        Me.CheckBoxLockZ.AccessibleDescription = resources.GetString("CheckBoxLockZ.AccessibleDescription")
        Me.CheckBoxLockZ.AccessibleName = resources.GetString("CheckBoxLockZ.AccessibleName")
        Me.CheckBoxLockZ.Anchor = CType(resources.GetObject("CheckBoxLockZ.Anchor"), System.Windows.Forms.AnchorStyles)
        Me.CheckBoxLockZ.Appearance = CType(resources.GetObject("CheckBoxLockZ.Appearance"), System.Windows.Forms.Appearance)
        Me.CheckBoxLockZ.BackColor = System.Drawing.SystemColors.Control
        Me.CheckBoxLockZ.BackgroundImage = CType(resources.GetObject("CheckBoxLockZ.BackgroundImage"), System.Drawing.Image)
        Me.CheckBoxLockZ.CheckAlign = CType(resources.GetObject("CheckBoxLockZ.CheckAlign"), System.Drawing.ContentAlignment)
        Me.CheckBoxLockZ.Dock = CType(resources.GetObject("CheckBoxLockZ.Dock"), System.Windows.Forms.DockStyle)
        Me.CheckBoxLockZ.Enabled = CType(resources.GetObject("CheckBoxLockZ.Enabled"), Boolean)
        Me.CheckBoxLockZ.FlatStyle = CType(resources.GetObject("CheckBoxLockZ.FlatStyle"), System.Windows.Forms.FlatStyle)
        Me.CheckBoxLockZ.Font = CType(resources.GetObject("CheckBoxLockZ.Font"), System.Drawing.Font)
        Me.CheckBoxLockZ.Image = CType(resources.GetObject("CheckBoxLockZ.Image"), System.Drawing.Image)
        Me.CheckBoxLockZ.ImageAlign = CType(resources.GetObject("CheckBoxLockZ.ImageAlign"), System.Drawing.ContentAlignment)
        Me.CheckBoxLockZ.ImageIndex = CType(resources.GetObject("CheckBoxLockZ.ImageIndex"), Integer)
        Me.CheckBoxLockZ.ImeMode = CType(resources.GetObject("CheckBoxLockZ.ImeMode"), System.Windows.Forms.ImeMode)
        Me.CheckBoxLockZ.Location = CType(resources.GetObject("CheckBoxLockZ.Location"), System.Drawing.Point)
        Me.CheckBoxLockZ.Name = "CheckBoxLockZ"
        Me.CheckBoxLockZ.RightToLeft = CType(resources.GetObject("CheckBoxLockZ.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.CheckBoxLockZ.Size = CType(resources.GetObject("CheckBoxLockZ.Size"), System.Drawing.Size)
        Me.CheckBoxLockZ.TabIndex = CType(resources.GetObject("CheckBoxLockZ.TabIndex"), Integer)
        Me.CheckBoxLockZ.Text = resources.GetString("CheckBoxLockZ.Text")
        Me.CheckBoxLockZ.TextAlign = CType(resources.GetObject("CheckBoxLockZ.TextAlign"), System.Drawing.ContentAlignment)
        Me.ToolTip1.SetToolTip(Me.CheckBoxLockZ, resources.GetString("CheckBoxLockZ.ToolTip"))
        Me.CheckBoxLockZ.Visible = CType(resources.GetObject("CheckBoxLockZ.Visible"), Boolean)
        '
        'TextBoxRobotZ
        '
        Me.TextBoxRobotZ.AccessibleDescription = resources.GetString("TextBoxRobotZ.AccessibleDescription")
        Me.TextBoxRobotZ.AccessibleName = resources.GetString("TextBoxRobotZ.AccessibleName")
        Me.TextBoxRobotZ.Anchor = CType(resources.GetObject("TextBoxRobotZ.Anchor"), System.Windows.Forms.AnchorStyles)
        Me.TextBoxRobotZ.AutoSize = CType(resources.GetObject("TextBoxRobotZ.AutoSize"), Boolean)
        Me.TextBoxRobotZ.BackgroundImage = CType(resources.GetObject("TextBoxRobotZ.BackgroundImage"), System.Drawing.Image)
        Me.TextBoxRobotZ.Dock = CType(resources.GetObject("TextBoxRobotZ.Dock"), System.Windows.Forms.DockStyle)
        Me.TextBoxRobotZ.Enabled = CType(resources.GetObject("TextBoxRobotZ.Enabled"), Boolean)
        Me.TextBoxRobotZ.Font = CType(resources.GetObject("TextBoxRobotZ.Font"), System.Drawing.Font)
        Me.TextBoxRobotZ.ImeMode = CType(resources.GetObject("TextBoxRobotZ.ImeMode"), System.Windows.Forms.ImeMode)
        Me.TextBoxRobotZ.Location = CType(resources.GetObject("TextBoxRobotZ.Location"), System.Drawing.Point)
        Me.TextBoxRobotZ.MaxLength = CType(resources.GetObject("TextBoxRobotZ.MaxLength"), Integer)
        Me.TextBoxRobotZ.Multiline = CType(resources.GetObject("TextBoxRobotZ.Multiline"), Boolean)
        Me.TextBoxRobotZ.Name = "TextBoxRobotZ"
        Me.TextBoxRobotZ.PasswordChar = CType(resources.GetObject("TextBoxRobotZ.PasswordChar"), Char)
        Me.TextBoxRobotZ.ReadOnly = True
        Me.TextBoxRobotZ.RightToLeft = CType(resources.GetObject("TextBoxRobotZ.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.TextBoxRobotZ.ScrollBars = CType(resources.GetObject("TextBoxRobotZ.ScrollBars"), System.Windows.Forms.ScrollBars)
        Me.TextBoxRobotZ.Size = CType(resources.GetObject("TextBoxRobotZ.Size"), System.Drawing.Size)
        Me.TextBoxRobotZ.TabIndex = CType(resources.GetObject("TextBoxRobotZ.TabIndex"), Integer)
        Me.TextBoxRobotZ.Text = resources.GetString("TextBoxRobotZ.Text")
        Me.TextBoxRobotZ.TextAlign = CType(resources.GetObject("TextBoxRobotZ.TextAlign"), System.Windows.Forms.HorizontalAlignment)
        Me.ToolTip1.SetToolTip(Me.TextBoxRobotZ, resources.GetString("TextBoxRobotZ.ToolTip"))
        Me.TextBoxRobotZ.Visible = CType(resources.GetObject("TextBoxRobotZ.Visible"), Boolean)
        Me.TextBoxRobotZ.WordWrap = CType(resources.GetObject("TextBoxRobotZ.WordWrap"), Boolean)
        '
        'CheckBoxLockY
        '
        Me.CheckBoxLockY.AccessibleDescription = resources.GetString("CheckBoxLockY.AccessibleDescription")
        Me.CheckBoxLockY.AccessibleName = resources.GetString("CheckBoxLockY.AccessibleName")
        Me.CheckBoxLockY.Anchor = CType(resources.GetObject("CheckBoxLockY.Anchor"), System.Windows.Forms.AnchorStyles)
        Me.CheckBoxLockY.Appearance = CType(resources.GetObject("CheckBoxLockY.Appearance"), System.Windows.Forms.Appearance)
        Me.CheckBoxLockY.BackColor = System.Drawing.SystemColors.Control
        Me.CheckBoxLockY.BackgroundImage = CType(resources.GetObject("CheckBoxLockY.BackgroundImage"), System.Drawing.Image)
        Me.CheckBoxLockY.CheckAlign = CType(resources.GetObject("CheckBoxLockY.CheckAlign"), System.Drawing.ContentAlignment)
        Me.CheckBoxLockY.Dock = CType(resources.GetObject("CheckBoxLockY.Dock"), System.Windows.Forms.DockStyle)
        Me.CheckBoxLockY.Enabled = CType(resources.GetObject("CheckBoxLockY.Enabled"), Boolean)
        Me.CheckBoxLockY.FlatStyle = CType(resources.GetObject("CheckBoxLockY.FlatStyle"), System.Windows.Forms.FlatStyle)
        Me.CheckBoxLockY.Font = CType(resources.GetObject("CheckBoxLockY.Font"), System.Drawing.Font)
        Me.CheckBoxLockY.Image = CType(resources.GetObject("CheckBoxLockY.Image"), System.Drawing.Image)
        Me.CheckBoxLockY.ImageAlign = CType(resources.GetObject("CheckBoxLockY.ImageAlign"), System.Drawing.ContentAlignment)
        Me.CheckBoxLockY.ImageIndex = CType(resources.GetObject("CheckBoxLockY.ImageIndex"), Integer)
        Me.CheckBoxLockY.ImeMode = CType(resources.GetObject("CheckBoxLockY.ImeMode"), System.Windows.Forms.ImeMode)
        Me.CheckBoxLockY.Location = CType(resources.GetObject("CheckBoxLockY.Location"), System.Drawing.Point)
        Me.CheckBoxLockY.Name = "CheckBoxLockY"
        Me.CheckBoxLockY.RightToLeft = CType(resources.GetObject("CheckBoxLockY.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.CheckBoxLockY.Size = CType(resources.GetObject("CheckBoxLockY.Size"), System.Drawing.Size)
        Me.CheckBoxLockY.TabIndex = CType(resources.GetObject("CheckBoxLockY.TabIndex"), Integer)
        Me.CheckBoxLockY.Text = resources.GetString("CheckBoxLockY.Text")
        Me.CheckBoxLockY.TextAlign = CType(resources.GetObject("CheckBoxLockY.TextAlign"), System.Drawing.ContentAlignment)
        Me.ToolTip1.SetToolTip(Me.CheckBoxLockY, resources.GetString("CheckBoxLockY.ToolTip"))
        Me.CheckBoxLockY.Visible = CType(resources.GetObject("CheckBoxLockY.Visible"), Boolean)
        '
        'TextBoxRobotY
        '
        Me.TextBoxRobotY.AccessibleDescription = resources.GetString("TextBoxRobotY.AccessibleDescription")
        Me.TextBoxRobotY.AccessibleName = resources.GetString("TextBoxRobotY.AccessibleName")
        Me.TextBoxRobotY.Anchor = CType(resources.GetObject("TextBoxRobotY.Anchor"), System.Windows.Forms.AnchorStyles)
        Me.TextBoxRobotY.AutoSize = CType(resources.GetObject("TextBoxRobotY.AutoSize"), Boolean)
        Me.TextBoxRobotY.BackgroundImage = CType(resources.GetObject("TextBoxRobotY.BackgroundImage"), System.Drawing.Image)
        Me.TextBoxRobotY.Dock = CType(resources.GetObject("TextBoxRobotY.Dock"), System.Windows.Forms.DockStyle)
        Me.TextBoxRobotY.Enabled = CType(resources.GetObject("TextBoxRobotY.Enabled"), Boolean)
        Me.TextBoxRobotY.Font = CType(resources.GetObject("TextBoxRobotY.Font"), System.Drawing.Font)
        Me.TextBoxRobotY.ImeMode = CType(resources.GetObject("TextBoxRobotY.ImeMode"), System.Windows.Forms.ImeMode)
        Me.TextBoxRobotY.Location = CType(resources.GetObject("TextBoxRobotY.Location"), System.Drawing.Point)
        Me.TextBoxRobotY.MaxLength = CType(resources.GetObject("TextBoxRobotY.MaxLength"), Integer)
        Me.TextBoxRobotY.Multiline = CType(resources.GetObject("TextBoxRobotY.Multiline"), Boolean)
        Me.TextBoxRobotY.Name = "TextBoxRobotY"
        Me.TextBoxRobotY.PasswordChar = CType(resources.GetObject("TextBoxRobotY.PasswordChar"), Char)
        Me.TextBoxRobotY.ReadOnly = True
        Me.TextBoxRobotY.RightToLeft = CType(resources.GetObject("TextBoxRobotY.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.TextBoxRobotY.ScrollBars = CType(resources.GetObject("TextBoxRobotY.ScrollBars"), System.Windows.Forms.ScrollBars)
        Me.TextBoxRobotY.Size = CType(resources.GetObject("TextBoxRobotY.Size"), System.Drawing.Size)
        Me.TextBoxRobotY.TabIndex = CType(resources.GetObject("TextBoxRobotY.TabIndex"), Integer)
        Me.TextBoxRobotY.Text = resources.GetString("TextBoxRobotY.Text")
        Me.TextBoxRobotY.TextAlign = CType(resources.GetObject("TextBoxRobotY.TextAlign"), System.Windows.Forms.HorizontalAlignment)
        Me.ToolTip1.SetToolTip(Me.TextBoxRobotY, resources.GetString("TextBoxRobotY.ToolTip"))
        Me.TextBoxRobotY.Visible = CType(resources.GetObject("TextBoxRobotY.Visible"), Boolean)
        Me.TextBoxRobotY.WordWrap = CType(resources.GetObject("TextBoxRobotY.WordWrap"), Boolean)
        '
        'CheckBoxLockX
        '
        Me.CheckBoxLockX.AccessibleDescription = resources.GetString("CheckBoxLockX.AccessibleDescription")
        Me.CheckBoxLockX.AccessibleName = resources.GetString("CheckBoxLockX.AccessibleName")
        Me.CheckBoxLockX.Anchor = CType(resources.GetObject("CheckBoxLockX.Anchor"), System.Windows.Forms.AnchorStyles)
        Me.CheckBoxLockX.Appearance = CType(resources.GetObject("CheckBoxLockX.Appearance"), System.Windows.Forms.Appearance)
        Me.CheckBoxLockX.BackColor = System.Drawing.SystemColors.Control
        Me.CheckBoxLockX.BackgroundImage = CType(resources.GetObject("CheckBoxLockX.BackgroundImage"), System.Drawing.Image)
        Me.CheckBoxLockX.CheckAlign = CType(resources.GetObject("CheckBoxLockX.CheckAlign"), System.Drawing.ContentAlignment)
        Me.CheckBoxLockX.Dock = CType(resources.GetObject("CheckBoxLockX.Dock"), System.Windows.Forms.DockStyle)
        Me.CheckBoxLockX.Enabled = CType(resources.GetObject("CheckBoxLockX.Enabled"), Boolean)
        Me.CheckBoxLockX.FlatStyle = CType(resources.GetObject("CheckBoxLockX.FlatStyle"), System.Windows.Forms.FlatStyle)
        Me.CheckBoxLockX.Font = CType(resources.GetObject("CheckBoxLockX.Font"), System.Drawing.Font)
        Me.CheckBoxLockX.Image = CType(resources.GetObject("CheckBoxLockX.Image"), System.Drawing.Image)
        Me.CheckBoxLockX.ImageAlign = CType(resources.GetObject("CheckBoxLockX.ImageAlign"), System.Drawing.ContentAlignment)
        Me.CheckBoxLockX.ImageIndex = CType(resources.GetObject("CheckBoxLockX.ImageIndex"), Integer)
        Me.CheckBoxLockX.ImeMode = CType(resources.GetObject("CheckBoxLockX.ImeMode"), System.Windows.Forms.ImeMode)
        Me.CheckBoxLockX.Location = CType(resources.GetObject("CheckBoxLockX.Location"), System.Drawing.Point)
        Me.CheckBoxLockX.Name = "CheckBoxLockX"
        Me.CheckBoxLockX.RightToLeft = CType(resources.GetObject("CheckBoxLockX.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.CheckBoxLockX.Size = CType(resources.GetObject("CheckBoxLockX.Size"), System.Drawing.Size)
        Me.CheckBoxLockX.TabIndex = CType(resources.GetObject("CheckBoxLockX.TabIndex"), Integer)
        Me.CheckBoxLockX.Text = resources.GetString("CheckBoxLockX.Text")
        Me.CheckBoxLockX.TextAlign = CType(resources.GetObject("CheckBoxLockX.TextAlign"), System.Drawing.ContentAlignment)
        Me.ToolTip1.SetToolTip(Me.CheckBoxLockX, resources.GetString("CheckBoxLockX.ToolTip"))
        Me.CheckBoxLockX.Visible = CType(resources.GetObject("CheckBoxLockX.Visible"), Boolean)
        '
        'TextBoxRobotX
        '
        Me.TextBoxRobotX.AccessibleDescription = resources.GetString("TextBoxRobotX.AccessibleDescription")
        Me.TextBoxRobotX.AccessibleName = resources.GetString("TextBoxRobotX.AccessibleName")
        Me.TextBoxRobotX.Anchor = CType(resources.GetObject("TextBoxRobotX.Anchor"), System.Windows.Forms.AnchorStyles)
        Me.TextBoxRobotX.AutoSize = CType(resources.GetObject("TextBoxRobotX.AutoSize"), Boolean)
        Me.TextBoxRobotX.BackgroundImage = CType(resources.GetObject("TextBoxRobotX.BackgroundImage"), System.Drawing.Image)
        Me.TextBoxRobotX.Dock = CType(resources.GetObject("TextBoxRobotX.Dock"), System.Windows.Forms.DockStyle)
        Me.TextBoxRobotX.Enabled = CType(resources.GetObject("TextBoxRobotX.Enabled"), Boolean)
        Me.TextBoxRobotX.Font = CType(resources.GetObject("TextBoxRobotX.Font"), System.Drawing.Font)
        Me.TextBoxRobotX.ImeMode = CType(resources.GetObject("TextBoxRobotX.ImeMode"), System.Windows.Forms.ImeMode)
        Me.TextBoxRobotX.Location = CType(resources.GetObject("TextBoxRobotX.Location"), System.Drawing.Point)
        Me.TextBoxRobotX.MaxLength = CType(resources.GetObject("TextBoxRobotX.MaxLength"), Integer)
        Me.TextBoxRobotX.Multiline = CType(resources.GetObject("TextBoxRobotX.Multiline"), Boolean)
        Me.TextBoxRobotX.Name = "TextBoxRobotX"
        Me.TextBoxRobotX.PasswordChar = CType(resources.GetObject("TextBoxRobotX.PasswordChar"), Char)
        Me.TextBoxRobotX.ReadOnly = True
        Me.TextBoxRobotX.RightToLeft = CType(resources.GetObject("TextBoxRobotX.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.TextBoxRobotX.ScrollBars = CType(resources.GetObject("TextBoxRobotX.ScrollBars"), System.Windows.Forms.ScrollBars)
        Me.TextBoxRobotX.Size = CType(resources.GetObject("TextBoxRobotX.Size"), System.Drawing.Size)
        Me.TextBoxRobotX.TabIndex = CType(resources.GetObject("TextBoxRobotX.TabIndex"), Integer)
        Me.TextBoxRobotX.Text = resources.GetString("TextBoxRobotX.Text")
        Me.TextBoxRobotX.TextAlign = CType(resources.GetObject("TextBoxRobotX.TextAlign"), System.Windows.Forms.HorizontalAlignment)
        Me.ToolTip1.SetToolTip(Me.TextBoxRobotX, resources.GetString("TextBoxRobotX.ToolTip"))
        Me.TextBoxRobotX.Visible = CType(resources.GetObject("TextBoxRobotX.Visible"), Boolean)
        Me.TextBoxRobotX.WordWrap = CType(resources.GetObject("TextBoxRobotX.WordWrap"), Boolean)
        '
        'lbPostName
        '
        Me.lbPostName.AccessibleDescription = resources.GetString("lbPostName.AccessibleDescription")
        Me.lbPostName.AccessibleName = resources.GetString("lbPostName.AccessibleName")
        Me.lbPostName.Anchor = CType(resources.GetObject("lbPostName.Anchor"), System.Windows.Forms.AnchorStyles)
        Me.lbPostName.AutoSize = CType(resources.GetObject("lbPostName.AutoSize"), Boolean)
        Me.lbPostName.Dock = CType(resources.GetObject("lbPostName.Dock"), System.Windows.Forms.DockStyle)
        Me.lbPostName.Enabled = CType(resources.GetObject("lbPostName.Enabled"), Boolean)
        Me.lbPostName.Font = CType(resources.GetObject("lbPostName.Font"), System.Drawing.Font)
        Me.lbPostName.Image = CType(resources.GetObject("lbPostName.Image"), System.Drawing.Image)
        Me.lbPostName.ImageAlign = CType(resources.GetObject("lbPostName.ImageAlign"), System.Drawing.ContentAlignment)
        Me.lbPostName.ImageIndex = CType(resources.GetObject("lbPostName.ImageIndex"), Integer)
        Me.lbPostName.ImeMode = CType(resources.GetObject("lbPostName.ImeMode"), System.Windows.Forms.ImeMode)
        Me.lbPostName.Location = CType(resources.GetObject("lbPostName.Location"), System.Drawing.Point)
        Me.lbPostName.Name = "lbPostName"
        Me.lbPostName.RightToLeft = CType(resources.GetObject("lbPostName.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.lbPostName.Size = CType(resources.GetObject("lbPostName.Size"), System.Drawing.Size)
        Me.lbPostName.TabIndex = CType(resources.GetObject("lbPostName.TabIndex"), Integer)
        Me.lbPostName.Text = resources.GetString("lbPostName.Text")
        Me.lbPostName.TextAlign = CType(resources.GetObject("lbPostName.TextAlign"), System.Drawing.ContentAlignment)
        Me.ToolTip1.SetToolTip(Me.lbPostName, resources.GetString("lbPostName.ToolTip"))
        Me.lbPostName.Visible = CType(resources.GetObject("lbPostName.Visible"), Boolean)
        '
        'DomainUpDown1
        '
        Me.DomainUpDown1.AccessibleDescription = resources.GetString("DomainUpDown1.AccessibleDescription")
        Me.DomainUpDown1.AccessibleName = resources.GetString("DomainUpDown1.AccessibleName")
        Me.DomainUpDown1.Anchor = CType(resources.GetObject("DomainUpDown1.Anchor"), System.Windows.Forms.AnchorStyles)
        Me.DomainUpDown1.Dock = CType(resources.GetObject("DomainUpDown1.Dock"), System.Windows.Forms.DockStyle)
        Me.DomainUpDown1.Enabled = CType(resources.GetObject("DomainUpDown1.Enabled"), Boolean)
        Me.DomainUpDown1.Font = CType(resources.GetObject("DomainUpDown1.Font"), System.Drawing.Font)
        Me.DomainUpDown1.ImeMode = CType(resources.GetObject("DomainUpDown1.ImeMode"), System.Windows.Forms.ImeMode)
        Me.DomainUpDown1.Items.Add(resources.GetString("resource"))
        Me.DomainUpDown1.Items.Add(resources.GetString("resource1"))
        Me.DomainUpDown1.Items.Add(resources.GetString("resource2"))
        Me.DomainUpDown1.Location = CType(resources.GetObject("DomainUpDown1.Location"), System.Drawing.Point)
        Me.DomainUpDown1.Name = "DomainUpDown1"
        Me.DomainUpDown1.RightToLeft = CType(resources.GetObject("DomainUpDown1.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.DomainUpDown1.Size = CType(resources.GetObject("DomainUpDown1.Size"), System.Drawing.Size)
        Me.DomainUpDown1.TabIndex = CType(resources.GetObject("DomainUpDown1.TabIndex"), Integer)
        Me.DomainUpDown1.Text = resources.GetString("DomainUpDown1.Text")
        Me.DomainUpDown1.TextAlign = CType(resources.GetObject("DomainUpDown1.TextAlign"), System.Windows.Forms.HorizontalAlignment)
        Me.ToolTip1.SetToolTip(Me.DomainUpDown1, resources.GetString("DomainUpDown1.ToolTip"))
        Me.DomainUpDown1.UpDownAlign = CType(resources.GetObject("DomainUpDown1.UpDownAlign"), System.Windows.Forms.LeftRightAlignment)
        Me.DomainUpDown1.Visible = CType(resources.GetObject("DomainUpDown1.Visible"), Boolean)
        Me.DomainUpDown1.Wrap = CType(resources.GetObject("DomainUpDown1.Wrap"), Boolean)
        '
        'Label1
        '
        Me.Label1.AccessibleDescription = resources.GetString("Label1.AccessibleDescription")
        Me.Label1.AccessibleName = resources.GetString("Label1.AccessibleName")
        Me.Label1.Anchor = CType(resources.GetObject("Label1.Anchor"), System.Windows.Forms.AnchorStyles)
        Me.Label1.AutoSize = CType(resources.GetObject("Label1.AutoSize"), Boolean)
        Me.Label1.Dock = CType(resources.GetObject("Label1.Dock"), System.Windows.Forms.DockStyle)
        Me.Label1.Enabled = CType(resources.GetObject("Label1.Enabled"), Boolean)
        Me.Label1.Font = CType(resources.GetObject("Label1.Font"), System.Drawing.Font)
        Me.Label1.Image = CType(resources.GetObject("Label1.Image"), System.Drawing.Image)
        Me.Label1.ImageAlign = CType(resources.GetObject("Label1.ImageAlign"), System.Drawing.ContentAlignment)
        Me.Label1.ImageIndex = CType(resources.GetObject("Label1.ImageIndex"), Integer)
        Me.Label1.ImeMode = CType(resources.GetObject("Label1.ImeMode"), System.Windows.Forms.ImeMode)
        Me.Label1.Location = CType(resources.GetObject("Label1.Location"), System.Drawing.Point)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = CType(resources.GetObject("Label1.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.Label1.Size = CType(resources.GetObject("Label1.Size"), System.Drawing.Size)
        Me.Label1.TabIndex = CType(resources.GetObject("Label1.TabIndex"), Integer)
        Me.Label1.Text = resources.GetString("Label1.Text")
        Me.Label1.TextAlign = CType(resources.GetObject("Label1.TextAlign"), System.Drawing.ContentAlignment)
        Me.ToolTip1.SetToolTip(Me.Label1, resources.GetString("Label1.ToolTip"))
        Me.Label1.Visible = CType(resources.GetObject("Label1.Visible"), Boolean)
        '
        'ButtonVolCal
        '
        Me.ButtonVolCal.AccessibleDescription = resources.GetString("ButtonVolCal.AccessibleDescription")
        Me.ButtonVolCal.AccessibleName = resources.GetString("ButtonVolCal.AccessibleName")
        Me.ButtonVolCal.Anchor = CType(resources.GetObject("ButtonVolCal.Anchor"), System.Windows.Forms.AnchorStyles)
        Me.ButtonVolCal.BackColor = System.Drawing.SystemColors.Control
        Me.ButtonVolCal.BackgroundImage = CType(resources.GetObject("ButtonVolCal.BackgroundImage"), System.Drawing.Image)
        Me.ButtonVolCal.Dock = CType(resources.GetObject("ButtonVolCal.Dock"), System.Windows.Forms.DockStyle)
        Me.ButtonVolCal.Enabled = CType(resources.GetObject("ButtonVolCal.Enabled"), Boolean)
        Me.ButtonVolCal.FlatStyle = CType(resources.GetObject("ButtonVolCal.FlatStyle"), System.Windows.Forms.FlatStyle)
        Me.ButtonVolCal.Font = CType(resources.GetObject("ButtonVolCal.Font"), System.Drawing.Font)
        Me.ButtonVolCal.Image = CType(resources.GetObject("ButtonVolCal.Image"), System.Drawing.Image)
        Me.ButtonVolCal.ImageAlign = CType(resources.GetObject("ButtonVolCal.ImageAlign"), System.Drawing.ContentAlignment)
        Me.ButtonVolCal.ImageIndex = CType(resources.GetObject("ButtonVolCal.ImageIndex"), Integer)
        Me.ButtonVolCal.ImageList = Me.ImageListGeneralTools
        Me.ButtonVolCal.ImeMode = CType(resources.GetObject("ButtonVolCal.ImeMode"), System.Windows.Forms.ImeMode)
        Me.ButtonVolCal.Location = CType(resources.GetObject("ButtonVolCal.Location"), System.Drawing.Point)
        Me.ButtonVolCal.Name = "ButtonVolCal"
        Me.ButtonVolCal.RightToLeft = CType(resources.GetObject("ButtonVolCal.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.ButtonVolCal.Size = CType(resources.GetObject("ButtonVolCal.Size"), System.Drawing.Size)
        Me.ButtonVolCal.TabIndex = CType(resources.GetObject("ButtonVolCal.TabIndex"), Integer)
        Me.ButtonVolCal.Text = resources.GetString("ButtonVolCal.Text")
        Me.ButtonVolCal.TextAlign = CType(resources.GetObject("ButtonVolCal.TextAlign"), System.Drawing.ContentAlignment)
        Me.ToolTip1.SetToolTip(Me.ButtonVolCal, resources.GetString("ButtonVolCal.ToolTip"))
        Me.ButtonVolCal.Visible = CType(resources.GetObject("ButtonVolCal.Visible"), Boolean)
        '
        'ImageListFiles
        '
        Me.ImageListFiles.ImageSize = CType(resources.GetObject("ImageListFiles.ImageSize"), System.Drawing.Size)
        Me.ImageListFiles.ImageStream = CType(resources.GetObject("ImageListFiles.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.ImageListFiles.TransparentColor = System.Drawing.Color.Transparent
        '
        'imageListElement
        '
        Me.imageListElement.ImageSize = CType(resources.GetObject("imageListElement.ImageSize"), System.Drawing.Size)
        Me.imageListElement.ImageStream = CType(resources.GetObject("imageListElement.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.imageListElement.TransparentColor = System.Drawing.Color.DarkGray
        '
        'ImageListReference
        '
        Me.ImageListReference.ImageSize = CType(resources.GetObject("ImageListReference.ImageSize"), System.Drawing.Size)
        Me.ImageListReference.ImageStream = CType(resources.GetObject("ImageListReference.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.ImageListReference.TransparentColor = System.Drawing.Color.White
        '
        'ReferenceCommandBlock
        '
        Me.ReferenceCommandBlock.AccessibleDescription = resources.GetString("ReferenceCommandBlock.AccessibleDescription")
        Me.ReferenceCommandBlock.AccessibleName = resources.GetString("ReferenceCommandBlock.AccessibleName")
        Me.ReferenceCommandBlock.AllowDrop = True
        Me.ReferenceCommandBlock.Anchor = CType(resources.GetObject("ReferenceCommandBlock.Anchor"), System.Windows.Forms.AnchorStyles)
        Me.ReferenceCommandBlock.Appearance = CType(resources.GetObject("ReferenceCommandBlock.Appearance"), System.Windows.Forms.ToolBarAppearance)
        Me.ReferenceCommandBlock.AutoSize = CType(resources.GetObject("ReferenceCommandBlock.AutoSize"), Boolean)
        Me.ReferenceCommandBlock.BackgroundImage = CType(resources.GetObject("ReferenceCommandBlock.BackgroundImage"), System.Drawing.Image)
        Me.ReferenceCommandBlock.Buttons.AddRange(New System.Windows.Forms.ToolBarButton() {Me.TBFiducialPt, Me.TBHeightPt})
        Me.ReferenceCommandBlock.ButtonSize = CType(resources.GetObject("ReferenceCommandBlock.ButtonSize"), System.Drawing.Size)
        Me.ReferenceCommandBlock.Divider = False
        Me.ReferenceCommandBlock.Dock = CType(resources.GetObject("ReferenceCommandBlock.Dock"), System.Windows.Forms.DockStyle)
        Me.ReferenceCommandBlock.DropDownArrows = CType(resources.GetObject("ReferenceCommandBlock.DropDownArrows"), Boolean)
        Me.ReferenceCommandBlock.Enabled = CType(resources.GetObject("ReferenceCommandBlock.Enabled"), Boolean)
        Me.ReferenceCommandBlock.Font = CType(resources.GetObject("ReferenceCommandBlock.Font"), System.Drawing.Font)
        Me.ReferenceCommandBlock.ImageList = Me.ImageListReference
        Me.ReferenceCommandBlock.ImeMode = CType(resources.GetObject("ReferenceCommandBlock.ImeMode"), System.Windows.Forms.ImeMode)
        Me.ReferenceCommandBlock.Location = CType(resources.GetObject("ReferenceCommandBlock.Location"), System.Drawing.Point)
        Me.ReferenceCommandBlock.Name = "ReferenceCommandBlock"
        Me.ReferenceCommandBlock.RightToLeft = CType(resources.GetObject("ReferenceCommandBlock.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.ReferenceCommandBlock.ShowToolTips = CType(resources.GetObject("ReferenceCommandBlock.ShowToolTips"), Boolean)
        Me.ReferenceCommandBlock.Size = CType(resources.GetObject("ReferenceCommandBlock.Size"), System.Drawing.Size)
        Me.ReferenceCommandBlock.TabIndex = CType(resources.GetObject("ReferenceCommandBlock.TabIndex"), Integer)
        Me.ReferenceCommandBlock.TextAlign = CType(resources.GetObject("ReferenceCommandBlock.TextAlign"), System.Windows.Forms.ToolBarTextAlign)
        Me.ToolTip1.SetToolTip(Me.ReferenceCommandBlock, resources.GetString("ReferenceCommandBlock.ToolTip"))
        Me.ReferenceCommandBlock.Visible = CType(resources.GetObject("ReferenceCommandBlock.Visible"), Boolean)
        Me.ReferenceCommandBlock.Wrappable = CType(resources.GetObject("ReferenceCommandBlock.Wrappable"), Boolean)
        '
        'TBFiducialPt
        '
        Me.TBFiducialPt.Enabled = CType(resources.GetObject("TBFiducialPt.Enabled"), Boolean)
        Me.TBFiducialPt.ImageIndex = CType(resources.GetObject("TBFiducialPt.ImageIndex"), Integer)
        Me.TBFiducialPt.Text = resources.GetString("TBFiducialPt.Text")
        Me.TBFiducialPt.ToolTipText = resources.GetString("TBFiducialPt.ToolTipText")
        Me.TBFiducialPt.Visible = CType(resources.GetObject("TBFiducialPt.Visible"), Boolean)
        '
        'TBHeightPt
        '
        Me.TBHeightPt.Enabled = CType(resources.GetObject("TBHeightPt.Enabled"), Boolean)
        Me.TBHeightPt.ImageIndex = CType(resources.GetObject("TBHeightPt.ImageIndex"), Integer)
        Me.TBHeightPt.Text = resources.GetString("TBHeightPt.Text")
        Me.TBHeightPt.ToolTipText = resources.GetString("TBHeightPt.ToolTipText")
        Me.TBHeightPt.Visible = CType(resources.GetObject("TBHeightPt.Visible"), Boolean)
        '
        'ElementsCommandBlock
        '
        Me.ElementsCommandBlock.AccessibleDescription = resources.GetString("ElementsCommandBlock.AccessibleDescription")
        Me.ElementsCommandBlock.AccessibleName = resources.GetString("ElementsCommandBlock.AccessibleName")
        Me.ElementsCommandBlock.Anchor = CType(resources.GetObject("ElementsCommandBlock.Anchor"), System.Windows.Forms.AnchorStyles)
        Me.ElementsCommandBlock.Appearance = CType(resources.GetObject("ElementsCommandBlock.Appearance"), System.Windows.Forms.ToolBarAppearance)
        Me.ElementsCommandBlock.AutoSize = CType(resources.GetObject("ElementsCommandBlock.AutoSize"), Boolean)
        Me.ElementsCommandBlock.BackgroundImage = CType(resources.GetObject("ElementsCommandBlock.BackgroundImage"), System.Drawing.Image)
        Me.ElementsCommandBlock.Buttons.AddRange(New System.Windows.Forms.ToolBarButton() {Me.TBBDot, Me.TBBLine, Me.TBBArc, Me.TBBRectangle, Me.TBBCircle, Me.TBBFilledRectangle, Me.TBBFilledCircle, Me.TBBLink, Me.TBBChipEdge, Me.TBBMove, Me.TBBWait, Me.TBBPurge, Me.TBBClean, Me.TBBQC, Me.TBBSubPattern, Me.TBBArray, Me.TBBGetIO, Me.TBBSetIO, Me.TBBOffset, Me.TBBMeasure, Me.TBBDotArray, Me.TBBVolumeCal})
        Me.ElementsCommandBlock.ButtonSize = CType(resources.GetObject("ElementsCommandBlock.ButtonSize"), System.Drawing.Size)
        Me.ElementsCommandBlock.Divider = False
        Me.ElementsCommandBlock.Dock = CType(resources.GetObject("ElementsCommandBlock.Dock"), System.Windows.Forms.DockStyle)
        Me.ElementsCommandBlock.DropDownArrows = CType(resources.GetObject("ElementsCommandBlock.DropDownArrows"), Boolean)
        Me.ElementsCommandBlock.Enabled = CType(resources.GetObject("ElementsCommandBlock.Enabled"), Boolean)
        Me.ElementsCommandBlock.Font = CType(resources.GetObject("ElementsCommandBlock.Font"), System.Drawing.Font)
        Me.ElementsCommandBlock.ImageList = Me.imageListElement
        Me.ElementsCommandBlock.ImeMode = CType(resources.GetObject("ElementsCommandBlock.ImeMode"), System.Windows.Forms.ImeMode)
        Me.ElementsCommandBlock.Location = CType(resources.GetObject("ElementsCommandBlock.Location"), System.Drawing.Point)
        Me.ElementsCommandBlock.Name = "ElementsCommandBlock"
        Me.ElementsCommandBlock.RightToLeft = CType(resources.GetObject("ElementsCommandBlock.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.ElementsCommandBlock.ShowToolTips = CType(resources.GetObject("ElementsCommandBlock.ShowToolTips"), Boolean)
        Me.ElementsCommandBlock.Size = CType(resources.GetObject("ElementsCommandBlock.Size"), System.Drawing.Size)
        Me.ElementsCommandBlock.TabIndex = CType(resources.GetObject("ElementsCommandBlock.TabIndex"), Integer)
        Me.ElementsCommandBlock.TextAlign = CType(resources.GetObject("ElementsCommandBlock.TextAlign"), System.Windows.Forms.ToolBarTextAlign)
        Me.ToolTip1.SetToolTip(Me.ElementsCommandBlock, resources.GetString("ElementsCommandBlock.ToolTip"))
        Me.ElementsCommandBlock.Visible = CType(resources.GetObject("ElementsCommandBlock.Visible"), Boolean)
        Me.ElementsCommandBlock.Wrappable = CType(resources.GetObject("ElementsCommandBlock.Wrappable"), Boolean)
        '
        'TBBDot
        '
        Me.TBBDot.Enabled = CType(resources.GetObject("TBBDot.Enabled"), Boolean)
        Me.TBBDot.ImageIndex = CType(resources.GetObject("TBBDot.ImageIndex"), Integer)
        Me.TBBDot.Text = resources.GetString("TBBDot.Text")
        Me.TBBDot.ToolTipText = resources.GetString("TBBDot.ToolTipText")
        Me.TBBDot.Visible = CType(resources.GetObject("TBBDot.Visible"), Boolean)
        '
        'TBBLine
        '
        Me.TBBLine.Enabled = CType(resources.GetObject("TBBLine.Enabled"), Boolean)
        Me.TBBLine.ImageIndex = CType(resources.GetObject("TBBLine.ImageIndex"), Integer)
        Me.TBBLine.Text = resources.GetString("TBBLine.Text")
        Me.TBBLine.ToolTipText = resources.GetString("TBBLine.ToolTipText")
        Me.TBBLine.Visible = CType(resources.GetObject("TBBLine.Visible"), Boolean)
        '
        'TBBArc
        '
        Me.TBBArc.Enabled = CType(resources.GetObject("TBBArc.Enabled"), Boolean)
        Me.TBBArc.ImageIndex = CType(resources.GetObject("TBBArc.ImageIndex"), Integer)
        Me.TBBArc.Text = resources.GetString("TBBArc.Text")
        Me.TBBArc.ToolTipText = resources.GetString("TBBArc.ToolTipText")
        Me.TBBArc.Visible = CType(resources.GetObject("TBBArc.Visible"), Boolean)
        '
        'TBBRectangle
        '
        Me.TBBRectangle.Enabled = CType(resources.GetObject("TBBRectangle.Enabled"), Boolean)
        Me.TBBRectangle.ImageIndex = CType(resources.GetObject("TBBRectangle.ImageIndex"), Integer)
        Me.TBBRectangle.Text = resources.GetString("TBBRectangle.Text")
        Me.TBBRectangle.ToolTipText = resources.GetString("TBBRectangle.ToolTipText")
        Me.TBBRectangle.Visible = CType(resources.GetObject("TBBRectangle.Visible"), Boolean)
        '
        'TBBCircle
        '
        Me.TBBCircle.Enabled = CType(resources.GetObject("TBBCircle.Enabled"), Boolean)
        Me.TBBCircle.ImageIndex = CType(resources.GetObject("TBBCircle.ImageIndex"), Integer)
        Me.TBBCircle.Text = resources.GetString("TBBCircle.Text")
        Me.TBBCircle.ToolTipText = resources.GetString("TBBCircle.ToolTipText")
        Me.TBBCircle.Visible = CType(resources.GetObject("TBBCircle.Visible"), Boolean)
        '
        'TBBFilledRectangle
        '
        Me.TBBFilledRectangle.Enabled = CType(resources.GetObject("TBBFilledRectangle.Enabled"), Boolean)
        Me.TBBFilledRectangle.ImageIndex = CType(resources.GetObject("TBBFilledRectangle.ImageIndex"), Integer)
        Me.TBBFilledRectangle.Text = resources.GetString("TBBFilledRectangle.Text")
        Me.TBBFilledRectangle.ToolTipText = resources.GetString("TBBFilledRectangle.ToolTipText")
        Me.TBBFilledRectangle.Visible = CType(resources.GetObject("TBBFilledRectangle.Visible"), Boolean)
        '
        'TBBFilledCircle
        '
        Me.TBBFilledCircle.Enabled = CType(resources.GetObject("TBBFilledCircle.Enabled"), Boolean)
        Me.TBBFilledCircle.ImageIndex = CType(resources.GetObject("TBBFilledCircle.ImageIndex"), Integer)
        Me.TBBFilledCircle.Text = resources.GetString("TBBFilledCircle.Text")
        Me.TBBFilledCircle.ToolTipText = resources.GetString("TBBFilledCircle.ToolTipText")
        Me.TBBFilledCircle.Visible = CType(resources.GetObject("TBBFilledCircle.Visible"), Boolean)
        '
        'TBBLink
        '
        Me.TBBLink.Enabled = CType(resources.GetObject("TBBLink.Enabled"), Boolean)
        Me.TBBLink.ImageIndex = CType(resources.GetObject("TBBLink.ImageIndex"), Integer)
        Me.TBBLink.Text = resources.GetString("TBBLink.Text")
        Me.TBBLink.ToolTipText = resources.GetString("TBBLink.ToolTipText")
        Me.TBBLink.Visible = CType(resources.GetObject("TBBLink.Visible"), Boolean)
        '
        'TBBChipEdge
        '
        Me.TBBChipEdge.Enabled = CType(resources.GetObject("TBBChipEdge.Enabled"), Boolean)
        Me.TBBChipEdge.ImageIndex = CType(resources.GetObject("TBBChipEdge.ImageIndex"), Integer)
        Me.TBBChipEdge.Text = resources.GetString("TBBChipEdge.Text")
        Me.TBBChipEdge.ToolTipText = resources.GetString("TBBChipEdge.ToolTipText")
        Me.TBBChipEdge.Visible = CType(resources.GetObject("TBBChipEdge.Visible"), Boolean)
        '
        'TBBMove
        '
        Me.TBBMove.Enabled = CType(resources.GetObject("TBBMove.Enabled"), Boolean)
        Me.TBBMove.ImageIndex = CType(resources.GetObject("TBBMove.ImageIndex"), Integer)
        Me.TBBMove.Text = resources.GetString("TBBMove.Text")
        Me.TBBMove.ToolTipText = resources.GetString("TBBMove.ToolTipText")
        Me.TBBMove.Visible = CType(resources.GetObject("TBBMove.Visible"), Boolean)
        '
        'TBBWait
        '
        Me.TBBWait.Enabled = CType(resources.GetObject("TBBWait.Enabled"), Boolean)
        Me.TBBWait.ImageIndex = CType(resources.GetObject("TBBWait.ImageIndex"), Integer)
        Me.TBBWait.Text = resources.GetString("TBBWait.Text")
        Me.TBBWait.ToolTipText = resources.GetString("TBBWait.ToolTipText")
        Me.TBBWait.Visible = CType(resources.GetObject("TBBWait.Visible"), Boolean)
        '
        'TBBPurge
        '
        Me.TBBPurge.Enabled = CType(resources.GetObject("TBBPurge.Enabled"), Boolean)
        Me.TBBPurge.ImageIndex = CType(resources.GetObject("TBBPurge.ImageIndex"), Integer)
        Me.TBBPurge.Text = resources.GetString("TBBPurge.Text")
        Me.TBBPurge.ToolTipText = resources.GetString("TBBPurge.ToolTipText")
        Me.TBBPurge.Visible = CType(resources.GetObject("TBBPurge.Visible"), Boolean)
        '
        'TBBClean
        '
        Me.TBBClean.Enabled = CType(resources.GetObject("TBBClean.Enabled"), Boolean)
        Me.TBBClean.ImageIndex = CType(resources.GetObject("TBBClean.ImageIndex"), Integer)
        Me.TBBClean.Text = resources.GetString("TBBClean.Text")
        Me.TBBClean.ToolTipText = resources.GetString("TBBClean.ToolTipText")
        Me.TBBClean.Visible = CType(resources.GetObject("TBBClean.Visible"), Boolean)
        '
        'TBBQC
        '
        Me.TBBQC.Enabled = CType(resources.GetObject("TBBQC.Enabled"), Boolean)
        Me.TBBQC.ImageIndex = CType(resources.GetObject("TBBQC.ImageIndex"), Integer)
        Me.TBBQC.Text = resources.GetString("TBBQC.Text")
        Me.TBBQC.ToolTipText = resources.GetString("TBBQC.ToolTipText")
        Me.TBBQC.Visible = CType(resources.GetObject("TBBQC.Visible"), Boolean)
        '
        'TBBSubPattern
        '
        Me.TBBSubPattern.Enabled = CType(resources.GetObject("TBBSubPattern.Enabled"), Boolean)
        Me.TBBSubPattern.ImageIndex = CType(resources.GetObject("TBBSubPattern.ImageIndex"), Integer)
        Me.TBBSubPattern.Text = resources.GetString("TBBSubPattern.Text")
        Me.TBBSubPattern.ToolTipText = resources.GetString("TBBSubPattern.ToolTipText")
        Me.TBBSubPattern.Visible = CType(resources.GetObject("TBBSubPattern.Visible"), Boolean)
        '
        'TBBArray
        '
        Me.TBBArray.Enabled = CType(resources.GetObject("TBBArray.Enabled"), Boolean)
        Me.TBBArray.ImageIndex = CType(resources.GetObject("TBBArray.ImageIndex"), Integer)
        Me.TBBArray.Text = resources.GetString("TBBArray.Text")
        Me.TBBArray.ToolTipText = resources.GetString("TBBArray.ToolTipText")
        Me.TBBArray.Visible = CType(resources.GetObject("TBBArray.Visible"), Boolean)
        '
        'TBBGetIO
        '
        Me.TBBGetIO.Enabled = CType(resources.GetObject("TBBGetIO.Enabled"), Boolean)
        Me.TBBGetIO.ImageIndex = CType(resources.GetObject("TBBGetIO.ImageIndex"), Integer)
        Me.TBBGetIO.Text = resources.GetString("TBBGetIO.Text")
        Me.TBBGetIO.ToolTipText = resources.GetString("TBBGetIO.ToolTipText")
        Me.TBBGetIO.Visible = CType(resources.GetObject("TBBGetIO.Visible"), Boolean)
        '
        'TBBSetIO
        '
        Me.TBBSetIO.Enabled = CType(resources.GetObject("TBBSetIO.Enabled"), Boolean)
        Me.TBBSetIO.ImageIndex = CType(resources.GetObject("TBBSetIO.ImageIndex"), Integer)
        Me.TBBSetIO.Text = resources.GetString("TBBSetIO.Text")
        Me.TBBSetIO.ToolTipText = resources.GetString("TBBSetIO.ToolTipText")
        Me.TBBSetIO.Visible = CType(resources.GetObject("TBBSetIO.Visible"), Boolean)
        '
        'TBBOffset
        '
        Me.TBBOffset.Enabled = CType(resources.GetObject("TBBOffset.Enabled"), Boolean)
        Me.TBBOffset.ImageIndex = CType(resources.GetObject("TBBOffset.ImageIndex"), Integer)
        Me.TBBOffset.Text = resources.GetString("TBBOffset.Text")
        Me.TBBOffset.ToolTipText = resources.GetString("TBBOffset.ToolTipText")
        Me.TBBOffset.Visible = CType(resources.GetObject("TBBOffset.Visible"), Boolean)
        '
        'TBBMeasure
        '
        Me.TBBMeasure.Enabled = CType(resources.GetObject("TBBMeasure.Enabled"), Boolean)
        Me.TBBMeasure.ImageIndex = CType(resources.GetObject("TBBMeasure.ImageIndex"), Integer)
        Me.TBBMeasure.Text = resources.GetString("TBBMeasure.Text")
        Me.TBBMeasure.ToolTipText = resources.GetString("TBBMeasure.ToolTipText")
        Me.TBBMeasure.Visible = CType(resources.GetObject("TBBMeasure.Visible"), Boolean)
        '
        'TBBDotArray
        '
        Me.TBBDotArray.Enabled = CType(resources.GetObject("TBBDotArray.Enabled"), Boolean)
        Me.TBBDotArray.ImageIndex = CType(resources.GetObject("TBBDotArray.ImageIndex"), Integer)
        Me.TBBDotArray.Text = resources.GetString("TBBDotArray.Text")
        Me.TBBDotArray.ToolTipText = resources.GetString("TBBDotArray.ToolTipText")
        Me.TBBDotArray.Visible = CType(resources.GetObject("TBBDotArray.Visible"), Boolean)
        '
        'TBBVolumeCal
        '
        Me.TBBVolumeCal.Enabled = CType(resources.GetObject("TBBVolumeCal.Enabled"), Boolean)
        Me.TBBVolumeCal.ImageIndex = CType(resources.GetObject("TBBVolumeCal.ImageIndex"), Integer)
        Me.TBBVolumeCal.Text = resources.GetString("TBBVolumeCal.Text")
        Me.TBBVolumeCal.ToolTipText = resources.GetString("TBBVolumeCal.ToolTipText")
        Me.TBBVolumeCal.Visible = CType(resources.GetObject("TBBVolumeCal.Visible"), Boolean)
        '
        'Label5
        '
        Me.Label5.AccessibleDescription = resources.GetString("Label5.AccessibleDescription")
        Me.Label5.AccessibleName = resources.GetString("Label5.AccessibleName")
        Me.Label5.Anchor = CType(resources.GetObject("Label5.Anchor"), System.Windows.Forms.AnchorStyles)
        Me.Label5.AutoSize = CType(resources.GetObject("Label5.AutoSize"), Boolean)
        Me.Label5.Dock = CType(resources.GetObject("Label5.Dock"), System.Windows.Forms.DockStyle)
        Me.Label5.Enabled = CType(resources.GetObject("Label5.Enabled"), Boolean)
        Me.Label5.Font = CType(resources.GetObject("Label5.Font"), System.Drawing.Font)
        Me.Label5.Image = CType(resources.GetObject("Label5.Image"), System.Drawing.Image)
        Me.Label5.ImageAlign = CType(resources.GetObject("Label5.ImageAlign"), System.Drawing.ContentAlignment)
        Me.Label5.ImageIndex = CType(resources.GetObject("Label5.ImageIndex"), Integer)
        Me.Label5.ImeMode = CType(resources.GetObject("Label5.ImeMode"), System.Windows.Forms.ImeMode)
        Me.Label5.Location = CType(resources.GetObject("Label5.Location"), System.Drawing.Point)
        Me.Label5.Name = "Label5"
        Me.Label5.RightToLeft = CType(resources.GetObject("Label5.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.Label5.Size = CType(resources.GetObject("Label5.Size"), System.Drawing.Size)
        Me.Label5.TabIndex = CType(resources.GetObject("Label5.TabIndex"), Integer)
        Me.Label5.Text = resources.GetString("Label5.Text")
        Me.Label5.TextAlign = CType(resources.GetObject("Label5.TextAlign"), System.Drawing.ContentAlignment)
        Me.ToolTip1.SetToolTip(Me.Label5, resources.GetString("Label5.ToolTip"))
        Me.Label5.Visible = CType(resources.GetObject("Label5.Visible"), Boolean)
        '
        'VisionMode
        '
        Me.VisionMode.AccessibleDescription = resources.GetString("VisionMode.AccessibleDescription")
        Me.VisionMode.AccessibleName = resources.GetString("VisionMode.AccessibleName")
        Me.VisionMode.Anchor = CType(resources.GetObject("VisionMode.Anchor"), System.Windows.Forms.AnchorStyles)
        Me.VisionMode.Appearance = CType(resources.GetObject("VisionMode.Appearance"), System.Windows.Forms.Appearance)
        Me.VisionMode.AutoCheck = False
        Me.VisionMode.BackgroundImage = CType(resources.GetObject("VisionMode.BackgroundImage"), System.Drawing.Image)
        Me.VisionMode.CheckAlign = CType(resources.GetObject("VisionMode.CheckAlign"), System.Drawing.ContentAlignment)
        Me.VisionMode.Checked = True
        Me.VisionMode.Dock = CType(resources.GetObject("VisionMode.Dock"), System.Windows.Forms.DockStyle)
        Me.VisionMode.Enabled = CType(resources.GetObject("VisionMode.Enabled"), Boolean)
        Me.VisionMode.FlatStyle = CType(resources.GetObject("VisionMode.FlatStyle"), System.Windows.Forms.FlatStyle)
        Me.VisionMode.Font = CType(resources.GetObject("VisionMode.Font"), System.Drawing.Font)
        Me.VisionMode.Image = CType(resources.GetObject("VisionMode.Image"), System.Drawing.Image)
        Me.VisionMode.ImageAlign = CType(resources.GetObject("VisionMode.ImageAlign"), System.Drawing.ContentAlignment)
        Me.VisionMode.ImageIndex = CType(resources.GetObject("VisionMode.ImageIndex"), Integer)
        Me.VisionMode.ImeMode = CType(resources.GetObject("VisionMode.ImeMode"), System.Windows.Forms.ImeMode)
        Me.VisionMode.Location = CType(resources.GetObject("VisionMode.Location"), System.Drawing.Point)
        Me.VisionMode.Name = "VisionMode"
        Me.VisionMode.RightToLeft = CType(resources.GetObject("VisionMode.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.VisionMode.Size = CType(resources.GetObject("VisionMode.Size"), System.Drawing.Size)
        Me.VisionMode.TabIndex = CType(resources.GetObject("VisionMode.TabIndex"), Integer)
        Me.VisionMode.TabStop = True
        Me.VisionMode.Text = resources.GetString("VisionMode.Text")
        Me.VisionMode.TextAlign = CType(resources.GetObject("VisionMode.TextAlign"), System.Drawing.ContentAlignment)
        Me.ToolTip1.SetToolTip(Me.VisionMode, resources.GetString("VisionMode.ToolTip"))
        Me.VisionMode.Visible = CType(resources.GetObject("VisionMode.Visible"), Boolean)
        '
        'CBExpandSpreadsheet
        '
        Me.CBExpandSpreadsheet.AccessibleDescription = resources.GetString("CBExpandSpreadsheet.AccessibleDescription")
        Me.CBExpandSpreadsheet.AccessibleName = resources.GetString("CBExpandSpreadsheet.AccessibleName")
        Me.CBExpandSpreadsheet.Anchor = CType(resources.GetObject("CBExpandSpreadsheet.Anchor"), System.Windows.Forms.AnchorStyles)
        Me.CBExpandSpreadsheet.Appearance = CType(resources.GetObject("CBExpandSpreadsheet.Appearance"), System.Windows.Forms.Appearance)
        Me.CBExpandSpreadsheet.BackgroundImage = CType(resources.GetObject("CBExpandSpreadsheet.BackgroundImage"), System.Drawing.Image)
        Me.CBExpandSpreadsheet.CheckAlign = CType(resources.GetObject("CBExpandSpreadsheet.CheckAlign"), System.Drawing.ContentAlignment)
        Me.CBExpandSpreadsheet.Dock = CType(resources.GetObject("CBExpandSpreadsheet.Dock"), System.Windows.Forms.DockStyle)
        Me.CBExpandSpreadsheet.Enabled = CType(resources.GetObject("CBExpandSpreadsheet.Enabled"), Boolean)
        Me.CBExpandSpreadsheet.FlatStyle = CType(resources.GetObject("CBExpandSpreadsheet.FlatStyle"), System.Windows.Forms.FlatStyle)
        Me.CBExpandSpreadsheet.Font = CType(resources.GetObject("CBExpandSpreadsheet.Font"), System.Drawing.Font)
        Me.CBExpandSpreadsheet.Image = CType(resources.GetObject("CBExpandSpreadsheet.Image"), System.Drawing.Image)
        Me.CBExpandSpreadsheet.ImageAlign = CType(resources.GetObject("CBExpandSpreadsheet.ImageAlign"), System.Drawing.ContentAlignment)
        Me.CBExpandSpreadsheet.ImageIndex = CType(resources.GetObject("CBExpandSpreadsheet.ImageIndex"), Integer)
        Me.CBExpandSpreadsheet.ImeMode = CType(resources.GetObject("CBExpandSpreadsheet.ImeMode"), System.Windows.Forms.ImeMode)
        Me.CBExpandSpreadsheet.Location = CType(resources.GetObject("CBExpandSpreadsheet.Location"), System.Drawing.Point)
        Me.CBExpandSpreadsheet.Name = "CBExpandSpreadsheet"
        Me.CBExpandSpreadsheet.RightToLeft = CType(resources.GetObject("CBExpandSpreadsheet.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.CBExpandSpreadsheet.Size = CType(resources.GetObject("CBExpandSpreadsheet.Size"), System.Drawing.Size)
        Me.CBExpandSpreadsheet.TabIndex = CType(resources.GetObject("CBExpandSpreadsheet.TabIndex"), Integer)
        Me.CBExpandSpreadsheet.Text = resources.GetString("CBExpandSpreadsheet.Text")
        Me.CBExpandSpreadsheet.TextAlign = CType(resources.GetObject("CBExpandSpreadsheet.TextAlign"), System.Drawing.ContentAlignment)
        Me.ToolTip1.SetToolTip(Me.CBExpandSpreadsheet, resources.GetString("CBExpandSpreadsheet.ToolTip"))
        Me.CBExpandSpreadsheet.Visible = CType(resources.GetObject("CBExpandSpreadsheet.Visible"), Boolean)
        '
        'NeedleMode
        '
        Me.NeedleMode.AccessibleDescription = resources.GetString("NeedleMode.AccessibleDescription")
        Me.NeedleMode.AccessibleName = resources.GetString("NeedleMode.AccessibleName")
        Me.NeedleMode.Anchor = CType(resources.GetObject("NeedleMode.Anchor"), System.Windows.Forms.AnchorStyles)
        Me.NeedleMode.Appearance = CType(resources.GetObject("NeedleMode.Appearance"), System.Windows.Forms.Appearance)
        Me.NeedleMode.AutoCheck = False
        Me.NeedleMode.BackColor = System.Drawing.SystemColors.Control
        Me.NeedleMode.BackgroundImage = CType(resources.GetObject("NeedleMode.BackgroundImage"), System.Drawing.Image)
        Me.NeedleMode.CheckAlign = CType(resources.GetObject("NeedleMode.CheckAlign"), System.Drawing.ContentAlignment)
        Me.NeedleMode.Dock = CType(resources.GetObject("NeedleMode.Dock"), System.Windows.Forms.DockStyle)
        Me.NeedleMode.Enabled = CType(resources.GetObject("NeedleMode.Enabled"), Boolean)
        Me.NeedleMode.FlatStyle = CType(resources.GetObject("NeedleMode.FlatStyle"), System.Windows.Forms.FlatStyle)
        Me.NeedleMode.Font = CType(resources.GetObject("NeedleMode.Font"), System.Drawing.Font)
        Me.NeedleMode.Image = CType(resources.GetObject("NeedleMode.Image"), System.Drawing.Image)
        Me.NeedleMode.ImageAlign = CType(resources.GetObject("NeedleMode.ImageAlign"), System.Drawing.ContentAlignment)
        Me.NeedleMode.ImageIndex = CType(resources.GetObject("NeedleMode.ImageIndex"), Integer)
        Me.NeedleMode.ImeMode = CType(resources.GetObject("NeedleMode.ImeMode"), System.Windows.Forms.ImeMode)
        Me.NeedleMode.Location = CType(resources.GetObject("NeedleMode.Location"), System.Drawing.Point)
        Me.NeedleMode.Name = "NeedleMode"
        Me.NeedleMode.RightToLeft = CType(resources.GetObject("NeedleMode.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.NeedleMode.Size = CType(resources.GetObject("NeedleMode.Size"), System.Drawing.Size)
        Me.NeedleMode.TabIndex = CType(resources.GetObject("NeedleMode.TabIndex"), Integer)
        Me.NeedleMode.Text = resources.GetString("NeedleMode.Text")
        Me.NeedleMode.TextAlign = CType(resources.GetObject("NeedleMode.TextAlign"), System.Drawing.ContentAlignment)
        Me.ToolTip1.SetToolTip(Me.NeedleMode, resources.GetString("NeedleMode.ToolTip"))
        Me.NeedleMode.Visible = CType(resources.GetObject("NeedleMode.Visible"), Boolean)
        '
        'ButtonClean
        '
        Me.ButtonClean.AccessibleDescription = resources.GetString("ButtonClean.AccessibleDescription")
        Me.ButtonClean.AccessibleName = resources.GetString("ButtonClean.AccessibleName")
        Me.ButtonClean.Anchor = CType(resources.GetObject("ButtonClean.Anchor"), System.Windows.Forms.AnchorStyles)
        Me.ButtonClean.BackColor = System.Drawing.SystemColors.Control
        Me.ButtonClean.BackgroundImage = CType(resources.GetObject("ButtonClean.BackgroundImage"), System.Drawing.Image)
        Me.ButtonClean.Dock = CType(resources.GetObject("ButtonClean.Dock"), System.Windows.Forms.DockStyle)
        Me.ButtonClean.Enabled = CType(resources.GetObject("ButtonClean.Enabled"), Boolean)
        Me.ButtonClean.FlatStyle = CType(resources.GetObject("ButtonClean.FlatStyle"), System.Windows.Forms.FlatStyle)
        Me.ButtonClean.Font = CType(resources.GetObject("ButtonClean.Font"), System.Drawing.Font)
        Me.ButtonClean.Image = CType(resources.GetObject("ButtonClean.Image"), System.Drawing.Image)
        Me.ButtonClean.ImageAlign = CType(resources.GetObject("ButtonClean.ImageAlign"), System.Drawing.ContentAlignment)
        Me.ButtonClean.ImageIndex = CType(resources.GetObject("ButtonClean.ImageIndex"), Integer)
        Me.ButtonClean.ImageList = Me.ImageListGeneralTools
        Me.ButtonClean.ImeMode = CType(resources.GetObject("ButtonClean.ImeMode"), System.Windows.Forms.ImeMode)
        Me.ButtonClean.Location = CType(resources.GetObject("ButtonClean.Location"), System.Drawing.Point)
        Me.ButtonClean.Name = "ButtonClean"
        Me.ButtonClean.RightToLeft = CType(resources.GetObject("ButtonClean.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.ButtonClean.Size = CType(resources.GetObject("ButtonClean.Size"), System.Drawing.Size)
        Me.ButtonClean.TabIndex = CType(resources.GetObject("ButtonClean.TabIndex"), Integer)
        Me.ButtonClean.Text = resources.GetString("ButtonClean.Text")
        Me.ButtonClean.TextAlign = CType(resources.GetObject("ButtonClean.TextAlign"), System.Drawing.ContentAlignment)
        Me.ToolTip1.SetToolTip(Me.ButtonClean, resources.GetString("ButtonClean.ToolTip"))
        Me.ButtonClean.Visible = CType(resources.GetObject("ButtonClean.Visible"), Boolean)
        '
        'ButtonNeedleCal
        '
        Me.ButtonNeedleCal.AccessibleDescription = resources.GetString("ButtonNeedleCal.AccessibleDescription")
        Me.ButtonNeedleCal.AccessibleName = resources.GetString("ButtonNeedleCal.AccessibleName")
        Me.ButtonNeedleCal.Anchor = CType(resources.GetObject("ButtonNeedleCal.Anchor"), System.Windows.Forms.AnchorStyles)
        Me.ButtonNeedleCal.BackColor = System.Drawing.SystemColors.Control
        Me.ButtonNeedleCal.BackgroundImage = CType(resources.GetObject("ButtonNeedleCal.BackgroundImage"), System.Drawing.Image)
        Me.ButtonNeedleCal.Dock = CType(resources.GetObject("ButtonNeedleCal.Dock"), System.Windows.Forms.DockStyle)
        Me.ButtonNeedleCal.Enabled = CType(resources.GetObject("ButtonNeedleCal.Enabled"), Boolean)
        Me.ButtonNeedleCal.FlatStyle = CType(resources.GetObject("ButtonNeedleCal.FlatStyle"), System.Windows.Forms.FlatStyle)
        Me.ButtonNeedleCal.Font = CType(resources.GetObject("ButtonNeedleCal.Font"), System.Drawing.Font)
        Me.ButtonNeedleCal.Image = CType(resources.GetObject("ButtonNeedleCal.Image"), System.Drawing.Image)
        Me.ButtonNeedleCal.ImageAlign = CType(resources.GetObject("ButtonNeedleCal.ImageAlign"), System.Drawing.ContentAlignment)
        Me.ButtonNeedleCal.ImageIndex = CType(resources.GetObject("ButtonNeedleCal.ImageIndex"), Integer)
        Me.ButtonNeedleCal.ImageList = Me.ImageListGeneralTools
        Me.ButtonNeedleCal.ImeMode = CType(resources.GetObject("ButtonNeedleCal.ImeMode"), System.Windows.Forms.ImeMode)
        Me.ButtonNeedleCal.Location = CType(resources.GetObject("ButtonNeedleCal.Location"), System.Drawing.Point)
        Me.ButtonNeedleCal.Name = "ButtonNeedleCal"
        Me.ButtonNeedleCal.RightToLeft = CType(resources.GetObject("ButtonNeedleCal.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.ButtonNeedleCal.Size = CType(resources.GetObject("ButtonNeedleCal.Size"), System.Drawing.Size)
        Me.ButtonNeedleCal.TabIndex = CType(resources.GetObject("ButtonNeedleCal.TabIndex"), Integer)
        Me.ButtonNeedleCal.Text = resources.GetString("ButtonNeedleCal.Text")
        Me.ButtonNeedleCal.TextAlign = CType(resources.GetObject("ButtonNeedleCal.TextAlign"), System.Drawing.ContentAlignment)
        Me.ToolTip1.SetToolTip(Me.ButtonNeedleCal, resources.GetString("ButtonNeedleCal.ToolTip"))
        Me.ButtonNeedleCal.Visible = CType(resources.GetObject("ButtonNeedleCal.Visible"), Boolean)
        '
        'ButtonHome
        '
        Me.ButtonHome.AccessibleDescription = resources.GetString("ButtonHome.AccessibleDescription")
        Me.ButtonHome.AccessibleName = resources.GetString("ButtonHome.AccessibleName")
        Me.ButtonHome.Anchor = CType(resources.GetObject("ButtonHome.Anchor"), System.Windows.Forms.AnchorStyles)
        Me.ButtonHome.BackColor = System.Drawing.SystemColors.Control
        Me.ButtonHome.BackgroundImage = CType(resources.GetObject("ButtonHome.BackgroundImage"), System.Drawing.Image)
        Me.ButtonHome.Dock = CType(resources.GetObject("ButtonHome.Dock"), System.Windows.Forms.DockStyle)
        Me.ButtonHome.Enabled = CType(resources.GetObject("ButtonHome.Enabled"), Boolean)
        Me.ButtonHome.FlatStyle = CType(resources.GetObject("ButtonHome.FlatStyle"), System.Windows.Forms.FlatStyle)
        Me.ButtonHome.Font = CType(resources.GetObject("ButtonHome.Font"), System.Drawing.Font)
        Me.ButtonHome.Image = CType(resources.GetObject("ButtonHome.Image"), System.Drawing.Image)
        Me.ButtonHome.ImageAlign = CType(resources.GetObject("ButtonHome.ImageAlign"), System.Drawing.ContentAlignment)
        Me.ButtonHome.ImageIndex = CType(resources.GetObject("ButtonHome.ImageIndex"), Integer)
        Me.ButtonHome.ImageList = Me.ImageListGeneralTools
        Me.ButtonHome.ImeMode = CType(resources.GetObject("ButtonHome.ImeMode"), System.Windows.Forms.ImeMode)
        Me.ButtonHome.Location = CType(resources.GetObject("ButtonHome.Location"), System.Drawing.Point)
        Me.ButtonHome.Name = "ButtonHome"
        Me.ButtonHome.RightToLeft = CType(resources.GetObject("ButtonHome.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.ButtonHome.Size = CType(resources.GetObject("ButtonHome.Size"), System.Drawing.Size)
        Me.ButtonHome.TabIndex = CType(resources.GetObject("ButtonHome.TabIndex"), Integer)
        Me.ButtonHome.Text = resources.GetString("ButtonHome.Text")
        Me.ButtonHome.TextAlign = CType(resources.GetObject("ButtonHome.TextAlign"), System.Drawing.ContentAlignment)
        Me.ToolTip1.SetToolTip(Me.ButtonHome, resources.GetString("ButtonHome.ToolTip"))
        Me.ButtonHome.Visible = CType(resources.GetObject("ButtonHome.Visible"), Boolean)
        '
        'CBDoorLock
        '
        Me.CBDoorLock.AccessibleDescription = resources.GetString("CBDoorLock.AccessibleDescription")
        Me.CBDoorLock.AccessibleName = resources.GetString("CBDoorLock.AccessibleName")
        Me.CBDoorLock.Anchor = CType(resources.GetObject("CBDoorLock.Anchor"), System.Windows.Forms.AnchorStyles)
        Me.CBDoorLock.Appearance = CType(resources.GetObject("CBDoorLock.Appearance"), System.Windows.Forms.Appearance)
        Me.CBDoorLock.BackColor = System.Drawing.SystemColors.Control
        Me.CBDoorLock.BackgroundImage = CType(resources.GetObject("CBDoorLock.BackgroundImage"), System.Drawing.Image)
        Me.CBDoorLock.CheckAlign = CType(resources.GetObject("CBDoorLock.CheckAlign"), System.Drawing.ContentAlignment)
        Me.CBDoorLock.Dock = CType(resources.GetObject("CBDoorLock.Dock"), System.Windows.Forms.DockStyle)
        Me.CBDoorLock.Enabled = CType(resources.GetObject("CBDoorLock.Enabled"), Boolean)
        Me.CBDoorLock.FlatStyle = CType(resources.GetObject("CBDoorLock.FlatStyle"), System.Windows.Forms.FlatStyle)
        Me.CBDoorLock.Font = CType(resources.GetObject("CBDoorLock.Font"), System.Drawing.Font)
        Me.CBDoorLock.Image = CType(resources.GetObject("CBDoorLock.Image"), System.Drawing.Image)
        Me.CBDoorLock.ImageAlign = CType(resources.GetObject("CBDoorLock.ImageAlign"), System.Drawing.ContentAlignment)
        Me.CBDoorLock.ImageIndex = CType(resources.GetObject("CBDoorLock.ImageIndex"), Integer)
        Me.CBDoorLock.ImageList = Me.ImageListMultiField
        Me.CBDoorLock.ImeMode = CType(resources.GetObject("CBDoorLock.ImeMode"), System.Windows.Forms.ImeMode)
        Me.CBDoorLock.Location = CType(resources.GetObject("CBDoorLock.Location"), System.Drawing.Point)
        Me.CBDoorLock.Name = "CBDoorLock"
        Me.CBDoorLock.RightToLeft = CType(resources.GetObject("CBDoorLock.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.CBDoorLock.Size = CType(resources.GetObject("CBDoorLock.Size"), System.Drawing.Size)
        Me.CBDoorLock.TabIndex = CType(resources.GetObject("CBDoorLock.TabIndex"), Integer)
        Me.CBDoorLock.Text = resources.GetString("CBDoorLock.Text")
        Me.CBDoorLock.TextAlign = CType(resources.GetObject("CBDoorLock.TextAlign"), System.Drawing.ContentAlignment)
        Me.ToolTip1.SetToolTip(Me.CBDoorLock, resources.GetString("CBDoorLock.ToolTip"))
        Me.CBDoorLock.Visible = CType(resources.GetObject("CBDoorLock.Visible"), Boolean)
        '
        'ImageListMultiField
        '
        Me.ImageListMultiField.ImageSize = CType(resources.GetObject("ImageListMultiField.ImageSize"), System.Drawing.Size)
        Me.ImageListMultiField.ImageStream = CType(resources.GetObject("ImageListMultiField.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.ImageListMultiField.TransparentColor = System.Drawing.Color.White
        '
        'AxSpreadsheetProgramming
        '
        Me.AxSpreadsheetProgramming.AccessibleDescription = resources.GetString("AxSpreadsheetProgramming.AccessibleDescription")
        Me.AxSpreadsheetProgramming.AccessibleName = resources.GetString("AxSpreadsheetProgramming.AccessibleName")
        Me.AxSpreadsheetProgramming.Anchor = CType(resources.GetObject("AxSpreadsheetProgramming.Anchor"), System.Windows.Forms.AnchorStyles)
        Me.AxSpreadsheetProgramming.BackgroundImage = CType(resources.GetObject("AxSpreadsheetProgramming.BackgroundImage"), System.Drawing.Image)
        Me.AxSpreadsheetProgramming.DataSource = Nothing
        Me.AxSpreadsheetProgramming.Dock = CType(resources.GetObject("AxSpreadsheetProgramming.Dock"), System.Windows.Forms.DockStyle)
        Me.AxSpreadsheetProgramming.Enabled = CType(resources.GetObject("AxSpreadsheetProgramming.Enabled"), Boolean)
        Me.AxSpreadsheetProgramming.Font = CType(resources.GetObject("AxSpreadsheetProgramming.Font"), System.Drawing.Font)
        Me.AxSpreadsheetProgramming.ImeMode = CType(resources.GetObject("AxSpreadsheetProgramming.ImeMode"), System.Windows.Forms.ImeMode)
        Me.AxSpreadsheetProgramming.Location = CType(resources.GetObject("AxSpreadsheetProgramming.Location"), System.Drawing.Point)
        Me.AxSpreadsheetProgramming.Name = "AxSpreadsheetProgramming"
        Me.AxSpreadsheetProgramming.OcxState = CType(resources.GetObject("AxSpreadsheetProgramming.OcxState"), System.Windows.Forms.AxHost.State)
        Me.AxSpreadsheetProgramming.RightToLeft = CType(resources.GetObject("AxSpreadsheetProgramming.RightToLeft"), Boolean)
        Me.AxSpreadsheetProgramming.Size = CType(resources.GetObject("AxSpreadsheetProgramming.Size"), System.Drawing.Size)
        Me.AxSpreadsheetProgramming.TabIndex = CType(resources.GetObject("AxSpreadsheetProgramming.TabIndex"), Integer)
        Me.AxSpreadsheetProgramming.Text = resources.GetString("AxSpreadsheetProgramming.Text")
        Me.ToolTip1.SetToolTip(Me.AxSpreadsheetProgramming, resources.GetString("AxSpreadsheetProgramming.ToolTip"))
        Me.AxSpreadsheetProgramming.Visible = CType(resources.GetObject("AxSpreadsheetProgramming.Visible"), Boolean)
        '
        'LabelMessege
        '
        Me.LabelMessege.AccessibleDescription = resources.GetString("LabelMessege.AccessibleDescription")
        Me.LabelMessege.AccessibleName = resources.GetString("LabelMessege.AccessibleName")
        Me.LabelMessege.Anchor = CType(resources.GetObject("LabelMessege.Anchor"), System.Windows.Forms.AnchorStyles)
        Me.LabelMessege.AutoSize = CType(resources.GetObject("LabelMessege.AutoSize"), Boolean)
        Me.LabelMessege.BackColor = System.Drawing.SystemColors.Control
        Me.LabelMessege.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.LabelMessege.Dock = CType(resources.GetObject("LabelMessege.Dock"), System.Windows.Forms.DockStyle)
        Me.LabelMessege.Enabled = CType(resources.GetObject("LabelMessege.Enabled"), Boolean)
        Me.LabelMessege.Font = CType(resources.GetObject("LabelMessege.Font"), System.Drawing.Font)
        Me.LabelMessege.ForeColor = System.Drawing.Color.Black
        Me.LabelMessege.Image = CType(resources.GetObject("LabelMessege.Image"), System.Drawing.Image)
        Me.LabelMessege.ImageAlign = CType(resources.GetObject("LabelMessege.ImageAlign"), System.Drawing.ContentAlignment)
        Me.LabelMessege.ImageIndex = CType(resources.GetObject("LabelMessege.ImageIndex"), Integer)
        Me.LabelMessege.ImeMode = CType(resources.GetObject("LabelMessege.ImeMode"), System.Windows.Forms.ImeMode)
        Me.LabelMessege.Location = CType(resources.GetObject("LabelMessege.Location"), System.Drawing.Point)
        Me.LabelMessege.Name = "LabelMessege"
        Me.LabelMessege.RightToLeft = CType(resources.GetObject("LabelMessege.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.LabelMessege.Size = CType(resources.GetObject("LabelMessege.Size"), System.Drawing.Size)
        Me.LabelMessege.TabIndex = CType(resources.GetObject("LabelMessege.TabIndex"), Integer)
        Me.LabelMessege.Text = resources.GetString("LabelMessege.Text")
        Me.LabelMessege.TextAlign = CType(resources.GetObject("LabelMessege.TextAlign"), System.Drawing.ContentAlignment)
        Me.ToolTip1.SetToolTip(Me.LabelMessege, resources.GetString("LabelMessege.ToolTip"))
        Me.LabelMessege.Visible = CType(resources.GetObject("LabelMessege.Visible"), Boolean)
        '
        'DispensingMode
        '
        Me.DispensingMode.AccessibleDescription = resources.GetString("DispensingMode.AccessibleDescription")
        Me.DispensingMode.AccessibleName = resources.GetString("DispensingMode.AccessibleName")
        Me.DispensingMode.Anchor = CType(resources.GetObject("DispensingMode.Anchor"), System.Windows.Forms.AnchorStyles)
        Me.DispensingMode.BackgroundImage = CType(resources.GetObject("DispensingMode.BackgroundImage"), System.Drawing.Image)
        Me.DispensingMode.Dock = CType(resources.GetObject("DispensingMode.Dock"), System.Windows.Forms.DockStyle)
        Me.DispensingMode.Enabled = CType(resources.GetObject("DispensingMode.Enabled"), Boolean)
        Me.DispensingMode.Font = CType(resources.GetObject("DispensingMode.Font"), System.Drawing.Font)
        Me.DispensingMode.ImeMode = CType(resources.GetObject("DispensingMode.ImeMode"), System.Windows.Forms.ImeMode)
        Me.DispensingMode.IntegralHeight = CType(resources.GetObject("DispensingMode.IntegralHeight"), Boolean)
        Me.DispensingMode.ItemHeight = CType(resources.GetObject("DispensingMode.ItemHeight"), Integer)
        Me.DispensingMode.Items.AddRange(New Object() {resources.GetString("DispensingMode.Items"), resources.GetString("DispensingMode.Items1"), resources.GetString("DispensingMode.Items2")})
        Me.DispensingMode.Location = CType(resources.GetObject("DispensingMode.Location"), System.Drawing.Point)
        Me.DispensingMode.MaxDropDownItems = CType(resources.GetObject("DispensingMode.MaxDropDownItems"), Integer)
        Me.DispensingMode.MaxLength = CType(resources.GetObject("DispensingMode.MaxLength"), Integer)
        Me.DispensingMode.Name = "DispensingMode"
        Me.DispensingMode.RightToLeft = CType(resources.GetObject("DispensingMode.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.DispensingMode.Size = CType(resources.GetObject("DispensingMode.Size"), System.Drawing.Size)
        Me.DispensingMode.TabIndex = CType(resources.GetObject("DispensingMode.TabIndex"), Integer)
        Me.DispensingMode.Text = resources.GetString("DispensingMode.Text")
        Me.ToolTip1.SetToolTip(Me.DispensingMode, resources.GetString("DispensingMode.ToolTip"))
        Me.DispensingMode.Visible = CType(resources.GetObject("DispensingMode.Visible"), Boolean)
        '
        'btStop
        '
        Me.btStop.AccessibleDescription = resources.GetString("btStop.AccessibleDescription")
        Me.btStop.AccessibleName = resources.GetString("btStop.AccessibleName")
        Me.btStop.Anchor = CType(resources.GetObject("btStop.Anchor"), System.Windows.Forms.AnchorStyles)
        Me.btStop.BackgroundImage = CType(resources.GetObject("btStop.BackgroundImage"), System.Drawing.Image)
        Me.btStop.Dock = CType(resources.GetObject("btStop.Dock"), System.Windows.Forms.DockStyle)
        Me.btStop.Enabled = CType(resources.GetObject("btStop.Enabled"), Boolean)
        Me.btStop.FlatStyle = CType(resources.GetObject("btStop.FlatStyle"), System.Windows.Forms.FlatStyle)
        Me.btStop.Font = CType(resources.GetObject("btStop.Font"), System.Drawing.Font)
        Me.btStop.Image = CType(resources.GetObject("btStop.Image"), System.Drawing.Image)
        Me.btStop.ImageAlign = CType(resources.GetObject("btStop.ImageAlign"), System.Drawing.ContentAlignment)
        Me.btStop.ImageIndex = CType(resources.GetObject("btStop.ImageIndex"), Integer)
        Me.btStop.ImeMode = CType(resources.GetObject("btStop.ImeMode"), System.Windows.Forms.ImeMode)
        Me.btStop.Location = CType(resources.GetObject("btStop.Location"), System.Drawing.Point)
        Me.btStop.Name = "btStop"
        Me.btStop.RightToLeft = CType(resources.GetObject("btStop.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.btStop.Size = CType(resources.GetObject("btStop.Size"), System.Drawing.Size)
        Me.btStop.TabIndex = CType(resources.GetObject("btStop.TabIndex"), Integer)
        Me.btStop.Text = resources.GetString("btStop.Text")
        Me.btStop.TextAlign = CType(resources.GetObject("btStop.TextAlign"), System.Drawing.ContentAlignment)
        Me.ToolTip1.SetToolTip(Me.btStop, resources.GetString("btStop.ToolTip"))
        Me.btStop.Visible = CType(resources.GetObject("btStop.Visible"), Boolean)
        '
        'btPause
        '
        Me.btPause.AccessibleDescription = resources.GetString("btPause.AccessibleDescription")
        Me.btPause.AccessibleName = resources.GetString("btPause.AccessibleName")
        Me.btPause.Anchor = CType(resources.GetObject("btPause.Anchor"), System.Windows.Forms.AnchorStyles)
        Me.btPause.BackgroundImage = CType(resources.GetObject("btPause.BackgroundImage"), System.Drawing.Image)
        Me.btPause.Dock = CType(resources.GetObject("btPause.Dock"), System.Windows.Forms.DockStyle)
        Me.btPause.Enabled = CType(resources.GetObject("btPause.Enabled"), Boolean)
        Me.btPause.FlatStyle = CType(resources.GetObject("btPause.FlatStyle"), System.Windows.Forms.FlatStyle)
        Me.btPause.Font = CType(resources.GetObject("btPause.Font"), System.Drawing.Font)
        Me.btPause.Image = CType(resources.GetObject("btPause.Image"), System.Drawing.Image)
        Me.btPause.ImageAlign = CType(resources.GetObject("btPause.ImageAlign"), System.Drawing.ContentAlignment)
        Me.btPause.ImageIndex = CType(resources.GetObject("btPause.ImageIndex"), Integer)
        Me.btPause.ImeMode = CType(resources.GetObject("btPause.ImeMode"), System.Windows.Forms.ImeMode)
        Me.btPause.Location = CType(resources.GetObject("btPause.Location"), System.Drawing.Point)
        Me.btPause.Name = "btPause"
        Me.btPause.RightToLeft = CType(resources.GetObject("btPause.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.btPause.Size = CType(resources.GetObject("btPause.Size"), System.Drawing.Size)
        Me.btPause.TabIndex = CType(resources.GetObject("btPause.TabIndex"), Integer)
        Me.btPause.Text = resources.GetString("btPause.Text")
        Me.btPause.TextAlign = CType(resources.GetObject("btPause.TextAlign"), System.Drawing.ContentAlignment)
        Me.ToolTip1.SetToolTip(Me.btPause, resources.GetString("btPause.ToolTip"))
        Me.btPause.Visible = CType(resources.GetObject("btPause.Visible"), Boolean)
        '
        'btPlay
        '
        Me.btPlay.AccessibleDescription = resources.GetString("btPlay.AccessibleDescription")
        Me.btPlay.AccessibleName = resources.GetString("btPlay.AccessibleName")
        Me.btPlay.Anchor = CType(resources.GetObject("btPlay.Anchor"), System.Windows.Forms.AnchorStyles)
        Me.btPlay.BackgroundImage = CType(resources.GetObject("btPlay.BackgroundImage"), System.Drawing.Image)
        Me.btPlay.Dock = CType(resources.GetObject("btPlay.Dock"), System.Windows.Forms.DockStyle)
        Me.btPlay.Enabled = CType(resources.GetObject("btPlay.Enabled"), Boolean)
        Me.btPlay.FlatStyle = CType(resources.GetObject("btPlay.FlatStyle"), System.Windows.Forms.FlatStyle)
        Me.btPlay.Font = CType(resources.GetObject("btPlay.Font"), System.Drawing.Font)
        Me.btPlay.Image = CType(resources.GetObject("btPlay.Image"), System.Drawing.Image)
        Me.btPlay.ImageAlign = CType(resources.GetObject("btPlay.ImageAlign"), System.Drawing.ContentAlignment)
        Me.btPlay.ImageIndex = CType(resources.GetObject("btPlay.ImageIndex"), Integer)
        Me.btPlay.ImeMode = CType(resources.GetObject("btPlay.ImeMode"), System.Windows.Forms.ImeMode)
        Me.btPlay.Location = CType(resources.GetObject("btPlay.Location"), System.Drawing.Point)
        Me.btPlay.Name = "btPlay"
        Me.btPlay.RightToLeft = CType(resources.GetObject("btPlay.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.btPlay.Size = CType(resources.GetObject("btPlay.Size"), System.Drawing.Size)
        Me.btPlay.TabIndex = CType(resources.GetObject("btPlay.TabIndex"), Integer)
        Me.btPlay.Text = resources.GetString("btPlay.Text")
        Me.btPlay.TextAlign = CType(resources.GetObject("btPlay.TextAlign"), System.Drawing.ContentAlignment)
        Me.ToolTip1.SetToolTip(Me.btPlay, resources.GetString("btPlay.ToolTip"))
        Me.btPlay.Visible = CType(resources.GetObject("btPlay.Visible"), Boolean)
        '
        'Panel1
        '
        Me.Panel1.AccessibleDescription = resources.GetString("Panel1.AccessibleDescription")
        Me.Panel1.AccessibleName = resources.GetString("Panel1.AccessibleName")
        Me.Panel1.Anchor = CType(resources.GetObject("Panel1.Anchor"), System.Windows.Forms.AnchorStyles)
        Me.Panel1.AutoScroll = CType(resources.GetObject("Panel1.AutoScroll"), Boolean)
        Me.Panel1.AutoScrollMargin = CType(resources.GetObject("Panel1.AutoScrollMargin"), System.Drawing.Size)
        Me.Panel1.AutoScrollMinSize = CType(resources.GetObject("Panel1.AutoScrollMinSize"), System.Drawing.Size)
        Me.Panel1.BackgroundImage = CType(resources.GetObject("Panel1.BackgroundImage"), System.Drawing.Image)
        Me.Panel1.Controls.Add(Me.GreenLight)
        Me.Panel1.Controls.Add(Me.AmberLight)
        Me.Panel1.Controls.Add(Me.RedLight)
        Me.Panel1.Dock = CType(resources.GetObject("Panel1.Dock"), System.Windows.Forms.DockStyle)
        Me.Panel1.Enabled = CType(resources.GetObject("Panel1.Enabled"), Boolean)
        Me.Panel1.Font = CType(resources.GetObject("Panel1.Font"), System.Drawing.Font)
        Me.Panel1.ImeMode = CType(resources.GetObject("Panel1.ImeMode"), System.Windows.Forms.ImeMode)
        Me.Panel1.Location = CType(resources.GetObject("Panel1.Location"), System.Drawing.Point)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.RightToLeft = CType(resources.GetObject("Panel1.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.Panel1.Size = CType(resources.GetObject("Panel1.Size"), System.Drawing.Size)
        Me.Panel1.TabIndex = CType(resources.GetObject("Panel1.TabIndex"), Integer)
        Me.Panel1.Text = resources.GetString("Panel1.Text")
        Me.ToolTip1.SetToolTip(Me.Panel1, resources.GetString("Panel1.ToolTip"))
        Me.Panel1.Visible = CType(resources.GetObject("Panel1.Visible"), Boolean)
        '
        'GreenLight
        '
        Me.GreenLight.AccessibleDescription = resources.GetString("GreenLight.AccessibleDescription")
        Me.GreenLight.AccessibleName = resources.GetString("GreenLight.AccessibleName")
        Me.GreenLight.Anchor = CType(resources.GetObject("GreenLight.Anchor"), System.Windows.Forms.AnchorStyles)
        Me.GreenLight.BackgroundImage = CType(resources.GetObject("GreenLight.BackgroundImage"), System.Drawing.Image)
        Me.GreenLight.Dock = CType(resources.GetObject("GreenLight.Dock"), System.Windows.Forms.DockStyle)
        Me.GreenLight.Enabled = CType(resources.GetObject("GreenLight.Enabled"), Boolean)
        Me.GreenLight.Font = CType(resources.GetObject("GreenLight.Font"), System.Drawing.Font)
        Me.GreenLight.Image = CType(resources.GetObject("GreenLight.Image"), System.Drawing.Image)
        Me.GreenLight.ImeMode = CType(resources.GetObject("GreenLight.ImeMode"), System.Windows.Forms.ImeMode)
        Me.GreenLight.Location = CType(resources.GetObject("GreenLight.Location"), System.Drawing.Point)
        Me.GreenLight.Name = "GreenLight"
        Me.GreenLight.RightToLeft = CType(resources.GetObject("GreenLight.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.GreenLight.Size = CType(resources.GetObject("GreenLight.Size"), System.Drawing.Size)
        Me.GreenLight.SizeMode = CType(resources.GetObject("GreenLight.SizeMode"), System.Windows.Forms.PictureBoxSizeMode)
        Me.GreenLight.TabIndex = CType(resources.GetObject("GreenLight.TabIndex"), Integer)
        Me.GreenLight.TabStop = False
        Me.GreenLight.Text = resources.GetString("GreenLight.Text")
        Me.ToolTip1.SetToolTip(Me.GreenLight, resources.GetString("GreenLight.ToolTip"))
        Me.GreenLight.Visible = CType(resources.GetObject("GreenLight.Visible"), Boolean)
        '
        'AmberLight
        '
        Me.AmberLight.AccessibleDescription = resources.GetString("AmberLight.AccessibleDescription")
        Me.AmberLight.AccessibleName = resources.GetString("AmberLight.AccessibleName")
        Me.AmberLight.Anchor = CType(resources.GetObject("AmberLight.Anchor"), System.Windows.Forms.AnchorStyles)
        Me.AmberLight.BackgroundImage = CType(resources.GetObject("AmberLight.BackgroundImage"), System.Drawing.Image)
        Me.AmberLight.Dock = CType(resources.GetObject("AmberLight.Dock"), System.Windows.Forms.DockStyle)
        Me.AmberLight.Enabled = CType(resources.GetObject("AmberLight.Enabled"), Boolean)
        Me.AmberLight.Font = CType(resources.GetObject("AmberLight.Font"), System.Drawing.Font)
        Me.AmberLight.Image = CType(resources.GetObject("AmberLight.Image"), System.Drawing.Image)
        Me.AmberLight.ImeMode = CType(resources.GetObject("AmberLight.ImeMode"), System.Windows.Forms.ImeMode)
        Me.AmberLight.Location = CType(resources.GetObject("AmberLight.Location"), System.Drawing.Point)
        Me.AmberLight.Name = "AmberLight"
        Me.AmberLight.RightToLeft = CType(resources.GetObject("AmberLight.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.AmberLight.Size = CType(resources.GetObject("AmberLight.Size"), System.Drawing.Size)
        Me.AmberLight.SizeMode = CType(resources.GetObject("AmberLight.SizeMode"), System.Windows.Forms.PictureBoxSizeMode)
        Me.AmberLight.TabIndex = CType(resources.GetObject("AmberLight.TabIndex"), Integer)
        Me.AmberLight.TabStop = False
        Me.AmberLight.Text = resources.GetString("AmberLight.Text")
        Me.ToolTip1.SetToolTip(Me.AmberLight, resources.GetString("AmberLight.ToolTip"))
        Me.AmberLight.Visible = CType(resources.GetObject("AmberLight.Visible"), Boolean)
        '
        'RedLight
        '
        Me.RedLight.AccessibleDescription = resources.GetString("RedLight.AccessibleDescription")
        Me.RedLight.AccessibleName = resources.GetString("RedLight.AccessibleName")
        Me.RedLight.Anchor = CType(resources.GetObject("RedLight.Anchor"), System.Windows.Forms.AnchorStyles)
        Me.RedLight.BackgroundImage = CType(resources.GetObject("RedLight.BackgroundImage"), System.Drawing.Image)
        Me.RedLight.Dock = CType(resources.GetObject("RedLight.Dock"), System.Windows.Forms.DockStyle)
        Me.RedLight.Enabled = CType(resources.GetObject("RedLight.Enabled"), Boolean)
        Me.RedLight.Font = CType(resources.GetObject("RedLight.Font"), System.Drawing.Font)
        Me.RedLight.Image = CType(resources.GetObject("RedLight.Image"), System.Drawing.Image)
        Me.RedLight.ImeMode = CType(resources.GetObject("RedLight.ImeMode"), System.Windows.Forms.ImeMode)
        Me.RedLight.Location = CType(resources.GetObject("RedLight.Location"), System.Drawing.Point)
        Me.RedLight.Name = "RedLight"
        Me.RedLight.RightToLeft = CType(resources.GetObject("RedLight.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.RedLight.Size = CType(resources.GetObject("RedLight.Size"), System.Drawing.Size)
        Me.RedLight.SizeMode = CType(resources.GetObject("RedLight.SizeMode"), System.Windows.Forms.PictureBoxSizeMode)
        Me.RedLight.TabIndex = CType(resources.GetObject("RedLight.TabIndex"), Integer)
        Me.RedLight.TabStop = False
        Me.RedLight.Text = resources.GetString("RedLight.Text")
        Me.ToolTip1.SetToolTip(Me.RedLight, resources.GetString("RedLight.ToolTip"))
        Me.RedLight.Visible = CType(resources.GetObject("RedLight.Visible"), Boolean)
        '
        'ButtonToggleMode
        '
        Me.ButtonToggleMode.AccessibleDescription = resources.GetString("ButtonToggleMode.AccessibleDescription")
        Me.ButtonToggleMode.AccessibleName = resources.GetString("ButtonToggleMode.AccessibleName")
        Me.ButtonToggleMode.Anchor = CType(resources.GetObject("ButtonToggleMode.Anchor"), System.Windows.Forms.AnchorStyles)
        Me.ButtonToggleMode.BackgroundImage = CType(resources.GetObject("ButtonToggleMode.BackgroundImage"), System.Drawing.Image)
        Me.ButtonToggleMode.Dock = CType(resources.GetObject("ButtonToggleMode.Dock"), System.Windows.Forms.DockStyle)
        Me.ButtonToggleMode.Enabled = CType(resources.GetObject("ButtonToggleMode.Enabled"), Boolean)
        Me.ButtonToggleMode.FlatStyle = CType(resources.GetObject("ButtonToggleMode.FlatStyle"), System.Windows.Forms.FlatStyle)
        Me.ButtonToggleMode.Font = CType(resources.GetObject("ButtonToggleMode.Font"), System.Drawing.Font)
        Me.ButtonToggleMode.Image = CType(resources.GetObject("ButtonToggleMode.Image"), System.Drawing.Image)
        Me.ButtonToggleMode.ImageAlign = CType(resources.GetObject("ButtonToggleMode.ImageAlign"), System.Drawing.ContentAlignment)
        Me.ButtonToggleMode.ImageIndex = CType(resources.GetObject("ButtonToggleMode.ImageIndex"), Integer)
        Me.ButtonToggleMode.ImeMode = CType(resources.GetObject("ButtonToggleMode.ImeMode"), System.Windows.Forms.ImeMode)
        Me.ButtonToggleMode.Location = CType(resources.GetObject("ButtonToggleMode.Location"), System.Drawing.Point)
        Me.ButtonToggleMode.Name = "ButtonToggleMode"
        Me.ButtonToggleMode.RightToLeft = CType(resources.GetObject("ButtonToggleMode.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.ButtonToggleMode.Size = CType(resources.GetObject("ButtonToggleMode.Size"), System.Drawing.Size)
        Me.ButtonToggleMode.TabIndex = CType(resources.GetObject("ButtonToggleMode.TabIndex"), Integer)
        Me.ButtonToggleMode.Text = resources.GetString("ButtonToggleMode.Text")
        Me.ButtonToggleMode.TextAlign = CType(resources.GetObject("ButtonToggleMode.TextAlign"), System.Drawing.ContentAlignment)
        Me.ToolTip1.SetToolTip(Me.ButtonToggleMode, resources.GetString("ButtonToggleMode.ToolTip"))
        Me.ButtonToggleMode.Visible = CType(resources.GetObject("ButtonToggleMode.Visible"), Boolean)
        '
        'PanelToBeAdded
        '
        Me.PanelToBeAdded.AccessibleDescription = resources.GetString("PanelToBeAdded.AccessibleDescription")
        Me.PanelToBeAdded.AccessibleName = resources.GetString("PanelToBeAdded.AccessibleName")
        Me.PanelToBeAdded.Anchor = CType(resources.GetObject("PanelToBeAdded.Anchor"), System.Windows.Forms.AnchorStyles)
        Me.PanelToBeAdded.AutoScroll = CType(resources.GetObject("PanelToBeAdded.AutoScroll"), Boolean)
        Me.PanelToBeAdded.AutoScrollMargin = CType(resources.GetObject("PanelToBeAdded.AutoScrollMargin"), System.Drawing.Size)
        Me.PanelToBeAdded.AutoScrollMinSize = CType(resources.GetObject("PanelToBeAdded.AutoScrollMinSize"), System.Drawing.Size)
        Me.PanelToBeAdded.BackColor = System.Drawing.SystemColors.Control
        Me.PanelToBeAdded.BackgroundImage = CType(resources.GetObject("PanelToBeAdded.BackgroundImage"), System.Drawing.Image)
        Me.PanelToBeAdded.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.PanelToBeAdded.Dock = CType(resources.GetObject("PanelToBeAdded.Dock"), System.Windows.Forms.DockStyle)
        Me.PanelToBeAdded.Enabled = CType(resources.GetObject("PanelToBeAdded.Enabled"), Boolean)
        Me.PanelToBeAdded.Font = CType(resources.GetObject("PanelToBeAdded.Font"), System.Drawing.Font)
        Me.PanelToBeAdded.ImeMode = CType(resources.GetObject("PanelToBeAdded.ImeMode"), System.Windows.Forms.ImeMode)
        Me.PanelToBeAdded.Location = CType(resources.GetObject("PanelToBeAdded.Location"), System.Drawing.Point)
        Me.PanelToBeAdded.Name = "PanelToBeAdded"
        Me.PanelToBeAdded.RightToLeft = CType(resources.GetObject("PanelToBeAdded.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.PanelToBeAdded.Size = CType(resources.GetObject("PanelToBeAdded.Size"), System.Drawing.Size)
        Me.PanelToBeAdded.TabIndex = CType(resources.GetObject("PanelToBeAdded.TabIndex"), Integer)
        Me.PanelToBeAdded.Text = resources.GetString("PanelToBeAdded.Text")
        Me.ToolTip1.SetToolTip(Me.PanelToBeAdded, resources.GetString("PanelToBeAdded.ToolTip"))
        Me.PanelToBeAdded.Visible = CType(resources.GetObject("PanelToBeAdded.Visible"), Boolean)
        '
        'ButtonStartFirstStage
        '
        Me.ButtonStartFirstStage.AccessibleDescription = resources.GetString("ButtonStartFirstStage.AccessibleDescription")
        Me.ButtonStartFirstStage.AccessibleName = resources.GetString("ButtonStartFirstStage.AccessibleName")
        Me.ButtonStartFirstStage.Anchor = CType(resources.GetObject("ButtonStartFirstStage.Anchor"), System.Windows.Forms.AnchorStyles)
        Me.ButtonStartFirstStage.BackgroundImage = CType(resources.GetObject("ButtonStartFirstStage.BackgroundImage"), System.Drawing.Image)
        Me.ButtonStartFirstStage.Dock = CType(resources.GetObject("ButtonStartFirstStage.Dock"), System.Windows.Forms.DockStyle)
        Me.ButtonStartFirstStage.Enabled = CType(resources.GetObject("ButtonStartFirstStage.Enabled"), Boolean)
        Me.ButtonStartFirstStage.FlatStyle = CType(resources.GetObject("ButtonStartFirstStage.FlatStyle"), System.Windows.Forms.FlatStyle)
        Me.ButtonStartFirstStage.Font = CType(resources.GetObject("ButtonStartFirstStage.Font"), System.Drawing.Font)
        Me.ButtonStartFirstStage.Image = CType(resources.GetObject("ButtonStartFirstStage.Image"), System.Drawing.Image)
        Me.ButtonStartFirstStage.ImageAlign = CType(resources.GetObject("ButtonStartFirstStage.ImageAlign"), System.Drawing.ContentAlignment)
        Me.ButtonStartFirstStage.ImageIndex = CType(resources.GetObject("ButtonStartFirstStage.ImageIndex"), Integer)
        Me.ButtonStartFirstStage.ImeMode = CType(resources.GetObject("ButtonStartFirstStage.ImeMode"), System.Windows.Forms.ImeMode)
        Me.ButtonStartFirstStage.Location = CType(resources.GetObject("ButtonStartFirstStage.Location"), System.Drawing.Point)
        Me.ButtonStartFirstStage.Name = "ButtonStartFirstStage"
        Me.ButtonStartFirstStage.RightToLeft = CType(resources.GetObject("ButtonStartFirstStage.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.ButtonStartFirstStage.Size = CType(resources.GetObject("ButtonStartFirstStage.Size"), System.Drawing.Size)
        Me.ButtonStartFirstStage.TabIndex = CType(resources.GetObject("ButtonStartFirstStage.TabIndex"), Integer)
        Me.ButtonStartFirstStage.Text = resources.GetString("ButtonStartFirstStage.Text")
        Me.ButtonStartFirstStage.TextAlign = CType(resources.GetObject("ButtonStartFirstStage.TextAlign"), System.Drawing.ContentAlignment)
        Me.ToolTip1.SetToolTip(Me.ButtonStartFirstStage, resources.GetString("ButtonStartFirstStage.ToolTip"))
        Me.ButtonStartFirstStage.Visible = CType(resources.GetObject("ButtonStartFirstStage.Visible"), Boolean)
        '
        'btRelease
        '
        Me.btRelease.AccessibleDescription = resources.GetString("btRelease.AccessibleDescription")
        Me.btRelease.AccessibleName = resources.GetString("btRelease.AccessibleName")
        Me.btRelease.Anchor = CType(resources.GetObject("btRelease.Anchor"), System.Windows.Forms.AnchorStyles)
        Me.btRelease.BackgroundImage = CType(resources.GetObject("btRelease.BackgroundImage"), System.Drawing.Image)
        Me.btRelease.Dock = CType(resources.GetObject("btRelease.Dock"), System.Windows.Forms.DockStyle)
        Me.btRelease.Enabled = CType(resources.GetObject("btRelease.Enabled"), Boolean)
        Me.btRelease.FlatStyle = CType(resources.GetObject("btRelease.FlatStyle"), System.Windows.Forms.FlatStyle)
        Me.btRelease.Font = CType(resources.GetObject("btRelease.Font"), System.Drawing.Font)
        Me.btRelease.Image = CType(resources.GetObject("btRelease.Image"), System.Drawing.Image)
        Me.btRelease.ImageAlign = CType(resources.GetObject("btRelease.ImageAlign"), System.Drawing.ContentAlignment)
        Me.btRelease.ImageIndex = CType(resources.GetObject("btRelease.ImageIndex"), Integer)
        Me.btRelease.ImeMode = CType(resources.GetObject("btRelease.ImeMode"), System.Windows.Forms.ImeMode)
        Me.btRelease.Location = CType(resources.GetObject("btRelease.Location"), System.Drawing.Point)
        Me.btRelease.Name = "btRelease"
        Me.btRelease.RightToLeft = CType(resources.GetObject("btRelease.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.btRelease.Size = CType(resources.GetObject("btRelease.Size"), System.Drawing.Size)
        Me.btRelease.TabIndex = CType(resources.GetObject("btRelease.TabIndex"), Integer)
        Me.btRelease.Text = resources.GetString("btRelease.Text")
        Me.btRelease.TextAlign = CType(resources.GetObject("btRelease.TextAlign"), System.Drawing.ContentAlignment)
        Me.ToolTip1.SetToolTip(Me.btRelease, resources.GetString("btRelease.ToolTip"))
        Me.btRelease.Visible = CType(resources.GetObject("btRelease.Visible"), Boolean)
        '
        'btRetrieve
        '
        Me.btRetrieve.AccessibleDescription = resources.GetString("btRetrieve.AccessibleDescription")
        Me.btRetrieve.AccessibleName = resources.GetString("btRetrieve.AccessibleName")
        Me.btRetrieve.Anchor = CType(resources.GetObject("btRetrieve.Anchor"), System.Windows.Forms.AnchorStyles)
        Me.btRetrieve.BackgroundImage = CType(resources.GetObject("btRetrieve.BackgroundImage"), System.Drawing.Image)
        Me.btRetrieve.Dock = CType(resources.GetObject("btRetrieve.Dock"), System.Windows.Forms.DockStyle)
        Me.btRetrieve.Enabled = CType(resources.GetObject("btRetrieve.Enabled"), Boolean)
        Me.btRetrieve.FlatStyle = CType(resources.GetObject("btRetrieve.FlatStyle"), System.Windows.Forms.FlatStyle)
        Me.btRetrieve.Font = CType(resources.GetObject("btRetrieve.Font"), System.Drawing.Font)
        Me.btRetrieve.Image = CType(resources.GetObject("btRetrieve.Image"), System.Drawing.Image)
        Me.btRetrieve.ImageAlign = CType(resources.GetObject("btRetrieve.ImageAlign"), System.Drawing.ContentAlignment)
        Me.btRetrieve.ImageIndex = CType(resources.GetObject("btRetrieve.ImageIndex"), Integer)
        Me.btRetrieve.ImeMode = CType(resources.GetObject("btRetrieve.ImeMode"), System.Windows.Forms.ImeMode)
        Me.btRetrieve.Location = CType(resources.GetObject("btRetrieve.Location"), System.Drawing.Point)
        Me.btRetrieve.Name = "btRetrieve"
        Me.btRetrieve.RightToLeft = CType(resources.GetObject("btRetrieve.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.btRetrieve.Size = CType(resources.GetObject("btRetrieve.Size"), System.Drawing.Size)
        Me.btRetrieve.TabIndex = CType(resources.GetObject("btRetrieve.TabIndex"), Integer)
        Me.btRetrieve.Text = resources.GetString("btRetrieve.Text")
        Me.btRetrieve.TextAlign = CType(resources.GetObject("btRetrieve.TextAlign"), System.Drawing.ContentAlignment)
        Me.ToolTip1.SetToolTip(Me.btRetrieve, resources.GetString("btRetrieve.ToolTip"))
        Me.btRetrieve.Visible = CType(resources.GetObject("btRetrieve.Visible"), Boolean)
        '
        'GroupBox1
        '
        Me.GroupBox1.AccessibleDescription = resources.GetString("GroupBox1.AccessibleDescription")
        Me.GroupBox1.AccessibleName = resources.GetString("GroupBox1.AccessibleName")
        Me.GroupBox1.Anchor = CType(resources.GetObject("GroupBox1.Anchor"), System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.BackgroundImage = CType(resources.GetObject("GroupBox1.BackgroundImage"), System.Drawing.Image)
        Me.GroupBox1.Controls.Add(Me.btStop)
        Me.GroupBox1.Controls.Add(Me.btPause)
        Me.GroupBox1.Controls.Add(Me.btPlay)
        Me.GroupBox1.Controls.Add(Me.DispensingMode)
        Me.GroupBox1.Controls.Add(Me.cbContinueTest)
        Me.GroupBox1.Dock = CType(resources.GetObject("GroupBox1.Dock"), System.Windows.Forms.DockStyle)
        Me.GroupBox1.Enabled = CType(resources.GetObject("GroupBox1.Enabled"), Boolean)
        Me.GroupBox1.Font = CType(resources.GetObject("GroupBox1.Font"), System.Drawing.Font)
        Me.GroupBox1.ImeMode = CType(resources.GetObject("GroupBox1.ImeMode"), System.Windows.Forms.ImeMode)
        Me.GroupBox1.Location = CType(resources.GetObject("GroupBox1.Location"), System.Drawing.Point)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.RightToLeft = CType(resources.GetObject("GroupBox1.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.GroupBox1.Size = CType(resources.GetObject("GroupBox1.Size"), System.Drawing.Size)
        Me.GroupBox1.TabIndex = CType(resources.GetObject("GroupBox1.TabIndex"), Integer)
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = resources.GetString("GroupBox1.Text")
        Me.ToolTip1.SetToolTip(Me.GroupBox1, resources.GetString("GroupBox1.ToolTip"))
        Me.GroupBox1.Visible = CType(resources.GetObject("GroupBox1.Visible"), Boolean)
        '
        'cbContinueTest
        '
        Me.cbContinueTest.AccessibleDescription = resources.GetString("cbContinueTest.AccessibleDescription")
        Me.cbContinueTest.AccessibleName = resources.GetString("cbContinueTest.AccessibleName")
        Me.cbContinueTest.Anchor = CType(resources.GetObject("cbContinueTest.Anchor"), System.Windows.Forms.AnchorStyles)
        Me.cbContinueTest.Appearance = CType(resources.GetObject("cbContinueTest.Appearance"), System.Windows.Forms.Appearance)
        Me.cbContinueTest.BackgroundImage = CType(resources.GetObject("cbContinueTest.BackgroundImage"), System.Drawing.Image)
        Me.cbContinueTest.CheckAlign = CType(resources.GetObject("cbContinueTest.CheckAlign"), System.Drawing.ContentAlignment)
        Me.cbContinueTest.Dock = CType(resources.GetObject("cbContinueTest.Dock"), System.Windows.Forms.DockStyle)
        Me.cbContinueTest.Enabled = CType(resources.GetObject("cbContinueTest.Enabled"), Boolean)
        Me.cbContinueTest.FlatStyle = CType(resources.GetObject("cbContinueTest.FlatStyle"), System.Windows.Forms.FlatStyle)
        Me.cbContinueTest.Font = CType(resources.GetObject("cbContinueTest.Font"), System.Drawing.Font)
        Me.cbContinueTest.Image = CType(resources.GetObject("cbContinueTest.Image"), System.Drawing.Image)
        Me.cbContinueTest.ImageAlign = CType(resources.GetObject("cbContinueTest.ImageAlign"), System.Drawing.ContentAlignment)
        Me.cbContinueTest.ImageIndex = CType(resources.GetObject("cbContinueTest.ImageIndex"), Integer)
        Me.cbContinueTest.ImeMode = CType(resources.GetObject("cbContinueTest.ImeMode"), System.Windows.Forms.ImeMode)
        Me.cbContinueTest.Location = CType(resources.GetObject("cbContinueTest.Location"), System.Drawing.Point)
        Me.cbContinueTest.Name = "cbContinueTest"
        Me.cbContinueTest.RightToLeft = CType(resources.GetObject("cbContinueTest.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.cbContinueTest.Size = CType(resources.GetObject("cbContinueTest.Size"), System.Drawing.Size)
        Me.cbContinueTest.TabIndex = CType(resources.GetObject("cbContinueTest.TabIndex"), Integer)
        Me.cbContinueTest.Text = resources.GetString("cbContinueTest.Text")
        Me.cbContinueTest.TextAlign = CType(resources.GetObject("cbContinueTest.TextAlign"), System.Drawing.ContentAlignment)
        Me.ToolTip1.SetToolTip(Me.cbContinueTest, resources.GetString("cbContinueTest.ToolTip"))
        Me.cbContinueTest.Visible = CType(resources.GetObject("cbContinueTest.Visible"), Boolean)
        '
        'gbConveyor
        '
        Me.gbConveyor.AccessibleDescription = resources.GetString("gbConveyor.AccessibleDescription")
        Me.gbConveyor.AccessibleName = resources.GetString("gbConveyor.AccessibleName")
        Me.gbConveyor.Anchor = CType(resources.GetObject("gbConveyor.Anchor"), System.Windows.Forms.AnchorStyles)
        Me.gbConveyor.BackgroundImage = CType(resources.GetObject("gbConveyor.BackgroundImage"), System.Drawing.Image)
        Me.gbConveyor.Controls.Add(Me.btReset)
        Me.gbConveyor.Controls.Add(Me.btEject)
        Me.gbConveyor.Controls.Add(Me.ButtonStartFirstStage)
        Me.gbConveyor.Controls.Add(Me.btRelease)
        Me.gbConveyor.Controls.Add(Me.btRetrieve)
        Me.gbConveyor.Dock = CType(resources.GetObject("gbConveyor.Dock"), System.Windows.Forms.DockStyle)
        Me.gbConveyor.Enabled = CType(resources.GetObject("gbConveyor.Enabled"), Boolean)
        Me.gbConveyor.Font = CType(resources.GetObject("gbConveyor.Font"), System.Drawing.Font)
        Me.gbConveyor.ImeMode = CType(resources.GetObject("gbConveyor.ImeMode"), System.Windows.Forms.ImeMode)
        Me.gbConveyor.Location = CType(resources.GetObject("gbConveyor.Location"), System.Drawing.Point)
        Me.gbConveyor.Name = "gbConveyor"
        Me.gbConveyor.RightToLeft = CType(resources.GetObject("gbConveyor.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.gbConveyor.Size = CType(resources.GetObject("gbConveyor.Size"), System.Drawing.Size)
        Me.gbConveyor.TabIndex = CType(resources.GetObject("gbConveyor.TabIndex"), Integer)
        Me.gbConveyor.TabStop = False
        Me.gbConveyor.Text = resources.GetString("gbConveyor.Text")
        Me.ToolTip1.SetToolTip(Me.gbConveyor, resources.GetString("gbConveyor.ToolTip"))
        Me.gbConveyor.Visible = CType(resources.GetObject("gbConveyor.Visible"), Boolean)
        '
        'btReset
        '
        Me.btReset.AccessibleDescription = resources.GetString("btReset.AccessibleDescription")
        Me.btReset.AccessibleName = resources.GetString("btReset.AccessibleName")
        Me.btReset.Anchor = CType(resources.GetObject("btReset.Anchor"), System.Windows.Forms.AnchorStyles)
        Me.btReset.BackgroundImage = CType(resources.GetObject("btReset.BackgroundImage"), System.Drawing.Image)
        Me.btReset.Dock = CType(resources.GetObject("btReset.Dock"), System.Windows.Forms.DockStyle)
        Me.btReset.Enabled = CType(resources.GetObject("btReset.Enabled"), Boolean)
        Me.btReset.FlatStyle = CType(resources.GetObject("btReset.FlatStyle"), System.Windows.Forms.FlatStyle)
        Me.btReset.Font = CType(resources.GetObject("btReset.Font"), System.Drawing.Font)
        Me.btReset.Image = CType(resources.GetObject("btReset.Image"), System.Drawing.Image)
        Me.btReset.ImageAlign = CType(resources.GetObject("btReset.ImageAlign"), System.Drawing.ContentAlignment)
        Me.btReset.ImageIndex = CType(resources.GetObject("btReset.ImageIndex"), Integer)
        Me.btReset.ImeMode = CType(resources.GetObject("btReset.ImeMode"), System.Windows.Forms.ImeMode)
        Me.btReset.Location = CType(resources.GetObject("btReset.Location"), System.Drawing.Point)
        Me.btReset.Name = "btReset"
        Me.btReset.RightToLeft = CType(resources.GetObject("btReset.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.btReset.Size = CType(resources.GetObject("btReset.Size"), System.Drawing.Size)
        Me.btReset.TabIndex = CType(resources.GetObject("btReset.TabIndex"), Integer)
        Me.btReset.Text = resources.GetString("btReset.Text")
        Me.btReset.TextAlign = CType(resources.GetObject("btReset.TextAlign"), System.Drawing.ContentAlignment)
        Me.ToolTip1.SetToolTip(Me.btReset, resources.GetString("btReset.ToolTip"))
        Me.btReset.Visible = CType(resources.GetObject("btReset.Visible"), Boolean)
        '
        'btEject
        '
        Me.btEject.AccessibleDescription = resources.GetString("btEject.AccessibleDescription")
        Me.btEject.AccessibleName = resources.GetString("btEject.AccessibleName")
        Me.btEject.Anchor = CType(resources.GetObject("btEject.Anchor"), System.Windows.Forms.AnchorStyles)
        Me.btEject.BackgroundImage = CType(resources.GetObject("btEject.BackgroundImage"), System.Drawing.Image)
        Me.btEject.Dock = CType(resources.GetObject("btEject.Dock"), System.Windows.Forms.DockStyle)
        Me.btEject.Enabled = CType(resources.GetObject("btEject.Enabled"), Boolean)
        Me.btEject.FlatStyle = CType(resources.GetObject("btEject.FlatStyle"), System.Windows.Forms.FlatStyle)
        Me.btEject.Font = CType(resources.GetObject("btEject.Font"), System.Drawing.Font)
        Me.btEject.Image = CType(resources.GetObject("btEject.Image"), System.Drawing.Image)
        Me.btEject.ImageAlign = CType(resources.GetObject("btEject.ImageAlign"), System.Drawing.ContentAlignment)
        Me.btEject.ImageIndex = CType(resources.GetObject("btEject.ImageIndex"), Integer)
        Me.btEject.ImeMode = CType(resources.GetObject("btEject.ImeMode"), System.Windows.Forms.ImeMode)
        Me.btEject.Location = CType(resources.GetObject("btEject.Location"), System.Drawing.Point)
        Me.btEject.Name = "btEject"
        Me.btEject.RightToLeft = CType(resources.GetObject("btEject.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.btEject.Size = CType(resources.GetObject("btEject.Size"), System.Drawing.Size)
        Me.btEject.TabIndex = CType(resources.GetObject("btEject.TabIndex"), Integer)
        Me.btEject.Text = resources.GetString("btEject.Text")
        Me.btEject.TextAlign = CType(resources.GetObject("btEject.TextAlign"), System.Drawing.ContentAlignment)
        Me.ToolTip1.SetToolTip(Me.btEject, resources.GetString("btEject.ToolTip"))
        Me.btEject.Visible = CType(resources.GetObject("btEject.Visible"), Boolean)
        '
        'tbOpenedFile
        '
        Me.tbOpenedFile.AccessibleDescription = resources.GetString("tbOpenedFile.AccessibleDescription")
        Me.tbOpenedFile.AccessibleName = resources.GetString("tbOpenedFile.AccessibleName")
        Me.tbOpenedFile.Anchor = CType(resources.GetObject("tbOpenedFile.Anchor"), System.Windows.Forms.AnchorStyles)
        Me.tbOpenedFile.AutoSize = CType(resources.GetObject("tbOpenedFile.AutoSize"), Boolean)
        Me.tbOpenedFile.BackgroundImage = CType(resources.GetObject("tbOpenedFile.BackgroundImage"), System.Drawing.Image)
        Me.tbOpenedFile.Dock = CType(resources.GetObject("tbOpenedFile.Dock"), System.Windows.Forms.DockStyle)
        Me.tbOpenedFile.Enabled = CType(resources.GetObject("tbOpenedFile.Enabled"), Boolean)
        Me.tbOpenedFile.Font = CType(resources.GetObject("tbOpenedFile.Font"), System.Drawing.Font)
        Me.tbOpenedFile.ImeMode = CType(resources.GetObject("tbOpenedFile.ImeMode"), System.Windows.Forms.ImeMode)
        Me.tbOpenedFile.Location = CType(resources.GetObject("tbOpenedFile.Location"), System.Drawing.Point)
        Me.tbOpenedFile.MaxLength = CType(resources.GetObject("tbOpenedFile.MaxLength"), Integer)
        Me.tbOpenedFile.Multiline = CType(resources.GetObject("tbOpenedFile.Multiline"), Boolean)
        Me.tbOpenedFile.Name = "tbOpenedFile"
        Me.tbOpenedFile.PasswordChar = CType(resources.GetObject("tbOpenedFile.PasswordChar"), Char)
        Me.tbOpenedFile.ReadOnly = True
        Me.tbOpenedFile.RightToLeft = CType(resources.GetObject("tbOpenedFile.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.tbOpenedFile.ScrollBars = CType(resources.GetObject("tbOpenedFile.ScrollBars"), System.Windows.Forms.ScrollBars)
        Me.tbOpenedFile.Size = CType(resources.GetObject("tbOpenedFile.Size"), System.Drawing.Size)
        Me.tbOpenedFile.TabIndex = CType(resources.GetObject("tbOpenedFile.TabIndex"), Integer)
        Me.tbOpenedFile.Text = resources.GetString("tbOpenedFile.Text")
        Me.tbOpenedFile.TextAlign = CType(resources.GetObject("tbOpenedFile.TextAlign"), System.Windows.Forms.HorizontalAlignment)
        Me.ToolTip1.SetToolTip(Me.tbOpenedFile, resources.GetString("tbOpenedFile.ToolTip"))
        Me.tbOpenedFile.Visible = CType(resources.GetObject("tbOpenedFile.Visible"), Boolean)
        Me.tbOpenedFile.WordWrap = CType(resources.GetObject("tbOpenedFile.WordWrap"), Boolean)
        '
        'GroupBox3
        '
        Me.GroupBox3.AccessibleDescription = resources.GetString("GroupBox3.AccessibleDescription")
        Me.GroupBox3.AccessibleName = resources.GetString("GroupBox3.AccessibleName")
        Me.GroupBox3.Anchor = CType(resources.GetObject("GroupBox3.Anchor"), System.Windows.Forms.AnchorStyles)
        Me.GroupBox3.BackgroundImage = CType(resources.GetObject("GroupBox3.BackgroundImage"), System.Drawing.Image)
        Me.GroupBox3.Controls.Add(Me.VisionMode)
        Me.GroupBox3.Controls.Add(Me.NeedleMode)
        Me.GroupBox3.Controls.Add(Me.Label5)
        Me.GroupBox3.Dock = CType(resources.GetObject("GroupBox3.Dock"), System.Windows.Forms.DockStyle)
        Me.GroupBox3.Enabled = CType(resources.GetObject("GroupBox3.Enabled"), Boolean)
        Me.GroupBox3.Font = CType(resources.GetObject("GroupBox3.Font"), System.Drawing.Font)
        Me.GroupBox3.ImeMode = CType(resources.GetObject("GroupBox3.ImeMode"), System.Windows.Forms.ImeMode)
        Me.GroupBox3.Location = CType(resources.GetObject("GroupBox3.Location"), System.Drawing.Point)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.RightToLeft = CType(resources.GetObject("GroupBox3.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.GroupBox3.Size = CType(resources.GetObject("GroupBox3.Size"), System.Drawing.Size)
        Me.GroupBox3.TabIndex = CType(resources.GetObject("GroupBox3.TabIndex"), Integer)
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = resources.GetString("GroupBox3.Text")
        Me.ToolTip1.SetToolTip(Me.GroupBox3, resources.GetString("GroupBox3.ToolTip"))
        Me.GroupBox3.Visible = CType(resources.GetObject("GroupBox3.Visible"), Boolean)
        '
        'btChangeSyringe
        '
        Me.btChangeSyringe.AccessibleDescription = resources.GetString("btChangeSyringe.AccessibleDescription")
        Me.btChangeSyringe.AccessibleName = resources.GetString("btChangeSyringe.AccessibleName")
        Me.btChangeSyringe.Anchor = CType(resources.GetObject("btChangeSyringe.Anchor"), System.Windows.Forms.AnchorStyles)
        Me.btChangeSyringe.BackColor = System.Drawing.SystemColors.Control
        Me.btChangeSyringe.BackgroundImage = CType(resources.GetObject("btChangeSyringe.BackgroundImage"), System.Drawing.Image)
        Me.btChangeSyringe.Dock = CType(resources.GetObject("btChangeSyringe.Dock"), System.Windows.Forms.DockStyle)
        Me.btChangeSyringe.Enabled = CType(resources.GetObject("btChangeSyringe.Enabled"), Boolean)
        Me.btChangeSyringe.FlatStyle = CType(resources.GetObject("btChangeSyringe.FlatStyle"), System.Windows.Forms.FlatStyle)
        Me.btChangeSyringe.Font = CType(resources.GetObject("btChangeSyringe.Font"), System.Drawing.Font)
        Me.btChangeSyringe.Image = CType(resources.GetObject("btChangeSyringe.Image"), System.Drawing.Image)
        Me.btChangeSyringe.ImageAlign = CType(resources.GetObject("btChangeSyringe.ImageAlign"), System.Drawing.ContentAlignment)
        Me.btChangeSyringe.ImageIndex = CType(resources.GetObject("btChangeSyringe.ImageIndex"), Integer)
        Me.btChangeSyringe.ImageList = Me.ImageListGeneralTools
        Me.btChangeSyringe.ImeMode = CType(resources.GetObject("btChangeSyringe.ImeMode"), System.Windows.Forms.ImeMode)
        Me.btChangeSyringe.Location = CType(resources.GetObject("btChangeSyringe.Location"), System.Drawing.Point)
        Me.btChangeSyringe.Name = "btChangeSyringe"
        Me.btChangeSyringe.RightToLeft = CType(resources.GetObject("btChangeSyringe.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.btChangeSyringe.Size = CType(resources.GetObject("btChangeSyringe.Size"), System.Drawing.Size)
        Me.btChangeSyringe.TabIndex = CType(resources.GetObject("btChangeSyringe.TabIndex"), Integer)
        Me.btChangeSyringe.Text = resources.GetString("btChangeSyringe.Text")
        Me.btChangeSyringe.TextAlign = CType(resources.GetObject("btChangeSyringe.TextAlign"), System.Drawing.ContentAlignment)
        Me.ToolTip1.SetToolTip(Me.btChangeSyringe, resources.GetString("btChangeSyringe.ToolTip"))
        Me.btChangeSyringe.Visible = CType(resources.GetObject("btChangeSyringe.Visible"), Boolean)
        '
        'btExit
        '
        Me.btExit.AccessibleDescription = resources.GetString("btExit.AccessibleDescription")
        Me.btExit.AccessibleName = resources.GetString("btExit.AccessibleName")
        Me.btExit.Anchor = CType(resources.GetObject("btExit.Anchor"), System.Windows.Forms.AnchorStyles)
        Me.btExit.BackColor = System.Drawing.SystemColors.Control
        Me.btExit.BackgroundImage = CType(resources.GetObject("btExit.BackgroundImage"), System.Drawing.Image)
        Me.btExit.Dock = CType(resources.GetObject("btExit.Dock"), System.Windows.Forms.DockStyle)
        Me.btExit.Enabled = CType(resources.GetObject("btExit.Enabled"), Boolean)
        Me.btExit.FlatStyle = CType(resources.GetObject("btExit.FlatStyle"), System.Windows.Forms.FlatStyle)
        Me.btExit.Font = CType(resources.GetObject("btExit.Font"), System.Drawing.Font)
        Me.btExit.Image = CType(resources.GetObject("btExit.Image"), System.Drawing.Image)
        Me.btExit.ImageAlign = CType(resources.GetObject("btExit.ImageAlign"), System.Drawing.ContentAlignment)
        Me.btExit.ImageIndex = CType(resources.GetObject("btExit.ImageIndex"), Integer)
        Me.btExit.ImeMode = CType(resources.GetObject("btExit.ImeMode"), System.Windows.Forms.ImeMode)
        Me.btExit.Location = CType(resources.GetObject("btExit.Location"), System.Drawing.Point)
        Me.btExit.Name = "btExit"
        Me.btExit.RightToLeft = CType(resources.GetObject("btExit.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.btExit.Size = CType(resources.GetObject("btExit.Size"), System.Drawing.Size)
        Me.btExit.TabIndex = CType(resources.GetObject("btExit.TabIndex"), Integer)
        Me.btExit.Text = resources.GetString("btExit.Text")
        Me.btExit.TextAlign = CType(resources.GetObject("btExit.TextAlign"), System.Drawing.ContentAlignment)
        Me.ToolTip1.SetToolTip(Me.btExit, resources.GetString("btExit.ToolTip"))
        Me.btExit.Visible = CType(resources.GetObject("btExit.Visible"), Boolean)
        '
        'btResetVolCalSettings
        '
        Me.btResetVolCalSettings.AccessibleDescription = resources.GetString("btResetVolCalSettings.AccessibleDescription")
        Me.btResetVolCalSettings.AccessibleName = resources.GetString("btResetVolCalSettings.AccessibleName")
        Me.btResetVolCalSettings.Anchor = CType(resources.GetObject("btResetVolCalSettings.Anchor"), System.Windows.Forms.AnchorStyles)
        Me.btResetVolCalSettings.BackColor = System.Drawing.SystemColors.Control
        Me.btResetVolCalSettings.BackgroundImage = CType(resources.GetObject("btResetVolCalSettings.BackgroundImage"), System.Drawing.Image)
        Me.btResetVolCalSettings.Dock = CType(resources.GetObject("btResetVolCalSettings.Dock"), System.Windows.Forms.DockStyle)
        Me.btResetVolCalSettings.Enabled = CType(resources.GetObject("btResetVolCalSettings.Enabled"), Boolean)
        Me.btResetVolCalSettings.FlatStyle = CType(resources.GetObject("btResetVolCalSettings.FlatStyle"), System.Windows.Forms.FlatStyle)
        Me.btResetVolCalSettings.Font = CType(resources.GetObject("btResetVolCalSettings.Font"), System.Drawing.Font)
        Me.btResetVolCalSettings.Image = CType(resources.GetObject("btResetVolCalSettings.Image"), System.Drawing.Image)
        Me.btResetVolCalSettings.ImageAlign = CType(resources.GetObject("btResetVolCalSettings.ImageAlign"), System.Drawing.ContentAlignment)
        Me.btResetVolCalSettings.ImageIndex = CType(resources.GetObject("btResetVolCalSettings.ImageIndex"), Integer)
        Me.btResetVolCalSettings.ImageList = Me.ImageListGeneralTools
        Me.btResetVolCalSettings.ImeMode = CType(resources.GetObject("btResetVolCalSettings.ImeMode"), System.Windows.Forms.ImeMode)
        Me.btResetVolCalSettings.Location = CType(resources.GetObject("btResetVolCalSettings.Location"), System.Drawing.Point)
        Me.btResetVolCalSettings.Name = "btResetVolCalSettings"
        Me.btResetVolCalSettings.RightToLeft = CType(resources.GetObject("btResetVolCalSettings.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.btResetVolCalSettings.Size = CType(resources.GetObject("btResetVolCalSettings.Size"), System.Drawing.Size)
        Me.btResetVolCalSettings.TabIndex = CType(resources.GetObject("btResetVolCalSettings.TabIndex"), Integer)
        Me.btResetVolCalSettings.Text = resources.GetString("btResetVolCalSettings.Text")
        Me.btResetVolCalSettings.TextAlign = CType(resources.GetObject("btResetVolCalSettings.TextAlign"), System.Drawing.ContentAlignment)
        Me.ToolTip1.SetToolTip(Me.btResetVolCalSettings, resources.GetString("btResetVolCalSettings.ToolTip"))
        Me.btResetVolCalSettings.Visible = CType(resources.GetObject("btResetVolCalSettings.Visible"), Boolean)
        '
        'tbLsHeight
        '
        Me.tbLsHeight.AccessibleDescription = resources.GetString("tbLsHeight.AccessibleDescription")
        Me.tbLsHeight.AccessibleName = resources.GetString("tbLsHeight.AccessibleName")
        Me.tbLsHeight.Anchor = CType(resources.GetObject("tbLsHeight.Anchor"), System.Windows.Forms.AnchorStyles)
        Me.tbLsHeight.AutoSize = CType(resources.GetObject("tbLsHeight.AutoSize"), Boolean)
        Me.tbLsHeight.BackgroundImage = CType(resources.GetObject("tbLsHeight.BackgroundImage"), System.Drawing.Image)
        Me.tbLsHeight.Dock = CType(resources.GetObject("tbLsHeight.Dock"), System.Windows.Forms.DockStyle)
        Me.tbLsHeight.Enabled = CType(resources.GetObject("tbLsHeight.Enabled"), Boolean)
        Me.tbLsHeight.Font = CType(resources.GetObject("tbLsHeight.Font"), System.Drawing.Font)
        Me.tbLsHeight.ImeMode = CType(resources.GetObject("tbLsHeight.ImeMode"), System.Windows.Forms.ImeMode)
        Me.tbLsHeight.Location = CType(resources.GetObject("tbLsHeight.Location"), System.Drawing.Point)
        Me.tbLsHeight.MaxLength = CType(resources.GetObject("tbLsHeight.MaxLength"), Integer)
        Me.tbLsHeight.Multiline = CType(resources.GetObject("tbLsHeight.Multiline"), Boolean)
        Me.tbLsHeight.Name = "tbLsHeight"
        Me.tbLsHeight.PasswordChar = CType(resources.GetObject("tbLsHeight.PasswordChar"), Char)
        Me.tbLsHeight.ReadOnly = True
        Me.tbLsHeight.RightToLeft = CType(resources.GetObject("tbLsHeight.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.tbLsHeight.ScrollBars = CType(resources.GetObject("tbLsHeight.ScrollBars"), System.Windows.Forms.ScrollBars)
        Me.tbLsHeight.Size = CType(resources.GetObject("tbLsHeight.Size"), System.Drawing.Size)
        Me.tbLsHeight.TabIndex = CType(resources.GetObject("tbLsHeight.TabIndex"), Integer)
        Me.tbLsHeight.Text = resources.GetString("tbLsHeight.Text")
        Me.tbLsHeight.TextAlign = CType(resources.GetObject("tbLsHeight.TextAlign"), System.Windows.Forms.HorizontalAlignment)
        Me.ToolTip1.SetToolTip(Me.tbLsHeight, resources.GetString("tbLsHeight.ToolTip"))
        Me.tbLsHeight.Visible = CType(resources.GetObject("tbLsHeight.Visible"), Boolean)
        Me.tbLsHeight.WordWrap = CType(resources.GetObject("tbLsHeight.WordWrap"), Boolean)
        '
        'btConfirm
        '
        Me.btConfirm.AccessibleDescription = resources.GetString("btConfirm.AccessibleDescription")
        Me.btConfirm.AccessibleName = resources.GetString("btConfirm.AccessibleName")
        Me.btConfirm.Anchor = CType(resources.GetObject("btConfirm.Anchor"), System.Windows.Forms.AnchorStyles)
        Me.btConfirm.BackgroundImage = CType(resources.GetObject("btConfirm.BackgroundImage"), System.Drawing.Image)
        Me.btConfirm.Dock = CType(resources.GetObject("btConfirm.Dock"), System.Windows.Forms.DockStyle)
        Me.btConfirm.Enabled = CType(resources.GetObject("btConfirm.Enabled"), Boolean)
        Me.btConfirm.FlatStyle = CType(resources.GetObject("btConfirm.FlatStyle"), System.Windows.Forms.FlatStyle)
        Me.btConfirm.Font = CType(resources.GetObject("btConfirm.Font"), System.Drawing.Font)
        Me.btConfirm.ForeColor = System.Drawing.Color.Green
        Me.btConfirm.Image = CType(resources.GetObject("btConfirm.Image"), System.Drawing.Image)
        Me.btConfirm.ImageAlign = CType(resources.GetObject("btConfirm.ImageAlign"), System.Drawing.ContentAlignment)
        Me.btConfirm.ImageIndex = CType(resources.GetObject("btConfirm.ImageIndex"), Integer)
        Me.btConfirm.ImeMode = CType(resources.GetObject("btConfirm.ImeMode"), System.Windows.Forms.ImeMode)
        Me.btConfirm.Location = CType(resources.GetObject("btConfirm.Location"), System.Drawing.Point)
        Me.btConfirm.Name = "btConfirm"
        Me.btConfirm.RightToLeft = CType(resources.GetObject("btConfirm.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.btConfirm.Size = CType(resources.GetObject("btConfirm.Size"), System.Drawing.Size)
        Me.btConfirm.TabIndex = CType(resources.GetObject("btConfirm.TabIndex"), Integer)
        Me.btConfirm.Text = resources.GetString("btConfirm.Text")
        Me.btConfirm.TextAlign = CType(resources.GetObject("btConfirm.TextAlign"), System.Drawing.ContentAlignment)
        Me.ToolTip1.SetToolTip(Me.btConfirm, resources.GetString("btConfirm.ToolTip"))
        Me.btConfirm.Visible = CType(resources.GetObject("btConfirm.Visible"), Boolean)
        '
        'btCancel
        '
        Me.btCancel.AccessibleDescription = resources.GetString("btCancel.AccessibleDescription")
        Me.btCancel.AccessibleName = resources.GetString("btCancel.AccessibleName")
        Me.btCancel.Anchor = CType(resources.GetObject("btCancel.Anchor"), System.Windows.Forms.AnchorStyles)
        Me.btCancel.BackgroundImage = CType(resources.GetObject("btCancel.BackgroundImage"), System.Drawing.Image)
        Me.btCancel.Dock = CType(resources.GetObject("btCancel.Dock"), System.Windows.Forms.DockStyle)
        Me.btCancel.Enabled = CType(resources.GetObject("btCancel.Enabled"), Boolean)
        Me.btCancel.FlatStyle = CType(resources.GetObject("btCancel.FlatStyle"), System.Windows.Forms.FlatStyle)
        Me.btCancel.Font = CType(resources.GetObject("btCancel.Font"), System.Drawing.Font)
        Me.btCancel.ForeColor = System.Drawing.Color.Red
        Me.btCancel.Image = CType(resources.GetObject("btCancel.Image"), System.Drawing.Image)
        Me.btCancel.ImageAlign = CType(resources.GetObject("btCancel.ImageAlign"), System.Drawing.ContentAlignment)
        Me.btCancel.ImageIndex = CType(resources.GetObject("btCancel.ImageIndex"), Integer)
        Me.btCancel.ImeMode = CType(resources.GetObject("btCancel.ImeMode"), System.Windows.Forms.ImeMode)
        Me.btCancel.Location = CType(resources.GetObject("btCancel.Location"), System.Drawing.Point)
        Me.btCancel.Name = "btCancel"
        Me.btCancel.RightToLeft = CType(resources.GetObject("btCancel.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.btCancel.Size = CType(resources.GetObject("btCancel.Size"), System.Drawing.Size)
        Me.btCancel.TabIndex = CType(resources.GetObject("btCancel.TabIndex"), Integer)
        Me.btCancel.Text = resources.GetString("btCancel.Text")
        Me.btCancel.TextAlign = CType(resources.GetObject("btCancel.TextAlign"), System.Drawing.ContentAlignment)
        Me.ToolTip1.SetToolTip(Me.btCancel, resources.GetString("btCancel.ToolTip"))
        Me.btCancel.Visible = CType(resources.GetObject("btCancel.Visible"), Boolean)
        '
        'Panel2
        '
        Me.Panel2.AccessibleDescription = resources.GetString("Panel2.AccessibleDescription")
        Me.Panel2.AccessibleName = resources.GetString("Panel2.AccessibleName")
        Me.Panel2.Anchor = CType(resources.GetObject("Panel2.Anchor"), System.Windows.Forms.AnchorStyles)
        Me.Panel2.AutoScroll = CType(resources.GetObject("Panel2.AutoScroll"), Boolean)
        Me.Panel2.AutoScrollMargin = CType(resources.GetObject("Panel2.AutoScrollMargin"), System.Drawing.Size)
        Me.Panel2.AutoScrollMinSize = CType(resources.GetObject("Panel2.AutoScrollMinSize"), System.Drawing.Size)
        Me.Panel2.BackgroundImage = CType(resources.GetObject("Panel2.BackgroundImage"), System.Drawing.Image)
        Me.Panel2.Controls.Add(Me.btConfirm)
        Me.Panel2.Controls.Add(Me.btCancel)
        Me.Panel2.Dock = CType(resources.GetObject("Panel2.Dock"), System.Windows.Forms.DockStyle)
        Me.Panel2.Enabled = CType(resources.GetObject("Panel2.Enabled"), Boolean)
        Me.Panel2.Font = CType(resources.GetObject("Panel2.Font"), System.Drawing.Font)
        Me.Panel2.ImeMode = CType(resources.GetObject("Panel2.ImeMode"), System.Windows.Forms.ImeMode)
        Me.Panel2.Location = CType(resources.GetObject("Panel2.Location"), System.Drawing.Point)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.RightToLeft = CType(resources.GetObject("Panel2.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.Panel2.Size = CType(resources.GetObject("Panel2.Size"), System.Drawing.Size)
        Me.Panel2.TabIndex = CType(resources.GetObject("Panel2.TabIndex"), Integer)
        Me.Panel2.Text = resources.GetString("Panel2.Text")
        Me.ToolTip1.SetToolTip(Me.Panel2, resources.GetString("Panel2.ToolTip"))
        Me.Panel2.Visible = CType(resources.GetObject("Panel2.Visible"), Boolean)
        '
        'Panel3
        '
        Me.Panel3.AccessibleDescription = resources.GetString("Panel3.AccessibleDescription")
        Me.Panel3.AccessibleName = resources.GetString("Panel3.AccessibleName")
        Me.Panel3.Anchor = CType(resources.GetObject("Panel3.Anchor"), System.Windows.Forms.AnchorStyles)
        Me.Panel3.AutoScroll = CType(resources.GetObject("Panel3.AutoScroll"), Boolean)
        Me.Panel3.AutoScrollMargin = CType(resources.GetObject("Panel3.AutoScrollMargin"), System.Drawing.Size)
        Me.Panel3.AutoScrollMinSize = CType(resources.GetObject("Panel3.AutoScrollMinSize"), System.Drawing.Size)
        Me.Panel3.BackgroundImage = CType(resources.GetObject("Panel3.BackgroundImage"), System.Drawing.Image)
        Me.Panel3.Controls.Add(Me.btStep)
        Me.Panel3.Controls.Add(Me.btEdit)
        Me.Panel3.Dock = CType(resources.GetObject("Panel3.Dock"), System.Windows.Forms.DockStyle)
        Me.Panel3.Enabled = CType(resources.GetObject("Panel3.Enabled"), Boolean)
        Me.Panel3.Font = CType(resources.GetObject("Panel3.Font"), System.Drawing.Font)
        Me.Panel3.ImeMode = CType(resources.GetObject("Panel3.ImeMode"), System.Windows.Forms.ImeMode)
        Me.Panel3.Location = CType(resources.GetObject("Panel3.Location"), System.Drawing.Point)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.RightToLeft = CType(resources.GetObject("Panel3.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.Panel3.Size = CType(resources.GetObject("Panel3.Size"), System.Drawing.Size)
        Me.Panel3.TabIndex = CType(resources.GetObject("Panel3.TabIndex"), Integer)
        Me.Panel3.Text = resources.GetString("Panel3.Text")
        Me.ToolTip1.SetToolTip(Me.Panel3, resources.GetString("Panel3.ToolTip"))
        Me.Panel3.Visible = CType(resources.GetObject("Panel3.Visible"), Boolean)
        '
        'btStep
        '
        Me.btStep.AccessibleDescription = resources.GetString("btStep.AccessibleDescription")
        Me.btStep.AccessibleName = resources.GetString("btStep.AccessibleName")
        Me.btStep.Anchor = CType(resources.GetObject("btStep.Anchor"), System.Windows.Forms.AnchorStyles)
        Me.btStep.BackgroundImage = CType(resources.GetObject("btStep.BackgroundImage"), System.Drawing.Image)
        Me.btStep.Dock = CType(resources.GetObject("btStep.Dock"), System.Windows.Forms.DockStyle)
        Me.btStep.Enabled = CType(resources.GetObject("btStep.Enabled"), Boolean)
        Me.btStep.FlatStyle = CType(resources.GetObject("btStep.FlatStyle"), System.Windows.Forms.FlatStyle)
        Me.btStep.Font = CType(resources.GetObject("btStep.Font"), System.Drawing.Font)
        Me.btStep.ForeColor = System.Drawing.Color.Green
        Me.btStep.Image = CType(resources.GetObject("btStep.Image"), System.Drawing.Image)
        Me.btStep.ImageAlign = CType(resources.GetObject("btStep.ImageAlign"), System.Drawing.ContentAlignment)
        Me.btStep.ImageIndex = CType(resources.GetObject("btStep.ImageIndex"), Integer)
        Me.btStep.ImeMode = CType(resources.GetObject("btStep.ImeMode"), System.Windows.Forms.ImeMode)
        Me.btStep.Location = CType(resources.GetObject("btStep.Location"), System.Drawing.Point)
        Me.btStep.Name = "btStep"
        Me.btStep.RightToLeft = CType(resources.GetObject("btStep.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.btStep.Size = CType(resources.GetObject("btStep.Size"), System.Drawing.Size)
        Me.btStep.TabIndex = CType(resources.GetObject("btStep.TabIndex"), Integer)
        Me.btStep.Text = resources.GetString("btStep.Text")
        Me.btStep.TextAlign = CType(resources.GetObject("btStep.TextAlign"), System.Drawing.ContentAlignment)
        Me.ToolTip1.SetToolTip(Me.btStep, resources.GetString("btStep.ToolTip"))
        Me.btStep.Visible = CType(resources.GetObject("btStep.Visible"), Boolean)
        '
        'btEdit
        '
        Me.btEdit.AccessibleDescription = resources.GetString("btEdit.AccessibleDescription")
        Me.btEdit.AccessibleName = resources.GetString("btEdit.AccessibleName")
        Me.btEdit.Anchor = CType(resources.GetObject("btEdit.Anchor"), System.Windows.Forms.AnchorStyles)
        Me.btEdit.BackgroundImage = CType(resources.GetObject("btEdit.BackgroundImage"), System.Drawing.Image)
        Me.btEdit.Dock = CType(resources.GetObject("btEdit.Dock"), System.Windows.Forms.DockStyle)
        Me.btEdit.Enabled = CType(resources.GetObject("btEdit.Enabled"), Boolean)
        Me.btEdit.FlatStyle = CType(resources.GetObject("btEdit.FlatStyle"), System.Windows.Forms.FlatStyle)
        Me.btEdit.Font = CType(resources.GetObject("btEdit.Font"), System.Drawing.Font)
        Me.btEdit.ForeColor = System.Drawing.Color.Orange
        Me.btEdit.Image = CType(resources.GetObject("btEdit.Image"), System.Drawing.Image)
        Me.btEdit.ImageAlign = CType(resources.GetObject("btEdit.ImageAlign"), System.Drawing.ContentAlignment)
        Me.btEdit.ImageIndex = CType(resources.GetObject("btEdit.ImageIndex"), Integer)
        Me.btEdit.ImeMode = CType(resources.GetObject("btEdit.ImeMode"), System.Windows.Forms.ImeMode)
        Me.btEdit.Location = CType(resources.GetObject("btEdit.Location"), System.Drawing.Point)
        Me.btEdit.Name = "btEdit"
        Me.btEdit.RightToLeft = CType(resources.GetObject("btEdit.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.btEdit.Size = CType(resources.GetObject("btEdit.Size"), System.Drawing.Size)
        Me.btEdit.TabIndex = CType(resources.GetObject("btEdit.TabIndex"), Integer)
        Me.btEdit.Text = resources.GetString("btEdit.Text")
        Me.btEdit.TextAlign = CType(resources.GetObject("btEdit.TextAlign"), System.Drawing.ContentAlignment)
        Me.ToolTip1.SetToolTip(Me.btEdit, resources.GetString("btEdit.ToolTip"))
        Me.btEdit.Visible = CType(resources.GetObject("btEdit.Visible"), Boolean)
        '
        'ImageListYesNo
        '
        Me.ImageListYesNo.ImageSize = CType(resources.GetObject("ImageListYesNo.ImageSize"), System.Drawing.Size)
        Me.ImageListYesNo.ImageStream = CType(resources.GetObject("ImageListYesNo.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.ImageListYesNo.TransparentColor = System.Drawing.Color.White
        '
        'ImageListOper
        '
        Me.ImageListOper.ImageSize = CType(resources.GetObject("ImageListOper.ImageSize"), System.Drawing.Size)
        Me.ImageListOper.ImageStream = CType(resources.GetObject("ImageListOper.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.ImageListOper.TransparentColor = System.Drawing.Color.Transparent
        '
        'NeedleContextMenu
        '
        Me.NeedleContextMenu.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem81, Me.MenuItem82, Me.MenuItem83, Me.MenuItem84, Me.MenuItem85, Me.MenuItem86})
        Me.NeedleContextMenu.RightToLeft = CType(resources.GetObject("NeedleContextMenu.RightToLeft"), System.Windows.Forms.RightToLeft)
        '
        'MenuItem81
        '
        Me.MenuItem81.Enabled = CType(resources.GetObject("MenuItem81.Enabled"), Boolean)
        Me.MenuItem81.Index = 0
        Me.MenuItem81.Shortcut = CType(resources.GetObject("MenuItem81.Shortcut"), System.Windows.Forms.Shortcut)
        Me.MenuItem81.ShowShortcut = CType(resources.GetObject("MenuItem81.ShowShortcut"), Boolean)
        Me.MenuItem81.Text = resources.GetString("MenuItem81.Text")
        Me.MenuItem81.Visible = CType(resources.GetObject("MenuItem81.Visible"), Boolean)
        '
        'MenuItem82
        '
        Me.MenuItem82.Enabled = CType(resources.GetObject("MenuItem82.Enabled"), Boolean)
        Me.MenuItem82.Index = 1
        Me.MenuItem82.Shortcut = CType(resources.GetObject("MenuItem82.Shortcut"), System.Windows.Forms.Shortcut)
        Me.MenuItem82.ShowShortcut = CType(resources.GetObject("MenuItem82.ShowShortcut"), Boolean)
        Me.MenuItem82.Text = resources.GetString("MenuItem82.Text")
        Me.MenuItem82.Visible = CType(resources.GetObject("MenuItem82.Visible"), Boolean)
        '
        'MenuItem83
        '
        Me.MenuItem83.Enabled = CType(resources.GetObject("MenuItem83.Enabled"), Boolean)
        Me.MenuItem83.Index = 2
        Me.MenuItem83.Shortcut = CType(resources.GetObject("MenuItem83.Shortcut"), System.Windows.Forms.Shortcut)
        Me.MenuItem83.ShowShortcut = CType(resources.GetObject("MenuItem83.ShowShortcut"), Boolean)
        Me.MenuItem83.Text = resources.GetString("MenuItem83.Text")
        Me.MenuItem83.Visible = CType(resources.GetObject("MenuItem83.Visible"), Boolean)
        '
        'MenuItem84
        '
        Me.MenuItem84.Enabled = CType(resources.GetObject("MenuItem84.Enabled"), Boolean)
        Me.MenuItem84.Index = 3
        Me.MenuItem84.Shortcut = CType(resources.GetObject("MenuItem84.Shortcut"), System.Windows.Forms.Shortcut)
        Me.MenuItem84.ShowShortcut = CType(resources.GetObject("MenuItem84.ShowShortcut"), Boolean)
        Me.MenuItem84.Text = resources.GetString("MenuItem84.Text")
        Me.MenuItem84.Visible = CType(resources.GetObject("MenuItem84.Visible"), Boolean)
        '
        'MenuItem85
        '
        Me.MenuItem85.Enabled = CType(resources.GetObject("MenuItem85.Enabled"), Boolean)
        Me.MenuItem85.Index = 4
        Me.MenuItem85.Shortcut = CType(resources.GetObject("MenuItem85.Shortcut"), System.Windows.Forms.Shortcut)
        Me.MenuItem85.ShowShortcut = CType(resources.GetObject("MenuItem85.ShowShortcut"), Boolean)
        Me.MenuItem85.Text = resources.GetString("MenuItem85.Text")
        Me.MenuItem85.Visible = CType(resources.GetObject("MenuItem85.Visible"), Boolean)
        '
        'MenuItem86
        '
        Me.MenuItem86.Enabled = CType(resources.GetObject("MenuItem86.Enabled"), Boolean)
        Me.MenuItem86.Index = 5
        Me.MenuItem86.Shortcut = CType(resources.GetObject("MenuItem86.Shortcut"), System.Windows.Forms.Shortcut)
        Me.MenuItem86.ShowShortcut = CType(resources.GetObject("MenuItem86.ShowShortcut"), Boolean)
        Me.MenuItem86.Text = resources.GetString("MenuItem86.Text")
        Me.MenuItem86.Visible = CType(resources.GetObject("MenuItem86.Visible"), Boolean)
        '
        'OpenPatternFileDialog
        '
        Me.OpenPatternFileDialog.Filter = resources.GetString("OpenPatternFileDialog.Filter")
        Me.OpenPatternFileDialog.Title = resources.GetString("OpenPatternFileDialog.Title")
        '
        'SavePatternFileDialog
        '
        Me.SavePatternFileDialog.Filter = resources.GetString("SavePatternFileDialog.Filter")
        Me.SavePatternFileDialog.Title = resources.GetString("SavePatternFileDialog.Title")
        '
        'ImageListLineArc
        '
        Me.ImageListLineArc.ImageSize = CType(resources.GetObject("ImageListLineArc.ImageSize"), System.Drawing.Size)
        Me.ImageListLineArc.ImageStream = CType(resources.GetObject("ImageListLineArc.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.ImageListLineArc.TransparentColor = System.Drawing.Color.Transparent
        '
        'TimerForUpdate
        '
        Me.TimerForUpdate.Interval = 1000
        '
        'SaveFileDialog1
        '
        Me.SaveFileDialog1.Filter = resources.GetString("SaveFileDialog1.Filter")
        Me.SaveFileDialog1.Title = resources.GetString("SaveFileDialog1.Title")
        '
        'IOCheck
        '
        Me.IOCheck.Interval = 50
        '
        'TowerLightImageList
        '
        Me.TowerLightImageList.ColorDepth = System.Windows.Forms.ColorDepth.Depth24Bit
        Me.TowerLightImageList.ImageSize = CType(resources.GetObject("TowerLightImageList.ImageSize"), System.Drawing.Size)
        Me.TowerLightImageList.ImageStream = CType(resources.GetObject("TowerLightImageList.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.TowerLightImageList.TransparentColor = System.Drawing.Color.Transparent
        '
        'FormProgramming
        '
        Me.AccessibleDescription = resources.GetString("$this.AccessibleDescription")
        Me.AccessibleName = resources.GetString("$this.AccessibleName")
        Me.AutoScaleBaseSize = CType(resources.GetObject("$this.AutoScaleBaseSize"), System.Drawing.Size)
        Me.AutoScroll = CType(resources.GetObject("$this.AutoScroll"), Boolean)
        Me.AutoScrollMargin = CType(resources.GetObject("$this.AutoScrollMargin"), System.Drawing.Size)
        Me.AutoScrollMinSize = CType(resources.GetObject("$this.AutoScrollMinSize"), System.Drawing.Size)
        Me.BackgroundImage = CType(resources.GetObject("$this.BackgroundImage"), System.Drawing.Image)
        Me.ClientSize = CType(resources.GetObject("$this.ClientSize"), System.Drawing.Size)
        Me.Controls.Add(Me.tbLsHeight)
        Me.Controls.Add(Me.tbOpenedFile)
        Me.Controls.Add(Me.ReferenceCommandBlock)
        Me.Controls.Add(Me.ElementsCommandBlock)
        Me.Controls.Add(Me.btExit)
        Me.Controls.Add(Me.btChangeSyringe)
        Me.Controls.Add(Me.GroupBox3)
        Me.Controls.Add(Me.gbConveyor)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.PanelToBeAdded)
        Me.Controls.Add(Me.ButtonToggleMode)
        Me.Controls.Add(Me.PanelVision)
        Me.Controls.Add(Me.PanelVisionCtrl)
        Me.Controls.Add(Me.CBExpandSpreadsheet)
        Me.Controls.Add(Me.LabelMessege)
        Me.Controls.Add(Me.AxSpreadsheetProgramming)
        Me.Controls.Add(Me.CBDoorLock)
        Me.Controls.Add(Me.ButtonHome)
        Me.Controls.Add(Me.ButtonNeedleCal)
        Me.Controls.Add(Me.ButtonClean)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.ButtonPurge)
        Me.Controls.Add(Me.ButtonVolCal)
        Me.Controls.Add(Me.btResetVolCalSettings)
        Me.Controls.Add(Me.Panel2)
        Me.Controls.Add(Me.Panel3)
        Me.Enabled = CType(resources.GetObject("$this.Enabled"), Boolean)
        Me.Font = CType(resources.GetObject("$this.Font"), System.Drawing.Font)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.ImeMode = CType(resources.GetObject("$this.ImeMode"), System.Windows.Forms.ImeMode)
        Me.Location = CType(resources.GetObject("$this.Location"), System.Drawing.Point)
        Me.MaximizeBox = False
        Me.MaximumSize = CType(resources.GetObject("$this.MaximumSize"), System.Drawing.Size)
        Me.Menu = Me.MainMenuProgramming
        Me.MinimizeBox = False
        Me.MinimumSize = CType(resources.GetObject("$this.MinimumSize"), System.Drawing.Size)
        Me.Name = "FormProgramming"
        Me.RightToLeft = CType(resources.GetObject("$this.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.StartPosition = CType(resources.GetObject("$this.StartPosition"), System.Windows.Forms.FormStartPosition)
        Me.Text = resources.GetString("$this.Text")
        Me.ToolTip1.SetToolTip(Me, resources.GetString("$this.ToolTip"))
        Me.PanelVisionCtrl.ResumeLayout(False)
        CType(Me.ValueBrightness, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.AxSpreadsheetProgramming, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel1.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        Me.gbConveyor.ResumeLayout(False)
        Me.GroupBox3.ResumeLayout(False)
        Me.Panel2.ResumeLayout(False)
        Me.Panel3.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

#Region "Winform start"

    Shared i As Integer = 0
    Dim j As Integer = 0
    Public gPatternFileName As String = ""
    Private iSubSheetName As String = ""
    Private SheetRowSelected As Boolean = False
    Private SheetColumnSelected As Boolean = False

    Private CopiedSheetName As String = ""
    Public ErrorSubSheet As SubSheetErrorStruct
    Public teachingMode As String = "Vision"
    Private m_NewProjectCreated As Boolean = False

    Public Shared logger As log4net.ILog ' for logging system
    Public Sub ErrorSubSheetStructIni(ByVal iMinItem, ByVal iMaxItem)
        'Change dimension of data structure
        ErrorSubSheet.ExtNumberOfSheets = 0
        ReDim ErrorSubSheet.IntAbstract(iMaxItem)
        ReDim ErrorSubSheet.ExtAbstract(iMaxItem)

        'Ini internal data structure
        Dim i As Integer
        For i = 0 To iMinItem - 1
            ErrorSubSheet.IntAbstract(i).SheetName = ""
            ErrorSubSheet.IntAbstract(i).SubFirstPt.X = 0
            ErrorSubSheet.IntAbstract(i).SubFirstPt.Y = 0
            ErrorSubSheet.IntAbstract(i).SubFirstPt.Z = 0
        Next


        'Ini external data structure
        For i = 0 To iMaxItem - 1
            ErrorSubSheet.ExtAbstract(i).SheetIndex = 0
            ErrorSubSheet.ExtAbstract(i).ExternalRefPt.X = 0
            ErrorSubSheet.ExtAbstract(i).ExternalRefPt.Y = 0
            ErrorSubSheet.ExtAbstract(i).ExternalRefPt.Z = 0
            ErrorSubSheet.ExtAbstract(i).RotationAngleRad = 0
            ErrorSubSheet.ExtAbstract(i).SheetLevel = 0
        Next
    End Sub

    Private Sub FormProgramming_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        KeyboardControl.GainControls()
        logger = log4net.LogManager.GetLogger("Pro")
        LogFile("Start Programming Mode")
        'AxSpreadsheetProgramming.Enabled = False
        MenuItem1.Enabled = False
        Init()
        While isInited = False
            Application.DoEvents()
        End While
        OffLaser()
        'reset!
        ResetToIdle()

        'display the stepping panel
        Controls.Add(m_Tri.SteppingButtons)
        m_Tri.SteppingButtons.BringToFront()
        m_Tri.SteppingButtons.Show()

        'change GUI settings unique for system setup here
        m_Tri.SteppingButtons.Location = PanelToBeAdded.Location()
        PanelVision.Controls.Add(Vision.FrmVision.PanelVision)

        ' get default value from the default pat file
        IDS.Data.ParameterID.RecordID = "FactoryDefault"
        IDSData.Admin.Folder.FileExtension = "Pat"
        IDSData.Admin.Folder.PatternPath = "C:\IDS\Pattern_Dir"
        IDS.Data.OpenData()
        SystemSetupDataRetrieve()
        m_Execution.m_Pattern.m_ErrorChk.GetErrorCheckParameter()

        'spreadsheet ini
        m_Execution.m_Pattern.SubCallSheetInitialization(200)     '200 subsheets maximum without duplicated name
        ErrorSubSheetStructIni(200, 500)    '500 subsheets maximum with duplicated name

        'error handling 
        Form_Service.ResetEventCode()

        'timers
        ThreadExecutor = New Threading.Thread(AddressOf StateChanger)
        ThreadExecutor.Priority = Threading.ThreadPriority.Normal
        ThreadExecutor.Start()

        ThreadMonitor = New Threading.Thread(AddressOf StateMonitor)
        ThreadMonitor.Priority = Threading.ThreadPriority.Normal
        ThreadMonitor.Start()

        MouseTimer = New System.Threading.Timer(AddressOf MouseJogging, Nothing, 0, 200)
        IDS.StartErrorCheck()
        IOCheck.Start()
        Conveyor.PositionTimer.Start()
        Me.Text = "Programming"

        'flags
        CurrentMode = "Program Editor"
        gExeMode = IDS.Data.Admin.User.RunApplication
        InitialPatternParameterSetup()
        m_NewProjectCreated = False

        'hardware
        Laser.OpenPort()
        Weighting_Scale.OpenPort()
        Conveyor.OpenPort()
        Dispenser.OpenPort()
        Conveyor.Command("Manual Mode")
        'motion controller
        'm_Tri.Connect_Controller()
        SetState("Homing")
        'vision
        SwitchToTeachCamera()
        ValueBrightness.Value = CDec(IDS.Data.Hardware.Camera.Brightness)
        'Disable part of the menu in File (GUI)
        MenuFileExport.Enabled = False
        MenuFileImport.Enabled = True
        MenuFileSave.Enabled = False
        MenuFileSaveAs.Enabled = False
        MenuEditUndo.Enabled = False
        MenuEditRedo.Enabled = False
        MenuEditCopy.Enabled = False
        MenuEditPaste.Enabled = False
        MenuEditCut.Enabled = False
        MenuEditDelete.Enabled = False
        MenuEditSelectAll.Enabled = False
        DisableTeachingToolbarOKButton()
        DisableTeachingToolbarCancelButton()
        DisableEditingToolbar()
        DisableTeachingButtons()

        'gui
        Me.Location = New Point(0, 0)
        Me.StartPosition = FormStartPosition.Manual
        Me.WindowState = FormWindowState.Maximized
        'EnableTeachingButtons()
        DisableElementsCommandBlockButton(10)
        DisableTeachingButtons()

    End Sub

    Private Sub FormProgramming_Closed(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        KeyboardControl.ReleaseControls()
        isInited = False
        Dim response As MsgBoxResult
        If (m_NewProjectCreated) And (SaveProgram.UnSave = True) Then
            response = MyMsgBox("Do you want to save the file before exit?", MsgBoxStyle.YesNo)
            If response = MsgBoxResult.Yes Then
                Local_SaveFile()
            End If
            SaveProgram.UnSave = False
        End If

        'error handling
        Form_Service.ResetEventCode()

        'vision
        SwitchToTeachCamera()
        IDS.Data.Hardware.Camera.Brightness = ValueBrightness.Value
        Vision.IDSV_SetBrightness(0)
        'motion controller
        'm_Tri.Move_Z(0)
        m_Tri.TrioStop()
        m_Tri.m_TriCtrl.Op(0)
        m_Tri.RunTrioMotionProgram("EXITIDS")
        m_Tri.TurnOff("Material Air")

        'timers stop
        IDS.StopErrorCheck()
        IOCheck.Stop()
        TimerForUpdate.Stop()
        MouseTimer.Dispose()
        ThreadMonitor.Abort()
        ThreadExecutor.Abort()

        InitThread.Abort()

        'hardware shutdown
        Conveyor.PositionTimer.Stop()
        MyConveyorSettings.CloseConveyorSetup()
        Conveyor.ClosePort()
        Weighting_Scale.ClosePort()
        OffLaser()
        Laser.ClosePort()
        OffTowerLamp()
        UnlockDoor()

        'gui
        PanelVision.Controls.Remove(Vision.FrmVision.PanelVision)
        m_Tri.Disconnect_Controller()

        IDS.Data.SaveData()
        IDS.FrmExecution.Hide()

    End Sub

    Public Function LogFile(ByVal message As String, Optional ByVal logLevel As Integer = 0)
        If message = "" Then
            Exit Function
        End If
        If logLevel = 0 Then
            logger.Info(message)
        ElseIf logLevel = 1 Then
            logger.Debug(message)
        ElseIf logLevel = 2 Then
            logger.Error(message)
        End If
    End Function

#End Region

    Public LaserHeightOffsetZ As Double = -1234
    'kr make public and exposed.
    Public CurrentMode As String 'flag for programming mode - program editor or basic setup
    Public tempPosX, tempPosY, tempPosZ As Double              'Save X and Y cells for "Cancel Process".  
    Public openEventViewer As Boolean = False                  'Flag for eventViewer not to update main UI.
    Public m_MachinePos(3) As Double
    Public DeletingRowFromExcel As Boolean = False             'Flag system is deleting program(row of excel)
    Public DeletingRowFinished As Boolean = False              'Flag for one cycle
    Public m_ReferPt() As Double = {0.0, 0.0, 0.0}
    Public m_BoardnRefBlkDist As Double = 0.0
    Public m_TeachMode As Integer = 0 '0:vision; 1:Left; 2: Right
    Public m_RunMode As Integer = 0 'runmode:  0-vision 1-dry 2-dry left 3-dry right 4-wet 5-dry left 6-dry right 
    Public Property GlobalQCEnabled() As Boolean
        Get
            Return IDS.Data.Hardware.Camera.DotQCEnable
        End Get
        Set(ByVal Value As Boolean)
            IDS.Data.Hardware.Camera.DotQCEnable = Value
        End Set
    End Property

#Region "Shen Jian"

    Private isJogON As Boolean = False
    Protected m_CamPos(3) As Double
    Private m_Xlocked As Boolean = False
    Private m_Ylocked As Boolean = False
    Private m_Zlocked As Boolean = False
    Public cell1, cell2 As Object

    Private m_SelectedType As String = ""
    Dim m_TrackBall As New DLL_Export_Device_Motor.Mouse(Me)
    Dim m_keyBoard As New DLL_Export_Device_Motor.Keyboard(Me)
    Public MouseTimer As System.Threading.Timer
    Private Declare Function InterlockedExchange Lib "kernel32" (ByRef Target As Integer, ByVal Value As Integer) As Integer

    Private GDRefX As Integer = 7 'lim 'to offset the negative value of the robot's position
    Private GDRefY As Integer = 305 'lim 'to offset the negative value of the robot's position

    Private Sub ButtonHome_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonHome.Click
        forceHome = True
        SetState("Homing")
    End Sub

    Private Sub ButtonPurge_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonPurge.Click
        If SetState("Purge") Then DoPurge()
    End Sub

    Private Sub ButtonClean_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonClean.Click
        If SetState("Clean") Then DoClean()
    End Sub

    Private Sub ButtonNeedleCal_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonNeedleCal.Click
        If ButtonNeedleCal.Text = "Need. Cal." Then
            SetState("Needle Calibration")
            ButtonNeedleCal.Text = "Stop Need. Cal."
        Else
            ButtonNeedleCal.Enabled = False
            StopDispensing()
        End If

    End Sub

    Private Sub ButtonVolCal_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonVolCal.Click
        'If gPatternFileName = "" Or gPatternFileName = Nothing Then
        '    MessageBox.Show("Cannot perform volume calibration. Please save the table to file or open/create a file")
        '    Return
        'End If
        If ButtonVolCal.Text = "Vol. Cal." Then
            SetState("Volume Calibration")
            ButtonVolCal.Text = "Stop Vol. Cal."
        Else
            ButtonVolCal.Enabled = False
            StopDispensing()
        End If

        'DoVolumeCalibration()
    End Sub
    Private zSafetyThreshold As Double = -3.0
    Private Sub VisionMode_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles VisionMode.Click

        If IsBusy() Then
            LabelMessage("Can't change mode when machine is running!")
            Exit Sub
        End If

        VisionMode.Checked = True
        NeedleMode.Checked = False
        'SwitchToTeachCamera()
        IDS.Devices.Vision.FrmVision.CameraResume()
        'If Not (m_ProgrammingStateFlag Or m_SteppingPostFlag) Then
        '    EnableReferenceCommandBlock()
        '    EnableElementsCommandBlock()
        'End If

        If m_Tri.ZPosition < zSafetyThreshold Then
            MoveZToSafePosition()
        End If
    End Sub

    Private Sub NeedleMode_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NeedleMode.Click
        If IsBusy() Then
            LabelMessage("Can't change mode when machine is running!")
            Exit Sub
        End If
        VisionMode.Checked = False
        NeedleMode.Checked = True
        'DisableElementsCommandBlockButton(gQCCmdIndex) 'QC
        'DisableElementsCommandBlockButton(gChipEdgeCmdIndex) 'ChipEdge
        'DisableReferenceCommandBlock()
        'EnableReferenceCommandBlockButton(gHeightCmdIndex) 'Height
        'If Not (m_ProgrammingStateFlag Or m_SteppingPostFlag) Then
        'DisableCommand_NeedleMode()
        'End If
        IDS.Devices.Vision.FrmVision.CameraIdle()
        If m_Tri.ZPosition < zSafetyThreshold Then
            MoveZToSafePosition()
        End If
    End Sub
    Public Sub DisableCommand_NeedleMode()
        DisableElementsCommandBlockButton(gQCCmdIndex) 'QC
        DisableElementsCommandBlockButton(gChipEdgeCmdIndex) 'ChipEdge
        DisableReferenceCommandBlock()
        EnableReferenceCommandBlockButton(gHeightCmdIndex) 'Height
    End Sub

    'SJ add door lock/unlock
    Private Sub CheckBox5_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CBDoorLock.CheckedChanged

        If IsBusy() Then
            LabelMessage("Can't unlock the door when machine running!")
            Exit Sub
        End If
        If CBDoorLock.Text = "Lock Door" Then
            CBDoorLock.Text = "Unlock Door"
            'CBDoorLock.ImageIndex = 6
            'TraceDoEvents()
            LockDoor()
        Else
            CBDoorLock.Text = "Lock Door"
            'CBDoorLock.ImageIndex = 5
            ''TraceDoEvents()
            UnlockDoor()

        End If
    End Sub

    Private Sub CheckBoxLockX_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBoxLockX.CheckedChanged

        If IsBusy() Then
            LabelMessage("Can't lock axis when machine running!")
            Exit Sub
        End If
        If CheckBoxLockX.Checked = True Then
            m_Tri.SetTrioMotionValues("Lock X")
            m_Xlocked = True
        Else
            m_Tri.SetTrioMotionValues("Unlock X")
            m_Xlocked = False
        End If

    End Sub


    Private Sub CheckBoxLockY_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBoxLockY.CheckedChanged

        If IsBusy() Then
            LabelMessage("Can't lock axis when machine running!")
            Exit Sub
        End If
        If CheckBoxLockY.Checked = True Then
            m_Tri.SetTrioMotionValues("Lock Y")
            m_Ylocked = True
        Else
            m_Tri.SetTrioMotionValues("Unlock Y")
            m_Ylocked = False
        End If

    End Sub

    Private Sub CheckBoxLockZ_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBoxLockZ.CheckedChanged

        If IsBusy() Then
            LabelMessage("Can't lock axis when machine running!")
            Exit Sub
        End If
        If CheckBoxLockZ.Checked = True Then
            m_Tri.SetTrioMotionValues("Lock Z")
            m_Zlocked = True
        Else
            m_Tri.SetTrioMotionValues("Unlock Z")
            m_Zlocked = False
        End If
    End Sub

    Private Sub MoveToSpreadsheetPoint(ByVal Pos() As Double, ByVal type As String, Optional ByVal needleGap As Double = 0.0, Optional ByVal selectedRow As Double = 1)
        Console.WriteLine("MoveTo SP Point")
        If IsBusy() Then
            LabelMessage("Axes are busy. Can't move at the moment. Try later")
            Exit Sub
        End If

        m_Tri.Set_XY_Speed(IDS.Data.Hardware.Gantry.ElementXYSpeed)
        m_Tri.Set_Z_Speed(IDS.Data.Hardware.Gantry.ElementZSpeed)

        Dim ref(3), inmmPos(3) As Double
        ref(0) = m_ReferPt(0)
        ref(1) = m_ReferPt(1)
        ref(2) = m_ReferPt(2)

        ReferToSys(Pos, Pos, ref)
        SysToHard(Pos, Pos)

        If type = "Needle" And NeedleMode.Checked Then
            Pos(0) = Pos(0) - gLeftNeedleOffs(0)
            Pos(1) = Pos(1) - gLeftNeedleOffs(1)
            LaserHeightOffsetZ = -1234
            If LaserHeightOffsetZ = -1234 Then
                If teachingMode = "Vision" Then
                    If GetExistingLaserHeightDifferent(LaserHeightOffsetZ, selectedRow) Then
                        Pos(2) = Pos(2) + IDS.Data.Hardware.Needle.Left.NeedleCalibrationPosition.Z - LaserHeightOffsetZ + needleGap
                    Else
                        Dim fm As InfoForm = New InfoForm
                        fm.SetTitle("Warning")
                        fm.OkOnly()
                        fm.AddNewLine("Please insert laser height command before run needle mode")
                        fm.ShowDialog()
                        Return
                    End If
                End If
            Else
                Pos(2) = Pos(2) + IDS.Data.Hardware.Needle.Left.NeedleCalibrationPosition.Z - LaserHeightOffsetZ + needleGap
            End If
        End If

        If Not m_Tri.Move_Z(SafePosition) Then Exit Sub
        If VisionMode.Checked Then
            If Not m_Tri.Move_XY(Pos) Then Exit Sub
        ElseIf NeedleMode.Checked Then
            If Not m_Tri.Move_XY(Pos) Then Exit Sub
            If Not m_Tri.Move_Z(Pos(2)) Then Exit Sub
        End If
    End Sub

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

    Private Sub MouseJogging(ByVal state As Object)

        If IsBusy() And Not IsJogging() Then Exit Sub
        If m_Tri.MachineHoming Or m_Tri.MachineRunning Or m_Tri.Calibrating Or m_Tri.Stepping Then Exit Sub

        'm_keyBoard.Poll()
        m_TrackBall.Poll()
        'isPress = m_keyBoard.State.Item(Key.LeftControl)
        isPress = KeyboardControl.ControlKeyPressed
        Dim x As Integer
        Dim y As Integer
        Dim VrData(3) As Single

        If isPress Then

            SetState("Jogging")

            VrData(0) = 0
            VrData(1) = 0.0
            VrData(2) = 0.0

            'Dim isPressAlt As Boolean = m_keyBoard.State.Item(Key.LeftAlt)
            'If isPressAlt Then
            '    Exit Sub
            'End If
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
                        If m_Xlocked = True Then
                            VrData(1) = 0.0
                        Else
                            VrData(1) = jogspeed
                        End If
                        If m_Ylocked = True Then
                            VrData(2) = 0.0
                        Else
                            VrData(2) = -jogspeed
                        End If

                        m_Tri.SetTrioMotionValues("Jogging", VrData)
                        isJogON = True
                    ElseIf (ratio < -ratioLB) And (ratio > -ratioUB) Then 'X+ Y+
                        VrData(0) = 1
                        If m_Xlocked = True Then
                            VrData(1) = 0.0
                        Else
                            VrData(1) = jogspeed
                        End If
                        If m_Ylocked = True Then
                            VrData(2) = 0.0
                        Else
                            VrData(2) = jogspeed
                        End If

                        m_Tri.SetTrioMotionValues("Jogging", VrData)
                        isJogON = True
                    Else   'X+
                        VrData(0) = 1
                        If m_Xlocked = True Then
                            VrData(1) = 0.0
                        Else
                            VrData(1) = jogspeed
                        End If
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
                        If m_Xlocked = True Then
                            VrData(1) = 0.0
                        Else
                            VrData(1) = -jogspeed
                        End If
                        If m_Ylocked = True Then
                            VrData(2) = 0.0
                        Else
                            VrData(2) = jogspeed
                        End If

                        m_Tri.SetTrioMotionValues("Jogging", VrData)
                        isJogON = True
                    ElseIf (ratio < -ratioLB) And (ratio > -ratioUB) Then 'X- Y-
                        VrData(0) = 1
                        If m_Xlocked = True Then
                            VrData(1) = 0.0
                        Else
                            VrData(1) = -jogspeed
                        End If
                        If m_Ylocked = True Then
                            VrData(2) = 0.0
                        Else
                            VrData(2) = -jogspeed
                        End If

                        m_Tri.SetTrioMotionValues("Jogging", VrData)
                        isJogON = True
                    Else   'X-
                        VrData(0) = 1
                        If m_Xlocked = True Then
                            VrData(1) = 0.0
                        Else
                            VrData(1) = -jogspeed
                        End If
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
                        If m_Xlocked = True Then
                            VrData(1) = 0.0
                        Else
                            VrData(1) = -jogspeed
                        End If
                        If m_Ylocked = True Then
                            VrData(2) = 0.0
                        Else
                            VrData(2) = jogspeed
                        End If

                        SetState("Jogging")
                        m_Tri.SetTrioMotionValues("Jogging", VrData)
                        isJogON = True
                    ElseIf (ratio < -ratioLB) And (ratio > -ratioUB) Then 'X+ Y+
                        VrData(0) = 1
                        If m_Xlocked = True Then
                            VrData(1) = 0.0
                        Else
                            VrData(1) = jogspeed
                        End If
                        If m_Ylocked = True Then
                            VrData(2) = 0.0
                        Else
                            VrData(2) = jogspeed
                        End If

                        m_Tri.SetTrioMotionValues("Jogging", VrData)
                        isJogON = True
                    Else   'Y+
                        VrData(0) = 1
                        VrData(1) = 0
                        If m_Ylocked = True Then
                            VrData(2) = 0.0
                        Else
                            VrData(2) = jogspeed
                        End If

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
                        If m_Xlocked = True Then
                            VrData(1) = 0.0
                        Else
                            VrData(1) = jogspeed
                        End If
                        If m_Ylocked = True Then
                            VrData(2) = 0.0
                        Else
                            VrData(2) = -jogspeed
                        End If

                        m_Tri.SetTrioMotionValues("Jogging", VrData)
                        isJogON = True
                    ElseIf (ratio < -ratioLB) And (ratio > -ratioUB) Then 'X- Y-
                        VrData(0) = 1
                        If m_Xlocked = True Then
                            VrData(1) = 0.0
                        Else
                            VrData(1) = -jogspeed
                        End If
                        If m_Ylocked = True Then
                            VrData(2) = 0.0
                        Else
                            VrData(2) = -jogspeed
                        End If

                        m_Tri.SetTrioMotionValues("Jogging", VrData)
                        isJogON = True
                    Else   'Y-
                        VrData(0) = 1
                        VrData(1) = 0
                        If m_Ylocked = True Then
                            VrData(2) = 0.0
                        Else
                            VrData(2) = -jogspeed
                        End If

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
                    If ProgrammingMode() Then
                        Programming.DispensingMode.Enabled = True
                        Programming.VisionMode.Enabled = True
                        Programming.NeedleMode.Enabled = True
                    End If
                Else
                    ResetToIdle()
                End If
            End If

            countMouseTimer += 1
            If (countMouseTimer >= 5) Then
                TraceGCCollect()
                countMouseTimer = 0
            End If
        End If

    End Sub

    Private Function CheckSoftLimitXYZ(ByVal pt() As Double, ByVal off() As Double) 'check whether to apply offset for ref/needle point when teaching
        Dim x, y, z As Double
        If VisionMode.Checked Then 'no offset, for refer
            x = pt(0)
            y = pt(1)
            z = pt(2)
        Else
            x = pt(0) - off(0)
            y = pt(1) - off(1)
            z = pt(2)
        End If

        '''''''''''''''''''''''''''''''''''''
        'Check(SoftLimit)                '
        '   Note: enumation for "status"    '
        '         1 = under Xmin            '
        '         2 = over Xmax             '
        '         3 = under Ymin            '
        '         4 = over Ymax             '
        '         5 = under Zmin            '
        '         6 = over Zmax             '
        '''''''''''''''''''''''''''''''''''''
        If (x < gWorkLimitXmin) Then
            MyMsgBox("Outside X limits")
        ElseIf (x > gWorkLimitXmax) Then
            MyMsgBox("Outside X limits")
        ElseIf (y < gWorkLimitYmin) Then
            MyMsgBox("Outside Y limits")
        ElseIf (y > gWorkLimitYmax) Then
            MyMsgBox("Outside Y limits")
        ElseIf (z < gWorkLimitZmin) Then
            MyMsgBox("Outside Z limits")
        ElseIf (z > gWorkLimitZmax) Then
            MyMsgBox("Outside Z limits")
        Else
            Return False
        End If
        Return True

    End Function

    Private Sub ErrorMessageBox(ByVal errVal As Integer)
        If (errVal = 1) Then
            LabelMessage("X-axis is over minimum of SoftLimit!")
        ElseIf (errVal = 2) Then
            LabelMessage("X-axis is over maximum of SoftLimit!")
        ElseIf (errVal = 3) Then
            LabelMessage("Y-axis is over minimum of SoftLimit!")
        ElseIf (errVal = 4) Then
            LabelMessage("Y-axis is over maximum of SoftLimit!")
        ElseIf (errVal = 5) Then
            LabelMessage("Z-axis is over minimum of SoftLimit!")
        ElseIf (errVal = 6) Then
            LabelMessage("Z-axis is over maximum of SoftLimit!")
        End If
    End Sub

    Public Sub SetRunMode(ByVal mode As Integer)
        m_RunMode = mode
    End Sub

    Dim itemIndex As Integer

    Private Sub ComboBox1_SelectedValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DispensingMode.SelectedValueChanged
        '0-vision 1-dry 2-dry left 3-dry right 4-wet 5-dry left 6-dry right  
        itemIndex = DispensingMode.SelectedIndex
        If itemIndex = 0 Then
            m_RunMode = 0
        ElseIf itemIndex = 1 Then
            m_RunMode = 1
        ElseIf itemIndex = 2 Then
            m_RunMode = 4
        End If
    End Sub

#End Region

    Private Sub ValueBrightness_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ValueBrightness.TextChanged
        If 0 <= ValueBrightness.Value < 255 Then
            Vision.IDSV_SetBrightness(ValueBrightness.Value)
        Else
            ValueBrightness.Value = 128
            Vision.IDSV_SetBrightness(ValueBrightness.Value)
        End If
    End Sub

#Region "Jiang"
    Private Sub MenuItem88_Click(ByVal sender As System.Object, _
        ByVal e As System.EventArgs) Handles MenuItem88.Click
        Dim FrmEventViewer As New EventViewer
        IDS.StopErrorCheck()
        openEventViewer = True

        FrmEventViewer.ShowDialog()

        IDS.StartErrorCheck()
        openEventViewer = False
    End Sub

    'FlushSpreadsheet activeX spreadsheet pattern data           
    '   Axspreadsheet:  instance of activeX spreadsheet  

    Private Sub FlushSpreadsheet()

        Dim TotalSheets As Integer = GetSheetCount()
        Dim i As Integer
        Dim SheetName As String

        For i = 1 To TotalSheets
            SheetName = AxSpreadsheetProgramming.Worksheets(1).name()

            If "Main" <> SheetName Then
                AxSpreadsheetProgramming.Worksheets(1).Delete()
            Else
                AxSpreadsheetProgramming.Worksheets(SheetName).Activate()
                AxSpreadsheetProgramming.ActiveSheet.UsedRange.Clear()
            End If
        Next i

        m_Execution.m_Pattern.SubCallSheetInitialization(200) '200 subsheet maximum
        m_Row = 1
        TraceGCCollect()
    End Sub




    'menu: File --> New: Create a new project        
    '                                                                               

    Private Sub MenuFileNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuFileNew.Click

        Try
            Dim response As MsgBoxResult
            'Handling the case with a projct already exists.
            If "" <> gPatternFileName Then
                '''''''''''''''''''''''''''''''''''''
                '   Xue Wen                         '
                '   check and save old program.     '
                '''''''''''''''''''''''''''''''''''''
                If (SaveProgram.UnSave = True) Then
                    response = MyMsgBox("Save current program before creating the new one.", MsgBoxStyle.YesNo, "Save Program")

                    If (MsgBoxResult.Yes = response) Then
                        Local_SaveFile()
                        SaveProgram.UnSave = False
                    End If
                End If

                response = MyMsgBox("Do you want to create a new project?", MsgBoxStyle.Question + MsgBoxStyle.YesNo)

                If response = MsgBoxResult.No Then
                    Return
                End If

            End If

            'Set defauly data for the DialogBox
            m_NewProjectCreated = True
            'SavePatternFileDialog.InitialDirectory = "C:\IDS\Pattern_Dir"
            'SavePatternFileDialog.AddExtension = True
            'SavePatternFileDialog.Filter = "Create a new project folder|"
            'SavePatternFileDialog.FileName() = ""

            'SavePatternFileDialog.ShowDialog()
            Dim fm As NewFileForm = New NewFileForm
            If fm.ShowDialog() = DialogResult.Cancel Then Return
            teachingMode = fm.teachingMode
            m_Execution.m_Pattern.teachingMode = teachingMode
            Dim file As String
            Dim folder As String = "C:\IDS\Pattern_Dir\" + fm.NewFileName 'SavePatternFileDialog.FileName()

            If folder <> Nothing Then

                If -1 <> folder.IndexOf(".") Then
                    'Project name cannot contain "."
                    MyMsgBox("Invalid project name!, Press Ok to return", _
                        MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly, "Create new project error")
                    Return
                ElseIf System.IO.Directory.Exists(folder) Then
                    'Project name is the same with an existing filename
                    'The file will be deleted
                    'If MessageBox.Show("Project folder exist! Do you want to delete all and overwrite it? Click no to abort.", "Warning", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) = DialogResult.Yes Then
                    System.IO.Directory.Delete(folder, True)
                    'Else : Return
                    'End If

                    'Create folder for the project
                End If
                Me.GlobalQCEnabled = False
                System.IO.Directory.CreateDirectory(folder)
                file = m_Execution.m_File.NameOnlyFromFullPath(folder)
                gPatternFileName = folder + "\" + file
                gFidFileName = gPatternFileName
                'Set ToolButton status
                If teachingMode = "Vision" Then
                    EnableTeachingButtons()
                    ' EnableElementsCommandBlockButton(gGlobalQCCmdIndex)
                Else
                    EnableTeachingButtons()
                    DisableCommand_NeedleMode()
                End If

                DisableElementsCommandBlockButton(gSeperatorCmdIndex)
                DisableElementsCommandBlockButton(gOffsetCmdIndex)
                FlushSpreadsheet()


                UndoData_Logging(0)

                'Set menu status
                MenuFileExport.Enabled = True
                MenuFileImport.Enabled = True
                MenuFileSave.Enabled = True
                MenuFileSaveAs.Enabled = True

                MenuEditUndo.Enabled = False
                MenuEditRedo.Enabled = False
                MenuEditCopy.Enabled = False
                MenuEditPaste.Enabled = False
                MenuEditCut.Enabled = False
                MenuEditDelete.Enabled = False
                MenuEditSelectAll.Enabled = True

                'lsgoh
                ' get default value from the default pat file
                IDS.Data.ParameterID.RecordID = "FactoryDefault"
                IDSData.Admin.Folder.FileExtension = "Pat"
                IDSData.Admin.Folder.PatternPath = "C:\IDS\Pattern_Dir"
                IDS.Data.OpenData()
                IDS.Data.SavePathFileData(gPatternFileName + ".pat")
                'lim
                Vision.SetFiducialFilename(gPatternFileName)
                tbOpenedFile.Text = gPatternFileName
                m_EditStateFlag = False
                MyConveyorSettings.InitializeConveyorSetup()
                Disp_Dispenser_Unit_info()

                SystemSetupDataRetrieve() 'SJ add 

                '   Set cursor position.                        '
                SelectCell(m_Row, 1)
                SaveProgram.UnSave = True
                EnableControl()
                LaserHeightOffsetZ = -1234
                m_Execution.m_Pattern.SavePatternPara(AxSpreadsheetProgramming, gPatternFileName + ".txt", False)

            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub
    Private Sub EnableControl()
        AxSpreadsheetProgramming.Enabled = True
        ButtonNeedleCal.Enabled = True
        ButtonVolCal.Enabled = True
    End Sub

    'Private function to export Excel data file,  to be called by MenuFileExport_Click
    '   file:  filename name for the export                                                                                

    Private Sub ExportExcelPatternFile(ByVal file As String)

        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor
        Dim SheetNamerArray(200) As String
        Dim SheetName As String

        Dim i, arrayDim As Integer
        Dim NumberOfSheets As Integer = AxSpreadsheetProgramming.Worksheets.Count()

        For i = 1 To NumberOfSheets
            SheetNamerArray(i - 1) = AxSpreadsheetProgramming.ActiveWorkbook.Worksheets(i).name()
        Next

        'Found all qualified subPage partialFileName.  This part is not quite necessory
        For i = 1 To NumberOfSheets

            SheetName = SheetNamerArray(i - 1)
            arrayDim = 1
            For j = 0 To Len(SheetName) - 1
                If "." = SheetName.Chars(j) Then
                    arrayDim += 1
                End If
            Next

            If 3 = arrayDim Or 2 = arrayDim Then
                AxSpreadsheetProgramming.Worksheets(SheetName).Delete()
            End If
        Next


        If file <> Nothing Then
            AxSpreadsheetProgramming.Worksheets("Main").Activate()

            SelectCell(1, gCommandNameColumn)
            AxSpreadsheetProgramming.Selection.EntireRow.Insert()
            SetCellValue(1, gCommandNameColumn, "Unit")
            AxSpreadsheetProgramming().Export(file, _
                OWC.SheetExportActionEnum.ssExportActionNone, OWC.SheetExportFormat.ssExportAsAppropriate)
            DeleteRow(1)
        End If

        Me.Cursor = System.Windows.Forms.Cursors.Default

        Dim filename As String
        Dim usedRange As Integer = AxSpreadsheetProgramming.Worksheets("Main").UsedRange.Rows.Count

        With AxSpreadsheetProgramming.Worksheets("Main")
            For i = 1 To usedRange
                If "SubPattern" = CStr(.Cells(i, gCommandNameColumn).value) Then
                    filename = CStr(.Cells(i, gSubnameColumn).value)
                    m_Execution.m_Pattern.LoadTxtPatternPara(AxSpreadsheetProgramming, filename, 1, 0, True)

                End If
            Next
        End With

        AxSpreadsheetProgramming.Worksheets("Main").Activate()
        TraceGCCollect()

    End Sub


    'Private function to export Pattern data file, to be called by MenuFileExport_Click
    '   file:  filename name for the export                                                                                

    Private Sub ExportTxtPatternFile(ByVal file As String)

        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor

        Try
            If file <> Nothing Then
                m_Execution.m_Pattern.SavePatternPara(AxSpreadsheetProgramming, file, False)
            End If
        Catch ex As Exception
        End Try

        Me.Cursor = System.Windows.Forms.Cursors.Default

    End Sub



    'menu: File-->Export, Private function to export data file
    '   file:  filename name for the export                                                                                

    Private Sub MenuFileExport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuFileExport.Click
        Dim strTmp(4) As String

        'should get from data mamager 
        SavePatternFileDialog.InitialDirectory = "C:\IDS\Pattern_Dir"
        SavePatternFileDialog.AddExtension = True
        SavePatternFileDialog.Filter = "Excel Pattern files (*.xls)|*.xls|Txt Pattern files (*.txt)|*.txt"
        SavePatternFileDialog.FileName = ""
        SavePatternFileDialog.ShowDialog()

        Dim file As String = SavePatternFileDialog.FileName()


        'OpenPatternFileDialog.ShowDialog()
        'Dim file As String = OpenPatternFileDialog.FileName()
        'If Nothing = file Then
        'Return
        'End If



        'savepatternfiledialog.

        If file <> Nothing Then
            strTmp = file.Split(".")
            If "xls" = strTmp(1).ToLower() Then
                ExportExcelPatternFile(file)
            ElseIf "txt" = strTmp(1).ToLower() Then
                ExportTxtPatternFile(file)
            End If
        End If

    End Sub



    'Import Excel pattern file          
    '   file:  input filename                                                                             

    Private Sub ImportExcelPatternFile(ByVal file As String)    'Internal function
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor
        gPatternFileName = ""
        gFidFileName = ""

        Try
            AxSpreadsheetProgramming.Caption = file
            m_Execution.m_Pattern.LoadExcelPatternFile(AxSpreadsheetProgramming, file, 1, False)
            m_Row = 2

            'Error checking for all the Spreadsheet
            If 0 <> m_Execution.m_Pattern.m_ErrorChk.CheckAllError(AxSpreadsheetProgramming, ErrorSubSheet) Then
                MyMsgBox("Click Ok to flush all the data", MsgBoxStyle.OKOnly, "Error found")

                FlushSpreadsheet()
                ErrorSubSheetStructIni(200, 500)
                gPatternFileName = ""
                gFidFileName = ""
                init_spreadsheet()
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

        Me.Cursor = System.Windows.Forms.Cursors.Default
        TraceGCCollect()
    End Sub



    'Import text pattern file          
    '   file:  input filename                                                                             

    Private Sub ImportTxtPatternFile(ByVal file As String)
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor
        gPatternFileName = ""
        gFidFileName = ""

        Try
            AxSpreadsheetProgramming.Caption = file
            'If 0 = m_Execution.m_Pattern.LoadTxtPatternPara(AxSpreadsheetProgramming, file, 2, m_Row, False) Then
            If 0 = m_Execution.m_Pattern.LoadTxtPatternPara(AxSpreadsheetProgramming, file, 2, 0, False) Then
                m_Row = 2

                'Error checking for all the Spreadsheet
                If 0 <> m_Execution.m_Pattern.m_ErrorChk.CheckAllError(AxSpreadsheetProgramming, ErrorSubSheet) Then
                    MyMsgBox("Click Ok to flush all the data", MsgBoxStyle.OKOnly, "Error found")

                    FlushSpreadsheet()
                    ErrorSubSheetStructIni(200, 500)
                    gPatternFileName = ""
                    gFidFileName = ""
                    init_spreadsheet()
                End If
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
        Me.Cursor = System.Windows.Forms.Cursors.Default
        TraceGCCollect()
    End Sub



    'menu: File-->Import (from an excel/text file)
    '   file:                                                                                  

    Private Sub MenuFileImport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuFileImport.Click

        'should get from data mamager 
        OpenPatternFileDialog.InitialDirectory = "C:\IDS\Pattern_Dir"
        OpenPatternFileDialog.AddExtension = True
        OpenPatternFileDialog.Filter = "Excel pattern files (*.xls)|*.xls|Txt pattern files(*.txt)|*.txt"
        OpenPatternFileDialog.FileName = ""

        OpenPatternFileDialog.ShowDialog()
        Dim file As String = OpenPatternFileDialog.FileName()
        If Nothing = file Then
            Return
        End If

        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor

        Dim strInfo(4) As String
        strInfo = file.Split(".")

        m_NewProjectCreated = True
        'FlushSpreadsheet()
        'init_spreadsheet()

        'If "XLS" = strInfo(1).ToUpper() Then
        '    ImportExcelPatternFile(file)
        'ElseIf "ptp" = strInfo(1) Then
        ImportTxtPatternFile(file)
        'gPatternFileName = file
        'End If
        Me.Cursor = System.Windows.Forms.Cursors.Default
        MenuFileSaveAs.Enabled = True
        MenuFileSave.Enabled = True
        TraceGCCollect()
    End Sub

    Public Function IsFilePresent() As Boolean
        If ProductionMode() And Production.TextBoxFilename.Text = "" Then Return False
        Return m_Execution.m_Pattern.m_ErrorChk.CheckAllError(AxSpreadsheetProgramming, ErrorSubSheet) <> 0
    End Function


    'Load an encrypt text pattern file for production //SJ add
    '   
    '   Return: 0=success

    Public Function ProductionFileOpen() As Integer
        Dim ErrorMsg As String
        Dim Rtn As MsgBoxResult

        Dim DialogPreview As New FormSelectPatternFile

        DialogPreview.ShowDialog()
        Dim file = DialogPreview.Path

        If Nothing = file Then

            If (Production.TextBoxFilename.Text <> "") Then
                Return -1
            End If

            Production.TextBoxFilename.Text = ""
            'Production.RichTextBoxNote.Text = ""
            Return -1
        End If

        FlushSpreadsheet()
        ErrorSubSheetStructIni(200, 500)
        gPatternFileName = ""
        gFidFileName = ""
        init_spreadsheet()
        m_Row = 0
        m_EditStateFlag = False

        Try
            'IDS.StopErrorCheck() 'kr?
            AxSpreadsheetProgramming.Caption = file
            m_Execution.m_Pattern.LoadTxtPatternPara(AxSpreadsheetProgramming, file, 0, 0, False) ' True)
            gPatternFileName = m_Execution.m_File.FolderWithNameFromFileName(file)
            gFidFileName = gPatternFileName
            m_Row = 2
            IDS.Data.OpenPathFileData(gPatternFileName + ".pat")
            MyConveyorSettings.InitializeConveyorSetup()
            SystemSetupDataRetrieve() 'SJ add 
            'IDS.StartErrorCheck() 'kr?

            IDS.newOpen = True

            AxSpreadsheetProgramming.Worksheets("Main").Activate()
            'Error checking for all the Spreadsheet
            If m_Execution.m_Pattern.m_ErrorChk.CheckAllError(AxSpreadsheetProgramming, ErrorSubSheet) <> 0 Then

                Rtn = MyMsgBox("Click Ok to flush all the data", MsgBoxStyle.Question + MsgBoxStyle.YesNo, "Error found, flush all the data or not?")
                If MsgBoxResult.Yes = Rtn Then
                    FlushSpreadsheet()
                    ErrorSubSheetStructIni(200, 500)
                    gPatternFileName = ""
                    gFidFileName = ""
                    init_spreadsheet()
                    Production.TextBoxFilename.Text = ""
                    Return -1
                Else

                End If
            Else

            End If

            Production.TextBoxFilename.Text = file

            'set gd settings 'lim
            If gPatternFileName = Nothing Then
            Else
                'lim
                Vision.SetFiducialFilename(gPatternFileName)
            End If
        Catch ex As Exception
            MessageBox.Show("Load pattern data unknown error", "Error information", MessageBoxButtons.OK)
        End Try
        TraceGCCollect()
        Return 0
    End Function



    'menu: File-->Load an encrypt text pattern file
    '   

    Private Sub MenuFileOpen_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuFileOpen.Click
        Dim ErrorMsg As String
        Dim Rtn As MsgBoxResult
        Dim DialogPreview As New FormSelectPatternFile

        '   Xue Wen                         '
        '   check and save old program.     '
        If (gPatternFileName <> "") And (SaveProgram.UnSave = True) Then
            Rtn = MyMsgBox("Save current program before uploading the old one.", MsgBoxStyle.YesNo, "Save Program")
            If (MsgBoxResult.Yes = Rtn) Then
                Local_SaveFile()
                SaveProgram.UnSave = False
            End If
        End If

        LabelMessage("Opening program file......")
        IDS.Data.Execution.Setting.DefaultFileToOpen = Me.ReadDefaultOpenFile()
        DialogPreview.defaultOpenFile = IDS.Data.Execution.Setting.DefaultFileToOpen
        If DialogPreview.ShowDialog() = DialogResult.Cancel Then
            LabelMessage("")
            Return
        End If
        Dim file = DialogPreview.Path

        If Nothing = file Then
            Return
        End If
        If DialogPreview.cbDefaultFile.Checked Then
            IDS.Data.Execution.Setting.DefaultFileToOpen = DialogPreview.Path
            Me.WriteDefaultOpenFile(DialogPreview.Path)
        End If
        'If DialogPreview.removeDefaultFile Then
        '    IDS.Data.Execution.Setting.DefaultFileToOpen = ""
        '    Me.WriteDefaultOpenFile("")
        'End If
        'AxSpreadsheetProgramming.ActiveSheet.UsedRange().Clear()
        m_NewProjectCreated = True
        gPatternFileName = ""
        gFidFileName = ""
        m_Row = 0
        m_EditStateFlag = False

        init_spreadsheet()
        FlushSpreadsheet()
        ErrorSubSheetStructIni(200, 500)
        LaserHeightOffsetZ = -1234
        Try
            Me.Cursor = Cursors.WaitCursor
            'IDS.StopErrorCheck() 'kr?
            AxSpreadsheetProgramming.Caption = file
            Me.GlobalQCEnabled = False
            m_Execution.m_Pattern.LoadTxtPatternPara(AxSpreadsheetProgramming, file, 0, 0, False)
            Me.Cursor = Cursors.WaitCursor
            gPatternFileName = m_Execution.m_File.FolderWithNameFromFileName(file)
            gFidFileName = gPatternFileName
            m_Row = 2
            Dim mode As String = ""
            If m_Execution.m_Pattern.GetTeachingMode(file, mode) Then
                teachingMode = mode
            End If
            EnableTeachingButtons()
            DisableElementsCommandBlockButton(gOffsetCmdIndex)
            IDS.Data.OpenPathFileData(gPatternFileName + ".pat")
            MyConveyorSettings.InitializeConveyorSetup()
            SystemSetupDataRetrieve() 'SJ add 
            'IDS.StartErrorCheck() 'kr?
            IDS.newOpen = True
            LogFile("Opening file " & gPatternFileName + ".pat")
            'Acitvate the "Main" page
            AxSpreadsheetProgramming.Worksheets("Main").Activate()
            LabelMessage("Please wait, Checking content.....")
            'Error checking for all the Spreadsheet
            If 0 <> m_Execution.m_Pattern.m_ErrorChk.CheckAllError(AxSpreadsheetProgramming, ErrorSubSheet) Then
                LabelMessage("File content invalid", True)
                Rtn = MyMsgBox("Click Ok to flush all the data", MsgBoxStyle.Question + MsgBoxStyle.YesNo, "Error found, flush all the data or not?")
                If MsgBoxResult.Yes = Rtn Then
                    FlushSpreadsheet()
                    ErrorSubSheetStructIni(200, 500)
                    gPatternFileName = ""
                    gFidFileName = ""
                    init_spreadsheet()

                    MenuFileExport.Enabled = False
                    MenuFileImport.Enabled = False
                    MenuFileSave.Enabled = False
                    MenuFileSaveAs.Enabled = False
                Else
                    MenuFileExport.Enabled = True
                    MenuFileImport.Enabled = True
                    MenuFileSave.Enabled = True
                    MenuFileSaveAs.Enabled = True
                    UndoData_Logging(0)
                End If
            Else
                LabelMessage("File opened")
                tbOpenedFile.Text = file
                UndoData_Logging(0)

                MenuFileExport.Enabled = True
                MenuFileImport.Enabled = True
                MenuFileSave.Enabled = True
                MenuFileSaveAs.Enabled = True

                MenuEditUndo.Enabled = False
                MenuEditRedo.Enabled = False
                MenuEditCopy.Enabled = False
                MenuEditPaste.Enabled = False
                MenuEditCut.Enabled = False
                MenuEditDelete.Enabled = False
                MenuEditSelectAll.Enabled = True

            End If

            'set gd settings 'lim
            If gPatternFileName = Nothing Then
            Else
                IDS.Data.OpenData()
                'lim
                Vision.SetFiducialFilename(gPatternFileName)
                ''''''''''''''
                '   Xue Wen                                                                     '
                '   Note: Do we need to read out other parameters such as "Pixel size",etc.     '
                ''''''''''''''
                ValueBrightness.Value = IDS.Data.Hardware.Camera.Brightness
                EnableControl()
            End If
            'If IDS.Data.Hardware.Camera.DotQCEnable Then
            '    DisableElementsCommandBlockButton(gGlobalQCCmdIndex)
            'Else
            '    EnableElementsCommandBlockButton(gGlobalQCCmdIndex)
            'End If
            SaveProgram.UnSave = False
            Me.Cursor = Cursors.Default

        Catch ex As Exception
            ExceptionDisplay(ex)
            MessageBox.Show("Loading data unknown error", "Error information", MessageBoxButtons.OK)
            MenuFileExport.Enabled = True
            MenuFileImport.Enabled = True
            MenuFileSave.Enabled = True
            MenuFileSaveAs.Enabled = True
            Me.Cursor = Cursors.Default
        End Try
        LabelMessage("")
        TraceGCCollect()
    End Sub



    'menu: File-->Save encrypt text pattern file

    Private Sub MenuFileSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuFileSave.Click

        Try
            If gPatternFileName <> Nothing Then

                '   When user edit a cell, but system doesn't update the data.                                  '
                '   System only update cell value in "Click" or "Tap" method. So, just add extra one command.   '
                '   Note:   May have better method.                                                             '
                '           Need to test more.                                                
                SelectCell(m_Row + 1, 1)
                IDS.Data.Hardware.Camera.Brightness = ValueBrightness.Value
                IDS.Data.SaveData()

                m_Execution.m_Pattern.SavePatternPara(AxSpreadsheetProgramming, gPatternFileName + ".txt", False)

                SaveProgram.UnSave = False
                MenuEditUndo.Enabled = False
                MenuEditRedo.Enabled = False
                EnableControl()
            Else
                '
                SavePatternFileDialog.InitialDirectory = "C:\IDS\Pattern_Dir"
                SavePatternFileDialog.AddExtension = True
                SavePatternFileDialog.Filter = "Input a new project name|"

                SavePatternFileDialog.ShowDialog()
                Dim file As String
                Dim folder As String = SavePatternFileDialog.FileName()
                If folder <> Nothing Then
                    If folder <> Nothing Then
                        If -1 <> folder.IndexOf(".") Then
                            'Project name cannot contain "."
                            MyMsgBox("Invalid project name!, Press Ok to return", _
                                MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly, "Create new project error")
                            Return
                        ElseIf System.IO.File.Exists(folder) Then
                            'Project name is the same with an existing filename
                            'The file will be deleted
                            System.IO.File.Delete(folder)
                        End If

                        System.IO.Directory.CreateDirectory(folder)
                        file = m_Execution.m_File.NameOnlyFromFullPath(folder)
                        gPatternFileName = folder + "\" + file
                        gFidFileName = gPatternFileName

                        '   When user edit a cell, but system doesn't update the data.                                  '
                        '   System only update cell value in "Click" or "Tap" method. So, just add extra one command.   '
                        '   Note:   May have better method.                                                             '
                        '           Need to test more.                                                                  '
                        SelectCell(m_Row + 1, 1)
                        m_Execution.m_Pattern.SavePatternPara(AxSpreadsheetProgramming, gPatternFileName + ".txt", False)
                        'lim
                        Vision.SetFiducialFilename(gPatternFileName)

                        '''''''''''''''''''''''''''''
                        '   Xue Wen                 '
                        '   Save the Brightness.    '
                        '''''''''''''''''''''''''''''
                        IDS.Data.Hardware.Camera.Brightness = ValueBrightness.Value
                        'IDS.Data.SaveData()
                        IDS.Data.SavePathFileData(gPatternFileName + ".pat")

                        SaveProgram.UnSave = False
                        MenuEditUndo.Enabled = False
                        MenuEditRedo.Enabled = False
                        EnableControl()
                    End If
                End If
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
        TraceGCCollect()
    End Sub

    Private Sub Local_SaveFile()

        Try
            If gPatternFileName <> Nothing Then
                '   When user edit a cell, but system doesn't update the data.                                  '
                '   System only update cell value in "Click" or "Tap" method. So, just add extra one command.   '
                '   Note:   May have better method.                                                             '
                '           Need to test more.                                                                  '
                SelectCell(m_Row + 1, 1)


                m_Execution.m_Pattern.SavePatternPara(AxSpreadsheetProgramming, gPatternFileName + ".txt", False)

                IDS.Data.Hardware.Camera.Brightness = ValueBrightness.Value
                IDS.Data.SaveData()
            Else
                '
                SavePatternFileDialog.InitialDirectory = "C:\IDS\Pattern_Dir"
                SavePatternFileDialog.AddExtension = True
                SavePatternFileDialog.Filter = "Input a new project name|"

                SavePatternFileDialog.ShowDialog()
                Dim file As String
                Dim folder As String = SavePatternFileDialog.FileName()
                If folder <> Nothing Then
                    If folder <> Nothing Then
                        If -1 <> folder.IndexOf(".") Then
                            'Project name cannot contain "."
                            MyMsgBox("Invalid project name!, Press Ok to return", _
                                MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly, "Create new project error")
                            Return
                        ElseIf System.IO.File.Exists(folder) Then
                            'Project name is the same with an existing filename
                            'The file will be deleted
                            System.IO.File.Delete(folder)
                        End If

                        System.IO.Directory.CreateDirectory(folder)
                        file = m_Execution.m_File.NameOnlyFromFullPath(folder)
                        gPatternFileName = folder + "\" + file
                        gFidFileName = gPatternFileName

                        '   When user edit a cell, but system doesn't update the data.                                  '
                        '   System only update cell value in "Click" or "Tap" method. So, just add extra one command.   '
                        '   Note:   May have better method.                                                             '
                        '           Need to test more.                                                                  

                        SelectCell(m_Row + 1, 1)
                        m_Execution.m_Pattern.SavePatternPara(AxSpreadsheetProgramming, gPatternFileName + ".txt", False)
                        Vision.SetFiducialFilename(gPatternFileName)
                        IDS.Data.Hardware.Camera.Brightness = ValueBrightness.Value
                        IDS.Data.SaveData()
                    End If
                End If
            End If
        Catch ex As Exception
            MessageBox.Show("Trying to save when there is no element data.")
            'MessageBox.Show(ex.Message)
        End Try
    End Sub

    'menu: File-->Save as encrypt text pattern file
    '   file:  filename                                                                                

    Private Sub MenuFileSaveAs_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuFileSaveAs.Click

        Try
            '
            SavePatternFileDialog.InitialDirectory = "C:\IDS\Pattern_Dir"
            SavePatternFileDialog.AddExtension = True
            SavePatternFileDialog.Filter = "Input a new project name|"

            SavePatternFileDialog.ShowDialog()
            Dim file As String
            Dim folder As String = SavePatternFileDialog.FileName()
            If folder <> Nothing Then
                If -1 <> folder.IndexOf(".") Then
                    'Project name cannot contain "."
                    MyMsgBox("Invalid project name!, Press Ok to return", _
                        MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly, "Create new project error")
                    Return
                ElseIf System.IO.File.Exists(folder + ".txt") Then
                    'Project name is the same with an existing filename
                    'The file will be deleted
                    System.IO.File.Delete(folder + ".txt")
                    file = m_Execution.m_File.NameOnlyFromFullPath(folder)
                    gPatternFileName = folder
                Else
                    'Create folder for the project
                    System.IO.Directory.CreateDirectory(folder)
                    file = m_Execution.m_File.NameOnlyFromFullPath(folder)
                    Dim gPatternFileName_OldPath As String = gPatternFileName 'lim, to copy fiducial from original folder to new folder
                    gPatternFileName = folder + "\" + file

                    '''''''''''''''''''''''''''''''''''''''''''''''''''''
                    '   Xue Wen                                         '
                    '   Check old file first before doing the copy.     '
                    '''''''''''''''''''''''''''''''''''''''''''''''''''''
                    If System.IO.File.Exists(gPatternFileName_OldPath & ".bmp") Then 'if second fiducial exist
                        System.IO.File.Copy(gPatternFileName_OldPath & ".bmp", gPatternFileName & ".bmp")  'GD
                    End If

                    If System.IO.File.Exists(gPatternFileName_OldPath & "1.mmo") Then 'if second fiducial exist
                        System.IO.File.Copy(gPatternFileName_OldPath & "1.mmo", gPatternFileName & "1.mmo") 'Fiducial File 1
                        System.IO.File.Copy(gPatternFileName_OldPath & "1.bmp", gPatternFileName & "1.bmp") 'Fiducial 1 image
                    End If

                    If System.IO.File.Exists(gPatternFileName_OldPath & "2.mmo") Then 'if second fiducial exist
                        System.IO.File.Copy(gPatternFileName_OldPath & "2.mmo", gPatternFileName & "2.mmo") 'Fiducial File 2
                        System.IO.File.Copy(gPatternFileName_OldPath & "2.bmp", gPatternFileName & "2.bmp") 'Fiducial 2 image
                    End If
                End If

                gFidFileName = gPatternFileName

                m_Execution.m_Pattern.SavePatternPara(AxSpreadsheetProgramming, gPatternFileName + ".txt", False)

                IDS.Data.SavePathFileData(gPatternFileName + ".pat")
                Vision.SetFiducialFilename(gPatternFileName)
                tbOpenedFile.Text = gPatternFileName
                SaveProgram.UnSave = False
                MenuEditUndo.Enabled = False
                MenuEditRedo.Enabled = False
                EnableControl()
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub MenuFileExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuFileExit.Click
        Close()
    End Sub


    'Logging data for Undo operation
    '   UndoLevel: 0 or 1; 
    '               0 is slow, used for multiRow, sub or array
    '               1 is fast, used for single row without sub and array            

    Private Sub UndoData_Logging(ByVal UndoLevel As Integer)

        'Return 'Disable undo operation. this function is terribely wrong...

        m_Execution.m_Undo.UndoLevel = UndoLevel

        If 0 = UndoLevel Then
            Dim FolderName As String = "C:\IDS\Pattern_Dir\SysSwapData"
            If Not System.IO.Directory.Exists(FolderName) Then
                System.IO.Directory.CreateDirectory(FolderName)
            End If

            'Backup the previous undo data
            If m_Execution.m_Undo.hasBackupData Then
                'The Excel filename extension can also be used as a text file format
                System.IO.File.Copy("C:\IDS\Pattern_Dir\SysSwapData\UndoStep_A.Xls", _
                    "C:\IDS\Pattern_Dir\SysSwapData\UndoStep_B.Xls", True)
                m_Execution.m_Undo.CurrentPageName_B = m_Execution.m_Undo.CurrentPageName_A
            End If
            m_Execution.m_Undo.UndoFilename = "C:\IDS\Pattern_Dir\SysSwapData\UndoStep_A.Xls"
            m_Execution.m_Undo.CurrentPageName_A = GetActiveSheetName()
            m_Execution.m_Undo.DataSaveFor_Undo(AxSpreadsheetProgramming)
        Else
            Dim sel As Microsoft.Office.Interop.OWC.Range = AxSpreadsheetProgramming.Selection

            Dim m_rowCount As Integer = sel.Rows.Count()
            Dim m_columnCount As Integer = sel.Columns.Count()
            Dim m_StartRow As Integer = sel.Row
            Dim m_columnStart As Integer = sel.Column

            m_Execution.m_Undo.UndoRow = m_StartRow

            Dim SheetName As String = GetActiveSheetName()
            If 1 = m_rowCount Then
                If m_Execution.m_Undo.hasBackupData Then
                    m_Execution.m_Undo.UndoPatternRec_B = m_Execution.m_Undo.UndoPatternRec_A
                End If

                m_Execution.m_Pattern.m_ErrorChk.ConvertToDataStruct(m_Execution.m_Undo.UndoPatternRec_A, _
                    AxSpreadsheetProgramming, SheetName, m_StartRow)
            End If
        End If


        MenuEditUndo.Enabled = True
        MenuEditRedo.Enabled = False
        TraceGCCollect()
    End Sub
    'yy
    'To change the way that undo operation before it got peformace issue like function above
    Dim undo As Object(,)
    Private Sub TUndoData_Logging(ByVal UndoLevel As Integer)
        m_Execution.m_Undo.UndoLevel = UndoLevel

        If 0 = UndoLevel Then
            Dim FolderName As String = "C:\IDS\Pattern_Dir\SysSwapData"
            If Not System.IO.Directory.Exists(FolderName) Then
                System.IO.Directory.CreateDirectory(FolderName)
            End If

            'Backup the previous undo data
            If m_Execution.m_Undo.hasBackupData Then
                'The Excel filename extension can also be used as a text file format
                System.IO.File.Copy("C:\IDS\Pattern_Dir\SysSwapData\UndoStep_A.Xls", _
                    "C:\IDS\Pattern_Dir\SysSwapData\UndoStep_B.Xls", True)
                m_Execution.m_Undo.CurrentPageName_B = m_Execution.m_Undo.CurrentPageName_A
            End If
            m_Execution.m_Undo.UndoFilename = "C:\IDS\Pattern_Dir\SysSwapData\UndoStep_A.Xls"
            m_Execution.m_Undo.CurrentPageName_A = GetActiveSheetName()
            m_Execution.m_Undo.DataSaveFor_Undo(AxSpreadsheetProgramming)

        Else
            Dim sel As Microsoft.Office.Interop.OWC.Range = AxSpreadsheetProgramming.Selection

            Dim m_rowCount As Integer = sel.Rows.Count()
            Dim m_columnCount As Integer = sel.Columns.Count()
            Dim m_StartRow As Integer = sel.Row
            Dim m_columnStart As Integer = sel.Column

            m_Execution.m_Undo.UndoRow = m_StartRow

            Dim SheetName As String = GetActiveSheetName()
            If 1 = m_rowCount Then
                If m_Execution.m_Undo.hasBackupData Then
                    m_Execution.m_Undo.UndoPatternRec_B = m_Execution.m_Undo.UndoPatternRec_A
                End If

                m_Execution.m_Pattern.m_ErrorChk.ConvertToDataStruct(m_Execution.m_Undo.UndoPatternRec_A, _
                    AxSpreadsheetProgramming, SheetName, m_StartRow)
            End If
        End If


        MenuEditUndo.Enabled = True
        MenuEditRedo.Enabled = False
        TraceGCCollect()
    End Sub


    'menu: Edit-->Redo                                                                           

    Private Sub MenuEditRedo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuEditRedo.Click

        If 0 = m_Execution.m_Undo.UndoLevel Then
            'm_Execution.m_Undo.UndoFilename = "C:\IDS\Pattern_Dir\SysSwapData\UndoStep_A.Xls"
            'm_Execution.m_Undo.CurrentPageName_A = GetActiveSheetName()
            'm_Execution.m_Undo.DataSaveFor_Undo(AxSpreadsheetProgramming)

            'm_Execution.m_Undo.UndoFilename = "C:\IDS\Pattern_Dir\SysSwapData\UndoStep_B.Xls"
            'FlushSpreadsheet()
            'm_Execution.m_Undo.DataLoadFor_Undo(AxSpreadsheetProgramming)
            m_Execution.m_Undo.UndoFilename = "C:\IDS\Pattern_Dir\SysSwapData\UndoStep_B.Xls"
            FlushSpreadsheet()
            m_Execution.m_Undo.DataLoadFor_Undo(AxSpreadsheetProgramming)

            System.IO.File.Copy("C:\IDS\Pattern_Dir\SysSwapData\UndoStep_B.Xls", _
                   "C:\IDS\Pattern_Dir\SysSwapData\tempHolder.Xls", True)

            System.IO.File.Copy("C:\IDS\Pattern_Dir\SysSwapData\UndoStep_A.Xls", _
                   "C:\IDS\Pattern_Dir\SysSwapData\UndoStep_B.Xls", True)

            System.IO.File.Copy("C:\IDS\Pattern_Dir\SysSwapData\tempHolder.Xls", _
                   "C:\IDS\Pattern_Dir\SysSwapData\UndoStep_A.Xls", True)

            AxSpreadsheetProgramming.Worksheets(m_Execution.m_Undo.CurrentPageName_B).Activate()
        Else
            Dim sel As Microsoft.Office.Interop.OWC.Range = AxSpreadsheetProgramming.Selection

            Dim m_rowCount As Integer = sel.Rows.Count()
            Dim m_columnCount As Integer = sel.Columns.Count()
            Dim m_StartRow As Integer = sel.Row
            Dim m_columnStart As Integer = sel.Column

            Dim SheetName As String = GetActiveSheetName()
            If 1 = m_rowCount Then
                m_Execution.m_Pattern.m_ErrorChk.ConvertToDataStruct(m_Execution.m_Undo.UndoPatternRec_A, _
                    AxSpreadsheetProgramming, SheetName, m_Execution.m_Undo.UndoRow)

                m_Execution.m_Pattern.m_ErrorChk.ConvertFromDataStruct(m_Execution.m_Undo.UndoPatternRec_B, _
                    AxSpreadsheetProgramming, m_Execution.m_Undo.UndoRow)
            End If
        End If

        MenuEditUndo.Enabled = True
        MenuEditRedo.Enabled = False
        TraceGCCollect()
    End Sub

    'menu: Edit-->Undo                                                   
    Private Sub MenuEditUndo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuEditUndo.Click

        If 0 = m_Execution.m_Undo.UndoLevel Then
            'm_Execution.m_Undo.UndoFilename = "C:\IDS\Pattern_Dir\SysSwapData\UndoStep_B.Xls"
            'm_Execution.m_Undo.CurrentPageName_B = GetActiveSheetName()
            'm_Execution.m_Undo.DataSaveFor_Undo(AxSpreadsheetProgramming)

            'm_Execution.m_Undo.UndoFilename = "C:\IDS\Pattern_Dir\SysSwapData\UndoStep_A.Xls"
            'FlushSpreadsheet()
            'm_Execution.m_Undo.DataLoadFor_Undo(AxSpreadsheetProgramming)

            m_Execution.m_Undo.UndoFilename = "C:\IDS\Pattern_Dir\SysSwapData\UndoStep_B.Xls"
            FlushSpreadsheet()
            m_Execution.m_Undo.DataLoadFor_Undo(AxSpreadsheetProgramming)

            System.IO.File.Copy("C:\IDS\Pattern_Dir\SysSwapData\UndoStep_B.Xls", _
                   "C:\IDS\Pattern_Dir\SysSwapData\tempHolder.Xls", True)

            System.IO.File.Copy("C:\IDS\Pattern_Dir\SysSwapData\UndoStep_A.Xls", _
                   "C:\IDS\Pattern_Dir\SysSwapData\UndoStep_B.Xls", True)

            System.IO.File.Copy("C:\IDS\Pattern_Dir\SysSwapData\tempHolder.Xls", _
                   "C:\IDS\Pattern_Dir\SysSwapData\UndoStep_A.Xls", True)
            Try

                AxSpreadsheetProgramming.Worksheets(m_Execution.m_Undo.CurrentPageName_A).Activate()
            Catch ex As Exception
                'AxSpreadsheetProgramming.Worksheets(m_Execution.m_Undo.CurrentPageName_A).Activate()
            End Try
        Else
            Dim sel As Microsoft.Office.Interop.OWC.Range = AxSpreadsheetProgramming.Selection

            Dim m_rowCount As Integer = sel.Rows.Count()
            Dim m_columnCount As Integer = sel.Columns.Count()
            Dim m_StartRow As Integer = sel.Row
            Dim m_columnStart As Integer = sel.Column

            Dim SheetName As String = GetActiveSheetName()
            If 1 = m_rowCount Then
                m_Execution.m_Pattern.m_ErrorChk.ConvertToDataStruct(m_Execution.m_Undo.UndoPatternRec_B, _
                    AxSpreadsheetProgramming, SheetName, m_Execution.m_Undo.UndoRow)

                m_Execution.m_Pattern.m_ErrorChk.ConvertFromDataStruct(m_Execution.m_Undo.UndoPatternRec_A, _
                    AxSpreadsheetProgramming, m_Execution.m_Undo.UndoRow)
            End If
        End If

        MenuEditUndo.Enabled = False
        MenuEditRedo.Enabled = True
        TraceGCCollect()
    End Sub

    'menu: Edit-->Cut: Cut the selected range with Array/Sub attached    
    Private Sub MenuEditCut_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuEditCut.Click
        Dim sel As Microsoft.Office.Interop.OWC.Range = AxSpreadsheetProgramming.Selection
        Dim m_rowCount As Integer = sel.Rows.Count()
        Dim m_columnCount As Integer = sel.Columns.Count()
        Dim m_StartRow As Integer = sel.Row
        Dim m_columnStart As Integer = sel.Column

        Dim SheetName As String = GetActiveSheetName()
        If m_columnCount = gMaxColumns Then
            CopiedSheetName = GetActiveSheetName()
            UndoData_Logging(0)
            'Copy selected data and also with the related Array/Sub data
            m_Execution.m_Pattern.SavePatternParaAllSheets(AxSpreadsheetProgramming, _
                "C:\IDS\Pattern_Dir\SysSwapData\CopyPaste.txt", 1, False)

            Cancel()
            MenuEditPaste.Enabled = True
        End If
        TraceGCCollect()
    End Sub


    'menu: Edit-->Copy: Copy the selected range with Array/Sub attached                                                                            

    Private Sub MenuEditCopy_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuEditCopy.Click
        Dim sel As Microsoft.Office.Interop.OWC.Range = AxSpreadsheetProgramming.Selection

        Dim m_rowCount As Integer = sel.Rows.Count()
        Dim m_columnCount As Integer = sel.Columns.Count()
        Dim m_StartRow As Integer = sel.Row
        Dim m_columnStart As Integer = sel.Column

        If m_columnCount = gMaxColumns Then
            CopiedSheetName = GetActiveSheetName()
            'Copy selected data and also with the related Array/Sub data
            m_Execution.m_Pattern.SavePatternParaAllSheets(AxSpreadsheetProgramming, _
                "C:\IDS\Pattern_Dir\SysSwapData\CopyPaste.txt", 1, False)
            MenuEditPaste.Enabled = True
        End If
        TraceGCCollect()
    End Sub


    'menu: Edit-->Paste: Paste the Copied range with Array/Sub attached                                                                            

    Private Sub MenuEditPaste_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuEditPaste.Click
        Dim sel As Microsoft.Office.Interop.OWC.Range = AxSpreadsheetProgramming.Selection

        Dim m_rowCount As Integer = sel.Rows.Count()
        Dim m_StartRow As Integer = sel.Row

        If CopiedSheetName <> GetActiveSheetName() Then
            MyMsgBox("Cannot be pasted in this Spreadsheet!", MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly, "Copy/Paste in Different Sheet")
            Return
        End If

        If "" = GetActiveSheetName() Then
            MyMsgBox("Copy contents cannot be used!", MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly, "Contents are invalid")
            Return
        End If

        If 1 = m_rowCount Then
            Spreadsheet_CheckForWithinLinkRange(False)
            If m_InLinkRange Then
                MyMsgBox("Paste is not allowed inside LINK range", MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly, "Paste is not alowed")
            Else
                m_Execution.m_Pattern.LoadTxtPatternParaAllSheets(AxSpreadsheetProgramming, _
                    "C:\IDS\Pattern_Dir\SysSwapData\CopyPaste.txt", _
                    CopiedSheetName, 2, m_StartRow, False)
                UndoData_Logging(0)
            End If
        End If
        TraceGCCollect()
    End Sub

    'menu: Edit-->Delete: Delecte all select                                                                        

    Private Sub MenuEditDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuEditDelete.Click
        UndoData_Logging(0)
        Cancel()
        TraceGCCollect()
    End Sub

    'menu: Edit-->SelectAll: Select all in the current spreadsheet  

    Private Sub MenuEditSelectAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuEditSelectAll.Click
        AxSpreadsheetProgramming.ActiveSheet.UsedRange.Select()
    End Sub

    Public Sub Disp_Dispenser_Unit_info()

        Dim DispensingType As String = IDS.Data.Hardware.Dispenser.Left.HeadType

    End Sub

    Private Sub ButtonToggleMode_click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonToggleMode.Click
        TimerForUpdate.Start() 'kr nov28
        If CurrentMode = "Program Editor" Then
            btExit.Visible = False
            btExit.Enabled = False
            OffLaser()
            'stepper panel
            'MySettings.PanelLeft.Controls.Add(m_Tri.SteppingButtons)
            MySettings.gbRelMove.Controls.Add(m_Tri.SteppingButtons)
            Controls.Remove(m_Tri.SteppingButtons)
            m_Tri.SteppingButtons.Location = New Point(17, 18) 'New Point(384, 12)
            m_Tri.SteppingButtons.BringToFront()
            m_Tri.SteppingButtons.Show()
            MySettings.AddDefaultView()
            CurrentMode = "Basic Setup"
            ButtonToggleMode.Text = "Go to Program Editor"
            Me.lbPostName.Text = "Robot"
            CBExpandSpreadsheet.Visible = False
            CBExpandSpreadsheet.Visible = False
            MySettings.PanelRight.Location = New Point(768, 33)
            MySettings.PanelLeft.Location = New Point(0, 33)
            Me.PanelVision.Location = New Point(0, 341)
            Me.PanelVisionCtrl.Location = New Point(0, 916)
            Me.PanelVisionCtrl.BringToFront()
            Me.PanelVision.BringToFront()
            Me.Controls.Add(MySettings.PanelRight)
            Me.Controls.Add(MySettings.PanelLeft)
            MySettings.PanelRight.BringToFront()
            MySettings.PanelLeft.BringToFront()
            MySettings.RichTextBox1.BringToFront()
            MySettings.RevertData()
            EnableDisableMenuBar()
            'Conveyor.Command("Reset Conveyor Status")
            Conveyor.Command("Manual Mode")
            If m_Tri.ZPosition < -5 Then
                m_Tri.Set_XY_Speed(IDS.Data.Hardware.Gantry.ServiceXYSpeed)
                m_Tri.Set_Z_Speed(IDS.Data.Hardware.Gantry.ServiceZSpeed)
                m_Tri.Move_Z(0)
            End If
        ElseIf CurrentMode = "Basic Setup" Then
            btExit.Enabled = True
            btExit.Visible = True
            MySettings.RemoveCurrentPanel()

            'in case you didn't exit conveyor settings but straightaway pressed program editor
            'Conveyor.PositionTimer.Stop()

            'stepper panel
            Controls.Add(m_Tri.SteppingButtons)
            MySettings.gbRelMove.Controls.Remove(m_Tri.SteppingButtons)
            m_Tri.SteppingButtons.Location = PanelToBeAdded.Location()
            m_Tri.SteppingButtons.BringToFront()
            m_Tri.SteppingButtons.Show()

            CurrentMode = "Program Editor"
            ButtonToggleMode.Text = "Go to Basic Setup"
            Me.lbPostName.Text = "Working Space"
            Disp_Dispenser_Unit_info()
            CBExpandSpreadsheet.Visible = True
            Me.Controls.Remove(MySettings.PanelRight)
            Me.Controls.Remove(MySettings.PanelLeft)


            Me.PanelVision.Location = New Point(84, 360)
            Me.PanelVisionCtrl.Location = New Point(84, 940)
            IDS.Data.OpenData()
            'SJ
            SystemSetupDataRetrieve()
            EnableDisableMenuBar()

        End If
    End Sub

    Private Sub EnableDisableMenuBar()
        If CurrentMode = "Basic Setup" Then
            MenuFileNew.Enabled = False
            MenuFileOpen.Enabled = False
            MenuFileImport.Enabled = False
            MenuFileExport.Enabled = False
            MenuFileSave.Enabled = False
            MenuFileSaveAs.Enabled = False

            MenuEditUndo.Enabled = False
            MenuEditRedo.Enabled = False
            MenuEditCut.Enabled = False
            MenuEditCopy.Enabled = False
            MenuEditPaste.Enabled = False
            MenuEditDelete.Enabled = False
            MenuEditSelectAll.Enabled = False

            MenuItem67.Enabled = True
            MenuItem88.Enabled = False

            MenuItem33.Enabled = False
            MenuItem32.Enabled = False
            MenuItem29.Enabled = False
            MenuItem30.Enabled = True
            MenuItem62.Enabled = False
            OptimizePath.Enabled = False

        ElseIf CurrentMode = "Program Editor" Then
            MenuFileNew.Enabled = True
            MenuFileOpen.Enabled = True
            MenuFileImport.Enabled = True
            MenuFileExport.Enabled = True
            MenuFileSave.Enabled = True
            MenuFileSaveAs.Enabled = True

            MenuEditUndo.Enabled = True
            MenuEditRedo.Enabled = True
            MenuEditCut.Enabled = True
            MenuEditCopy.Enabled = True
            MenuEditPaste.Enabled = True
            MenuEditDelete.Enabled = True
            MenuEditSelectAll.Enabled = True

            MenuItem67.Enabled = False
            MenuItem88.Enabled = True

            MenuItem33.Enabled = True
            MenuItem32.Enabled = True
            MenuItem29.Enabled = True
            MenuItem30.Enabled = False
            MenuItem62.Enabled = True
            OptimizePath.Enabled = True
        End If
    End Sub

    Private Sub CheckBox3_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CBExpandSpreadsheet.CheckedChanged
        If CBExpandSpreadsheet.Checked = True Then
            AxSpreadsheetProgramming.Height = 906
            AxSpreadsheetProgramming.BringToFront()
        Else
            AxSpreadsheetProgramming.Height = 260
            'AxSpreadsheetProgramming.SendToBack()
        End If
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub



    'Update link for next/previous row

    Private Sub UpdateLinkForNextRow()
        Spreadsheet_CheckForWithinLinkRange(True)
        If m_InLinkRange Then
            If GetCellValue(m_Row, gCommandNameColumn) = "Line" Then
                SetCellValue(m_Row + 1, gPos1XColumn, GetCellValue(m_Row, gPos2XColumn))
                SetCellValue(m_Row + 1, gPos1YColumn, GetCellValue(m_Row, gPos2YColumn))
            ElseIf GetCellValue(m_Row, gCommandNameColumn) = "Arc" Then
                SetCellValue(m_Row + 1, gPos1XColumn, GetCellValue(m_Row, gPos3XColumn))
                SetCellValue(m_Row + 1, gPos1YColumn, GetCellValue(m_Row, gPos3YColumn))
            End If
        End If
    End Sub

    Private Sub UpdateLinkForPreviousRow()
        Spreadsheet_CheckForWithinLinkRange(False)
        If m_InLinkRange Then
            If GetCellValue(m_Row - 1, gCommandNameColumn) = "Line" Then
                SetCellValue(m_Row - 1, gPos2XColumn, GetCellValue(m_Row, gPos1XColumn))
                SetCellValue(m_Row - 1, gPos2YColumn, GetCellValue(m_Row, gPos1YColumn))
            ElseIf GetCellValue(m_Row - 1, gCommandNameColumn) = "Arc" Then
                SetCellValue(m_Row - 1, gPos3XColumn, GetCellValue(m_Row, gPos1XColumn))
                SetCellValue(m_Row - 1, gPos3YColumn, GetCellValue(m_Row, gPos1YColumn))
            End If
        End If
    End Sub




    'EditingToolbar_ButtonClick: edit xyz coordinates using toolbar
    '                                                                                  

    Private Sub EditingToolbar_ButtonClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ToolBarButtonClickEventArgs)
        'If Me.AxSpreadsheetProgramming.Enabled = False Then
        '    Return
        'End If
        'Dim idFlag As Integer = 0
        'If e.Button Is EditingToolBar1.Buttons(0) Then
        '    Dim sel As Microsoft.Office.Interop.OWC.Range = AxSpreadsheetProgramming.Selection
        '    Dim m_rowCount As Integer = sel.Rows.Count()
        '    Dim m_StartRow As Integer = sel.Row
        '    Dim commandName As String = GetCellValue(m_StartRow, gCommandNameColumn)
        '    If m_rowCount = 1 And commandName = Nothing Then
        '        If Not (m_EditStateFlag) And Not (m_ProgrammingStateFlag) Then
        '            LabelMessage("")
        '        End If
        '        Return
        '    End If
        '    idFlag = 1
        '    m_EditStateFlag = False
        '    m_SteppingPostFlag = True
        'End If
        'Dim CmdName As String
        'm_Row = GetActiveCellRow()
        'CmdName = GetCellValue(m_Row, gCommandNameColumn)
        'If CmdName = "GlobalQC" Then
        '    m_EditStateFlag = True
        '    idFlag = 1 'edit
        'End If
        'EditingToolbar_Implementation(idFlag)
    End Sub



    'EditingToolbar_ButtonClick: edit xyz coordinates using toolbar
    '    idFlag: 1=goto next point, 0=edit current point
    '

    Private Sub EditingToolbar_Implementation(ByVal idFlag As Integer)
        MenuEditCopy.Enabled = False
        MenuEditPaste.Enabled = False
        MenuEditCut.Enabled = False
        MenuEditUndo.Enabled = False
        MenuEditRedo.Enabled = False


        Dim i As Integer
        Dim posX, posNoMax As Integer
        Dim pos(3) As Double
        Dim cell1, cell2 As Object
        Dim CmdName As String
        m_Row = GetActiveCellRow()
        CmdName = GetCellValue(m_Row, gCommandNameColumn)

        Dim sel As Microsoft.Office.Interop.OWC.Range = AxSpreadsheetProgramming.Selection
        Dim m_rowCount As Integer = sel.Rows.Count()
        Dim m_columnCount As Integer = sel.Columns.Count()
        Dim m_StartRow As Integer = sel.Row
        Dim m_StartColumn As Integer = sel.Column

        If 1 = idFlag Then             'Goto next point, next point button
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '   Xue Wen                                             '
            '   It sould be flaged afer pressing "Edit" button.     '
            '   Note: Test more.                                    '
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            'm_EditStateFlag = True
            If (CmdName = "GlobalQC") Then
                DisableCoordinateUpdateInSpreadsheet()
                Cancel()
            Else
                DisableCoordinateUpdateInSpreadsheet()
                DisableEditingToolbarSwitchButton() '
                DisableEditingToolbarEditButton() '
                LabelMessage("Edit or goto next Pt")
                If m_Execution.m_Pattern.m_ErrorChk.CheckValidPtXY(CmdName) Then

                    If m_Execution.m_Pattern.m_ErrorChk.CheckRequiredPt3XY(CmdName) Then
                        posNoMax = 3
                    ElseIf m_Execution.m_Pattern.m_ErrorChk.CheckRequiredPt2XY(CmdName) Then
                        posNoMax = 2
                    Else
                        posNoMax = 1
                        DisableEditingToolbarEditButton()    'Goto next point. Position jump is not allowed
                    End If

                    'EXPERIMENTAL 11/23/15 
                    'start position
                    If m_columnCount = gMaxColumns Or m_columnCount = 1 Then
                        m_TeachStepNumber = 1
                        posX = gPos1XColumn

                        '(from 1st position go back to 2nd position)
                    ElseIf gPos1XColumn <= m_StartColumn And gPos1ZColumn >= m_StartColumn Then
                        If posNoMax >= 2 Then
                            m_TeachStepNumber = 2
                            posX = gPos2XColumn
                        Else
                            m_TeachStepNumber = 1
                            posX = gPos1XColumn
                        End If

                        cell1 = GetCell(m_Row, gPos1XColumn)
                        cell2 = GetCell(m_Row, gPos1XColumn + 2)
                        m_Execution.m_Pattern.Spreadsheet_CellGrey(cell1, cell2, True, AxSpreadsheetProgramming)

                        'Previous confirmed data is the begaining of the row.  It will be used to update the end of previous row
                        UpdateLinkForPreviousRow()

                        '(from 2nd position go back to 3rd position)
                    ElseIf gPos2XColumn <= m_StartColumn And gPos2ZColumn >= m_StartColumn Then
                        If posNoMax >= 3 Then
                            m_TeachStepNumber = 3
                            posX = gPos3XColumn
                        Else
                            m_TeachStepNumber = 1
                            posX = gPos1XColumn
                        End If

                        cell1 = GetCell(m_Row, gPos2XColumn)
                        cell2 = GetCell(m_Row, gPos2XColumn + 2)
                        m_Execution.m_Pattern.Spreadsheet_CellGrey(cell1, cell2, True, AxSpreadsheetProgramming)

                        'Previous confirmed data is the end of the row.  It will be used to update the begaining of next row
                        UpdateLinkForNextRow()

                        '(from 3rd position go back to 1st position)
                    ElseIf gPos3XColumn <= m_StartColumn And gPos3ZColumn >= m_StartColumn Then
                        m_TeachStepNumber = 1
                        posX = gPos1XColumn

                        cell1 = GetCell(m_Row, gPos3XColumn)
                        cell2 = GetCell(m_Row, gPos3XColumn + 2)
                        m_Execution.m_Pattern.Spreadsheet_CellGrey(cell1, cell2, True, AxSpreadsheetProgramming)

                        'Previous confirmed data is the end of the row.  It will be used to update the begaining of next row
                        UpdateLinkForNextRow()
                    End If


                    cell1 = GetCell(m_Row, posX)
                    cell2 = GetCell(m_Row, posX + 2)
                    SelectRange(cell1, cell2)
                    pos(0) = GetCellValue(m_Row, posX)
                    pos(1) = GetCellValue(m_Row, posX + 1)
                    pos(2) = GetCellValue(m_Row, posX + 2)

                    pos(0) = pos(0)
                    pos(1) = pos(1)
                    pos(2) = pos(2)
                    Dim needleGap As Double = GetCellValue(m_StartRow, gNeedleGapColumn)

                    If CmdName = "ChipEdge" Or CmdName = "QC" Or CmdName = "Height" Or CmdName = "Fiducial" Or CmdName = "Reject" Then
                        MoveToSpreadsheetPoint(pos, "Vision")
                    Else
                        'MoveToSpreadsheetPoint(pos, "Needle")
                        MoveToSpreadsheetPoint(pos, "Needle", needleGap, m_StartRow)
                    End If

                End If
                EnableEditingToolbarSwitchButton() '
                DisableTeachingToolbarOKButton()
                EnableTeachingToolbarCancelButton()
                EnableEditingToolbarEditButton() 'Edit start
                DisableTeachingButtons()
                UndoData_Logging(1)
            End If

        ElseIf 0 = idFlag Then         'Edit current point, pen button
            ToolBarSwitch("Edit")
            LabelMessage("Edit or confirm Pt")

            If m_TeachStepNumber = 1 Then
                posX = gPos1XColumn
            ElseIf m_TeachStepNumber = 2 Then
                posX = gPos2XColumn
            ElseIf m_TeachStepNumber = 3 Then
                posX = gPos3XColumn
            End If

            cell1 = GetCell(m_Row, posX)
            cell2 = GetCell(m_Row, posX + 2)
            m_Execution.m_Pattern.Spreadsheet_CellGrey(cell1, cell2, False, AxSpreadsheetProgramming)
            m_EditStateFlag = True
            DisableTeachingButtons()
            DisableEditingToolbarEditButton() 'Edit start
            EnableTeachingToolbarCancelButton()

            '''''''''''''''''''''''''''''''''''''''''''''
            '   Xue Wen                                 '
            '   Save as tempData for vision system.     '
            '''''''''''''''''''''''''''''''''''''''''''''
            If m_TeachStepNumber = 1 Then
                tempPosX = GetCellValue(m_Row, gPos1XColumn)
                tempPosY = GetCellValue(m_Row, gPos1YColumn)
                tempPosZ = GetCellValue(m_Row, gPos1ZColumn)
            ElseIf m_TeachStepNumber = 2 Then
                tempPosX = GetCellValue(m_Row, gPos2XColumn)
                tempPosY = GetCellValue(m_Row, gPos2YColumn)
                tempPosZ = GetCellValue(m_Row, gPos2ZColumn)
            Else
                tempPosX = GetCellValue(m_Row, gPos1XColumn)
                tempPosY = GetCellValue(m_Row, gPos1YColumn)
                tempPosZ = GetCellValue(m_Row, gPos1ZColumn)
            End If


        End If
        TraceGCCollect()
    End Sub

    Private Sub InitialPatternParameterSetup()

        Dim gpID As String = IDS.Data.Admin.User.Group.ID
        'SJ test only
        Dim recID As String
        m_Execution.m_Pattern.LoadDefaultPara(gpID, recID, 0)
        m_TeachStepNumber = 1
        m_CamPos(0) = 0.0
        m_CamPos(1) = 0.0
        m_CamPos(2) = 0.0
        TraceGCCollect()
    End Sub

    Public Sub AddGlobalQCToTop()
        SelectCell(1, 1)
        AxSpreadsheetProgramming.Selection.EntireRow.Insert()
        SetCellValue(1, gCommandNameColumn, "GlobalQC")
    End Sub

    Public Sub AddCommandToSpreadsheet(ByVal type As String)
        Dim cell1, cell2 As Object
        Dim SkippedFirstPt As Boolean = False
        Dim NotArray As Boolean = type <> "Array"
        Dim Cmd As CIDSData.CIDSPatternExecution.CIDSTemplate
        Cmd = m_Execution.m_Pattern.AddCommand(type, m_CamPos)

        SelectCell(m_Row, 1)
        AxSpreadsheetProgramming.Selection.EntireRow.Insert()
        ClearRow(m_Row)

        SetCellValue(m_Row, gCommandNameColumn, type)
        If type.ToUpper = "LINKSTART" And type.ToUpper = "LINKEND" Then
            Exit Sub
        Else
            m_Execution.m_ItemLUT.Cmd2Index(type)
            If m_Execution.m_ItemLUT.GetFlag(gPos1XColumn) Then
                If NotArray Then
                    If m_InLinkRange Then
                        If GetCellValue(m_Row - 1, gCommandNameColumn) = "Line" Then
                            SetCellValue(m_Row, gPos1XColumn, GetCellValue(m_Row - 1, gPos2XColumn))
                            SetCellValue(m_Row, gPos1YColumn, GetCellValue(m_Row - 1, gPos2YColumn))
                            SetCellValue(m_Row, gPos1ZColumn, GetCellValue(m_Row - 1, gPos2ZColumn))
                            m_TeachStepNumber = 2
                            cell1 = GetCell(m_Row, gPos1XColumn)
                            cell2 = GetCell(m_Row, gPos1ZColumn)
                            m_Execution.m_Pattern.Spreadsheet_CellGrey(cell1, cell2, True, AxSpreadsheetProgramming)
                            SelectCell(m_Row, gPos2XColumn)
                            SkippedFirstPt = True
                        ElseIf GetCellValue(m_Row - 1, gCommandNameColumn) = "Arc" Then
                            SetCellValue(m_Row, gPos1XColumn, GetCellValue(m_Row - 1, gPos3XColumn))
                            SetCellValue(m_Row, gPos1YColumn, GetCellValue(m_Row - 1, gPos3YColumn))
                            SetCellValue(m_Row, gPos1ZColumn, GetCellValue(m_Row - 1, gPos3ZColumn))
                            m_TeachStepNumber = 2
                            cell1 = GetCell(m_Row, gPos1XColumn)
                            cell2 = GetCell(m_Row, gPos1ZColumn)
                            m_Execution.m_Pattern.Spreadsheet_CellGrey(cell1, cell2, True, AxSpreadsheetProgramming)
                            SelectCell(m_Row, gPos2XColumn)
                            SkippedFirstPt = True
                        Else
                            SetCellValue(m_Row, gPos1XColumn, Cmd.Pos1.X + 1000)
                            SetCellValue(m_Row, gPos1YColumn, Cmd.Pos1.Y + 1000)
                            SetCellValue(m_Row, gPos1ZColumn, Cmd.Pos1.Z + 1000)
                            cell1 = GetCell(m_Row, gPos1XColumn)
                            cell2 = GetCell(m_Row, gPos1ZColumn)
                            SelectRange(cell1, cell2)
                        End If
                    Else
                        SetCellValue(m_Row, gPos1XColumn, Cmd.Pos1.X)
                        SetCellValue(m_Row, gPos1YColumn, Cmd.Pos1.Y)
                        SetCellValue(m_Row, gPos1ZColumn, Cmd.Pos1.Z)
                    End If
                End If
            End If
            If m_Execution.m_ItemLUT.GetFlag(gPos2XColumn) Then
                If NotArray Then
                    SetCellValue(m_Row, gPos2XColumn, Cmd.Pos2.X)
                    SetCellValue(m_Row, gPos2YColumn, Cmd.Pos2.Y)
                    SetCellValue(m_Row, gPos2ZColumn, Cmd.Pos2.Z)
                    If SkippedFirstPt = True Then
                        cell1 = GetCell(m_Row, gPos2XColumn)
                        cell2 = GetCell(m_Row, gPos2ZColumn)
                        SelectRange(cell1, cell2)
                    End If
                End If
            End If
            If m_Execution.m_ItemLUT.GetFlag(gPos3XColumn) Then
                If NotArray Then
                    SetCellValue(m_Row, gPos3XColumn, Cmd.pos3.X)
                    SetCellValue(m_Row, gPos3YColumn, Cmd.pos3.Y)
                    SetCellValue(m_Row, gPos3ZColumn, Cmd.pos3.Z)
                End If
            End If
            If m_Execution.m_ItemLUT.GetFlag(gNeedleColumn) Then
                SetCellValue(m_Row, gNeedleColumn, "Left")   'Cmd.Needle
            End If
            If m_Execution.m_ItemLUT.GetFlag(gDispensColumn) Then
                'gSubnameColumn use the same slot
                If Cmd.DispenseFlag = "True" Or Cmd.DispenseFlag = "On" Then
                    Cmd.DispenseFlag = "On"
                Else
                    Cmd.DispenseFlag = "Off"
                End If
                SetCellValue(m_Row, gDispensColumn, Cmd.DispenseFlag)
            End If

            If m_Execution.m_ItemLUT.GetFlag(gTravelSpeedColumn) Then
                SetCellValue(m_Row, gTravelSpeedColumn, Cmd.TravelSpeed)
            End If
            If m_Execution.m_ItemLUT.GetFlag(gNeedleGapColumn) Then
                SetCellValue(m_Row, gNeedleGapColumn, Cmd.NeedleGap)
            End If
            If m_Execution.m_ItemLUT.GetFlag(gDurationColumn) Then
                SetCellValue(m_Row, gDurationColumn, Cmd.DispenseDuration)
            End If
            If m_Execution.m_ItemLUT.GetFlag(gTravelDelayColumn) Then
                SetCellValue(m_Row, gTravelDelayColumn, Cmd.TravelDelay)
            End If
            If m_Execution.m_ItemLUT.GetFlag(gRetractDelayColumn) Then
                SetCellValue(m_Row, gRetractDelayColumn, Cmd.RetractDelay)
            End If
            If m_Execution.m_ItemLUT.GetFlag(gApproachHtColumn) Then
                SetCellValue(m_Row, gApproachHtColumn, Cmd.ApproachDispHeight)
            End If
            If m_Execution.m_ItemLUT.GetFlag(gRetractSpeedColumn) Then
                SetCellValue(m_Row, gRetractSpeedColumn, Cmd.RetractSpeed)
            End If
            If m_Execution.m_ItemLUT.GetFlag(gRetractHtColumn) Then
                SetCellValue(m_Row, gRetractHtColumn, Cmd.RetractHeight)
            End If
            If m_Execution.m_ItemLUT.GetFlag(gClearanceHtColumn) Then
                SetCellValue(m_Row, gClearanceHtColumn, Cmd.ClearanceHeight)
            End If
            If m_Execution.m_ItemLUT.GetFlag(gDTaiDistColumn) Then
                SetCellValue(m_Row, gDTaiDistColumn, Cmd.DetailingDistance)
            End If
            If m_Execution.m_ItemLUT.GetFlag(gArcRadiusColumn) Then
                SetCellValue(m_Row, gArcRadiusColumn, Cmd.ArcRadius)
            End If
            If m_Execution.m_ItemLUT.GetFlag(gPitchColumn) Then
                SetCellValue(m_Row, gPitchColumn, Cmd.Pitch)
            End If
            If m_Execution.m_ItemLUT.GetFlag(gFillHeightColumn) Then
                SetCellValue(m_Row, gFillHeightColumn, Cmd.FilledHeight)
            End If
            If m_Execution.m_ItemLUT.GetFlag(gNoOfRunColumn) Then
                SetCellValue(m_Row, gNoOfRunColumn, Cmd.NumofRun)  '1
            End If
            If m_Execution.m_ItemLUT.GetFlag(gSprialColumn) Then
                SetCellValue(m_Row, gSprialColumn, Cmd.SpiralFlag)
            End If
            If m_Execution.m_ItemLUT.GetFlag(gRtAngleColumn) Then
                SetCellValue(m_Row, gRtAngleColumn, Cmd.RotatingAngle)
            End If
            If m_Execution.m_ItemLUT.GetFlag(gEdgeClearColumn) Then
                SetCellValue(m_Row, gEdgeClearColumn, Cmd.EdgeClearance)
            End If
            If m_Execution.m_ItemLUT.GetFlag(gIONumberColumn) Then
                SetCellValue(m_Row, gIONumberColumn, 1)
            End If
        End If

        TraceGCCollect()
    End Sub


    'SetDefaultPositionToLastAbovePt: Not in use
    '                                                                                  

    Private Sub SetDefaultPositionToLastAbovePt(ByVal CmdStr As String)
        'Find the last valid point above current row.  We will set this as the last point used for programming.
        'This function is not in use currently.
        Dim pos(3) As Double
        Dim RefPos(3) As Double
        Dim ValidPtFound As Boolean = False

        Dim CurrentRow As Integer = GetActiveCellRow()
        If m_Execution.m_Pattern.m_ErrorChk.CheckValidPtXY(CmdStr) Then

            If CurrentRow = 1 Then
                CurrentRow = 2
            End If

            ValidPtFound = m_Execution.m_Pattern.Spreadsheet_FindLastValidPoint( _
                AxSpreadsheetProgramming, 1, CurrentRow - 1, pos)

            If False = ValidPtFound Then
                pos(0) = 0
                pos(1) = 0
                pos(2) = 0
            End If

            'Reset default pt position
            m_CamPos(0) = pos(0)
            m_CamPos(1) = pos(1)
            m_CamPos(2) = pos(2)

            'Set ref point
            m_Execution.m_Pattern.Spreadsheet_GetRowLocalReference( _
                AxSpreadsheetProgramming, CurrentRow, RefPos)
            Spreadsheet_SetMotionRef(RefPos)
            pos(0) = pos(0)
            pos(1) = pos(1)
            pos(2) = pos(2)
            MoveToSpreadsheetPoint(pos, "Vision")

        End If

    End Sub

    'Add reference pattern command
    Private Sub TBrefer_ButtonClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ToolBarButtonClickEventArgs) Handles ReferenceCommandBlock.ButtonClick
        Dim buttonText As String = e.Button.Text
        buttonText = buttonText.Trim(" ")
        m_EditStateFlag = False
        m_ProgrammingStateFlag = True
        UndoData_Logging(0)
        TBrefer_Implementation(buttonText)
        TraceGCCollect()
        AxSpreadsheetProgramming.Enabled = False
    End Sub

    Private Sub TBrefer_Implementation(ByVal ButtonText As String)
        Dim cell1, cell2 As Object
        NeedleMode.Enabled = True
        SheetRowSelected = False
        ToggleButtonsForTeachingStart()
        HideSpreadsheetTabs()

        Dim sel As Microsoft.Office.Interop.OWC.Range = AxSpreadsheetProgramming.Selection
        Dim m_rowStart As Integer = sel.Row

        Dim RefPos(3) As Double
        If "Reference" = ButtonText Then
            RefPos(0) = 0 : RefPos(1) = 0 : RefPos(2) = 0
            Spreadsheet_SetMotionRef(RefPos)
        Else
            m_Execution.m_Pattern.Spreadsheet_GetRowLocalReference( _
                AxSpreadsheetProgramming, m_rowStart, RefPos)
            Spreadsheet_SetMotionRef(RefPos)
        End If

        type = ButtonText
        m_TeachStepNumber = 1
        Try
            Select Case (ButtonText)
                Case "Fiducial"
                    VisionMode.Checked = True
                    NeedleMode.Checked = False
                    VisionMode.Enabled = False
                    NeedleMode.Enabled = False
                    TeachCommand(type, cell1, cell2)
                Case "Reference"            'Absolute coordinate only
                    m_ReferPt(0) = 0.0
                    m_ReferPt(1) = 0.0
                    m_ReferPt(2) = 0.0
                    TeachCommand(type, cell1, cell2)
                Case "Height"
                    TeachCommand(type, cell1, cell2)
                    NeedleMode.Enabled = False
                Case "Reject"
                    TeachCommand(type, cell1, cell2)
            End Select

        Catch ex As SystemException
            ExceptionDisplay(ex)
        End Try

    End Sub

    'Add element pattern command
    Friend Sub TBElements_ButtonClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ToolBarButtonClickEventArgs) Handles ElementsCommandBlock.ButtonClick

        Dim buttonText As String = e.Button.Text
        buttonText = buttonText.Trim(" ")

        Dim SheetName As String = GetActiveSheetName()
        If m_Execution.m_Pattern.Spreadsheet_IsAnArraySheet(AxSpreadsheetProgramming, SheetName) Then
            MessageBox.Show("Command is not allowed in an Array sheet", "Warnning!")
        Else
            m_EditStateFlag = False
            m_ProgrammingStateFlag = True
            AxSpreadsheetProgramming.Enabled = False
            NeedleMode.Enabled = True
            TeachElementCommand(buttonText)
            If "Measure" <> buttonText Then
                UndoData_Logging(0)
            End If
        End If

    End Sub
    Public ArrayDlg As ArrayGenerate
    'Add element pattern command implementation
    Friend Sub TeachElementCommand(ByVal ButtonText As String)

        Dim i As Integer
        Dim Offset(3) As Double

        SheetRowSelected = False
        ToggleButtonsForTeachingStart()

        Dim pos(3), RefPos(3) As Double
        Dim sel As Microsoft.Office.Interop.OWC.Range = AxSpreadsheetProgramming.Selection
        Dim m_rowStart As Integer = sel.Row
        Dim m_rowEnd As Integer = (m_rowStart + sel.Rows.Count() - 1)

        Try
            HideSpreadsheetTabs()

            If "Offset" = ButtonText Then
            ElseIf "Reference" = ButtonText Then
                RefPos(0) = 0 : RefPos(1) = 0 : RefPos(2) = 0
                Spreadsheet_SetMotionRef(RefPos)
            Else
                m_Execution.m_Pattern.Spreadsheet_GetRowLocalReference( _
                    AxSpreadsheetProgramming, m_rowStart, RefPos)
                Spreadsheet_SetMotionRef(RefPos)
            End If

            Select Case (ButtonText)
                Case "Offset"
                    DisableTeachingToolbarOKButton()
                    Dim xyzOffsetDlg As New FormOffsetInput
                    If xyzOffsetDlg.ShowDialog() = DialogResult.OK Then
                        Offset(0) = CDbl(xyzOffsetDlg.TextBoxX.Text)
                        Offset(1) = CDbl(xyzOffsetDlg.TextBoxY.Text)
                        Offset(2) = CDbl(xyzOffsetDlg.TextBoxZ.Text)

                        Dim CmdStr As String

                        For i = m_rowStart To m_rowEnd
                            CmdStr = GetCellValue(i, gCommandNameColumn)
                            If m_Execution.m_Pattern.m_ErrorChk.CheckValidPtXY(CmdStr) Then
                                If m_Execution.m_Pattern.m_ErrorChk.CheckRequiredPt3XY(CmdStr) Then
                                    SetCellValue(i, gPos1XColumn, (GetCellValue(i, gPos1XColumn) + Offset(0)))
                                    SetCellValue(i, gPos1YColumn, (GetCellValue(i, gPos1YColumn) + Offset(1)))
                                    SetCellValue(i, gPos1ZColumn, (GetCellValue(i, gPos1ZColumn) + Offset(2)))
                                    SetCellValue(i, gPos2XColumn, (GetCellValue(i, gPos2XColumn) + Offset(0)))
                                    SetCellValue(i, gPos2YColumn, (GetCellValue(i, gPos2YColumn) + Offset(1)))
                                    SetCellValue(i, gPos2ZColumn, (GetCellValue(i, gPos2ZColumn) + Offset(2)))
                                    SetCellValue(i, gPos3XColumn, (GetCellValue(i, gPos3XColumn) + Offset(0)))
                                    SetCellValue(i, gPos3YColumn, (GetCellValue(i, gPos3YColumn) + Offset(1)))
                                    SetCellValue(i, gPos3ZColumn, (GetCellValue(i, gPos3ZColumn) + Offset(2)))
                                ElseIf m_Execution.m_Pattern.m_ErrorChk.CheckRequiredPt2XY(CmdStr) Then
                                    SetCellValue(i, gPos1XColumn, (GetCellValue(i, gPos1XColumn) + Offset(0)))
                                    SetCellValue(i, gPos1YColumn, (GetCellValue(i, gPos1YColumn) + Offset(1)))
                                    SetCellValue(i, gPos1ZColumn, (GetCellValue(i, gPos1ZColumn) + Offset(2)))
                                    SetCellValue(i, gPos2XColumn, (GetCellValue(i, gPos2XColumn) + Offset(0)))
                                    SetCellValue(i, gPos2YColumn, (GetCellValue(i, gPos2YColumn) + Offset(1)))
                                    SetCellValue(i, gPos2ZColumn, (GetCellValue(i, gPos2ZColumn) + Offset(2)))
                                Else
                                    SetCellValue(i, gPos1XColumn, (GetCellValue(i, gPos1XColumn) + Offset(0)))
                                    SetCellValue(i, gPos1YColumn, (GetCellValue(i, gPos1YColumn) + Offset(1)))
                                    If m_Execution.m_Pattern.m_ErrorChk.CheckRequiredPtHeight(CmdStr) Then
                                        SetCellValue(i, gPos1ZColumn, (GetCellValue(i, gPos1ZColumn) + Offset(2)))
                                    End If
                                End If
                            Else
                                m_ProgrammingStateFlag = False
                                Exit Select
                            End If
                        Next
                        'Move robot arm to the last offset position
                        pos(0) = GetCellValue(m_rowEnd, gPos1XColumn)
                        pos(1) = GetCellValue(m_rowEnd, gPos1YColumn)
                        pos(2) = GetCellValue(m_rowEnd, gPos1ZColumn)
                        m_Execution.m_Pattern.Spreadsheet_GetRowLocalReference(AxSpreadsheetProgramming, m_rowEnd, RefPos)
                        Spreadsheet_SetMotionRef(RefPos)
                        pos(0) = pos(0)
                        pos(1) = pos(1)
                        pos(2) = pos(2)
                        MoveToSpreadsheetPoint(pos, "Vision")
                        AxSpreadsheetProgramming.Enabled = True
                    End If
                    AxSpreadsheetProgramming.Enabled = True
                    m_ProgrammingStateFlag = False

                Case "Measure"
                    AxSpreadsheetProgramming.Enabled = True
                Case "Dot", "Move", "Line", "Arc", "Circle", "Rectangle", _
                    "FillCircle", "FillRectangle", "ChipEdge", "QC", "SubPattern", "Wait", "EdgeDetection", "DotArray", "GlobalQC"
                    type = ButtonText
                    If Not (type = "GlobalQC") Then
                        m_TeachStepNumber = 1
                        m_Row = GetActiveCellRow()
                        AddCommandToSpreadsheet(type)
                        DisableCoordinateUpdateInSpreadsheet()
                    End If
                    If m_InLinkRange Then
                    Else
                        cell1 = GetCell(m_Row, gPos1XColumn)
                        cell2 = GetCell(m_Row, gPos1ZColumn)
                        SelectRange(cell1, cell2) 'Select the related cell in spreadsheet
                    End If
                    If Not (ButtonText = "SubPattern") And Not (ButtonText = "QC") And Not (ButtonText = "GlobalQC") And Not (ButtonText = "ChipEdge") And Not (ButtonText = "DotArray") Then
                        EnableCoordinateUpdateInSpreadsheet() 'yy

                    End If

                    ToolBarSwitch("YesNo")

                    If ButtonText = "ChipEdge" Then
                        'DisableTeachingButtons()
                        'DisableTeachingToolbar()
                        Vision.FrmVision.EnableClickToMoveMode(True)
                        DisableElementsCommandBlockButton(gOffsetCmdIndex)
                        DisableProgrammingBrightnessToggle()
                        Vision.FrmVision.ResetChipEdgePoint()
                        Vision.IDSV_Form_CE(ValueBrightness.Value)
                        Vision.FrmVision.ChipEdge5PointsSet = New DLL_Export_Device_Vision.FormVision.ChipEdgeSetDelegate(AddressOf Me.ChipEdgeAdjustXY)
                        Dim status As Integer = 0
                        Dim x, y As Double
                        DisableCalibButtons()
                        Vision.ChipEdgePoints_form.FormCloseEvent = New DLL_Export_Device_Vision.ChipEdgePoints.FormCloseDelegate(AddressOf Me.AddChipEdgeCommandFormResponse)
                        EnableClickToMove()
                        '    While Not (Vision.RobotMotionOffset(x, y) = True Or Vision.GetChipEdgeStatus = 2)
                        '        TraceDoEvents()
                        '    End While
                        '    'moverobot
                        '    pos(0) = x
                        '    pos(1) = -y
                        '    pos(2) = 0
                        '    If Not (x = 0 And y = 0) Then
                        '        m_Tri.Set_XY_Speed(IDS.Data.Hardware.Gantry.ElementXYSpeed)
                        '        m_Tri.MoveRelative_XY(pos)
                        '    End If

                        '    If status = 3 Then
                        '        status = 0 'reset status after 5 points being reset
                        '    End If
                        '    While status = 0
                        '        TraceDoEvents()
                        '        status = Vision.GetChipEdgeStatus()
                        '    End While
                        'Loop While status = 3 'status 3= reset 5 points
                        'EnableCalibButtons()
                        ''DelayForRowDelete()
                        'DisableCoordinateUpdateInSpreadsheet()

                        'If status = 2 Then
                        '    DeleteRow(m_Row)
                        '    UpdateSpreadsheet()
                        '    DeletingRowFromExcel = False
                        '    DeletingRowFinished = False
                        '    AxSpreadsheetProgramming.Enabled = True
                        'ElseIf status = 1 Then 'chipedge finished settings
                        '    SetChipEdgeSettings()
                        '    ElementsCommandBlock.Enabled = True
                        '    ReferenceCommandBlock.Enabled = True
                        '    DisplaySpreadsheetTabs()
                        '    SelectCell(m_Row + 1, 1)
                        'End If
                        'AxSpreadsheetProgramming.Enabled = True
                        'ToggleButtonsForTeachingStop()
                        'ClearAndDisplayIndicator()
                        'EnableTeachingButtons()
                        'EnableTeachModeSwitching()
                        'm_ProgrammingStateFlag = False
                        'm_EditStateFlag = False
                        'EnableProgrammingBrightnessToggle()
                        'DisableElementsCommandBlockButton(gSeperatorCmdIndex)
                        'DisableElementsCommandBlockButton(gOffsetCmdIndex)
                        'DisableTeachingToolbarOKButton()
                        'EnableTeachingToolbarCancelButton()
                        'DisplaySpreadsheetTabs()

                    ElseIf ButtonText = "QC" Or ButtonText = "GlobalQC" Then
                        Dim str As String = GetCellValue(1, gCommandNameColumn)
                        If str.ToUpper = "GLOBALQC" Then
                            Vision.SetAllowGlobalQC(False)
                        Else
                            Vision.SetAllowGlobalQC(True)
                        End If
                        DisableTeachingButtons()
                        DisableElementsCommandBlockButton(gOffsetCmdIndex)
                        DisableTeachingToolbar()
                        DisableProgrammingBrightnessToggle()
                        Vision.QC_form.FormCloseEvent = New DLL_Export_Device_Vision.QC.FormCloseDelegate(AddressOf Me.AddQCCommandFormResponse)
                        Vision.IDSV_Form_QC(ValueBrightness.Value)
                        DisableCalibButtons()
                        Me.EnableClickToMove()
                        'Dim status As Integer = 0
                        'Try
                        '    While status = 0 Or status = 3
                        '        Do
                        '            TraceDoEvents()
                        '            status = Vision.GetQCStatus()
                        '        Loop While status = 3
                        '    End While
                        'Catch ex As Exception
                        '    ExceptionDisplay(ex)
                        'End Try
                        'EnableCalibButtons()
                        'If status = 2 Then 'Cancel
                        '    DelayForRowDelete()
                        '    DisableCoordinateUpdateInSpreadsheet()
                        '    DeleteRow(m_Row)
                        '    UpdateSpreadsheet()
                        '    DeletingRowFromExcel = False
                        '    DeletingRowFinished = False
                        'ElseIf status = 1 Then 'Ok
                        '    If Vision.GetIsGlobalQC() Then
                        '        DeleteRow(m_Row)
                        '        Me.AddGlobalQCToTop()
                        '        DisableCoordinateUpdateInSpreadsheet()
                        '        SetGlobalQCSettings()
                        '        IDS.Data.Hardware.Camera.DotQCEnable = True
                        '        SetCellValue(1, gPos1XColumn, Nothing)
                        '        SetCellValue(1, gPos1YColumn, Nothing)
                        '        SetCellValue(1, gPos1ZColumn, Nothing)
                        '        ' DisableElementsCommandBlockButton(gGlobalQCCmdIndex)
                        '    Else
                        '        SetQCSettings()
                        '    End If
                        '    DisableCoordinateUpdateInSpreadsheet()
                        '    SelectCell(m_Row + 1, 1)
                        'End If

                        'EnableProgrammingBrightnessToggle()
                        'EnableTeachModeSwitching()
                        'EnableTeachingButtons()
                        'DisableElementsCommandBlockButton(gSeperatorCmdIndex)
                        'DisableElementsCommandBlockButton(gOffsetCmdIndex)
                        'DisableTeachingToolbarOKButton()
                        'EnableTeachingToolbarCancelButton()
                        'DisplaySpreadsheetTabs()
                        'm_ProgrammingStateFlag = False
                        'ClearAndDisplayIndicator()
                        'AxSpreadsheetProgramming.Enabled = True
                    End If

                Case "Link"

                    type = "LinkStart"
                    m_TeachStepNumber = 1
                    m_Row = GetActiveCellRow()
                    AddCommandToSpreadsheet(type)
                    SelectCell(m_Row + 1, 1)
                    type = "LinkEnd"
                    m_Row = GetActiveCellRow()
                    AddCommandToSpreadsheet(type)
                    ToolBarSwitch("YesNo")
                    m_Row = GetActiveCellRow()
                    DisableCoordinateUpdateInSpreadsheet()

                    EnableElementsCommandBlockButton(gLineCmdIndex)
                    EnableElementsCommandBlockButton(gArcCmdIndex)
                    DisableTeachingToolbarOKButton()
                    EnableTeachingToolbarCancelButton()
                    type = "Link"

                Case "Clean", "Purge", "GetIO", "SetIO", "VolumeCalibration"

                    type = ButtonText
                    m_TeachStepNumber = 1
                    m_Row = GetActiveCellRow()
                    AddCommandToSpreadsheet(type)
                    DisableCoordinateUpdateInSpreadsheet()
                    ToggleButtonsForTeachingStop()
                    m_Row = m_Row + 1
                    SelectCell(m_Row, 1)

                    m_ProgrammingStateFlag = False
                    AxSpreadsheetProgramming.Enabled = True

                Case "Array"
                    CBExpandSpreadsheet.Enabled = False

                    DisableTeachingToolbarOKButton()
                    DisableTeachingToolbarCancelButton()
                    type = ButtonText
                    'sj modify 28/09/09
                    m_Row = GetActiveCellRow()

                    AddCommandToSpreadsheet(type)

                    EnableCoordinateUpdateInSpreadsheet()
                    ArrayDlg = New ArrayGenerate
                    ArrayDlg.isVisionMode = VisionMode.Checked
                    ArrayDlg.SetDefaultPara(m_Execution.m_Pattern.m_CurrentDPara)
                    If ArrayDlg.isVisionMode Then
                        ArrayDlg.SetNeedleGap("0.5")
                    Else
                        ArrayDlg.SetNeedleGap("0.000")
                    End If

                    ArrayDlg.SetDispenseDuration("500")
                    ArrayDlg.BringToFront()
                    ArrayDlg.FormCloseEvent = New ArrayGenerate.FormCloseDelegate(AddressOf Me.AddArrayCommandFormResponse)
                    DisableCalibButtons()
                    ArrayDlg.Show()
                    'Dim DlgReturn = ArrayDlg.DialogResult()
                    'Dim PointX, PointY, PointZ As Double
                    'DisableCalibButtons()
                    'Do
                    '    TraceDoEvents()
                    '    PointX = CDbl(GetCellValue(m_Row, gPos1XColumn))
                    '    PointY = CDbl(GetCellValue(m_Row, gPos1YColumn))
                    '    PointZ = CDbl(GetCellValue(m_Row, gPos1ZColumn))

                    '    ArrayDlg.SetPoint(PointX, PointY, PointZ)
                    '    DlgReturn = ArrayDlg.DialogResult()
                    'Loop While Nothing = DlgReturn
                    'EnableCalibButtons()
                    'DisableCoordinateUpdateInSpreadsheet()

                    'If DialogResult.OK = DlgReturn Then
                    '    'Get "Array" data succesfully.  Data will be used to generate "Array"
                    '    LabelMessage("Array generating starts")
                    '    Dim PatternLineRecord(3) As CIDSPattern.PatternRecord
                    '    Dim arraydata As ArrayData
                    '    arraydata.FirstX = 10.0
                    '    Dim dotarray As New FormArraySetting(arraydata)

                    '    If dotarray.ShowDialog() = DialogResult.OK Then
                    '        dotarray.GetArrayData(arraydata)

                    '        Dim cmdTmpString As String = ArrayDlg.ArrayPara1.Name

                    '        Spreadsheet_GetArrayRecord(PatternLineRecord(0), 1, ArrayDlg)
                    '        Spreadsheet_GetArrayRecord(PatternLineRecord(1), 2, ArrayDlg)
                    '        Spreadsheet_GetArrayRecord(PatternLineRecord(2), 3, ArrayDlg)

                    '        Dim ActSheetName As String = GetActiveSheetName()
                    '        Spreadsheet_GenerateArraySubSheetName(iSubSheetName, ActSheetName, cmdTmpString)

                    '        Dim file As New CIDSFileHandler
                    '        ClearRow(m_Row)

                    '        SetCellValue(m_Row, gCommandNameColumn, "Array")
                    '        SetCellValue(m_Row, gSubnameColumn, file.ExtOnlyFromFullPath(iSubSheetName))
                    '        SelectCell(m_Row + 1, gCommandNameColumn)

                    '        AxSpreadsheetProgramming.Sheets.Add.Name = iSubSheetName
                    '        AxSpreadsheetProgramming.ActiveWindow.FreezePanes = False
                    '        AxSpreadsheetProgramming.Worksheets(iSubSheetName).Range("B1:B1").Select()
                    '        AxSpreadsheetProgramming.ActiveWindow.FreezePanes = True

                    '        'Build Sub array sheet
                    '        Spreadsheet_AddSheetRecord(PatternLineRecord(0), _
                    '            PatternLineRecord(1), PatternLineRecord(2), arraydata)

                    '        'Load sub and array data within the sub
                    '        'Spreadsheet_AddSubandArrayInSub(PatternLineRecord(0).pPara.DispenseFlag)
                    '        'yy
                    '        If Not (ArrayDlg.ArrayPara1.DispenseFlag.ToUpper = "ON" Or ArrayDlg.ArrayPara1.DispenseFlag.ToUpper = "OFF ") Then
                    '            Dim file1 As String = ArrayDlg.ArrayPara1.DispenseFlag
                    '            If 0 = m_Execution.m_Pattern.Spreadsheet_CheckSubsheetExist( _
                    '        AxSpreadsheetProgramming, m_Execution.m_File.NameOnlyFromFullPath(file1)) Then
                    '                m_Execution.m_Pattern.LoadTxtPatternPara(AxSpreadsheetProgramming, file1, 1, 0, False)
                    '            End If
                    '        End If

                    '        EnableTeachingButtons()
                    '        DisableElementsCommandBlockButton(gOffsetCmdIndex)
                    '        DisableTeachingToolbarOKButton()
                    '        EnableTeachingToolbarCancelButton()
                    '    Else
                    '        DelayForRowDelete()
                    '        EnableTeachingButtons()
                    '        DisableElementsCommandBlockButton(gOffsetCmdIndex)
                    '        DisableTeachingToolbarOKButton()
                    '        EnableTeachingToolbarCancelButton()
                    '        TraceDoEvents()
                    '        DeleteRow(m_Row)
                    '        DeletingRowFromExcel = False
                    '        DeletingRowFinished = False
                    '    End If
                    'Else
                    '    DelayForRowDelete()
                    '    EnableTeachingButtons()
                    '    DisableElementsCommandBlockButton(gOffsetCmdIndex)
                    '    DisableTeachingToolbarOKButton()
                    '    EnableTeachingToolbarCancelButton()
                    '    TraceDoEvents()
                    '    DeleteRow(m_Row)
                    '    DeletingRowFromExcel = False
                    '    DeletingRowFinished = False

                    'End If
                    'AxSpreadsheetProgramming.Enabled = True
                    'm_ProgrammingStateFlag = False

                    'CBExpandSpreadsheet.Enabled = True
            End Select
        Catch ex As SystemException

            'kr should i put this in
            m_TeachStepNumber = 1
            DisableCoordinateUpdateInSpreadsheet()
            ToggleButtonsForTeachingStart()
            m_ProgrammingStateFlag = False
            DeletingRowFromExcel = False
            DeletingRowFinished = False
            EnableTeachingButtons()

            ExceptionDisplay(ex)
        End Try
        TraceGCCollect()
    End Sub

    'spreadsheet operation ini
    '                                                                                  

    Public Sub init_spreadsheet()

        Dim Title(gMaxColumns) As String
        Title(gCommandNameColumn - 1) = "Command"
        Title(gNeedleColumn - 1) = "Needle"
        Title(gDispensColumn - 1) = "Dispensing"
        Title(gPos1XColumn - 1) = "X1"
        Title(gPos1YColumn - 1) = "Y1"
        Title(gPos1ZColumn - 1) = "Z1"
        Title(gPos2XColumn - 1) = "X2"
        Title(gPos2YColumn - 1) = "Y2"
        Title(gPos2ZColumn - 1) = "Z2"
        Title(gPos3XColumn - 1) = "X3"
        Title(gPos3YColumn - 1) = "Y3"
        Title(gPos3ZColumn - 1) = "Z3"
        Title(gTravelSpeedColumn - 1) = "TravelSpeed"

        Title(gNeedleGapColumn - 1) = "NeedleGap"
        Title(gDurationColumn - 1) = "Duration"
        Title(gTravelDelayColumn - 1) = "TravelDelay"
        Title(gDTaiDistColumn - 1) = "DTail Dist"
        Title(gApproachHtColumn - 1) = "Appr. Height"
        Title(gRetractDelayColumn - 1) = "RetractDelay"
        Title(gRetractSpeedColumn - 1) = "RetractSpeed"
        Title(gRetractHtColumn - 1) = "RetractHeight"
        Title(gClearanceHtColumn - 1) = "ClearanceHeight"
        Title(gArcRadiusColumn - 1) = "Arc Radius"
        Title(gPitchColumn - 1) = "Pitch"
        Title(gNoOfRunColumn - 1) = "RoundNo"
        Title(gFillHeightColumn - 1) = "FillHeight"
        Title(gRtAngleColumn - 1) = "Rt Angle"
        Title(gEdgeClearColumn - 1) = "EdgeClear"
        Title(gSprialColumn - 1) = "Sprial"
        Title(gIONumberColumn - 1) = "IONumber"

        Dim i As Integer
        With AxSpreadsheetProgramming
            .ViewableRange = "$A1:$AD30000"
            .DisplayColumnHeadings = True
            For i = gCommandNameColumn To gIONumberColumn
                .Windows(1).ColumnHeadings(i).Caption = Title(i - 1)
            Next i
            .ActiveSheet.Range("A:A").ColumnWidth = 10
            .ActiveSheet.Range("B:B").ColumnWidth = 6
            .ActiveSheet.Range("C:C").ColumnWidth = 10
            .ActiveSheet.Range("D:D").ColumnWidth = 8   'X1
            .ActiveSheet.Range("E:E").ColumnWidth = 8   'Y1
            .ActiveSheet.Range("F:F").ColumnWidth = 8   'Z1
            .ActiveSheet.Range("G:G").ColumnWidth = 8   'X2
            .ActiveSheet.Range("H:H").ColumnWidth = 8   'Y2
            .ActiveSheet.Range("I:I").ColumnWidth = 8   'Z2
            .ActiveSheet.Range("J:J").ColumnWidth = 8   'X3
            .ActiveSheet.Range("K:K").ColumnWidth = 8   'Y3
            .ActiveSheet.Range("L:L").ColumnWidth = 8   'Z3
        End With
        SelectCell(1, 1)
        AxSpreadsheetProgramming.ActiveWindow.EnableResize = False
        DisplaySpreadsheetTabs()
        TraceGCCollect()
    End Sub



    'Initialize a sheet
    '   e: ActiveX component event handler
    '

    Private Sub Spreadsheet_Initialize(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AxSpreadsheetProgramming.Initialize
        init_spreadsheet()
        m_Row = 1
        TraceGCCollect()
    End Sub


    'Activiate a sheet
    '   e: ActiveX component event handler
    '

    Private Sub Spreadsheet_SheetActivate(ByVal sender As System.Object, ByVal e As AxOWC10.ISpreadsheetEventSink_SheetActivateEvent) Handles AxSpreadsheetProgramming.SheetActivate

        Dim sheetname As String = e.sh.Name
        init_spreadsheet()
        If GetActiveSheetName() = "Main" Then
            SelectCell(m_Row + 1, 1)
        End If

        TraceGCCollect()
    End Sub



    'Check it can be copied or not
    '   AxSpreadsheet: ActiveX component
    '
    'Return: True=Can be copied, False=Cannot

    Private Function Spreadsheet_CanBeCopyCut(ByRef AxSpreadsheet As AxOWC10.AxSpreadsheet) As Boolean
        Dim Rtn As Boolean = False
        Dim sel As Microsoft.Office.Interop.OWC.Range = AxSpreadsheetProgramming.Selection

        Dim m_rowCount As Integer = sel.Rows.Count()
        Dim m_columnCount As Integer = sel.Columns.Count()
        Dim m_StartRow As Integer = sel.Row
        Dim m_EndRow As Integer = m_StartRow + m_rowCount - 1
        Dim SelectedPageName As String = AxSpreadsheet.ActiveWorkbook.ActiveSheet.Name()

        Dim bStartRowInLink As Boolean = Spreadsheet_WithinLinkRangeForCopyCut(m_StartRow)
        Dim bEndRowInLink As Boolean = Spreadsheet_WithinLinkRangeForCopyCut(m_EndRow)
        Dim bStartEndRowInTheSameLink As Boolean


        If 1 = m_rowCount Then
            If "LinkStart" = GetCellValue(m_StartRow, gCommandNameColumn) Or _
                "LinkEnd" = GetCellValue(m_StartRow, gCommandNameColumn) Then
                Return Rtn
            End If
        End If

        If gMaxColumns <> m_columnCount Then
            Return Rtn
        End If

        If bStartRowInLink And bEndRowInLink Then       'May possible be copied or cut, need to do further chk
            Rtn = Spreadsheet_WithinLinkRangeForCopyCut(m_StartRow, m_EndRow)  'Is a identity link

        ElseIf bStartRowInLink Then                     'may not be copied and cut
            If GetCellValue(m_StartRow, gCommandNameColumn) <> "LinkStart" Then
                Rtn = False
            End If

        ElseIf bEndRowInLink Then                       'may not be copied and cut
            If GetCellValue(m_EndRow, gCommandNameColumn) <> "LinkEnd" Then
                Rtn = False
            Else
                Rtn = True
            End If
        Else                                            'Can be copied or cut
            Rtn = True
        End If

        TraceGCCollect()
        Return Rtn

    End Function



    'Handle Del Key
    '   e: ActiveX component event handler
    '

    Dim cellString As String
    Private Sub Spreadsheet_BeforeButtonDelInput(ByVal sender As System.Object, ByVal e As AxOWC10.ISpreadsheetEventSink_BeforeKeyDownEvent) Handles AxSpreadsheetProgramming.BeforeKeyDown

        Dim sel As Microsoft.Office.Interop.OWC.Range = AxSpreadsheetProgramming.Selection

        Dim m_rowCount As Integer = sel.Rows.Count()
        Dim m_columnCount As Integer = sel.Columns.Count()
        Dim row As Integer = GetActiveCellRow()
        Dim colum As Integer = GetActiveCellColumn()
        Dim keyValue As Integer



        If e.keyCode = 46 Then  'the delete key 
            If m_columnCount = gMaxColumns And m_rowCount = 1 Then
            Else
                cellString = GetCellValue(row, gSubnameColumn)
            End If
        Else
            If Not (colum >= gCommandNameColumn And colum <= gPos3ZColumn) Then
                If (colum <> gDTaiDistColumn) Then
                    keyValue = e.keyCode

                    If (keyValue = 16) Or (keyValue = 107) Or (keyValue = 109) Or (keyValue = 187) Or (keyValue = 189) Then
                        LabelMessage("Invalid key!")
                        e.cancel.Value = True
                    Else
                        cellString = GetCellValue(row, colum)
                    End If
                End If
            End If
        End If
        TraceGCCollect()
    End Sub

    Dim specialKey As Boolean = False
    Private Sub Spreadsheet_AfterButtonDelInput(ByVal sender As System.Object, ByVal e As AxOWC10.ISpreadsheetEventSink_KeyDownEvent) Handles AxSpreadsheetProgramming.KeyDownEvent

        Dim sel As Microsoft.Office.Interop.OWC.Range = AxSpreadsheetProgramming.Selection

        Dim m_rowCount As Integer = sel.Rows.Count()
        Dim m_columnCount As Integer = sel.Columns.Count()
        Dim row As Integer = GetActiveCellRow()
        Dim colum As Integer = GetActiveCellColumn()
        Dim cellValue As String

        If e.keyCode = 46 Then  'the delete key 
            If m_columnCount = gMaxColumns And m_rowCount = 1 Then
                'ConfirmOrCancel_Implementation(0)
            Else
                SetCellValue(row, gSubnameColumn, cellString)
            End If
        Else
            '''''''''''''''''''''''''''''''''''''''''
            '   Xue Wen                             '
            '   Checking for "-", "+" and "=".      '
            '''''''''''''''''''''''''''''''''''''''''
            Dim keyValue As Integer

            keyValue = e.keyCode

            If (keyValue = 16) Or (keyValue = 107) Or (keyValue = 109) Or (keyValue = 187) Or (keyValue = 189) Then
                specialKey = True
            End If
        End If
        If e.keyCode = Keys.ControlKey Then
            isPress = True
        End If
        TraceGCCollect()
    End Sub


    'Start cell edit, tentative result may be rejected
    '   e: ActiveX component event handler
    '

    Private Sub Spreadsheet_StartEdit(ByVal sender As System.Object, ByVal e As AxOWC10.ISpreadsheetEventSink_StartEditEvent) Handles AxSpreadsheetProgramming.StartEdit
        Console.WriteLine("Sp start edit")

        Dim row As Integer = GetActiveCellRow()
        Dim colum As Integer = GetActiveCellColumn()
        Dim cellValue As Object = GetActiveCellValue()
        DisableCoordinateUpdateInSpreadsheet()

        Dim command As String = GetCellValue(row, gCommandNameColumn)
        If command Is Nothing Then
            SetCellValue(row, colum, Nothing)
            Return
        End If
        LabelMessage("Editing......")
        Dim m_ItemLUT As New CIDSItemsLUT
        m_ItemLUT.Cmd2Index(command.ToUpper)
        'This is to check if data insert is allowed for this cell of the particular command.
        If Not (command.ToUpper = "FIDUCIAL") Then
            If m_ItemLUT.GetFlag(colum) = 0 Then
                LabelMessage("Add Value to This Cell Is Not Allowed!")
                SetCellValue(row, colum, Nothing)
                Return
            End If
        End If

        If (colum >= gPos1XColumn And colum <= gPos3ZColumn) Then
            LabelMessage("Can't change value directly")
            '           System.Windows.Forms.MessageBox.Show("not allowed to edit", "Add Dot")
            SetCellValue(row, colum, cellValue)
            Return
        End If
        Console.WriteLine("After Check")
        If colum = gTravelSpeedColumn And cellValue = "DotArray" Then
            Exit Sub
        End If

        If (colum >= gCommandNameColumn And colum <= gDispensColumn) Then
            If colum = gCommandNameColumn Then
                If cellValue = "SubPattern" Or cellValue = "Array" Then
                    Dim SubSheetName As String
                    SubSheetName = GetCellValue(row, gSubnameColumn)
                    SetActiveCellValue(cellValue)
                    AxSpreadsheetProgramming.Worksheets(SubSheetName).Activate()
                    LabelMessage("")
                    Exit Sub
                Else
                    SetCellValue(row, colum, cellValue)
                    LabelMessage("Can't change value directly")
                End If
            ElseIf (colum >= gNeedleColumn And colum <= gDispensColumn) Then
                LabelMessage("Can't change value directly")
                SetCellValue(row, colum, cellValue)
            End If

        ElseIf (colum >= gPos1XColumn And colum <= gPos3ZColumn) Then
            LabelMessage("Can't change value directly")
            '           System.Windows.Forms.MessageBox.Show("not allowed to edit", "Add Dot")
            SetCellValue(row, colum, cellValue)
        Else
            '''''''''''''''''''''''''''''''''''''
            '   Xue Wen                         '
            '   Edit for tailing distance.      '
            '''''''''''''''''''''''''''''''''''''
            Dim rtnVal, oldData, CmdStr As String
            If (colum = gDTaiDistColumn) Then

                CmdStr = GetCellValue(row, gCommandNameColumn)

                If (CmdStr = "Line") Or (CmdStr = "Arc") Or (CmdStr = "Rectangle") Or (CmdStr = "Circle") _
                Or (CmdStr = "FillRectangle") Or (CmdStr = "FillCircle") Or (CmdStr = "ChipEdge") Then

                    oldData = GetCellValue(row, colum)
                    rtnVal = InputBox("Please key in the tailing distance.", "Tailing Distance", oldData)

                    If (rtnVal = "") Then
                        SetCellValue(row, colum, oldData)
                    Else
                        SetCellValue(row, colum, rtnVal)
                    End If

                    Exit Sub
                Else
                    If (specialKey = True) Then
                        LabelMessage("Invalid key!")
                        If (cellString = Nothing) Then
                            SetCellValue(row, colum, Nothing)
                        Else
                            SetCellValue(row, colum, cellString)
                        End If
                        specialKey = False
                        Exit Sub
                    End If
                End If
            Else
                '''''''''''''''''''''''''''''''''''''''''
                '   Checking for "-", "+" and "=".      '
                '''''''''''''''''''''''''''''''''''''''''

                If (specialKey = True) Then
                    LabelMessage("Invalid key!")
                    If (cellString = Nothing) Then
                        SetCellValue(row, colum, Nothing)
                    Else
                        SetCellValue(row, colum, cellString)
                    End If
                    specialKey = False
                    Exit Sub
                End If
            End If

            ''''''''''''''''''''''''''
            '   Xue Wen                                                                                 '
            '   If the system is not in "Teaching" or "Editing" mode, it should not show this message.  '
            '   Note: Test more.                                                                        '
            ''''''''''''''''''''''''''
            If Not ((m_EditStateFlag = False) And (m_ProgrammingStateFlag = False)) Then
                LabelMessage("Please input a new value")
            End If

            If m_TeachStepNumber = 1 Then
                If (colum >= gPos1XColumn And colum <= gPos1ZColumn) Then
                Else
                End If
            ElseIf m_TeachStepNumber = 2 Then
                If (colum >= gPos2XColumn And colum <= gPos2ZColumn) Then
                Else
                End If
            ElseIf m_TeachStepNumber = 3 Then
                If (colum >= gPos3XColumn And colum <= gPos3ZColumn) Then
                Else
                End If
            End If
        End If

        TraceGCCollect()
    End Sub


    'End cell edit, result will be finalized
    '   e: ActiveX component event handler
    '

    Private Sub Spreadsheet_EndEdit(ByVal sender As System.Object, ByVal e As AxOWC10.ISpreadsheetEventSink_EndEditEvent) Handles AxSpreadsheetProgramming.EndEdit
        'Return
        Console.WriteLine("Sp end edit")
        LabelMessage("")
        Dim row As Integer = GetActiveCellRow()
        Dim colum As Integer = GetActiveCellColumn()
        Dim dVal As Double

        Dim ErrorFound As Boolean = False

        Dim CmdStr As String = GetCellValue(row, gCommandNameColumn)
        Dim ActSheetName As String = GetActiveSheetName()
        Dim InputStr As String = e.finalValue.Value.ToString()

        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '   Xue Wen                                             '
        '   To avoid "System.NullReferenceException" error.     '
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        specialKey = False
        If (InputStr = Nothing) Then
            Exit Sub
        End If

        If (colum >= gPos1XColumn And colum <= gPos1YColumn) Or _
            (colum >= gPos2XColumn And colum <= gPos2YColumn) Or _
            (colum >= gPos3XColumn And colum <= gPos3YColumn) Then

            ErrorFound = m_Execution.m_Pattern.m_ErrorChk.CheckCellXYError(ActSheetName, InputStr, row, colum, AxSpreadsheetProgramming, ErrorSubSheet)

            If ErrorFound Then
                MyMsgBox("Found edit error, click Ok to restore", MsgBoxStyle.Information + MsgBoxStyle.OKOnly, "Input error")
                e.cancel.Value = True
            Else
                m_Execution.m_Pattern.UpdateDispara(m_Execution.m_Pattern.m_CurrentDPara, InputStr, colum)
                'Move to the place
                Dim pos(3) As Double
                Try
                    Select Case colum
                        Case gPos1XColumn
                            pos(0) = CDbl(InputStr)
                            pos(1) = CDbl(GetCellValue(row, gPos1YColumn))
                            pos(2) = CDbl(GetCellValue(row, gPos1ZColumn))
                        Case gPos1YColumn
                            pos(0) = CDbl(GetCellValue(row, gPos1XColumn))
                            pos(1) = CDbl(InputStr)
                            pos(2) = CDbl(GetCellValue(row, gPos1ZColumn))
                        Case gPos2XColumn
                            pos(0) = CDbl(InputStr)
                            pos(1) = CDbl(GetCellValue(row, gPos2YColumn))
                            pos(2) = CDbl(GetCellValue(row, gPos2ZColumn))
                        Case gPos2YColumn
                            pos(0) = CDbl(GetCellValue(row, gPos2XColumn))
                            pos(1) = CDbl(InputStr)
                            pos(2) = CDbl(GetCellValue(row, gPos2ZColumn))
                        Case gPos3XColumn
                            pos(0) = CDbl(InputStr)
                            pos(1) = CDbl(GetCellValue(row, gPos3YColumn))
                            pos(2) = CDbl(GetCellValue(row, gPos3ZColumn))
                        Case gPos3YColumn
                            pos(0) = CDbl(GetCellValue(row, gPos3XColumn))
                            pos(1) = CDbl(InputStr)
                            pos(2) = CDbl(GetCellValue(row, gPos3ZColumn))
                    End Select
                    'If CmdStr = "ChipEdge" Or CmdStr = "QC" Or CmdStr = "Height" Or CmdStr = "Fiducial" Or CmdStr = "Reject" Then
                    '    MoveToSpreadsheetPoint(pos, "Vision")
                    'Else
                    '    MoveToSpreadsheetPoint(pos, "Needle")
                    'End If
                Catch ex As Exception
                    LabelMessage("invalid value")
                    e.cancel.Value = True 'Cancel the edit
                End Try
            End If

        Else

            ErrorFound = m_Execution.m_Pattern.m_ErrorChk.CheckOtherCellCharError(CmdStr, InputStr, colum)

            If Not ErrorFound Then
                If colum = gPos1ZColumn Or colum = gPos2ZColumn Or colum = gPos3ZColumn Or colum = gClearanceHtColumn Or _
                    colum = gRetractHtColumn Or colum = gNeedleGapColumn Then

                    If m_Execution.m_Pattern.m_ErrorChk.CheckRequiredPtHeight(CmdStr) Then
                        Dim CurrentHeight, fClearanceHeight, fRetractHeight, fNeedleGap As Double

                        CurrentHeight = CDbl(GetCellValue(row, gApproachHtColumn))
                        fClearanceHeight = CDbl(GetCellValue(row, gClearanceHtColumn))
                        fRetractHeight = CDbl(GetCellValue(row, gRetractHtColumn))
                        fNeedleGap = CDbl(GetCellValue(row, gNeedleGapColumn))

                        'Using tentative value for height error checking
                        If colum = gApproachHtColumn Then
                            CurrentHeight = CDbl(InputStr)
                        ElseIf colum = gClearanceHtColumn Then
                            fClearanceHeight = CDbl(InputStr)
                        ElseIf colum = gRetractHtColumn Then
                            fRetractHeight = CDbl(InputStr)
                        ElseIf colum = gNeedleGapColumn Then
                            fNeedleGap = CDbl(InputStr)
                        End If

                        ErrorFound = m_Execution.m_Pattern.m_ErrorChk.CheckCellHeightError(CurrentHeight, fClearanceHeight, fRetractHeight, fNeedleGap)
                    End If

                End If
            End If

            If ErrorFound Then
                '''''''''''''''''''''''''''''''''''''''''
                If (colum = gTravelSpeedColumn) Or (colum = gRetractSpeedColumn) Then
                    MyMsgBox("The range of travel speed is from 1 to 800, click 'Ok' to restore", MsgBoxStyle.Information + MsgBoxStyle.OKOnly, "Input error")
                ElseIf (colum = gDurationColumn) Or (colum = gTravelDelayColumn) Or (colum = gRetractDelayColumn) Then
                    MyMsgBox("Time doesn't have negetive value, click 'Ok' to restore", MsgBoxStyle.Information + MsgBoxStyle.OKOnly, "Input error")
                Else
                    MyMsgBox("Found edit error:" & m_Execution.m_Pattern.m_ErrorChk.ErrorMessage() & ". Click Ok to restore", MsgBoxStyle.Information + MsgBoxStyle.OKOnly, "Input error")
                End If

                e.cancel.Value = True
            Else
                m_Execution.m_Pattern.UpdateDispara(m_Execution.m_Pattern.m_CurrentDPara, InputStr, colum)
            End If
        End If

        TraceGCCollect()
    End Sub


    'set motion reference point
    '   RefPos: input ref point
    Private Sub Spreadsheet_SetMotionRef(ByVal RefPos() As Double)
        m_ReferPt(0) = RefPos(0)
        m_ReferPt(1) = RefPos(1)
        m_ReferPt(2) = RefPos(2)
    End Sub

    'Check the click event, detail Buttons and flags will be reset accordingly
    '   e: ActiveX component event handler
    Private Sub Spreadsheet_ClickEvent(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AxSpreadsheetProgramming.ClickEvent
        If m_ProgrammingStateFlag Then
            'Dim selected As Microsoft.Office.Interop.OWC.Range = Programming.AxSpreadsheetProgramming.ActiveSheet.Rows(m_Row)
            'Dim c1 As Object = GetCell(m_Row, 3)
            'Dim c2 As Object = GetCell(m_Row, 3 + 2)
            'Programming.AxSpreadsheetProgramming.ActiveSheet.Cells(c1, c2).sel()
            'Return
        End If
        'Console.WriteLine("Click")
        Dim sel As Microsoft.Office.Interop.OWC.Range = AxSpreadsheetProgramming.Selection
        Dim m_rowCount As Integer = sel.Rows.Count()
        Dim m_columnCount As Integer = sel.Columns.Count()
        Dim m_StartRow As Integer = sel.Row
        Dim m_columnStart As Integer = sel.Column
        Dim m_EndRow As Integer = m_StartRow + m_rowCount - 1
        Dim RefPos() As Double = {0, 0, 0}
        Dim commandName As String = GetCellValue(m_StartRow, gCommandNameColumn)
        If m_rowCount = 1 And commandName = Nothing And Not (m_EditStateFlag) And Not (m_ProgrammingStateFlag) Then
            LabelMessage("")
        End If
        If "" = gPatternFileName Then
            Exit Sub
        End If

        Dim CellStr As String
        Dim CmdStr As String

        Dim cell1, cell2 As Object
        Dim inCurrentEditSlot As Boolean = True
        Dim ReselectCells As Boolean = False

        '   Xue Wen                 '
        '   To refresh the label.   '

        If True = m_EditStateFlag Then 'when edit pen clicked and m_EditStageFlag assigned to true
            MenuEditSelectAll.Enabled = False
            CellStr = GetCellValue(m_Row, m_columnStart)
            CmdStr = GetCellValue(m_Row, gCommandNameColumn)
            If m_TeachStepNumber = 1 Then
                If m_columnStart < gPos1XColumn Or m_columnStart > gPos1ZColumn Then
                    inCurrentEditSlot = False
                End If
                If m_Execution.m_Pattern.m_ErrorChk.CheckRequiredPt1XY(CmdStr) Then
                    cell1 = GetCell(m_Row, gPos1XColumn)
                    cell2 = GetCell(m_Row, gPos1ZColumn)
                    ReselectCells = True
                End If

            ElseIf m_TeachStepNumber = 2 Then
                If m_columnStart < gPos2XColumn Or m_columnStart > gPos2ZColumn Then
                    inCurrentEditSlot = False
                End If
                If m_Execution.m_Pattern.m_ErrorChk.CheckRequiredPt2XY(CmdStr) Then
                    cell1 = GetCell(m_Row, gPos2XColumn)
                    cell2 = GetCell(m_Row, gPos2ZColumn)
                    ReselectCells = True
                End If

            ElseIf m_TeachStepNumber = 3 Then
                If m_columnStart < gPos3XColumn Or m_columnStart > gPos3ZColumn Then
                    inCurrentEditSlot = False
                End If
                If m_Execution.m_Pattern.m_ErrorChk.CheckRequiredPt3XY(CmdStr) Then
                    cell1 = GetCell(m_Row, gPos3XColumn)
                    cell2 = GetCell(m_Row, gPos3ZColumn)
                    ReselectCells = True
                End If

            End If

            If m_Row <> m_StartRow Or inCurrentEditSlot = False Then
                If ReselectCells = True Then
                    SelectRange(cell1, cell2)
                End If
                Exit Sub
            End If

        ElseIf True = m_ProgrammingStateFlag Then 'When One of the teaching tool (Fiducial/Line) was selected/clicked
            MenuEditSelectAll.Enabled = False
            CmdStr = GetCellValue(m_Row, gCommandNameColumn)
            If "LinkEnd" = CmdStr Then
                inCurrentEditSlot = False
                cell1 = GetCell(m_Row, gCommandNameColumn)
                cell2 = GetCell(m_Row, gCommandNameColumn)
                ReselectCells = True
            ElseIf m_TeachStepNumber = 1 Then
                If m_columnStart < gPos1XColumn Or m_columnStart > gPos1ZColumn Then
                    inCurrentEditSlot = False
                End If
                If m_Execution.m_Pattern.m_ErrorChk.CheckRequiredPt1XY(CmdStr) Then
                    cell1 = GetCell(m_Row, gPos1XColumn)
                    cell2 = GetCell(m_Row, gPos1ZColumn)
                    ReselectCells = True
                End If

            ElseIf m_TeachStepNumber = 2 Then
                If m_columnStart < gPos2XColumn Or m_columnStart > gPos2ZColumn Then
                    inCurrentEditSlot = False
                End If
                If m_Execution.m_Pattern.m_ErrorChk.CheckRequiredPt2XY(CmdStr) Then
                    cell1 = GetCell(m_Row, gPos2XColumn)
                    cell2 = GetCell(m_Row, gPos2ZColumn)
                    ReselectCells = True
                End If

            ElseIf m_TeachStepNumber = 3 Then
                If m_columnStart < gPos3XColumn Or m_columnStart > gPos3ZColumn Then
                    inCurrentEditSlot = False
                End If
                If m_Execution.m_Pattern.m_ErrorChk.CheckRequiredPt3XY(CmdStr) Then
                    cell1 = GetCell(m_Row, gPos3XColumn)
                    cell2 = GetCell(m_Row, gPos3ZColumn)
                    ReselectCells = True
                End If

            End If

            If m_Row <> m_StartRow Or inCurrentEditSlot = False Then
                If m_Row <> m_StartRow Or ReselectCells = True Then
                    If cell1 <> Nothing Or cell2 <> Nothing Then
                        SelectRange(cell1, cell2)
                    End If
                End If
                Exit Sub
            End If
        End If

        m_Row = sel.Row
        SheetRowSelected = False
        DisableCoordinateUpdateInSpreadsheet()
        MenuEditCopy.Enabled = False
        CellStr = GetCellValue(m_Row, m_columnStart)
        CmdStr = GetCellValue(m_Row, gCommandNameColumn)

        'Set motion ref for any single row selected
        If 1 = m_rowCount Then
            m_Execution.m_Pattern.Spreadsheet_GetRowLocalReference(AxSpreadsheetProgramming, m_Row, RefPos)
            Spreadsheet_SetMotionRef(RefPos) 'This to update the XYZ position on GUI
        End If

        If gMaxRows = m_rowCount And gMaxColumns = m_columnCount Then
            Dim SubSheetName1 As String
            Dim SubSheetName2 As String
            SubSheetName1 = GetActiveSheetName()
            SubSheetName2 = m_Execution.m_Pattern.Spreadsheet_FindParrentSheet(SubSheetName1)

            AxSpreadsheetProgramming.Worksheets(SubSheetName2).Activate()
            DisplaySpreadsheetTabs()
        ElseIf 1 = m_rowCount And 1 = m_columnCount Then
            Dim SubSheetName As String
            If "Array" = CellStr Or "SubPattern" = CellStr Then
                If "SubPattern" = CellStr Then
                    SubSheetName = GetCellValue(m_Row, gSubnameColumn)
                    SubSheetName = m_Execution.m_File.NameOnlyFromFullPath(SubSheetName)
                ElseIf "Array" = CellStr Then
                    SubSheetName = GetActiveSheetName() + "."
                    SubSheetName = SubSheetName + GetCellValue(m_Row, gSubnameColumn)
                End If

                AxSpreadsheetProgramming.Worksheets(SubSheetName).Activate()

                DisableElementsCommandBlockButton(gOffsetCmdIndex)
                DisableEditingToolbar()
            ElseIf m_Execution.m_Pattern.m_ErrorChk.CheckValidPtXY(CmdStr) Then
                If m_Execution.m_Pattern.m_ErrorChk.CheckRequiredPt2XY(CmdStr) Then
                    EnableEditingToolbarSwitchButton()                 'Goto next
                Else
                    DisableEditingToolbarEditButton()
                End If

                DisableEditingToolbarEditButton()                    'Edit start
                EnableTeachingToolbarCancelButton()
                m_SteppingPostFlag = False 'Disable this flag to allow user to delete the entire row when selected new row and X button clicked
            Else
                DisableElementsCommandBlockButton(gOffsetCmdIndex)
                'DisableEditingToolbar()
                Programming.btEdit.Enabled = False
                'Programming.EditingToolBar1.Buttons(1).Enabled = False
            End If

        ElseIf gMaxColumns = m_columnCount Then
            Dim pos(3) As Double

            MenuEditDelete.Enabled = True
            DisableTeachingToolbarOKButton()
            EnableTeachingToolbarCancelButton()

            'Check it can be copied or not
            If Spreadsheet_CanBeCopyCut(AxSpreadsheetProgramming) Then
                MenuEditCopy.Enabled = True
                MenuEditCut.Enabled = True
            Else
                MenuEditCopy.Enabled = False
                MenuEditCut.Enabled = False
            End If

            If 1 = m_rowCount Then
                SheetRowSelected = True
                DisplaySpreadsheetTabs()

                If m_Execution.m_Pattern.m_ErrorChk.CheckValidPtXY(CmdStr) Then
                    m_TeachStepNumber = 1
                    pos(0) = GetCellValue(m_Row, gPos1XColumn)
                    pos(1) = GetCellValue(m_Row, gPos1YColumn)
                    pos(2) = GetCellValue(m_Row, gPos1ZColumn)
                    pos(0) = pos(0)
                    pos(1) = pos(1)
                    pos(2) = pos(2)

                    AxSpreadsheetProgramming.Enabled = False

                    If CmdStr = "ChipEdge" Or CmdStr = "QC" Or CmdStr = "Height" Or CmdStr = "Fiducial" Or CmdStr = "Reject" Then
                        'MoveToSpreadsheetPoint(pos, "Vision")
                    Else
                        'MoveToSpreadsheetPoint(pos, "Needle")
                    End If

                    AxSpreadsheetProgramming.Enabled = True
                    HideSpreadsheetTabs()
                    AxSpreadsheetProgramming.Refresh()

                    'DisableElementsCommandBlock()
                    EnableElementsCommandBlockButton(gOffsetCmdIndex)         'Offset

                    If CmdStr = "GlobalQC" Then
                        EnableEditingToolbarEditButton()
                        DisableEditingToolbarSwitchButton() 'Goto next
                    Else
                        EnableEditingToolbarSwitchButton()                     'Goto next
                        DisableEditingToolbarEditButton()                    'Edit start
                    End If

                    m_SteppingPostFlag = False 'Disable this flag to allow user to delete the entire row when selected new row and X button clicked
                ElseIf "" = CmdStr Then
                    DisableElementsCommandBlockButton(gOffsetCmdIndex)
                    DisableEditingToolbar()
                    DisplaySpreadsheetTabs()

                Else
                    DisableElementsCommandBlockButton(gOffsetCmdIndex)
                    DisableEditingToolbar()

                End If
                Spreadsheet_CheckForWithinLinkRange(True)
            Else
                'For multi-row OFFSET selection, we need to find the last valid point first.  
                DisableTeachingButtons()

                If m_Execution.m_Pattern.Spreadsheet_FindLastValidPoint(AxSpreadsheetProgramming, m_StartRow, m_EndRow, pos) Then
                    EnableElementsCommandBlockButton(gOffsetCmdIndex) 'Offset enabled

                    'For the valid Offset selections, we get the motion ref for the last row
                    m_Execution.m_Pattern.Spreadsheet_GetRowLocalReference(AxSpreadsheetProgramming, m_EndRow, RefPos)
                    Spreadsheet_SetMotionRef(RefPos)

                    'Move robot arm to the point last valid point
                    pos(0) = pos(0)
                    pos(1) = pos(1)
                    pos(2) = pos(2)

                    AxSpreadsheetProgramming.Enabled = False
                    'MoveToSpreadsheetPoint(pos, "Vision")
                    AxSpreadsheetProgramming.Enabled = True
                Else
                    'For the invalid Offset selections, we get the motion ref for the first row
                    m_Execution.m_Pattern.Spreadsheet_GetRowLocalReference(AxSpreadsheetProgramming, m_Row, RefPos)
                    Spreadsheet_SetMotionRef(RefPos)

                    DisableElementsCommandBlockButton(gOffsetCmdIndex)
                End If
                DisableEditingToolbar()
            End If
            'When the Header of column clicked
        ElseIf 1 = m_columnCount Then   'Update elements column-wise, need to be improved later to
            m_Column = sel.Column       'exclude part of data
            DisableTeachingToolbar()
            DisableEditingToolbar()
            DisableTeachingButtons()

            If (m_Column < gTravelSpeedColumn Or m_Column > gSprialColumn Or m_rowCount < 1) And Not (m_Column = gDispensColumn) Then
                Exit Sub
            End If
            If GetCellValue(m_Row, gCommandNameColumn) = "" Or CStr(GetCellValue(m_Row, m_Column)) = "" Then
                Exit Sub
            End If

            Dim response As MsgBoxResult
            response = MyMsgBox("Update the column based on the first valid value?", MsgBoxStyle.YesNo)

            If response = MsgBoxResult.Yes Then
                If m_rowCount > 1 Then
                    Spreadsheet_UpdateColumn(m_Row, m_Column, m_rowCount, AxSpreadsheetProgramming)
                Else
                    MyMsgBox("Update column is not allowed", MsgBoxStyle.OKOnly)
                End If
            End If
            'when the header of row clicked
        ElseIf 1 = m_rowCount Then
            m_Execution.m_Pattern.Spreadsheet_GetRowLocalReference(AxSpreadsheetProgramming, m_Row, RefPos)
            Spreadsheet_SetMotionRef(RefPos)

            If 3 = m_columnCount Then
                If m_Execution.m_Pattern.m_ErrorChk.CheckValidPtXY(CmdStr) Then

                End If
            End If
        Else
            m_Execution.m_Pattern.Spreadsheet_GetRowLocalReference(AxSpreadsheetProgramming, m_Row, RefPos)
            Spreadsheet_SetMotionRef(RefPos)

            DisableEditingToolbar()
            DisableTeachingButtons()
        End If
        TraceGCCollect()
    End Sub


    'Check the select row is valid for Copy/cut or not                
    '   Axspreadsheet:  instance of activeX spreadsheet               
    '   m_LocalRow:     The row which will be checked                 
    '                                                                 

    Private Function Spreadsheet_WithinLinkRangeForCopyCut(ByVal m_LocalRow As Integer) As Boolean
        'If selected row is within Link range, part of Element tool icon should be disabled.
        'For fast speed, only +- 250 line will be checked, agreed by EET on 21/11/2007

        Dim Rtn As Boolean = False
        Dim FurtherChecking As Boolean = False


        Dim sel As Microsoft.Office.Interop.OWC.Range = AxSpreadsheetProgramming.Selection
        Dim m_rowCount As Integer = sel.Rows.Count()
        Dim m_columnCount As Integer = sel.Columns.Count()
        Dim m_LocalType As String

        Dim ii, jj As Integer

        If m_LocalRow < 251 Then
            ii = 1
        Else
            ii = m_LocalRow - 250
        End If

        Try
            m_LocalType = GetCellValue(m_LocalRow, gCommandNameColumn)

            If "Line" = m_LocalType Or "Arc" = m_LocalType Or "LinkStart" = m_LocalType Or "LinkEnd" = m_LocalType Then
                If "LinkEnd" = m_LocalType Or "LinkEnd" = m_LocalType Then
                    Rtn = True
                    FurtherChecking = False
                Else
                    FurtherChecking = True
                End If
            Else
                FurtherChecking = False
            End If

            If FurtherChecking Then
                'Backword search
                For jj = m_LocalRow To ii Step -1
                    m_LocalType = GetCellValue(jj, gCommandNameColumn)

                    If "LinkEnd" = m_LocalType Then
                        Exit Try
                    ElseIf "LinkStart" = m_LocalType Then
                        Rtn = True
                        Exit For
                    End If
                Next
                'Forword search
                For jj = m_LocalRow To m_LocalRow + 250
                    m_LocalType = GetCellValue(jj, gCommandNameColumn)
                    If "LinkStart" = m_LocalType Then
                        Exit Try
                    ElseIf "LinkEnd" = m_LocalType Then
                        Rtn = True
                        Exit For
                    End If
                Next
            End If

        Catch ex As SystemException
            ExceptionDisplay(ex)
        End Try

        Return (Rtn)
        TraceGCCollect()
    End Function


    'Check the select range is a link identity or not                 
    '   Axspreadsheet:  instance of activeX spreadsheet               
    '   m_LocalStartRow:  The start row of the range                  
    '   m_LocalEndRow:    The end row of the range                    
    '                                                                 
    '   Return Rtn=True if it is a identity                           

    Private Function Spreadsheet_WithinLinkRangeForCopyCut( _
            ByVal m_LocalStartRow As Integer, ByVal m_LocalEndRow As Integer) As Boolean
        Dim Rtn As Boolean = True
        Dim FurtherChecking As Boolean = False


        Dim sel As Microsoft.Office.Interop.OWC.Range = AxSpreadsheetProgramming.Selection
        Dim m_rowCount As Integer = sel.Rows.Count()
        Dim m_columnCount As Integer = sel.Columns.Count()
        Dim m_LocalType As String

        Dim ii, jj As Integer

        Try
            For jj = m_LocalStartRow To m_LocalEndRow
                m_LocalType = GetCellValue(jj, gCommandNameColumn)
                If "LinkStart" = m_LocalType Or "LinkEnd" = m_LocalType Then
                    If "LinkStart" = m_LocalType Then
                        If jj <> m_LocalStartRow Then
                            Rtn = False
                            Exit For
                        End If
                    Else
                        If jj <> m_LocalEndRow Then
                            Rtn = False
                            Exit For
                        End If
                    End If
                ElseIf "Line" <> m_LocalType And "Arc" <> m_LocalType Then 'Further enhancement
                    Rtn = False
                    Exit For
                End If
            Next

        Catch ex As SystemException
            ExceptionDisplay(ex)
        End Try

        Return (Rtn)
        TraceGCCollect()
    End Function



    'Check the current row is in the link range or not                                               
    '   UpdateToolBar:      To indicate the toolbar enabled/disabled will be updated or not
    '                       True=will be updated, False=not be updated

    Private Sub Spreadsheet_CheckForWithinLinkRange(ByVal UpdateToolBar As Boolean)
        'If selected row is within Link range, part of Element tool icon should be disabled.
        'For fast speed, only +- 250 line will be checked, agreed by EET on 21/11/2007

        Dim CurrentOffsetButtonStare As Boolean = ElementsCommandBlock.Buttons(gOffsetCmdIndex - 1).Enabled
        Dim FurtherChecking As Boolean = False

        Try
            Dim sel As Microsoft.Office.Interop.OWC.Range = AxSpreadsheetProgramming.Selection
            Dim m_rowCount As Integer = sel.Rows.Count()
            Dim m_columnCount As Integer = sel.Columns.Count()
            m_InLinkRange = False

            If m_rowCount <> 1 Then Exit Sub

            Dim m_LocalRow = sel.Row()
            Dim ii, jj As Integer

            If m_LocalRow < 251 Then
                ii = 1
            Else
                ii = m_LocalRow - 250
            End If

            Dim m_LocalType As String = GetCellValue(m_LocalRow, gCommandNameColumn)

            If "Line" = m_LocalType Or "Arc" = m_LocalType Or "LinkEnd" = m_LocalType Then
                If "LinkEnd" = m_LocalType Then
                    m_InLinkRange = True
                    FurtherChecking = False
                Else
                    FurtherChecking = True
                End If
            Else
                FurtherChecking = False
            End If

            If FurtherChecking Then
                'Backword search
                For jj = m_LocalRow To ii Step -1
                    m_LocalType = GetCellValue(jj, gCommandNameColumn)

                    If "LinkEnd" = m_LocalType Then
                        If UpdateToolBar Then
                            EnableTeachingButtons()
                        End If
                        Exit Sub
                    ElseIf "LinkStart" = m_LocalType Then
                        m_InLinkRange = True
                        Exit For
                    End If
                Next
                'Forword search
                For jj = m_LocalRow To m_LocalRow + 250
                    m_LocalType = GetCellValue(jj, gCommandNameColumn)
                    If "LinkStart" = m_LocalType Then
                        Exit Sub
                    ElseIf "LinkEnd" = m_LocalType Then
                        m_InLinkRange = True
                        Exit For
                    End If
                Next
            End If

            If UpdateToolBar Then
                If m_InLinkRange = True Then
                    DisableTeachingButtons()
                    EnableElementsCommandBlockButton(gLineCmdIndex)
                    EnableElementsCommandBlockButton(gArcCmdIndex)
                Else
                    EnableTeachingButtons()
                    If CurrentOffsetButtonStare = True Then
                        EnableElementsCommandBlockButton(gOffsetCmdIndex)
                    Else
                        DisableElementsCommandBlockButton(gOffsetCmdIndex)
                    End If
                End If
            End If
        Catch ex As SystemException
            ExceptionDisplay(ex)
        End Try
        TraceGCCollect()
    End Sub

    'Update one selected column in spreadsheet                                               
    '   AxSpreadsheet:      ActiveX spreadsheet                                                    
    '   startRow:           Starting rows
    '   m_rowCount:         Rows to be updated  
    '   SelectedColumn:     Selected column 

    Private Sub Spreadsheet_UpdateColumn(ByVal startRow As Integer, ByVal SelectedColumn As Integer, _
    ByVal m_rowCount As Integer, ByRef AxSpreadsheet As AxOWC10.AxSpreadsheet)
        Dim emptyLinCount As Integer = 0
        Dim i, j As Integer
        Dim StrTmp As String
        UndoData_Logging(0)

        i = startRow
        Do
            StrTmp = GetCellValue(i, gCommandNameColumn)
            If "" = StrTmp Or StrTmp.ToUpper = "ARRAY" Or StrTmp.ToUpper = "SUBPATTERN" Then
                emptyLinCount = emptyLinCount + 1
            Else

                m_Execution.m_ItemLUT.Cmd2Index(StrTmp.ToUpper)
                If m_Execution.m_ItemLUT.GetFlag(SelectedColumn) Then
                    SetCellValue(i, SelectedColumn, GetCellValue(startRow, SelectedColumn))
                End If

                emptyLinCount = 0
            End If
            i = i + 1
        Loop Until emptyLinCount > 20 Or i > m_rowCount + startRow - 1
        TraceGCCollect()
    End Sub


    'Delete multi-row                                                
    '   AxSpreadsheet:      ActiveX spreadsheet                                                                                                              

    Public Sub Spreadsheet_DeleteMultiRow(ByRef AxSpreadsheet As AxOWC10.AxSpreadsheet)
        Dim sel As Microsoft.Office.Interop.OWC.Range = AxSpreadsheetProgramming.Selection

        If AxSpreadsheetProgramming.ActiveSheet.Name <> "Main" Then
            If sel.Count >= AxSpreadsheetProgramming.ActiveSheet.UsedRange.Count Then
                MessageBox.Show("Cannot clear all items in other sheet except in main sheet!")
                Return
            End If
        End If
        Dim m_rowCount As Integer = sel.Rows.Count()
        Dim m_rowLocal As Integer = sel.Row

        Dim DeltedRowNo As Integer = 0
        Dim DeletedRowNoEachtime As Integer = 1

        If m_rowCount < 1 Then Return
        Dim name As String = ""
        Do
            name = GetCellValue(m_Row, gCommandNameColumn)
            If name.ToUpper = "GLOBALQC" Then
                IDS.Data.Hardware.Camera.DotQCEnable = False
            End If
            Spreadsheet_DeleteOneRow(DeletedRowNoEachtime, m_rowLocal, AxSpreadsheet)

            DeltedRowNo = DeltedRowNo + DeletedRowNoEachtime
        Loop Until (DeltedRowNo >= m_rowCount)
        'Dim dir As Object = Microsoft.Office.Interop.OWC.XlDeleteShiftDirection.xlShiftUp
        'sel.Delete()
    End Sub


    Public Sub Spreadsheet_DeleteMultiRow2(ByRef AxSpreadsheet As AxOWC10.AxSpreadsheet)
        '''''''''''''''''''''''''''''''''''''''''''''''''
        '   Xue Wen                                     '
        '   Method (2) for deleting "Multiple Rows".    '
        '   Note: Test more.                            '
        '''''''''''''''''''''''''''''''''''''''''''''''''
        Dim sel As Microsoft.Office.Interop.OWC.Range = AxSpreadsheetProgramming.Selection
        Dim m_rowCount As Integer = sel.Rows.Count()
        Dim m_rowLocal As Integer = sel.Row

        Dim UsedRowNumber1 As Integer = AxSpreadsheet.ActiveSheet.UsedRange.Rows.Count
        Dim UsedRowNumber2 As Integer = 1
        Dim UsedRowNumberTmp As Integer = 1

        Dim DeltedRowNo, Del_i As Integer

        If m_rowCount < 1 Then Return

        For Del_i = 1 To UsedRowNumber1
            If (UsedRowNumber2 <= 0) Or (Del_i > UsedRowNumber1) Then
                Exit For
            End If

            If (m_rowLocal = Del_i) Then
                m_rowCount = sel.Rows.Count()

                Spreadsheet_DeleteOneRow(UsedRowNumberTmp, Del_i, AxSpreadsheet)

                UsedRowNumber2 = AxSpreadsheet.ActiveSheet.UsedRange.Rows.Count
                DeltedRowNo = UsedRowNumber1 - UsedRowNumber2

                If (m_rowCount = DeltedRowNo) Then
                    Exit For
                End If

                If (m_rowCount > DeltedRowNo) Then
                    m_rowLocal = sel.Row
                ElseIf (m_rowCount < DeltedRowNo) Then
                    If (m_rowLocal = m_Row) Then
                        Dim E_Row As Integer

                        E_Row = m_rowLocal + (m_rowCount - 1)
                        If (m_rowLocal + (DeltedRowNo - 1)) > E_Row Then
                            m_rowLocal = 0
                        End If
                    Else
                        m_rowCount = DeltedRowNo + (m_rowCount - 1)

                        If (m_rowCount > DeltedRowNo) Then
                            m_rowLocal = sel.Row
                        Else
                            m_rowLocal = 0
                        End If
                    End If
                Else
                    m_rowLocal = 0
                End If

                Del_i = m_Row - 1
                UsedRowNumber1 = UsedRowNumber2
            End If
        Next

        m_Row = UsedRowNumber2 + 1
    End Sub

    'Delete one row                                                
    '   AxSpreadsheet:      ActiveX spreadsheet                                                   
    '   row:                Row to be deleted                         
    Public Sub Spreadsheet_DeleteOneRow(ByRef NumOfRowDeleted As Integer, ByVal row As Integer, ByRef AxSpreadsheet As AxOWC10.AxSpreadsheet)

        '''''''''''''''''''''''''''''''''''''
        '   Xue Wen                         '
        '   Here, should not use "m_row".   '
        '''''''''''''''''''''''''''''''''''''
        NumOfRowDeleted = 1

        m_Row = row
        Dim m_type As String = GetCellValue(m_Row, gCommandNameColumn)

        If "" = m_type Then
            DeleteRow(m_Row)

        ElseIf "SUBPATTERN" = m_type.ToUpper Or "Array" = m_type Then
            Dim SubSheetName As String
            Dim CurrentSheetName As String

            SubSheetName = GetCellValue(m_Row, gSubnameColumn)
            CurrentSheetName = GetActiveSheetName()

            'If it is also used by others, the sub with other subs attached to it will not be deleted.
            'If 0 = Spreadsheet_QuickCountExistedSubsheet(AxSpreadsheetProgramming, SubSheetName) Then
            'The current sub sheet will be removed.  Before delete the sub sheet, we also need to 
            'check if we should to delete the sub->sub sheets and arrays sheets attached to it.

            'Find all the subsheets, then classify the usages

            If "SUBPATTERN" = m_type.ToUpper Then
                SubSheetName = m_Execution.m_File.NameOnlyFromFullPath(SubSheetName)
            Else
                If CurrentSheetName = "Main" Then

                Else
                    SubSheetName = CurrentSheetName + "." + SubSheetName
                End If
            End If

            Spreadsheet_FindSubTotalUsage(AxSpreadsheetProgramming, SubSheetName)
            Spreadsheet_BuildRootConnectedSub(AxSpreadsheetProgramming, SubSheetName)
            Spreadsheet_FindSubSelfUsage(AxSpreadsheetProgramming, SubSheetName)
            Spreadsheet_FlatternTotalUsage(AxSpreadsheetProgramming, SubSheetName)
            Spreadsheet_DeleteSubSheet(AxSpreadsheetProgramming)

            DeleteRow(m_Row)
        ElseIf "Line" = m_type Or "Arc" = m_type Then
            Spreadsheet_CheckForWithinLinkRange(False)
            If m_InLinkRange Then
                With AxSpreadsheet.ActiveSheet
                    If GetCellValue(m_Row - 1, gCommandNameColumn) = "Link" Then
                    ElseIf GetCellValue(m_Row - 1, gCommandNameColumn) = "Line" Then
                        If GetCellValue(m_Row + 1, gCommandNameColumn) = "Line" Or _
                        GetCellValue(m_Row + 1, gCommandNameColumn) = "Arc" Then
                            .Cells(m_Row + 1, gPos1XColumn) = .Cells(m_Row - 1, gPos2XColumn)
                            .Cells(m_Row + 1, gPos1YColumn) = .Cells(m_Row - 1, gPos2YColumn)
                            .Cells(m_Row + 1, gPos1ZColumn) = .Cells(m_Row - 1, gPos2ZColumn)
                        End If
                    ElseIf GetCellValue(m_Row - 1, gCommandNameColumn) = "Arc" Then
                        If GetCellValue(m_Row + 1, gCommandNameColumn) = "Line" Or _
                            GetCellValue(m_Row + 1, gCommandNameColumn) = "Arc" Then
                            .Cells(m_Row + 1, gPos1XColumn) = .Cells(m_Row - 1, gPos3XColumn)
                            .Cells(m_Row + 1, gPos1YColumn) = .Cells(m_Row - 1, gPos3YColumn)
                            .Cells(m_Row + 1, gPos1ZColumn) = .Cells(m_Row - 1, gPos3ZColumn)
                        End If
                    End If
                End With
            End If
            DeleteRow(m_Row)
        ElseIf "LINKSTART" = m_type.ToUpper Or "LINKEND" = m_type.ToUpper Then
            Dim ii As Integer = 0
            If "LINKSTART" = m_type.ToUpper Then
                Do
                    type = GetCellValue(m_Row, gCommandNameColumn)
                    DeleteRow(m_Row)
                    ii = ii + 1
                    If Nothing = type Then
                        type = "TMPSTRING"
                    End If

                    '   Change "m_type.ToUpper" to "type.ToUpper" becuase "type" is updated cell value.  '
                    '   If not, system will do deleting 500 times.                                      '
                Loop Until "LINKEND" = type.ToUpper Or ii > 500
                SelectCell(m_Row, 1)

                NumOfRowDeleted = ii

            ElseIf "LINKEND" = m_type.ToUpper Then
                Do
                    If m_Row - ii < 1 Then Exit Sub

                    type = GetCellValue(m_Row - ii, gCommandNameColumn)
                    DeleteRow(m_Row - ii)
                    ii = ii + 1

                    'Change "m_type.ToUpper" to "type.ToUpper" becuase "type" is updated cell value.  
                    If Nothing = type Then
                        type = "TMPSTRING"
                    End If
                Loop Until "LINKSTART" = type.ToUpper Or ii > 500

                'Updating will also do in end of delete. So, we just need to set row position.   '
                SelectCell(m_Row - ii + 1, 1)
                NumOfRowDeleted = ii

            End If
        Else
            DeleteRow(m_Row)
        End If
        TraceGCCollect()
    End Sub


    'Shen Jian's part
    Private Sub Spreadsheet_BeforeContextMenu(ByVal sender As System.Object, ByVal e As AxOWC10.ISpreadsheetEventSink_BeforeContextMenuEvent) Handles AxSpreadsheetProgramming.BeforeContextMenu

        Dim oMenu1, oMenu2, oMenu3, oMenu3_1, oMenu4, oMenu5, oMenu6, oMenu7 As Object()
        Dim oMenu As Object()
        If AxSpreadsheetProgramming.ActiveCell.Column = gNeedleColumn Then
            oMenu = New Object() {}
            'If (chkDualHead.Checked = False) Then
            '    oMenu1 = New Object() {"&Left", "Left"}
            '    oMenu = New Object() {oMenu1}
            'Else
            '    oMenu1 = New Object() {"&Left", "Left"}
            '    oMenu2 = New Object() {"&Right", "Right"}
            '    oMenu = New Object() {oMenu1, oMenu2}
            'End If

        ElseIf AxSpreadsheetProgramming.ActiveCell.Column = gDispensColumn Then
            Dim row As Integer = GetActiveCellRow()
            Dim command As String = GetCellValue(row, gCommandNameColumn)
            Dim CmdName() As String = {"FIDUCIAL", "HEIGHT", "REFERENCE", "REJECT", "LINKSTART", "LINKEND", "QC", "SUBPATTERN", "ARRAY", "SETIO", "GETIO"}
            If command.ToUpper = "FIDUCIAL" Or command.ToUpper = "HEIGHT" Or command.ToUpper = "LINKSTART" Or command.ToUpper = "LINKEND" Or command.ToUpper = "REFERENCE" Or command.ToUpper = "GETIO" Or command.ToUpper = "PURGE" Or command.ToUpper = "CLEAN" Or command.ToUpper = "ARRAY" Or command.ToUpper = "SUBPATTERN" Or command.ToUpper = "QC" Then
                oMenu = New Object() {}
            Else
                oMenu1 = New Object() {"&On", "On"}
                oMenu2 = New Object() {"&Off", "Off"}
                oMenu = New Object() {oMenu1, oMenu2}
            End If

            'oMenu3 = New Object() {"____________", ""}
            'oMenu4 = New Object() {"&Undo", "Undo"}
            'oMenu5 = New Object() {"&Insert Row", "Insert"}
            'oMenu6 = New Object() {"&Delete Row", "Delete"}
            'oMenu = New Object() {oMenu1, oMenu2, oMenu3, oMenu4, oMenu5, oMenu6}

        Else
            oMenu = New Object() {}

            'ElseIf AxSpreadsheetProgramming.ActiveCell.Column = gCommandNameColumn Then
            '    'lsgoh - start 
            '    oMenu1 = New Object() {"RowsCut", "MCut"}
            '    oMenu2 = New Object() {"RowsCopy", "MCopy"}
            '    oMenu3 = New Object() {"RowsPaste", "MPaste"}
            '    oMenu4 = New Object() {"RowsInsert", "MInsert"}
            '    oMenu5 = New Object() {"____________", ""}
            '    'lsgoh end

            '    oMenu6 = New Object() {"&Undo", "Undo"}
            '    oMenu7 = New Object() {"&Insert Row", "Insert"}
            '    oMenu3_1 = New Object() {"&Delete Row", "Delete"}
            '    'oMenu = New Object() {oMenu1, oMenu2, oMenu3}
            '    oMenu = New Object() {oMenu1, oMenu2, oMenu3, oMenu4, oMenu5, oMenu6, oMenu7, oMenu3_1} 'lsgoh
            'Else
            '    oMenu1 = New Object() {"&Cut", "Cut"}
            '    oMenu2 = New Object() {"&Copy", "Copy"}
            '    oMenu3 = New Object() {"&Paste", "Paste"}
            '    oMenu4 = New Object() {"____________", ""}
            '    oMenu5 = New Object() {"&Undo", "Undo"}
            '    oMenu6 = New Object() {"&Insert Row", "Insert"}
            '    oMenu7 = New Object() {"&Delete Row", "Delete"}

            '    oMenu = New Object() {oMenu1, oMenu2, oMenu3, oMenu4, oMenu5, oMenu6, oMenu7}
        End If
        e.menu.Value = oMenu
        'e.menu.Value = Me.NeedleContextMenu

    End Sub

    Private Function CheckNeedleChangeFlag(ByVal type As String) As Boolean
        Dim CmdSize As Integer = 10
        Dim CmdName() As String = {"FIDUCIAL", "HEIGHT", "REFERENCE", "REJECT", "LINKSTART", "LINKEND", "QC", "SUBPATTERN", "ARRAY", "SETIO", _
                                    "GETIO"}
        Dim cmdtype As Integer
        Dim cmdStr As String = type.ToUpper
        cmdtype = -1
        For i = 0 To CmdSize
            If CmdName(i) <> "" Then
                If cmdStr.IndexOf(CmdName(i)) = 0 Then
                    cmdtype = i
                    Exit For
                End If
            End If
        Next i
        If cmdtype >= 0 Then
            Return False
        Else
            Return True
        End If
    End Function

    Private Function CheckDispenseONOFFFlag(ByVal type As String) As Boolean
        Dim CmdSize As Integer = 13
        Dim CmdName() As String = {"FIDUCIAL", "HEIGHT", "REFERENCE", "REJECT", "LINKSTART", "LINKEND", "QC", "SUBPATTERN", "ARRAY", _
                                    "GETIO", "MOVE", "WAIT", "PURGE", "CLEAN", "VOLUMECALIBRATION"}
        Dim cmdtype As Integer
        Dim cmdStr As String = type.ToUpper
        cmdtype = -1
        For i = 0 To CmdSize
            If CmdName(i) <> "" Then
                If cmdStr.IndexOf(CmdName(i)) = 0 Then
                    cmdtype = i
                    Exit For
                End If
            End If
        Next i
        If cmdtype >= 0 Then
            Return False
        Else
            Return True
        End If
    End Function

    Private Sub Spreadsheet_CommandExecute(ByVal sender As System.Object, ByVal e As AxOWC10.ISpreadsheetEventSink_CommandExecuteEvent) Handles AxSpreadsheetProgramming.CommandExecute
        Dim row As Integer = GetActiveCellRow()
        Dim colum As Integer = GetActiveCellColumn()
        Dim cmdType As String = GetCellValue(row, colum)
        Static m_rowCount As Integer

        UndoData_Logging(1)

        Select Case e.command
            Case "Left", "Right"
                If CheckNeedleChangeFlag(cmdType) Then
                    SetCellValue(row, colum, e.command)
                Else
                    LabelMessage("Can't insert needle for this command!")
                    LabelMessege.Update()
                End If

            Case "On"
                If CheckDispenseONOFFFlag(cmdType) Then
                    SetCellValue(row, colum, "On")
                Else
                    LabelMessage("Can't insert ON/OFF for this command!")
                    LabelMessege.Update()
                End If

            Case "Off"
                If CheckDispenseONOFFFlag(cmdType) Then
                    SetCellValue(row, colum, "Off")
                Else
                    LabelMessage("Can't insert ON/OFF for this command!")
                    LabelMessege.Update()
                End If

            Case "Cut"
                AxSpreadsheetProgramming.ActiveCell.Cut()
            Case "Copy"
                AxSpreadsheetProgramming.ActiveCell.Copy()
            Case "Paste"
                AxSpreadsheetProgramming.ActiveCell.Paste()

                'lsgoh - start here
            Case "MCut"
                Dim sel As Microsoft.Office.Interop.OWC.Range = AxSpreadsheetProgramming.Selection 'lsgoh
                m_rowCount = sel.Rows.Count()
                AxSpreadsheetProgramming.Selection.Cut()
            Case "MCopy"
                Dim sel As Microsoft.Office.Interop.OWC.Range = AxSpreadsheetProgramming.Selection 'lsgoh
                m_rowCount = sel.Rows.Count()
                AxSpreadsheetProgramming.Selection.Copy()
            Case "MPaste"
                SelectCell(row, 1)
                AxSpreadsheetProgramming.Selection.Paste()
                SelectCell(row, 1)

            Case "MInsert"
                SelectCell(row, 1)
                Dim I As Integer = 0
                If m_rowCount > 0 Then
                    For I = 0 To m_rowCount - 1
                        AxSpreadsheetProgramming.Selection.EntireRow.Insert()
                    Next
                End If
                AxSpreadsheetProgramming.Selection.Paste()
                SelectCell(row, 1)
                'end here

            Case "Undo"
                AxSpreadsheetProgramming.Undo()
            Case "Insert"
                SelectCell(row + 1, 1)
                AxSpreadsheetProgramming.Selection.EntireRow.Insert()
                ClearRow(row + 1)
                If (m_Row > row) Then
                    m_Row = m_Row + 1
                End If

            Case "Delete"
                SelectCell(row, 1)
                AxSpreadsheetProgramming.Selection.EntireRow.Delete()
                If (m_Row > row) Then
                    m_Row = m_Row - 1
                ElseIf m_Row = row Then
                    DisableCoordinateUpdateInSpreadsheet()
                End If
                If (m_Row < 2) Then
                    m_Row = 2
                    DisableCoordinateUpdateInSpreadsheet()
                End If
        End Select


        TraceGCCollect()
    End Sub





    'Automatically Generate Array or SubSheet Name
    '   SheetName: Sheet name to be generated 
    '   ActSheetName: Current actSheet name       

    Private Sub Spreadsheet_GenerateArraySubSheetName(ByRef SheetName, ByVal ActSheetName, ByVal MyType)

        Dim ArraySheetSerialNumber As Integer = 1

        'kr added to shorten sheet names
        If MyType = "Rectangle" Then MyType = "Rect"
        If MyType = "Circle" Then MyType = "Circ"
        If MyType = "FillRectangle" Then MyType = "FillRect"
        If MyType = "FillCircle" Then MyType = "FillCirc"

        If "Main" = ActSheetName Then
            Do
                SheetName = MyType + "_" + CStr(ArraySheetSerialNumber)
                If 0 = m_Execution.m_Pattern.Spreadsheet_CheckSubsheetExist(AxSpreadsheetProgramming, SheetName) Then
                    Exit Do
                End If
                ArraySheetSerialNumber = ArraySheetSerialNumber + 1
            Loop
        Else
            Do
                SheetName = ActSheetName + "." + MyType + "_" + CStr(ArraySheetSerialNumber)
                If 0 = m_Execution.m_Pattern.Spreadsheet_CheckSubsheetExist(AxSpreadsheetProgramming, SheetName) Then
                    Exit Do
                End If
                ArraySheetSerialNumber = ArraySheetSerialNumber + 1
            Loop
        End If

        TraceGCCollect()
    End Sub



    'Check the sub or array existance changes, not be used currently
    '   SheetName: Sheet to be checked
    '   
    '   Return: The changes of existance

    Private Function Spreadsheet_QuickCountExistedSubsheet(ByRef AxSpreadsheet As AxOWC10.AxSpreadsheet, _
        ByRef SheetName As String) As Integer

        Dim Rtn As Integer = 0
        Dim i, j, emptyLine As Integer
        Dim NumberOfSheet As Integer = AxSpreadsheet.Worksheets.Count()
        Dim SpreadSheetName As String
        Dim strTmp As String

        For i = 1 To NumberOfSheet
            SpreadSheetName = AxSpreadsheet.ActiveWorkbook.Worksheets(i).name()
            j = 0
            emptyLine = 0
            With AxSpreadsheet.Worksheets(SpreadSheetName)
                Do
                    j = j + 1
                    strTmp = .Cells(j, gCommandNameColumn).Value

                    If "" = strTmp Then
                        emptyLine = emptyLine + 1
                    ElseIf "Array" = strTmp Then
                        emptyLine = 0
                        If SheetName = .Cells(j, gSubnameColumn).Value Then
                            Rtn = Rtn + 1
                            Exit For
                        End If
                    ElseIf "SubPattern" = strTmp Then
                        emptyLine = 0
                        If SheetName = .Cells(j, gSubnameColumn).Value Then
                            Rtn = Rtn + 1
                            If Rtn > 1 Then 'if more than one were found
                                Exit For
                            End If
                        End If
                    Else
                        emptyLine = 0
                    End If
                Loop Until 20 < emptyLine 'Successive 20 empty excel lines will terminate the saving record.
            End With
        Next i

        Return Rtn
    End Function


    'Total usage of all sub/Array Spreadsheet                                                
    '   AxSpreadsheet:     ActiveX spreadsheet                                                    
    '   RootDelectSheetName: Root sheet name, not in use                                                             

    Private Sub Spreadsheet_FindSubTotalUsage(ByRef AxSpreadsheet As AxOWC10.AxSpreadsheet, _
        ByRef RootDelectSheetName As String)
        Dim Rtn As Integer = 0
        Dim i, j, k, emptyLine As Integer
        Dim NumberOfSheets As Integer = AxSpreadsheet.Worksheets.Count()
        Dim CurrentSheetName As String
        Dim strTmp As String
        Dim SheetName As String



        m_Execution.m_Pattern.SubCallSheetName.TotalNumberOfSheets = NumberOfSheets
        For i = 0 To NumberOfSheets - 1
            m_Execution.m_Pattern.SubCallSheetName.SheetName(i).SheetName = _
                AxSpreadsheet.ActiveWorkbook.Worksheets(i + 1).name()
            m_Execution.m_Pattern.SubCallSheetName.SheetName(i).SheetInSelfUse = 0
            m_Execution.m_Pattern.SubCallSheetName.SheetName(i).SheetInTotalUse = 0
            m_Execution.m_Pattern.SubCallSheetName.SheetName(i).SheetRootRelated = 0
        Next i

        For i = 0 To NumberOfSheets - 1
            CurrentSheetName = m_Execution.m_Pattern.SubCallSheetName.SheetName(i).SheetName

            j = 0
            emptyLine = 0
            With AxSpreadsheet.Worksheets(CurrentSheetName)
                Do
                    j = j + 1
                    strTmp = .Cells(j, gCommandNameColumn).Value

                    If "" = strTmp Then
                        emptyLine = emptyLine + 1
                    ElseIf "Array" = strTmp Or "SubPattern" = strTmp Then
                        emptyLine = 0
                        SheetName = .Cells(j, gSubnameColumn).Value

                        If "SubPattern" = strTmp Then
                            SheetName = m_Execution.m_File.NameOnlyFromFullPath(SheetName)
                        Else
                            If CurrentSheetName = "Main" Then

                            Else
                                SheetName = CurrentSheetName + "." + SheetName
                            End If
                        End If

                        For k = 0 To NumberOfSheets - 1
                            If SheetName = m_Execution.m_Pattern.SubCallSheetName.SheetName(k).SheetName Then
                                m_Execution.m_Pattern.SubCallSheetName.SheetName(k).SheetInTotalUse += 1
                                Exit For
                            End If
                        Next k
                    Else
                        emptyLine = 0
                    End If
                Loop Until 20 < emptyLine 'Successive 20 empty excel lines will terminate the saving record.
            End With
        Next i
        TraceGCCollect()
    End Sub



    'Root connection of a sheet, itself is level 1                                              
    '   AxSpreadsheet:     ActiveX spreadsheet                                                    
    '   RootSheetName: Root sheet name                                                              

    Private Sub Spreadsheet_BuildRootConnectedSub(ByRef AxSpreadsheet As AxOWC10.AxSpreadsheet, _
        ByRef RootSheetName As String)
        Dim Rtn As Integer = 0
        Dim i, j, k, emptyLine As Integer
        Dim NumberOfSheet As Integer = AxSpreadsheet.Worksheets.Count()
        Dim SpreadSheetName As String
        Dim CurrentSheetName As String
        Dim strTmp As String
        Dim SheetName As String

        NumberOfSheet = m_Execution.m_Pattern.SubCallSheetName.TotalNumberOfSheets

        'Build the connectivity of the first sub sheet.  It's also in self use.
        For k = 0 To NumberOfSheet - 1
            If RootSheetName = m_Execution.m_Pattern.SubCallSheetName.SheetName(k).SheetName Then
                m_Execution.m_Pattern.SubCallSheetName.SheetName(k).SheetRootRelated = 1
                m_Execution.m_Pattern.SubCallSheetName.SheetName(k).SheetInSelfUse = 1
                Exit For
            End If
        Next k



        'Build the connectivity of the remaining sub sheets
        Dim iLayer As Integer = 0
        Dim iConnected As Integer = 0

        'Loop for all the layers
        Do
            iLayer += 1
            iConnected = 0
            'Loop for all the sheets
            For i = 0 To NumberOfSheet - 1

                If iLayer = m_Execution.m_Pattern.SubCallSheetName.SheetName(i).SheetRootRelated Then
                    CurrentSheetName = m_Execution.m_Pattern.SubCallSheetName.SheetName(i).SheetName

                    j = 0
                    emptyLine = 0
                    With AxSpreadsheet.Worksheets(CurrentSheetName)
                        'Loop within the sheet
                        Do
                            j = j + 1
                            strTmp = .Cells(j, gCommandNameColumn).Value

                            If "" = strTmp Then
                                emptyLine = emptyLine + 1
                            ElseIf "Array" = strTmp Or "SubPattern" = strTmp Then
                                emptyLine = 0

                                SheetName = .Cells(j, gSubnameColumn).Value

                                If "SubPattern" = strTmp Then
                                    SheetName = m_Execution.m_File.NameOnlyFromFullPath(SheetName)
                                Else
                                    If CurrentSheetName = "Main" Then

                                    Else
                                        SheetName = CurrentSheetName + "." + SheetName
                                    End If
                                End If


                                For k = 0 To NumberOfSheet - 1
                                    'Find only the unconnected sheet.  Mark the layer
                                    If SheetName = m_Execution.m_Pattern.SubCallSheetName.SheetName(k).SheetName And _
                                    0 = m_Execution.m_Pattern.SubCallSheetName.SheetName(k).SheetRootRelated Then
                                        m_Execution.m_Pattern.SubCallSheetName.SheetName(k).SheetRootRelated = iLayer + 1
                                        iConnected = 1
                                        Exit For
                                    End If
                                Next k
                            Else
                                emptyLine = 0
                            End If
                        Loop Until 20 < emptyLine 'Successive 20 empty excel lines will terminate the saving record.
                    End With
                End If
            Next i
        Loop Until iConnected = 0
        TraceGCCollect()
    End Sub



    'Self usage of a sub/Array Spreadsheet                                              
    '   AxSpreadsheet:     ActiveX spreadsheet                                                    
    '   RootDelectSheetName: Root sheet name                                                              

    Private Sub Spreadsheet_FindSubSelfUsage(ByRef AxSpreadsheet As AxOWC10.AxSpreadsheet, _
        ByRef RootDelectSheetName As String)
        Dim Rtn As Integer = 0
        Dim i, j, k, emptyLine As Integer
        Dim NumberOfSheet As Integer = AxSpreadsheet.Worksheets.Count()
        Dim CurrentSheetName As String
        Dim strTmp As String
        Dim SheetName As String

        NumberOfSheet = m_Execution.m_Pattern.SubCallSheetName.TotalNumberOfSheets

        For i = 0 To NumberOfSheet - 1

            If 0 < m_Execution.m_Pattern.SubCallSheetName.SheetName(i).SheetRootRelated Then
                CurrentSheetName = m_Execution.m_Pattern.SubCallSheetName.SheetName(i).SheetName

                j = 0
                emptyLine = 0
                With AxSpreadsheet.Worksheets(CurrentSheetName)
                    Do
                        j = j + 1
                        strTmp = .Cells(j, gCommandNameColumn).Value

                        If "" = strTmp Then
                            emptyLine = emptyLine + 1
                        ElseIf "Array" = strTmp Or "SubPattern" = strTmp Then
                            emptyLine = 0
                            SheetName = .Cells(j, gSubnameColumn).Value


                            If "SubPattern" = strTmp Then
                                SheetName = m_Execution.m_File.NameOnlyFromFullPath(SheetName)
                            Else
                                If CurrentSheetName = "Main" Then

                                Else
                                    SheetName = CurrentSheetName + "." + SheetName
                                End If
                            End If

                            For k = 0 To NumberOfSheet - 1
                                If SheetName = m_Execution.m_Pattern.SubCallSheetName.SheetName(k).SheetName Then
                                    m_Execution.m_Pattern.SubCallSheetName.SheetName(k).SheetInSelfUse += 1
                                    Exit For
                                End If
                            Next k
                        Else
                            emptyLine = 0
                        End If
                    Loop Until 20 < emptyLine 'Successive 20 empty excel lines will terminate the saving record.
                End With
            End If
        Next i
        TraceGCCollect()
    End Sub



    'Self usage of a sub/Array Spreadsheet to include external as as an add on                                               
    '   AxSpreadsheet:     ActiveX spreadsheet                                                    
    '   RootDelectSheetName: Root sheet name                                                              

    Private Sub Spreadsheet_FlatternTotalUsage(ByRef AxSpreadsheet As AxOWC10.AxSpreadsheet, _
        ByRef RootDelectSheetName As String)
        Dim Rtn As Integer = 0
        Dim i, j, k, emptyLine As Integer
        Dim NumberOfSheets As Integer = AxSpreadsheet.Worksheets.Count()
        Dim CurrentSheetName As String
        Dim strTmp As String
        Dim SheetName As String
        Dim RootSheetInTotalUse, RootSheetInSelfUse As Integer


        For k = 0 To NumberOfSheets - 1
            If RootDelectSheetName = m_Execution.m_Pattern.SubCallSheetName.SheetName(k).SheetName Then
                RootSheetInTotalUse = m_Execution.m_Pattern.SubCallSheetName.SheetName(k).SheetInTotalUse
                RootSheetInSelfUse = m_Execution.m_Pattern.SubCallSheetName.SheetName(k).SheetInSelfUse
                Exit For
            End If
        Next k


        For k = 0 To NumberOfSheets - 1
            'Exclude the sub/Array itself
            If 1 < m_Execution.m_Pattern.SubCallSheetName.SheetName(k).SheetRootRelated Then
                'This is not a real number.  But it is Ok to use it to indicate the difference
                m_Execution.m_Pattern.SubCallSheetName.SheetName(k).SheetInTotalUse += _
                    (RootSheetInTotalUse - RootSheetInSelfUse)

                Exit For
            End If
        Next k

        TraceGCCollect()
    End Sub


    'Delete one page of Spreadsheet                                                
    '   AxSpreadsheet:     ActiveX spreadsheet                                                    
    '                                                                 

    Private Sub Spreadsheet_DeleteSubSheet(ByRef AxSpreadsheet As AxOWC10.AxSpreadsheet)
        Dim i As Integer
        Dim NumberOfSheet As Integer
        Dim SpreadSheetName As String

        NumberOfSheet = m_Execution.m_Pattern.SubCallSheetName.TotalNumberOfSheets

        For i = 0 To NumberOfSheet - 1

            If 0 < m_Execution.m_Pattern.SubCallSheetName.SheetName(i).SheetRootRelated Then
                'Only delete the subsheet which has different total and self use number
                If m_Execution.m_Pattern.SubCallSheetName.SheetName(i).SheetInTotalUse = m_Execution.m_Pattern.SubCallSheetName.SheetName(i).SheetInSelfUse Then
                    SpreadSheetName = m_Execution.m_Pattern.SubCallSheetName.SheetName(i).SheetName
                    AxSpreadsheetProgramming.Worksheets(SpreadSheetName).Delete()
                End If
            End If
        Next i
        TraceGCCollect()
    End Sub




    'Build Array data based on the 3 rows input                                                
    '   AxSpreadsheet:      ActiveX spreadsheet                                                    
    '   Record1:            The first row for the array   
    '   Record2:            The second row for the array  
    '   Record3:            The third row for the array  
    '   arraydata:          Array parameters 

    Private Sub Spreadsheet_AddSheetRecord( _
        ByRef Record1 As CIDSPattern.PatternRecord, _
        ByRef Record2 As CIDSPattern.PatternRecord, _
        ByRef Record3 As CIDSPattern.PatternRecord, _
        ByVal arraydata As ArrayData)
        Dim i, j, ij As Integer


        Dim dXh, dXv, dYh, dYv, Xij, Yij As Double
        Dim pt1(3), pt2(3), pt3(3), cen1(2), cen2(2), cen3(2), radius As Double
        Dim Record As CIDSPattern.PatternRecord

        'EXPERIMENTAL 11/23/15
        'If "Circle" = Record1.pPara.Name Or "FillRectangle" = Record1.pPara.Name Then
        '    pt1(0) = Record1.pPara.Pos1.X : pt1(1) = Record1.pPara.Pos1.Y : pt1(2) = Record1.pPara.Pos1.Z
        '    pt2(0) = Record1.pPara.Pos2.X : pt2(1) = Record1.pPara.Pos2.Y : pt2(2) = Record1.pPara.Pos2.Z
        '    pt3(0) = Record1.pPara.Pos3.X : pt3(1) = Record1.pPara.Pos3.Y : pt3(2) = Record1.pPara.Pos3.Z
        '    MathFunction.Get3dCircleCentre(pt1, pt2, pt3, cen1, radius)

        '    pt1(0) = Record2.pPara.Pos1.X : pt1(1) = Record2.pPara.Pos1.Y : pt1(2) = Record2.pPara.Pos1.Z
        '    pt2(0) = Record2.pPara.Pos2.X : pt2(1) = Record2.pPara.Pos2.Y : pt2(2) = Record2.pPara.Pos2.Z
        '    pt3(0) = Record2.pPara.Pos3.X : pt3(1) = Record2.pPara.Pos3.Y : pt3(2) = Record2.pPara.Pos3.Z
        '    MathFunction.Get3dCircleCentre(pt1, pt2, pt3, cen2, radius)

        '    pt1(0) = Record3.pPara.Pos1.X : pt1(1) = Record3.pPara.Pos1.Y : pt1(2) = Record3.pPara.Pos1.Z
        '    pt2(0) = Record3.pPara.Pos2.X : pt2(1) = Record3.pPara.Pos2.Y : pt2(2) = Record3.pPara.Pos2.Z
        '    pt3(0) = Record3.pPara.Pos3.X : pt3(1) = Record3.pPara.Pos3.Y : pt3(2) = Record3.pPara.Pos3.Z
        '    MathFunction.Get3dCircleCentre(pt1, pt2, pt3, cen3, radius)

        '    dXh = (cen2(0) - cen1(0)) / (arraydata.ColumNo - 1)
        '    dYh = (cen2(1) - cen1(1)) / (arraydata.ColumNo - 1)
        '    dXv = (cen3(0) - cen1(0)) / (arraydata.RowNo - 1)
        '    dYv = (cen3(1) - cen1(1)) / (arraydata.RowNo - 1)
        'Else
        dXh = (Record2.pPara.Pos1.X - Record1.pPara.Pos1.X) / (arraydata.ColumNo - 1)
        dYh = (Record2.pPara.Pos1.Y - Record1.pPara.Pos1.Y) / (arraydata.ColumNo - 1)
        dXv = (Record3.pPara.Pos1.X - Record2.pPara.Pos1.X) / (arraydata.RowNo - 1)
        dYv = (Record3.pPara.Pos1.Y - Record2.pPara.Pos1.Y) / (arraydata.RowNo - 1)
        'End If

        Dim iFlagOptimize As Boolean = True
        Dim iInc As Integer
        ij = -arraydata.ColumNo + 1         'Set ij=0 for unoptimized

        For i = 0 To arraydata.RowNo - 1
            'Optimization array coordinate.  If unoptimized, set iInc=1
            If False = iFlagOptimize Then
                iInc = -1
                iFlagOptimize = True
                ij = ij + arraydata.ColumNo + 1
            Else
                iInc = 1
                iFlagOptimize = False
                ij = ij + arraydata.ColumNo - 1
            End If
            'End of array coordinate optimization

            For j = 0 To arraydata.ColumNo - 1
                ij = ij + iInc
                Xij = dXh * j + dXv * i
                Yij = dYh * j + dYv * i

                SetCellValue(ij, gCommandNameColumn, Record1.pPara.Name)
                m_Execution.m_ItemLUT.Cmd2Index((Record1.pPara.Name).ToUpper)

                If m_Execution.m_ItemLUT.GetFlag(gNeedleColumn) Then
                    SetCellValue(ij, gNeedleColumn, Record1.pPara.Needle)
                End If
                If m_Execution.m_ItemLUT.GetFlag(gDispensColumn) Then
                    SetCellValue(ij, gDispensColumn, Record1.pPara.DispenseFlag)
                End If

                If m_Execution.m_ItemLUT.GetFlag(gPos1XColumn) Then
                    SetCellValue(ij, gPos1XColumn, Record1.pPara.Pos1.X + Xij)
                    SetCellValue(ij, gPos1YColumn, Record1.pPara.Pos1.Y + Yij)
                    SetCellValue(ij, gPos1ZColumn, Record1.pPara.Pos1.Z)
                End If
                If m_Execution.m_ItemLUT.GetFlag(gPos2XColumn) Then
                    SetCellValue(ij, gPos2XColumn, Record1.pPara.Pos2.X + Xij)
                    SetCellValue(ij, gPos2YColumn, Record1.pPara.Pos2.Y + Yij)
                    SetCellValue(ij, gPos2ZColumn, Record1.pPara.Pos2.Z)
                End If
                If m_Execution.m_ItemLUT.GetFlag(gPos3XColumn) Then
                    SetCellValue(ij, gPos3XColumn, Record1.pPara.Pos3.X + Xij)
                    SetCellValue(ij, gPos3YColumn, Record1.pPara.Pos3.Y + Yij)
                    SetCellValue(ij, gPos3ZColumn, Record1.pPara.Pos3.Z)
                End If


                If m_Execution.m_ItemLUT.GetFlag(gTravelSpeedColumn) Then
                    SetCellValue(ij, gTravelSpeedColumn, Record1.pPara.TravelSpeed)
                End If
                If m_Execution.m_ItemLUT.GetFlag(gNeedleGapColumn) Then
                    SetCellValue(ij, gNeedleGapColumn, Record1.pPara.NeedleGap)
                End If
                If m_Execution.m_ItemLUT.GetFlag(gDurationColumn) Then
                    SetCellValue(ij, gDurationColumn, Record1.pPara.DispenseDuration)
                End If
                If m_Execution.m_ItemLUT.GetFlag(gTravelDelayColumn) Then
                    SetCellValue(ij, gTravelDelayColumn, Record1.pPara.TravelDelay)
                End If
                If m_Execution.m_ItemLUT.GetFlag(gRetractDelayColumn) Then
                    SetCellValue(ij, gRetractDelayColumn, Record1.pPara.RetractDelay)
                End If
                If m_Execution.m_ItemLUT.GetFlag(gApproachHtColumn) Then
                    SetCellValue(ij, gApproachHtColumn, Record1.pPara.ApproachDispHeight)
                End If
                If m_Execution.m_ItemLUT.GetFlag(gRetractSpeedColumn) Then
                    SetCellValue(ij, gRetractSpeedColumn, Record1.pPara.RetractSpeed)
                End If
                If m_Execution.m_ItemLUT.GetFlag(gRetractHtColumn) Then
                    SetCellValue(ij, gRetractHtColumn, Record1.pPara.RetractHeight)
                End If
                If m_Execution.m_ItemLUT.GetFlag(gClearanceHtColumn) Then
                    SetCellValue(ij, gClearanceHtColumn, Record1.pPara.ClearanceHeight)
                End If
                If m_Execution.m_ItemLUT.GetFlag(gDTaiDistColumn) Then
                    SetCellValue(ij, gDTaiDistColumn, Record1.pPara.DetailingDistance)
                End If
                If m_Execution.m_ItemLUT.GetFlag(gArcRadiusColumn) Then
                    SetCellValue(ij, gArcRadiusColumn, Record1.pPara.ArcRadius)
                End If
                If m_Execution.m_ItemLUT.GetFlag(gPitchColumn) Then
                    SetCellValue(ij, gPitchColumn, Record1.pPara.Pitch)
                End If
                If m_Execution.m_ItemLUT.GetFlag(gFillHeightColumn) Then
                    SetCellValue(ij, gFillHeightColumn, Record1.pPara.FilledHeight)
                End If
                If m_Execution.m_ItemLUT.GetFlag(gNoOfRunColumn) Then
                    SetCellValue(ij, gNoOfRunColumn, Record1.pPara.NumberOfRun)
                End If
                If m_Execution.m_ItemLUT.GetFlag(gSprialColumn) Then
                    SetCellValue(ij, gSprialColumn, Record1.pPara.SpiralFlag)
                End If
                If m_Execution.m_ItemLUT.GetFlag(gRtAngleColumn) Then
                    SetCellValue(ij, gRtAngleColumn, Record1.pPara.RotatingAngle)
                End If
                If m_Execution.m_ItemLUT.GetFlag(gEdgeClearColumn) Then
                    SetCellValue(ij, gEdgeClearColumn, Record1.pPara.EdgeClear)
                End If
                If m_Execution.m_ItemLUT.GetFlag(gIONumberColumn) Then
                    SetCellValue(ij, gIONumberColumn, Record1.pPara.IONumber)
                End If

            Next j
        Next i
        TraceGCCollect()
    End Sub




    'Build SubSub and/or SubArray data based on the sub input                                                                                                   
    '   iSubSheetName:            The full sub sheet name   

    Private Sub Spreadsheet_AddSubandArrayInSub(ByRef iSubSheetName As String)


        'Get current sheet name and row number
        Dim TmpSheetName As String = GetActiveSheetName()
        Dim TmpRowNumber As Integer = m_Row

        If 0 = m_Execution.m_Pattern.Spreadsheet_CheckSubsheetExist( _
            AxSpreadsheetProgramming, m_Execution.m_File.NameOnlyFromFullPath(iSubSheetName)) Then
            m_Execution.m_Pattern.LoadTxtPatternPara(AxSpreadsheetProgramming, iSubSheetName, 1, 0, False)
        End If

        'Restore to the previous sheet and rows
        m_Row = TmpRowNumber

        AxSpreadsheetProgramming.Worksheets(TmpSheetName).Activate()

    End Sub




    'Get array data from the DialogBox                                                
    '   AxSpreadsheet:      ActiveX spreadsheet                                                    
    '   Record:             One row for the array   
    '   CurrentPt:          The number of position, such x1y1z1=1, x2y2z2=2....  
    '   ArrayDlg:           Array generation DialogBox 

    Private Sub Spreadsheet_GetArrayRecord( _
            ByRef Record As CIDSPattern.PatternRecord, _
            ByVal CurrentPt As Integer, ByVal ArrayDlg As ArrayGenerate)

        Dim Z1 As Double = CDbl(ArrayDlg.ArrayPara1.Pos1.Z)

        Record.pPara.Name = ArrayDlg.ArrayPara1.Name
        m_Execution.m_ItemLUT.Cmd2Index((Record.pPara.Name).ToUpper)

        If m_Execution.m_ItemLUT.GetFlag(gNeedleColumn) Then
            Record.pPara.Needle = ArrayDlg.ArrayPara1.Needle
        End If
        If m_Execution.m_ItemLUT.GetFlag(gDispensColumn) Then
            Record.pPara.DispenseFlag = ArrayDlg.ArrayPara1.DispenseFlag
        End If

        If 1 = CurrentPt Then
            If m_Execution.m_ItemLUT.GetFlag(gPos1XColumn) Then
                Record.pPara.Pos1.X = CDbl(ArrayDlg.ArrayPara1.Pos1.X)
                Record.pPara.Pos1.Y = CDbl(ArrayDlg.ArrayPara1.Pos1.Y)
                Record.pPara.Pos1.Z = Z1
            End If
            If m_Execution.m_ItemLUT.GetFlag(gPos2XColumn) Then
                Record.pPara.Pos2.X = CDbl(ArrayDlg.ArrayPara1.Pos2.X)
                Record.pPara.Pos2.Y = CDbl(ArrayDlg.ArrayPara1.Pos2.Y)
                Record.pPara.Pos2.Z = Z1
            End If
            If m_Execution.m_ItemLUT.GetFlag(gPos3XColumn) Then
                Record.pPara.Pos3.X = CDbl(ArrayDlg.ArrayPara1.pos3.X)
                Record.pPara.Pos3.Y = CDbl(ArrayDlg.ArrayPara1.pos3.Y)
                Record.pPara.Pos3.Z = Z1
            End If
        ElseIf 2 = CurrentPt Then
            If m_Execution.m_ItemLUT.GetFlag(gPos1XColumn) Then
                Record.pPara.Pos1.X = CDbl(ArrayDlg.ArrayPara2.Pos1.X)
                Record.pPara.Pos1.Y = CDbl(ArrayDlg.ArrayPara2.Pos1.Y)
                Record.pPara.Pos1.Z = Z1
            End If
            If m_Execution.m_ItemLUT.GetFlag(gPos2XColumn) And _
                    ("Circle" = Record.pPara.Name Or "FillCircle" = Record.pPara.Name) Then
                Record.pPara.Pos2.X = CDbl(ArrayDlg.ArrayPara2.Pos2.X)
                Record.pPara.Pos2.Y = CDbl(ArrayDlg.ArrayPara2.Pos2.Y)
                Record.pPara.Pos2.Z = Z1
            End If
            If m_Execution.m_ItemLUT.GetFlag(gPos3XColumn) And _
                    ("Circle" = Record.pPara.Name Or "FillCircle" = Record.pPara.Name) Then
                Record.pPara.Pos3.X = CDbl(ArrayDlg.ArrayPara2.pos3.X)
                Record.pPara.Pos3.Y = CDbl(ArrayDlg.ArrayPara2.pos3.Y)
                Record.pPara.Pos3.Z = Z1
            End If
        ElseIf 3 = CurrentPt Then
            If m_Execution.m_ItemLUT.GetFlag(gPos1XColumn) Then
                Record.pPara.Pos1.X = CDbl(ArrayDlg.ArrayPara3.Pos1.X)
                Record.pPara.Pos1.Y = CDbl(ArrayDlg.ArrayPara3.Pos1.Y)
                Record.pPara.Pos1.Z = Z1
            End If
            If m_Execution.m_ItemLUT.GetFlag(gPos2XColumn) And _
                    ("Circle" = Record.pPara.Name Or "FillCircle" = Record.pPara.Name) Then
                Record.pPara.Pos2.X = CDbl(ArrayDlg.ArrayPara3.Pos2.X)
                Record.pPara.Pos2.Y = CDbl(ArrayDlg.ArrayPara3.Pos2.Y)
                Record.pPara.Pos2.Z = Z1
            End If
            If m_Execution.m_ItemLUT.GetFlag(gPos3XColumn) And _
                    ("Circle" = Record.pPara.Name Or "FillCircle" = Record.pPara.Name) Then
                Record.pPara.Pos3.X = CDbl(ArrayDlg.ArrayPara3.pos3.X)
                Record.pPara.Pos3.Y = CDbl(ArrayDlg.ArrayPara3.pos3.Y)
                Record.pPara.Pos3.Z = Z1
            End If
        End If


        If m_Execution.m_ItemLUT.GetFlag(gTravelSpeedColumn) Then
            Record.pPara.TravelSpeed = CDbl(ArrayDlg.ArrayPara1.TravelSpeed)
        End If
        If m_Execution.m_ItemLUT.GetFlag(gNeedleGapColumn) Then
            Record.pPara.NeedleGap = CDbl(ArrayDlg.ArrayPara1.NeedleGap)
        End If
        If m_Execution.m_ItemLUT.GetFlag(gDurationColumn) Then
            Record.pPara.DispenseDuration = CDbl(ArrayDlg.ArrayPara1.DispenseDuration)
        End If
        If m_Execution.m_ItemLUT.GetFlag(gTravelDelayColumn) Then
            Record.pPara.TravelDelay = CDbl(ArrayDlg.ArrayPara1.TravelDelay)
        End If
        If m_Execution.m_ItemLUT.GetFlag(gRetractDelayColumn) Then
            Record.pPara.RetractDelay = CDbl(ArrayDlg.ArrayPara1.RetractDelay)
        End If
        If m_Execution.m_ItemLUT.GetFlag(gApproachHtColumn) Then
            Record.pPara.ApproachDispHeight = CDbl(ArrayDlg.ArrayPara1.ApproachDispHeight)
        End If
        If m_Execution.m_ItemLUT.GetFlag(gRetractSpeedColumn) Then
            Record.pPara.RetractSpeed = CDbl(ArrayDlg.ArrayPara1.RetractSpeed)
        End If
        If m_Execution.m_ItemLUT.GetFlag(gRetractHtColumn) Then
            Record.pPara.RetractHeight = CDbl(ArrayDlg.ArrayPara1.RetractHeight)
        End If
        If m_Execution.m_ItemLUT.GetFlag(gClearanceHtColumn) Then
            Record.pPara.ClearanceHeight = CDbl(ArrayDlg.ArrayPara1.ClearanceHeight)
        End If
        If m_Execution.m_ItemLUT.GetFlag(gDTaiDistColumn) Then
            Record.pPara.DetailingDistance = CDbl(ArrayDlg.ArrayPara1.DetailingDistance)
        End If
        If m_Execution.m_ItemLUT.GetFlag(gArcRadiusColumn) Then
            Record.pPara.ArcRadius = CDbl(ArrayDlg.ArrayPara1.ArcRadius)
        End If
        If m_Execution.m_ItemLUT.GetFlag(gPitchColumn) Then
            Record.pPara.Pitch = CDbl(ArrayDlg.ArrayPara1.Pitch)
        End If
        If m_Execution.m_ItemLUT.GetFlag(gFillHeightColumn) Then
            Record.pPara.FilledHeight = CDbl(ArrayDlg.ArrayPara1.FilledHeight)
        End If
        If m_Execution.m_ItemLUT.GetFlag(gNoOfRunColumn) Then
            Record.pPara.NumberOfRun = CDbl(ArrayDlg.ArrayPara1.NumofRun)
        End If
        If m_Execution.m_ItemLUT.GetFlag(gSprialColumn) Then
            If "FILLCIRCLE" = (Record.pPara.Name).ToUpper Or "FILLRECTANGLE" = (Record.pPara.Name).ToUpper Then
                Record.pPara.SpiralFlag = ArrayDlg.ArrayPara1.SpiralFlag
            Else
                Record.pPara.SpiralFlag = CBool(ArrayDlg.ArrayPara1.SpiralFlag)
            End If
        End If
        If m_Execution.m_ItemLUT.GetFlag(gRtAngleColumn) Then
            Record.pPara.RotatingAngle = CDbl(ArrayDlg.ArrayPara1.RotatingAngle)
        End If
        If m_Execution.m_ItemLUT.GetFlag(gEdgeClearColumn) Then
            Record.pPara.EdgeClear = CDbl(ArrayDlg.ArrayPara1.EdgeClearance)
        End If
        If m_Execution.m_ItemLUT.GetFlag(gIONumberColumn) Then
            Record.pPara.IONumber = CInt(ArrayDlg.ArrayPara1.IONumber)
        End If

        TraceGCCollect()
    End Sub

    'Get array data from a row                                                
    '   AxSpreadsheetProgramming: ActiveX spreadsheet, used as global                                                   
    '   Record:             One row data record   
    '   CurrentRow:         The row in the Spreadsheet  
    Private Sub Spreadsheet_GetSheetRecord( _
            ByRef Record As CIDSPattern.PatternRecord, _
            ByVal CurrentRow As Integer)

        Dim strName As String = GetCellValue(CurrentRow, gCommandNameColumn)
        Dim strName2(2) As String
        strName2 = strName.Split("_")
        strName = strName2(1)


        With AxSpreadsheetProgramming.ActiveSheet
            Record.pPara.Name = strName
            Record.pPara.Needle = .Cells(CurrentRow, gNeedleColumn).Value
            Record.pPara.DispenseFlag = .Cells(CurrentRow, gDispensColumn).Value

            Record.pPara.Pos1.X = .Cells(CurrentRow, gPos1XColumn).Value
            Record.pPara.Pos1.Y = .Cells(CurrentRow, gPos1YColumn).Value
            Record.pPara.Pos1.Z = .Cells(CurrentRow, gPos1ZColumn).Value
            Record.pPara.Pos2.X = .Cells(CurrentRow, gPos2XColumn).Value
            Record.pPara.Pos2.Y = .Cells(CurrentRow, gPos2YColumn).Value
            Record.pPara.Pos2.Z = .Cells(CurrentRow, gPos2ZColumn).Value
            Record.pPara.Pos3.X = .Cells(CurrentRow, gPos3XColumn).Value
            Record.pPara.Pos3.Y = .Cells(CurrentRow, gPos3YColumn).Value
            Record.pPara.Pos3.Z = .Cells(CurrentRow, gPos3ZColumn).Value


            Record.pPara.TravelSpeed = .Cells(CurrentRow, gTravelSpeedColumn).Value
            Record.pPara.NeedleGap = .Cells(CurrentRow, gNeedleGapColumn).Value
            Record.pPara.DispenseDuration = .Cells(CurrentRow, gDurationColumn).Value
            Record.pPara.TravelDelay = .Cells(CurrentRow, gTravelDelayColumn).Value
            Record.pPara.RetractDelay = .Cells(CurrentRow, gRetractDelayColumn).Value
            Record.pPara.ApproachDispHeight = .Cells(CurrentRow, gApproachHtColumn).Value
            Record.pPara.RetractSpeed = .Cells(CurrentRow, gRetractSpeedColumn).Value
            Record.pPara.RetractHeight = .Cells(CurrentRow, gRetractHtColumn).Value


            Record.pPara.ClearanceHeight = .Cells(CurrentRow, gClearanceHtColumn).Value
            Record.pPara.DetailingDistance = .Cells(CurrentRow, gDTaiDistColumn).Value
            Record.pPara.ArcRadius = .Cells(CurrentRow, gArcRadiusColumn).Value
            Record.pPara.Pitch = .Cells(CurrentRow, gPitchColumn).Value
            Record.pPara.FilledHeight = .Cells(CurrentRow, gFillHeightColumn).Value


            Record.pPara.NumberOfRun = .Cells(CurrentRow, gNoOfRunColumn).Value
            Record.pPara.SpiralFlag = .Cells(CurrentRow, gSprialColumn).Value
            Record.pPara.RotatingAngle = .Cells(CurrentRow, gRtAngleColumn).Value


            Record.pPara.EdgeClear = .Cells(CurrentRow, gEdgeClearColumn).Value
            Record.pPara.IONumber = .Cells(CurrentRow, gIONumberColumn).Value

        End With
        TraceGCCollect()
    End Sub


    'Is the row empty?
    '   row:  The target row                                                                                
    '
    'Return: True=empty, False=Not empty

    Public Function isEmptyRow(ByVal row As Integer) As Boolean
        Dim i As Integer
        Dim str As String

        For i = 1 To gMaxColumns
            str = GetCellValue(row, i)
            If str <> Nothing Then
                str = str.Trim(" ")
                If str <> "" Then
                    Return False
                End If
            End If
        Next
        Return True
    End Function


    'Private Sub ToolBarYesNo_ButtonClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ToolBarButtonClickEventArgs)
    '    If MachineRunning() Then
    '        Return
    '    End If
    '    offset(0) = gLeftNeedleOffs(0)
    '    offset(1) = gLeftNeedleOffs(1)
    '    offset(2) = gLeftNeedleOffs(2)

    '    If e.Button Is TeachingToolBar1.Buttons(0) Then
    '        If type <> "Move" Then
    '            If CheckSoftLimitXYZ(m_MachinePos, offset) Then Exit Sub
    '        End If
    '        Confirm()   'Add rows
    '    Else
    '        If m_EditStateFlag = False Then
    '            'DelayForRowDelete()
    '            If m_SteppingPostFlag Then
    '                DisableCoordinateUpdateInSpreadsheet()
    '                Spreadsheet_CheckForWithinLinkRange(True)
    '                DisableTeachingToolbarOKButton()
    '                EnableTeachingToolbarCancelButton()
    '                DisableEditingToolbar()

    '                SelectCell(m_Row, gCommandNameColumn)
    '                m_SteppingPostFlag = False
    '                LabelMessage("")
    '            Else
    '                If (MessageBox.Show("Are you sure you want to delete the row/rows?", "Warning", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) = DialogResult.Yes) Then
    '                    DisableCoordinateUpdateInSpreadsheet()
    '                    'DelayForRowDelete()
    '                    type = GetCellValue(m_Row, gCommandNameColumn)
    '                    If type = "GlobalQC" Then
    '                        IDS.Data.Hardware.Camera.DotQCEnable = False
    '                    End If
    '                    Cancel()   'Cancel or delete rows
    '                    AxSpreadsheetProgramming.Enabled = True
    '                End If
    '            End If
    '            'DeletingRowFromExcel = False
    '            'DeletingRowFinished = False
    '        Else
    '            'm_EditStateFlag = False
    '            'm_SteppingPostFlag = False
    '            Console.WriteLine("Disable spredsheet update")
    '            DisableCoordinateUpdateInSpreadsheet()
    '            Spreadsheet_CheckForWithinLinkRange(True)
    '            DisableTeachingToolbarOKButton()
    '            EnableTeachingToolbarCancelButton()
    '            DisableEditingToolbar()

    '            SelectCell(m_Row, gCommandNameColumn)
    '            LabelMessage("")
    '            Cancel()
    '            'If (MessageBox.Show("Are you sure you want to delete the row/rows?", "Warning", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) = DialogResult.Yes) Then
    '            'Cancel()   'Cancel or delete rows
    '            'End If

    '        End If
    '    End If

    '    EnableTeachModeSwitching()

    'End Sub

    Private Sub DelayForRowDelete()
        Dim W_Time, Current_Time As Single
        DeletingRowFromExcel = True
        W_Time = Microsoft.VisualBasic.Timer
        Do While (m_PosUpdate = True And DeletingRowFinished = False)
            Current_Time = Microsoft.VisualBasic.Timer
            If (Current_Time < W_Time) Then
                Current_Time = Current_Time - (86400 - W_Time)
            Else
                Current_Time = Current_Time - W_Time
            End If
            If (Current_Time >= 0.8) Then
                Exit Do
            End If
            TraceDoEvents()
            MySleep(20)
        Loop
    End Sub

    Dim x1, x2, x3 As Double
    Dim y1, y2, y3 As Double

    Private Sub Cancel()

        Dim i As Integer
        'status: 1=ok, 2=cancel, 3=end with first fiducial

        Dim status As Integer = 0
        Dim PatternLineRecord(3) As CIDSPattern.PatternRecord
        Dim cell1, cell2 As Object

        m_Row = GetActiveCellRow() 'Get current row number

        Dim sel As Microsoft.Office.Interop.OWC.Range = AxSpreadsheetProgramming.Selection
        Dim m_rowCount As Integer = sel.Rows.Count()
        Dim currenCmdName As String = ""
        Try
            If m_rowCount = 1 Then
                type = GetCellValue(m_Row, gCommandNameColumn)
                'UndoData_Logging(0)

                If "" = type Then
                    DeleteRow(m_Row)

                    SaveProgram.UnSave = True
                ElseIf m_EditStateFlag = True Then 'edit 
                    If m_Execution.m_Pattern.m_ErrorChk.CheckRequiredPt3XY(type) Then
                        cell1 = GetCell(m_Row, gPos1XColumn)
                        cell2 = GetCell(m_Row, gPos3ZColumn)
                        m_Execution.m_Pattern.Spreadsheet_CellGrey(cell1, cell2, True, AxSpreadsheetProgramming)
                        UpdateLinkForNextRow()
                        UpdateLinkForPreviousRow()

                    ElseIf m_Execution.m_Pattern.m_ErrorChk.CheckRequiredPt2XY(type) Then
                        cell1 = GetCell(m_Row, gPos1XColumn)
                        cell2 = GetCell(m_Row, gPos2ZColumn)
                        m_Execution.m_Pattern.Spreadsheet_CellGrey(cell1, cell2, True, AxSpreadsheetProgramming)
                        UpdateLinkForNextRow()
                        UpdateLinkForPreviousRow()

                    ElseIf m_Execution.m_Pattern.m_ErrorChk.CheckRequiredPt1XY(type) Then
                        cell1 = GetCell(m_Row, gPos1XColumn)
                        cell2 = GetCell(m_Row, gPos1ZColumn)
                        m_Execution.m_Pattern.Spreadsheet_CellGrey(cell1, cell2, True, AxSpreadsheetProgramming)

                    End If

                    'lim
                    Dim CmdName As String
                    m_Row = GetActiveCellRow()
                    CmdName = GetCellValue(m_Row, gCommandNameColumn)
                    currenCmdName = CmdName
                    If "HEIGHT" = CmdName.ToUpper Then
                        Dim posBack(3), posTo(3), posOffset(3) As Double
                        posBack(0) = GetCellValue(m_Row, gPos1XColumn)
                        posBack(1) = GetCellValue(m_Row, gPos1YColumn)
                        posBack(2) = 0
                        posOffset(0) = posOffset(1) = posOffset(2) = 0
                        posTo(0) = posBack(0) - gLaserOffX
                        posTo(1) = posBack(1) - gLaserOffY
                        posTo(2) = posBack(2)

                        MoveToSpreadsheetPoint(posTo, "Vision")

                        Dim rtn As Integer
                        OnLaser()
                        rtn = Laser.WaitForReadingToStabilize()
                        OffLaser()
                        If rtn Then                      'No laser readout error
                            posOffset(2) = Laser.LASER_Reading - IDS.Data.Hardware.HeightSensor.Laser.HeightReference
                            posOffset(2) = CInt(posOffset(2) * 1000) / 1000
                            Console.WriteLine("Laser Reading: " & Laser.LASER_Reading)
                            SetCellValue(m_Row, gPos1ZColumn, posOffset(2))
                            Dim p As Double = Laser.LASER_Reading
                            Me.tbLsHeight.Text = p.ToString("0.000") + "-" + IDS.Data.Hardware.HeightSensor.Laser.HeightReference.ToString("0.000") + " = " + posOffset(2).ToString("0.000")
                        Else
                            MessageBox.Show("laser sensor out of range")
                            m_TeachStepNumber = 0
                            ToggleButtonsForTeachingStop()
                            DisableElementsCommandBlockButton(gSeperatorCmdIndex)
                            DeleteRow(m_Row)
                        End If
                        posBack(0) = posBack(0)
                        posBack(1) = posBack(1)
                        posBack(2) = posBack(2)
                        MoveToSpreadsheetPoint(posBack, "Vision")
                        'SelectCell(m_Row + 1, gCommandNameColumn)
                    ElseIf "FIDUCIAL" = CmdName.ToUpper Then
                        'For Lim's part

                        Dim brightness As Integer

                        If m_TeachStepNumber = 1 Then

                            SwitchToVisionTeachMode()
                            DisableTeachModeSwitching()

                            brightness = GetCellValue(m_Row, gBrightnessColumn)
                            DisableProgrammingBrightnessToggle()
                            Me.DisableCalibButtons()
                            Vision.FiducialMark_form.FormCloseEvent = New DLL_Export_Device_Vision.FiducialForm.FormCloseDelegate(AddressOf Me.EditFiducialCommandFormResponse)
                            Vision.IDSV_Form_FI_Edit(1, gPatternFileName + "1.mmo", brightness)
                            Me.EnableClickToMove()
                            'While status = 0
                            '    TraceDoEvents()
                            '    status = Vision.GetFiducialStatus()
                            'End While
                            'Dim fidname As String
                            'If status = 1 Then
                            '    fidname = Vision.GetFiducialFilename()
                            '    Dim Brightness1 As Integer = Vision.GetFiducialBrightness
                            '    SetCellValue(m_Row, gFid1Column, fidname)
                            '    SetCellValue(m_Row, gBrightnessColumn, Brightness1)
                            'ElseIf 2 = status Then  'Cancel command
                            '    DisableCoordinateUpdateInSpreadsheet()
                            '    m_TeachStepNumber = 0
                            '    ToggleButtonsForTeachingStop()
                            '    DisableElementsCommandBlockButton(gSeperatorCmdIndex)
                            'ElseIf 3 = status Then
                            '    fidname = Vision.GetFiducialFilename()
                            '    SetCellValue(m_Row, gFid1Column, fidname)
                            'End If
                        ElseIf m_TeachStepNumber = 2 Then
                            brightness = GetCellValue(m_Row, gThresholdColumn) 'fiducial 2 brightness value
                            Vision.FrmVision.SetBrightness(brightness)
                            DisableProgrammingBrightnessToggle()
                            Vision.FiducialMark_form.FormCloseEvent = New DLL_Export_Device_Vision.FiducialForm.FormCloseDelegate(AddressOf Me.EditFiducialCommandFormResponse)
                            Me.EnableClickToMove()
                            Vision.IDSV_Form_FI_Edit(2, gPatternFileName + "2.mmo", brightness)
                            'While status = 0
                            '    TraceDoEvents()
                            '    status = Vision.GetFiducialStatus()
                            'End While
                            'Dim fidname As String
                            'If status = 1 Then
                            '    fidname = Vision.GetFiducialFilename()
                            '    Dim Brightness2 As Integer = Vision.GetFiducialBrightness
                            '    SetCellValue(m_Row, gFid2Column, fidname)
                            '    SetCellValue(m_Row, gThresholdColumn, Brightness2) 'For brightness fiducial no.2
                            'ElseIf 2 = status Then
                            'End If
                        End If
                        'EnableProgrammingBrightnessToggle()
                    ElseIf "REJECT" = CmdName.ToUpper Then 'For Lim's part
                        EnableCoordinateUpdateInSpreadsheet()

                        Dim vPara As DLL_Export_Device_Vision.RejectPoint.RMParam
                        m_Row = GetActiveCellRow()
                        vpara._AcceptRatio = GetCellValue(m_Row, gAcceptRatioCoulumn)
                        vpara._Binarized = GetCellValue(m_Row, gBinarizedColumn)
                        vpara._BlackWithoutRM = GetCellValue(m_Row, gBlackWithoutRMCoulumn)
                        vpara._BlackWithRM = GetCellValue(m_Row, gBlackWithRMCoulumn)
                        vpara._Brightness = GetCellValue(m_Row, gBrightnessColumn)
                        vpara._MRegionX = GetCellValue(m_Row, gMRegionXColumn)
                        vpara._MRegionY = GetCellValue(m_Row, gMRegionYColumn)
                        vpara._MROIx = GetCellValue(m_Row, gMROIxColumn)
                        vpara._MROIy = GetCellValue(m_Row, gMROIyColumn)
                        vpara._WhiteWithoutRM = GetCellValue(m_Row, gWhiteWithoutRMCoulumn)
                        vpara._WhiteWithRM = GetCellValue(m_Row, gWhiteWithRMCoulumn)
                        vpara._WoRM = GetCellValue(m_Row, gWoRMCoulumn)
                        vpara._WRM = GetCellValue(m_Row, gWRMCoulumn)
                        Vision.IDSV_Form_RM_Edit(vpara)

                        Dim x, y As Double

                        While status = 0
                            TraceDoEvents()
                            status = Vision.GetRMStatus
                        End While
                        DisableCoordinateUpdateInSpreadsheet()
                        EnableTeachingButtons()
                        If status = 2 Then          'Cancel
                            ElementsCommandBlock.Enabled = True
                            ReferenceCommandBlock.Enabled = True
                            UpdateSpreadsheet()
                            SelectCell(m_Row, gCommandNameColumn)
                        ElseIf status = 1 Then      'Ok
                            SetRejectMarkSettings()
                            SelectCell(m_Row + 1, gCommandNameColumn)
                        End If

                    ElseIf "CHIPEDGE" = CmdName.ToUpper Then 'lim
                        EnableCoordinateUpdateInSpreadsheet()

                        Dim vPara As DLL_Export_Device_Vision.ChipEdgePoints.ChipEdgeParam
                        m_Row = GetActiveCellRow()
                        vPara._Brightness = GetCellValue(m_Row, gBrightnessColumn)
                        vPara._EdgeClearance = GetCellValue(m_Row, gEdgeClearColumn)
                        vPara._CheckBox_ChipRec_Enable = GetCellValue(m_Row, gCheckBoxColumn)
                        vPara._Cw_CCw = GetCellValue(m_Row, gCwCCwColumn)
                        vPara._DispenseModel = GetCellValue(m_Row, gDispModelColumn)
                        vPara._Inside_out = GetCellValue(m_Row, gInOutColumn)
                        vPara._MainEdge = GetCellValue(m_Row, gMainEdgeColumn)
                        vPara._PointX1 = GetCellValue(m_Row, gPointX1Column)
                        vPara._PointX2 = GetCellValue(m_Row, gPointX2Column)
                        vPara._PointX3 = GetCellValue(m_Row, gPointX3Column)
                        vPara._PointX4 = GetCellValue(m_Row, gPointX4Column)
                        vPara._PointX5 = GetCellValue(m_Row, gPointX5Column)
                        vPara._PointY1 = GetCellValue(m_Row, gPointY1Column)
                        vPara._PointY2 = GetCellValue(m_Row, gPointY2Column)
                        vPara._PointY3 = GetCellValue(m_Row, gPointY3Column)
                        vPara._PointY4 = GetCellValue(m_Row, gPointY4Column)
                        vPara._PointY5 = GetCellValue(m_Row, gPointY5Column)
                        vPara._Pos = GetCellValue(m_Row, gPosColumn)
                        vPara._PosX = GetCellValue(m_Row, gPosXColumn)
                        vPara._PosY = GetCellValue(m_Row, gPosYColumn)
                        vPara._ROI = GetCellValue(m_Row, gROIColumn)
                        vPara._Rot = GetCellValue(m_Row, gRotColumn)
                        vPara._Size = GetCellValue(m_Row, gSizeColumn)
                        vPara._SizeX = GetCellValue(m_Row, gSizeXColumn)
                        vPara._SizeY = GetCellValue(m_Row, gSizeYColumn)
                        vPara._Threshold = GetCellValue(m_Row, gThresholdColumn)
                        vPara._Vertical = GetCellValue(m_Row, gVerticalColumn)
                        vPara._DotDispensingDuration = GetCellValue(m_Row, gDurationColumn)
                        vPara._Contrast = GetCellValue(m_Row, gCompactnessColumn)
                        DisableProgrammingBrightnessToggle()
                        Vision.ChipEdgePoints_form.FormCloseEvent = New DLL_Export_Device_Vision.ChipEdgePoints.FormCloseDelegate(AddressOf Me.EditChipEdgeCommandFormResponse)
                        Vision.IDSV_Form_CE_Edit(vPara)
                        Vision.FrmVision.ChipEdge5PointsSet = New DLL_Export_Device_Vision.FormVision.ChipEdgeSetDelegate(AddressOf Me.ChipEdgeAdjustXY)
                        EnableClickToMove()
                        'Dim x, y As Double
                        'Dim pos(3) As Double
                        'Do
                        '    While Not (Vision.RobotMotionOffset(x, y) = True Or Vision.GetChipEdgeStatus = 2 Or Vision.GetChipEdgeStatus = 1)
                        '        TraceDoEvents()
                        '    End While

                        '    'moverobot
                        '    pos(0) = x
                        '    pos(1) = -y
                        '    pos(2) = 0

                        '    m_Tri.Set_XY_Speed(IDS.Data.Hardware.Gantry.ElementXYSpeed)
                        '    m_Tri.MoveRelative_XY(pos)

                        '    If status = 3 Then
                        '        status = 0 'reset status after 5 points being reset
                        '    End If

                        '    While status = 0
                        '        TraceDoEvents()
                        '        status = Vision.GetChipEdgeStatus()
                        '    End While

                        'Loop While status = 3 'status 3= reset 5 points

                        'DisableCoordinateUpdateInSpreadsheet()
                        'ElementsCommandBlock.Enabled = True
                        'ReferenceCommandBlock.Enabled = True
                        'EnableTeachingButtons()
                        'EnableProgrammingBrightnessToggle()
                        'DisableElementsCommandBlockButton(gSeperatorCmdIndex)
                        'm_ProgrammingStateFlag = False
                        'If status = 2 Then
                        '    UpdateSpreadsheet()
                        '    SelectCell(m_Row, gCommandNameColumn)
                        'ElseIf status = 1 Then
                        '    Dim bb As Boolean
                        '    Vision.GetChipEdgeParameters(vPara)
                        '    Vision.SetCEReset()

                        '    SetCellValue(m_Row, gEdgeClearColumn, vPara._EdgeClearance)
                        '    SetCellValue(m_Row, gBrightnessColumn, vPara._Brightness)
                        '    SetCellValue(m_Row, gCheckBoxColumn, vPara._CheckBox_ChipRec_Enable)
                        '    SetCellValue(m_Row, gCwCCwColumn, vPara._Cw_CCw)
                        '    SetCellValue(m_Row, gDispModelColumn, vPara._DispenseModel)
                        '    SetCellValue(m_Row, gInOutColumn, vPara._Inside_out)
                        '    SetCellValue(m_Row, gMainEdgeColumn, vPara._MainEdge)
                        '    SetCellValue(m_Row, gPointX1Column, vPara._PointX1)
                        '    SetCellValue(m_Row, gPointX2Column, vPara._PointX2)
                        '    SetCellValue(m_Row, gPointX3Column, vPara._PointX3)
                        '    SetCellValue(m_Row, gPointX4Column, vPara._PointX4)
                        '    SetCellValue(m_Row, gPointX5Column, vPara._PointX5)
                        '    SetCellValue(m_Row, gPointY1Column, vPara._PointY1)
                        '    SetCellValue(m_Row, gPointY2Column, vPara._PointY2)
                        '    SetCellValue(m_Row, gPointY3Column, vPara._PointY3)
                        '    SetCellValue(m_Row, gPointY4Column, vPara._PointY4)
                        '    SetCellValue(m_Row, gPointY5Column, vPara._PointY5)
                        '    SetCellValue(m_Row, gPosColumn, vPara._Pos)
                        '    SetCellValue(m_Row, gPosXColumn, vPara._PosX)
                        '    SetCellValue(m_Row, gPosYColumn, vPara._PosY)
                        '    SetCellValue(m_Row, gROIColumn, vPara._ROI)
                        '    SetCellValue(m_Row, gRotColumn, vPara._Rot)
                        '    SetCellValue(m_Row, gSizeColumn, vPara._Size)
                        '    SetCellValue(m_Row, gSizeXColumn, vPara._SizeX)
                        '    SetCellValue(m_Row, gSizeYColumn, vPara._SizeY)
                        '    SetCellValue(m_Row, gThresholdColumn, vPara._Threshold)
                        '    SetCellValue(m_Row, gVerticalColumn, vPara._Vertical)
                        '    SetCellValue(m_Row, gDurationColumn, vPara._DotDispensingDuration)
                        '    SetCellValue(m_Row, gCompactnessColumn, vPara._Contrast)
                        '    cell1 = GetCell(m_Row, gPos1XColumn)
                        '    cell2 = GetCell(m_Row, gPos1ZColumn)
                        '    m_Execution.m_Pattern.Spreadsheet_CellGrey(cell1, cell2, True, AxSpreadsheetProgramming)
                        '    SelectCell(m_Row, 1)
                        'End If

                    ElseIf "QC" = CmdName.ToUpper Or "GLOBALQC" = CmdName.ToUpper Then  'lim
                        If "QC" = CmdName.ToUpper Then
                            EnableCoordinateUpdateInSpreadsheet()
                            Vision.SetGlobalQCMode(False)
                        Else
                            Vision.SetGlobalQCMode(True)
                        End If

                        Dim vPara As DLL_Export_Device_Vision.QC.QCParam
                        m_Row = GetActiveCellRow()
                        vPara._Brightness = GetCellValue(m_Row, gBrightnessColumn)
                        vPara._BlackDot = GetCellValue(m_Row, gBlackDotColumn)
                        vPara._Binarized = GetCellValue(m_Row, gBinarizedColumn)
                        vPara._MaxArea = GetCellValue(m_Row, gMaxAreaColumn)
                        vPara._MinArea = GetCellValue(m_Row, gMinAreaColumn)
                        vPara._Close = GetCellValue(m_Row, gCloseColumn)
                        vPara._Open = GetCellValue(m_Row, gOpenColumn)
                        vPara._Roughness = GetCellValue(m_Row, gRoughnessColumn)
                        vPara._Compactness = GetCellValue(m_Row, gCompactnessColumn)
                        vPara._MRegionX = GetCellValue(m_Row, gMRegionXColumn)
                        vPara._MRegionY = GetCellValue(m_Row, gMRegionYColumn)
                        vPara._MROIx = GetCellValue(m_Row, gMROIxColumn)
                        vPara._MROIy = GetCellValue(m_Row, gMROIyColumn)
                        vPara._Tolerance = GetCellValue(m_Row, gToleranceColumn)
                        vPara._Diameter = GetCellValue(m_Row, gDiameterColumn)
                        DisableProgrammingBrightnessToggle()
                        'If Vision.QC_form.FormCloseEvent Is Nothing Then
                        Vision.QC_form.FormCloseEvent = New DLL_Export_Device_Vision.QC.FormCloseDelegate(AddressOf Me.QCFormResponse)
                        'End If
                        DisableCalibButtons()
                        EnableClickToMove()
                        Vision.IDSV_Form_QC_Edit(vPara)

                        'While status = 0 Or status = 3
                        '    Do
                        '        TraceDoEvents()
                        '        status = Vision.GetQCStatus()
                        '    Loop While status = 3
                        'End While
                        'DisableCoordinateUpdateInSpreadsheet()
                        'EnableProgrammingBrightnessToggle()
                        'EnableTeachingButtons()
                        'DisableElementsCommandBlockButton(gSeperatorCmdIndex)
                        'm_ProgrammingStateFlag = False
                        'If status = 2 Then 'Cancel
                        '    SelectCell(m_Row, 1)
                        '    Vision.SetQCReset()
                        'ElseIf status = 1 Then                         'Ok
                        '    SetQCSettings()
                        '    SelectCell(m_Row + 1, 1)
                        'End If
                    End If
                    If Not ("QC" = currenCmdName.ToUpper Or "GLOBALQC" = currenCmdName.ToUpper) Then
                        MenuEditCopy.Enabled = True
                        If "" <> CopiedSheetName Then
                            MenuEditPaste.Enabled = True
                        End If
                        MenuEditCut.Enabled = True
                        MenuEditUndo.Enabled = True
                        MenuEditRedo.Enabled = False
                        SaveProgram.UnSave = True
                    End If
                ElseIf m_ProgrammingStateFlag And "LINKSTART" <> type.ToUpper And "LINKEND" <> type.ToUpper Then
                    DeleteRow(m_Row)
                    SelectCell(m_Row, 1)
                    m_ProgrammingStateFlag = False

                    SaveProgram.UnSave = True
                Else
                    Spreadsheet_DeleteMultiRow(AxSpreadsheetProgramming)
                    m_ProgrammingStateFlag = False

                    SaveProgram.UnSave = True
                End If
            Else
                'Multiple rows delete
                'UndoData_Logging(0)
                Spreadsheet_DeleteMultiRow(AxSpreadsheetProgramming)
                SaveProgram.UnSave = True
            End If
            If Not ("CHIPEDGE" = currenCmdName.ToUpper) Then
                If Not ("QC" = currenCmdName.ToUpper Or "GLOBALQC" = currenCmdName.ToUpper) Then
                    UndoData_Logging(0)
                    m_EditStateFlag = False
                    Spreadsheet_CheckForWithinLinkRange(True)
                    DisableTeachingToolbarOKButton()
                    EnableTeachingToolbarCancelButton()
                    DisableEditingToolbar()
                    DisableElementsCommandBlockButton(gSeperatorCmdIndex)
                    DisableElementsCommandBlockButton(gOffsetCmdIndex)
                    DisplaySpreadsheetTabs()
                    If status = 2 Then
                        If Not ("GLOBALQC" = currenCmdName.ToUpper) Then
                            SetCellValue(m_Row, gPos1XColumn, tempPosX)
                            SetCellValue(m_Row, gPos1YColumn, tempPosY)
                            SetCellValue(m_Row, gPos1ZColumn, tempPosZ)
                            SelectCell(m_Row, gCommandNameColumn)
                        End If
                    ElseIf status = 1 Then      'Ok
                        SelectCell(m_Row + 1, gCommandNameColumn)
                    End If
                    DisableCoordinateUpdateInSpreadsheet() 'for delete
                    UpdateSpreadsheet()
                    DeletingRowFromExcel = False
                    ClearAndDisplayIndicator()
                End If
            End If
        Catch ex As SystemException
            'kr should i put here
            ExceptionDisplay(ex)
            DeletingRowFromExcel = False
        End Try
        TraceGCCollect()

    End Sub

    Private Sub QCFormResponse()
        EnableCalibButtons()
        Dim status As Integer
        status = Vision.GetQCStatus()
        DisableCoordinateUpdateInSpreadsheet()
        EnableProgrammingBrightnessToggle()
        EnableTeachingButtons()
        DisableElementsCommandBlockButton(gSeperatorCmdIndex)
        m_ProgrammingStateFlag = False
        'Vision.FrmVision.EnableClickToMoveMode(False)
        If status = 2 Then 'Cancel
            SelectCell(m_Row, 1)
            Vision.SetQCReset()
        ElseIf status = 1 Then                         'Ok
            SetQCSettings()
            SelectCell(m_Row + 1, 1)
        End If
        If status = 2 Then
            Dim name As String = GetCellValue(m_Row, gCommandNameColumn)
            If Not ("GLOBALQC" = name.ToUpper) Then
                SetCellValue(m_Row, gPos1XColumn, tempPosX)
                SetCellValue(m_Row, gPos1YColumn, tempPosY)
                SetCellValue(m_Row, gPos1ZColumn, tempPosZ)
                SelectCell(m_Row, gCommandNameColumn)
            End If
        ElseIf status = 1 Then      'Ok
            If ("GLOBALQC" = name.ToUpper) Then
                SetCellValue(m_Row, gPos1XColumn, Nothing)
                SetCellValue(m_Row, gPos1YColumn, Nothing)
                SetCellValue(m_Row, gPos1ZColumn, Nothing)
                SelectCell(m_Row, gCommandNameColumn)
            End If
            SelectCell(m_Row, gCommandNameColumn)
        End If

        MenuEditCopy.Enabled = True
        If "" <> CopiedSheetName Then
            MenuEditPaste.Enabled = True
        End If
        MenuEditCut.Enabled = True
        MenuEditUndo.Enabled = True
        MenuEditRedo.Enabled = False
        SaveProgram.UnSave = True

        UndoData_Logging(0)
        m_EditStateFlag = False
        Spreadsheet_CheckForWithinLinkRange(True)
        DisableTeachingToolbarOKButton()
        EnableTeachingToolbarCancelButton()
        DisableEditingToolbar()
        DisableElementsCommandBlockButton(gSeperatorCmdIndex)
        DisableElementsCommandBlockButton(gOffsetCmdIndex)
        DisplaySpreadsheetTabs()
        DisableCoordinateUpdateInSpreadsheet() 'for delete
        EnableEditingToolbarSwitchButton()
        UpdateSpreadsheet()
        DeletingRowFromExcel = False
        ClearAndDisplayIndicator()
    End Sub

    Private Sub AddQCCommandFormResponse()
        EnableCalibButtons()
        Dim status As Integer = Vision.GetQCStatus()
        'Vision.FrmVision.EnableClickToMoveMode(False)
        If status = 2 Then 'Cancel
            DelayForRowDelete()
            DisableCoordinateUpdateInSpreadsheet()
            DeleteRow(m_Row)
            UpdateSpreadsheet()
            DeletingRowFromExcel = False
            DeletingRowFinished = False
        ElseIf status = 1 Then 'Ok
            If Vision.GetIsGlobalQC() Then
                DeleteRow(m_Row)
                Me.AddGlobalQCToTop()
                DisableCoordinateUpdateInSpreadsheet()
                SetGlobalQCSettings()
                IDS.Data.Hardware.Camera.DotQCEnable = True
                SetCellValue(1, gPos1XColumn, Nothing)
                SetCellValue(1, gPos1YColumn, Nothing)
                SetCellValue(1, gPos1ZColumn, Nothing)
                ' DisableElementsCommandBlockButton(gGlobalQCCmdIndex)
            Else
                SetQCSettings()
            End If
            DisableCoordinateUpdateInSpreadsheet()
            SelectCell(m_Row + 1, 1)
        End If

        EnableProgrammingBrightnessToggle()
        EnableTeachModeSwitching()
        EnableTeachingButtons()
        DisableElementsCommandBlockButton(gSeperatorCmdIndex)
        DisableElementsCommandBlockButton(gOffsetCmdIndex)
        DisableTeachingToolbarOKButton()
        EnableTeachingToolbarCancelButton()
        DisplaySpreadsheetTabs()
        m_ProgrammingStateFlag = False
        ClearAndDisplayIndicator()
        AxSpreadsheetProgramming.Enabled = True
    End Sub

    Private Sub EditFiducialCommandFormResponse()
        Dim status As Integer = Vision.GetFiducialStatus()
        'Vision.FrmVision.EnableClickToMoveMode(False)
        Dim fidname As String
        If status = 1 Then
            fidname = Vision.GetFiducialFilename()
            Dim Brightness As Integer = Vision.GetFiducialBrightness
            'SetCellValue(m_Row, gFid1Column, fidname)
            'SetCellValue(m_Row, gBrightnessColumn, Brightness)
            If m_TeachStepNumber = 1 Then
                SetCellValue(m_Row, gFid1Column, fidname)
                SetCellValue(m_Row, gBrightnessColumn, Brightness)
                'LabelMessage("Please confirm the 2nd Fiducial pt")
            Else
                SetCellValue(m_Row, gFid2Column, fidname)
                SetCellValue(m_Row, gThresholdColumn, Brightness)
                EnableProgrammingBrightnessToggle()
            End If
        ElseIf 2 = status Then  'Cancel command
            DisableCoordinateUpdateInSpreadsheet()
            ToggleButtonsForTeachingStop()
            DisableElementsCommandBlockButton(gSeperatorCmdIndex)
            EnableProgrammingBrightnessToggle()
            If m_TeachStepNumber = 1 Then
                SetCellValue(m_Row, gPos1XColumn, tempPosX)
                SetCellValue(m_Row, gPos1YColumn, tempPosY)
                SetCellValue(m_Row, gPos1ZColumn, tempPosZ)
                SelectCell(m_Row, gCommandNameColumn)
            ElseIf m_TeachStepNumber = 2 Then
                SetCellValue(m_Row, gPos2XColumn, tempPosX)
                SetCellValue(m_Row, gPos2YColumn, tempPosY)
                SetCellValue(m_Row, gPos2ZColumn, tempPosZ)
                SelectCell(m_Row, gCommandNameColumn)
            End If
            m_TeachStepNumber = 0
        ElseIf 3 = status Then
            fidname = Vision.GetFiducialFilename()
            SetCellValue(m_Row, gFid1Column, fidname)
        End If
        EnableEditingToolbarSwitchButton()
        EnableCalibButtons()
    End Sub

    Private Sub AddFiducialCommandFormResponse()
        Dim fidname As String
        Dim status As Integer = Vision.GetFiducialStatus()
        'Vision.FrmVision.EnableClickToMoveMode(False)
        If status = 1 Then
            fidname = Vision.GetFiducialFilename()
            Dim Brightness1 As Integer = Vision.GetFiducialBrightness
            'SetCellValue(m_Row, gFid1Column, fidname)
            '
            If m_TeachStepNumber = 1 Then
                SetCellValue(m_Row, gFid1Column, fidname)
                SetCellValue(m_Row, gBrightnessColumn, Brightness1)
                LabelMessage("Please confirm the 2nd Fiducial pt")
            Else
                SetCellValue(m_Row, gFid2Column, fidname)
                SetCellValue(m_Row, gThresholdColumn, Brightness1)
            End If
        ElseIf 2 = status Then  'Cancel command
            DelayForRowDelete()
            DisableCoordinateUpdateInSpreadsheet()
            DeleteRow(m_Row)
            SelectCell(m_Row, gCommandNameColumn)
            DeletingRowFromExcel = False
            DeletingRowFinished = False
            m_TeachStepNumber = 0
            ToggleButtonsForTeachingStop()
            DisableElementsCommandBlockButton(gSeperatorCmdIndex)
            DisableElementsCommandBlockButton(gOffsetCmdIndex)
            Spreadsheet_CheckForWithinLinkRange(True)
            DisplaySpreadsheetTabs()
            m_ProgrammingStateFlag = False
            LabelMessage("")
            AxSpreadsheetProgramming.Enabled = True
            EnableCalibButtons()
        ElseIf 3 = status Then
            DisableCoordinateUpdateInSpreadsheet()
            m_TeachStepNumber = 0
            SetCellValue(m_Row, gPos2XColumn, GetCellValue(m_Row, gPos1XColumn))
            SetCellValue(m_Row, gPos2YColumn, GetCellValue(m_Row, gPos1YColumn))
            cell1 = GetCell(m_Row, gPos2XColumn)
            cell2 = GetCell(m_Row, gPos2ZColumn)
            m_Execution.m_Pattern.Spreadsheet_CellGrey(cell1, cell2, True, AxSpreadsheetProgramming)
            ToggleButtonsForTeachingStop()
            DisableElementsCommandBlockButton(gSeperatorCmdIndex)
            DisableElementsCommandBlockButton(gOffsetCmdIndex)
            Spreadsheet_CheckForWithinLinkRange(True)
            DisplaySpreadsheetTabs()
            m_ProgrammingStateFlag = False
        End If

        If m_TeachStepNumber = 2 Then
            DisableCoordinateUpdateInSpreadsheet()
            DisplaySpreadsheetTabs()
            SelectCell(m_Row + 1, 1)
            m_TeachStepNumber = 1
            DisableTeachingToolbar()
        Else
            If (m_ProgrammingStateFlag = True) Then
                m_TeachStepNumber = 2
            End If
        End If
        EnableProgrammingBrightnessToggle()
    End Sub

    Private Sub AddChipEdgeCommandFormResponse()
        'Dim x, y As Double

        Dim status As Integer = Vision.GetChipEdgeStatus
        'Vision.FrmVision.EnableClickToMoveMode(False)
        ''moverobot
        'Vision.RobotMotionOffset(x, y)
        'Dim pos(2) As Double
        'pos(0) = x
        'pos(1) = -y
        'pos(2) = 0
        'If Not (x = 0 And y = 0) Then
        '    m_Tri.Set_XY_Speed(IDS.Data.Hardware.Gantry.ElementXYSpeed)
        '    m_Tri.MoveRelative_XY(pos)
        'End If

        'If status = 3 Then
        '    status = 0 'reset status after 5 points being reset
        'End If
        'While status = 0
        '    TraceDoEvents()
        '    status = Vision.GetChipEdgeStatus()
        'End While
        EnableCalibButtons()
        'DelayForRowDelete()
        DisableCoordinateUpdateInSpreadsheet()

        If status = 2 Then
            DeleteRow(m_Row)
            UpdateSpreadsheet()
            DeletingRowFromExcel = False
            DeletingRowFinished = False
            AxSpreadsheetProgramming.Enabled = True
        ElseIf status = 1 Then 'chipedge finished settings
            SetChipEdgeSettings()
            ElementsCommandBlock.Enabled = True
            ReferenceCommandBlock.Enabled = True
            DisplaySpreadsheetTabs()
            SelectCell(m_Row + 1, 1)
        End If
        AxSpreadsheetProgramming.Enabled = True
        ToggleButtonsForTeachingStop()
        ClearAndDisplayIndicator()
        EnableTeachingButtons()
        EnableTeachModeSwitching()
        m_ProgrammingStateFlag = False
        m_EditStateFlag = False
        EnableProgrammingBrightnessToggle()
        DisableElementsCommandBlockButton(gSeperatorCmdIndex)
        DisableElementsCommandBlockButton(gOffsetCmdIndex)
        DisableTeachingToolbarOKButton()
        EnableTeachingToolbarCancelButton()
        DisplaySpreadsheetTabs()

    End Sub

    Private Sub EditChipEdgeCommandFormResponse()
        Dim x, y As Double
        Dim pos(3) As Double
        Dim status As Integer
        Dim vPara As DLL_Export_Device_Vision.ChipEdgePoints.ChipEdgeParam
        'Do
        '    While Not (Vision.RobotMotionOffset(x, y) = True Or Vision.GetChipEdgeStatus = 2 Or Vision.GetChipEdgeStatus = 1)
        '        TraceDoEvents()
        '    End While

        '    'moverobot
        '    pos(0) = x
        '    pos(1) = -y
        '    pos(2) = 0

        '    m_Tri.Set_XY_Speed(IDS.Data.Hardware.Gantry.ElementXYSpeed)
        '    m_Tri.MoveRelative_XY(pos)

        '    If status = 3 Then
        '        status = 0 'reset status after 5 points being reset
        '    End If

        '    While status = 0
        '        TraceDoEvents()
        '        status = Vision.GetChipEdgeStatus()
        '    End While

        'Loop While status = 3 'status 3= reset 5 points
        status = Vision.GetChipEdgeStatus()
        'Vision.FrmVision.EnableClickToMoveMode(False)
        'moverobot
        'Vision.RobotMotionOffset(x, y)
        'pos(0) = x
        'pos(1) = -y
        'pos(2) = 0
        'If Not (x = 0 And y = 0) Then
        '    m_Tri.Set_XY_Speed(IDS.Data.Hardware.Gantry.ElementXYSpeed)
        '    m_Tri.MoveRelative_XY(pos)
        'End If
        EnableCalibButtons()
        DisableCoordinateUpdateInSpreadsheet()
        ElementsCommandBlock.Enabled = True
        ReferenceCommandBlock.Enabled = True
        EnableTeachingButtons()
        EnableProgrammingBrightnessToggle()
        DisableElementsCommandBlockButton(gSeperatorCmdIndex)
        m_ProgrammingStateFlag = False
        m_EditStateFlag = False
        EnableEditingToolbarSwitchButton()
        If status = 2 Then
            UpdateSpreadsheet()
            'SelectCell(m_Row, gCommandNameColumn)

            SetCellValue(m_Row, gPos1XColumn, tempPosX)
            SetCellValue(m_Row, gPos1YColumn, tempPosY)
            SetCellValue(m_Row, gPos1ZColumn, tempPosZ)
            SelectCell(m_Row, gCommandNameColumn)
        ElseIf status = 1 Then
            Dim bb As Boolean
            Vision.GetChipEdgeParameters(vPara)
            Vision.SetCEReset()

            SetCellValue(m_Row, gEdgeClearColumn, vPara._EdgeClearance)
            SetCellValue(m_Row, gBrightnessColumn, vPara._Brightness)
            SetCellValue(m_Row, gCheckBoxColumn, vPara._CheckBox_ChipRec_Enable)
            SetCellValue(m_Row, gCwCCwColumn, vPara._Cw_CCw)
            SetCellValue(m_Row, gDispModelColumn, vPara._DispenseModel)
            SetCellValue(m_Row, gInOutColumn, vPara._Inside_out)
            SetCellValue(m_Row, gMainEdgeColumn, vPara._MainEdge)
            SetCellValue(m_Row, gPointX1Column, vPara._PointX1)
            SetCellValue(m_Row, gPointX2Column, vPara._PointX2)
            SetCellValue(m_Row, gPointX3Column, vPara._PointX3)
            SetCellValue(m_Row, gPointX4Column, vPara._PointX4)
            SetCellValue(m_Row, gPointX5Column, vPara._PointX5)
            SetCellValue(m_Row, gPointY1Column, vPara._PointY1)
            SetCellValue(m_Row, gPointY2Column, vPara._PointY2)
            SetCellValue(m_Row, gPointY3Column, vPara._PointY3)
            SetCellValue(m_Row, gPointY4Column, vPara._PointY4)
            SetCellValue(m_Row, gPointY5Column, vPara._PointY5)
            SetCellValue(m_Row, gPosColumn, vPara._Pos)
            SetCellValue(m_Row, gPosXColumn, vPara._PosX)
            SetCellValue(m_Row, gPosYColumn, vPara._PosY)
            SetCellValue(m_Row, gROIColumn, vPara._ROI)
            SetCellValue(m_Row, gRotColumn, vPara._Rot)
            SetCellValue(m_Row, gSizeColumn, vPara._Size)
            SetCellValue(m_Row, gSizeXColumn, vPara._SizeX)
            SetCellValue(m_Row, gSizeYColumn, vPara._SizeY)
            SetCellValue(m_Row, gThresholdColumn, vPara._Threshold)
            SetCellValue(m_Row, gVerticalColumn, vPara._Vertical)
            SetCellValue(m_Row, gDurationColumn, vPara._DotDispensingDuration)
            SetCellValue(m_Row, gCompactnessColumn, vPara._Contrast)
            cell1 = GetCell(m_Row, gPos1XColumn)
            cell2 = GetCell(m_Row, gPos1ZColumn)
            m_Execution.m_Pattern.Spreadsheet_CellGrey(cell1, cell2, True, AxSpreadsheetProgramming)
            SelectCell(m_Row, 1)
        End If
    End Sub

    Private Sub AddArrayCommandFormResponse()
        Dim DlgReturn = ArrayDlg.DialogResult()
        'Dim PointX, PointY, PointZ As Double
        'Do
        '    TraceDoEvents()
        '    PointX = CDbl(GetCellValue(m_Row, gPos1XColumn))
        '    PointY = CDbl(GetCellValue(m_Row, gPos1YColumn))
        '    PointZ = CDbl(GetCellValue(m_Row, gPos1ZColumn))

        '    ArrayDlg.SetPoint(PointX, PointY, PointZ)
        '    DlgReturn = ArrayDlg.DialogResult()
        'Loop While Nothing = DlgReturn
        EnableCalibButtons()
        DisableCoordinateUpdateInSpreadsheet()

        If DialogResult.OK = DlgReturn Then
            'Get "Array" data succesfully.  Data will be used to generate "Array"
            LabelMessage("Array generating starts")
            Dim PatternLineRecord(3) As CIDSPattern.PatternRecord
            Dim arraydata As ArrayData
            arraydata.FirstX = 10.0
            Dim dotarray As New FormArraySetting(arraydata)
            dotarray.TopMost = True
            If dotarray.ShowDialog() = DialogResult.OK Then
                dotarray.GetArrayData(arraydata)

                Dim cmdTmpString As String = ArrayDlg.ArrayPara1.Name

                Spreadsheet_GetArrayRecord(PatternLineRecord(0), 1, ArrayDlg)
                Spreadsheet_GetArrayRecord(PatternLineRecord(1), 2, ArrayDlg)
                Spreadsheet_GetArrayRecord(PatternLineRecord(2), 3, ArrayDlg)

                Dim ActSheetName As String = GetActiveSheetName()
                Spreadsheet_GenerateArraySubSheetName(iSubSheetName, ActSheetName, cmdTmpString)

                Dim file As New CIDSFileHandler
                ClearRow(m_Row)

                SetCellValue(m_Row, gCommandNameColumn, "Array")
                SetCellValue(m_Row, gSubnameColumn, file.ExtOnlyFromFullPath(iSubSheetName))
                SelectCell(m_Row + 1, gCommandNameColumn)

                AxSpreadsheetProgramming.Sheets.Add.Name = iSubSheetName
                AxSpreadsheetProgramming.ActiveWindow.FreezePanes = False
                AxSpreadsheetProgramming.Worksheets(iSubSheetName).Range("B1:B1").Select()
                AxSpreadsheetProgramming.ActiveWindow.FreezePanes = True

                'Build Sub array sheet
                Spreadsheet_AddSheetRecord(PatternLineRecord(0), _
                    PatternLineRecord(1), PatternLineRecord(2), arraydata)

                'Load sub and array data within the sub
                'Spreadsheet_AddSubandArrayInSub(PatternLineRecord(0).pPara.DispenseFlag)
                'yy
                If Not (ArrayDlg.ArrayPara1.DispenseFlag.ToUpper = "ON" Or ArrayDlg.ArrayPara1.DispenseFlag.ToUpper = "OFF ") Then
                    Dim file1 As String = ArrayDlg.ArrayPara1.DispenseFlag
                    If 0 = m_Execution.m_Pattern.Spreadsheet_CheckSubsheetExist( _
                AxSpreadsheetProgramming, m_Execution.m_File.NameOnlyFromFullPath(file1)) Then
                        m_Execution.m_Pattern.LoadTxtPatternPara(AxSpreadsheetProgramming, file1, 1, 0, False)
                    End If
                End If

                EnableTeachingButtons()
                DisableElementsCommandBlockButton(gOffsetCmdIndex)
                DisableTeachingToolbarOKButton()
                EnableTeachingToolbarCancelButton()
            Else
                DelayForRowDelete()
                EnableTeachingButtons()
                DisableElementsCommandBlockButton(gOffsetCmdIndex)
                DisableTeachingToolbarOKButton()
                EnableTeachingToolbarCancelButton()
                TraceDoEvents()
                DeleteRow(m_Row)
                DeletingRowFromExcel = False
                DeletingRowFinished = False
            End If
        Else
            DelayForRowDelete()
            EnableTeachingButtons()
            DisableElementsCommandBlockButton(gOffsetCmdIndex)
            DisableTeachingToolbarOKButton()
            EnableTeachingToolbarCancelButton()
            TraceDoEvents()
            DeleteRow(m_Row)
            DeletingRowFromExcel = False
            DeletingRowFinished = False

        End If
        AxSpreadsheetProgramming.Enabled = True
        m_ProgrammingStateFlag = False

        CBExpandSpreadsheet.Enabled = True
    End Sub

#Region "teaching"
    Public Sub ConfirmPurge()
        DisableCoordinateUpdateInSpreadsheet()
        DisplaySpreadsheetTabs()
        SelectCell(m_Row + 1, 1)
        m_Row = GetActiveCellRow()
        DisableTeachingToolbar()
        m_ProgrammingStateFlag = False
    End Sub
    Public Sub ConfirmReject()
        EnableCoordinateUpdateInSpreadsheet()
        DisableProgrammingBrightnessToggle()
        Vision.IDSV_Form_RM(ValueBrightness.Value)
        Dim status As Integer = 0
        Dim x, y As Double
        While status = 0
            TraceDoEvents()
            status = Vision.GetRMStatus
        End While
        If status = 2 Then
            DelayForRowDelete()
            DisableCoordinateUpdateInSpreadsheet()
            DisplaySpreadsheetTabs()
            DeleteRow(m_Row)
            SelectCell(m_Row, gCommandNameColumn)
            DeletingRowFromExcel = False
            DeletingRowFinished = False
        ElseIf status = 1 Then      'Ok
            Dim bb As Boolean
            Dim vpara As DLL_Export_Device_Vision.RejectPoint.RMParam
            Vision.GetRMParameters(vpara)
            Vision.SetRMReset()
            DisableCoordinateUpdateInSpreadsheet()
            SetCellValue(m_Row, gAcceptRatioCoulumn, vpara._AcceptRatio)
            SetCellValue(m_Row, gBinarizedColumn, vpara._Binarized)
            SetCellValue(m_Row, gBlackWithoutRMCoulumn, vpara._BlackWithoutRM)
            SetCellValue(m_Row, gBlackWithRMCoulumn, vpara._BlackWithRM)
            SetCellValue(m_Row, gBrightnessColumn, vpara._Brightness)
            SetCellValue(m_Row, gMRegionXColumn, vpara._MRegionX)
            SetCellValue(m_Row, gMRegionYColumn, vpara._MRegionY)
            SetCellValue(m_Row, gMROIxColumn, vpara._MROIx)
            SetCellValue(m_Row, gMROIyColumn, vpara._MROIy)
            SetCellValue(m_Row, gWhiteWithoutRMCoulumn, vpara._WhiteWithoutRM)
            SetCellValue(m_Row, gWhiteWithRMCoulumn, vpara._WhiteWithRM)
            SetCellValue(m_Row, gWoRMCoulumn, vpara._WoRM)
            SetCellValue(m_Row, gWRMCoulumn, vpara._WRM)
            cell1 = GetCell(m_Row, gPos1XColumn)
            cell2 = GetCell(m_Row, gPos1ZColumn)
            m_Execution.m_Pattern.Spreadsheet_CellGrey(cell1, cell2, True, AxSpreadsheetProgramming)
            DisplaySpreadsheetTabs()
            SelectCell(m_Row + 1, gCommandNameColumn)
        End If
        EnableProgrammingBrightnessToggle()
        m_ProgrammingStateFlag = False
        EnableTeachingButtons()
        DisableElementsCommandBlockButton(gOffsetCmdIndex)
        DisableTeachingToolbarOKButton()
        EnableTeachingToolbarCancelButton()
    End Sub
    Public Sub ConfirmHeight()
        DisableCoordinateUpdateInSpreadsheet()
        DisplaySpreadsheetTabs()
        cell1 = GetCell(m_Row, gPos1XColumn)
        cell2 = GetCell(m_Row, gPos1ZColumn)
        Dim x1 As Double = GetCellValue(m_Row, gPos1XColumn)
        Dim y1 As Double = GetCellValue(m_Row, gPos1YColumn)

        ToggleButtonsForTeachingStop()
        DisableElementsCommandBlockButton(gSeperatorCmdIndex)

        m_Execution.m_Pattern.Spreadsheet_CellGrey(cell1, cell2, True, AxSpreadsheetProgramming)
        Dim posFrom(3), posTo(3), posOffset As Double
        posFrom(0) = GetCellValue(m_Row, gPos1XColumn)
        posFrom(1) = GetCellValue(m_Row, gPos1YColumn)
        posFrom(2) = 0
        posTo(0) = posFrom(0) - gLaserOffX
        posTo(1) = posFrom(1) - gLaserOffY
        posTo(2) = posFrom(2)

        LockMovementButtons()
        MoveToSpreadsheetPoint(posTo, "Vision")

        OnLaser()
        Dim rtn As Integer = Laser.WaitForReadingToStabilize()
        OffLaser()
        If rtn Then                     'No laser readout error
            posOffset = Laser.LASER_Reading - IDS.Data.Hardware.HeightSensor.Laser.HeightReference
            LaserHeightOffsetZ = posOffset
            posOffset = CInt(posOffset * 1000) / 1000
            SetCellValue(m_Row, gPos1ZColumn, posOffset)
        Else
            MessageBox.Show("laser sensor out of range")
            m_TeachStepNumber = 0
            ToggleButtonsForTeachingStop()
            DisableElementsCommandBlockButton(gSeperatorCmdIndex)
            DisableElementsCommandBlockButton(gOffsetCmdIndex)
            DeleteRow(m_Row)
        End If
        posFrom(0) = posFrom(0)
        posFrom(1) = posFrom(1)
        posFrom(2) = posFrom(2)

        MoveToSpreadsheetPoint(posFrom, "Vision")
        ChangeButtonState("Idle")
        UnlockMovementButtons()

        SelectCell(m_Row + 1, gCommandNameColumn)
        m_ProgrammingStateFlag = False
    End Sub
    Public Sub ConfirmSubPattern()
        DisableCoordinateUpdateInSpreadsheet()
        DisplaySpreadsheetTabs()
        cell1 = GetCell(m_Row, gPos1XColumn)
        cell2 = GetCell(m_Row, gPos1ZColumn)
        Dim x1 As Double = GetCellValue(m_Row, gPos1XColumn)
        Dim y1 As Double = GetCellValue(m_Row, gPos1YColumn)
        DisableTeachingButtons()
        m_Execution.m_Pattern.Spreadsheet_CellGrey(cell1, cell2, True, AxSpreadsheetProgramming)
        SelectCell(m_Row + 1, gCommandNameColumn)
        m_ProgrammingStateFlag = False
    End Sub
    Public Sub ConfirmSinglePointElement()
        DisableCoordinateUpdateInSpreadsheet()
        DisplaySpreadsheetTabs()
        cell1 = GetCell(m_Row, gPos1XColumn)
        cell2 = GetCell(m_Row, gPos1ZColumn)
        Dim x1 As Double = GetCellValue(m_Row, gPos1XColumn)
        Dim y1 As Double = GetCellValue(m_Row, gPos1YColumn)
        ToggleButtonsForTeachingStop()
        DisableElementsCommandBlockButton(gSeperatorCmdIndex)
        m_Execution.m_Pattern.Spreadsheet_CellGrey(cell1, cell2, True, AxSpreadsheetProgramming)
        SelectCell(m_Row + 1, gCommandNameColumn)
        m_ProgrammingStateFlag = False
    End Sub
#End Region

    Private Sub Confirm()
        Dim i As Integer
        Dim PatternLineRecord(3) As CIDSPattern.PatternRecord
        Dim cell1, cell2 As Object

        m_Row = GetActiveCellRow() 'Get current row number

        Try
            If type = Nothing Then Exit Sub
            Select Case (type.ToUpper)
                Case "PURGE", "CLEAN"
                    ConfirmPurge()
                    AxSpreadsheetProgramming.Enabled = True
                Case "REJECT"
                    ConfirmReject()
                    AxSpreadsheetProgramming.Enabled = True
                Case "HEIGHT"
                    ConfirmHeight()
                    AxSpreadsheetProgramming.Enabled = True
                Case "SUBPATTERN"
                    ConfirmSubPattern()
                    AxSpreadsheetProgramming.Enabled = True
                Case "DOT", "REFERENCE", "QC", "MOVE", "CHIPEDGE", "WAIT"
                    ConfirmSinglePointElement()
                    AxSpreadsheetProgramming.Enabled = True
                Case "LINE", "FIDUCIAL"
                    EnableCoordinateUpdateInSpreadsheet()
                    HideSpreadsheetTabs()

                    If m_TeachStepNumber = 1 Then
                        LabelMessage("Please confirm the 2nd pt")
                        cell1 = GetCell(m_Row, gPos1XColumn)
                        cell2 = GetCell(m_Row, gPos1ZColumn)
                        x1 = GetCellValue(m_Row, gPos1XColumn)
                        y1 = GetCellValue(m_Row, gPos1YColumn)
                        m_Execution.m_Pattern.Spreadsheet_CellGrey(cell1, cell2, True, AxSpreadsheetProgramming)
                        cell1 = GetCell(m_Row, gPos2XColumn)
                        cell2 = GetCell(m_Row, gPos2ZColumn)
                        SelectRange(cell1, cell2)
                        Spreadsheet_CheckForWithinLinkRange(False)
                        If m_InLinkRange Then
                            UpdateLinkForPreviousRow()
                        ElseIf type.ToUpper = "FIDUCIAL" Then
                            LabelMessage("Please confirm the 1st Fiducial pt")
                            IDS.Data.SaveData()
                            DisableProgrammingBrightnessToggle()
                            Vision.FiducialMark_form.FormCloseEvent = New DLL_Export_Device_Vision.FiducialForm.FormCloseDelegate(AddressOf Me.AddFiducialCommandFormResponse)
                            Vision.IDSV_Form_FI(1, ValueBrightness.Value)
                            Me.EnableClickToMove()
                            'Dim status As Integer = 0 'status: 1=ok, 2=cancel, 3=end with first fiducial
                            'While status = 0
                            '    TraceDoEvents()
                            '    status = Vision.GetFiducialStatus()
                            'End While
                            'Dim fidname As String
                            'If status = 1 Then
                            '    fidname = Vision.GetFiducialFilename()
                            '    Dim Brightness1 As Integer = Vision.GetFiducialBrightness
                            '    SetCellValue(m_Row, gFid1Column, fidname)
                            '    SetCellValue(m_Row, gBrightnessColumn, Brightness1)
                            '    LabelMessage("Please confirm the 2nd Fiducial pt")
                            'ElseIf 2 = status Then  'Cancel command
                            '    DelayForRowDelete()
                            '    DisableCoordinateUpdateInSpreadsheet()
                            '    DeleteRow(m_Row)
                            '    SelectCell(m_Row, gCommandNameColumn)
                            '    DeletingRowFromExcel = False
                            '    DeletingRowFinished = False
                            '    m_TeachStepNumber = 0
                            '    ToggleButtonsForTeachingStop()
                            '    DisableElementsCommandBlockButton(gSeperatorCmdIndex)
                            '    DisableElementsCommandBlockButton(gOffsetCmdIndex)
                            '    Spreadsheet_CheckForWithinLinkRange(True)
                            '    DisplaySpreadsheetTabs()
                            '    m_ProgrammingStateFlag = False
                            '    LabelMessage("")
                            'ElseIf 3 = status Then
                            '    DisableCoordinateUpdateInSpreadsheet()
                            '    m_TeachStepNumber = 0
                            '    SetCellValue(m_Row, gPos2XColumn, GetCellValue(m_Row, gPos1XColumn))
                            '    SetCellValue(m_Row, gPos2YColumn, GetCellValue(m_Row, gPos1YColumn))
                            '    cell1 = GetCell(m_Row, gPos2XColumn)
                            '    cell2 = GetCell(m_Row, gPos2ZColumn)
                            '    m_Execution.m_Pattern.Spreadsheet_CellGrey(cell1, cell2, True, AxSpreadsheetProgramming)
                            '    ToggleButtonsForTeachingStop()
                            '    DisableElementsCommandBlockButton(gSeperatorCmdIndex)
                            '    DisableElementsCommandBlockButton(gOffsetCmdIndex)
                            '    Spreadsheet_CheckForWithinLinkRange(True)
                            '    DisplaySpreadsheetTabs()
                            '    m_ProgrammingStateFlag = False
                            'End If
                            'EnableProgrammingBrightnessToggle()
                        End If

                    ElseIf m_TeachStepNumber = 2 Then
                        AxSpreadsheetProgramming.Enabled = True
                        m_ProgrammingStateFlag = False
                        cell1 = GetCell(m_Row, gPos2XColumn)
                        cell2 = GetCell(m_Row, gPos2ZColumn)
                        x2 = GetCellValue(m_Row, gPos2XColumn)
                        y2 = GetCellValue(m_Row, gPos2YColumn)
                        If type.ToUpper = "LINE" Then
                        ElseIf type.ToUpper = "FIDUCIAL" Then
                            DisableProgrammingBrightnessToggle()
                            Vision.FiducialMark_form.FormCloseEvent = New DLL_Export_Device_Vision.FiducialForm.FormCloseDelegate(AddressOf Me.AddFiducialCommandFormResponse)
                            Vision.IDSV_Form_FI(2, ValueBrightness.Value)
                            Me.EnableClickToMove()
                            'Dim status As Integer = 0
                            'While status = 0
                            '    TraceDoEvents()
                            '    status = Vision.GetFiducialStatus()
                            'End While
                            'Dim fidname As String
                            'If status = 1 Then
                            '    fidname = Vision.GetFiducialFilename()
                            '    SetCellValue(m_Row, gFid2Column, fidname)
                            '    Dim Brightness2 As Integer = Vision.GetFiducialBrightness
                            '    SetCellValue(m_Row, gThresholdColumn, Brightness2) 'For brightness fiducial no.2
                            'ElseIf 2 = status Then
                            '    'SetCellValue(m_Row, gPos2XColumn, GetCellValue(m_Row, gPos1XColumn))
                            '    'SetCellValue(m_Row, gPos2YColumn, GetCellValue(m_Row, gPos1YColumn))
                            'End If
                        End If
                        If Not type.ToUpper = "FIDUCIAL" Then
                            UpdateLinkForNextRow()
                            DisableCoordinateUpdateInSpreadsheet()
                            m_Execution.m_Pattern.Spreadsheet_CellGrey(cell1, cell2, True, AxSpreadsheetProgramming)
                            SelectCell(m_Row + 1, 1)
                        End If
                    End If

                    If Not type.ToUpper = "FIDUCIAL" Then
                        If m_TeachStepNumber = 2 Then
                            DisableCoordinateUpdateInSpreadsheet()
                            DisplaySpreadsheetTabs()
                            m_TeachStepNumber = 1
                            DisableTeachingToolbar()
                        Else
                            If (m_ProgrammingStateFlag = True) Then
                                m_TeachStepNumber = 2
                            End If
                        End If
                    End If

                    'If type.ToUpper = "FIDUCIAL" Then EnableProgrammingBrightnessToggle()

                Case "ARC", "CIRCLE", "RECTANGLE", "FILLCIRCLE", "FILLRECTANGLE", "ARRAY", "DOTARRAY"
                    EnableCoordinateUpdateInSpreadsheet()
                    HideSpreadsheetTabs()

                    If (m_TeachStepNumber = 1) Then
                        cell1 = GetCell(m_Row, gPos1XColumn)
                        cell2 = GetCell(m_Row, gPos1ZColumn)
                        x1 = GetCellValue(m_Row, gPos1XColumn)
                        y1 = GetCellValue(m_Row, gPos1YColumn)
                        m_Execution.m_Pattern.Spreadsheet_CellGrey(cell1, cell2, True, AxSpreadsheetProgramming)
                        cell1 = GetCell(m_Row, gPos2XColumn)
                        cell2 = GetCell(m_Row, gPos2ZColumn)
                        SelectRange(cell1, cell2)
                        m_TeachStepNumber = 2
                        UpdateLinkForPreviousRow()
                    ElseIf (m_TeachStepNumber = 2) Then
                        cell1 = GetCell(m_Row, gPos2XColumn)
                        cell2 = GetCell(m_Row, gPos2ZColumn)
                        x2 = GetCellValue(m_Row, gPos2XColumn)
                        y2 = GetCellValue(m_Row, gPos2YColumn)
                        m_TeachStepNumber = 3
                        m_Execution.m_Pattern.Spreadsheet_CellGrey(cell1, cell2, True, AxSpreadsheetProgramming)
                        cell1 = GetCell(m_Row, gPos3XColumn)
                        cell2 = GetCell(m_Row, gPos3ZColumn)
                        AxSpreadsheetProgramming.ActiveSheet.Range(cell1, cell2).Select()
                    ElseIf (m_TeachStepNumber = 3) Then
                        AxSpreadsheetProgramming.Enabled = True
                        m_ProgrammingStateFlag = False
                        cell1 = GetCell(m_Row, gPos3XColumn)
                        cell2 = GetCell(m_Row, gPos3ZColumn)
                        x3 = GetCellValue(m_Row, gPos3XColumn)
                        y3 = GetCellValue(m_Row, gPos3YColumn)
                        If type.ToUpper = "DOTARRAY" Then
                            Dim arraydata As ArrayData
                            arraydata.FirstX = 10.0
                            Dim dotarray As New FormArraySetting(arraydata)
                            If dotarray.ShowDialog() = DialogResult.OK Then
                                dotarray.GetArrayData(arraydata)
                                SetCellValue(m_Row, gDotArrayRowsColumn, arraydata.RowNo.ToString + "x" + arraydata.ColumNo.ToString)
                            Else
                                MyMsgBox("Setting dot array rows and numbers to 2x2.")
                                SetCellValue(m_Row, gDotArrayRowsColumn, "2x2")
                            End If
                        End If
                        DisableCoordinateUpdateInSpreadsheet()
                        m_TeachStepNumber = 1
                        DisableTeachingToolbar()
                        UpdateLinkForNextRow()
                        m_Execution.m_Pattern.Spreadsheet_CellGrey(cell1, cell2, True, AxSpreadsheetProgramming)
                        SelectCell(m_Row + 1, 1)
                    End If

                    If m_TeachStepNumber = 1 Then
                        DisplaySpreadsheetTabs()
                        If type = "Array" Then
                            EnableElementsCommandBlock()
                            DisableElementsCommandBlockButton(gSeperatorCmdIndex)
                        End If
                    End If
            End Select

            If type.ToUpper = "SUBPATTERN" And m_TeachStepNumber = 1 And SheetRowSelected = False Then

                Dim DialogPreview As New FormSelectPatternFile

                DialogPreview.ShowDialog()
                Dim file As String = DialogPreview.Path

                If Nothing = file Then
                    'Clear rows based on the rejected input.  Array of Sub needs to celar more rows.

                    DeleteRow(m_Row)
                    SelectCell(m_Row, gCommandNameColumn)

                    'Enable cmd button if it is rejected
                    ToggleButtonsForTeachingStop()
                    DisableElementsCommandBlockButton(gSeperatorCmdIndex)

                    type = "Dot"

                Else
                    SetCellValue(m_Row, gSubnameColumn, file)
                    'Get current sheet name and row number
                    Dim TmpSheetName As String = GetActiveSheetName()
                    Dim TmpRowNumber As Integer = m_Row

                    If 0 = m_Execution.m_Pattern.Spreadsheet_CheckSubsheetExist( _
                        AxSpreadsheetProgramming, m_Execution.m_File.NameOnlyFromFullPath(file)) Then
                        m_Execution.m_Pattern.LoadTxtPatternPara(AxSpreadsheetProgramming, file, 1, 0, False)
                    End If
                    'Restore to the previous sheet and rows
                    m_Row = TmpRowNumber
                    AxSpreadsheetProgramming.Worksheets(TmpSheetName).Activate()
                    SelectCell(m_Row + 1, gCommandNameColumn)
                    ToggleButtonsForTeachingStop()
                    DisableElementsCommandBlockButton(gSeperatorCmdIndex)

                End If
            End If
            SaveProgram.UnSave = True
        Catch ex As SystemException
            ExceptionDisplay(ex)
        End Try
        TraceGCCollect()
    End Sub

    Private Sub EditOrCancel(ByVal editcancelFlag As Integer)
        Try
            Dim ran As String = "C" + m_Row.ToString + ":E" + m_Row.ToString
            Dim backup(3) As Object
            backup(0) = GetCellValue(m_Row, 3)
            backup(1) = GetCellValue(m_Row, 4)
            backup(2) = GetCellValue(m_Row, 5)
            If editcancelFlag = 1 Then
                'm_row = getactivecellrow
                ElementsCommandBlock.Enabled = False
                ReferenceCommandBlock.Enabled = False
                EnableCoordinateUpdateInSpreadsheet()

                LabelMessage("Move robot to teach point or keyin ponit")
                ToolBarSwitch("YesNo")
                AxSpreadsheetProgramming.ActiveSheet.Range(ran).Clear()
                SetCellValue(m_Row, 3, backup(0))
                SetCellValue(m_Row, 4, backup(1))
                SetCellValue(m_Row, 5, backup(2))

            ElseIf editcancelFlag = 0 Then
                DisableCoordinateUpdateInSpreadsheet()

                ClearRow(m_Row)
                UpdateSpreadsheet()
            End If

        Catch ex As SystemException
            ExceptionDisplay(ex)
        End Try
    End Sub

    Public Sub ToolBarSwitch(ByVal toggle_str As String)
        If toggle_str = "YesNo" Then
            LabelMessage("Move robot to teach point.")
        Else
            LabelMessage("Move robot or keyin coordinates.")
        End If
        EnableCoordinateUpdateInSpreadsheet()
    End Sub

#End Region

#Region "Others"
    Private Sub OptimizePath_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OptimizePath.Click

        m_Execution.m_Command.SetOptimizFlag(1)
        Dim m_PatternList As New ArrayList 'kr
        Dim m_DispenseList As New ArrayList 'kr
        Dim comP As New IDSPattnCompiler(m_PatternList) 'kr
        Dim loadSheet As New CIDSPatternLoader(AxSpreadsheetProgramming) 'kr
        UpdateSpreadsheet()
        Dim rtn As Integer = m_Execution.m_Command.ReadPattern(AxSpreadsheetProgramming)
        'comP.Compile(m_DispenseList) 'kr

        Dim optm As ArrayList
        Dim row As Integer

        If rtn = 0 Then
            optm = m_Execution.m_Command.PattenList
            row = optm.Count
        Else
            LabelMessage("Empty sheet or data error!")
            Exit Sub
        End If

        Dim folder As String = gPatternFileName
        folder = folder.TrimEnd("\") + "_optm"
        If System.IO.File.Exists(folder) Then
            'Project name is the same with an existing filename
            'The file will be deleted
            System.IO.File.Delete(folder)
            'Create folder for the project
        End If

        System.IO.Directory.CreateDirectory(folder)

        'gPatternFileName = folder + "\" + File


        Dim FileName As String '= gPatternFileName
        FileName = folder + "\" + m_Execution.m_File.NameOnlyFromFullPath(folder)

        'lsgoh
        ' get default value from the default pat file

        IDS.Data.ParameterID.RecordID = "FactoryDefault"
        IDSData.Admin.Folder.FileExtension = "Pat"
        IDSData.Admin.Folder.PatternPath = "C:\IDS\Pattern_Dir"
        IDS.Data.OpenData()
        IDS.Data.SavePathFileData(FileName + ".pat")

        Try
            If FileName <> Nothing Then
                m_Execution.m_Pattern.SavePatternParaOpt(optm, FileName + ".txt", False)
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
        LabelMessage("Dot optimisation file saved!")
        TraceGCCollect()

    End Sub

    Private Sub ButtonPro_LEdit_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        IDS.Data.SaveData()
    End Sub

    Private Sub ButtonPro_REdit_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        IDS.Data.SaveData()
    End Sub

#End Region

    Private Sub MenuItem33_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem33.Click
        '''''''''''''''''
        '   Offset      '
        '''''''''''''''''
        TeachElementCommand("Offset")
    End Sub

    Private Sub MenuItem32_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem32.Click
        '''''''''''''''''''''
        '   SubPattern      '
        '''''''''''''''''''''
        TeachElementCommand("SubPattern")
    End Sub

    Private Sub MenuItem29_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem29.Click
        '''''''''''''''''
        '   SetIO       '
        '''''''''''''''''
        TeachElementCommand("SetIO")
    End Sub

    Private Sub MenuItem30_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem30.Click
        '''''''''''''''''
        '   ResetIO     '
        '''''''''''''''''
        Dim Respond As DialogResult = MessageBox.Show("Do You want to initialize the IO bits?", "", MessageBoxButtons.YesNo)

        If Respond = DialogResult.Yes Then
            'Initialize trio controller IO 
            Dim I As Integer = 0
            m_Tri.SetAllDIOsOff()

            'Initialize PCIO card
            For I = 0 To 7   'iterate thru the io bits
                IDS.Devices.DIO.SetOutputBit(0, I, False)
                IDS.Devices.DIO.SetOutputBit(1, I, False)
            Next

            'Initialize CAN IO 
            For I = 32 To 47    'iterate thru the io bits
                MySleep(10)
                m_Tri.SetDIOs(1, I, False)
            Next
        End If
    End Sub

    Private Sub MenuItem62_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem62.Click
        TeachElementCommand("Measurement")
    End Sub

#Region "thermal"

    ' reads temperature from heater

    Dim temp_num As Integer = 0
    Dim alarm_num As Integer = 0
    Dim switching As Boolean = True

    Public Sub TimerForUpdate_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TimerForUpdate.Tick

        'If switching = True Then
        '    If IDS.Devices.Thermal.FrmThermal.command = "received" Then
        '        temp_num += 1
        '        If temp_num = 6 Then
        '            temp_num = 1
        '        End If
        '        If MyHeaterSettings.Heater_Enabled(temp_num - 1) = True Then
        '            IDS.Devices.Thermal.FrmThermal.command = "receiving"
        '            IDS.Devices.Thermal.FrmThermal.Get_Temperature(temp_num)
        '        End If
        '        switching = False
        '    End If
        'Else
        '    If IDS.Devices.Thermal.FrmThermal.command = "received" Then
        '        alarm_num += 1
        '        If alarm_num = 6 Then
        '            alarm_num = 1
        '        End If
        '        If MyHeaterSettings.Heater_Enabled(alarm_num - 1) = True Then
        '            IDS.Devices.Thermal.FrmThermal.command = "receiving"
        '            IDS.Devices.Thermal.FrmThermal.Get_Alarm_Status(alarm_num)
        '        End If
        '        switching = True
        '    End If
        'End If

        'IDS.Devices.Thermal.FrmThermal.check_individual_reading()

        'LB_NeedleHeater.Text = IDS.Devices.Thermal.FrmThermal.Temperature(0)
        'Production.Needle.Text = IDS.Devices.Thermal.FrmThermal.Temperature(0)

        'LB_SyrHeater.Text = IDS.Devices.Thermal.FrmThermal.Temperature(1)
        'Production.Syringe.Text = IDS.Devices.Thermal.FrmThermal.Temperature(1)

        'LB_PreHeater.Text = IDS.Devices.Thermal.FrmThermal.Temperature(2)
        'Production.Station1.Text = IDS.Devices.Thermal.FrmThermal.Temperature(2)

        'LB_DispHeater.Text = IDS.Devices.Thermal.FrmThermal.Temperature(3)
        'Production.Station2.Text = IDS.Devices.Thermal.FrmThermal.Temperature(3)

        'LB_PostHeater.Text = IDS.Devices.Thermal.FrmThermal.Temperature(4)
        'Production.Station3.Text = IDS.Devices.Thermal.FrmThermal.Temperature(4)

        'If IDS.Devices.Thermal.FrmThermal.switched_off <> 0 Then
        '    Select Case (IDS.Devices.Thermal.FrmThermal.switched_off)
        '        Case 1
        '            HeaterControl0.Text = "Heater Off"
        '        Case 2
        '            HeaterControl1.Text = "Heater Off"
        '        Case 3
        '            HeaterControl2.Text = "Heater Off"
        '        Case 4
        '            HeaterControl3.Text = "Heater Off"
        '        Case 5
        '            HeaterControl4.Text = "Heater Off"
        '    End Select
        '    IDS.Devices.Thermal.FrmThermal.switched_off = 0
        'End If

    End Sub

    Public Function SetControl(ByVal command_type As String, ByVal node_num As Integer, ByVal toggle As Boolean)

        'If command_type = "heater control" Then
        '    IDS.Devices.Thermal.FrmThermal.On_Off_thermal(node_num, toggle)
        'ElseIf command_type = "temperature control" Then
        '    If toggle = True Then ' true for operation, false for standby
        '        Select Case node_num
        '            Case 1
        '                IDS.Devices.Thermal.FrmThermal.Set_Temperature(node_num, IDS.Data.Hardware.Thermal.Needle.OperationTemp)
        '            Case 2
        '                IDS.Devices.Thermal.FrmThermal.Set_Temperature(node_num, IDS.Data.Hardware.Thermal.Syringe.OperationTemp)
        '            Case 3
        '                IDS.Devices.Thermal.FrmThermal.Set_Temperature(node_num, IDS.Data.Hardware.Thermal.Station1.OperationTemp)
        '            Case 4
        '                IDS.Devices.Thermal.FrmThermal.Set_Temperature(node_num, IDS.Data.Hardware.Thermal.Station2.OperationTemp)
        '            Case 5
        '                IDS.Devices.Thermal.FrmThermal.Set_Temperature(node_num, IDS.Data.Hardware.Thermal.Station3.OperationTemp)
        '        End Select
        '    Else
        '        Select Case node_num
        '            Case 1
        '                IDS.Devices.Thermal.FrmThermal.Set_Temperature(node_num, IDS.Data.Hardware.Thermal.Needle.StandbyTemp)
        '            Case 2
        '                IDS.Devices.Thermal.FrmThermal.Set_Temperature(node_num, IDS.Data.Hardware.Thermal.Syringe.StandbyTemp)
        '            Case 3
        '                IDS.Devices.Thermal.FrmThermal.Set_Temperature(node_num, IDS.Data.Hardware.Thermal.Station1.StandbyTemp)
        '            Case 4
        '                IDS.Devices.Thermal.FrmThermal.Set_Temperature(node_num, IDS.Data.Hardware.Thermal.Station2.StandbyTemp)
        '            Case 5
        '                IDS.Devices.Thermal.FrmThermal.Set_Temperature(node_num, IDS.Data.Hardware.Thermal.Station3.StandbyTemp)
        '        End Select
        '    End If
        'End If

    End Function

    '''''''''''''''
    '                                       kr                                       '
    ' set the heater_enabled
    '''''''''''''''

    Private Sub check_and_set_thermal()

        'If IDS.Data.Hardware.Thermal.Needle.OnOffControl = True Then
        '    LB_NeedleHeater.Visible = True
        '    HeaterControl0.Enabled = True
        '    HeaterTempControl0.Enabled = True
        '    MyHeaterSettings.Heater_Enabled(0) = True
        '    IDS.Devices.Thermal.FrmThermal.Set_Alarm(1, IDS.Data.Hardware.Thermal.Needle.AlarmThreshold)
        'Else
        '    LB_NeedleHeater.Visible = False
        '    HeaterControl0.Enabled = False
        '    HeaterTempControl0.Enabled = False
        'End If

        'If IDS.Data.Hardware.Thermal.Syringe.OnOffControl = True Then
        '    LB_SyrHeater.Visible = True
        '    HeaterControl1.Enabled = True
        '    HeaterTempControl1.Enabled = True
        '    MyHeaterSettings.Heater_Enabled(1) = True
        '    IDS.Devices.Thermal.FrmThermal.Set_Alarm(2, IDS.Data.Hardware.Thermal.Syringe.AlarmThreshold)
        'Else
        '    LB_SyrHeater.Visible = False
        '    HeaterControl1.Enabled = False
        '    HeaterTempControl1.Enabled = False
        'End If

        'If IDS.Data.Hardware.Thermal.Station1.OnOffControl = True Then
        '    LB_PreHeater.Visible = True
        '    HeaterControl2.Enabled = True
        '    HeaterTempControl2.Enabled = True
        '    MyHeaterSettings.Heater_Enabled(2) = True
        '    IDS.Devices.Thermal.FrmThermal.Set_Alarm(3, IDS.Data.Hardware.Thermal.Station1.AlarmThreshold)
        'Else
        '    LB_PreHeater.Visible = False
        '    HeaterControl2.Enabled = False
        '    HeaterTempControl2.Enabled = False
        'End If

        'If IDS.Data.Hardware.Thermal.Station2.OnOffControl = True Then
        '    LB_DispHeater.Visible = True
        '    HeaterControl3.Enabled = True
        '    HeaterTempControl3.Enabled = True
        '    MyHeaterSettings.Heater_Enabled(3) = True
        '    IDS.Devices.Thermal.FrmThermal.Set_Alarm(4, IDS.Data.Hardware.Thermal.Station3.AlarmThreshold)
        'Else
        '    LB_DispHeater.Visible = False
        '    HeaterControl3.Enabled = False
        '    HeaterTempControl3.Enabled = False
        'End If

        'If IDS.Data.Hardware.Thermal.Station3.OnOffControl = True Then
        '    LB_PostHeater.Visible = True
        '    HeaterControl4.Enabled = True
        '    HeaterTempControl4.Enabled = True
        '    MyHeaterSettings.Heater_Enabled(4) = True
        '    IDS.Devices.Thermal.FrmThermal.Set_Alarm(5, IDS.Data.Hardware.Thermal.Station3.AlarmThreshold)
        'Else
        '    LB_PostHeater.Visible = False
        '    HeaterControl4.Enabled = False
        '    HeaterTempControl4.Enabled = False
        'End If

    End Sub

#End Region


    Sub SelectRange(ByVal cell1 As Object, ByVal cell2 As Object)
        Try
            AxSpreadsheetProgramming.ActiveSheet.Range(cell1, cell2).Select()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Sub SelectCell(ByVal row As Integer, ByVal column As Integer)
        Try
            AxSpreadsheetProgramming.ActiveSheet.Cells(row, column).Select()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Sub SetChipEdgeSettings()
        Dim bb As Boolean
        Dim vpara As DLL_Export_Device_Vision.ChipEdgePoints.ChipEdgeParam
        Vision.GetChipEdgeParameters(vpara)
        Vision.SetCEReset()
        SetCellValue(m_Row, gEdgeClearColumn, vpara._EdgeClearance)
        SetCellValue(m_Row, gBrightnessColumn, vpara._Brightness)
        SetCellValue(m_Row, gCheckBoxColumn, vpara._CheckBox_ChipRec_Enable)
        SetCellValue(m_Row, gCwCCwColumn, vpara._Cw_CCw)
        SetCellValue(m_Row, gDispModelColumn, vpara._DispenseModel)
        SetCellValue(m_Row, gInOutColumn, vpara._Inside_out)
        SetCellValue(m_Row, gMainEdgeColumn, vpara._MainEdge)
        SetCellValue(m_Row, gPointX1Column, vpara._PointX1)
        SetCellValue(m_Row, gPointX2Column, vpara._PointX2)
        SetCellValue(m_Row, gPointX3Column, vpara._PointX3)
        SetCellValue(m_Row, gPointX4Column, vpara._PointX4)
        SetCellValue(m_Row, gPointX5Column, vpara._PointX5)
        SetCellValue(m_Row, gPointY1Column, vpara._PointY1)
        SetCellValue(m_Row, gPointY2Column, vpara._PointY2)
        SetCellValue(m_Row, gPointY3Column, vpara._PointY3)
        SetCellValue(m_Row, gPointY4Column, vpara._PointY4)
        SetCellValue(m_Row, gPointY5Column, vpara._PointY5)
        SetCellValue(m_Row, gPosColumn, vpara._Pos)
        SetCellValue(m_Row, gPosXColumn, vpara._PosX)
        SetCellValue(m_Row, gPosYColumn, vpara._PosY)
        SetCellValue(m_Row, gROIColumn, vpara._ROI)
        SetCellValue(m_Row, gRotColumn, vpara._Rot)
        SetCellValue(m_Row, gSizeColumn, vpara._Size)
        SetCellValue(m_Row, gSizeXColumn, vpara._SizeX)
        SetCellValue(m_Row, gSizeYColumn, vpara._SizeY)
        SetCellValue(m_Row, gThresholdColumn, vpara._Threshold)
        SetCellValue(m_Row, gVerticalColumn, vpara._Vertical)

        SetCellValue(m_Row, gDurationColumn, vpara._DotDispensingDuration)

        cell1 = GetCell(m_Row, gPos1XColumn)
        cell2 = GetCell(m_Row, gPos1ZColumn)
        m_Execution.m_Pattern.Spreadsheet_CellGrey(cell1, cell2, True, AxSpreadsheetProgramming)
    End Sub
    Sub SetGlobalQCSettings()
        Dim bb As Boolean
        Dim vpara As DLL_Export_Device_Vision.QC.QCParam
        Vision.GetQCParameters(vpara)
        Vision.SetQCReset()
        Dim m_Row As Integer = 1
        SetCellValue(m_Row, gBrightnessColumn, vpara._Brightness)
        SetCellValue(m_Row, gBinarizedColumn, vpara._Binarized)
        SetCellValue(m_Row, gBlackDotColumn, vpara._BlackDot)
        SetCellValue(m_Row, gOpenColumn, vpara._Open)
        SetCellValue(m_Row, gCloseColumn, vpara._Close)
        SetCellValue(m_Row, gCompactnessColumn, vpara._Compactness)
        SetCellValue(1, gMaxAreaColumn, vpara._MaxArea)
        SetCellValue(m_Row, gMinAreaColumn, vpara._MinArea)
        SetCellValue(m_Row, gMRegionXColumn, vpara._MRegionX)
        SetCellValue(m_Row, gMRegionYColumn, vpara._MRegionY)
        SetCellValue(m_Row, gMROIxColumn, vpara._MROIx)
        SetCellValue(m_Row, gMROIyColumn, vpara._MROIy)
        SetCellValue(m_Row, gRoughnessColumn, vpara._Roughness)
        SetCellValue(m_Row, gToleranceColumn, vpara._Tolerance)
        SetCellValue(m_Row, gDiameterColumn, vpara._Diameter)
        cell1 = GetCell(m_Row, gPos1XColumn)
        cell2 = GetCell(m_Row, gPos1ZColumn)
        Vision.SetQCReset()
        m_Execution.m_Pattern.Spreadsheet_CellGrey(cell1, cell2, True, AxSpreadsheetProgramming)
    End Sub
    Sub SetQCSettings()
        Dim bb As Boolean
        Dim vpara As DLL_Export_Device_Vision.QC.QCParam
        Vision.GetQCParameters(vpara)
        Vision.SetQCReset()
        SetCellValue(m_Row, gBrightnessColumn, vpara._Brightness)
        SetCellValue(m_Row, gBinarizedColumn, vpara._Binarized)
        SetCellValue(m_Row, gBlackDotColumn, vpara._BlackDot)
        SetCellValue(m_Row, gOpenColumn, vpara._Open)
        SetCellValue(m_Row, gCloseColumn, vpara._Close)
        SetCellValue(m_Row, gCompactnessColumn, vpara._Compactness)
        SetCellValue(m_Row, gMaxAreaColumn, vpara._MaxArea)
        SetCellValue(m_Row, gMinAreaColumn, vpara._MinArea)
        SetCellValue(m_Row, gMRegionXColumn, vpara._MRegionX)
        SetCellValue(m_Row, gMRegionYColumn, vpara._MRegionY)
        SetCellValue(m_Row, gMROIxColumn, vpara._MROIx)
        SetCellValue(m_Row, gMROIyColumn, vpara._MROIy)
        SetCellValue(m_Row, gRoughnessColumn, vpara._Roughness)
        SetCellValue(m_Row, gToleranceColumn, vpara._Tolerance)
        SetCellValue(m_Row, gDiameterColumn, vpara._Diameter)
        cell1 = GetCell(m_Row, gPos1XColumn)
        cell2 = GetCell(m_Row, gPos1ZColumn)
        Vision.SetQCReset()
        m_Execution.m_Pattern.Spreadsheet_CellGrey(cell1, cell2, True, AxSpreadsheetProgramming)
    End Sub
    Sub SetRejectMarkSettings()
        Dim vPara As DLL_Export_Device_Vision.RejectPoint.RMParam
        Dim bb As Boolean
        Vision.GetRMParameters(vPara)
        Vision.SetRMReset()
        SetCellValue(m_Row, gAcceptRatioCoulumn, vPara._AcceptRatio)
        SetCellValue(m_Row, gBinarizedColumn, vPara._Binarized)
        SetCellValue(m_Row, gBlackWithoutRMCoulumn, vPara._BlackWithoutRM)
        SetCellValue(m_Row, gBlackWithRMCoulumn, vPara._BlackWithRM)
        SetCellValue(m_Row, gBrightnessColumn, vPara._Brightness)
        SetCellValue(m_Row, gMRegionXColumn, vPara._MRegionX)
        SetCellValue(m_Row, gMRegionYColumn, vPara._MRegionY)
        SetCellValue(m_Row, gMROIxColumn, vPara._MROIx)
        SetCellValue(m_Row, gMROIyColumn, vPara._MROIy)
        SetCellValue(m_Row, gWhiteWithoutRMCoulumn, vPara._WhiteWithoutRM)
        SetCellValue(m_Row, gWhiteWithRMCoulumn, vPara._WhiteWithRM)
        SetCellValue(m_Row, gWoRMCoulumn, vPara._WoRM)
        SetCellValue(m_Row, gWRMCoulumn, vPara._WRM)
        cell1 = GetCell(m_Row, gPos1XColumn)
        cell2 = GetCell(m_Row, gPos1ZColumn)
        m_Execution.m_Pattern.Spreadsheet_CellGrey(cell1, cell2, True, AxSpreadsheetProgramming)
    End Sub

    Sub TeachCommand(ByVal type As String, ByVal cell1 As Object, ByVal cell2 As Object)
        m_Row = GetActiveCellRow()
        AddCommandToSpreadsheet(type)
        cell1 = GetCell(m_Row, gPos1XColumn)
        cell2 = GetCell(m_Row, gPos1ZColumn)
        AxSpreadsheetProgramming.ActiveSheet.Range(cell1, cell2).Select()
        ToolBarSwitch("YesNo")
    End Sub

    Dim debounce_counter As Integer = 5
    Public Sub IOCheck_tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles IOCheck.Tick
        Try

            'any consequence if we move this from statemonitor to a timer
            If m_Tri.EStopActivated And Not EStop Then
                EStopSequence()
                EStop = True
            ElseIf Form_Service.EStopPressedOk() And Not m_Tri.EStopActivated Then
                EStop = False
                LabelMessage("Restarting motion controller and carrying out homing.")
                DIO_Service.Off_Siren()
                m_Tri.Disconnect_Controller()
                m_Tri.Connect_Controller()
                SetState("Homing")
                Form_Service.ResetEventCode()
            End If

            If ProgrammingMode() And CurrentMode = "Basic Setup" Then Exit Sub
            If debounce_counter < 4 Then
                debounce_counter += 1
                Exit Sub
            End If
            Dim button_num_pressed As Integer = IDS.Devices.DIO.DIO.CheckOperationSwitch()
            If button_num_pressed <> 0 Then
                IOCheck.Stop()
                Select Case button_num_pressed
                    Case 1  'Park/ChangeSyr
                        SetState("Park")
                    Case 2  'Purge
                        If SetState("Purge") Then DoPurge()
                    Case 3  'Clean
                        If SetState("Clean") Then DoClean()
                    Case 4  'NeedleCal
                        SetState("Needle Calibration")
                    Case 5  'VolumeCal
                        SetState("Volume Calibration")
                    Case 6  'Run
                        SetState("Start")
                    Case 7  'Stop
                        StopDispensing()
                    Case 8  'Pause
                        PauseDispensing()
                    Case 9  'Reset
                        'SetState("Reset")
                    Case 10
                End Select
                IOCheck.Start()
                debounce_counter = 0
            End If
        Catch ex As Exception
            ExceptionDisplay(ex)
        End Try
    End Sub

    Public Function IsVisionTeachMode()
        Return VisionMode.Checked
    End Function

    Public Function IsNeedleTeachMode()
        Return NeedleMode.Checked
    End Function

    Public Sub DisableTeachModeSwitching()
        NeedleMode.Enabled = False
        VisionMode.Enabled = False
    End Sub

    Public Sub EnableTeachModeSwitching()
        NeedleMode.Enabled = True
        VisionMode.Enabled = True
    End Sub

    Public Sub SwitchToVisionTeachMode()
        NeedleMode.Checked = False
        VisionMode.Checked = True
    End Sub

    Public Sub SwitchToNeedleTeachMode()
        NeedleMode.Checked = True
        VisionMode.Checked = False
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Conveyor.MoveTo(IDS.Data.Hardware.Conveyor.Width)
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Conveyor.Command("Read Position")
    End Sub

    Private Sub ButtonStartFirstStage_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonStartFirstStage.Click
        DIO_Service.TriggerUpstream()
    End Sub

    'Private Sub FormProgramming_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
    '    If e.KeyCode = Keys.ControlKey Then
    '        isPress = True
    '    End If
    'End Sub

    'Private Sub FormProgramming_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyUp
    '    If e.KeyCode = Keys.ControlKey Then
    '        isPress = False
    '    End If
    'End Sub

    'Private Sub AxSpreadsheetProgramming_KeyUpEvent(ByVal sender As Object, ByVal e As AxOWC10.ISpreadsheetEventSink_KeyUpEvent) Handles AxSpreadsheetProgramming.KeyUpEvent
    '    If e.keyCode = Keys.ControlKey Then
    '        isPress = False
    '    End If
    'End Sub

    Private Sub AxSpreadsheetProgramming_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles AxSpreadsheetProgramming.LostFocus

    End Sub
    Private Function GetExistingLaserHeightDifferent(ByRef height As Double, ByVal selectedRow As Integer) As Boolean
        Dim Rows, Colums As Integer
        Dim sheetname As String = "Main"
        Dim sheet As OWC10.Worksheet = AxSpreadsheetProgramming.Workbooks.Item(1).Worksheets(sheetname)
        If sheet Is Nothing Then
            Return False
        End If
        Rows = sheet.UsedRange.Rows.Count
        Colums = sheet.UsedRange.Columns.Count
        If Rows < 1 Then
            Return False
        End If
        Dim array As Object(,) ' array start at (1,1)
        array = sheet.Range("A1", "AD" & Rows).Value
        Dim i As Integer = 1
        Dim foundHeight As Boolean = False
        For i = 1 To selectedRow
            Dim type As String = array((selectedRow + 1) - i, gCommandNameColumn)
            If type.ToUpper() = "HEIGHT" Then
                height = array((selectedRow + 1) - i, gPos1ZColumn)
                Console.WriteLine("Laser Height offset: " & height)
                Return True
            End If
        Next
    End Function
    Private Sub btPlay_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btPlay.Click
        If m_RunMode = 1 Or m_RunMode = 4 Then
            Dim Rows, Colums As Integer
            Dim sheetname As String = "Main"
            Dim sheet As OWC10.Worksheet = AxSpreadsheetProgramming.Workbooks.Item(1).Worksheets(sheetname)
            If sheet Is Nothing Then
                LogFile("Empy sheet when trying to run process")
                Return
            End If
            Rows = sheet.UsedRange.Rows.Count
            Colums = sheet.UsedRange.Columns.Count
            If Rows < 1 Then
                LogFile("Empy file when trying to run process")
                Return
            End If
            If teachingMode = "Vision" Then
                Dim array As Object(,) ' array start at (1,1)
                array = sheet.Range("A1", "AD" & Rows).Value
                Dim i As Integer = 1
                Dim foundHeight As Boolean = False
                For i = 1 To Rows
                    Dim type As String = array(i, gCommandNameColumn)
                    If type.ToUpper() = "HEIGHT" Then
                        foundHeight = True
                        Exit For
                    End If
                Next
                If Not (foundHeight) Then
                    Dim fm As InfoForm = New InfoForm
                    fm.SetTitle("Warning")
                    fm.OkOnly()
                    fm.AddNewLine("Please insert the laser height command before running Dry/Wet needle mode")
                    fm.SetOKButtonText("Ok")
                    LogFile("Try to run process without insert laser height command", 2)
                    fm.ShowDialog()
                    Return
                End If
            End If
        End If
        SetState("Start")
        btExit.Enabled = False
    End Sub

    Private Sub btPause_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btPause.Click
        PauseDispensing()
    End Sub

    Private Sub btStop_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btStop.Click
        StopDispensing()
        'btExit.Enabled = True
    End Sub

    Private Sub btRetrieve_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btRetrieve.Click
        Conveyor.Command("Retrieve")
    End Sub

    Private Sub btRelease_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btRelease.Click
        Conveyor.Command("Release")
    End Sub

    Private Sub AxSpreadsheetProgramming_DblClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles AxSpreadsheetProgramming.DblClick
        Dim pos(3) As Double
        Dim needleGap As Double
        Dim sel As Microsoft.Office.Interop.OWC.Range = AxSpreadsheetProgramming.Selection
        Dim m_rowCount As Integer = sel.Rows.Count()
        Dim m_columnCount As Integer = sel.Columns.Count()
        Dim m_StartRow As Integer = sel.Row
        Dim m_columnStart As Integer = sel.Column
        Dim m_EndRow As Integer = m_StartRow + m_rowCount - 1
        Dim RefPos() As Double = {0, 0, 0}
        Dim CmdStr As String = GetCellValue(m_StartRow, gCommandNameColumn)
        If CmdStr.ToUpper = "GLOBALQC" Then Return
        'If m_rowCount = 1 And CmdStr = Nothing Then
        '    LabelMessage("")
        'End If
        If Not (m_columnStart = 1) Then
            Return
        End If
        If 1 = m_rowCount Then
            SheetRowSelected = True
            DisplaySpreadsheetTabs()

            If m_Execution.m_Pattern.m_ErrorChk.CheckValidPtXY(CmdStr) Then
                m_TeachStepNumber = 1
                pos(0) = GetCellValue(m_StartRow, gPos1XColumn)
                pos(1) = GetCellValue(m_StartRow, gPos1YColumn)
                pos(2) = GetCellValue(m_StartRow, gPos1ZColumn)
                pos(0) = pos(0)
                pos(1) = pos(1)
                pos(2) = pos(2)
                needleGap = GetCellValue(m_StartRow, gNeedleGapColumn)

                AxSpreadsheetProgramming.Enabled = False

                If CmdStr = "ChipEdge" Or CmdStr = "QC" Or CmdStr = "Height" Or CmdStr = "Fiducial" Or CmdStr = "Reject" Then
                    MoveToSpreadsheetPoint(pos, "Vision")
                Else
                    MoveToSpreadsheetPoint(pos, "Needle", needleGap, m_StartRow)
                End If
                AxSpreadsheetProgramming.Enabled = True
            End If
        End If
    End Sub

    Private Sub btChangeSyringe_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btChangeSyringe.Click
        SetState("Change Syringe")
    End Sub
    Public Sub DisableCalibButtons()
        Dim eb As Boolean = False
        Me.ButtonNeedleCal.Enabled = eb
        Me.ButtonHome.Enabled = eb
        Me.ButtonVolCal.Enabled = eb
        Me.ButtonPurge.Enabled = eb
        Me.ButtonClean.Enabled = eb
        Me.btChangeSyringe.Enabled = eb
        Me.gbConveyor.Enabled = eb
        Me.VisionMode.Enabled = eb
        Me.NeedleMode.Enabled = eb
        ButtonToggleMode.Enabled = eb
        MenuItem1.Enabled = eb
    End Sub
    Public Sub EnableCalibButtons()
        Dim eb As Boolean = True
        Me.ButtonNeedleCal.Enabled = eb
        Me.ButtonHome.Enabled = eb
        Me.ButtonVolCal.Enabled = eb
        Me.ButtonPurge.Enabled = eb
        Me.ButtonClean.Enabled = eb
        Me.btChangeSyringe.Enabled = eb
        Me.gbConveyor.Enabled = eb
        Me.VisionMode.Enabled = eb
        MenuItem1.Enabled = eb
        Me.NeedleMode.Enabled = eb
        ButtonToggleMode.Enabled = eb
    End Sub

    Private Sub btExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btExit.Click
        Dim fm As InfoForm = New InfoForm
        fm.SetTitle("Warning")
        fm.AddNewLine("Are you sure you want to exit programming program?")
        fm.SetOKButtonText("Yes")
        fm.SetCancelButtonText("No")
        If fm.ShowDialog = DialogResult.Cancel Then
            Return
        End If
        m_Tri.AbortMotionDone()
        m_Tri.Move_Z(SafePosition)
        Close()
    End Sub

    Private Sub cbDisplayIndicator_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbDisplayIndicator.Click
        cbDisplayIndicator.Checked = True
        ClearAndDisplayIndicator()
    End Sub

    Private Sub btResetVolCalSettings_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btResetVolCalSettings.Click
        Laser.Instance.Show()
        'Production.ResetPressure()
    End Sub

    'Private Sub ToolBar1_ButtonClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ToolBarButtonClickEventArgs)
    '    If MachineRunning() Then
    '        Return
    '    End If
    '    offset(0) = gLeftNeedleOffs(0)
    '    offset(1) = gLeftNeedleOffs(1)
    '    offset(2) = gLeftNeedleOffs(2)

    '    If e.Button Is TeachingToolBar1.Buttons(0) Then
    '        If type <> "Move" Then
    '            If CheckSoftLimitXYZ(m_MachinePos, offset) Then Exit Sub
    '        End If
    '        Confirm()   'Add rows
    '    Else
    '        If m_EditStateFlag = False Then
    '            'DelayForRowDelete()
    '            If m_SteppingPostFlag Then
    '                DisableCoordinateUpdateInSpreadsheet()
    '                Spreadsheet_CheckForWithinLinkRange(True)
    '                DisableTeachingToolbarOKButton()
    '                EnableTeachingToolbarCancelButton()
    '                DisableEditingToolbar()
    '                EnableEditingToolbarSwitchButton()
    '                SelectCell(m_Row, gCommandNameColumn)
    '                m_SteppingPostFlag = False
    '                LabelMessage("")
    '            Else
    '                If (MessageBox.Show("Are you sure you want to delete the row/rows?", "Warning", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) = DialogResult.Yes) Then
    '                    DisableCoordinateUpdateInSpreadsheet()
    '                    'DelayForRowDelete()
    '                    type = GetCellValue(m_Row, gCommandNameColumn)
    '                    If type = "GlobalQC" Then
    '                        IDS.Data.Hardware.Camera.DotQCEnable = False
    '                    End If
    '                    Cancel()   'Cancel or delete rows
    '                    AxSpreadsheetProgramming.Enabled = True
    '                End If
    '            End If
    '            'DeletingRowFromExcel = False
    '            'DeletingRowFinished = False
    '        Else
    '            'm_EditStateFlag = False
    '            'm_SteppingPostFlag = False
    '            Console.WriteLine("Disable spredsheet update")
    '            DisableCoordinateUpdateInSpreadsheet()
    '            Spreadsheet_CheckForWithinLinkRange(True)
    '            DisableTeachingToolbarOKButton()
    '            EnableTeachingToolbarCancelButton()
    '            DisableEditingToolbar()

    '            SelectCell(m_Row, gCommandNameColumn)
    '            LabelMessage("")
    '            Cancel()
    '            'If (MessageBox.Show("Are you sure you want to delete the row/rows?", "Warning", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) = DialogResult.Yes) Then
    '            'Cancel()   'Cancel or delete rows
    '            'End If

    '        End If
    '    End If

    '    EnableTeachModeSwitching()
    'End Sub

    Private Sub ChipEdgeAdjustXY(ByVal xOffset As Double, ByVal yOffset As Double)
        Dim pos(2) As Double
        pos(0) = xOffset
        pos(1) = -yOffset
        pos(2) = 0
        m_Tri.Set_XY_Speed(IDS.Data.Hardware.Gantry.ElementXYSpeed)
        m_Tri.MoveRelative_XY(pos)
    End Sub
    Private Function ClickToMove(ByVal xOffset As Double, ByVal yOffset As Double) As Boolean
        If IsBusy() Or IsJogging() Then Exit Function
        If m_Tri.MachineHoming Or m_Tri.MachineRunning Or m_Tri.Calibrating Or m_Tri.Stepping Then Exit Function
        Dim pos(2) As Double
        pos(0) = xOffset
        pos(1) = -yOffset
        pos(2) = 0
        If KeyboardControl.ShiftKeyPressed Then
            m_Tri.Set_XY_Speed(IDS.Data.Hardware.Gantry.ElementXYSpeed)
            m_Tri.MoveRelative_XY(pos)
            Return True
        End If
        Return False
    End Function

    Public Function EnableClickToMove()
        Vision.FrmVision.EnableClickToMoveMode(True)
        Vision.FrmVision.ClickToMove = New DLL_Export_Device_Vision.FormVision.ClickToMoveDelegate(AddressOf Me.ClickToMove)
    End Function

    Private Function ReadDefaultOpenFile() As String
        If File.Exists("C:\Program Files\TIDS\f.dat") Then
            Dim reader As StreamReader = New StreamReader("C:\Program Files\TIDS\f.dat")
            Dim pth As String = reader.ReadLine()
            reader.Close()
            Return pth
        End If
        Return ""
    End Function
    Private Function WriteDefaultOpenFile(ByVal str As String) As Boolean
        Dim writer As StreamWriter = New StreamWriter("C:\Program Files\TIDS\f.dat")
        writer.WriteLine(str)
        writer.Close()
    End Function
    Private loadFirstTime As Boolean = True
    Public Function OpenDefaultFile()
        If Not loadFirstTime Then Exit Function
        Dim File As String = ReadDefaultOpenFile()
        If Not (System.IO.File.Exists(File)) Then Exit Function
        Dim Rtn As Boolean = False
        If Not (File = "") Then
            init_spreadsheet()
            FlushSpreadsheet()
            ErrorSubSheetStructIni(200, 500)
            LaserHeightOffsetZ = -1234
            Try
                Me.Cursor = Cursors.WaitCursor
                'IDS.StopErrorCheck() 'kr?
                AxSpreadsheetProgramming.Caption = File
                Me.GlobalQCEnabled = False
                m_Execution.m_Pattern.LoadTxtPatternPara(AxSpreadsheetProgramming, File, 0, 0, False)
                Me.Cursor = Cursors.WaitCursor
                gPatternFileName = m_Execution.m_File.FolderWithNameFromFileName(File)
                gFidFileName = gPatternFileName
                m_Row = 2
                Dim mode As String = ""
                If m_Execution.m_Pattern.GetTeachingMode(File, mode) Then
                    teachingMode = mode
                End If
                EnableTeachingButtons()
                DisableElementsCommandBlockButton(gOffsetCmdIndex)
                IDS.Data.OpenPathFileData(gPatternFileName + ".pat")

                MyConveyorSettings.InitializeConveyorSetup()
                SystemSetupDataRetrieve() 'SJ add 
                'IDS.StartErrorCheck() 'kr?
                IDS.newOpen = True
                LogFile("Opening file " & gPatternFileName + ".pat")
                'Acitvate the "Main" page
                AxSpreadsheetProgramming.Worksheets("Main").Activate()
                LabelMessage("Please wait, Checking content.....")
                'Error checking for all the Spreadsheet
                If 0 <> m_Execution.m_Pattern.m_ErrorChk.CheckAllError(AxSpreadsheetProgramming, ErrorSubSheet) Then
                    LabelMessage("Default file content invalid", True)

                    MenuFileExport.Enabled = True
                    MenuFileImport.Enabled = True
                    MenuFileSave.Enabled = True
                    MenuFileSaveAs.Enabled = True
                    UndoData_Logging(0)
                Else
                    LabelMessage("Default file opened")
                    tbOpenedFile.Text = File
                    UndoData_Logging(0)
                    MenuFileExport.Enabled = True
                    MenuFileImport.Enabled = True
                    MenuFileSave.Enabled = True
                    MenuFileSaveAs.Enabled = True
                    MenuEditUndo.Enabled = False
                    MenuEditRedo.Enabled = False
                    MenuEditCopy.Enabled = False
                    MenuEditPaste.Enabled = False
                    MenuEditCut.Enabled = False
                    MenuEditDelete.Enabled = False
                    MenuEditSelectAll.Enabled = True
                End If
                If gPatternFileName = Nothing Then
                Else
                    IDS.Data.OpenData()
                    Vision.SetFiducialFilename(gPatternFileName)
                    ValueBrightness.Value = IDS.Data.Hardware.Camera.Brightness
                    EnableControl()
                End If
                SaveProgram.UnSave = False
                Me.Cursor = Cursors.Default
            Catch ex As Exception
                ExceptionDisplay(ex)
                MessageBox.Show("Unknown error when opening default file", "Error information", MessageBoxButtons.OK)
                MenuFileExport.Enabled = True
                MenuFileImport.Enabled = True
                MenuFileSave.Enabled = True
                MenuFileSaveAs.Enabled = True
                Me.Cursor = Cursors.Default
            End Try
            LabelMessage("")
            loadFirstTime = False
        End If
    End Function

    Private Sub btEject_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btEject.Click
        DIO_Service.TriggerDownstream()
    End Sub

    Private Sub btReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btReset.Click
        Conveyor.Command("Reset Conveyor Status")
    End Sub

    Private Sub btConfirm_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btConfirm.Click
        If MachineRunning() Then
            Return
        End If
        offset(0) = gLeftNeedleOffs(0)
        offset(1) = gLeftNeedleOffs(1)
        offset(2) = gLeftNeedleOffs(2)
        If type <> "Move" Then
            If CheckSoftLimitXYZ(m_MachinePos, offset) Then Exit Sub
        End If
        Confirm()   'Add rows
        EnableTeachModeSwitching()
    End Sub

    Private Sub btCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btCancel.Click
        If MachineRunning() Then
            Return
        End If
        offset(0) = gLeftNeedleOffs(0)
        offset(1) = gLeftNeedleOffs(1)
        offset(2) = gLeftNeedleOffs(2)
        If m_EditStateFlag = False Then
            'DelayForRowDelete()
            If m_SteppingPostFlag Then
                DisableCoordinateUpdateInSpreadsheet()
                Spreadsheet_CheckForWithinLinkRange(True)
                DisableTeachingToolbarOKButton()
                EnableTeachingToolbarCancelButton()
                DisableEditingToolbar()
                EnableEditingToolbarSwitchButton()
                SelectCell(m_Row, gCommandNameColumn)
                m_SteppingPostFlag = False
                LabelMessage("")
            Else
                If (MessageBox.Show("Are you sure you want to delete the row/rows?", "Warning", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) = DialogResult.Yes) Then
                    DisableCoordinateUpdateInSpreadsheet()
                    'DelayForRowDelete()
                    type = GetCellValue(m_Row, gCommandNameColumn)
                    If type = "GlobalQC" Then
                        IDS.Data.Hardware.Camera.DotQCEnable = False
                    End If
                    Cancel()   'Cancel or delete rows
                    AxSpreadsheetProgramming.Enabled = True
                End If
            End If
            'DeletingRowFromExcel = False
            'DeletingRowFinished = False
        Else
            'm_EditStateFlag = False
            'm_SteppingPostFlag = False
            Console.WriteLine("Disable spredsheet update")
            DisableCoordinateUpdateInSpreadsheet()
            Spreadsheet_CheckForWithinLinkRange(True)
            DisableTeachingToolbarOKButton()
            EnableTeachingToolbarCancelButton()
            DisableEditingToolbar()

            SelectCell(m_Row, gCommandNameColumn)
            LabelMessage("")
            Cancel()
        End If
        EnableTeachModeSwitching()
    End Sub

    Private Sub btStep_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btStep.Click
        If Me.AxSpreadsheetProgramming.Enabled = False Then
            Return
        End If
        Dim idFlag As Integer = 0
            Dim sel As Microsoft.Office.Interop.OWC.Range = AxSpreadsheetProgramming.Selection
            Dim m_rowCount As Integer = sel.Rows.Count()
            Dim m_StartRow As Integer = sel.Row
            Dim commandName As String = GetCellValue(m_StartRow, gCommandNameColumn)
            If m_rowCount = 1 And commandName = Nothing Then
                If Not (m_EditStateFlag) And Not (m_ProgrammingStateFlag) Then
                    LabelMessage("")
                End If
                Return
            End If
            idFlag = 1
            m_EditStateFlag = False
            m_SteppingPostFlag = True
        Dim CmdName As String
        m_Row = GetActiveCellRow()
        EditingToolbar_Implementation(idFlag)
    End Sub

    Private Sub btEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btEdit.Click
        If Me.AxSpreadsheetProgramming.Enabled = False Then
            Return
        End If
        Dim idFlag As Integer = 0
        Dim CmdName As String
        m_Row = GetActiveCellRow()
        CmdName = GetCellValue(m_Row, gCommandNameColumn)
        If CmdName = "GlobalQC" Then
            m_EditStateFlag = True
            idFlag = 1 'edit
        End If
        EditingToolbar_Implementation(idFlag)
    End Sub
End Class
