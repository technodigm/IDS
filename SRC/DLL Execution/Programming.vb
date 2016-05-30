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
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents ImageListFiles As System.Windows.Forms.ImageList
    Friend WithEvents imageListElement As System.Windows.Forms.ImageList
    Friend WithEvents ImageListReference As System.Windows.Forms.ImageList
    Friend WithEvents TBFiducialPt As System.Windows.Forms.ToolBarButton
    Friend WithEvents TBHeightPt As System.Windows.Forms.ToolBarButton
    Friend WithEvents TBReferencePt As System.Windows.Forms.ToolBarButton
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents ImageListYesNo As System.Windows.Forms.ImageList
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Panel4 As System.Windows.Forms.Panel
    Friend WithEvents MenuItem76 As System.Windows.Forms.MenuItem
    Friend WithEvents TBRejectPt As System.Windows.Forms.ToolBarButton
    Friend WithEvents CBExpandSpreadsheet As System.Windows.Forms.CheckBox
    Friend WithEvents TBBOk As System.Windows.Forms.ToolBarButton
    Friend WithEvents TBBCancel As System.Windows.Forms.ToolBarButton
    Friend WithEvents TBBSwitch As System.Windows.Forms.ToolBarButton
    Friend WithEvents TBBEdit As System.Windows.Forms.ToolBarButton
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
    Friend WithEvents PBRed As System.Windows.Forms.PictureBox
    Friend WithEvents PBGreen As System.Windows.Forms.PictureBox
    Friend WithEvents TBOperation As System.Windows.Forms.ToolBar
    Friend WithEvents ToolBarButton1 As System.Windows.Forms.ToolBarButton
    Friend WithEvents ToolBarButton2 As System.Windows.Forms.ToolBarButton
    Friend WithEvents ToolBarButton5 As System.Windows.Forms.ToolBarButton
    Friend WithEvents PBYellow As System.Windows.Forms.PictureBox
    Public WithEvents PanelMotion As System.Windows.Forms.Panel
    Friend WithEvents ButtonToggleMode As System.Windows.Forms.Button
    Public WithEvents PanelToBeAdded As System.Windows.Forms.Panel
    Friend WithEvents ComboBoxFineStep As System.Windows.Forms.NumericUpDown
    Friend WithEvents ButtonStepZdown As System.Windows.Forms.Button
    Friend WithEvents ButtonStepZup As System.Windows.Forms.Button
    Friend WithEvents ButtonStepXminus As System.Windows.Forms.Button
    Friend WithEvents ButtonStepXplus As System.Windows.Forms.Button
    Friend WithEvents ButtonStepYminus As System.Windows.Forms.Button
    Friend WithEvents ButtonStepYplus As System.Windows.Forms.Button
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents LabelStep As System.Windows.Forms.Label
    Friend WithEvents TeachingToolbar As System.Windows.Forms.ToolBar
    Friend WithEvents EditingToolbar As System.Windows.Forms.ToolBar
    Friend WithEvents VisionMode As System.Windows.Forms.RadioButton
    Friend WithEvents ReferenceCommandBlock As System.Windows.Forms.ToolBar
    Friend WithEvents ElementsCommandBlock As System.Windows.Forms.ToolBar
    Friend WithEvents DispensingMode As System.Windows.Forms.ComboBox
    Friend WithEvents TBBVolumeCal As System.Windows.Forms.ToolBarButton
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox2 As System.Windows.Forms.TextBox
    Friend WithEvents ValueBrightness As System.Windows.Forms.NumericUpDown
    Friend WithEvents cross_num As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents ButtonStartFirstStage As System.Windows.Forms.Button
    Friend WithEvents ButtonCV_Prod_Release As System.Windows.Forms.Button
    Friend WithEvents ButtonCV_Prod_Retrieve As System.Windows.Forms.Button
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
        Me.ValueBrightness = New System.Windows.Forms.NumericUpDown
        Me.CheckBoxLockZ = New System.Windows.Forms.CheckBox
        Me.TextBoxRobotZ = New System.Windows.Forms.TextBox
        Me.CheckBoxLockY = New System.Windows.Forms.CheckBox
        Me.TextBoxRobotY = New System.Windows.Forms.TextBox
        Me.CheckBoxLockX = New System.Windows.Forms.CheckBox
        Me.TextBoxRobotX = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.DomainUpDown1 = New System.Windows.Forms.DomainUpDown
        Me.Label1 = New System.Windows.Forms.Label
        Me.ImageListFiles = New System.Windows.Forms.ImageList(Me.components)
        Me.imageListElement = New System.Windows.Forms.ImageList(Me.components)
        Me.ImageListReference = New System.Windows.Forms.ImageList(Me.components)
        Me.ReferenceCommandBlock = New System.Windows.Forms.ToolBar
        Me.TBFiducialPt = New System.Windows.Forms.ToolBarButton
        Me.TBHeightPt = New System.Windows.Forms.ToolBarButton
        Me.TBReferencePt = New System.Windows.Forms.ToolBarButton
        Me.TBRejectPt = New System.Windows.Forms.ToolBarButton
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
        Me.ImageListYesNo = New System.Windows.Forms.ImageList(Me.components)
        Me.Label5 = New System.Windows.Forms.Label
        Me.VisionMode = New System.Windows.Forms.RadioButton
        Me.CBExpandSpreadsheet = New System.Windows.Forms.CheckBox
        Me.Panel4 = New System.Windows.Forms.Panel
        Me.NeedleMode = New System.Windows.Forms.RadioButton
        Me.TeachingToolbar = New System.Windows.Forms.ToolBar
        Me.TBBOk = New System.Windows.Forms.ToolBarButton
        Me.TBBCancel = New System.Windows.Forms.ToolBarButton
        Me.EditingToolbar = New System.Windows.Forms.ToolBar
        Me.TBBSwitch = New System.Windows.Forms.ToolBarButton
        Me.TBBEdit = New System.Windows.Forms.ToolBarButton
        Me.ImageListOper = New System.Windows.Forms.ImageList(Me.components)
        Me.ButtonClean = New System.Windows.Forms.Button
        Me.ButtonNeedleCal = New System.Windows.Forms.Button
        Me.ButtonVolCal = New System.Windows.Forms.Button
        Me.ButtonHome = New System.Windows.Forms.Button
        Me.ImageListMultiField = New System.Windows.Forms.ImageList(Me.components)
        Me.CBDoorLock = New System.Windows.Forms.CheckBox
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
        Me.AxSpreadsheetProgramming = New AxOWC10.AxSpreadsheet
        Me.IOCheck = New System.Windows.Forms.Timer(Me.components)
        Me.LabelMessege = New System.Windows.Forms.Label
        Me.DispensingMode = New System.Windows.Forms.ComboBox
        Me.PBRed = New System.Windows.Forms.PictureBox
        Me.PBGreen = New System.Windows.Forms.PictureBox
        Me.TBOperation = New System.Windows.Forms.ToolBar
        Me.ToolBarButton1 = New System.Windows.Forms.ToolBarButton
        Me.ToolBarButton2 = New System.Windows.Forms.ToolBarButton
        Me.ToolBarButton5 = New System.Windows.Forms.ToolBarButton
        Me.PBYellow = New System.Windows.Forms.PictureBox
        Me.PanelMotion = New System.Windows.Forms.Panel
        Me.ButtonToggleMode = New System.Windows.Forms.Button
        Me.PanelToBeAdded = New System.Windows.Forms.Panel
        Me.ComboBoxFineStep = New System.Windows.Forms.NumericUpDown
        Me.ButtonStepZdown = New System.Windows.Forms.Button
        Me.ButtonStepZup = New System.Windows.Forms.Button
        Me.ButtonStepXminus = New System.Windows.Forms.Button
        Me.ButtonStepXplus = New System.Windows.Forms.Button
        Me.ButtonStepYminus = New System.Windows.Forms.Button
        Me.ButtonStepYplus = New System.Windows.Forms.Button
        Me.Label6 = New System.Windows.Forms.Label
        Me.LabelStep = New System.Windows.Forms.Label
        Me.TextBox1 = New System.Windows.Forms.TextBox
        Me.TextBox2 = New System.Windows.Forms.TextBox
        Me.cross_num = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.ButtonStartFirstStage = New System.Windows.Forms.Button
        Me.ButtonCV_Prod_Release = New System.Windows.Forms.Button
        Me.ButtonCV_Prod_Retrieve = New System.Windows.Forms.Button
        Me.PanelVisionCtrl.SuspendLayout()
        CType(Me.ValueBrightness, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel4.SuspendLayout()
        CType(Me.AxSpreadsheetProgramming, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.PanelMotion.SuspendLayout()
        Me.PanelToBeAdded.SuspendLayout()
        CType(Me.ComboBoxFineStep, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'MainMenuProgramming
        '
        Me.MainMenuProgramming.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem1, Me.MenuItem8, Me.MenuItem11, Me.MenuItem25, Me.MenuItem64})
        '
        'MenuItem1
        '
        Me.MenuItem1.Index = 0
        Me.MenuItem1.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuFileNew, Me.MenuFileOpen, Me.MenuItem19, Me.MenuFileImport, Me.MenuFileExport, Me.MenuItem20, Me.MenuFileSave, Me.MenuFileSaveAs, Me.MenuItem6, Me.MenuFileExit})
        Me.MenuItem1.Text = "File"
        '
        'MenuFileNew
        '
        Me.MenuFileNew.Index = 0
        Me.MenuFileNew.Text = "New"
        '
        'MenuFileOpen
        '
        Me.MenuFileOpen.Index = 1
        Me.MenuFileOpen.Text = "Open"
        '
        'MenuItem19
        '
        Me.MenuItem19.Index = 2
        Me.MenuItem19.Text = "-"
        '
        'MenuFileImport
        '
        Me.MenuFileImport.Index = 3
        Me.MenuFileImport.Text = "Import"
        '
        'MenuFileExport
        '
        Me.MenuFileExport.Index = 4
        Me.MenuFileExport.Text = "Export"
        '
        'MenuItem20
        '
        Me.MenuItem20.Index = 5
        Me.MenuItem20.Text = "-"
        '
        'MenuFileSave
        '
        Me.MenuFileSave.Index = 6
        Me.MenuFileSave.Text = "Save"
        '
        'MenuFileSaveAs
        '
        Me.MenuFileSaveAs.Index = 7
        Me.MenuFileSaveAs.Text = "Save As..."
        '
        'MenuItem6
        '
        Me.MenuItem6.Index = 8
        Me.MenuItem6.Text = "-"
        '
        'MenuFileExit
        '
        Me.MenuFileExit.Index = 9
        Me.MenuFileExit.Text = "Exit"
        '
        'MenuItem8
        '
        Me.MenuItem8.Index = 1
        Me.MenuItem8.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuEditUndo, Me.MenuEditRedo, Me.MenuItem15, Me.MenuEditCut, Me.MenuEditCopy, Me.MenuEditPaste, Me.MenuEditDelete, Me.MenuItem17, Me.MenuEditSelectAll})
        Me.MenuItem8.Text = "Edit"
        '
        'MenuEditUndo
        '
        Me.MenuEditUndo.Index = 0
        Me.MenuEditUndo.Text = "Undo"
        '
        'MenuEditRedo
        '
        Me.MenuEditRedo.Index = 1
        Me.MenuEditRedo.Text = "Redo"
        '
        'MenuItem15
        '
        Me.MenuItem15.Index = 2
        Me.MenuItem15.Text = "-"
        '
        'MenuEditCut
        '
        Me.MenuEditCut.Index = 3
        Me.MenuEditCut.Text = "Cut"
        '
        'MenuEditCopy
        '
        Me.MenuEditCopy.Index = 4
        Me.MenuEditCopy.Text = "Copy"
        '
        'MenuEditPaste
        '
        Me.MenuEditPaste.Index = 5
        Me.MenuEditPaste.Text = "Paste"
        '
        'MenuEditDelete
        '
        Me.MenuEditDelete.Index = 6
        Me.MenuEditDelete.Text = "Delete"
        '
        'MenuItem17
        '
        Me.MenuItem17.Index = 7
        Me.MenuItem17.Text = "-"
        '
        'MenuEditSelectAll
        '
        Me.MenuEditSelectAll.Index = 8
        Me.MenuEditSelectAll.Text = "Select All"
        '
        'MenuItem11
        '
        Me.MenuItem11.Index = 2
        Me.MenuItem11.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem67, Me.MenuItem87, Me.MenuItem88})
        Me.MenuItem11.Text = "View"
        '
        'MenuItem67
        '
        Me.MenuItem67.Index = 0
        Me.MenuItem67.Text = "Information"
        '
        'MenuItem87
        '
        Me.MenuItem87.Index = 1
        Me.MenuItem87.Text = "-"
        '
        'MenuItem88
        '
        Me.MenuItem88.Index = 2
        Me.MenuItem88.Text = "Open Event Viewer"
        '
        'MenuItem25
        '
        Me.MenuItem25.Index = 3
        Me.MenuItem25.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem33, Me.MenuItem32, Me.MenuItem35, Me.MenuItem29, Me.MenuItem30, Me.MenuItem36, Me.MenuItem62, Me.MenuItem76, Me.OptimizePath, Me.MenuItem3})
        Me.MenuItem25.Text = "Tools"
        '
        'MenuItem33
        '
        Me.MenuItem33.Index = 0
        Me.MenuItem33.Text = "Offset"
        '
        'MenuItem32
        '
        Me.MenuItem32.Index = 1
        Me.MenuItem32.Text = "Sub Pattern"
        '
        'MenuItem35
        '
        Me.MenuItem35.Index = 2
        Me.MenuItem35.Text = "-"
        '
        'MenuItem29
        '
        Me.MenuItem29.Index = 3
        Me.MenuItem29.Text = "Set I/O"
        '
        'MenuItem30
        '
        Me.MenuItem30.Index = 4
        Me.MenuItem30.Text = "Reset I/O"
        '
        'MenuItem36
        '
        Me.MenuItem36.Index = 5
        Me.MenuItem36.Text = "-"
        '
        'MenuItem62
        '
        Me.MenuItem62.Index = 6
        Me.MenuItem62.Text = "Measurement"
        '
        'MenuItem76
        '
        Me.MenuItem76.Index = 7
        Me.MenuItem76.Text = "-"
        '
        'OptimizePath
        '
        Me.OptimizePath.Index = 8
        Me.OptimizePath.Text = "Optimize Dispensing Path"
        '
        'MenuItem3
        '
        Me.MenuItem3.Index = 9
        Me.MenuItem3.Text = "-"
        '
        'MenuItem64
        '
        Me.MenuItem64.Index = 4
        Me.MenuItem64.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem65})
        Me.MenuItem64.Text = "Help"
        '
        'MenuItem65
        '
        Me.MenuItem65.Index = 0
        Me.MenuItem65.Text = "About IDS"
        '
        'PanelVision
        '
        Me.PanelVision.BackColor = System.Drawing.Color.SlateGray
        Me.PanelVision.Location = New System.Drawing.Point(84, 360)
        Me.PanelVision.Name = "PanelVision"
        Me.PanelVision.Size = New System.Drawing.Size(768, 576)
        Me.PanelVision.TabIndex = 0
        '
        'ImageListGeneralTools
        '
        Me.ImageListGeneralTools.ImageSize = New System.Drawing.Size(36, 28)
        Me.ImageListGeneralTools.ImageStream = CType(resources.GetObject("ImageListGeneralTools.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.ImageListGeneralTools.TransparentColor = System.Drawing.Color.White
        '
        'ButtonPurge
        '
        Me.ButtonPurge.BackColor = System.Drawing.SystemColors.InactiveCaptionText
        Me.ButtonPurge.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonPurge.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.ButtonPurge.ImageIndex = 4
        Me.ButtonPurge.ImageList = Me.ImageListGeneralTools
        Me.ButtonPurge.Location = New System.Drawing.Point(856, 530)
        Me.ButtonPurge.Name = "ButtonPurge"
        Me.ButtonPurge.Size = New System.Drawing.Size(88, 80)
        Me.ButtonPurge.TabIndex = 57
        Me.ButtonPurge.Text = "Purge On"
        Me.ButtonPurge.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'PanelVisionCtrl
        '
        Me.PanelVisionCtrl.Controls.Add(Me.ValueBrightness)
        Me.PanelVisionCtrl.Controls.Add(Me.CheckBoxLockZ)
        Me.PanelVisionCtrl.Controls.Add(Me.TextBoxRobotZ)
        Me.PanelVisionCtrl.Controls.Add(Me.CheckBoxLockY)
        Me.PanelVisionCtrl.Controls.Add(Me.TextBoxRobotY)
        Me.PanelVisionCtrl.Controls.Add(Me.CheckBoxLockX)
        Me.PanelVisionCtrl.Controls.Add(Me.TextBoxRobotX)
        Me.PanelVisionCtrl.Controls.Add(Me.Label4)
        Me.PanelVisionCtrl.Controls.Add(Me.DomainUpDown1)
        Me.PanelVisionCtrl.Controls.Add(Me.Label1)
        Me.PanelVisionCtrl.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.PanelVisionCtrl.Location = New System.Drawing.Point(84, 940)
        Me.PanelVisionCtrl.Name = "PanelVisionCtrl"
        Me.PanelVisionCtrl.Size = New System.Drawing.Size(768, 32)
        Me.PanelVisionCtrl.TabIndex = 2
        '
        'ValueBrightness
        '
        Me.ValueBrightness.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ValueBrightness.Location = New System.Drawing.Point(72, 3)
        Me.ValueBrightness.Maximum = New Decimal(New Integer() {255, 0, 0, 0})
        Me.ValueBrightness.Name = "ValueBrightness"
        Me.ValueBrightness.Size = New System.Drawing.Size(45, 21)
        Me.ValueBrightness.TabIndex = 76
        Me.ValueBrightness.Value = New Decimal(New Integer() {10, 0, 0, 0})
        '
        'CheckBoxLockZ
        '
        Me.CheckBoxLockZ.BackColor = System.Drawing.SystemColors.Control
        Me.CheckBoxLockZ.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CheckBoxLockZ.Image = CType(resources.GetObject("CheckBoxLockZ.Image"), System.Drawing.Image)
        Me.CheckBoxLockZ.Location = New System.Drawing.Point(724, 5)
        Me.CheckBoxLockZ.Name = "CheckBoxLockZ"
        Me.CheckBoxLockZ.Size = New System.Drawing.Size(40, 16)
        Me.CheckBoxLockZ.TabIndex = 75
        '
        'TextBoxRobotZ
        '
        Me.TextBoxRobotZ.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBoxRobotZ.Location = New System.Drawing.Point(651, 4)
        Me.TextBoxRobotZ.Name = "TextBoxRobotZ"
        Me.TextBoxRobotZ.ReadOnly = True
        Me.TextBoxRobotZ.Size = New System.Drawing.Size(74, 21)
        Me.TextBoxRobotZ.TabIndex = 74
        Me.TextBoxRobotZ.Text = "Z: 100.000"
        '
        'CheckBoxLockY
        '
        Me.CheckBoxLockY.BackColor = System.Drawing.SystemColors.Control
        Me.CheckBoxLockY.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CheckBoxLockY.Image = CType(resources.GetObject("CheckBoxLockY.Image"), System.Drawing.Image)
        Me.CheckBoxLockY.Location = New System.Drawing.Point(611, 5)
        Me.CheckBoxLockY.Name = "CheckBoxLockY"
        Me.CheckBoxLockY.Size = New System.Drawing.Size(40, 16)
        Me.CheckBoxLockY.TabIndex = 73
        '
        'TextBoxRobotY
        '
        Me.TextBoxRobotY.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBoxRobotY.Location = New System.Drawing.Point(535, 4)
        Me.TextBoxRobotY.Name = "TextBoxRobotY"
        Me.TextBoxRobotY.ReadOnly = True
        Me.TextBoxRobotY.Size = New System.Drawing.Size(74, 21)
        Me.TextBoxRobotY.TabIndex = 72
        Me.TextBoxRobotY.Text = "Y: 100.000"
        '
        'CheckBoxLockX
        '
        Me.CheckBoxLockX.BackColor = System.Drawing.SystemColors.Control
        Me.CheckBoxLockX.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CheckBoxLockX.Image = CType(resources.GetObject("CheckBoxLockX.Image"), System.Drawing.Image)
        Me.CheckBoxLockX.Location = New System.Drawing.Point(495, 5)
        Me.CheckBoxLockX.Name = "CheckBoxLockX"
        Me.CheckBoxLockX.Size = New System.Drawing.Size(40, 16)
        Me.CheckBoxLockX.TabIndex = 71
        '
        'TextBoxRobotX
        '
        Me.TextBoxRobotX.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBoxRobotX.Location = New System.Drawing.Point(420, 4)
        Me.TextBoxRobotX.Name = "TextBoxRobotX"
        Me.TextBoxRobotX.ReadOnly = True
        Me.TextBoxRobotX.Size = New System.Drawing.Size(74, 21)
        Me.TextBoxRobotX.TabIndex = 6
        Me.TextBoxRobotX.Text = "X: 100.000"
        '
        'Label4
        '
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(374, 8)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(42, 16)
        Me.Label4.TabIndex = 7
        Me.Label4.Text = "Robot"
        '
        'DomainUpDown1
        '
        Me.DomainUpDown1.Items.Add("Aa")
        Me.DomainUpDown1.Items.Add("b")
        Me.DomainUpDown1.Items.Add("c")
        Me.DomainUpDown1.Location = New System.Drawing.Point(72, 3)
        Me.DomainUpDown1.Name = "DomainUpDown1"
        Me.DomainUpDown1.Size = New System.Drawing.Size(40, 20)
        Me.DomainUpDown1.TabIndex = 1
        Me.DomainUpDown1.Text = "50"
        Me.DomainUpDown1.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(2, 5)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(68, 16)
        Me.Label1.TabIndex = 3
        Me.Label1.Text = "Brightness"
        '
        'ImageListFiles
        '
        Me.ImageListFiles.ImageSize = New System.Drawing.Size(16, 16)
        Me.ImageListFiles.ImageStream = CType(resources.GetObject("ImageListFiles.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.ImageListFiles.TransparentColor = System.Drawing.Color.Transparent
        '
        'imageListElement
        '
        Me.imageListElement.ImageSize = New System.Drawing.Size(30, 30)
        Me.imageListElement.ImageStream = CType(resources.GetObject("imageListElement.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.imageListElement.TransparentColor = System.Drawing.Color.DarkGray
        '
        'ImageListReference
        '
        Me.ImageListReference.ImageSize = New System.Drawing.Size(30, 30)
        Me.ImageListReference.ImageStream = CType(resources.GetObject("ImageListReference.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.ImageListReference.TransparentColor = System.Drawing.Color.White
        '
        'ReferenceCommandBlock
        '
        Me.ReferenceCommandBlock.AllowDrop = True
        Me.ReferenceCommandBlock.Anchor = System.Windows.Forms.AnchorStyles.Left
        Me.ReferenceCommandBlock.Buttons.AddRange(New System.Windows.Forms.ToolBarButton() {Me.TBFiducialPt, Me.TBHeightPt, Me.TBReferencePt, Me.TBRejectPt})
        Me.ReferenceCommandBlock.ButtonSize = New System.Drawing.Size(42, 42)
        Me.ReferenceCommandBlock.Dock = System.Windows.Forms.DockStyle.None
        Me.ReferenceCommandBlock.DropDownArrows = True
        Me.ReferenceCommandBlock.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ReferenceCommandBlock.ImageList = Me.ImageListReference
        Me.ReferenceCommandBlock.Location = New System.Drawing.Point(0, 339)
        Me.ReferenceCommandBlock.Name = "ReferenceCommandBlock"
        Me.ReferenceCommandBlock.ShowToolTips = True
        Me.ReferenceCommandBlock.Size = New System.Drawing.Size(86, 90)
        Me.ReferenceCommandBlock.TabIndex = 82
        '
        'TBFiducialPt
        '
        Me.TBFiducialPt.ImageIndex = 1
        Me.TBFiducialPt.Text = "      Fiducial"
        Me.TBFiducialPt.ToolTipText = "Fiducial Point"
        '
        'TBHeightPt
        '
        Me.TBHeightPt.ImageIndex = 2
        Me.TBHeightPt.Text = "      Height"
        Me.TBHeightPt.ToolTipText = "Height Point"
        '
        'TBReferencePt
        '
        Me.TBReferencePt.ImageIndex = 0
        Me.TBReferencePt.Text = "     Reference"
        Me.TBReferencePt.ToolTipText = "Reference Point"
        '
        'TBRejectPt
        '
        Me.TBRejectPt.ImageIndex = 3
        Me.TBRejectPt.Text = "     Reject"
        Me.TBRejectPt.ToolTipText = "Reject Point"
        '
        'ElementsCommandBlock
        '
        Me.ElementsCommandBlock.Anchor = System.Windows.Forms.AnchorStyles.Left
        Me.ElementsCommandBlock.Buttons.AddRange(New System.Windows.Forms.ToolBarButton() {Me.TBBDot, Me.TBBLine, Me.TBBArc, Me.TBBRectangle, Me.TBBCircle, Me.TBBFilledRectangle, Me.TBBFilledCircle, Me.TBBLink, Me.TBBChipEdge, Me.TBBMove, Me.TBBWait, Me.TBBPurge, Me.TBBClean, Me.TBBQC, Me.TBBSubPattern, Me.TBBArray, Me.TBBGetIO, Me.TBBSetIO, Me.TBBOffset, Me.TBBMeasure, Me.TBBDotArray, Me.TBBVolumeCal})
        Me.ElementsCommandBlock.ButtonSize = New System.Drawing.Size(42, 42)
        Me.ElementsCommandBlock.Dock = System.Windows.Forms.DockStyle.None
        Me.ElementsCommandBlock.DropDownArrows = True
        Me.ElementsCommandBlock.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ElementsCommandBlock.ImageList = Me.imageListElement
        Me.ElementsCommandBlock.Location = New System.Drawing.Point(0, 443)
        Me.ElementsCommandBlock.Name = "ElementsCommandBlock"
        Me.ElementsCommandBlock.ShowToolTips = True
        Me.ElementsCommandBlock.Size = New System.Drawing.Size(88, 468)
        Me.ElementsCommandBlock.TabIndex = 83
        '
        'TBBDot
        '
        Me.TBBDot.ImageIndex = 0
        Me.TBBDot.Text = "       Dot"
        Me.TBBDot.ToolTipText = "Dot"
        '
        'TBBLine
        '
        Me.TBBLine.ImageIndex = 1
        Me.TBBLine.Text = "      Line"
        Me.TBBLine.ToolTipText = "Line"
        '
        'TBBArc
        '
        Me.TBBArc.ImageIndex = 2
        Me.TBBArc.Text = "       Arc"
        Me.TBBArc.ToolTipText = "Arc"
        '
        'TBBRectangle
        '
        Me.TBBRectangle.ImageIndex = 3
        Me.TBBRectangle.Text = "     Rectangle"
        Me.TBBRectangle.ToolTipText = "Rectangle"
        '
        'TBBCircle
        '
        Me.TBBCircle.ImageIndex = 4
        Me.TBBCircle.Text = "     Circle"
        Me.TBBCircle.ToolTipText = "Circle"
        '
        'TBBFilledRectangle
        '
        Me.TBBFilledRectangle.ImageIndex = 5
        Me.TBBFilledRectangle.Text = "     FillRectangle"
        Me.TBBFilledRectangle.ToolTipText = "FillRectangle"
        '
        'TBBFilledCircle
        '
        Me.TBBFilledCircle.ImageIndex = 6
        Me.TBBFilledCircle.Text = "      FillCircle"
        Me.TBBFilledCircle.ToolTipText = "FillCircle"
        '
        'TBBLink
        '
        Me.TBBLink.ImageIndex = 7
        Me.TBBLink.Text = "     Link"
        Me.TBBLink.ToolTipText = "Link"
        '
        'TBBChipEdge
        '
        Me.TBBChipEdge.ImageIndex = 8
        Me.TBBChipEdge.Text = "     ChipEdge"
        Me.TBBChipEdge.ToolTipText = "ChipEdge"
        '
        'TBBMove
        '
        Me.TBBMove.ImageIndex = 9
        Me.TBBMove.Text = "     Move"
        Me.TBBMove.ToolTipText = "Move"
        '
        'TBBWait
        '
        Me.TBBWait.ImageIndex = 10
        Me.TBBWait.Text = "    Wait"
        Me.TBBWait.ToolTipText = "Wait"
        '
        'TBBPurge
        '
        Me.TBBPurge.ImageIndex = 11
        Me.TBBPurge.Text = "     Purge"
        Me.TBBPurge.ToolTipText = "Purge"
        '
        'TBBClean
        '
        Me.TBBClean.ImageIndex = 12
        Me.TBBClean.Text = "     Clean"
        Me.TBBClean.ToolTipText = "Clean"
        '
        'TBBQC
        '
        Me.TBBQC.ImageIndex = 13
        Me.TBBQC.Text = "     QC"
        Me.TBBQC.ToolTipText = "QCCheck"
        '
        'TBBSubPattern
        '
        Me.TBBSubPattern.ImageIndex = 14
        Me.TBBSubPattern.Text = "      SubPattern"
        Me.TBBSubPattern.ToolTipText = "Call Sub Pattern"
        '
        'TBBArray
        '
        Me.TBBArray.ImageIndex = 15
        Me.TBBArray.Text = "     Array"
        Me.TBBArray.ToolTipText = "Array"
        '
        'TBBGetIO
        '
        Me.TBBGetIO.ImageIndex = 16
        Me.TBBGetIO.Text = "    GetIO"
        Me.TBBGetIO.ToolTipText = "Get Digital Input"
        '
        'TBBSetIO
        '
        Me.TBBSetIO.ImageIndex = 17
        Me.TBBSetIO.Text = "     SetIO"
        Me.TBBSetIO.ToolTipText = "Set Digital Output"
        '
        'TBBOffset
        '
        Me.TBBOffset.ImageIndex = 18
        Me.TBBOffset.Text = "      Offset"
        Me.TBBOffset.ToolTipText = "Offset"
        '
        'TBBMeasure
        '
        Me.TBBMeasure.ImageIndex = 19
        Me.TBBMeasure.Text = "    Measure"
        Me.TBBMeasure.ToolTipText = "Measure"
        '
        'TBBDotArray
        '
        Me.TBBDotArray.ImageIndex = 20
        Me.TBBDotArray.Text = "    DotArray"
        Me.TBBDotArray.ToolTipText = "Dot Array"
        '
        'TBBVolumeCal
        '
        Me.TBBVolumeCal.Text = "     VolumeCalibration"
        '
        'ImageListYesNo
        '
        Me.ImageListYesNo.ImageSize = New System.Drawing.Size(30, 30)
        Me.ImageListYesNo.ImageStream = CType(resources.GetObject("ImageListYesNo.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.ImageListYesNo.TransparentColor = System.Drawing.Color.White
        '
        'Label5
        '
        Me.Label5.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label5.Location = New System.Drawing.Point(46, 16)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(112, 23)
        Me.Label5.TabIndex = 87
        Me.Label5.Text = "Teach Mode:"
        '
        'VisionMode
        '
        Me.VisionMode.AutoCheck = False
        Me.VisionMode.Checked = True
        Me.VisionMode.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.VisionMode.Location = New System.Drawing.Point(182, 16)
        Me.VisionMode.Name = "VisionMode"
        Me.VisionMode.Size = New System.Drawing.Size(80, 24)
        Me.VisionMode.TabIndex = 86
        Me.VisionMode.TabStop = True
        Me.VisionMode.Text = "Vision"
        '
        'CBExpandSpreadsheet
        '
        Me.CBExpandSpreadsheet.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.CBExpandSpreadsheet.Location = New System.Drawing.Point(1084, 7)
        Me.CBExpandSpreadsheet.Name = "CBExpandSpreadsheet"
        Me.CBExpandSpreadsheet.Size = New System.Drawing.Size(192, 24)
        Me.CBExpandSpreadsheet.TabIndex = 93
        Me.CBExpandSpreadsheet.Text = "Expand Spreadsheet"
        '
        'Panel4
        '
        Me.Panel4.BackColor = System.Drawing.SystemColors.ScrollBar
        Me.Panel4.Controls.Add(Me.VisionMode)
        Me.Panel4.Controls.Add(Me.NeedleMode)
        Me.Panel4.Controls.Add(Me.Label5)
        Me.Panel4.Location = New System.Drawing.Point(852, 320)
        Me.Panel4.Name = "Panel4"
        Me.Panel4.Size = New System.Drawing.Size(428, 50)
        Me.Panel4.TabIndex = 94
        '
        'NeedleMode
        '
        Me.NeedleMode.AutoCheck = False
        Me.NeedleMode.BackColor = System.Drawing.SystemColors.ScrollBar
        Me.NeedleMode.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.NeedleMode.Location = New System.Drawing.Point(286, 16)
        Me.NeedleMode.Name = "NeedleMode"
        Me.NeedleMode.Size = New System.Drawing.Size(96, 24)
        Me.NeedleMode.TabIndex = 89
        Me.NeedleMode.Text = "Needle"
        '
        'TeachingToolbar
        '
        Me.TeachingToolbar.BackColor = System.Drawing.Color.FromArgb(CType(192, Byte), CType(255, Byte), CType(255, Byte))
        Me.TeachingToolbar.Buttons.AddRange(New System.Windows.Forms.ToolBarButton() {Me.TBBOk, Me.TBBCancel})
        Me.TeachingToolbar.ButtonSize = New System.Drawing.Size(42, 42)
        Me.TeachingToolbar.Dock = System.Windows.Forms.DockStyle.None
        Me.TeachingToolbar.DropDownArrows = True
        Me.TeachingToolbar.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.TeachingToolbar.ImageList = Me.ImageListYesNo
        Me.TeachingToolbar.Location = New System.Drawing.Point(768, 320)
        Me.TeachingToolbar.Name = "TeachingToolbar"
        Me.TeachingToolbar.ShowToolTips = True
        Me.TeachingToolbar.Size = New System.Drawing.Size(84, 48)
        Me.TeachingToolbar.TabIndex = 0
        '
        'TBBOk
        '
        Me.TBBOk.ImageIndex = 0
        Me.TBBOk.ToolTipText = "Ok"
        '
        'TBBCancel
        '
        Me.TBBCancel.ImageIndex = 1
        Me.TBBCancel.ToolTipText = "Cancel"
        '
        'EditingToolbar
        '
        Me.EditingToolbar.BackColor = System.Drawing.Color.FromArgb(CType(192, Byte), CType(255, Byte), CType(255, Byte))
        Me.EditingToolbar.Buttons.AddRange(New System.Windows.Forms.ToolBarButton() {Me.TBBSwitch, Me.TBBEdit})
        Me.EditingToolbar.ButtonSize = New System.Drawing.Size(42, 42)
        Me.EditingToolbar.Dock = System.Windows.Forms.DockStyle.None
        Me.EditingToolbar.DropDownArrows = True
        Me.EditingToolbar.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.EditingToolbar.ImageList = Me.ImageListYesNo
        Me.EditingToolbar.Location = New System.Drawing.Point(688, 320)
        Me.EditingToolbar.Name = "EditingToolbar"
        Me.EditingToolbar.ShowToolTips = True
        Me.EditingToolbar.Size = New System.Drawing.Size(84, 48)
        Me.EditingToolbar.TabIndex = 105
        '
        'TBBSwitch
        '
        Me.TBBSwitch.ImageIndex = 2
        Me.TBBSwitch.ToolTipText = "Switch to Next Point"
        '
        'TBBEdit
        '
        Me.TBBEdit.ImageIndex = 3
        Me.TBBEdit.ToolTipText = "Edit Element"
        '
        'ImageListOper
        '
        Me.ImageListOper.ImageSize = New System.Drawing.Size(36, 36)
        Me.ImageListOper.ImageStream = CType(resources.GetObject("ImageListOper.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.ImageListOper.TransparentColor = System.Drawing.Color.Transparent
        '
        'ButtonClean
        '
        Me.ButtonClean.BackColor = System.Drawing.SystemColors.InactiveCaptionText
        Me.ButtonClean.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonClean.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.ButtonClean.ImageIndex = 5
        Me.ButtonClean.ImageList = Me.ImageListGeneralTools
        Me.ButtonClean.Location = New System.Drawing.Point(856, 690)
        Me.ButtonClean.Name = "ButtonClean"
        Me.ButtonClean.Size = New System.Drawing.Size(88, 80)
        Me.ButtonClean.TabIndex = 58
        Me.ButtonClean.Text = "Clean On"
        Me.ButtonClean.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'ButtonNeedleCal
        '
        Me.ButtonNeedleCal.BackColor = System.Drawing.SystemColors.InactiveCaptionText
        Me.ButtonNeedleCal.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonNeedleCal.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.ButtonNeedleCal.ImageIndex = 3
        Me.ButtonNeedleCal.ImageList = Me.ImageListGeneralTools
        Me.ButtonNeedleCal.Location = New System.Drawing.Point(944, 610)
        Me.ButtonNeedleCal.Name = "ButtonNeedleCal"
        Me.ButtonNeedleCal.Size = New System.Drawing.Size(84, 80)
        Me.ButtonNeedleCal.TabIndex = 56
        Me.ButtonNeedleCal.Text = "Needle Calib"
        Me.ButtonNeedleCal.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'ButtonVolCal
        '
        Me.ButtonVolCal.BackColor = System.Drawing.SystemColors.InactiveCaptionText
        Me.ButtonVolCal.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonVolCal.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.ButtonVolCal.ImageIndex = 2
        Me.ButtonVolCal.ImageList = Me.ImageListGeneralTools
        Me.ButtonVolCal.Location = New System.Drawing.Point(1028, 610)
        Me.ButtonVolCal.Name = "ButtonVolCal"
        Me.ButtonVolCal.Size = New System.Drawing.Size(84, 80)
        Me.ButtonVolCal.TabIndex = 55
        Me.ButtonVolCal.Text = "Volume Calib"
        Me.ButtonVolCal.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'ButtonHome
        '
        Me.ButtonHome.BackColor = System.Drawing.SystemColors.InactiveCaptionText
        Me.ButtonHome.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonHome.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.ButtonHome.ImageIndex = 0
        Me.ButtonHome.ImageList = Me.ImageListGeneralTools
        Me.ButtonHome.Location = New System.Drawing.Point(1112, 610)
        Me.ButtonHome.Name = "ButtonHome"
        Me.ButtonHome.Size = New System.Drawing.Size(84, 80)
        Me.ButtonHome.TabIndex = 53
        Me.ButtonHome.Text = "Do Homing"
        Me.ButtonHome.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'ImageListMultiField
        '
        Me.ImageListMultiField.ImageSize = New System.Drawing.Size(36, 28)
        Me.ImageListMultiField.ImageStream = CType(resources.GetObject("ImageListMultiField.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.ImageListMultiField.TransparentColor = System.Drawing.Color.White
        '
        'CBDoorLock
        '
        Me.CBDoorLock.Appearance = System.Windows.Forms.Appearance.Button
        Me.CBDoorLock.BackColor = System.Drawing.SystemColors.InactiveCaptionText
        Me.CBDoorLock.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CBDoorLock.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.CBDoorLock.ImageIndex = 5
        Me.CBDoorLock.ImageList = Me.ImageListMultiField
        Me.CBDoorLock.Location = New System.Drawing.Point(1196, 610)
        Me.CBDoorLock.Name = "CBDoorLock"
        Me.CBDoorLock.Size = New System.Drawing.Size(84, 80)
        Me.CBDoorLock.TabIndex = 117
        Me.CBDoorLock.Text = "Lock Door"
        Me.CBDoorLock.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'NeedleContextMenu
        '
        Me.NeedleContextMenu.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem81, Me.MenuItem82, Me.MenuItem83, Me.MenuItem84, Me.MenuItem85, Me.MenuItem86})
        '
        'MenuItem81
        '
        Me.MenuItem81.Index = 0
        Me.MenuItem81.Text = "&Left"
        '
        'MenuItem82
        '
        Me.MenuItem82.Index = 1
        Me.MenuItem82.Text = "&Right"
        '
        'MenuItem83
        '
        Me.MenuItem83.Index = 2
        Me.MenuItem83.Text = "-"
        '
        'MenuItem84
        '
        Me.MenuItem84.Index = 3
        Me.MenuItem84.Text = "&Undo"
        '
        'MenuItem85
        '
        Me.MenuItem85.Index = 4
        Me.MenuItem85.Text = "&Insert Row"
        '
        'MenuItem86
        '
        Me.MenuItem86.Index = 5
        Me.MenuItem86.Text = "&Delete Row"
        '
        'ImageListLineArc
        '
        Me.ImageListLineArc.ImageSize = New System.Drawing.Size(32, 32)
        Me.ImageListLineArc.ImageStream = CType(resources.GetObject("ImageListLineArc.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.ImageListLineArc.TransparentColor = System.Drawing.Color.Transparent
        '
        'TimerForUpdate
        '
        Me.TimerForUpdate.Interval = 1000
        '
        'AxSpreadsheetProgramming
        '
        Me.AxSpreadsheetProgramming.DataSource = Nothing
        Me.AxSpreadsheetProgramming.Enabled = True
        Me.AxSpreadsheetProgramming.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.AxSpreadsheetProgramming.Location = New System.Drawing.Point(0, 40)
        Me.AxSpreadsheetProgramming.Name = "AxSpreadsheetProgramming"
        Me.AxSpreadsheetProgramming.OcxState = CType(resources.GetObject("AxSpreadsheetProgramming.OcxState"), System.Windows.Forms.AxHost.State)
        Me.AxSpreadsheetProgramming.Size = New System.Drawing.Size(1280, 280)
        Me.AxSpreadsheetProgramming.TabIndex = 128
        '
        'IOCheck
        '
        Me.IOCheck.Interval = 50
        '
        'LabelMessege
        '
        Me.LabelMessege.BackColor = System.Drawing.SystemColors.Info
        Me.LabelMessege.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.LabelMessege.Font = New System.Drawing.Font("Microsoft Sans Serif", 16.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelMessege.ForeColor = System.Drawing.Color.Blue
        Me.LabelMessege.Location = New System.Drawing.Point(0, 320)
        Me.LabelMessege.Name = "LabelMessege"
        Me.LabelMessege.Size = New System.Drawing.Size(688, 42)
        Me.LabelMessege.TabIndex = 84
        Me.LabelMessege.Text = "Message Bar"
        Me.LabelMessege.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'DispensingMode
        '
        Me.DispensingMode.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.DispensingMode.ItemHeight = 20
        Me.DispensingMode.Items.AddRange(New Object() {"Vision Mode", "Dry Needle Mode", "Wet Needle Mode"})
        Me.DispensingMode.Location = New System.Drawing.Point(61, 112)
        Me.DispensingMode.Name = "DispensingMode"
        Me.DispensingMode.Size = New System.Drawing.Size(214, 28)
        Me.DispensingMode.TabIndex = 52
        Me.DispensingMode.Text = "Vision Mode"
        '
        'PBRed
        '
        Me.PBRed.BackColor = System.Drawing.Color.FromArgb(CType(192, Byte), CType(192, Byte), CType(255, Byte))
        Me.PBRed.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.PBRed.Image = CType(resources.GetObject("PBRed.Image"), System.Drawing.Image)
        Me.PBRed.Location = New System.Drawing.Point(64, 32)
        Me.PBRed.Name = "PBRed"
        Me.PBRed.Size = New System.Drawing.Size(209, 56)
        Me.PBRed.TabIndex = 61
        Me.PBRed.TabStop = False
        '
        'PBGreen
        '
        Me.PBGreen.BackColor = System.Drawing.Color.FromArgb(CType(192, Byte), CType(192, Byte), CType(255, Byte))
        Me.PBGreen.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.PBGreen.Image = CType(resources.GetObject("PBGreen.Image"), System.Drawing.Image)
        Me.PBGreen.Location = New System.Drawing.Point(64, 32)
        Me.PBGreen.Name = "PBGreen"
        Me.PBGreen.Size = New System.Drawing.Size(209, 56)
        Me.PBGreen.TabIndex = 63
        Me.PBGreen.TabStop = False
        Me.PBGreen.Visible = False
        '
        'TBOperation
        '
        Me.TBOperation.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.TBOperation.BackColor = System.Drawing.Color.FromArgb(CType(192, Byte), CType(192, Byte), CType(255, Byte))
        Me.TBOperation.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.TBOperation.Buttons.AddRange(New System.Windows.Forms.ToolBarButton() {Me.ToolBarButton1, Me.ToolBarButton2, Me.ToolBarButton5})
        Me.TBOperation.ButtonSize = New System.Drawing.Size(70, 50)
        Me.TBOperation.Dock = System.Windows.Forms.DockStyle.None
        Me.TBOperation.DropDownArrows = True
        Me.TBOperation.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.TBOperation.ImageList = Me.ImageListOper
        Me.TBOperation.Location = New System.Drawing.Point(61, 159)
        Me.TBOperation.Name = "TBOperation"
        Me.TBOperation.ShowToolTips = True
        Me.TBOperation.Size = New System.Drawing.Size(214, 58)
        Me.TBOperation.TabIndex = 0
        '
        'ToolBarButton1
        '
        Me.ToolBarButton1.ImageIndex = 0
        '
        'ToolBarButton2
        '
        Me.ToolBarButton2.Enabled = False
        Me.ToolBarButton2.ImageIndex = 1
        '
        'ToolBarButton5
        '
        Me.ToolBarButton5.Enabled = False
        Me.ToolBarButton5.ImageIndex = 2
        '
        'PBYellow
        '
        Me.PBYellow.BackColor = System.Drawing.Color.FromArgb(CType(192, Byte), CType(192, Byte), CType(255, Byte))
        Me.PBYellow.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.PBYellow.Image = CType(resources.GetObject("PBYellow.Image"), System.Drawing.Image)
        Me.PBYellow.Location = New System.Drawing.Point(64, 32)
        Me.PBYellow.Name = "PBYellow"
        Me.PBYellow.Size = New System.Drawing.Size(209, 56)
        Me.PBYellow.TabIndex = 62
        Me.PBYellow.TabStop = False
        Me.PBYellow.Visible = False
        '
        'PanelMotion
        '
        Me.PanelMotion.BackColor = System.Drawing.Color.FromArgb(CType(192, Byte), CType(192, Byte), CType(255, Byte))
        Me.PanelMotion.Controls.Add(Me.DispensingMode)
        Me.PanelMotion.Controls.Add(Me.PBRed)
        Me.PanelMotion.Controls.Add(Me.PBGreen)
        Me.PanelMotion.Controls.Add(Me.TBOperation)
        Me.PanelMotion.Controls.Add(Me.PBYellow)
        Me.PanelMotion.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.PanelMotion.Location = New System.Drawing.Point(944, 370)
        Me.PanelMotion.Name = "PanelMotion"
        Me.PanelMotion.Size = New System.Drawing.Size(336, 240)
        Me.PanelMotion.TabIndex = 101
        '
        'ButtonToggleMode
        '
        Me.ButtonToggleMode.Location = New System.Drawing.Point(2, 2)
        Me.ButtonToggleMode.Name = "ButtonToggleMode"
        Me.ButtonToggleMode.Size = New System.Drawing.Size(222, 30)
        Me.ButtonToggleMode.TabIndex = 129
        Me.ButtonToggleMode.Text = "Go to Basic Setup"
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
        Me.PanelToBeAdded.Controls.Add(Me.Label6)
        Me.PanelToBeAdded.Controls.Add(Me.LabelStep)
        Me.PanelToBeAdded.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.PanelToBeAdded.Location = New System.Drawing.Point(944, 690)
        Me.PanelToBeAdded.Name = "PanelToBeAdded"
        Me.PanelToBeAdded.Size = New System.Drawing.Size(336, 280)
        Me.PanelToBeAdded.TabIndex = 130
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
        'Label6
        '
        Me.Label6.BackColor = System.Drawing.Color.LightSteelBlue
        Me.Label6.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label6.Location = New System.Drawing.Point(248, 240)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(40, 24)
        Me.Label6.TabIndex = 2
        Me.Label6.Text = "mm"
        Me.Label6.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
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
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(856, 376)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(88, 27)
        Me.TextBox1.TabIndex = 131
        Me.TextBox1.Text = "prev state"
        '
        'TextBox2
        '
        Me.TextBox2.Location = New System.Drawing.Point(856, 408)
        Me.TextBox2.Name = "TextBox2"
        Me.TextBox2.Size = New System.Drawing.Size(88, 27)
        Me.TextBox2.TabIndex = 131
        Me.TextBox2.Text = "curr state"
        '
        'cross_num
        '
        Me.cross_num.Location = New System.Drawing.Point(856, 472)
        Me.cross_num.Name = "cross_num"
        Me.cross_num.Size = New System.Drawing.Size(88, 27)
        Me.cross_num.TabIndex = 131
        Me.cross_num.Text = "0"
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(856, 448)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(88, 24)
        Me.Label2.TabIndex = 132
        Me.Label2.Text = "clear num"
        '
        'ButtonStartFirstStage
        '
        Me.ButtonStartFirstStage.Location = New System.Drawing.Point(856, 612)
        Me.ButtonStartFirstStage.Name = "ButtonStartFirstStage"
        Me.ButtonStartFirstStage.Size = New System.Drawing.Size(88, 80)
        Me.ButtonStartFirstStage.TabIndex = 145
        Me.ButtonStartFirstStage.Text = "Start"
        '
        'ButtonCV_Prod_Release
        '
        Me.ButtonCV_Prod_Release.ForeColor = System.Drawing.SystemColors.ControlText
        Me.ButtonCV_Prod_Release.Image = CType(resources.GetObject("ButtonCV_Prod_Release.Image"), System.Drawing.Image)
        Me.ButtonCV_Prod_Release.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.ButtonCV_Prod_Release.Location = New System.Drawing.Point(856, 824)
        Me.ButtonCV_Prod_Release.Name = "ButtonCV_Prod_Release"
        Me.ButtonCV_Prod_Release.Size = New System.Drawing.Size(88, 56)
        Me.ButtonCV_Prod_Release.TabIndex = 144
        Me.ButtonCV_Prod_Release.Text = "Release"
        Me.ButtonCV_Prod_Release.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'ButtonCV_Prod_Retrieve
        '
        Me.ButtonCV_Prod_Retrieve.ForeColor = System.Drawing.SystemColors.ControlText
        Me.ButtonCV_Prod_Retrieve.Image = CType(resources.GetObject("ButtonCV_Prod_Retrieve.Image"), System.Drawing.Image)
        Me.ButtonCV_Prod_Retrieve.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.ButtonCV_Prod_Retrieve.Location = New System.Drawing.Point(856, 768)
        Me.ButtonCV_Prod_Retrieve.Name = "ButtonCV_Prod_Retrieve"
        Me.ButtonCV_Prod_Retrieve.Size = New System.Drawing.Size(88, 56)
        Me.ButtonCV_Prod_Retrieve.TabIndex = 143
        Me.ButtonCV_Prod_Retrieve.Text = "Retrieve"
        Me.ButtonCV_Prod_Retrieve.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'FormProgramming
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(8, 20)
        Me.ClientSize = New System.Drawing.Size(1276, 927)
        Me.Controls.Add(Me.ButtonStartFirstStage)
        Me.Controls.Add(Me.ButtonCV_Prod_Release)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.PanelToBeAdded)
        Me.Controls.Add(Me.ButtonPurge)
        Me.Controls.Add(Me.ButtonToggleMode)
        Me.Controls.Add(Me.PanelVision)
        Me.Controls.Add(Me.TeachingToolbar)
        Me.Controls.Add(Me.EditingToolbar)
        Me.Controls.Add(Me.ReferenceCommandBlock)
        Me.Controls.Add(Me.ElementsCommandBlock)
        Me.Controls.Add(Me.TextBox1)
        Me.Controls.Add(Me.TextBox2)
        Me.Controls.Add(Me.cross_num)
        Me.Controls.Add(Me.PanelMotion)
        Me.Controls.Add(Me.PanelVisionCtrl)
        Me.Controls.Add(Me.CBExpandSpreadsheet)
        Me.Controls.Add(Me.Panel4)
        Me.Controls.Add(Me.LabelMessege)
        Me.Controls.Add(Me.AxSpreadsheetProgramming)
        Me.Controls.Add(Me.CBDoorLock)
        Me.Controls.Add(Me.ButtonHome)
        Me.Controls.Add(Me.ButtonVolCal)
        Me.Controls.Add(Me.ButtonNeedleCal)
        Me.Controls.Add(Me.ButtonClean)
        Me.Controls.Add(Me.ButtonCV_Prod_Retrieve)
        Me.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.MaximizeBox = False
        Me.Menu = Me.MainMenuProgramming
        Me.MinimizeBox = False
        Me.Name = "FormProgramming"
        Me.Text = "Programming"
        Me.PanelVisionCtrl.ResumeLayout(False)
        CType(Me.ValueBrightness, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel4.ResumeLayout(False)
        CType(Me.AxSpreadsheetProgramming, System.ComponentModel.ISupportInitialize).EndInit()
        Me.PanelMotion.ResumeLayout(False)
        Me.PanelToBeAdded.ResumeLayout(False)
        CType(Me.ComboBoxFineStep, System.ComponentModel.ISupportInitialize).EndInit()
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
    Private m_NewProjectCreated As Boolean = False


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
        Init()
        While isInited = False
            Application.DoEvents()
        End While
        'debugging
        cross_num.Text = CStr(0)

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
        'motion controller
        'm_Tri.Connect_Controller()
        'SetState("Homing")
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
        DisableElementsCommandBlockButton(10)

    End Sub

    Private Sub FormProgramming_Closed(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        KeyboardControl.ReleaseControls()
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

#End Region

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
        SetState("Homing")
    End Sub

    Private Sub ButtonPurge_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonPurge.Click
        If SetState("Purge") Then DoPurge()
    End Sub

    Private Sub ButtonClean_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonClean.Click
        If SetState("Clean") Then DoClean()
    End Sub

    Private Sub ButtonNeedleCal_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonNeedleCal.Click
        SetState("Needle Calibration")
    End Sub

    Private Sub ButtonVolCal_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonVolCal.Click
        SetState("Volume Calibration")
    End Sub

    Private Sub VisionMode_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles VisionMode.Click

        If IsBusy() Then
            LabelMessage("Can't change mode when machine is running!")
            Exit Sub
        End If

        VisionMode.Checked = Not VisionMode.Checked
        NeedleMode.Checked = Not NeedleMode.Checked

        EnableReferenceCommandBlock()
        EnableElementsCommandBlock()

        MoveZToSafePosition()

    End Sub

    Private Sub NeedleMode_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NeedleMode.Click

        If IsBusy() Then
            LabelMessage("Can't change mode when machine is running!")
            Exit Sub
        End If

        MoveZToSafePosition()

        VisionMode.Checked = Not VisionMode.Checked
        NeedleMode.Checked = Not NeedleMode.Checked

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
        If CBDoorLock.Checked = False Then
            CBDoorLock.Text = "Unlock Door"
            CBDoorLock.ImageIndex = 6
            TraceDoEvents()
            LockDoor()
        Else
            CBDoorLock.Text = "Lock Door"
            CBDoorLock.ImageIndex = 5
            TraceDoEvents()
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

    'run program
    Private Sub TBOperation_ButtonClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ToolBarButtonClickEventArgs) Handles TBOperation.ButtonClick

        Try
            If e.Button Is TBOperation.Buttons(0) Then 'run

                SetState("Start")

            ElseIf e.Button Is TBOperation.Buttons(1) Then 'pause

                PauseDispensing()

            ElseIf e.Button Is TBOperation.Buttons(2) Then 'stop

                StopDispensing()
            End If

        Catch ex As SystemException
            ExceptionDisplay(ex)
        End Try
    End Sub

    Private Sub MoveToSpreadsheetPoint(ByVal Pos() As Double, ByVal type As String)

        If IsBusy() Then
            LabelMessage("Can't move when program is running!")
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
        End If

        If Not m_Tri.Move_Z(SafePosition) Then Exit Sub
        If Not m_Tri.Move_XY(Pos) Then Exit Sub

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
            SavePatternFileDialog.InitialDirectory = "C:\IDS\Pattern_Dir"
            SavePatternFileDialog.AddExtension = True
            SavePatternFileDialog.Filter = "Create a new project folder|"
            SavePatternFileDialog.FileName() = ""

            SavePatternFileDialog.ShowDialog()
            Dim file As String
            Dim folder As String = SavePatternFileDialog.FileName()

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
                    'Create folder for the project
                End If

                System.IO.Directory.CreateDirectory(folder)
                file = m_Execution.m_File.NameOnlyFromFullPath(folder)
                gPatternFileName = folder + "\" + file

                gFidFileName = gPatternFileName


                'Set ToolButton status
                EnableTeachingButtons()
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

                m_EditStateFlag = False
                MyConveyorSettings.InitializeConveyorSetup()
                Disp_Dispenser_Unit_info()

                SystemSetupDataRetrieve() 'SJ add 

                '   Set cursor position.                        '
                SelectCell(m_Row, 1)
                SaveProgram.UnSave = True
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

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
        SavePatternFileDialog.Filter = "Excel Pattern files (*.Xls)|*.Xls|Txt Pattern files (*.ptp)|*.ptp"
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
            If "Xls" = strTmp(1) Then
                ExportExcelPatternFile(file)
            ElseIf "ptp" = strTmp(1) Then
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
            If 0 = m_Execution.m_Pattern.LoadTxtPatternPara(AxSpreadsheetProgramming, file, 2, m_Row, False) Then
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
        OpenPatternFileDialog.Filter = "Excel pattern files (*.Xls)|*.Xls|Txt pattern files(*.ptp)|*.ptp"
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

        If "Xls" = strInfo(1) Then
            ImportExcelPatternFile(file)
        ElseIf "ptp" = strInfo(1) Then
            ImportTxtPatternFile(file)
        End If
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
            Production.RichTextBoxNote.Text = ""
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

        LabelMessage("Please wait, system is uploading.....")

        DialogPreview.ShowDialog()
        Dim file = DialogPreview.Path

        If Nothing = file Then
            Return
        End If

        'AxSpreadsheetProgramming.ActiveSheet.UsedRange().Clear()
        m_NewProjectCreated = True
        gPatternFileName = ""
        gFidFileName = ""
        m_Row = 0
        m_EditStateFlag = False

        init_spreadsheet()
        FlushSpreadsheet()
        ErrorSubSheetStructIni(200, 500)

        Try
            Me.Cursor = Cursors.WaitCursor
            'IDS.StopErrorCheck() 'kr?
            AxSpreadsheetProgramming.Caption = file
            m_Execution.m_Pattern.LoadTxtPatternPara(AxSpreadsheetProgramming, file, 0, 0, False)
            Me.Cursor = Cursors.WaitCursor
            gPatternFileName = m_Execution.m_File.FolderWithNameFromFileName(file)
            gFidFileName = gPatternFileName
            m_Row = 2
            EnableTeachingButtons()

            DisableElementsCommandBlockButton(gOffsetCmdIndex)
            IDS.Data.OpenPathFileData(gPatternFileName + ".pat")
            MyConveyorSettings.InitializeConveyorSetup()
            SystemSetupDataRetrieve() 'SJ add 
            'IDS.StartErrorCheck() 'kr?
            IDS.newOpen = True

            'Acitvate the "Main" page
            AxSpreadsheetProgramming.Worksheets("Main").Activate()

            LabelMessage("Please wait, Checking content.....")
            'Error checking for all the Spreadsheet
            If 0 <> m_Execution.m_Pattern.m_ErrorChk.CheckAllError(AxSpreadsheetProgramming, ErrorSubSheet) Then

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
            End If

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

        LabelMessage("Finish.....")

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
                        IDS.Data.SaveData()

                        SaveProgram.UnSave = False
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

                SaveProgram.UnSave = False
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

        Return 'Disable undo operation. this function is terribely wrong...

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
            m_Execution.m_Undo.UndoFilename = "C:\IDS\Pattern_Dir\SysSwapData\UndoStep_A.Xls"
            m_Execution.m_Undo.CurrentPageName_A = GetActiveSheetName()
            m_Execution.m_Undo.DataSaveFor_Undo(AxSpreadsheetProgramming)

            m_Execution.m_Undo.UndoFilename = "C:\IDS\Pattern_Dir\SysSwapData\UndoStep_B.Xls"
            FlushSpreadsheet()
            m_Execution.m_Undo.DataLoadFor_Undo(AxSpreadsheetProgramming)

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
            m_Execution.m_Undo.UndoFilename = "C:\IDS\Pattern_Dir\SysSwapData\UndoStep_B.Xls"
            m_Execution.m_Undo.CurrentPageName_B = GetActiveSheetName()
            m_Execution.m_Undo.DataSaveFor_Undo(AxSpreadsheetProgramming)

            m_Execution.m_Undo.UndoFilename = "C:\IDS\Pattern_Dir\SysSwapData\UndoStep_A.Xls"
            FlushSpreadsheet()
            m_Execution.m_Undo.DataLoadFor_Undo(AxSpreadsheetProgramming)

            AxSpreadsheetProgramming.Worksheets(m_Execution.m_Undo.CurrentPageName_A).Activate()
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
                UndoData_Logging(0)

                m_Execution.m_Pattern.LoadTxtPatternParaAllSheets(AxSpreadsheetProgramming, _
                    "C:\IDS\Pattern_Dir\SysSwapData\CopyPaste.txt", _
                    CopiedSheetName, 2, m_StartRow, False)
            End If
        End If
        TraceGCCollect()
    End Sub

    'menu: Edit-->Delete: Delecte all select                                                                        

    Private Sub MenuEditDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuEditDelete.Click
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

            OffLaser()

            'stepper panel
            MySettings.PanelLeft.Controls.Add(m_Tri.SteppingButtons)
            Controls.Remove(m_Tri.SteppingButtons)
            m_Tri.SteppingButtons.Location = New Point(384, 12)
            m_Tri.SteppingButtons.BringToFront()
            m_Tri.SteppingButtons.Show()

            CurrentMode = "Basic Setup"
            ButtonToggleMode.Text = "Go to Program Editor"
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

        ElseIf CurrentMode = "Basic Setup" Then

            MySettings.RemoveCurrentPanel()

            'in case you didn't exit conveyor settings but straightaway pressed program editor
            Conveyor.PositionTimer.Stop()

            'stepper panel
            Controls.Add(m_Tri.SteppingButtons)
            MySettings.PanelLeft.Controls.Remove(m_Tri.SteppingButtons)
            m_Tri.SteppingButtons.Location = PanelToBeAdded.Location()
            m_Tri.SteppingButtons.BringToFront()
            m_Tri.SteppingButtons.Show()

            CurrentMode = "Program Editor"
            ButtonToggleMode.Text = "Go to Basic Setup"
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

    Private Sub EditingToolbar_ButtonClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ToolBarButtonClickEventArgs) Handles EditingToolbar.ButtonClick

        Dim idFlag As Integer = 0
        If e.Button Is EditingToolbar.Buttons(0) Then
            idFlag = 1
        End If

        EditingToolbar_Implementation(idFlag)
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
            DisableCoordinateUpdateInSpreadsheet()
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


                If CmdName = "ChipEdge" Or CmdName = "QC" Or CmdName = "Height" Or CmdName = "Fiducial" Or CmdName = "Reject" Then
                    MoveToSpreadsheetPoint(pos, "Vision")
                Else
                    MoveToSpreadsheetPoint(pos, "Needle")
                End If

            End If
            DisableTeachingToolbarOKButton()
            EnableTeachingToolbarCancelButton()
            EnableEditingToolbarEditButton() 'Edit start
            DisableTeachingButtons()
            UndoData_Logging(1)

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
            tempPosX = GetCellValue(m_Row, gPos1XColumn)
            tempPosY = GetCellValue(m_Row, gPos1YColumn)
            tempPosZ = GetCellValue(m_Row, gPos1ZColumn)

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
                If Cmd.DispenseFlag = "True" Then
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
    End Sub

    Private Sub TBrefer_Implementation(ByVal ButtonText As String)
        Dim cell1, cell2 As Object

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
        m_EditStateFlag = False
        m_ProgrammingStateFlag = True

        Dim SheetName As String = GetActiveSheetName()
        If m_Execution.m_Pattern.Spreadsheet_IsAnArraySheet(AxSpreadsheetProgramming, SheetName) Then
            MessageBox.Show("Command is not allowed in an Array sheet", "Warnning!")
        Else
            If "Measure" <> buttonText Then
                UndoData_Logging(0)
            End If
            TeachElementCommand(buttonText)
        End If

    End Sub

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
                    End If
                    m_ProgrammingStateFlag = False

                Case "Measure"

                Case "Dot", "Move", "Line", "Arc", "Circle", "Rectangle", _
                    "FillCircle", "FillRectangle", "ChipEdge", "QC", "SubPattern", "Wait", "EdgeDetection", "DotArray"

                    type = ButtonText
                    m_TeachStepNumber = 1
                    m_Row = GetActiveCellRow()
                    AddCommandToSpreadsheet(type)

                    If m_InLinkRange Then

                    Else
                        cell1 = GetCell(m_Row, gPos1XColumn)
                        cell2 = GetCell(m_Row, gPos1ZColumn)
                        SelectRange(cell1, cell2) 'Select the related cell in spreadsheet
                    End If

                    ToolBarSwitch("YesNo")

                    If ButtonText = "ChipEdge" Then
                        'DisableTeachingButtons()
                        'DisableTeachingToolbar()
                        DisableElementsCommandBlockButton(gOffsetCmdIndex)
                        DisableProgrammingBrightnessToggle()
                        Vision.IDSV_Form_CE(ValueBrightness.Value)
                        Dim status As Integer = 0
                        Dim x, y As Double
                        Do
                            While Not (Vision.RobotMotionOffset(x, y) = True Or Vision.GetChipEdgeStatus = 2)
                                TraceDoEvents()
                            End While
                            'moverobot
                            pos(0) = x
                            pos(1) = -y
                            pos(2) = 0

                            m_Tri.Set_XY_Speed(IDS.Data.Hardware.Gantry.ElementXYSpeed)
                            m_Tri.MoveRelative_XY(pos)
                            If status = 3 Then
                                status = 0 'reset status after 5 points being reset
                            End If
                            While status = 0
                                TraceDoEvents()
                                status = Vision.GetChipEdgeStatus()
                            End While
                        Loop While status = 3 'status 3= reset 5 points

                        DelayForRowDelete()
                        DisableCoordinateUpdateInSpreadsheet()

                        If status = 2 Then
                            DeleteRow(m_Row)
                            UpdateSpreadsheet()
                            DeletingRowFromExcel = False
                            DeletingRowFinished = False
                        ElseIf status = 1 Then 'chipedge finished settings
                            SetChipEdgeSettings()
                            ElementsCommandBlock.Enabled = True
                            ReferenceCommandBlock.Enabled = True
                            DisplaySpreadsheetTabs()
                            SelectCell(m_Row + 1, 1)
                        End If
                        ToggleButtonsForTeachingStop()
                        ClearAndDisplayIndicator()
                        EnableTeachingButtons()
                        EnableTeachModeSwitching()
                        m_ProgrammingStateFlag = False
                        EnableProgrammingBrightnessToggle()
                        DisableElementsCommandBlockButton(gSeperatorCmdIndex)
                        DisableElementsCommandBlockButton(gOffsetCmdIndex)
                        DisableTeachingToolbarOKButton()
                        EnableTeachingToolbarCancelButton()
                        DisplaySpreadsheetTabs()

                    ElseIf ButtonText = "QC" Then
                        DisableTeachingButtons()
                        DisableElementsCommandBlockButton(gOffsetCmdIndex)
                        DisableTeachingToolbar()
                        DisableProgrammingBrightnessToggle()
                        Vision.IDSV_Form_QC(ValueBrightness.Value)
                        Dim status As Integer = 0
                        Try
                            While status = 0 Or status = 3
                                Do
                                    TraceDoEvents()
                                    status = Vision.GetQCStatus()
                                Loop While status = 3
                            End While
                        Catch ex As Exception
                            ExceptionDisplay(ex)
                        End Try

                        If status = 2 Then 'Cancel
                            DelayForRowDelete()
                            DisableCoordinateUpdateInSpreadsheet()
                            DeleteRow(m_Row)
                            UpdateSpreadsheet()
                            DeletingRowFromExcel = False
                            DeletingRowFinished = False
                        ElseIf status = 1 Then 'Ok
                            SetQCSettings()
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

                Case "Array"
                    CBExpandSpreadsheet.Enabled = False

                    DisableTeachingToolbarOKButton()
                    DisableTeachingToolbarCancelButton()
                    type = ButtonText
                    'sj modify 28/09/09
                    m_Row = GetActiveCellRow()

                    AddCommandToSpreadsheet(type)

                    EnableCoordinateUpdateInSpreadsheet()
                    Dim ArrayDlg As New ArrayGenerate
                    ArrayDlg.SetDefaultPara(m_Execution.m_Pattern.m_CurrentDPara)
                    ArrayDlg.BringToFront()
                    ArrayDlg.Show()
                    Dim DlgReturn = ArrayDlg.DialogResult()
                    Dim PointX, PointY, PointZ As Double

                    Do
                        TraceDoEvents()
                        PointX = CDbl(GetCellValue(m_Row, gPos1XColumn))
                        PointY = CDbl(GetCellValue(m_Row, gPos1YColumn))
                        PointZ = CDbl(GetCellValue(m_Row, gPos1ZColumn))

                        ArrayDlg.SetPoint(PointX, PointY, PointZ)
                        DlgReturn = ArrayDlg.DialogResult()
                    Loop While Nothing = DlgReturn

                    DisableCoordinateUpdateInSpreadsheet()

                    If DialogResult.OK = DlgReturn Then
                        'Get "Array" data succesfully.  Data will be used to generate "Array"
                        LabelMessage("Array generating starts")
                        Dim PatternLineRecord(3) As CIDSPattern.PatternRecord
                        Dim arraydata As ArrayData
                        arraydata.FirstX = 10.0
                        Dim dotarray As New FormArraySetting(arraydata)

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
                            Spreadsheet_AddSubandArrayInSub(PatternLineRecord(0).pPara.DispenseFlag)

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
                    m_ProgrammingStateFlag = False

                    CBExpandSpreadsheet.Enabled = True
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

        Dim row As Integer = GetActiveCellRow()
        Dim colum As Integer = GetActiveCellColumn()
        Dim cellValue As Object = GetActiveCellValue()
        DisableCoordinateUpdateInSpreadsheet()

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
                    Exit Sub
                Else
                    SetCellValue(row, colum, cellValue)
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
            (colum >= gPos3XColumn And colum <= gPos3YColumn) Then ' And _
            '       (row > 0) And ( ) Then

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
                    If CmdStr = "ChipEdge" Or CmdStr = "QC" Or CmdStr = "Height" Or CmdStr = "Fiducial" Or CmdStr = "Reject" Then
                        MoveToSpreadsheetPoint(pos, "Vision")
                    Else
                        MoveToSpreadsheetPoint(pos, "Needle")
                    End If
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
                    MyMsgBox("Found edit error, click Ok to restore", MsgBoxStyle.Information + MsgBoxStyle.OKOnly, "Input error")
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

        Dim sel As Microsoft.Office.Interop.OWC.Range = AxSpreadsheetProgramming.Selection

        Dim m_rowCount As Integer = sel.Rows.Count()
        Dim m_columnCount As Integer = sel.Columns.Count()
        Dim m_StartRow As Integer = sel.Row
        Dim m_columnStart As Integer = sel.Column
        Dim m_EndRow As Integer = m_StartRow + m_rowCount - 1
        Dim RefPos() As Double = {0, 0, 0}

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

        If True = m_EditStateFlag Then
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

        ElseIf True = m_ProgrammingStateFlag Then
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
                    SelectRange(cell1, cell2)
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
            Spreadsheet_SetMotionRef(RefPos)
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
            Else
                DisableElementsCommandBlockButton(gOffsetCmdIndex)
                DisableEditingToolbar()
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
                        MoveToSpreadsheetPoint(pos, "Vision")
                    Else
                        MoveToSpreadsheetPoint(pos, "Needle")
                    End If

                    AxSpreadsheetProgramming.Enabled = True
                    HideSpreadsheetTabs()
                    AxSpreadsheetProgramming.Refresh()

                    DisableElementsCommandBlock()
                    EnableElementsCommandBlockButton(gOffsetCmdIndex)         'Offset
                    EnableEditingToolbarSwitchButton()                     'Goto next
                    DisableEditingToolbarEditButton()                    'Edit start
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
                    EnableElementsCommandBlockButton(gOffsetCmdIndex)   'Offset enabled

                    'For the valid Offset selections, we get the motion ref for the last row
                    m_Execution.m_Pattern.Spreadsheet_GetRowLocalReference(AxSpreadsheetProgramming, m_EndRow, RefPos)
                    Spreadsheet_SetMotionRef(RefPos)

                    'Move robot arm to the point last valid point
                    pos(0) = pos(0)
                    pos(1) = pos(1)
                    pos(2) = pos(2)

                    AxSpreadsheetProgramming.Enabled = False
                    MoveToSpreadsheetPoint(pos, "Vision")
                    AxSpreadsheetProgramming.Enabled = True
                Else
                    'For the invalid Offset selections, we get the motion ref for the first row
                    m_Execution.m_Pattern.Spreadsheet_GetRowLocalReference(AxSpreadsheetProgramming, m_Row, RefPos)
                    Spreadsheet_SetMotionRef(RefPos)

                    DisableElementsCommandBlockButton(gOffsetCmdIndex)
                End If
                DisableEditingToolbar()
            End If

        ElseIf 1 = m_columnCount Then   'Update elements column-wise, need to be improved later to
            m_Column = sel.Column       'exclude part of data
            DisableTeachingToolbar()
            DisableEditingToolbar()
            DisableTeachingButtons()

            If m_Column < gTravelSpeedColumn Or m_Column > gSprialColumn Or m_rowCount < 1 Then
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
            If "" = StrTmp Then
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
        Dim m_rowCount As Integer = sel.Rows.Count()
        Dim m_rowLocal As Integer = sel.Row

        Dim DeltedRowNo As Integer = 0
        Dim DeletedRowNoEachtime As Integer = 1

        If m_rowCount < 1 Then Return

        Do
            Spreadsheet_DeleteOneRow(DeletedRowNoEachtime, m_rowLocal, AxSpreadsheet)

            DeltedRowNo = DeltedRowNo + DeletedRowNoEachtime
        Loop Until (DeltedRowNo >= m_rowCount)

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

            'If (chkDualHead.Checked = False) Then
            '    oMenu1 = New Object() {"&Left", "Left"}
            '    oMenu = New Object() {oMenu1}
            'Else
            '    oMenu1 = New Object() {"&Left", "Left"}
            '    oMenu2 = New Object() {"&Right", "Right"}
            '    oMenu = New Object() {oMenu1, oMenu2}
            'End If

        ElseIf AxSpreadsheetProgramming.ActiveCell.Column = gDispensColumn Then
            oMenu1 = New Object() {"&On", "On"}
            oMenu2 = New Object() {"&Off", "Off"}
            'oMenu3 = New Object() {"____________", ""}
            'oMenu4 = New Object() {"&Undo", "Undo"}
            'oMenu5 = New Object() {"&Insert Row", "Insert"}
            'oMenu6 = New Object() {"&Delete Row", "Delete"}
            'oMenu = New Object() {oMenu1, oMenu2, oMenu3, oMenu4, oMenu5, oMenu6}
            oMenu = New Object() {oMenu1, oMenu2}
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


    Private Sub ToolBarYesNo_ButtonClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ToolBarButtonClickEventArgs) Handles TeachingToolbar.ButtonClick

        offset(0) = gLeftNeedleOffs(0)
        offset(1) = gLeftNeedleOffs(1)
        offset(2) = gLeftNeedleOffs(2)

        If e.Button Is TeachingToolbar.Buttons(0) Then
            If type <> "Move" Then
                If CheckSoftLimitXYZ(m_MachinePos, offset) Then Exit Sub
            End If
            Confirm()   'Add rows
        Else
            If m_EditStateFlag = False Then
                DelayForRowDelete()
                Cancel()   'Cancel or delete rows
                DeletingRowFromExcel = False
                DeletingRowFinished = False
            Else
                Cancel()   'Cancel or delete rows
            End If
        End If

        EnableTeachModeSwitching()

    End Sub

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

        Try
            If m_rowCount = 1 Then
                type = GetCellValue(m_Row, gCommandNameColumn)
                UndoData_Logging(0)

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
                            posOffset(2) = Laser.MM_Reading - IDS.Data.Hardware.HeightSensor.Laser.HeightReference
                            SetCellValue(m_Row, gPos1ZColumn, CInt(posOffset(2)))
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
                        SelectCell(m_Row + 1, gCommandNameColumn)
                    ElseIf "FIDUCIAL" = CmdName.ToUpper Then
                        'For Lim's part

                        Dim brightness As Integer

                        If m_TeachStepNumber = 1 Then

                            SwitchToVisionTeachMode()
                            DisableTeachModeSwitching()

                            brightness = GetCellValue(m_Row, gBrightnessColumn)
                            DisableProgrammingBrightnessToggle()
                            Vision.IDSV_Form_FI_Edit(1, gPatternFileName + "1.mmo", brightness)
                            While status = 0
                                TraceDoEvents()
                                status = Vision.GetFiducialStatus()
                            End While
                            Dim fidname As String
                            If status = 1 Then
                                fidname = Vision.GetFiducialFilename()
                                Dim Brightness1 As Integer = Vision.GetFiducialBrightness
                                SetCellValue(m_Row, gFid1Column, fidname)
                                SetCellValue(m_Row, gBrightnessColumn, Brightness1)
                            ElseIf 2 = status Then  'Cancel command
                                DisableCoordinateUpdateInSpreadsheet()
                                m_TeachStepNumber = 0
                                ToggleButtonsForTeachingStop()
                                DisableElementsCommandBlockButton(gSeperatorCmdIndex)
                            ElseIf 3 = status Then
                                fidname = Vision.GetFiducialFilename()
                                SetCellValue(m_Row, gFid1Column, fidname)
                            End If
                        ElseIf m_TeachStepNumber = 2 Then
                            brightness = GetCellValue(m_Row, gThresholdColumn) 'fiducial 2 brightness value
                            Vision.FrmVision.SetBrightness(brightness)
                            DisableProgrammingBrightnessToggle()
                            Vision.IDSV_Form_FI_Edit(2, gPatternFileName + "2.mmo", brightness)
                            While status = 0
                                TraceDoEvents()
                                status = Vision.GetFiducialStatus()
                            End While
                            Dim fidname As String
                            If status = 1 Then
                                fidname = Vision.GetFiducialFilename()
                                Dim Brightness2 As Integer = Vision.GetFiducialBrightness
                                SetCellValue(m_Row, gFid2Column, fidname)
                                SetCellValue(m_Row, gThresholdColumn, Brightness2) 'For brightness fiducial no.2
                            ElseIf 2 = status Then
                            End If
                        End If
                        EnableProgrammingBrightnessToggle()
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
                        DisableProgrammingBrightnessToggle()
                        Vision.IDSV_Form_CE_Edit(vPara)

                        Dim x, y As Double
                        Dim pos(3) As Double

                        Do
                            While Not (Vision.RobotMotionOffset(x, y) = True Or Vision.GetChipEdgeStatus = 2 Or Vision.GetChipEdgeStatus = 1)
                                TraceDoEvents()
                            End While

                            'moverobot
                            pos(0) = x
                            pos(1) = -y
                            pos(2) = 0

                            m_Tri.Set_XY_Speed(IDS.Data.Hardware.Gantry.ElementXYSpeed)
                            m_Tri.MoveRelative_XY(pos)

                            If status = 3 Then
                                status = 0 'reset status after 5 points being reset
                            End If

                            While status = 0
                                TraceDoEvents()
                                status = Vision.GetChipEdgeStatus()
                            End While

                        Loop While status = 3 'status 3= reset 5 points

                        DisableCoordinateUpdateInSpreadsheet()
                        ElementsCommandBlock.Enabled = True
                        ReferenceCommandBlock.Enabled = True
                        EnableTeachingButtons()
                        EnableProgrammingBrightnessToggle()
                        DisableElementsCommandBlockButton(gSeperatorCmdIndex)
                        m_ProgrammingStateFlag = False
                        If status = 2 Then
                            UpdateSpreadsheet()
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
                            cell1 = GetCell(m_Row, gPos1XColumn)
                            cell2 = GetCell(m_Row, gPos1ZColumn)
                            m_Execution.m_Pattern.Spreadsheet_CellGrey(cell1, cell2, True, AxSpreadsheetProgramming)
                            SelectCell(m_Row, 1)
                        End If

                    ElseIf "QC" = CmdName.ToUpper Then 'lim
                        EnableCoordinateUpdateInSpreadsheet()

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
                        Vision.IDSV_Form_QC_Edit(vPara)

                        While status = 0 Or status = 3
                            Do
                                TraceDoEvents()
                                status = Vision.GetQCStatus()
                            Loop While status = 3
                        End While
                        DisableCoordinateUpdateInSpreadsheet()
                        EnableProgrammingBrightnessToggle()
                        EnableTeachingButtons()
                        DisableElementsCommandBlockButton(gSeperatorCmdIndex)
                        m_ProgrammingStateFlag = False
                        If status = 2 Then 'Cancel
                            SelectCell(m_Row, 1)
                            Vision.SetQCReset()
                        ElseIf status = 1 Then                         'Ok
                            SetQCSettings()
                            SelectCell(m_Row + 1, 1)
                        End If
                    End If
                    MenuEditCopy.Enabled = True
                    If "" <> CopiedSheetName Then
                        MenuEditPaste.Enabled = True
                    End If
                    MenuEditCut.Enabled = True
                    MenuEditUndo.Enabled = True
                    MenuEditRedo.Enabled = False
                    SaveProgram.UnSave = True

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
                UndoData_Logging(0)
                Spreadsheet_DeleteMultiRow(AxSpreadsheetProgramming)
                SaveProgram.UnSave = True
            End If

            m_EditStateFlag = False
            Spreadsheet_CheckForWithinLinkRange(True)
            DisableTeachingToolbarOKButton()
            EnableTeachingToolbarCancelButton()
            DisableEditingToolbar()
            DisableElementsCommandBlockButton(gSeperatorCmdIndex)
            DisableElementsCommandBlockButton(gOffsetCmdIndex)
            DisplaySpreadsheetTabs()
            If status = 2 Then
                SetCellValue(m_Row, gPos1XColumn, tempPosX)
                SetCellValue(m_Row, gPos1YColumn, tempPosY)
                SetCellValue(m_Row, gPos1ZColumn, tempPosZ)
                SelectCell(m_Row, gCommandNameColumn)
            ElseIf status = 1 Then      'Ok
                SelectCell(m_Row + 1, gCommandNameColumn)
            End If
            DisableCoordinateUpdateInSpreadsheet() 'for delete
            UpdateSpreadsheet()
            ClearAndDisplayIndicator()

        Catch ex As SystemException
            'kr should i put here
            ExceptionDisplay(ex)
        End Try
        TraceGCCollect()

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
            posOffset = Laser.MM_Reading - IDS.Data.Hardware.HeightSensor.Laser.HeightReference
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

                Case "REJECT"
                    ConfirmReject()

                Case "HEIGHT"
                    ConfirmHeight()

                Case "SUBPATTERN"
                    ConfirmSubPattern()

                Case "DOT", "REFERENCE", "QC", "MOVE", "CHIPEDGE", "WAIT"
                    ConfirmSinglePointElement()

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
                            Vision.IDSV_Form_FI(1, ValueBrightness.Value)
                            Dim status As Integer = 0 'status: 1=ok, 2=cancel, 3=end with first fiducial
                            While status = 0
                                TraceDoEvents()
                                status = Vision.GetFiducialStatus()
                            End While
                            Dim fidname As String
                            If status = 1 Then
                                fidname = Vision.GetFiducialFilename()
                                Dim Brightness1 As Integer = Vision.GetFiducialBrightness
                                SetCellValue(m_Row, gFid1Column, fidname)
                                SetCellValue(m_Row, gBrightnessColumn, Brightness1)
                                LabelMessage("Please confirm the 2nd Fiducial pt")
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
                            EnableProgrammingBrightnessToggle()
                        End If

                    ElseIf m_TeachStepNumber = 2 Then
                        m_ProgrammingStateFlag = False
                        cell1 = GetCell(m_Row, gPos2XColumn)
                        cell2 = GetCell(m_Row, gPos2ZColumn)
                        x2 = GetCellValue(m_Row, gPos2XColumn)
                        y2 = GetCellValue(m_Row, gPos2YColumn)
                        If type.ToUpper = "LINE" Then
                        ElseIf type.ToUpper = "FIDUCIAL" Then
                            DisableProgrammingBrightnessToggle()
                            Vision.IDSV_Form_FI(2, ValueBrightness.Value)
                            Dim status As Integer = 0
                            While status = 0
                                TraceDoEvents()
                                status = Vision.GetFiducialStatus()
                            End While
                            Dim fidname As String
                            If status = 1 Then
                                fidname = Vision.GetFiducialFilename()
                                SetCellValue(m_Row, gFid2Column, fidname)
                                Dim Brightness2 As Integer = Vision.GetFiducialBrightness
                                SetCellValue(m_Row, gThresholdColumn, Brightness2) 'For brightness fiducial no.2
                            ElseIf 2 = status Then
                                SetCellValue(m_Row, gPos2XColumn, GetCellValue(m_Row, gPos1XColumn))
                                SetCellValue(m_Row, gPos2YColumn, GetCellValue(m_Row, gPos1YColumn))
                            End If
                        End If
                        UpdateLinkForNextRow()
                        DisableCoordinateUpdateInSpreadsheet()
                        m_Execution.m_Pattern.Spreadsheet_CellGrey(cell1, cell2, True, AxSpreadsheetProgramming)
                        SelectCell(m_Row + 1, 1)
                    End If

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
                    If type.ToUpper = "FIDUCIAL" Then EnableProgrammingBrightnessToggle()

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
        cell1 = GetCell(m_Row, gPos1XColumn)
        cell2 = GetCell(m_Row, gPos1ZColumn)
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

    Private Sub ButtonCV_Prod_Retrieve_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonCV_Prod_Retrieve.Click
        Conveyor.Command("Retrieve")
    End Sub

    Private Sub ButtonCV_Prod_Release_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonCV_Prod_Release.Click
        Conveyor.Command("Release")
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
End Class
