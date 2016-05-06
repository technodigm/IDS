Imports DLL_DataManager
Imports DLL_Export_Devices
Imports DLL_Export_Device_Motor
Imports DLL_Export_Service
Imports Microsoft.DirectX.DirectInput
Imports DLL_SetupAndCaliberate
Imports System.Threading
Imports Microsoft.win32
Imports Microsoft.Office.Interop
Imports System.Messaging
Imports Microsoft.VisualBasic.ComClassAttribute
Imports System.Math
Imports System.IO

'<ComVisibleAttribute(True)> Public NotInheritable Class Monitor

Public Class FormProgramming
    Inherits System.Windows.Forms.Form
    Private m_TriProgMemHandle As TrioMemClass.ClassTrioMem = DLL_SetupAndCaliberate.MySSMain.m_TriMemHandle
    Private Delegate Sub ButtonSwitchControl(ByVal switchID As Integer)
    Private bswitchCtrl As ButtonSwitchControl = New ButtonSwitchControl(AddressOf DealSwitchControl)

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
    Friend WithEvents MenuItem23 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem24 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem25 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem29 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem30 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem32 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem33 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem34 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem35 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem36 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem37 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem45 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem48 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem62 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem63 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem64 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem65 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem66 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem67 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem68 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem69 As System.Windows.Forms.MenuItem
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents DomainUpDown1 As System.Windows.Forms.DomainUpDown
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents ImageListFiles As System.Windows.Forms.ImageList
    Friend WithEvents imageListElement As System.Windows.Forms.ImageList
    Friend WithEvents ImageListReference As System.Windows.Forms.ImageList
    Friend WithEvents TBrefer As System.Windows.Forms.ToolBar
    Friend WithEvents TBFiducialPt As System.Windows.Forms.ToolBarButton
    Friend WithEvents TBHeightPt As System.Windows.Forms.ToolBarButton
    Friend WithEvents TBReferencePt As System.Windows.Forms.ToolBarButton
    Friend WithEvents TBElements As System.Windows.Forms.ToolBar
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents LabelMessege As System.Windows.Forms.Label
    Friend WithEvents ImageListYesNo As System.Windows.Forms.ImageList
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Panel4 As System.Windows.Forms.Panel
    Public WithEvents PanelMotion As System.Windows.Forms.Panel
    Friend WithEvents Label27 As System.Windows.Forms.Label
    Friend WithEvents MenuItem71 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem72 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem73 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem76 As System.Windows.Forms.MenuItem
    Friend WithEvents TBRejectPt As System.Windows.Forms.ToolBarButton
    Friend WithEvents CBExpandSpreadsheet As System.Windows.Forms.CheckBox
    Friend WithEvents RBLNeedleMode As System.Windows.Forms.RadioButton
    Friend WithEvents RBRNeedleMode As System.Windows.Forms.RadioButton
    Friend WithEvents RBVisionMode As System.Windows.Forms.RadioButton
    Friend WithEvents TBBOk As System.Windows.Forms.ToolBarButton
    Friend WithEvents TBBCancel As System.Windows.Forms.ToolBarButton
    Friend WithEvents TBBSwitch As System.Windows.Forms.ToolBarButton
    Friend WithEvents TBBEdit As System.Windows.Forms.ToolBarButton
    Friend WithEvents TBBEditOnly As System.Windows.Forms.ToolBarButton
    '   Friend WithEvents AxSpreadsheet1 As AxOWC10.AxSpreadsheet
    Friend WithEvents TBYesNo As System.Windows.Forms.ToolBar
    Friend WithEvents TBNextEdit As System.Windows.Forms.ToolBar
    Friend WithEvents TBEditOnly As System.Windows.Forms.ToolBar
    Friend WithEvents ImageListGeneralTools As System.Windows.Forms.ImageList
    Friend WithEvents MenuItem78 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem79 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem80 As System.Windows.Forms.MenuItem
    Friend WithEvents TBMultiDisplayField As System.Windows.Forms.ToolBar
    Friend WithEvents BTHome As System.Windows.Forms.Button
    Friend WithEvents BTChgSyringe As System.Windows.Forms.Button
    Friend WithEvents BTVolCal As System.Windows.Forms.Button
    Friend WithEvents BTNeedleCal As System.Windows.Forms.Button
    Friend WithEvents BTPurge As System.Windows.Forms.Button
    Friend WithEvents BTClean As System.Windows.Forms.Button
    Friend WithEvents PBYellow As System.Windows.Forms.PictureBox
    Friend WithEvents PBGreen As System.Windows.Forms.PictureBox
    Friend WithEvents ImageListMultiField As System.Windows.Forms.ImageList
    Friend WithEvents PBRed As System.Windows.Forms.PictureBox
    Friend WithEvents CBDoorLock As System.Windows.Forms.CheckBox
    Friend WithEvents PanelGraphDisp As System.Windows.Forms.Panel
    Friend WithEvents TBBGraphDisp As System.Windows.Forms.ToolBarButton
    Friend WithEvents TBBStepper As System.Windows.Forms.ToolBarButton
    Friend WithEvents TBBDispenser As System.Windows.Forms.ToolBarButton
    Friend WithEvents TBBThermal As System.Windows.Forms.ToolBarButton
    Friend WithEvents TBBConveyor As System.Windows.Forms.ToolBarButton
    Friend WithEvents LabelStep As System.Windows.Forms.Label
    Friend WithEvents DomainUpDown3 As System.Windows.Forms.DomainUpDown
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents PanelStepper As System.Windows.Forms.Panel
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents Label17 As System.Windows.Forms.Label
    Friend WithEvents Label18 As System.Windows.Forms.Label
    Friend WithEvents Label19 As System.Windows.Forms.Label
    Friend WithEvents Label20 As System.Windows.Forms.Label
    Friend WithEvents Label21 As System.Windows.Forms.Label
    Friend WithEvents PanelThermal As System.Windows.Forms.Panel
    Friend WithEvents PictureBox2 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox3 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox4 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox5 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox6 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox7 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox8 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox9 As System.Windows.Forms.PictureBox
    Friend WithEvents PanelDispenser As System.Windows.Forms.Panel
    Friend WithEvents PanelConveyor As System.Windows.Forms.Panel
    Friend WithEvents Label30 As System.Windows.Forms.Label
    Friend WithEvents PanelVision As System.Windows.Forms.Panel
    Friend WithEvents PanelVisionCtrl As System.Windows.Forms.Panel
    Friend WithEvents Reset As System.Windows.Forms.Button
    Friend WithEvents Download As System.Windows.Forms.Button
    Friend WithEvents BTStepZdown As System.Windows.Forms.Button
    Friend WithEvents BTStepZup As System.Windows.Forms.Button
    Friend WithEvents BTStepXminus As System.Windows.Forms.Button
    Friend WithEvents BTStepXplus As System.Windows.Forms.Button
    Friend WithEvents BTStepYminus As System.Windows.Forms.Button
    Friend WithEvents BTStepYplus As System.Windows.Forms.Button
    Friend WithEvents ComboBoxFineStep As System.Windows.Forms.ComboBox
    Friend WithEvents NeedleContextMenu As System.Windows.Forms.ContextMenu
    Friend WithEvents MenuItem81 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem82 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem83 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem84 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem85 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem86 As System.Windows.Forms.MenuItem
    Friend WithEvents OpenPatternFileDialog As System.Windows.Forms.OpenFileDialog
    Friend WithEvents SavePatternFileDialog As System.Windows.Forms.SaveFileDialog
    Public WithEvents RBProgramEditor As System.Windows.Forms.RadioButton
    Public WithEvents RBBasicSetup As System.Windows.Forms.RadioButton
    Friend WithEvents TextBoxRobotZ As System.Windows.Forms.TextBox
    Friend WithEvents TextBoxRobotY As System.Windows.Forms.TextBox
    Friend WithEvents TextBoxRobotX As System.Windows.Forms.TextBox
    Friend WithEvents ImageListLineArc As System.Windows.Forms.ImageList
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents MenuItem87 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem88 As System.Windows.Forms.MenuItem
    Friend WithEvents NumericUpDownSysSpeed As System.Windows.Forms.NumericUpDown
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents ToolBarButton1 As System.Windows.Forms.ToolBarButton
    Friend WithEvents ToolBarButton2 As System.Windows.Forms.ToolBarButton
    Friend WithEvents ToolBarButton5 As System.Windows.Forms.ToolBarButton
    Friend WithEvents ImageListOper As System.Windows.Forms.ImageList
    Friend WithEvents ComboBox1 As System.Windows.Forms.ComboBox
    Friend WithEvents TBOperation As System.Windows.Forms.ToolBar
    Friend WithEvents ToolBarButton3 As System.Windows.Forms.ToolBarButton
    Friend WithEvents ToolBarButton4 As System.Windows.Forms.ToolBarButton
    Friend WithEvents ToolBarLineArc As System.Windows.Forms.ToolBar
    Friend WithEvents LB_MM As System.Windows.Forms.Label
    Friend WithEvents LB_INCH As System.Windows.Forms.Label
    Friend WithEvents MainMenuProgramming As System.Windows.Forms.MainMenu
    Friend WithEvents Button4 As System.Windows.Forms.Button
    Friend WithEvents AxSpreadsheetProgramming As AxOWC10.AxSpreadsheet
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
    Friend WithEvents TBBOffest As System.Windows.Forms.ToolBarButton
    Friend WithEvents TBBMeasure As System.Windows.Forms.ToolBarButton
    Friend WithEvents MenuFileSave As System.Windows.Forms.MenuItem
    Friend WithEvents MenuFileSaveAs As System.Windows.Forms.MenuItem
    Friend WithEvents MenuFileExit As System.Windows.Forms.MenuItem
    Friend WithEvents MenuCommandsReference As System.Windows.Forms.MenuItem
    Friend WithEvents MenuCommandsFiducial As System.Windows.Forms.MenuItem
    Friend WithEvents MenuCommandsHeight As System.Windows.Forms.MenuItem
    Friend WithEvents MenuCommandsReject As System.Windows.Forms.MenuItem
    Friend WithEvents MenuCommandsDot As System.Windows.Forms.MenuItem
    Friend WithEvents MenuCommandsLine As System.Windows.Forms.MenuItem
    Friend WithEvents MenuCommandsEdit As System.Windows.Forms.MenuItem
    Friend WithEvents MenuCommandsRectangle As System.Windows.Forms.MenuItem
    Friend WithEvents MenuCommandsFillRectangle As System.Windows.Forms.MenuItem
    Friend WithEvents MenuCommandsLink As System.Windows.Forms.MenuItem
    Friend WithEvents MenuCommandsChipEdge As System.Windows.Forms.MenuItem
    Friend WithEvents MenuCommandsPurge As System.Windows.Forms.MenuItem
    Friend WithEvents MenuCommandsClean As System.Windows.Forms.MenuItem
    Friend WithEvents MenuCommandsQC As System.Windows.Forms.MenuItem
    Friend WithEvents MenuCommandsSetIO As System.Windows.Forms.MenuItem
    Friend WithEvents MenuCommandsGetIO As System.Windows.Forms.MenuItem
    Friend WithEvents MenuCommandsMove As System.Windows.Forms.MenuItem
    Friend WithEvents MenuCommandsWait As System.Windows.Forms.MenuItem
    Friend WithEvents MenuCommandsSubPattern As System.Windows.Forms.MenuItem
    Friend WithEvents MenuCommandsArc As System.Windows.Forms.MenuItem
    Friend WithEvents MenuCommandsCircle As System.Windows.Forms.MenuItem
    Friend WithEvents MenuCommandsFillCircle As System.Windows.Forms.MenuItem
    Friend WithEvents MenuCommandsArray As System.Windows.Forms.MenuItem
    Friend WithEvents ToolBarButton6 As System.Windows.Forms.ToolBarButton
    Friend WithEvents MenuItem2 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuCommandsOffset As System.Windows.Forms.MenuItem
    Friend WithEvents MenuCommandsMeasurement As System.Windows.Forms.MenuItem
    Friend WithEvents NumericUpDown_Brightness As System.Windows.Forms.NumericUpDown
    Friend WithEvents BT_HeaterCancel As System.Windows.Forms.Button
    Friend WithEvents BT_HeaterUpdate As System.Windows.Forms.Button
    Friend WithEvents TB_NeedleHeater As System.Windows.Forms.TextBox
    Friend WithEvents TB_SyrHeater As System.Windows.Forms.TextBox
    Friend WithEvents LB_NeedleHeater As System.Windows.Forms.Label
    Friend WithEvents LB_SyrHeater As System.Windows.Forms.Label
    Friend WithEvents LB_PostHeater As System.Windows.Forms.Label
    Friend WithEvents LB_DispHeater As System.Windows.Forms.Label
    Friend WithEvents LB_PreHeater As System.Windows.Forms.Label
    Friend WithEvents TB_PostHeater As System.Windows.Forms.TextBox
    Friend WithEvents TB_DispHeater As System.Windows.Forms.TextBox
    Friend WithEvents TB_PreHeater As System.Windows.Forms.TextBox
    Friend WithEvents LB_Pro_LPressure As System.Windows.Forms.Label
    Friend WithEvents TB_Pro_RREV As System.Windows.Forms.TextBox
    Friend WithEvents LB_Pro_RRPM As System.Windows.Forms.Label
    Friend WithEvents LB_Pro_RPressure As System.Windows.Forms.Label
    Friend WithEvents TB_Pro_RPressure As System.Windows.Forms.TextBox
    Friend WithEvents TB_Pro_RRPM As System.Windows.Forms.TextBox
    Friend WithEvents TB_Pro_LREV As System.Windows.Forms.TextBox
    Friend WithEvents LB_Pro_LRPM As System.Windows.Forms.Label
    Friend WithEvents TB_Pro_LPressure As System.Windows.Forms.TextBox
    Friend WithEvents TB_Pro_LRPM As System.Windows.Forms.TextBox
    Friend WithEvents LB_Pro_LREV As System.Windows.Forms.Label
    Friend WithEvents TB_Pro_LProg As System.Windows.Forms.TextBox
    Friend WithEvents TB_Pro_LValve As System.Windows.Forms.TextBox
    Friend WithEvents LB_Pro_LProg As System.Windows.Forms.Label
    Friend WithEvents LB_Pro_LValve As System.Windows.Forms.Label
    Friend WithEvents TB_Pro_RProg As System.Windows.Forms.TextBox
    Friend WithEvents LB_Pro_RProg As System.Windows.Forms.Label
    Friend WithEvents TB_Pro_RValve As System.Windows.Forms.TextBox
    Friend WithEvents LB_Pro_RValve As System.Windows.Forms.Label
    Friend WithEvents GB_RTP As System.Windows.Forms.GroupBox
    Friend WithEvents GB_RJetter As System.Windows.Forms.GroupBox
    Friend WithEvents GB_LTP As System.Windows.Forms.GroupBox
    Friend WithEvents GB_LJetter As System.Windows.Forms.GroupBox
    Friend WithEvents BT_CV_Release As System.Windows.Forms.Button
    Friend WithEvents BT_CV_Retrieve As System.Windows.Forms.Button
    Friend WithEvents TimerForUpdate As System.Windows.Forms.Timer
    Friend WithEvents BT_Pro_REdit As System.Windows.Forms.CheckBox
    Friend WithEvents BT_Pro_LEdit As System.Windows.Forms.CheckBox
    Friend WithEvents BT_Pro_RDownload As System.Windows.Forms.Button
    Friend WithEvents BT_Pro_RUpdate As System.Windows.Forms.Button
    Friend WithEvents BT_Pro_LDownload As System.Windows.Forms.Button
    Friend WithEvents BT_Pro_LUpdate As System.Windows.Forms.Button
    Friend WithEvents TB_Pro_RSuck As System.Windows.Forms.TextBox
    Friend WithEvents TB_Pro_LSuck As System.Windows.Forms.TextBox
    Friend WithEvents BT_RJetter_Browse As System.Windows.Forms.Button
    Friend WithEvents BT_RJetter_Open As System.Windows.Forms.Button
    Friend WithEvents BT_LJetter_Browse As System.Windows.Forms.Button
    Friend WithEvents BT_LJetter_Open As System.Windows.Forms.Button
    Friend WithEvents TB_RJetter_Path As System.Windows.Forms.TextBox
    Friend WithEvents TB_LJetter_Path As System.Windows.Forms.TextBox
    Friend WithEvents BT_RJetter_Edit As System.Windows.Forms.CheckBox
    Friend WithEvents BT_LJetter_Edit As System.Windows.Forms.CheckBox
    Friend WithEvents LB_IDSPressure_RJetter As System.Windows.Forms.Label
    Friend WithEvents LB_IDSPressure_LJetter As System.Windows.Forms.Label
    Friend WithEvents LB_IDSPressure_RTP As System.Windows.Forms.Label
    Friend WithEvents LB_IDSPressure_LTP As System.Windows.Forms.Label
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
    Friend WithEvents MenuSystemHome As System.Windows.Forms.MenuItem
    Friend WithEvents MenuVolmCalib As System.Windows.Forms.MenuItem
    Friend WithEvents MenuNeedleCalib As System.Windows.Forms.MenuItem
    Friend WithEvents MenuNeedleChng As System.Windows.Forms.MenuItem
    Friend WithEvents MenuPurge As System.Windows.Forms.MenuItem
    Friend WithEvents MenuClean As System.Windows.Forms.MenuItem
    Friend WithEvents MenuDoorUnLock As System.Windows.Forms.MenuItem
    Friend WithEvents MenuDoorLock As System.Windows.Forms.MenuItem
    Friend WithEvents PictureBoxEstopRed As System.Windows.Forms.PictureBox
    Friend WithEvents Button5 As System.Windows.Forms.Button
    Friend WithEvents OptimizePath As System.Windows.Forms.MenuItem
    Friend WithEvents SaveFileDialog1 As System.Windows.Forms.SaveFileDialog
    Friend WithEvents LB_Pro_RSuck As System.Windows.Forms.Label
    Friend WithEvents LB_Pro_LSuck As System.Windows.Forms.Label
    Friend WithEvents LB_Pro_RREV As System.Windows.Forms.Label
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
        Me.MenuItem23 = New System.Windows.Forms.MenuItem
        Me.MenuItem24 = New System.Windows.Forms.MenuItem
        Me.MenuItem66 = New System.Windows.Forms.MenuItem
        Me.MenuItem67 = New System.Windows.Forms.MenuItem
        Me.MenuItem68 = New System.Windows.Forms.MenuItem
        Me.MenuItem69 = New System.Windows.Forms.MenuItem
        Me.MenuItem87 = New System.Windows.Forms.MenuItem
        Me.MenuItem88 = New System.Windows.Forms.MenuItem
        Me.MenuItem25 = New System.Windows.Forms.MenuItem
        Me.MenuSystemHome = New System.Windows.Forms.MenuItem
        Me.MenuVolmCalib = New System.Windows.Forms.MenuItem
        Me.MenuNeedleCalib = New System.Windows.Forms.MenuItem
        Me.MenuNeedleChng = New System.Windows.Forms.MenuItem
        Me.MenuPurge = New System.Windows.Forms.MenuItem
        Me.MenuClean = New System.Windows.Forms.MenuItem
        Me.MenuItem34 = New System.Windows.Forms.MenuItem
        Me.MenuItem79 = New System.Windows.Forms.MenuItem
        Me.MenuItem80 = New System.Windows.Forms.MenuItem
        Me.MenuItem78 = New System.Windows.Forms.MenuItem
        Me.MenuItem33 = New System.Windows.Forms.MenuItem
        Me.MenuItem32 = New System.Windows.Forms.MenuItem
        Me.MenuItem35 = New System.Windows.Forms.MenuItem
        Me.MenuItem29 = New System.Windows.Forms.MenuItem
        Me.MenuItem30 = New System.Windows.Forms.MenuItem
        Me.MenuItem36 = New System.Windows.Forms.MenuItem
        Me.MenuItem62 = New System.Windows.Forms.MenuItem
        Me.MenuItem63 = New System.Windows.Forms.MenuItem
        Me.MenuItem71 = New System.Windows.Forms.MenuItem
        Me.MenuItem72 = New System.Windows.Forms.MenuItem
        Me.MenuItem73 = New System.Windows.Forms.MenuItem
        Me.MenuDoorLock = New System.Windows.Forms.MenuItem
        Me.MenuDoorUnLock = New System.Windows.Forms.MenuItem
        Me.MenuItem76 = New System.Windows.Forms.MenuItem
        Me.OptimizePath = New System.Windows.Forms.MenuItem
        Me.MenuItem37 = New System.Windows.Forms.MenuItem
        Me.MenuItem48 = New System.Windows.Forms.MenuItem
        Me.MenuCommandsReference = New System.Windows.Forms.MenuItem
        Me.MenuCommandsFiducial = New System.Windows.Forms.MenuItem
        Me.MenuCommandsHeight = New System.Windows.Forms.MenuItem
        Me.MenuCommandsReject = New System.Windows.Forms.MenuItem
        Me.MenuItem45 = New System.Windows.Forms.MenuItem
        Me.MenuCommandsDot = New System.Windows.Forms.MenuItem
        Me.MenuCommandsLine = New System.Windows.Forms.MenuItem
        Me.MenuCommandsArc = New System.Windows.Forms.MenuItem
        Me.MenuCommandsRectangle = New System.Windows.Forms.MenuItem
        Me.MenuCommandsFillRectangle = New System.Windows.Forms.MenuItem
        Me.MenuCommandsCircle = New System.Windows.Forms.MenuItem
        Me.MenuCommandsFillCircle = New System.Windows.Forms.MenuItem
        Me.MenuCommandsLink = New System.Windows.Forms.MenuItem
        Me.MenuCommandsChipEdge = New System.Windows.Forms.MenuItem
        Me.MenuCommandsMove = New System.Windows.Forms.MenuItem
        Me.MenuCommandsWait = New System.Windows.Forms.MenuItem
        Me.MenuCommandsPurge = New System.Windows.Forms.MenuItem
        Me.MenuCommandsClean = New System.Windows.Forms.MenuItem
        Me.MenuCommandsQC = New System.Windows.Forms.MenuItem
        Me.MenuCommandsSubPattern = New System.Windows.Forms.MenuItem
        Me.MenuCommandsArray = New System.Windows.Forms.MenuItem
        Me.MenuCommandsSetIO = New System.Windows.Forms.MenuItem
        Me.MenuCommandsGetIO = New System.Windows.Forms.MenuItem
        Me.MenuItem2 = New System.Windows.Forms.MenuItem
        Me.MenuCommandsOffset = New System.Windows.Forms.MenuItem
        Me.MenuCommandsMeasurement = New System.Windows.Forms.MenuItem
        Me.MenuCommandsEdit = New System.Windows.Forms.MenuItem
        Me.MenuItem64 = New System.Windows.Forms.MenuItem
        Me.MenuItem65 = New System.Windows.Forms.MenuItem
        Me.PanelVision = New System.Windows.Forms.Panel
        Me.ImageListGeneralTools = New System.Windows.Forms.ImageList(Me.components)
        Me.BTChgSyringe = New System.Windows.Forms.Button
        Me.BTPurge = New System.Windows.Forms.Button
        Me.PanelVisionCtrl = New System.Windows.Forms.Panel
        Me.NumericUpDown_Brightness = New System.Windows.Forms.NumericUpDown
        Me.CheckBoxLockZ = New System.Windows.Forms.CheckBox
        Me.TextBoxRobotZ = New System.Windows.Forms.TextBox
        Me.CheckBoxLockY = New System.Windows.Forms.CheckBox
        Me.TextBoxRobotY = New System.Windows.Forms.TextBox
        Me.CheckBoxLockX = New System.Windows.Forms.CheckBox
        Me.TextBoxRobotX = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.TextBox1 = New System.Windows.Forms.TextBox
        Me.DomainUpDown1 = New System.Windows.Forms.DomainUpDown
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.ImageListFiles = New System.Windows.Forms.ImageList(Me.components)
        Me.imageListElement = New System.Windows.Forms.ImageList(Me.components)
        Me.ImageListReference = New System.Windows.Forms.ImageList(Me.components)
        Me.TBrefer = New System.Windows.Forms.ToolBar
        Me.TBFiducialPt = New System.Windows.Forms.ToolBarButton
        Me.TBHeightPt = New System.Windows.Forms.ToolBarButton
        Me.TBReferencePt = New System.Windows.Forms.ToolBarButton
        Me.TBRejectPt = New System.Windows.Forms.ToolBarButton
        Me.TBElements = New System.Windows.Forms.ToolBar
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
        Me.ToolBarButton6 = New System.Windows.Forms.ToolBarButton
        Me.TBBOffest = New System.Windows.Forms.ToolBarButton
        Me.TBBMeasure = New System.Windows.Forms.ToolBarButton
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.LabelMessege = New System.Windows.Forms.Label
        Me.ImageListYesNo = New System.Windows.Forms.ImageList(Me.components)
        Me.RBLNeedleMode = New System.Windows.Forms.RadioButton
        Me.RBRNeedleMode = New System.Windows.Forms.RadioButton
        Me.Label5 = New System.Windows.Forms.Label
        Me.RBVisionMode = New System.Windows.Forms.RadioButton
        Me.RBProgramEditor = New System.Windows.Forms.RadioButton
        Me.RBBasicSetup = New System.Windows.Forms.RadioButton
        Me.CBExpandSpreadsheet = New System.Windows.Forms.CheckBox
        Me.Panel4 = New System.Windows.Forms.Panel
        Me.TBYesNo = New System.Windows.Forms.ToolBar
        Me.TBBOk = New System.Windows.Forms.ToolBarButton
        Me.TBBCancel = New System.Windows.Forms.ToolBarButton
        Me.TBEditOnly = New System.Windows.Forms.ToolBar
        Me.TBBEditOnly = New System.Windows.Forms.ToolBarButton
        Me.TBNextEdit = New System.Windows.Forms.ToolBar
        Me.TBBSwitch = New System.Windows.Forms.ToolBarButton
        Me.TBBEdit = New System.Windows.Forms.ToolBarButton
        Me.PanelMotion = New System.Windows.Forms.Panel
        Me.ComboBox1 = New System.Windows.Forms.ComboBox
        Me.Label27 = New System.Windows.Forms.Label
        Me.PBYellow = New System.Windows.Forms.PictureBox
        Me.PBRed = New System.Windows.Forms.PictureBox
        Me.PBGreen = New System.Windows.Forms.PictureBox
        Me.TBOperation = New System.Windows.Forms.ToolBar
        Me.ToolBarButton1 = New System.Windows.Forms.ToolBarButton
        Me.ToolBarButton2 = New System.Windows.Forms.ToolBarButton
        Me.ToolBarButton5 = New System.Windows.Forms.ToolBarButton
        Me.ImageListOper = New System.Windows.Forms.ImageList(Me.components)
        Me.BTClean = New System.Windows.Forms.Button
        Me.BTNeedleCal = New System.Windows.Forms.Button
        Me.BTVolCal = New System.Windows.Forms.Button
        Me.BTHome = New System.Windows.Forms.Button
        Me.ImageListMultiField = New System.Windows.Forms.ImageList(Me.components)
        Me.TBMultiDisplayField = New System.Windows.Forms.ToolBar
        Me.TBBGraphDisp = New System.Windows.Forms.ToolBarButton
        Me.TBBStepper = New System.Windows.Forms.ToolBarButton
        Me.TBBDispenser = New System.Windows.Forms.ToolBarButton
        Me.TBBThermal = New System.Windows.Forms.ToolBarButton
        Me.TBBConveyor = New System.Windows.Forms.ToolBarButton
        Me.CBDoorLock = New System.Windows.Forms.CheckBox
        Me.PanelGraphDisp = New System.Windows.Forms.Panel
        Me.PanelStepper = New System.Windows.Forms.Panel
        Me.ComboBoxFineStep = New System.Windows.Forms.ComboBox
        Me.BTStepZdown = New System.Windows.Forms.Button
        Me.BTStepZup = New System.Windows.Forms.Button
        Me.BTStepXminus = New System.Windows.Forms.Button
        Me.BTStepXplus = New System.Windows.Forms.Button
        Me.BTStepYminus = New System.Windows.Forms.Button
        Me.BTStepYplus = New System.Windows.Forms.Button
        Me.Label6 = New System.Windows.Forms.Label
        Me.DomainUpDown3 = New System.Windows.Forms.DomainUpDown
        Me.LabelStep = New System.Windows.Forms.Label
        Me.PanelDispenser = New System.Windows.Forms.Panel
        Me.GB_RTP = New System.Windows.Forms.GroupBox
        Me.TB_Pro_RRPM = New System.Windows.Forms.TextBox
        Me.BT_Pro_RDownload = New System.Windows.Forms.Button
        Me.BT_Pro_RUpdate = New System.Windows.Forms.Button
        Me.TB_Pro_RSuck = New System.Windows.Forms.TextBox
        Me.TB_Pro_RREV = New System.Windows.Forms.TextBox
        Me.LB_Pro_RSuck = New System.Windows.Forms.Label
        Me.LB_Pro_RRPM = New System.Windows.Forms.Label
        Me.LB_Pro_RPressure = New System.Windows.Forms.Label
        Me.TB_Pro_RPressure = New System.Windows.Forms.TextBox
        Me.LB_Pro_RREV = New System.Windows.Forms.Label
        Me.BT_Pro_REdit = New System.Windows.Forms.CheckBox
        Me.LB_IDSPressure_RTP = New System.Windows.Forms.Label
        Me.TB_Pro_RProg = New System.Windows.Forms.TextBox
        Me.LB_Pro_RProg = New System.Windows.Forms.Label
        Me.TB_Pro_RValve = New System.Windows.Forms.TextBox
        Me.LB_Pro_RValve = New System.Windows.Forms.Label
        Me.GB_RJetter = New System.Windows.Forms.GroupBox
        Me.LB_IDSPressure_RJetter = New System.Windows.Forms.Label
        Me.BT_RJetter_Edit = New System.Windows.Forms.CheckBox
        Me.TB_RJetter_Path = New System.Windows.Forms.TextBox
        Me.BT_RJetter_Browse = New System.Windows.Forms.Button
        Me.BT_RJetter_Open = New System.Windows.Forms.Button
        Me.Label30 = New System.Windows.Forms.Label
        Me.GB_LJetter = New System.Windows.Forms.GroupBox
        Me.LB_IDSPressure_LJetter = New System.Windows.Forms.Label
        Me.BT_LJetter_Edit = New System.Windows.Forms.CheckBox
        Me.TB_LJetter_Path = New System.Windows.Forms.TextBox
        Me.BT_LJetter_Browse = New System.Windows.Forms.Button
        Me.BT_LJetter_Open = New System.Windows.Forms.Button
        Me.Label15 = New System.Windows.Forms.Label
        Me.GB_LTP = New System.Windows.Forms.GroupBox
        Me.TB_Pro_LRPM = New System.Windows.Forms.TextBox
        Me.LB_IDSPressure_LTP = New System.Windows.Forms.Label
        Me.BT_Pro_LEdit = New System.Windows.Forms.CheckBox
        Me.BT_Pro_LDownload = New System.Windows.Forms.Button
        Me.BT_Pro_LUpdate = New System.Windows.Forms.Button
        Me.TB_Pro_LSuck = New System.Windows.Forms.TextBox
        Me.TB_Pro_LREV = New System.Windows.Forms.TextBox
        Me.LB_Pro_LSuck = New System.Windows.Forms.Label
        Me.LB_Pro_LRPM = New System.Windows.Forms.Label
        Me.LB_Pro_LPressure = New System.Windows.Forms.Label
        Me.TB_Pro_LPressure = New System.Windows.Forms.TextBox
        Me.LB_Pro_LREV = New System.Windows.Forms.Label
        Me.TB_Pro_LProg = New System.Windows.Forms.TextBox
        Me.TB_Pro_LValve = New System.Windows.Forms.TextBox
        Me.LB_Pro_LProg = New System.Windows.Forms.Label
        Me.LB_Pro_LValve = New System.Windows.Forms.Label
        Me.PanelThermal = New System.Windows.Forms.Panel
        Me.PictureBox6 = New System.Windows.Forms.PictureBox
        Me.PictureBox5 = New System.Windows.Forms.PictureBox
        Me.PictureBox4 = New System.Windows.Forms.PictureBox
        Me.PictureBox3 = New System.Windows.Forms.PictureBox
        Me.PictureBox2 = New System.Windows.Forms.PictureBox
        Me.LB_PostHeater = New System.Windows.Forms.Label
        Me.LB_DispHeater = New System.Windows.Forms.Label
        Me.LB_PreHeater = New System.Windows.Forms.Label
        Me.LB_NeedleHeater = New System.Windows.Forms.Label
        Me.LB_SyrHeater = New System.Windows.Forms.Label
        Me.BT_HeaterCancel = New System.Windows.Forms.Button
        Me.BT_HeaterUpdate = New System.Windows.Forms.Button
        Me.TB_PostHeater = New System.Windows.Forms.TextBox
        Me.TB_DispHeater = New System.Windows.Forms.TextBox
        Me.TB_PreHeater = New System.Windows.Forms.TextBox
        Me.TB_NeedleHeater = New System.Windows.Forms.TextBox
        Me.TB_SyrHeater = New System.Windows.Forms.TextBox
        Me.Label21 = New System.Windows.Forms.Label
        Me.Label20 = New System.Windows.Forms.Label
        Me.Label19 = New System.Windows.Forms.Label
        Me.Label18 = New System.Windows.Forms.Label
        Me.Label17 = New System.Windows.Forms.Label
        Me.PanelConveyor = New System.Windows.Forms.Panel
        Me.BT_CV_Release = New System.Windows.Forms.Button
        Me.BT_CV_Retrieve = New System.Windows.Forms.Button
        Me.PictureBox9 = New System.Windows.Forms.PictureBox
        Me.PictureBox8 = New System.Windows.Forms.PictureBox
        Me.PictureBox7 = New System.Windows.Forms.PictureBox
        Me.Reset = New System.Windows.Forms.Button
        Me.Download = New System.Windows.Forms.Button
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
        Me.Button1 = New System.Windows.Forms.Button
        Me.NumericUpDownSysSpeed = New System.Windows.Forms.NumericUpDown
        Me.Button2 = New System.Windows.Forms.Button
        Me.Button3 = New System.Windows.Forms.Button
        Me.ToolBarButton3 = New System.Windows.Forms.ToolBarButton
        Me.ToolBarButton4 = New System.Windows.Forms.ToolBarButton
        Me.ToolBarLineArc = New System.Windows.Forms.ToolBar
        Me.LB_MM = New System.Windows.Forms.Label
        Me.LB_INCH = New System.Windows.Forms.Label
        Me.Button4 = New System.Windows.Forms.Button
        Me.AxSpreadsheetProgramming = New AxOWC10.AxSpreadsheet
        Me.TimerForUpdate = New System.Windows.Forms.Timer(Me.components)
        Me.PictureBoxEstopRed = New System.Windows.Forms.PictureBox
        Me.Button5 = New System.Windows.Forms.Button
        Me.SaveFileDialog1 = New System.Windows.Forms.SaveFileDialog
        Me.PanelVisionCtrl.SuspendLayout()
        CType(Me.NumericUpDown_Brightness, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel4.SuspendLayout()
        Me.PanelMotion.SuspendLayout()
        Me.PanelStepper.SuspendLayout()
        Me.PanelDispenser.SuspendLayout()
        Me.GB_RTP.SuspendLayout()
        Me.GB_RJetter.SuspendLayout()
        Me.GB_LJetter.SuspendLayout()
        Me.GB_LTP.SuspendLayout()
        Me.PanelThermal.SuspendLayout()
        Me.PanelConveyor.SuspendLayout()
        CType(Me.NumericUpDownSysSpeed, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.AxSpreadsheetProgramming, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'MainMenuProgramming
        '
        Me.MainMenuProgramming.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem1, Me.MenuItem8, Me.MenuItem11, Me.MenuItem25, Me.MenuItem37, Me.MenuItem64})
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
        Me.MenuItem11.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem23, Me.MenuItem24, Me.MenuItem66, Me.MenuItem67, Me.MenuItem68, Me.MenuItem69, Me.MenuItem87, Me.MenuItem88})
        Me.MenuItem11.Text = "View"
        '
        'MenuItem23
        '
        Me.MenuItem23.Index = 0
        Me.MenuItem23.Text = "Expanse Spreadsheet"
        '
        'MenuItem24
        '
        Me.MenuItem24.Index = 1
        Me.MenuItem24.Text = "Normal Spreadsheet"
        '
        'MenuItem66
        '
        Me.MenuItem66.Index = 2
        Me.MenuItem66.Text = "-"
        '
        'MenuItem67
        '
        Me.MenuItem67.Index = 3
        Me.MenuItem67.Text = "Information"
        '
        'MenuItem68
        '
        Me.MenuItem68.Index = 4
        Me.MenuItem68.Text = "Basic Setup"
        '
        'MenuItem69
        '
        Me.MenuItem69.Index = 5
        Me.MenuItem69.Text = "Program Editor"
        '
        'MenuItem87
        '
        Me.MenuItem87.Index = 6
        Me.MenuItem87.Text = "-"
        '
        'MenuItem88
        '
        Me.MenuItem88.Index = 7
        Me.MenuItem88.Text = "Open Event Viewer"
        '
        'MenuItem25
        '
        Me.MenuItem25.Index = 3
        Me.MenuItem25.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuSystemHome, Me.MenuVolmCalib, Me.MenuNeedleCalib, Me.MenuNeedleChng, Me.MenuPurge, Me.MenuClean, Me.MenuItem34, Me.MenuItem79, Me.MenuItem80, Me.MenuItem78, Me.MenuItem33, Me.MenuItem32, Me.MenuItem35, Me.MenuItem29, Me.MenuItem30, Me.MenuItem36, Me.MenuItem62, Me.MenuItem63, Me.MenuItem71, Me.MenuItem72, Me.MenuItem73, Me.MenuDoorLock, Me.MenuDoorUnLock, Me.MenuItem76, Me.OptimizePath})
        Me.MenuItem25.Text = "Tools"
        '
        'MenuSystemHome
        '
        Me.MenuSystemHome.Index = 0
        Me.MenuSystemHome.Text = "System Home"
        '
        'MenuVolmCalib
        '
        Me.MenuVolmCalib.Index = 1
        Me.MenuVolmCalib.Text = "Volume Calibration"
        '
        'MenuNeedleCalib
        '
        Me.MenuNeedleCalib.Index = 2
        Me.MenuNeedleCalib.Text = "Needle Calibration"
        '
        'MenuNeedleChng
        '
        Me.MenuNeedleChng.Index = 3
        Me.MenuNeedleChng.Text = "Change Needle"
        '
        'MenuPurge
        '
        Me.MenuPurge.Index = 4
        Me.MenuPurge.Text = "Purge"
        '
        'MenuClean
        '
        Me.MenuClean.Index = 5
        Me.MenuClean.Text = "Clean"
        '
        'MenuItem34
        '
        Me.MenuItem34.Index = 6
        Me.MenuItem34.Text = "-"
        '
        'MenuItem79
        '
        Me.MenuItem79.Index = 7
        Me.MenuItem79.Text = "Volume Calibration Setup"
        '
        'MenuItem80
        '
        Me.MenuItem80.Index = 8
        Me.MenuItem80.Text = "Needle Calibration Setup"
        '
        'MenuItem78
        '
        Me.MenuItem78.Index = 9
        Me.MenuItem78.Text = "-"
        '
        'MenuItem33
        '
        Me.MenuItem33.Index = 10
        Me.MenuItem33.Text = "Offset"
        '
        'MenuItem32
        '
        Me.MenuItem32.Index = 11
        Me.MenuItem32.Text = "Sub Pattern"
        '
        'MenuItem35
        '
        Me.MenuItem35.Index = 12
        Me.MenuItem35.Text = "-"
        '
        'MenuItem29
        '
        Me.MenuItem29.Index = 13
        Me.MenuItem29.Text = "Set I/O"
        '
        'MenuItem30
        '
        Me.MenuItem30.Index = 14
        Me.MenuItem30.Text = "Reset I/O"
        '
        'MenuItem36
        '
        Me.MenuItem36.Index = 15
        Me.MenuItem36.Text = "-"
        '
        'MenuItem62
        '
        Me.MenuItem62.Index = 16
        Me.MenuItem62.Text = "Measurement"
        '
        'MenuItem63
        '
        Me.MenuItem63.Index = 17
        Me.MenuItem63.Text = "-"
        '
        'MenuItem71
        '
        Me.MenuItem71.Index = 18
        Me.MenuItem71.Text = "Retrieve Board"
        '
        'MenuItem72
        '
        Me.MenuItem72.Index = 19
        Me.MenuItem72.Text = "Release Board"
        '
        'MenuItem73
        '
        Me.MenuItem73.Index = 20
        Me.MenuItem73.Text = "-"
        '
        'MenuDoorLock
        '
        Me.MenuDoorLock.Index = 21
        Me.MenuDoorLock.Text = "Lock Door"
        '
        'MenuDoorUnLock
        '
        Me.MenuDoorUnLock.Index = 22
        Me.MenuDoorUnLock.Text = "Unlock Door"
        '
        'MenuItem76
        '
        Me.MenuItem76.Index = 23
        Me.MenuItem76.Text = "-"
        '
        'OptimizePath
        '
        Me.OptimizePath.Index = 24
        Me.OptimizePath.Text = "Optimize Dispensing Path"
        '
        'MenuItem37
        '
        Me.MenuItem37.Index = 4
        Me.MenuItem37.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem48, Me.MenuItem45, Me.MenuCommandsEdit})
        Me.MenuItem37.Text = "Commands"
        '
        'MenuItem48
        '
        Me.MenuItem48.Index = 0
        Me.MenuItem48.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuCommandsReference, Me.MenuCommandsFiducial, Me.MenuCommandsHeight, Me.MenuCommandsReject})
        Me.MenuItem48.Text = "Reference"
        '
        'MenuCommandsReference
        '
        Me.MenuCommandsReference.Index = 0
        Me.MenuCommandsReference.Text = "Reference Point"
        '
        'MenuCommandsFiducial
        '
        Me.MenuCommandsFiducial.Index = 1
        Me.MenuCommandsFiducial.Text = "Fiducial Point"
        '
        'MenuCommandsHeight
        '
        Me.MenuCommandsHeight.Index = 2
        Me.MenuCommandsHeight.Text = "Height Point"
        '
        'MenuCommandsReject
        '
        Me.MenuCommandsReject.Index = 3
        Me.MenuCommandsReject.Text = "Reject Mark"
        '
        'MenuItem45
        '
        Me.MenuItem45.Index = 1
        Me.MenuItem45.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuCommandsDot, Me.MenuCommandsLine, Me.MenuCommandsArc, Me.MenuCommandsRectangle, Me.MenuCommandsFillRectangle, Me.MenuCommandsCircle, Me.MenuCommandsFillCircle, Me.MenuCommandsLink, Me.MenuCommandsChipEdge, Me.MenuCommandsMove, Me.MenuCommandsWait, Me.MenuCommandsPurge, Me.MenuCommandsClean, Me.MenuCommandsQC, Me.MenuCommandsSubPattern, Me.MenuCommandsArray, Me.MenuCommandsSetIO, Me.MenuCommandsGetIO, Me.MenuItem2, Me.MenuCommandsOffset, Me.MenuCommandsMeasurement})
        Me.MenuItem45.Text = "Elements"
        '
        'MenuCommandsDot
        '
        Me.MenuCommandsDot.Index = 0
        Me.MenuCommandsDot.Text = "Dot"
        '
        'MenuCommandsLine
        '
        Me.MenuCommandsLine.Index = 1
        Me.MenuCommandsLine.Text = "Line"
        '
        'MenuCommandsArc
        '
        Me.MenuCommandsArc.Index = 2
        Me.MenuCommandsArc.Text = "Arc"
        '
        'MenuCommandsRectangle
        '
        Me.MenuCommandsRectangle.Index = 3
        Me.MenuCommandsRectangle.Text = "Rectangle"
        '
        'MenuCommandsFillRectangle
        '
        Me.MenuCommandsFillRectangle.Index = 4
        Me.MenuCommandsFillRectangle.Text = "FillRectangle"
        '
        'MenuCommandsCircle
        '
        Me.MenuCommandsCircle.Index = 5
        Me.MenuCommandsCircle.Text = "Circle"
        '
        'MenuCommandsFillCircle
        '
        Me.MenuCommandsFillCircle.Index = 6
        Me.MenuCommandsFillCircle.Text = "FillCircle"
        '
        'MenuCommandsLink
        '
        Me.MenuCommandsLink.Index = 7
        Me.MenuCommandsLink.Text = "Link"
        '
        'MenuCommandsChipEdge
        '
        Me.MenuCommandsChipEdge.Index = 8
        Me.MenuCommandsChipEdge.Text = "ChipEdge"
        '
        'MenuCommandsMove
        '
        Me.MenuCommandsMove.Index = 9
        Me.MenuCommandsMove.Text = "Move"
        '
        'MenuCommandsWait
        '
        Me.MenuCommandsWait.Index = 10
        Me.MenuCommandsWait.Text = "Wait"
        '
        'MenuCommandsPurge
        '
        Me.MenuCommandsPurge.Index = 11
        Me.MenuCommandsPurge.Text = "Purge"
        '
        'MenuCommandsClean
        '
        Me.MenuCommandsClean.Index = 12
        Me.MenuCommandsClean.Text = "Clean"
        '
        'MenuCommandsQC
        '
        Me.MenuCommandsQC.Index = 13
        Me.MenuCommandsQC.Text = "QC"
        '
        'MenuCommandsSubPattern
        '
        Me.MenuCommandsSubPattern.Index = 14
        Me.MenuCommandsSubPattern.Text = "SubPattern"
        '
        'MenuCommandsArray
        '
        Me.MenuCommandsArray.Index = 15
        Me.MenuCommandsArray.Text = "Array"
        '
        'MenuCommandsSetIO
        '
        Me.MenuCommandsSetIO.Index = 16
        Me.MenuCommandsSetIO.Text = "SetIO"
        '
        'MenuCommandsGetIO
        '
        Me.MenuCommandsGetIO.Index = 17
        Me.MenuCommandsGetIO.Text = "GetIO"
        '
        'MenuItem2
        '
        Me.MenuItem2.Index = 18
        Me.MenuItem2.Text = "-"
        '
        'MenuCommandsOffset
        '
        Me.MenuCommandsOffset.Index = 19
        Me.MenuCommandsOffset.Text = "Offset"
        '
        'MenuCommandsMeasurement
        '
        Me.MenuCommandsMeasurement.Index = 20
        Me.MenuCommandsMeasurement.Text = "Measurement"
        '
        'MenuCommandsEdit
        '
        Me.MenuCommandsEdit.Index = 2
        Me.MenuCommandsEdit.Text = "Edit"
        '
        'MenuItem64
        '
        Me.MenuItem64.Index = 5
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
        Me.PanelVision.Location = New System.Drawing.Point(84, 341)
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
        'BTChgSyringe
        '
        Me.BTChgSyringe.BackColor = System.Drawing.SystemColors.InactiveCaptionText
        Me.BTChgSyringe.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.BTChgSyringe.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.BTChgSyringe.ImageIndex = 1
        Me.BTChgSyringe.ImageList = Me.ImageListGeneralTools
        Me.BTChgSyringe.Location = New System.Drawing.Point(926, 362)
        Me.BTChgSyringe.Name = "BTChgSyringe"
        Me.BTChgSyringe.Size = New System.Drawing.Size(74, 60)
        Me.BTChgSyringe.TabIndex = 54
        Me.BTChgSyringe.TabStop = False
        Me.BTChgSyringe.Text = "Chg. Syr."
        Me.BTChgSyringe.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'BTPurge
        '
        Me.BTPurge.BackColor = System.Drawing.SystemColors.InactiveCaptionText
        Me.BTPurge.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.BTPurge.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.BTPurge.ImageIndex = 4
        Me.BTPurge.ImageList = Me.ImageListGeneralTools
        Me.BTPurge.Location = New System.Drawing.Point(852, 482)
        Me.BTPurge.Name = "BTPurge"
        Me.BTPurge.Size = New System.Drawing.Size(74, 60)
        Me.BTPurge.TabIndex = 57
        Me.BTPurge.Text = "Purge"
        Me.BTPurge.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'PanelVisionCtrl
        '
        Me.PanelVisionCtrl.Controls.Add(Me.NumericUpDown_Brightness)
        Me.PanelVisionCtrl.Controls.Add(Me.CheckBoxLockZ)
        Me.PanelVisionCtrl.Controls.Add(Me.TextBoxRobotZ)
        Me.PanelVisionCtrl.Controls.Add(Me.CheckBoxLockY)
        Me.PanelVisionCtrl.Controls.Add(Me.TextBoxRobotY)
        Me.PanelVisionCtrl.Controls.Add(Me.CheckBoxLockX)
        Me.PanelVisionCtrl.Controls.Add(Me.TextBoxRobotX)
        Me.PanelVisionCtrl.Controls.Add(Me.Label4)
        Me.PanelVisionCtrl.Controls.Add(Me.TextBox1)
        Me.PanelVisionCtrl.Controls.Add(Me.DomainUpDown1)
        Me.PanelVisionCtrl.Controls.Add(Me.Label1)
        Me.PanelVisionCtrl.Controls.Add(Me.Label3)
        Me.PanelVisionCtrl.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.PanelVisionCtrl.Location = New System.Drawing.Point(84, 916)
        Me.PanelVisionCtrl.Name = "PanelVisionCtrl"
        Me.PanelVisionCtrl.Size = New System.Drawing.Size(768, 28)
        Me.PanelVisionCtrl.TabIndex = 2
        '
        'NumericUpDown_Brightness
        '
        Me.NumericUpDown_Brightness.Location = New System.Drawing.Point(302, 4)
        Me.NumericUpDown_Brightness.Maximum = New Decimal(New Integer() {255, 0, 0, 0})
        Me.NumericUpDown_Brightness.Name = "NumericUpDown_Brightness"
        Me.NumericUpDown_Brightness.Size = New System.Drawing.Size(40, 20)
        Me.NumericUpDown_Brightness.TabIndex = 76
        Me.NumericUpDown_Brightness.Value = New Decimal(New Integer() {10, 0, 0, 0})
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
        Me.TextBoxRobotZ.Location = New System.Drawing.Point(668, 4)
        Me.TextBoxRobotZ.Name = "TextBoxRobotZ"
        Me.TextBoxRobotZ.ReadOnly = True
        Me.TextBoxRobotZ.Size = New System.Drawing.Size(56, 20)
        Me.TextBoxRobotZ.TabIndex = 74
        Me.TextBoxRobotZ.Text = "Z 100.000"
        '
        'CheckBoxLockY
        '
        Me.CheckBoxLockY.BackColor = System.Drawing.SystemColors.Control
        Me.CheckBoxLockY.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CheckBoxLockY.Image = CType(resources.GetObject("CheckBoxLockY.Image"), System.Drawing.Image)
        Me.CheckBoxLockY.Location = New System.Drawing.Point(616, 5)
        Me.CheckBoxLockY.Name = "CheckBoxLockY"
        Me.CheckBoxLockY.Size = New System.Drawing.Size(40, 16)
        Me.CheckBoxLockY.TabIndex = 73
        '
        'TextBoxRobotY
        '
        Me.TextBoxRobotY.Location = New System.Drawing.Point(560, 4)
        Me.TextBoxRobotY.Name = "TextBoxRobotY"
        Me.TextBoxRobotY.ReadOnly = True
        Me.TextBoxRobotY.Size = New System.Drawing.Size(56, 20)
        Me.TextBoxRobotY.TabIndex = 72
        Me.TextBoxRobotY.Text = "Y 100.000"
        '
        'CheckBoxLockX
        '
        Me.CheckBoxLockX.BackColor = System.Drawing.SystemColors.Control
        Me.CheckBoxLockX.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CheckBoxLockX.Image = CType(resources.GetObject("CheckBoxLockX.Image"), System.Drawing.Image)
        Me.CheckBoxLockX.Location = New System.Drawing.Point(508, 5)
        Me.CheckBoxLockX.Name = "CheckBoxLockX"
        Me.CheckBoxLockX.Size = New System.Drawing.Size(40, 16)
        Me.CheckBoxLockX.TabIndex = 71
        '
        'TextBoxRobotX
        '
        Me.TextBoxRobotX.Location = New System.Drawing.Point(452, 4)
        Me.TextBoxRobotX.Name = "TextBoxRobotX"
        Me.TextBoxRobotX.ReadOnly = True
        Me.TextBoxRobotX.Size = New System.Drawing.Size(56, 20)
        Me.TextBoxRobotX.TabIndex = 6
        Me.TextBoxRobotX.Text = "X 100.000"
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(409, 6)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(64, 23)
        Me.Label4.TabIndex = 7
        Me.Label4.Text = "Robot"
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(52, 4)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.ReadOnly = True
        Me.TextBox1.Size = New System.Drawing.Size(136, 20)
        Me.TextBox1.TabIndex = 0
        Me.TextBox1.Text = "X: 100.000,  Y: 100.000"
        '
        'DomainUpDown1
        '
        Me.DomainUpDown1.Items.Add("Aa")
        Me.DomainUpDown1.Items.Add("b")
        Me.DomainUpDown1.Items.Add("c")
        Me.DomainUpDown1.Location = New System.Drawing.Point(301, 4)
        Me.DomainUpDown1.Name = "DomainUpDown1"
        Me.DomainUpDown1.Size = New System.Drawing.Size(40, 20)
        Me.DomainUpDown1.TabIndex = 1
        Me.DomainUpDown1.Text = "50"
        Me.DomainUpDown1.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(245, 6)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(64, 23)
        Me.Label1.TabIndex = 3
        Me.Label1.Text = "Brightness"
        '
        'Label3
        '
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label3.Location = New System.Drawing.Point(4, 6)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(40, 16)
        Me.Label3.TabIndex = 5
        Me.Label3.Text = "Cursor"
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
        Me.imageListElement.TransparentColor = System.Drawing.Color.White
        '
        'ImageListReference
        '
        Me.ImageListReference.ImageSize = New System.Drawing.Size(30, 30)
        Me.ImageListReference.ImageStream = CType(resources.GetObject("ImageListReference.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.ImageListReference.TransparentColor = System.Drawing.Color.White
        '
        'TBrefer
        '
        Me.TBrefer.AllowDrop = True
        Me.TBrefer.Anchor = System.Windows.Forms.AnchorStyles.Left
        Me.TBrefer.Buttons.AddRange(New System.Windows.Forms.ToolBarButton() {Me.TBFiducialPt, Me.TBHeightPt, Me.TBReferencePt, Me.TBRejectPt})
        Me.TBrefer.ButtonSize = New System.Drawing.Size(42, 42)
        Me.TBrefer.Dock = System.Windows.Forms.DockStyle.None
        Me.TBrefer.DropDownArrows = True
        Me.TBrefer.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TBrefer.ImageList = Me.ImageListReference
        Me.TBrefer.Location = New System.Drawing.Point(0, 349)
        Me.TBrefer.Name = "TBrefer"
        Me.TBrefer.ShowToolTips = True
        Me.TBrefer.Size = New System.Drawing.Size(86, 90)
        Me.TBrefer.TabIndex = 82
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
        'TBElements
        '
        Me.TBElements.Anchor = System.Windows.Forms.AnchorStyles.Left
        Me.TBElements.Buttons.AddRange(New System.Windows.Forms.ToolBarButton() {Me.TBBDot, Me.TBBLine, Me.TBBArc, Me.TBBRectangle, Me.TBBCircle, Me.TBBFilledRectangle, Me.TBBFilledCircle, Me.TBBLink, Me.TBBChipEdge, Me.TBBMove, Me.TBBWait, Me.TBBPurge, Me.TBBClean, Me.TBBQC, Me.TBBSubPattern, Me.TBBArray, Me.TBBGetIO, Me.TBBSetIO, Me.ToolBarButton6, Me.TBBOffest, Me.TBBMeasure})
        Me.TBElements.ButtonSize = New System.Drawing.Size(42, 42)
        Me.TBElements.Dock = System.Windows.Forms.DockStyle.None
        Me.TBElements.DropDownArrows = True
        Me.TBElements.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TBElements.ImageList = Me.imageListElement
        Me.TBElements.Location = New System.Drawing.Point(0, 451)
        Me.TBElements.Name = "TBElements"
        Me.TBElements.ShowToolTips = True
        Me.TBElements.Size = New System.Drawing.Size(88, 468)
        Me.TBElements.TabIndex = 83
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
        'ToolBarButton6
        '
        Me.ToolBarButton6.Style = System.Windows.Forms.ToolBarButtonStyle.Separator
        '
        'TBBOffest
        '
        Me.TBBOffest.ImageIndex = 18
        Me.TBBOffest.Text = "    Offset"
        Me.TBBOffest.ToolTipText = "Offset"
        '
        'TBBMeasure
        '
        Me.TBBMeasure.ImageIndex = 19
        Me.TBBMeasure.Text = "    Measure"
        Me.TBBMeasure.ToolTipText = "Measure"
        '
        'LabelMessege
        '
        Me.LabelMessege.BackColor = System.Drawing.SystemColors.Info
        Me.LabelMessege.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.LabelMessege.Font = New System.Drawing.Font("Microsoft Sans Serif", 16.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelMessege.ForeColor = System.Drawing.Color.Blue
        Me.LabelMessege.Location = New System.Drawing.Point(0, 299)
        Me.LabelMessege.Name = "LabelMessege"
        Me.LabelMessege.Size = New System.Drawing.Size(684, 42)
        Me.LabelMessege.TabIndex = 84
        Me.LabelMessege.Text = "Message Bar"
        '
        'ImageListYesNo
        '
        Me.ImageListYesNo.ImageSize = New System.Drawing.Size(30, 30)
        Me.ImageListYesNo.ImageStream = CType(resources.GetObject("ImageListYesNo.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.ImageListYesNo.TransparentColor = System.Drawing.Color.White
        '
        'RBLNeedleMode
        '
        Me.RBLNeedleMode.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.RBLNeedleMode.Location = New System.Drawing.Point(223, 14)
        Me.RBLNeedleMode.Name = "RBLNeedleMode"
        Me.RBLNeedleMode.Size = New System.Drawing.Size(96, 24)
        Me.RBLNeedleMode.TabIndex = 89
        Me.RBLNeedleMode.Text = "L Needle"
        '
        'RBRNeedleMode
        '
        Me.RBRNeedleMode.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.RBRNeedleMode.Location = New System.Drawing.Point(323, 14)
        Me.RBRNeedleMode.Name = "RBRNeedleMode"
        Me.RBRNeedleMode.TabIndex = 88
        Me.RBRNeedleMode.Text = "R Needle"
        '
        'Label5
        '
        Me.Label5.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label5.Location = New System.Drawing.Point(9, 15)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(112, 23)
        Me.Label5.TabIndex = 87
        Me.Label5.Text = "Teach Mode:"
        '
        'RBVisionMode
        '
        Me.RBVisionMode.Checked = True
        Me.RBVisionMode.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.RBVisionMode.Location = New System.Drawing.Point(123, 14)
        Me.RBVisionMode.Name = "RBVisionMode"
        Me.RBVisionMode.Size = New System.Drawing.Size(80, 24)
        Me.RBVisionMode.TabIndex = 86
        Me.RBVisionMode.TabStop = True
        Me.RBVisionMode.Text = "Vision"
        '
        'RBProgramEditor
        '
        Me.RBProgramEditor.Appearance = System.Windows.Forms.Appearance.Button
        Me.RBProgramEditor.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.RBProgramEditor.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.RBProgramEditor.Location = New System.Drawing.Point(159, 3)
        Me.RBProgramEditor.Name = "RBProgramEditor"
        Me.RBProgramEditor.Size = New System.Drawing.Size(144, 31)
        Me.RBProgramEditor.TabIndex = 91
        Me.RBProgramEditor.Text = "Program Editor"
        Me.RBProgramEditor.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'RBBasicSetup
        '
        Me.RBBasicSetup.Appearance = System.Windows.Forms.Appearance.Button
        Me.RBBasicSetup.BackColor = System.Drawing.SystemColors.Control
        Me.RBBasicSetup.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.RBBasicSetup.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.RBBasicSetup.Location = New System.Drawing.Point(7, 3)
        Me.RBBasicSetup.Name = "RBBasicSetup"
        Me.RBBasicSetup.Size = New System.Drawing.Size(144, 31)
        Me.RBBasicSetup.TabIndex = 92
        Me.RBBasicSetup.Text = "Basic Setup"
        Me.RBBasicSetup.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'CBExpandSpreadsheet
        '
        Me.CBExpandSpreadsheet.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.CBExpandSpreadsheet.Location = New System.Drawing.Point(1084, 7)
        Me.CBExpandSpreadsheet.Name = "CBExpandSpreadsheet"
        Me.CBExpandSpreadsheet.Size = New System.Drawing.Size(192, 24)
        Me.CBExpandSpreadsheet.TabIndex = 93
        Me.CBExpandSpreadsheet.Text = "Expand Spreadsheet"
        Me.CBExpandSpreadsheet.Visible = False
        '
        'Panel4
        '
        Me.Panel4.BackColor = System.Drawing.SystemColors.ScrollBar
        Me.Panel4.Controls.Add(Me.RBVisionMode)
        Me.Panel4.Controls.Add(Me.RBLNeedleMode)
        Me.Panel4.Controls.Add(Me.RBRNeedleMode)
        Me.Panel4.Controls.Add(Me.Label5)
        Me.Panel4.Location = New System.Drawing.Point(852, 294)
        Me.Panel4.Name = "Panel4"
        Me.Panel4.Size = New System.Drawing.Size(428, 50)
        Me.Panel4.TabIndex = 94
        '
        'TBYesNo
        '
        Me.TBYesNo.BackColor = System.Drawing.Color.FromArgb(CType(192, Byte), CType(255, Byte), CType(255, Byte))
        Me.TBYesNo.Buttons.AddRange(New System.Windows.Forms.ToolBarButton() {Me.TBBOk, Me.TBBCancel})
        Me.TBYesNo.ButtonSize = New System.Drawing.Size(42, 42)
        Me.TBYesNo.Dock = System.Windows.Forms.DockStyle.None
        Me.TBYesNo.DropDownArrows = True
        Me.TBYesNo.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.TBYesNo.ImageList = Me.ImageListYesNo
        Me.TBYesNo.Location = New System.Drawing.Point(768, 295)
        Me.TBYesNo.Name = "TBYesNo"
        Me.TBYesNo.ShowToolTips = True
        Me.TBYesNo.Size = New System.Drawing.Size(84, 48)
        Me.TBYesNo.TabIndex = 0
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
        'TBEditOnly
        '
        Me.TBEditOnly.BackColor = System.Drawing.Color.FromArgb(CType(192, Byte), CType(255, Byte), CType(255, Byte))
        Me.TBEditOnly.Buttons.AddRange(New System.Windows.Forms.ToolBarButton() {Me.TBBEditOnly})
        Me.TBEditOnly.ButtonSize = New System.Drawing.Size(42, 42)
        Me.TBEditOnly.Dock = System.Windows.Forms.DockStyle.None
        Me.TBEditOnly.DropDownArrows = True
        Me.TBEditOnly.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.TBEditOnly.ImageList = Me.ImageListYesNo
        Me.TBEditOnly.Location = New System.Drawing.Point(726, 295)
        Me.TBEditOnly.Name = "TBEditOnly"
        Me.TBEditOnly.ShowToolTips = True
        Me.TBEditOnly.Size = New System.Drawing.Size(42, 48)
        Me.TBEditOnly.TabIndex = 106
        Me.TBEditOnly.Visible = False
        '
        'TBBEditOnly
        '
        Me.TBBEditOnly.ImageIndex = 3
        Me.TBBEditOnly.ToolTipText = "Edit Element"
        '
        'TBNextEdit
        '
        Me.TBNextEdit.BackColor = System.Drawing.Color.FromArgb(CType(192, Byte), CType(255, Byte), CType(255, Byte))
        Me.TBNextEdit.Buttons.AddRange(New System.Windows.Forms.ToolBarButton() {Me.TBBSwitch, Me.TBBEdit})
        Me.TBNextEdit.ButtonSize = New System.Drawing.Size(42, 42)
        Me.TBNextEdit.Dock = System.Windows.Forms.DockStyle.None
        Me.TBNextEdit.DropDownArrows = True
        Me.TBNextEdit.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.TBNextEdit.ImageList = Me.ImageListYesNo
        Me.TBNextEdit.Location = New System.Drawing.Point(684, 295)
        Me.TBNextEdit.Name = "TBNextEdit"
        Me.TBNextEdit.ShowToolTips = True
        Me.TBNextEdit.Size = New System.Drawing.Size(84, 48)
        Me.TBNextEdit.TabIndex = 105
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
        'PanelMotion
        '
        Me.PanelMotion.BackColor = System.Drawing.Color.FromArgb(CType(192, Byte), CType(192, Byte), CType(255, Byte))
        Me.PanelMotion.Controls.Add(Me.ComboBox1)
        Me.PanelMotion.Controls.Add(Me.Label27)
        Me.PanelMotion.Controls.Add(Me.PBYellow)
        Me.PanelMotion.Controls.Add(Me.PBRed)
        Me.PanelMotion.Controls.Add(Me.PBGreen)
        Me.PanelMotion.Controls.Add(Me.TBOperation)
        Me.PanelMotion.Location = New System.Drawing.Point(1000, 341)
        Me.PanelMotion.Name = "PanelMotion"
        Me.PanelMotion.Size = New System.Drawing.Size(280, 221)
        Me.PanelMotion.TabIndex = 101
        '
        'ComboBox1
        '
        Me.ComboBox1.Items.AddRange(New Object() {"Vision Mode", "Dry Needle Mode", "Dry Needle (Left Only)", "Dry Needle (Right Only)", "Wet Needle Mode", "Wet Needle (Left Only)", "Wet Needle (Right Only)"})
        Me.ComboBox1.Location = New System.Drawing.Point(34, 104)
        Me.ComboBox1.Name = "ComboBox1"
        Me.ComboBox1.Size = New System.Drawing.Size(214, 31)
        Me.ComboBox1.TabIndex = 52
        Me.ComboBox1.Text = "Vision Mode"
        '
        'Label27
        '
        Me.Label27.BackColor = System.Drawing.Color.FromArgb(CType(192, Byte), CType(192, Byte), CType(255, Byte))
        Me.Label27.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label27.Location = New System.Drawing.Point(78, 80)
        Me.Label27.Name = "Label27"
        Me.Label27.Size = New System.Drawing.Size(128, 16)
        Me.Label27.TabIndex = 50
        Me.Label27.Text = "Test Run Mode:"
        Me.Label27.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'PBYellow
        '
        Me.PBYellow.Image = CType(resources.GetObject("PBYellow.Image"), System.Drawing.Image)
        Me.PBYellow.Location = New System.Drawing.Point(36, 16)
        Me.PBYellow.Name = "PBYellow"
        Me.PBYellow.Size = New System.Drawing.Size(209, 56)
        Me.PBYellow.TabIndex = 62
        Me.PBYellow.TabStop = False
        Me.PBYellow.Visible = False
        '
        'PBRed
        '
        Me.PBRed.Image = CType(resources.GetObject("PBRed.Image"), System.Drawing.Image)
        Me.PBRed.Location = New System.Drawing.Point(36, 16)
        Me.PBRed.Name = "PBRed"
        Me.PBRed.Size = New System.Drawing.Size(209, 56)
        Me.PBRed.TabIndex = 61
        Me.PBRed.TabStop = False
        '
        'PBGreen
        '
        Me.PBGreen.Image = CType(resources.GetObject("PBGreen.Image"), System.Drawing.Image)
        Me.PBGreen.Location = New System.Drawing.Point(36, 16)
        Me.PBGreen.Name = "PBGreen"
        Me.PBGreen.Size = New System.Drawing.Size(209, 56)
        Me.PBGreen.TabIndex = 63
        Me.PBGreen.TabStop = False
        Me.PBGreen.Visible = False
        '
        'TBOperation
        '
        Me.TBOperation.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.TBOperation.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.TBOperation.Buttons.AddRange(New System.Windows.Forms.ToolBarButton() {Me.ToolBarButton1, Me.ToolBarButton2, Me.ToolBarButton5})
        Me.TBOperation.ButtonSize = New System.Drawing.Size(70, 50)
        Me.TBOperation.Dock = System.Windows.Forms.DockStyle.None
        Me.TBOperation.DropDownArrows = True
        Me.TBOperation.ImageList = Me.ImageListOper
        Me.TBOperation.Location = New System.Drawing.Point(34, 149)
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
        'ImageListOper
        '
        Me.ImageListOper.ImageSize = New System.Drawing.Size(36, 36)
        Me.ImageListOper.ImageStream = CType(resources.GetObject("ImageListOper.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.ImageListOper.TransparentColor = System.Drawing.Color.Transparent
        '
        'BTClean
        '
        Me.BTClean.BackColor = System.Drawing.SystemColors.InactiveCaptionText
        Me.BTClean.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.BTClean.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.BTClean.ImageIndex = 5
        Me.BTClean.ImageList = Me.ImageListGeneralTools
        Me.BTClean.Location = New System.Drawing.Point(926, 482)
        Me.BTClean.Name = "BTClean"
        Me.BTClean.Size = New System.Drawing.Size(74, 60)
        Me.BTClean.TabIndex = 58
        Me.BTClean.Text = "Clean"
        Me.BTClean.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'BTNeedleCal
        '
        Me.BTNeedleCal.BackColor = System.Drawing.SystemColors.InactiveCaptionText
        Me.BTNeedleCal.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.BTNeedleCal.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.BTNeedleCal.ImageIndex = 3
        Me.BTNeedleCal.ImageList = Me.ImageListGeneralTools
        Me.BTNeedleCal.Location = New System.Drawing.Point(852, 422)
        Me.BTNeedleCal.Name = "BTNeedleCal"
        Me.BTNeedleCal.Size = New System.Drawing.Size(74, 60)
        Me.BTNeedleCal.TabIndex = 56
        Me.BTNeedleCal.Text = "Ndl. Cal."
        Me.BTNeedleCal.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'BTVolCal
        '
        Me.BTVolCal.BackColor = System.Drawing.SystemColors.InactiveCaptionText
        Me.BTVolCal.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.BTVolCal.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.BTVolCal.ImageIndex = 2
        Me.BTVolCal.ImageList = Me.ImageListGeneralTools
        Me.BTVolCal.Location = New System.Drawing.Point(926, 422)
        Me.BTVolCal.Name = "BTVolCal"
        Me.BTVolCal.Size = New System.Drawing.Size(74, 60)
        Me.BTVolCal.TabIndex = 55
        Me.BTVolCal.Text = "Vol. Cal."
        Me.BTVolCal.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'BTHome
        '
        Me.BTHome.BackColor = System.Drawing.SystemColors.InactiveCaptionText
        Me.BTHome.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.BTHome.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.BTHome.ImageIndex = 0
        Me.BTHome.ImageList = Me.ImageListGeneralTools
        Me.BTHome.Location = New System.Drawing.Point(852, 362)
        Me.BTHome.Name = "BTHome"
        Me.BTHome.Size = New System.Drawing.Size(74, 60)
        Me.BTHome.TabIndex = 53
        Me.BTHome.Text = "Home"
        Me.BTHome.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'ImageListMultiField
        '
        Me.ImageListMultiField.ImageSize = New System.Drawing.Size(36, 28)
        Me.ImageListMultiField.ImageStream = CType(resources.GetObject("ImageListMultiField.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.ImageListMultiField.TransparentColor = System.Drawing.Color.White
        '
        'TBMultiDisplayField
        '
        Me.TBMultiDisplayField.BackColor = System.Drawing.SystemColors.Menu
        Me.TBMultiDisplayField.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.TBMultiDisplayField.Buttons.AddRange(New System.Windows.Forms.ToolBarButton() {Me.TBBGraphDisp, Me.TBBStepper, Me.TBBDispenser, Me.TBBThermal, Me.TBBConveyor})
        Me.TBMultiDisplayField.ButtonSize = New System.Drawing.Size(80, 56)
        Me.TBMultiDisplayField.Dock = System.Windows.Forms.DockStyle.None
        Me.TBMultiDisplayField.DropDownArrows = True
        Me.TBMultiDisplayField.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.TBMultiDisplayField.ForeColor = System.Drawing.Color.Black
        Me.TBMultiDisplayField.ImageList = Me.ImageListMultiField
        Me.TBMultiDisplayField.Location = New System.Drawing.Point(1200, 562)
        Me.TBMultiDisplayField.Name = "TBMultiDisplayField"
        Me.TBMultiDisplayField.ShowToolTips = True
        Me.TBMultiDisplayField.Size = New System.Drawing.Size(80, 287)
        Me.TBMultiDisplayField.TabIndex = 105
        '
        'TBBGraphDisp
        '
        Me.TBBGraphDisp.ImageIndex = 0
        Me.TBBGraphDisp.Pushed = True
        Me.TBBGraphDisp.Style = System.Windows.Forms.ToolBarButtonStyle.ToggleButton
        Me.TBBGraphDisp.Text = "Graph. Disp."
        '
        'TBBStepper
        '
        Me.TBBStepper.ImageIndex = 1
        Me.TBBStepper.Style = System.Windows.Forms.ToolBarButtonStyle.ToggleButton
        Me.TBBStepper.Text = "Stepper"
        '
        'TBBDispenser
        '
        Me.TBBDispenser.ImageIndex = 2
        Me.TBBDispenser.Style = System.Windows.Forms.ToolBarButtonStyle.ToggleButton
        Me.TBBDispenser.Text = "Dispenser"
        '
        'TBBThermal
        '
        Me.TBBThermal.ImageIndex = 3
        Me.TBBThermal.Style = System.Windows.Forms.ToolBarButtonStyle.ToggleButton
        Me.TBBThermal.Text = "Thermal"
        '
        'TBBConveyor
        '
        Me.TBBConveyor.ImageIndex = 4
        Me.TBBConveyor.Style = System.Windows.Forms.ToolBarButtonStyle.ToggleButton
        Me.TBBConveyor.Text = "Conveyor"
        '
        'CBDoorLock
        '
        Me.CBDoorLock.Appearance = System.Windows.Forms.Appearance.Button
        Me.CBDoorLock.BackColor = System.Drawing.Color.LightGray
        Me.CBDoorLock.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.CBDoorLock.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.CBDoorLock.ImageIndex = 5
        Me.CBDoorLock.ImageList = Me.ImageListMultiField
        Me.CBDoorLock.Location = New System.Drawing.Point(1200, 888)
        Me.CBDoorLock.Name = "CBDoorLock"
        Me.CBDoorLock.Size = New System.Drawing.Size(80, 56)
        Me.CBDoorLock.TabIndex = 117
        Me.CBDoorLock.Text = "Lock Door"
        Me.CBDoorLock.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'PanelGraphDisp
        '
        Me.PanelGraphDisp.BackColor = System.Drawing.SystemColors.ScrollBar
        Me.PanelGraphDisp.Location = New System.Drawing.Point(852, 562)
        Me.PanelGraphDisp.Name = "PanelGraphDisp"
        Me.PanelGraphDisp.Size = New System.Drawing.Size(348, 382)
        Me.PanelGraphDisp.TabIndex = 119
        '
        'PanelStepper
        '
        Me.PanelStepper.BackColor = System.Drawing.SystemColors.ScrollBar
        Me.PanelStepper.Controls.Add(Me.ComboBoxFineStep)
        Me.PanelStepper.Controls.Add(Me.BTStepZdown)
        Me.PanelStepper.Controls.Add(Me.BTStepZup)
        Me.PanelStepper.Controls.Add(Me.BTStepXminus)
        Me.PanelStepper.Controls.Add(Me.BTStepXplus)
        Me.PanelStepper.Controls.Add(Me.BTStepYminus)
        Me.PanelStepper.Controls.Add(Me.BTStepYplus)
        Me.PanelStepper.Controls.Add(Me.Label6)
        Me.PanelStepper.Controls.Add(Me.DomainUpDown3)
        Me.PanelStepper.Controls.Add(Me.LabelStep)
        Me.PanelStepper.Location = New System.Drawing.Point(852, 562)
        Me.PanelStepper.Name = "PanelStepper"
        Me.PanelStepper.Size = New System.Drawing.Size(348, 382)
        Me.PanelStepper.TabIndex = 0
        Me.PanelStepper.Visible = False
        '
        'ComboBoxFineStep
        '
        Me.ComboBoxFineStep.Items.AddRange(New Object() {"0.001", "0.01", "0.1", "1", "10"})
        Me.ComboBoxFineStep.Location = New System.Drawing.Point(173, 288)
        Me.ComboBoxFineStep.Name = "ComboBoxFineStep"
        Me.ComboBoxFineStep.Size = New System.Drawing.Size(72, 31)
        Me.ComboBoxFineStep.TabIndex = 9
        Me.ComboBoxFineStep.Text = "0.1"
        '
        'BTStepZdown
        '
        Me.BTStepZdown.BackColor = System.Drawing.Color.Linen
        Me.BTStepZdown.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.5!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.BTStepZdown.Image = CType(resources.GetObject("BTStepZdown.Image"), System.Drawing.Image)
        Me.BTStepZdown.ImageAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.BTStepZdown.Location = New System.Drawing.Point(269, 144)
        Me.BTStepZdown.Name = "BTStepZdown"
        Me.BTStepZdown.Size = New System.Drawing.Size(48, 72)
        Me.BTStepZdown.TabIndex = 8
        Me.BTStepZdown.Text = "Dn"
        Me.BTStepZdown.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'BTStepZup
        '
        Me.BTStepZup.BackColor = System.Drawing.Color.Linen
        Me.BTStepZup.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.5!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.BTStepZup.Image = CType(resources.GetObject("BTStepZup.Image"), System.Drawing.Image)
        Me.BTStepZup.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.BTStepZup.Location = New System.Drawing.Point(269, 56)
        Me.BTStepZup.Name = "BTStepZup"
        Me.BTStepZup.Size = New System.Drawing.Size(48, 72)
        Me.BTStepZup.TabIndex = 7
        Me.BTStepZup.Text = "Up"
        Me.BTStepZup.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'BTStepXminus
        '
        Me.BTStepXminus.BackColor = System.Drawing.Color.LightGoldenrodYellow
        Me.BTStepXminus.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.5!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.BTStepXminus.Image = CType(resources.GetObject("BTStepXminus.Image"), System.Drawing.Image)
        Me.BTStepXminus.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BTStepXminus.Location = New System.Drawing.Point(29, 112)
        Me.BTStepXminus.Name = "BTStepXminus"
        Me.BTStepXminus.Size = New System.Drawing.Size(80, 48)
        Me.BTStepXminus.TabIndex = 6
        Me.BTStepXminus.Text = "X-"
        Me.BTStepXminus.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'BTStepXplus
        '
        Me.BTStepXplus.BackColor = System.Drawing.Color.LightGoldenrodYellow
        Me.BTStepXplus.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.5!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.BTStepXplus.Image = CType(resources.GetObject("BTStepXplus.Image"), System.Drawing.Image)
        Me.BTStepXplus.ImageAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.BTStepXplus.Location = New System.Drawing.Point(157, 112)
        Me.BTStepXplus.Name = "BTStepXplus"
        Me.BTStepXplus.Size = New System.Drawing.Size(80, 48)
        Me.BTStepXplus.TabIndex = 5
        Me.BTStepXplus.Text = "X+"
        Me.BTStepXplus.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'BTStepYminus
        '
        Me.BTStepYminus.BackColor = System.Drawing.Color.LightGoldenrodYellow
        Me.BTStepYminus.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.5!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.BTStepYminus.Image = CType(resources.GetObject("BTStepYminus.Image"), System.Drawing.Image)
        Me.BTStepYminus.ImageAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.BTStepYminus.Location = New System.Drawing.Point(109, 160)
        Me.BTStepYminus.Name = "BTStepYminus"
        Me.BTStepYminus.Size = New System.Drawing.Size(48, 80)
        Me.BTStepYminus.TabIndex = 4
        Me.BTStepYminus.Text = "Y-"
        Me.BTStepYminus.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'BTStepYplus
        '
        Me.BTStepYplus.BackColor = System.Drawing.Color.LightGoldenrodYellow
        Me.BTStepYplus.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.5!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.BTStepYplus.Image = CType(resources.GetObject("BTStepYplus.Image"), System.Drawing.Image)
        Me.BTStepYplus.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.BTStepYplus.Location = New System.Drawing.Point(109, 32)
        Me.BTStepYplus.Name = "BTStepYplus"
        Me.BTStepYplus.Size = New System.Drawing.Size(48, 80)
        Me.BTStepYplus.TabIndex = 3
        Me.BTStepYplus.Text = "Y+"
        Me.BTStepYplus.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'Label6
        '
        Me.Label6.Location = New System.Drawing.Point(253, 288)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(40, 24)
        Me.Label6.TabIndex = 2
        Me.Label6.Text = "mm"
        Me.Label6.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'DomainUpDown3
        '
        Me.DomainUpDown3.Location = New System.Drawing.Point(173, 232)
        Me.DomainUpDown3.Name = "DomainUpDown3"
        Me.DomainUpDown3.Size = New System.Drawing.Size(72, 27)
        Me.DomainUpDown3.TabIndex = 1
        Me.DomainUpDown3.Text = "0.002"
        Me.DomainUpDown3.Visible = False
        '
        'LabelStep
        '
        Me.LabelStep.Location = New System.Drawing.Point(61, 288)
        Me.LabelStep.Name = "LabelStep"
        Me.LabelStep.Size = New System.Drawing.Size(88, 24)
        Me.LabelStep.TabIndex = 0
        Me.LabelStep.Text = "Fine Step:"
        Me.LabelStep.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'PanelDispenser
        '
        Me.PanelDispenser.BackColor = System.Drawing.SystemColors.ScrollBar
        Me.PanelDispenser.Controls.Add(Me.GB_RTP)
        Me.PanelDispenser.Controls.Add(Me.GB_RJetter)
        Me.PanelDispenser.Controls.Add(Me.GB_LJetter)
        Me.PanelDispenser.Controls.Add(Me.GB_LTP)
        Me.PanelDispenser.Location = New System.Drawing.Point(852, 562)
        Me.PanelDispenser.Name = "PanelDispenser"
        Me.PanelDispenser.Size = New System.Drawing.Size(348, 382)
        Me.PanelDispenser.TabIndex = 16
        Me.PanelDispenser.Visible = False
        '
        'GB_RTP
        '
        Me.GB_RTP.Controls.Add(Me.TB_Pro_RRPM)
        Me.GB_RTP.Controls.Add(Me.BT_Pro_RDownload)
        Me.GB_RTP.Controls.Add(Me.BT_Pro_RUpdate)
        Me.GB_RTP.Controls.Add(Me.TB_Pro_RSuck)
        Me.GB_RTP.Controls.Add(Me.TB_Pro_RREV)
        Me.GB_RTP.Controls.Add(Me.LB_Pro_RSuck)
        Me.GB_RTP.Controls.Add(Me.LB_Pro_RRPM)
        Me.GB_RTP.Controls.Add(Me.LB_Pro_RPressure)
        Me.GB_RTP.Controls.Add(Me.TB_Pro_RPressure)
        Me.GB_RTP.Controls.Add(Me.LB_Pro_RREV)
        Me.GB_RTP.Controls.Add(Me.BT_Pro_REdit)
        Me.GB_RTP.Controls.Add(Me.LB_IDSPressure_RTP)
        Me.GB_RTP.Controls.Add(Me.TB_Pro_RProg)
        Me.GB_RTP.Controls.Add(Me.LB_Pro_RProg)
        Me.GB_RTP.Controls.Add(Me.TB_Pro_RValve)
        Me.GB_RTP.Controls.Add(Me.LB_Pro_RValve)
        Me.GB_RTP.ForeColor = System.Drawing.SystemColors.ActiveCaption
        Me.GB_RTP.Location = New System.Drawing.Point(8, 186)
        Me.GB_RTP.Name = "GB_RTP"
        Me.GB_RTP.Size = New System.Drawing.Size(331, 184)
        Me.GB_RTP.TabIndex = 1
        Me.GB_RTP.TabStop = False
        Me.GB_RTP.Text = "RIGHT - Time Pressure/AVC"
        Me.GB_RTP.Visible = False
        '
        'TB_Pro_RRPM
        '
        Me.TB_Pro_RRPM.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.TB_Pro_RRPM.Location = New System.Drawing.Point(184, 32)
        Me.TB_Pro_RRPM.Name = "TB_Pro_RRPM"
        Me.TB_Pro_RRPM.Size = New System.Drawing.Size(48, 26)
        Me.TB_Pro_RRPM.TabIndex = 18
        Me.TB_Pro_RRPM.Text = "100"
        '
        'BT_Pro_RDownload
        '
        Me.BT_Pro_RDownload.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.BT_Pro_RDownload.ForeColor = System.Drawing.SystemColors.ControlText
        Me.BT_Pro_RDownload.Location = New System.Drawing.Point(116, 96)
        Me.BT_Pro_RDownload.Name = "BT_Pro_RDownload"
        Me.BT_Pro_RDownload.Size = New System.Drawing.Size(88, 32)
        Me.BT_Pro_RDownload.TabIndex = 26
        Me.BT_Pro_RDownload.Text = "Download"
        '
        'BT_Pro_RUpdate
        '
        Me.BT_Pro_RUpdate.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.BT_Pro_RUpdate.ForeColor = System.Drawing.SystemColors.ControlText
        Me.BT_Pro_RUpdate.Location = New System.Drawing.Point(16, 96)
        Me.BT_Pro_RUpdate.Name = "BT_Pro_RUpdate"
        Me.BT_Pro_RUpdate.Size = New System.Drawing.Size(88, 32)
        Me.BT_Pro_RUpdate.TabIndex = 25
        Me.BT_Pro_RUpdate.Text = "Update"
        Me.BT_Pro_RUpdate.Visible = False
        '
        'TB_Pro_RSuck
        '
        Me.TB_Pro_RSuck.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.TB_Pro_RSuck.Location = New System.Drawing.Point(80, 64)
        Me.TB_Pro_RSuck.Name = "TB_Pro_RSuck"
        Me.TB_Pro_RSuck.Size = New System.Drawing.Size(48, 26)
        Me.TB_Pro_RSuck.TabIndex = 17
        Me.TB_Pro_RSuck.Text = "-0.25"
        '
        'TB_Pro_RREV
        '
        Me.TB_Pro_RREV.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.TB_Pro_RREV.Location = New System.Drawing.Point(184, 64)
        Me.TB_Pro_RREV.Name = "TB_Pro_RREV"
        Me.TB_Pro_RREV.Size = New System.Drawing.Size(48, 26)
        Me.TB_Pro_RREV.TabIndex = 19
        Me.TB_Pro_RREV.Text = "-0.1"
        '
        'LB_Pro_RSuck
        '
        Me.LB_Pro_RSuck.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.LB_Pro_RSuck.ForeColor = System.Drawing.SystemColors.ControlText
        Me.LB_Pro_RSuck.Location = New System.Drawing.Point(8, 64)
        Me.LB_Pro_RSuck.Name = "LB_Pro_RSuck"
        Me.LB_Pro_RSuck.Size = New System.Drawing.Size(72, 24)
        Me.LB_Pro_RSuck.TabIndex = 22
        Me.LB_Pro_RSuck.Text = "Suck Bk"
        Me.LB_Pro_RSuck.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'LB_Pro_RRPM
        '
        Me.LB_Pro_RRPM.Enabled = False
        Me.LB_Pro_RRPM.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.LB_Pro_RRPM.ForeColor = System.Drawing.SystemColors.ControlText
        Me.LB_Pro_RRPM.Location = New System.Drawing.Point(144, 32)
        Me.LB_Pro_RRPM.Name = "LB_Pro_RRPM"
        Me.LB_Pro_RRPM.Size = New System.Drawing.Size(48, 24)
        Me.LB_Pro_RRPM.TabIndex = 23
        Me.LB_Pro_RRPM.Text = "RPM"
        Me.LB_Pro_RRPM.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'LB_Pro_RPressure
        '
        Me.LB_Pro_RPressure.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.LB_Pro_RPressure.ForeColor = System.Drawing.SystemColors.ControlText
        Me.LB_Pro_RPressure.Location = New System.Drawing.Point(8, 32)
        Me.LB_Pro_RPressure.Name = "LB_Pro_RPressure"
        Me.LB_Pro_RPressure.Size = New System.Drawing.Size(72, 24)
        Me.LB_Pro_RPressure.TabIndex = 21
        Me.LB_Pro_RPressure.Text = "Pressure"
        Me.LB_Pro_RPressure.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'TB_Pro_RPressure
        '
        Me.TB_Pro_RPressure.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.TB_Pro_RPressure.Location = New System.Drawing.Point(80, 32)
        Me.TB_Pro_RPressure.Name = "TB_Pro_RPressure"
        Me.TB_Pro_RPressure.Size = New System.Drawing.Size(48, 26)
        Me.TB_Pro_RPressure.TabIndex = 16
        Me.TB_Pro_RPressure.Text = "2.00"
        '
        'LB_Pro_RREV
        '
        Me.LB_Pro_RREV.Enabled = False
        Me.LB_Pro_RREV.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.LB_Pro_RREV.ForeColor = System.Drawing.SystemColors.ControlText
        Me.LB_Pro_RREV.Location = New System.Drawing.Point(144, 64)
        Me.LB_Pro_RREV.Name = "LB_Pro_RREV"
        Me.LB_Pro_RREV.Size = New System.Drawing.Size(40, 24)
        Me.LB_Pro_RREV.TabIndex = 24
        Me.LB_Pro_RREV.Text = "REV"
        Me.LB_Pro_RREV.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'BT_Pro_REdit
        '
        Me.BT_Pro_REdit.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.BT_Pro_REdit.ForeColor = System.Drawing.SystemColors.ControlText
        Me.BT_Pro_REdit.Location = New System.Drawing.Point(220, 96)
        Me.BT_Pro_REdit.Name = "BT_Pro_REdit"
        Me.BT_Pro_REdit.TabIndex = 27
        Me.BT_Pro_REdit.Text = "Allow Edit"
        '
        'LB_IDSPressure_RTP
        '
        Me.LB_IDSPressure_RTP.Enabled = False
        Me.LB_IDSPressure_RTP.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.LB_IDSPressure_RTP.ForeColor = System.Drawing.Color.FromArgb(CType(192, Byte), CType(0, Byte), CType(0, Byte))
        Me.LB_IDSPressure_RTP.Location = New System.Drawing.Point(12, 152)
        Me.LB_IDSPressure_RTP.Name = "LB_IDSPressure_RTP"
        Me.LB_IDSPressure_RTP.Size = New System.Drawing.Size(304, 24)
        Me.LB_IDSPressure_RTP.TabIndex = 28
        Me.LB_IDSPressure_RTP.Text = "Cur. Smart Control: (None)"
        Me.LB_IDSPressure_RTP.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.LB_IDSPressure_RTP.Visible = False
        '
        'TB_Pro_RProg
        '
        Me.TB_Pro_RProg.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.TB_Pro_RProg.Location = New System.Drawing.Point(296, 64)
        Me.TB_Pro_RProg.Name = "TB_Pro_RProg"
        Me.TB_Pro_RProg.Size = New System.Drawing.Size(16, 26)
        Me.TB_Pro_RProg.TabIndex = 19
        Me.TB_Pro_RProg.Text = "4"
        '
        'LB_Pro_RProg
        '
        Me.LB_Pro_RProg.Enabled = False
        Me.LB_Pro_RProg.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.LB_Pro_RProg.ForeColor = System.Drawing.SystemColors.ControlText
        Me.LB_Pro_RProg.Location = New System.Drawing.Point(248, 64)
        Me.LB_Pro_RProg.Name = "LB_Pro_RProg"
        Me.LB_Pro_RProg.Size = New System.Drawing.Size(48, 24)
        Me.LB_Pro_RProg.TabIndex = 24
        Me.LB_Pro_RProg.Text = "Prog"
        Me.LB_Pro_RProg.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'TB_Pro_RValve
        '
        Me.TB_Pro_RValve.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.TB_Pro_RValve.Location = New System.Drawing.Point(296, 32)
        Me.TB_Pro_RValve.Name = "TB_Pro_RValve"
        Me.TB_Pro_RValve.Size = New System.Drawing.Size(16, 26)
        Me.TB_Pro_RValve.TabIndex = 18
        Me.TB_Pro_RValve.Text = "4"
        '
        'LB_Pro_RValve
        '
        Me.LB_Pro_RValve.Enabled = False
        Me.LB_Pro_RValve.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.LB_Pro_RValve.ForeColor = System.Drawing.SystemColors.ControlText
        Me.LB_Pro_RValve.Location = New System.Drawing.Point(248, 32)
        Me.LB_Pro_RValve.Name = "LB_Pro_RValve"
        Me.LB_Pro_RValve.Size = New System.Drawing.Size(48, 24)
        Me.LB_Pro_RValve.TabIndex = 23
        Me.LB_Pro_RValve.Text = "Valve"
        Me.LB_Pro_RValve.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'GB_RJetter
        '
        Me.GB_RJetter.Controls.Add(Me.LB_IDSPressure_RJetter)
        Me.GB_RJetter.Controls.Add(Me.BT_RJetter_Edit)
        Me.GB_RJetter.Controls.Add(Me.TB_RJetter_Path)
        Me.GB_RJetter.Controls.Add(Me.BT_RJetter_Browse)
        Me.GB_RJetter.Controls.Add(Me.BT_RJetter_Open)
        Me.GB_RJetter.Controls.Add(Me.Label30)
        Me.GB_RJetter.ForeColor = System.Drawing.SystemColors.ActiveCaption
        Me.GB_RJetter.Location = New System.Drawing.Point(8, 186)
        Me.GB_RJetter.Name = "GB_RJetter"
        Me.GB_RJetter.Size = New System.Drawing.Size(331, 184)
        Me.GB_RJetter.TabIndex = 26
        Me.GB_RJetter.TabStop = False
        Me.GB_RJetter.Text = "RIGHT - Jetter"
        '
        'LB_IDSPressure_RJetter
        '
        Me.LB_IDSPressure_RJetter.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.LB_IDSPressure_RJetter.ForeColor = System.Drawing.Color.FromArgb(CType(192, Byte), CType(0, Byte), CType(0, Byte))
        Me.LB_IDSPressure_RJetter.Location = New System.Drawing.Point(12, 152)
        Me.LB_IDSPressure_RJetter.Name = "LB_IDSPressure_RJetter"
        Me.LB_IDSPressure_RJetter.Size = New System.Drawing.Size(304, 24)
        Me.LB_IDSPressure_RJetter.TabIndex = 30
        Me.LB_IDSPressure_RJetter.Text = "Cur. Smart Control: P = 3.00, S = -0.25"
        Me.LB_IDSPressure_RJetter.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'BT_RJetter_Edit
        '
        Me.BT_RJetter_Edit.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.BT_RJetter_Edit.ForeColor = System.Drawing.SystemColors.ControlText
        Me.BT_RJetter_Edit.Location = New System.Drawing.Point(220, 96)
        Me.BT_RJetter_Edit.Name = "BT_RJetter_Edit"
        Me.BT_RJetter_Edit.TabIndex = 29
        Me.BT_RJetter_Edit.Text = "Allow Edit"
        '
        'TB_RJetter_Path
        '
        Me.TB_RJetter_Path.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.TB_RJetter_Path.Location = New System.Drawing.Point(56, 64)
        Me.TB_RJetter_Path.Name = "TB_RJetter_Path"
        Me.TB_RJetter_Path.Size = New System.Drawing.Size(256, 26)
        Me.TB_RJetter_Path.TabIndex = 28
        Me.TB_RJetter_Path.Text = "C:\IDS\Dispenser\JP1"
        '
        'BT_RJetter_Browse
        '
        Me.BT_RJetter_Browse.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.BT_RJetter_Browse.ForeColor = System.Drawing.SystemColors.ControlText
        Me.BT_RJetter_Browse.Location = New System.Drawing.Point(72, 96)
        Me.BT_RJetter_Browse.Name = "BT_RJetter_Browse"
        Me.BT_RJetter_Browse.Size = New System.Drawing.Size(75, 32)
        Me.BT_RJetter_Browse.TabIndex = 26
        Me.BT_RJetter_Browse.Text = "Browse"
        '
        'BT_RJetter_Open
        '
        Me.BT_RJetter_Open.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.BT_RJetter_Open.ForeColor = System.Drawing.SystemColors.ControlText
        Me.BT_RJetter_Open.Location = New System.Drawing.Point(72, 24)
        Me.BT_RJetter_Open.Name = "BT_RJetter_Open"
        Me.BT_RJetter_Open.Size = New System.Drawing.Size(184, 32)
        Me.BT_RJetter_Open.TabIndex = 25
        Me.BT_RJetter_Open.Text = "Open Jetter Program"
        '
        'Label30
        '
        Me.Label30.Enabled = False
        Me.Label30.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label30.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label30.Location = New System.Drawing.Point(16, 64)
        Me.Label30.Name = "Label30"
        Me.Label30.Size = New System.Drawing.Size(40, 24)
        Me.Label30.TabIndex = 27
        Me.Label30.Text = "File:"
        Me.Label30.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'GB_LJetter
        '
        Me.GB_LJetter.Controls.Add(Me.LB_IDSPressure_LJetter)
        Me.GB_LJetter.Controls.Add(Me.BT_LJetter_Edit)
        Me.GB_LJetter.Controls.Add(Me.TB_LJetter_Path)
        Me.GB_LJetter.Controls.Add(Me.BT_LJetter_Browse)
        Me.GB_LJetter.Controls.Add(Me.BT_LJetter_Open)
        Me.GB_LJetter.Controls.Add(Me.Label15)
        Me.GB_LJetter.ForeColor = System.Drawing.SystemColors.ActiveCaption
        Me.GB_LJetter.Location = New System.Drawing.Point(8, 2)
        Me.GB_LJetter.Name = "GB_LJetter"
        Me.GB_LJetter.Size = New System.Drawing.Size(331, 184)
        Me.GB_LJetter.TabIndex = 25
        Me.GB_LJetter.TabStop = False
        Me.GB_LJetter.Text = "LEFT - Jetter"
        Me.GB_LJetter.Visible = False
        '
        'LB_IDSPressure_LJetter
        '
        Me.LB_IDSPressure_LJetter.Enabled = False
        Me.LB_IDSPressure_LJetter.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.LB_IDSPressure_LJetter.ForeColor = System.Drawing.Color.FromArgb(CType(192, Byte), CType(0, Byte), CType(0, Byte))
        Me.LB_IDSPressure_LJetter.Location = New System.Drawing.Point(12, 152)
        Me.LB_IDSPressure_LJetter.Name = "LB_IDSPressure_LJetter"
        Me.LB_IDSPressure_LJetter.Size = New System.Drawing.Size(304, 24)
        Me.LB_IDSPressure_LJetter.TabIndex = 24
        Me.LB_IDSPressure_LJetter.Text = "Cur. Smart Control: P = 3.00, S = -0.25"
        Me.LB_IDSPressure_LJetter.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'BT_LJetter_Edit
        '
        Me.BT_LJetter_Edit.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.BT_LJetter_Edit.ForeColor = System.Drawing.SystemColors.ControlText
        Me.BT_LJetter_Edit.Location = New System.Drawing.Point(220, 96)
        Me.BT_LJetter_Edit.Name = "BT_LJetter_Edit"
        Me.BT_LJetter_Edit.TabIndex = 23
        Me.BT_LJetter_Edit.Text = "Allow Edit"
        '
        'TB_LJetter_Path
        '
        Me.TB_LJetter_Path.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.TB_LJetter_Path.Location = New System.Drawing.Point(56, 64)
        Me.TB_LJetter_Path.Name = "TB_LJetter_Path"
        Me.TB_LJetter_Path.Size = New System.Drawing.Size(256, 26)
        Me.TB_LJetter_Path.TabIndex = 17
        Me.TB_LJetter_Path.Text = "C:\IDS\Dispenser\JP1"
        '
        'BT_LJetter_Browse
        '
        Me.BT_LJetter_Browse.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.BT_LJetter_Browse.ForeColor = System.Drawing.SystemColors.ControlText
        Me.BT_LJetter_Browse.Location = New System.Drawing.Point(72, 96)
        Me.BT_LJetter_Browse.Name = "BT_LJetter_Browse"
        Me.BT_LJetter_Browse.Size = New System.Drawing.Size(75, 32)
        Me.BT_LJetter_Browse.TabIndex = 1
        Me.BT_LJetter_Browse.Text = "Browse"
        '
        'BT_LJetter_Open
        '
        Me.BT_LJetter_Open.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.BT_LJetter_Open.ForeColor = System.Drawing.SystemColors.ControlText
        Me.BT_LJetter_Open.Location = New System.Drawing.Point(72, 24)
        Me.BT_LJetter_Open.Name = "BT_LJetter_Open"
        Me.BT_LJetter_Open.Size = New System.Drawing.Size(184, 32)
        Me.BT_LJetter_Open.TabIndex = 0
        Me.BT_LJetter_Open.Text = "Open Jetter Program"
        '
        'Label15
        '
        Me.Label15.Enabled = False
        Me.Label15.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label15.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label15.Location = New System.Drawing.Point(16, 64)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(40, 24)
        Me.Label15.TabIndex = 16
        Me.Label15.Text = "File:"
        Me.Label15.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'GB_LTP
        '
        Me.GB_LTP.Controls.Add(Me.TB_Pro_LRPM)
        Me.GB_LTP.Controls.Add(Me.LB_IDSPressure_LTP)
        Me.GB_LTP.Controls.Add(Me.BT_Pro_LEdit)
        Me.GB_LTP.Controls.Add(Me.BT_Pro_LDownload)
        Me.GB_LTP.Controls.Add(Me.BT_Pro_LUpdate)
        Me.GB_LTP.Controls.Add(Me.TB_Pro_LSuck)
        Me.GB_LTP.Controls.Add(Me.TB_Pro_LREV)
        Me.GB_LTP.Controls.Add(Me.LB_Pro_LSuck)
        Me.GB_LTP.Controls.Add(Me.LB_Pro_LRPM)
        Me.GB_LTP.Controls.Add(Me.LB_Pro_LPressure)
        Me.GB_LTP.Controls.Add(Me.TB_Pro_LPressure)
        Me.GB_LTP.Controls.Add(Me.LB_Pro_LREV)
        Me.GB_LTP.Controls.Add(Me.TB_Pro_LProg)
        Me.GB_LTP.Controls.Add(Me.TB_Pro_LValve)
        Me.GB_LTP.Controls.Add(Me.LB_Pro_LProg)
        Me.GB_LTP.Controls.Add(Me.LB_Pro_LValve)
        Me.GB_LTP.ForeColor = System.Drawing.SystemColors.ActiveCaption
        Me.GB_LTP.Location = New System.Drawing.Point(8, 2)
        Me.GB_LTP.Name = "GB_LTP"
        Me.GB_LTP.Size = New System.Drawing.Size(331, 184)
        Me.GB_LTP.TabIndex = 0
        Me.GB_LTP.TabStop = False
        Me.GB_LTP.Text = "LEFT - Time Pressure/AVC"
        '
        'TB_Pro_LRPM
        '
        Me.TB_Pro_LRPM.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.TB_Pro_LRPM.Location = New System.Drawing.Point(184, 32)
        Me.TB_Pro_LRPM.Name = "TB_Pro_LRPM"
        Me.TB_Pro_LRPM.Size = New System.Drawing.Size(48, 26)
        Me.TB_Pro_LRPM.TabIndex = 31
        Me.TB_Pro_LRPM.Text = "100"
        '
        'LB_IDSPressure_LTP
        '
        Me.LB_IDSPressure_LTP.Enabled = False
        Me.LB_IDSPressure_LTP.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.LB_IDSPressure_LTP.ForeColor = System.Drawing.Color.FromArgb(CType(192, Byte), CType(0, Byte), CType(0, Byte))
        Me.LB_IDSPressure_LTP.Location = New System.Drawing.Point(12, 152)
        Me.LB_IDSPressure_LTP.Name = "LB_IDSPressure_LTP"
        Me.LB_IDSPressure_LTP.Size = New System.Drawing.Size(304, 24)
        Me.LB_IDSPressure_LTP.TabIndex = 40
        Me.LB_IDSPressure_LTP.Text = "Cur. Smart Control: (None)"
        Me.LB_IDSPressure_LTP.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.LB_IDSPressure_LTP.Visible = False
        '
        'BT_Pro_LEdit
        '
        Me.BT_Pro_LEdit.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.BT_Pro_LEdit.ForeColor = System.Drawing.SystemColors.ControlText
        Me.BT_Pro_LEdit.Location = New System.Drawing.Point(220, 96)
        Me.BT_Pro_LEdit.Name = "BT_Pro_LEdit"
        Me.BT_Pro_LEdit.TabIndex = 39
        Me.BT_Pro_LEdit.Text = "Allow Edit"
        '
        'BT_Pro_LDownload
        '
        Me.BT_Pro_LDownload.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.BT_Pro_LDownload.ForeColor = System.Drawing.SystemColors.ControlText
        Me.BT_Pro_LDownload.Location = New System.Drawing.Point(118, 96)
        Me.BT_Pro_LDownload.Name = "BT_Pro_LDownload"
        Me.BT_Pro_LDownload.Size = New System.Drawing.Size(88, 32)
        Me.BT_Pro_LDownload.TabIndex = 38
        Me.BT_Pro_LDownload.Text = "Download"
        '
        'BT_Pro_LUpdate
        '
        Me.BT_Pro_LUpdate.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.BT_Pro_LUpdate.ForeColor = System.Drawing.SystemColors.ControlText
        Me.BT_Pro_LUpdate.Location = New System.Drawing.Point(16, 96)
        Me.BT_Pro_LUpdate.Name = "BT_Pro_LUpdate"
        Me.BT_Pro_LUpdate.Size = New System.Drawing.Size(88, 32)
        Me.BT_Pro_LUpdate.TabIndex = 37
        Me.BT_Pro_LUpdate.Text = "Update"
        Me.BT_Pro_LUpdate.Visible = False
        '
        'TB_Pro_LSuck
        '
        Me.TB_Pro_LSuck.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.TB_Pro_LSuck.Location = New System.Drawing.Point(80, 64)
        Me.TB_Pro_LSuck.Name = "TB_Pro_LSuck"
        Me.TB_Pro_LSuck.Size = New System.Drawing.Size(48, 26)
        Me.TB_Pro_LSuck.TabIndex = 30
        Me.TB_Pro_LSuck.Text = "-0.25"
        '
        'TB_Pro_LREV
        '
        Me.TB_Pro_LREV.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.TB_Pro_LREV.Location = New System.Drawing.Point(184, 64)
        Me.TB_Pro_LREV.Name = "TB_Pro_LREV"
        Me.TB_Pro_LREV.Size = New System.Drawing.Size(48, 26)
        Me.TB_Pro_LREV.TabIndex = 32
        Me.TB_Pro_LREV.Text = "-0.1"
        '
        'LB_Pro_LSuck
        '
        Me.LB_Pro_LSuck.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.LB_Pro_LSuck.ForeColor = System.Drawing.SystemColors.ControlText
        Me.LB_Pro_LSuck.Location = New System.Drawing.Point(8, 64)
        Me.LB_Pro_LSuck.Name = "LB_Pro_LSuck"
        Me.LB_Pro_LSuck.Size = New System.Drawing.Size(72, 24)
        Me.LB_Pro_LSuck.TabIndex = 34
        Me.LB_Pro_LSuck.Text = "Suck Bk"
        Me.LB_Pro_LSuck.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'LB_Pro_LRPM
        '
        Me.LB_Pro_LRPM.Enabled = False
        Me.LB_Pro_LRPM.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.LB_Pro_LRPM.ForeColor = System.Drawing.SystemColors.ControlText
        Me.LB_Pro_LRPM.Location = New System.Drawing.Point(144, 32)
        Me.LB_Pro_LRPM.Name = "LB_Pro_LRPM"
        Me.LB_Pro_LRPM.Size = New System.Drawing.Size(48, 24)
        Me.LB_Pro_LRPM.TabIndex = 35
        Me.LB_Pro_LRPM.Text = "RPM"
        Me.LB_Pro_LRPM.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'LB_Pro_LPressure
        '
        Me.LB_Pro_LPressure.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.LB_Pro_LPressure.ForeColor = System.Drawing.SystemColors.ControlText
        Me.LB_Pro_LPressure.Location = New System.Drawing.Point(8, 32)
        Me.LB_Pro_LPressure.Name = "LB_Pro_LPressure"
        Me.LB_Pro_LPressure.Size = New System.Drawing.Size(72, 24)
        Me.LB_Pro_LPressure.TabIndex = 33
        Me.LB_Pro_LPressure.Text = "Pressure"
        Me.LB_Pro_LPressure.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'TB_Pro_LPressure
        '
        Me.TB_Pro_LPressure.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.TB_Pro_LPressure.Location = New System.Drawing.Point(80, 32)
        Me.TB_Pro_LPressure.Name = "TB_Pro_LPressure"
        Me.TB_Pro_LPressure.Size = New System.Drawing.Size(48, 26)
        Me.TB_Pro_LPressure.TabIndex = 29
        Me.TB_Pro_LPressure.Text = "2.00"
        '
        'LB_Pro_LREV
        '
        Me.LB_Pro_LREV.Enabled = False
        Me.LB_Pro_LREV.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.LB_Pro_LREV.ForeColor = System.Drawing.SystemColors.ControlText
        Me.LB_Pro_LREV.Location = New System.Drawing.Point(144, 64)
        Me.LB_Pro_LREV.Name = "LB_Pro_LREV"
        Me.LB_Pro_LREV.Size = New System.Drawing.Size(40, 24)
        Me.LB_Pro_LREV.TabIndex = 36
        Me.LB_Pro_LREV.Text = "REV"
        Me.LB_Pro_LREV.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'TB_Pro_LProg
        '
        Me.TB_Pro_LProg.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.TB_Pro_LProg.Location = New System.Drawing.Point(296, 64)
        Me.TB_Pro_LProg.Name = "TB_Pro_LProg"
        Me.TB_Pro_LProg.Size = New System.Drawing.Size(16, 26)
        Me.TB_Pro_LProg.TabIndex = 32
        Me.TB_Pro_LProg.Text = "4"
        '
        'TB_Pro_LValve
        '
        Me.TB_Pro_LValve.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.TB_Pro_LValve.Location = New System.Drawing.Point(296, 32)
        Me.TB_Pro_LValve.Name = "TB_Pro_LValve"
        Me.TB_Pro_LValve.Size = New System.Drawing.Size(16, 26)
        Me.TB_Pro_LValve.TabIndex = 31
        Me.TB_Pro_LValve.Text = "4"
        '
        'LB_Pro_LProg
        '
        Me.LB_Pro_LProg.Enabled = False
        Me.LB_Pro_LProg.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.LB_Pro_LProg.ForeColor = System.Drawing.SystemColors.ControlText
        Me.LB_Pro_LProg.Location = New System.Drawing.Point(248, 64)
        Me.LB_Pro_LProg.Name = "LB_Pro_LProg"
        Me.LB_Pro_LProg.Size = New System.Drawing.Size(48, 24)
        Me.LB_Pro_LProg.TabIndex = 36
        Me.LB_Pro_LProg.Text = "Prog"
        Me.LB_Pro_LProg.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'LB_Pro_LValve
        '
        Me.LB_Pro_LValve.Enabled = False
        Me.LB_Pro_LValve.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.LB_Pro_LValve.ForeColor = System.Drawing.SystemColors.ControlText
        Me.LB_Pro_LValve.Location = New System.Drawing.Point(248, 32)
        Me.LB_Pro_LValve.Name = "LB_Pro_LValve"
        Me.LB_Pro_LValve.Size = New System.Drawing.Size(48, 24)
        Me.LB_Pro_LValve.TabIndex = 35
        Me.LB_Pro_LValve.Text = "Valve"
        Me.LB_Pro_LValve.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'PanelThermal
        '
        Me.PanelThermal.BackColor = System.Drawing.SystemColors.ScrollBar
        Me.PanelThermal.Controls.Add(Me.PictureBox6)
        Me.PanelThermal.Controls.Add(Me.PictureBox5)
        Me.PanelThermal.Controls.Add(Me.PictureBox4)
        Me.PanelThermal.Controls.Add(Me.PictureBox3)
        Me.PanelThermal.Controls.Add(Me.PictureBox2)
        Me.PanelThermal.Controls.Add(Me.LB_PostHeater)
        Me.PanelThermal.Controls.Add(Me.LB_DispHeater)
        Me.PanelThermal.Controls.Add(Me.LB_PreHeater)
        Me.PanelThermal.Controls.Add(Me.LB_NeedleHeater)
        Me.PanelThermal.Controls.Add(Me.LB_SyrHeater)
        Me.PanelThermal.Controls.Add(Me.BT_HeaterCancel)
        Me.PanelThermal.Controls.Add(Me.BT_HeaterUpdate)
        Me.PanelThermal.Controls.Add(Me.TB_PostHeater)
        Me.PanelThermal.Controls.Add(Me.TB_DispHeater)
        Me.PanelThermal.Controls.Add(Me.TB_PreHeater)
        Me.PanelThermal.Controls.Add(Me.TB_NeedleHeater)
        Me.PanelThermal.Controls.Add(Me.TB_SyrHeater)
        Me.PanelThermal.Controls.Add(Me.Label21)
        Me.PanelThermal.Controls.Add(Me.Label20)
        Me.PanelThermal.Controls.Add(Me.Label19)
        Me.PanelThermal.Controls.Add(Me.Label18)
        Me.PanelThermal.Controls.Add(Me.Label17)
        Me.PanelThermal.Location = New System.Drawing.Point(852, 562)
        Me.PanelThermal.Name = "PanelThermal"
        Me.PanelThermal.Size = New System.Drawing.Size(348, 382)
        Me.PanelThermal.TabIndex = 0
        Me.PanelThermal.Visible = False
        '
        'PictureBox6
        '
        Me.PictureBox6.Image = CType(resources.GetObject("PictureBox6.Image"), System.Drawing.Image)
        Me.PictureBox6.Location = New System.Drawing.Point(229, 144)
        Me.PictureBox6.Name = "PictureBox6"
        Me.PictureBox6.Size = New System.Drawing.Size(112, 32)
        Me.PictureBox6.TabIndex = 21
        Me.PictureBox6.TabStop = False
        '
        'PictureBox5
        '
        Me.PictureBox5.Image = CType(resources.GetObject("PictureBox5.Image"), System.Drawing.Image)
        Me.PictureBox5.Location = New System.Drawing.Point(117, 144)
        Me.PictureBox5.Name = "PictureBox5"
        Me.PictureBox5.Size = New System.Drawing.Size(112, 32)
        Me.PictureBox5.TabIndex = 20
        Me.PictureBox5.TabStop = False
        '
        'PictureBox4
        '
        Me.PictureBox4.Image = CType(resources.GetObject("PictureBox4.Image"), System.Drawing.Image)
        Me.PictureBox4.Location = New System.Drawing.Point(5, 144)
        Me.PictureBox4.Name = "PictureBox4"
        Me.PictureBox4.Size = New System.Drawing.Size(112, 32)
        Me.PictureBox4.TabIndex = 19
        Me.PictureBox4.TabStop = False
        '
        'PictureBox3
        '
        Me.PictureBox3.Image = CType(resources.GetObject("PictureBox3.Image"), System.Drawing.Image)
        Me.PictureBox3.Location = New System.Drawing.Point(146, 32)
        Me.PictureBox3.Name = "PictureBox3"
        Me.PictureBox3.Size = New System.Drawing.Size(40, 16)
        Me.PictureBox3.TabIndex = 18
        Me.PictureBox3.TabStop = False
        '
        'PictureBox2
        '
        Me.PictureBox2.Image = CType(resources.GetObject("PictureBox2.Image"), System.Drawing.Image)
        Me.PictureBox2.Location = New System.Drawing.Point(146, 80)
        Me.PictureBox2.Name = "PictureBox2"
        Me.PictureBox2.Size = New System.Drawing.Size(40, 16)
        Me.PictureBox2.TabIndex = 17
        Me.PictureBox2.TabStop = False
        '
        'LB_PostHeater
        '
        Me.LB_PostHeater.ForeColor = System.Drawing.Color.Red
        Me.LB_PostHeater.Location = New System.Drawing.Point(253, 248)
        Me.LB_PostHeater.Name = "LB_PostHeater"
        Me.LB_PostHeater.Size = New System.Drawing.Size(56, 27)
        Me.LB_PostHeater.TabIndex = 16
        Me.LB_PostHeater.Text = "50 oC"
        Me.LB_PostHeater.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'LB_DispHeater
        '
        Me.LB_DispHeater.ForeColor = System.Drawing.Color.Red
        Me.LB_DispHeater.Location = New System.Drawing.Point(141, 248)
        Me.LB_DispHeater.Name = "LB_DispHeater"
        Me.LB_DispHeater.Size = New System.Drawing.Size(56, 27)
        Me.LB_DispHeater.TabIndex = 15
        Me.LB_DispHeater.Text = "50 oC"
        Me.LB_DispHeater.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'LB_PreHeater
        '
        Me.LB_PreHeater.ForeColor = System.Drawing.Color.Red
        Me.LB_PreHeater.Location = New System.Drawing.Point(29, 248)
        Me.LB_PreHeater.Name = "LB_PreHeater"
        Me.LB_PreHeater.Size = New System.Drawing.Size(56, 27)
        Me.LB_PreHeater.TabIndex = 14
        Me.LB_PreHeater.Text = "50 oC"
        Me.LB_PreHeater.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'LB_NeedleHeater
        '
        Me.LB_NeedleHeater.ForeColor = System.Drawing.Color.Red
        Me.LB_NeedleHeater.Location = New System.Drawing.Point(282, 72)
        Me.LB_NeedleHeater.Name = "LB_NeedleHeater"
        Me.LB_NeedleHeater.Size = New System.Drawing.Size(56, 27)
        Me.LB_NeedleHeater.TabIndex = 13
        Me.LB_NeedleHeater.Text = "50 oC"
        Me.LB_NeedleHeater.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.LB_NeedleHeater.Visible = False
        '
        'LB_SyrHeater
        '
        Me.LB_SyrHeater.ForeColor = System.Drawing.Color.Red
        Me.LB_SyrHeater.Location = New System.Drawing.Point(282, 24)
        Me.LB_SyrHeater.Name = "LB_SyrHeater"
        Me.LB_SyrHeater.Size = New System.Drawing.Size(56, 27)
        Me.LB_SyrHeater.TabIndex = 12
        Me.LB_SyrHeater.Text = "50 oC"
        Me.LB_SyrHeater.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'BT_HeaterCancel
        '
        Me.BT_HeaterCancel.Location = New System.Drawing.Point(210, 304)
        Me.BT_HeaterCancel.Name = "BT_HeaterCancel"
        Me.BT_HeaterCancel.Size = New System.Drawing.Size(75, 48)
        Me.BT_HeaterCancel.TabIndex = 11
        Me.BT_HeaterCancel.Text = "Cancel"
        '
        'BT_HeaterUpdate
        '
        Me.BT_HeaterUpdate.Location = New System.Drawing.Point(66, 304)
        Me.BT_HeaterUpdate.Name = "BT_HeaterUpdate"
        Me.BT_HeaterUpdate.Size = New System.Drawing.Size(75, 48)
        Me.BT_HeaterUpdate.TabIndex = 10
        Me.BT_HeaterUpdate.Text = "Update"
        '
        'TB_PostHeater
        '
        Me.TB_PostHeater.Location = New System.Drawing.Point(253, 216)
        Me.TB_PostHeater.Name = "TB_PostHeater"
        Me.TB_PostHeater.Size = New System.Drawing.Size(56, 27)
        Me.TB_PostHeater.TabIndex = 9
        Me.TB_PostHeater.Text = "60"
        '
        'TB_DispHeater
        '
        Me.TB_DispHeater.Location = New System.Drawing.Point(141, 216)
        Me.TB_DispHeater.Name = "TB_DispHeater"
        Me.TB_DispHeater.Size = New System.Drawing.Size(56, 27)
        Me.TB_DispHeater.TabIndex = 8
        Me.TB_DispHeater.Text = "60"
        '
        'TB_PreHeater
        '
        Me.TB_PreHeater.Location = New System.Drawing.Point(29, 216)
        Me.TB_PreHeater.Name = "TB_PreHeater"
        Me.TB_PreHeater.Size = New System.Drawing.Size(56, 27)
        Me.TB_PreHeater.TabIndex = 7
        Me.TB_PreHeater.Text = "60"
        '
        'TB_NeedleHeater
        '
        Me.TB_NeedleHeater.Enabled = False
        Me.TB_NeedleHeater.Location = New System.Drawing.Point(210, 72)
        Me.TB_NeedleHeater.Name = "TB_NeedleHeater"
        Me.TB_NeedleHeater.Size = New System.Drawing.Size(56, 27)
        Me.TB_NeedleHeater.TabIndex = 6
        Me.TB_NeedleHeater.Text = "60"
        '
        'TB_SyrHeater
        '
        Me.TB_SyrHeater.Location = New System.Drawing.Point(210, 24)
        Me.TB_SyrHeater.Name = "TB_SyrHeater"
        Me.TB_SyrHeater.Size = New System.Drawing.Size(56, 27)
        Me.TB_SyrHeater.TabIndex = 5
        Me.TB_SyrHeater.Text = "60"
        '
        'Label21
        '
        Me.Label21.Location = New System.Drawing.Point(229, 184)
        Me.Label21.Name = "Label21"
        Me.Label21.Size = New System.Drawing.Size(112, 23)
        Me.Label21.TabIndex = 4
        Me.Label21.Text = "Post Heater"
        Me.Label21.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label20
        '
        Me.Label20.Location = New System.Drawing.Point(117, 184)
        Me.Label20.Name = "Label20"
        Me.Label20.Size = New System.Drawing.Size(112, 23)
        Me.Label20.TabIndex = 3
        Me.Label20.Text = "Disp. Heater"
        Me.Label20.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label19
        '
        Me.Label19.Location = New System.Drawing.Point(5, 184)
        Me.Label19.Name = "Label19"
        Me.Label19.Size = New System.Drawing.Size(112, 23)
        Me.Label19.TabIndex = 2
        Me.Label19.Text = "Pre Heater"
        Me.Label19.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label18
        '
        Me.Label18.Enabled = False
        Me.Label18.Location = New System.Drawing.Point(10, 72)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(128, 23)
        Me.Label18.TabIndex = 1
        Me.Label18.Text = "Needle Heater"
        Me.Label18.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label17
        '
        Me.Label17.Location = New System.Drawing.Point(10, 24)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(128, 23)
        Me.Label17.TabIndex = 0
        Me.Label17.Text = "Syringe Heater"
        Me.Label17.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'PanelConveyor
        '
        Me.PanelConveyor.BackColor = System.Drawing.SystemColors.ScrollBar
        Me.PanelConveyor.Controls.Add(Me.BT_CV_Release)
        Me.PanelConveyor.Controls.Add(Me.BT_CV_Retrieve)
        Me.PanelConveyor.Controls.Add(Me.PictureBox9)
        Me.PanelConveyor.Controls.Add(Me.PictureBox8)
        Me.PanelConveyor.Controls.Add(Me.PictureBox7)
        Me.PanelConveyor.Location = New System.Drawing.Point(852, 562)
        Me.PanelConveyor.Name = "PanelConveyor"
        Me.PanelConveyor.Size = New System.Drawing.Size(348, 382)
        Me.PanelConveyor.TabIndex = 0
        Me.PanelConveyor.Visible = False
        '
        'BT_CV_Release
        '
        Me.BT_CV_Release.Image = CType(resources.GetObject("BT_CV_Release.Image"), System.Drawing.Image)
        Me.BT_CV_Release.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.BT_CV_Release.Location = New System.Drawing.Point(210, 200)
        Me.BT_CV_Release.Name = "BT_CV_Release"
        Me.BT_CV_Release.Size = New System.Drawing.Size(88, 56)
        Me.BT_CV_Release.TabIndex = 26
        Me.BT_CV_Release.Text = "Release"
        Me.BT_CV_Release.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'BT_CV_Retrieve
        '
        Me.BT_CV_Retrieve.Image = CType(resources.GetObject("BT_CV_Retrieve.Image"), System.Drawing.Image)
        Me.BT_CV_Retrieve.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.BT_CV_Retrieve.Location = New System.Drawing.Point(66, 200)
        Me.BT_CV_Retrieve.Name = "BT_CV_Retrieve"
        Me.BT_CV_Retrieve.Size = New System.Drawing.Size(88, 56)
        Me.BT_CV_Retrieve.TabIndex = 25
        Me.BT_CV_Retrieve.Text = "Retrieve"
        Me.BT_CV_Retrieve.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'PictureBox9
        '
        Me.PictureBox9.Image = CType(resources.GetObject("PictureBox9.Image"), System.Drawing.Image)
        Me.PictureBox9.Location = New System.Drawing.Point(230, 80)
        Me.PictureBox9.Name = "PictureBox9"
        Me.PictureBox9.Size = New System.Drawing.Size(112, 64)
        Me.PictureBox9.TabIndex = 24
        Me.PictureBox9.TabStop = False
        '
        'PictureBox8
        '
        Me.PictureBox8.Image = CType(resources.GetObject("PictureBox8.Image"), System.Drawing.Image)
        Me.PictureBox8.Location = New System.Drawing.Point(118, 80)
        Me.PictureBox8.Name = "PictureBox8"
        Me.PictureBox8.Size = New System.Drawing.Size(112, 64)
        Me.PictureBox8.TabIndex = 23
        Me.PictureBox8.TabStop = False
        '
        'PictureBox7
        '
        Me.PictureBox7.Image = CType(resources.GetObject("PictureBox7.Image"), System.Drawing.Image)
        Me.PictureBox7.Location = New System.Drawing.Point(6, 80)
        Me.PictureBox7.Name = "PictureBox7"
        Me.PictureBox7.Size = New System.Drawing.Size(112, 64)
        Me.PictureBox7.TabIndex = 22
        Me.PictureBox7.TabStop = False
        '
        'Reset
        '
        Me.Reset.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Reset.Location = New System.Drawing.Point(424, 8)
        Me.Reset.Name = "Reset"
        Me.Reset.Size = New System.Drawing.Size(56, 23)
        Me.Reset.TabIndex = 121
        Me.Reset.Text = "Reset"
        Me.Reset.Visible = False
        '
        'Download
        '
        Me.Download.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Download.Location = New System.Drawing.Point(488, 8)
        Me.Download.Name = "Download"
        Me.Download.Size = New System.Drawing.Size(80, 23)
        Me.Download.TabIndex = 122
        Me.Download.Text = "Download"
        Me.Download.Visible = False
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
        'Button1
        '
        Me.Button1.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Button1.Location = New System.Drawing.Point(576, 8)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(80, 23)
        Me.Button1.TabIndex = 123
        Me.Button1.Text = "OnlineRun"
        Me.Button1.Visible = False
        '
        'NumericUpDownSysSpeed
        '
        Me.NumericUpDownSysSpeed.Enabled = False
        Me.NumericUpDownSysSpeed.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.NumericUpDownSysSpeed.Increment = New Decimal(New Integer() {100, 0, 0, 0})
        Me.NumericUpDownSysSpeed.Location = New System.Drawing.Point(360, 8)
        Me.NumericUpDownSysSpeed.Maximum = New Decimal(New Integer() {800, 0, 0, 0})
        Me.NumericUpDownSysSpeed.Minimum = New Decimal(New Integer() {100, 0, 0, 0})
        Me.NumericUpDownSysSpeed.Name = "NumericUpDownSysSpeed"
        Me.NumericUpDownSysSpeed.Size = New System.Drawing.Size(56, 23)
        Me.NumericUpDownSysSpeed.TabIndex = 124
        Me.NumericUpDownSysSpeed.Value = New Decimal(New Integer() {100, 0, 0, 0})
        Me.NumericUpDownSysSpeed.Visible = False
        '
        'Button2
        '
        Me.Button2.BackColor = System.Drawing.SystemColors.Control
        Me.Button2.Enabled = False
        Me.Button2.Location = New System.Drawing.Point(852, 341)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(148, 21)
        Me.Button2.TabIndex = 102
        '
        'Button3
        '
        Me.Button3.BackColor = System.Drawing.SystemColors.Control
        Me.Button3.Enabled = False
        Me.Button3.Location = New System.Drawing.Point(852, 541)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(148, 21)
        Me.Button3.TabIndex = 103
        '
        'ToolBarButton3
        '
        Me.ToolBarButton3.ImageIndex = 1
        Me.ToolBarButton3.ToolTipText = "Line"
        '
        'ToolBarButton4
        '
        Me.ToolBarButton4.ImageIndex = 2
        Me.ToolBarButton4.ToolTipText = "Arc"
        '
        'ToolBarLineArc
        '
        Me.ToolBarLineArc.BackColor = System.Drawing.Color.FromArgb(CType(192, Byte), CType(255, Byte), CType(255, Byte))
        Me.ToolBarLineArc.Buttons.AddRange(New System.Windows.Forms.ToolBarButton() {Me.ToolBarButton3, Me.ToolBarButton4})
        Me.ToolBarLineArc.ButtonSize = New System.Drawing.Size(42, 42)
        Me.ToolBarLineArc.Dock = System.Windows.Forms.DockStyle.None
        Me.ToolBarLineArc.DropDownArrows = True
        Me.ToolBarLineArc.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.ToolBarLineArc.ImageList = Me.imageListElement
        Me.ToolBarLineArc.Location = New System.Drawing.Point(700, 295)
        Me.ToolBarLineArc.Name = "ToolBarLineArc"
        Me.ToolBarLineArc.ShowToolTips = True
        Me.ToolBarLineArc.Size = New System.Drawing.Size(84, 48)
        Me.ToolBarLineArc.TabIndex = 107
        Me.ToolBarLineArc.Visible = False
        '
        'LB_MM
        '
        Me.LB_MM.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LB_MM.Location = New System.Drawing.Point(825, 7)
        Me.LB_MM.Name = "LB_MM"
        Me.LB_MM.Size = New System.Drawing.Size(152, 25)
        Me.LB_MM.TabIndex = 125
        Me.LB_MM.Text = "Current Unit: mm"
        '
        'LB_INCH
        '
        Me.LB_INCH.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LB_INCH.Location = New System.Drawing.Point(825, 7)
        Me.LB_INCH.Name = "LB_INCH"
        Me.LB_INCH.Size = New System.Drawing.Size(152, 25)
        Me.LB_INCH.TabIndex = 126
        Me.LB_INCH.Text = "Current Unit: inch"
        '
        'Button4
        '
        Me.Button4.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Button4.ForeColor = System.Drawing.Color.Red
        Me.Button4.Location = New System.Drawing.Point(664, 8)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(80, 23)
        Me.Button4.TabIndex = 127
        Me.Button4.Text = "Production"
        Me.Button4.Visible = False
        '
        'AxSpreadsheetProgramming
        '
        Me.AxSpreadsheetProgramming.DataSource = Nothing
        Me.AxSpreadsheetProgramming.Enabled = True
        Me.AxSpreadsheetProgramming.Location = New System.Drawing.Point(8, 40)
        Me.AxSpreadsheetProgramming.Name = "AxSpreadsheetProgramming"
        Me.AxSpreadsheetProgramming.OcxState = CType(resources.GetObject("AxSpreadsheetProgramming.OcxState"), System.Windows.Forms.AxHost.State)
        Me.AxSpreadsheetProgramming.Size = New System.Drawing.Size(1256, 256)
        Me.AxSpreadsheetProgramming.TabIndex = 128
        '
        'TimerForUpdate
        '
        Me.TimerForUpdate.Interval = 1000
        '
        'PictureBoxEstopRed
        '
        Me.PictureBoxEstopRed.Image = CType(resources.GetObject("PictureBoxEstopRed.Image"), System.Drawing.Image)
        Me.PictureBoxEstopRed.Location = New System.Drawing.Point(1032, 8)
        Me.PictureBoxEstopRed.Name = "PictureBoxEstopRed"
        Me.PictureBoxEstopRed.Size = New System.Drawing.Size(16, 16)
        Me.PictureBoxEstopRed.TabIndex = 129
        Me.PictureBoxEstopRed.TabStop = False
        Me.PictureBoxEstopRed.Visible = False
        '
        'Button5
        '
        Me.Button5.Location = New System.Drawing.Point(752, 8)
        Me.Button5.Name = "Button5"
        Me.Button5.Size = New System.Drawing.Size(80, 24)
        Me.Button5.TabIndex = 130
        Me.Button5.Text = "Button5"
        Me.Button5.Visible = False
        '
        'FormProgramming
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(8, 20)
        Me.ClientSize = New System.Drawing.Size(1277, 940)
        Me.Controls.Add(Me.Button5)
        Me.Controls.Add(Me.PictureBoxEstopRed)
        Me.Controls.Add(Me.AxSpreadsheetProgramming)
        Me.Controls.Add(Me.Button4)
        Me.Controls.Add(Me.PanelVision)
        Me.Controls.Add(Me.PanelMotion)
        Me.Controls.Add(Me.NumericUpDownSysSpeed)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.PanelVisionCtrl)
        Me.Controls.Add(Me.Download)
        Me.Controls.Add(Me.Reset)
        Me.Controls.Add(Me.CBDoorLock)
        Me.Controls.Add(Me.CBExpandSpreadsheet)
        Me.Controls.Add(Me.RBBasicSetup)
        Me.Controls.Add(Me.RBProgramEditor)
        Me.Controls.Add(Me.TBMultiDisplayField)
        Me.Controls.Add(Me.TBrefer)
        Me.Controls.Add(Me.TBElements)
        Me.Controls.Add(Me.TBYesNo)
        Me.Controls.Add(Me.TBEditOnly)
        Me.Controls.Add(Me.TBNextEdit)
        Me.Controls.Add(Me.ToolBarLineArc)
        Me.Controls.Add(Me.BTNeedleCal)
        Me.Controls.Add(Me.BTVolCal)
        Me.Controls.Add(Me.BTHome)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.LabelMessege)
        Me.Controls.Add(Me.Panel4)
        Me.Controls.Add(Me.LB_MM)
        Me.Controls.Add(Me.BTChgSyringe)
        Me.Controls.Add(Me.BTPurge)
        Me.Controls.Add(Me.BTClean)
        Me.Controls.Add(Me.LB_INCH)
        Me.Controls.Add(Me.PanelThermal)
        Me.Controls.Add(Me.PanelStepper)
        Me.Controls.Add(Me.PanelGraphDisp)
        Me.Controls.Add(Me.PanelDispenser)
        Me.Controls.Add(Me.PanelConveyor)
        Me.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Menu = Me.MainMenuProgramming
        Me.Name = "FormProgramming"
        Me.Text = "Programming"
        Me.PanelVisionCtrl.ResumeLayout(False)
        CType(Me.NumericUpDown_Brightness, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel4.ResumeLayout(False)
        Me.PanelMotion.ResumeLayout(False)
        Me.PanelStepper.ResumeLayout(False)
        Me.PanelDispenser.ResumeLayout(False)
        Me.GB_RTP.ResumeLayout(False)
        Me.GB_RJetter.ResumeLayout(False)
        Me.GB_LJetter.ResumeLayout(False)
        Me.GB_LTP.ResumeLayout(False)
        Me.PanelThermal.ResumeLayout(False)
        Me.PanelConveyor.ResumeLayout(False)
        CType(Me.NumericUpDownSysSpeed, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.AxSpreadsheetProgramming, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

#Region "Winform start"
    Private Declare Sub Sleep Lib "kernel32" Alias "Sleep" (ByVal dwMilliseconds As Long)  'delay function

    Shared i As Integer = 0
    Dim j As Integer = 0
    Public ThreadExecution As Threading.Thread 'lsgoh
    Private ThreadMonitor As Threading.Thread
    Private threadMain As Threading.Thread
    Private gPatternFileName As String = ""
    Private iSubSheetName As String = ""
    Private SheetRowSelected As Boolean = False
    Private SheetColumnSelected As Boolean = False

    Private CopiedSheetName As String = ""
    Public ErrorSubSheet As SubSheetErrorStruct


    Private Sub ErrorSubSheetStructIni(ByVal iMinItem, ByVal iMaxItem)
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


    'Private modbusClient As New SocketsClient
    ''Private Sub Form1xx_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    'Dim fw As StreamWriter
    'Private Sub Log(ByVal readinteger As Integer)
    '    fw = New StreamWriter("C:\logging.txt", False)
    '    fw.WriteLine(readinteger)
    '    fw.Close()
    'End Sub

    Private Sub FormProgramming_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        PanelVision.Controls.Add(IDS.Devices.Vision.FrmVision.PanelVision) 'lsgoh
        'IDS.Devices.Vision.FrmVision.PanelVision.Location = New Point(0, 0)
        If IDS.Data.Admin.User.RunApplication = "Programmer" Then
            CBExpandSpreadsheet.Visible = True
            Me.Text = "Programming"
        ElseIf IDS.Data.Admin.User.RunApplication = "Maintenace" Then
            Me.Text = "Maintenance"
            CBExpandSpreadsheet.Visible = False

        ElseIf IDS.Data.Admin.User.RunApplication = "Operator" Then
            Me.Text = "Production"
            CBExpandSpreadsheet.Visible = False
        Else
            CBExpandSpreadsheet.Visible = False
        End If

        If IDS.Data.Hardware.Gantry.SystemUnit = False Then 'mm
            LB_MM.Visible = True
            LB_INCH.Visible = False
        Else
            LB_MM.Visible = False
            LB_INCH.Visible = True
        End If

        Disp_Dispenser_Unit_info()

        'SJ

        InitialPatternParameterSetup()

        Connect_Controller()

        m_Execution.m_Pattern.SubCallSheetInitialization(200)     '200 subsheets maximum without duplicated name

        ErrorSubSheetStructIni(200, 500)    '500 subsheets maximum with duplicated name

        ''''''''''''''''''''''''''''''''''''''''''''''
        ' get default value from the default pat file
        ''''''''''''''''''''''''''''''''''''''''''''''

        IDS.Data.ParameterID.RecordID = "FactoryDefault"
        IDSData.Admin.Folder.FileExtension = "Pat"
        IDSData.Admin.Folder.PatternPath = "C:\IDS\Pattern_Dir"
        IDS.Data.OpenData()
        ''''''''''''''''''''''''''''''''''''''''''''''

        SystemSetupDataRetrieve()
        m_Execution.m_Pattern.m_ErrorChk.GetErrorCheckParameter()

        IDS.__ID.Text = 0  'xu long
        IDS.__IDEXE.Text = 0 'xu long
        IDS.T1.Interval = 250 'xu long
        IDS.T2.Interval = 36000 'xu long
        IDS.IN_PS = True   'xu long
        IDS.Devices.InitializeHardware() 'xu long 'open comm ports, etc.
        IDS.T1.Start() ' xu long 
        DispInforPanel()   'xu long 190707 

        'lsgoh - using different thread
        Dim CP As Process = Process.GetCurrentProcess()
        CP.PriorityClass = ProcessPriorityClass.High
        threadMain = Thread.CurrentThread
        threadMain.Priority = ThreadPriority.BelowNormal

        'Dim T As Threading.Thread = Thread.CurrentThread
        'T.Priority = ThreadPriority.Highest

        ThreadMonitor = New Threading.Thread(AddressOf TheThreadMonitor)
        ThreadMonitor.Priority = Threading.ThreadPriority.BelowNormal
        'ThreadMonitor.Priority = Threading.ThreadPriority.Normal
        ThreadMonitor.Start()

        'Disable part of the menu in File
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

        TBYesNo.Buttons(0).Enabled = False
        TBYesNo.Buttons(1).Enabled = False
        EnableTBNextEdit(False)
        EnableTBRefers(False)
        EnableTBElements(False)

        If IDS.Data.Admin.User.RunApplication = "Operator" Then
            Me.Controls.Add(FrmProduction.PanelProduction)
            FrmProduction.PanelProduction.BringToFront()
            FrmProduction.PanelProduction.Location = New Point(0, 0)
            Me.Menu = FrmProduction.MainMenuProduction
            'PanelVision.Controls.Remove(IDS.Devices.Vision.FrmVision.PanelVision) 'lim
            'FrmProduction.Panel1.Controls.Add(IDS.Devices.Vision.FrmVision.PanelVision) 'lim
            PanelVision.BringToFront() 'lsgoh

        End If

        'SJ
        gExeMode = IDS.Data.Admin.User.RunApplication
        'lim
        FrmGD.SetGD(300, 300, 7, 305)
        PanelGraphDisp.Show()
        PanelGraphDisp.Controls.Add(FrmGD.Panel1)
        FrmGD.Panel1.Location = New Point(0, 0)
        PanelStepper.Hide()
        PanelDispenser.Hide()
        PanelThermal.Hide()
        PanelConveyor.Hide()
        TBMultiDisplayField.Buttons(0).Pushed = True
        TBMultiDisplayField.Buttons(1).Pushed = False
        TBMultiDisplayField.Buttons(2).Pushed = False
        TBMultiDisplayField.Buttons(3).Pushed = False
        TBMultiDisplayField.Buttons(4).Pushed = False

        Me.Location = New Point(0, 0)
        Me.StartPosition = FormStartPosition.Manual
        Me.WindowState = FormWindowState.Maximized

    End Sub

    Private Sub FormProgramming_Closed(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Closed
        CloseDevicesSetup() ' xu long. This function is in - Setup DLL - Module.vb -
        MyPSStationSettings.Timer1.Stop()
        IDS.T1.Stop() 'xu long
        IDS.Devices.Conveyor.FrmConveyor.CloseConveyorPort() 'xu long
        IDS.Devices.Thermal.FrmThermal.close_ThermalCom()   'xu long
        IDS.Devices.Weighting.FrmWeight.CloseWeighingMchinePort()  'xu long
        PanelVision.Controls.Remove(IDS.Devices.Vision.FrmVision.PanelVision) 'lsgoh
        FormProgrammingClosed_Implementaton()
        If IDS.Data.Admin.User.RunApplication = "Operator" Then 'lim
            FrmProduction.Panel1.Controls.Remove(IDS.Devices.Vision.FrmVision.PanelVision) 'lim
        End If
    End Sub

    Private Sub FormProgrammingClosed_Implementaton()
        'SJ
        Disconnect_Controller()



        'Thread1.Abort() 'lsgoh
        '        MouseTimer.Change(System.Threading.Timeout.Infinite, System.Threading.Timeout.Infinite)
        ThreadMonitor.Abort()
    End Sub

    Private Sub FormProgramming_HandleCreated(ByVal sender As Object, _
        ByVal e As System.EventArgs) Handles MyBase.HandleCreated
        'InitialPatternParameterSetup()
    End Sub

#End Region

#Region "Shen Jian"
    Friend m_Tri As CIDSTrioController = IDS.Devices.Motor  ' New CIDSTrioController '
    'Public m_Execution As New IDSExecution   ' moved to Module1.vb by Xu Long. is it OK?
    Private m_PosUpdate As Boolean = False
    Private m_KeyinOK As Boolean = False
    Private m_row As Integer = 2
    Private m_InLinkRange As Boolean = False
    Private m_cmdCompleted As Integer = 0
    Private m_ArrayRow As Integer = 0
    Private m_column As Integer = 2
    Private m_CtrlKey As Boolean = False
    Private m_CodingMode As Integer = 0
    Private m_PosNo As Integer = 1
    Private m_EditStateFlag As Boolean = False
    Private m_ProgrammingStateFlag As Boolean = False
    Private type As String = ""
    Private isJogON As Boolean = False
    Protected m_CamPos(3) As Double
    Private m_MachinePos(3) As Double
    Private m_ReferPt() As Double = {0.0, 0.0, 0.0}
    Private m_BoardnRefBlkDist As Double = 0.0
    Private m_RunMode As Integer = 0
    Private m_RunState As Integer = 0
    Private m_PrevRunState As Integer = -1
    Private m_CurrentNeedle As String = "VISION"
    Private m_TeachMode As Integer = 0 '0:vision; 1:Left; 2: Right
    Private m_Xlocked As Boolean = False
    Private m_Ylocked As Boolean = False
    Private m_Zlocked As Boolean = False

    Private m_OpBtnClick As Integer = 3  '1:run; 2:pause 3:stop
    Private m_SelectedType As String = ""
    Dim m_TrackBall As New Mouse(Me)
    Dim m_keyBoard As New Keyboard(Me)
    Private MouseTimer As New System.Threading.Timer(AddressOf MouseJogging, Nothing, 0, 200)
    Private GDRefX As Integer = 7 'lim 'to offset the negative value of the robot's position
    Private GDRefY As Integer = 305 'lim 'to offset the negative value of the robot's position


    Private Sub TheThreadMonitor()

        'Dim counter As Int64
        Dim qcRequest As Integer = 0
        Do
            machineStatus()
            m_Tri.m_TriCtrl.GetTable(200, 1, qcRequest)
            If qcRequest = 1 Then
                m_Tri.m_TriCtrl.SetTable(200, 1, 0)

                bswitchCtrl(10)
                'QCcheck()
            End If
            Thread.Sleep(200)
        Loop
    End Sub

    Private Sub TheThreadExecution()
        ' Exit Sub

        ' m_Tri.m_TriCtrl.SetTable(21, 1, 2)
        Dim rtn As Integer = DispensingCtrl()
        If rtn = 0 Then
        Else
            '  m_Tri.m_TriCtrl.SetVr(21, 0)
            m_Tri.m_TriCtrl.SetTable(21, 1, 0)
            If gExeMode <> "Operator" Then
                LabelMessege.Text = "Trio controller execution error"
            Else
                FrmProduction.LabelMessege.Text = "Trio controller execution error"
            End If

        End If
        Dim t As Threading.Thread = Thread.CurrentThread
        t.Abort()
    End Sub

    Public Sub HardHomming()
        If m_Tri.m_TriCtrl.IsOpen(0) Then
            'Dim mode As Integer = 0
            'm_Tri.m_TriCtrl.GetTable(21, 1, mode)
            'If mode <> 0 Then Exit Sub
            m_Tri.m_TriCtrl.Execute("RUN SETDATUM")
        Else

        End If
    End Sub

    Private Sub BTHome_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTHome.Click
        Dim mode As Integer = 0
        m_Tri.m_TriCtrl.GetTable(21, 1, mode)
        If mode <> 0 Then
            LabelMessege.Text = "Can't home when program running!"
            Exit Sub
        End If
        HardHomming()
    End Sub

    Private Sub MenuSystemHome_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles MenuSystemHome.Click
        Dim mode As Integer = 0
        m_Tri.m_TriCtrl.GetTable(21, 1, mode)
        If mode <> 0 Then
            LabelMessege.Text = "Can't home when program running!"
            Exit Sub
        End If
        HardHomming()
    End Sub


    Public Function exePurge(ByVal needle As String) As Boolean
        Dim off(3), pos(3) As Double
        Dim needleIO As Integer
        Dim LSlideOn, RSlideOn As Integer
        LSlideOn = 0
        RSlideOn = 0

        If needle.ToUpper = "LEFT" Then
            off(0) = gLeftNeedleOffs(0)
            off(1) = gLeftNeedleOffs(1)
            off(2) = gLeftNeedleOffs(2) + gToolPurgePos(2)
            needleIO = gIOLeftNeedle
            LSlideOn = 1
            RSlideOn = 0
            'gPurgeDuration = gLeftPurgeDuration
        ElseIf needle.ToUpper = "RIGHT" Then
            off(0) = gRightNeedleOffs(0)
            off(1) = gRightNeedleOffs(1)
            off(2) = gRightNeedleOffs(2) + gToolPurgePos(2)
            needleIO = gIORightNeedle
            LSlideOn = 0
            RSlideOn = 1
            'gPurgeDuration = gRightPurgeDuration
        Else
            Return False
        End If
        pos(0) = gToolPurgePos(0) - off(0)
        pos(1) = gToolPurgePos(1) - off(1)

        m_Tri.m_TriCtrl.SetTable(21, 1, 2)

        Dim z As Double = off(2) 'gToolPurgePos(2)
        Dim cmdStr As String = "SPEED AXIS(2)=" & gSystemSpeed.ToString
        m_Tri.m_TriCtrl.Execute(cmdStr)
        cmdStr = "SPEED AXIS(1)=" & gSystemSpeed.ToString
        m_Tri.m_TriCtrl.Execute(cmdStr)
        cmdStr = "SPEED AXIS(0)=" & gSystemSpeed.ToString
        m_Tri.m_TriCtrl.Execute(cmdStr)
        Dim zz As Double = gSoftHome(2) + gSystemOrigin(2)
        m_Tri.m_TriCtrl.MoveAbs(1, zz, 2)

        m_Tri.SetDIOs(0, gIORNeedleSlide, RSlideOn)
        m_Tri.SetDIOs(0, gIOLNeedleSlide, LSlideOn)

        m_Tri.m_TriCtrl.MoveAbs(2, pos, 0)
        m_Tri.WaitForEndOfMove()
        m_Tri.m_TriCtrl.MoveAbs(1, z, 2)
        m_Tri.WaitForEndOfMove()
        m_Tri.m_TriCtrl.SetTable(21, 1, 0)

        m_Tri.m_TriCtrl.Op(needleIO, 1)

        '????????????????????????????
        'Dim duration As Integer = gPurgeDuration
        'cmdStr = "WA(" + duration.ToString + ")"
        'm_Tri.m_TriCtrl.Execute(cmdStr)
        'm_Tri.m_TriCtrl.Op(needleIO, 0)

        Return True
    End Function

    Public Function exeClean(ByVal needle As String) As Boolean
        Dim off(3), pos(3) As Double
        Dim needleIO As Integer
        Dim LSlideOn, RSlideOn As Integer
        LSlideOn = 0
        RSlideOn = 0
        If needle.ToUpper = "LEFT" Then
            off(0) = gLeftNeedleOffs(0)
            off(1) = gLeftNeedleOffs(1)
            off(2) = gLeftNeedleOffs(2) + gToolCleanPos(2)
            needleIO = gIOLeftNeedle
            LSlideOn = 1
            RSlideOn = 0
            'gCleanDuration = gLeftCleanDuration
        ElseIf needle.ToUpper = "RIGHT" Then
            off(0) = gRightNeedleOffs(0)
            off(1) = gRightNeedleOffs(1)
            off(2) = gRightNeedleOffs(2) + gToolCleanPos(2)
            needleIO = gIORightNeedle
            LSlideOn = 0
            RSlideOn = 1
            'gCleanDuration = gRightCleanDuration
        Else
            Return False
        End If
        pos(0) = gToolCleanPos(0) - off(0)
        pos(1) = gToolCleanPos(1) - off(1)
        m_Tri.m_TriCtrl.SetTable(21, 1, 2)
        Dim z As Double = off(2) 'gToolCleanPos(2)
        Dim cmdStr As String = "SPEED AXIS(2)=" & gSystemSpeed.ToString
        m_Tri.m_TriCtrl.Execute(cmdStr)
        cmdStr = "SPEED AXIS(1)=" & gSystemSpeed.ToString
        m_Tri.m_TriCtrl.Execute(cmdStr)
        cmdStr = "SPEED AXIS(0)=" & gSystemSpeed.ToString
        m_Tri.m_TriCtrl.Execute(cmdStr)
        Dim zz As Double = gSoftHome(2) + gSystemOrigin(2)
        m_Tri.m_TriCtrl.MoveAbs(1, zz, 2)

        m_Tri.SetDIOs(0, gIORNeedleSlide, RSlideOn)
        m_Tri.SetDIOs(0, gIOLNeedleSlide, LSlideOn)

        m_Tri.m_TriCtrl.MoveAbs(2, pos, 0)
        m_Tri.WaitForEndOfMove()
        m_Tri.m_TriCtrl.MoveAbs(1, z, 2)
        m_Tri.WaitForEndOfMove()
        m_Tri.m_TriCtrl.SetTable(21, 1, 0)

        m_Tri.m_TriCtrl.Op(gIOClean, 1)
        'If Not IDS.Devices.DIO.DIO.SetOutputBit(gPCioPort, gIOClean, True) Then Return False
        '????????????????????????????
        'Dim duration As Integer = gCleanDuration
        'cmdStr = "WA(" + duration.ToString + ")"
        'm_Tri.m_TriCtrl.Execute(cmdStr)
        'm_Tri.m_TriCtrl.Op(gIOClean, 0)

        Return True
    End Function

    Private Function exeChangeSyringe(ByVal needle As String) As Boolean
        Dim off(3), pos(3) As Double
        Dim needleIO As Integer

        If needle.ToUpper = "LEFT" Then
            off(0) = gLeftNeedleOffs(0)
            off(1) = gLeftNeedleOffs(1)
            off(2) = gLeftNeedleOffs(2) + gSyringeChangePos(2)
            needleIO = gIOLeftNeedle
        ElseIf needle.ToUpper = "RIGHT" Then
            off(0) = gRightNeedleOffs(0)
            off(1) = gRightNeedleOffs(1)
            off(2) = gRightNeedleOffs(2) + gSyringeChangePos(2)
            needleIO = gIORightNeedle
        Else
            Return False
        End If
        pos(0) = gSyringeChangePos(0) - off(0)
        pos(1) = gSyringeChangePos(1) - off(1)
        m_Tri.m_TriCtrl.SetTable(21, 1, 2)
        Dim z As Double = off(2) 'gSyringeChangePos(2)
        Dim cmdStr As String = "SPEED AXIS(2)=" & gSystemSpeed.ToString
        m_Tri.m_TriCtrl.Execute(cmdStr)
        cmdStr = "SPEED AXIS(1)=" & gSystemSpeed.ToString
        m_Tri.m_TriCtrl.Execute(cmdStr)
        cmdStr = "SPEED AXIS(0)=" & gSystemSpeed.ToString
        m_Tri.m_TriCtrl.Execute(cmdStr)
        Dim zz As Double = gSoftHome(2) + gSystemOrigin(2)
        m_Tri.m_TriCtrl.MoveAbs(1, zz, 2)
        ' m_Tri.WaitNextFree()
        m_Tri.m_TriCtrl.MoveAbs(2, pos, 0)
        m_Tri.WaitForEndOfMove()
        m_Tri.m_TriCtrl.MoveAbs(1, z, 2)
        m_Tri.WaitForEndOfMove()
        m_Tri.m_TriCtrl.SetTable(21, 1, 0)
        Return True
    End Function

    Public Sub ChangeSyringeCtrl(ByVal needle As String)
        Dim mode As Integer = 0
        m_Tri.m_TriCtrl.GetTable(21, 1, mode)
        If mode <> 0 Then
            LabelMessege.Text = "Can't change syringe when program running!"
            Exit Sub
        End If
        If exeChangeSyringe(needle) Then
        Else
            LabelMessege.Text = "Wrong teaching mode"
            LabelMessege.Refresh()
        End If
    End Sub

    Private Sub BTChgSyringe_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTChgSyringe.Click
        ChangeSyringeCtrl(m_CurrentNeedle)
    End Sub

    Private Sub MenuNeedleChng_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles MenuNeedleChng.Click
        ChangeSyringeCtrl(m_CurrentNeedle)
    End Sub

    Private m_PurgeClicked As Boolean = False
    Public Sub PurgeCtrl(ByVal needle As String)
        Dim mode As Integer = 0
        m_Tri.m_TriCtrl.GetTable(21, 1, mode)
        If mode <> 0 Then
            LabelMessege.Text = "Can't purge when program running!"
            Exit Sub
        End If
        m_PurgeClicked = Not m_PurgeClicked
        If m_PurgeClicked Then
            If exePurge(needle) Then
            Else
                m_PurgeClicked = Not m_PurgeClicked
                LabelMessege.Text = "Wrong teaching mode"
                LabelMessege.Refresh()
                Exit Sub
            End If
            BTPurge.Text = "Purge Off"
            MenuPurge.Text = "Purge Off"
        Else
            Dim needleIO As Integer
            If needle.ToUpper = "LEFT" Then
                needleIO = gIOLeftNeedle
            ElseIf needle.ToUpper = "RIGHT" Then
                needleIO = gIORightNeedle
            Else
                m_PurgeClicked = Not m_PurgeClicked
                LabelMessege.Text = "Wrong teaching mode"
                LabelMessege.Refresh()
                Exit Sub
            End If

            m_Tri.m_TriCtrl.Op(needleIO, 0)
            Dim z As Double = gSoftHome(2) + gSystemOrigin(2)
            m_Tri.m_TriCtrl.MoveAbs(1, z, 2)
            m_Tri.WaitForEndOfMove()
            ' m_Tri.m_TriCtrl.SetTable(21, 1, 0)
            BTPurge.Text = "Purge"
            MenuPurge.Text = "Purge"
        End If
    End Sub

    Private Sub BTPurge_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTPurge.Click
        PurgeCtrl(m_CurrentNeedle)
    End Sub

    Private Sub MenuPurge_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles MenuPurge.Click
        PurgeCtrl(m_CurrentNeedle)
    End Sub


    Private m_CleanClicked As Boolean = False
    Public Sub CleanCtrl(ByVal needle As String)
        Dim mode As Integer = 0
        m_Tri.m_TriCtrl.GetTable(21, 1, mode)
        If mode <> 0 Then
            LabelMessege.Text = "Can't clean when program running!"
            Exit Sub
        End If
        m_CleanClicked = Not m_CleanClicked
        If m_CleanClicked Then
            If exeClean(needle) Then
            Else
                m_CleanClicked = Not m_CleanClicked
                LabelMessege.Text = "Wrong teaching mode"
                LabelMessege.Refresh()
                Exit Sub
            End If
            BTClean.Text = "Clean Off"
            MenuClean.Text = "Clean Off"
        Else
            m_Tri.m_TriCtrl.Op(gIOClean, 0)
            Dim z As Double = gSoftHome(2) + gSystemOrigin(2)
            m_Tri.m_TriCtrl.MoveAbs(1, z, 2)
            m_Tri.WaitForEndOfMove()
            ' m_Tri.m_TriCtrl.SetTable(21, 1, 0)

            BTClean.Text = "Clean"
            MenuClean.Text = "Clean"
        End If
    End Sub

    Private Sub BTClean_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTClean.Click
        CleanCtrl(m_CurrentNeedle)
    End Sub

    Private Sub MenuClean_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles MenuClean.Click
        CleanCtrl(m_CurrentNeedle)
    End Sub

    Public Sub NeedleCalibCtrl(ByVal needle As String)
        Dim status As Integer
        m_Tri.m_TriCtrl.GetTable(21, 1, status)
        If status <> 0 Then
            LabelMessege.Text = "Can't do needle calibration when program running!"
            Exit Sub
        End If

        m_Tri.m_TriCtrl.SetTable(21, 1, 2)

        If needle.ToUpper = "LEFT" Then 'Left
            If MyPSNeedleCal5.L_Needle_Cal() Then
                SystemSetupDataRetrieve()
            Else
                LabelMessege.Text = "Needle calibration fail!"
            End If
        ElseIf needle.ToUpper = "RIGHT" Then  'Right
            If MyPSNeedleCal5.R_Needle_Cal() Then
                SystemSetupDataRetrieve()
            Else
                LabelMessege.Text = "Needle calibration fail!"
            End If
        Else
            LabelMessege.Text = "Wrong teaching mode!"
        End If
        m_Tri.m_TriCtrl.SetTable(21, 1, 0)
    End Sub

    Private Sub BTNeedleCal_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNeedleCal.Click
        'm_Tri.m_TriCtrl.SetTable(21, 1, 2)
        NeedleCalibCtrl(m_CurrentNeedle)
        'm_Tri.m_TriCtrl.SetTable(21, 1, 0)
    End Sub

    Private Sub MenuNeedleCalib_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles MenuNeedleCalib.Click
        'm_Tri.m_TriCtrl.SetTable(21, 1, 2)
        NeedleCalibCtrl(m_CurrentNeedle)
        'm_Tri.m_TriCtrl.SetTable(21, 1, 0)
    End Sub

    Public Function MoveToVolumCalPos(ByVal needle As String) As Integer
        Dim off(3), pos(3) As Double
        Dim needleIO As Integer

        If needle.ToUpper = "LEFT" Then
            off(0) = gLeftNeedleOffs(0)
            off(1) = gLeftNeedleOffs(1)
            off(2) = gLeftNeedleOffs(2) + gVolumCalPos(2)
        ElseIf needle.ToUpper = "RIGHT" Then
            off(0) = gRightNeedleOffs(0)
            off(1) = gRightNeedleOffs(1)
            off(2) = gRightNeedleOffs(2) + gVolumCalPos(2)
        Else
            Return -1
        End If
        pos(0) = gVolumCalPos(0) - off(0)
        pos(1) = gVolumCalPos(1) - off(1)

        Dim z As Double = off(2)
        Dim cmdStr As String = "SPEED AXIS(2)=" & gSystemSpeed.ToString
        m_Tri.m_TriCtrl.Execute(cmdStr)
        cmdStr = "SPEED AXIS(1)=" & gSystemSpeed.ToString
        m_Tri.m_TriCtrl.Execute(cmdStr)
        cmdStr = "SPEED AXIS(0)=" & gSystemSpeed.ToString
        m_Tri.m_TriCtrl.Execute(cmdStr)
        Dim zz As Double = gSoftHome(2) + gSystemOrigin(2)
        m_Tri.m_TriCtrl.MoveAbs(1, zz, 2)

        m_Tri.m_TriCtrl.MoveAbs(2, pos, 0)
        m_Tri.WaitForEndOfMove()
        m_Tri.m_TriCtrl.MoveAbs(1, z, 2)
        m_Tri.WaitForEndOfMove()
        Return 0
    End Function

    Public Sub VolumCalibCtrl(ByVal needle As String)
        Dim status As Integer
        m_Tri.m_TriCtrl.GetTable(21, 1, status)
        If status <> 0 Then
            LabelMessege.Text = "Can't do volume calibration when program running!"
            Exit Sub
        End If
        m_Tri.m_TriCtrl.SetTable(21, 1, 2)

        MoveToVolumCalPos(needle)

        If needle.ToUpper = "LEFT" Then  'Left
            If MyPSVolumeCal.L_VolCal() Then
                SystemSetupDataRetrieve()
            Else
                LabelMessege.Text = "Needle calibration fail!"
            End If
        ElseIf needle.ToUpper = "RIGHT" Then  'Right
            If MyPSVolumeCal.R_VolCal() Then
                SystemSetupDataRetrieve()
            Else
                LabelMessege.Text = "Needle calibration fail!"
            End If
        Else
            LabelMessege.Text = "Wrong teaching mode!"
        End If

        Dim zz As Double = gSoftHome(2) + gSystemOrigin(2)
        m_Tri.m_TriCtrl.MoveAbs(1, zz, 2)

        m_Tri.m_TriCtrl.SetTable(21, 1, 0)
    End Sub

    Private Sub BTVolCal_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTVolCal.Click
        VolumCalibCtrl(m_CurrentNeedle)
    End Sub

    Private Sub MenuVolmCalib_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuVolmCalib.Click
        VolumCalibCtrl(m_CurrentNeedle)
    End Sub

    Private Sub RBVisionMode_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RBVisionMode.Click
        Dim status As Integer
        m_Tri.m_TriCtrl.GetTable(21, 1, status)
        If status <> 0 Then
            LabelMessege.Text = "Can't change mode when program running!"
            Exit Sub
        End If

        m_TeachMode = 0
        m_CurrentNeedle = "VISION"
        m_Tri.SetDIOs(0, gIOLNeedleSlide, 0)
        m_Tri.SetDIOs(0, gIORNeedleSlide, 0)
        LabelMessege.Text = ""
        LabelMessege.Refresh()
    End Sub

    Private Sub RBLNeedleMode_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RBLNeedleMode.Click
        Dim status As Integer
        m_Tri.m_TriCtrl.GetTable(21, 1, status)
        If status <> 0 Then
            LabelMessege.Text = "Can't change mode when program running!"
            Exit Sub
        End If

        Dim cmdStr As String = "SPEED AXIS(2)=" & gSystemSpeed.ToString
        m_Tri.m_TriCtrl.Execute(cmdStr)
        cmdStr = "SPEED AXIS(1)=" & gSystemSpeed.ToString
        m_Tri.m_TriCtrl.Execute(cmdStr)
        cmdStr = "SPEED AXIS(0)=" & gSystemSpeed.ToString
        m_Tri.m_TriCtrl.Execute(cmdStr)
        Dim zz As Double = gSoftHome(2) + gSystemOrigin(2)
        m_Tri.m_TriCtrl.MoveAbs(1, zz, 2)
        m_Tri.WaitForEndOfMove()

        m_TeachMode = 1
        m_CurrentNeedle = "LEFT"
        m_Tri.SetDIOs(0, gIOLNeedleSlide, 1)
        m_Tri.SetDIOs(0, gIORNeedleSlide, 0)
        LabelMessege.Text = ""
        LabelMessege.Refresh()
    End Sub

    Private Sub RBRNeedleMode_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RBRNeedleMode.Click
        Dim status As Integer
        m_Tri.m_TriCtrl.GetTable(21, 1, status)
        If status <> 0 Then
            LabelMessege.Text = "Can't change mode when program running!"
            Exit Sub
        End If

        Dim cmdStr As String = "SPEED AXIS(2)=" & gSystemSpeed.ToString
        m_Tri.m_TriCtrl.Execute(cmdStr)
        cmdStr = "SPEED AXIS(1)=" & gSystemSpeed.ToString
        m_Tri.m_TriCtrl.Execute(cmdStr)
        cmdStr = "SPEED AXIS(0)=" & gSystemSpeed.ToString
        m_Tri.m_TriCtrl.Execute(cmdStr)
        Dim zz As Double = gSoftHome(2) + gSystemOrigin(2)
        m_Tri.m_TriCtrl.MoveAbs(1, zz, 2)
        m_Tri.WaitForEndOfMove()

        m_TeachMode = 2
        m_CurrentNeedle = "RIGHT"
        m_Tri.SetDIOs(0, gIORNeedleSlide, 1)
        m_Tri.SetDIOs(0, gIOLNeedleSlide, 0)
        LabelMessege.Text = ""
        LabelMessege.Refresh()
    End Sub

    'SJ add door lock/unlock
    Private Sub CheckBox5_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CBDoorLock.CheckedChanged
        Dim status As Integer
        m_Tri.m_TriCtrl.GetTable(21, 1, status)
        If status <> 0 Then
            LabelMessege.Text = "Can't unlock the door when program running!"
            Exit Sub
        End If
        If CBDoorLock.Checked = True Then
            CBDoorLock.Text = "Unlock Door"
            CBDoorLock.ImageIndex = 6
            MenuDoorLock.Checked = True
            MenuDoorUnLock.Checked = False
            IDS.Devices.DIO.DIO.DoorLock(True)
        Else
            CBDoorLock.Text = "Lock Door"
            CBDoorLock.ImageIndex = 5
            MenuDoorLock.Checked = False
            MenuDoorUnLock.Checked = True
            IDS.Devices.DIO.DIO.DoorLock(False)
        End If
    End Sub

    Private Sub MenuDoorLock_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles MenuDoorLock.Click
        IDS.Devices.DIO.DIO.DoorLock(True)
        CBDoorLock.Text = "Unlock Door"
        CBDoorLock.ImageIndex = 6
        MenuDoorLock.Checked = True
        MenuDoorUnLock.Checked = False
    End Sub

    Private Sub MenuDoorUnLock_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles MenuDoorUnLock.Click
        Dim status As Integer
        m_Tri.m_TriCtrl.GetTable(21, 1, status)
        If status <> 0 Then
            LabelMessege.Text = "Can't unlock the door when program running!"
            Exit Sub
        End If
        IDS.Devices.DIO.DIO.DoorLock(False)
        CBDoorLock.Text = "Lock Door"
        CBDoorLock.ImageIndex = 5
        MenuDoorLock.Checked = False
        MenuDoorUnLock.Checked = True
    End Sub

    Private Sub CheckBoxLockX_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBoxLockX.CheckedChanged
        Dim status As Integer
        m_Tri.m_TriCtrl.GetTable(21, 1, status)
        If status <> 0 Then
            LabelMessege.Text = "Can't lock axis when program running!"
            Exit Sub
        End If
        If CheckBoxLockX.Checked = True Then
            m_Tri.m_TriCtrl.SetTable(10, 1, 1)
            m_Xlocked = True
        Else
            m_Tri.m_TriCtrl.SetTable(10, 1, 0)
            m_Xlocked = False
        End If
    End Sub


    Private Sub CheckBoxLockY_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBoxLockY.CheckedChanged
        Dim status As Integer
        m_Tri.m_TriCtrl.GetTable(21, 1, status)
        If status <> 0 Then
            LabelMessege.Text = "Can't lock axis when program running!"
            Exit Sub
        End If
        If CheckBoxLockY.Checked = True Then
            m_Tri.m_TriCtrl.SetTable(11, 1, 1)
            m_Ylocked = True
        Else
            m_Tri.m_TriCtrl.SetTable(11, 1, 0)
            m_Ylocked = False
        End If
    End Sub

    Private Sub CheckBoxLockZ_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBoxLockZ.CheckedChanged
        Dim status As Integer
        m_Tri.m_TriCtrl.GetTable(21, 1, status)
        If status <> 0 Then
            LabelMessege.Text = "Can't lock axis when program running!"
            Exit Sub
        End If
        If CheckBoxLockZ.Checked = True Then
            m_Tri.m_TriCtrl.SetTable(12, 1, 1)
            m_Zlocked = True
        Else
            m_Tri.m_TriCtrl.SetTable(12, 1, 0)
            m_Zlocked = False
        End If
    End Sub

    Public Function RunPatternProg() As Integer
        Dim mode As Integer = 0
        m_Tri.m_TriCtrl.GetTable(21, 1, mode)
        If mode = 2 Then  'is running
            Exit Function
        ElseIf mode = 1 Then 'currenetly pause
            m_Tri.m_TriCtrl.SetTable(21, 1, 2) 'resume 
        Else  'currently stop
            If gExeMode <> "Operator" Then
                LabelMessege.Text = "Downloading ..."
                LabelMessege.Refresh()
            Else
                FrmProduction.LabelMessege.Text = "Downloading ..."
                FrmProduction.LabelMessege.Refresh()
            End If

            m_Tri.m_TriCtrl.SetTable(21, 1, 2)
            If gExeMode <> "Operator" Then
                TBOperation.Buttons(0).Enabled = False
                TBOperation.Buttons(1).Enabled = True
                TBOperation.Buttons(2).Enabled = True
            Else
                ' FrmProduction.TBOperation.Buttons(0).Enabled = False
                'FrmProduction.TBOperation.Buttons(1).Enabled = True
                'FrmProduction.TBOperation.Buttons(2).Enabled = True
            End If

            m_Execution.m_Command.SetOptimizFlag(0)
            Dim rtn As Integer = m_Execution.m_Command.ReadPattern(AxSpreadsheetProgramming)
            If rtn < 0 Then
                m_Tri.m_TriCtrl.SetTable(21, 1, 0)
                If gExeMode <> "Operator" Then
                    LabelMessege.Text = "Empty sheet or Error"
                    LabelMessege.Refresh()
                    TBOperation.Buttons(0).Enabled = True
                    TBOperation.Buttons(1).Enabled = False
                    TBOperation.Buttons(2).Enabled = False
                Else
                    FrmProduction.LabelMessege.Text = "Empty sheet or Error"
                    FrmProduction.LabelMessege.Refresh()
                    '  FrmProduction.TBOperation.Buttons(0).Enabled = True
                    ' FrmProduction.TBOperation.Buttons(1).Enabled = False
                    'FrmProduction.TBOperation.Buttons(2).Enabled = False
                End If

                'TBOperation.Enabled = True
                Return rtn
            ElseIf rtn > 0 Then
                m_Tri.m_TriCtrl.SetTable(21, 1, 0)
                If gExeMode <> "Operator" Then
                    LabelMessege.Text = "Execution stopped"
                    LabelMessege.Refresh()
                    TBOperation.Buttons(0).Enabled = True
                    TBOperation.Buttons(1).Enabled = False
                    TBOperation.Buttons(2).Enabled = False
                Else
                    FrmProduction.LabelMessege.Text = "Execution stopped"
                    FrmProduction.LabelMessege.Refresh()
                    '  FrmProduction.TBOperation.Buttons(0).Enabled = True
                    ' FrmProduction.TBOperation.Buttons(1).Enabled = False
                    'FrmProduction.TBOperation.Buttons(2).Enabled = False
                End If

                'TBOperation.Enabled = True
                Return rtn
            End If
            If gExeMode <> "Operator" Then
                LabelMessege.Text = "Compiling..."
                LabelMessege.Update()
            Else
                FrmProduction.LabelMessege.Text = "Compiling..."
                FrmProduction.LabelMessege.Update()
            End If
            ''''SJ test
            Application.DoEvents()
            If m_Execution.m_Command.Compile(m_RunMode) < 0 Then
                m_Tri.m_TriCtrl.SetTable(21, 1, 0)
                If gExeMode <> "Operator" Then
                    LabelMessege.Text = "Empty sheet or Error"
                    LabelMessege.Refresh()
                    TBOperation.Buttons(0).Enabled = True
                    TBOperation.Buttons(1).Enabled = False
                    TBOperation.Buttons(2).Enabled = False
                Else
                    FrmProduction.LabelMessege.Text = "Empty sheet or Error"
                    FrmProduction.LabelMessege.Refresh()
                    ' FrmProduction.TBOperation.Buttons(0).Enabled = True
                    'FrmProduction.TBOperation.Buttons(1).Enabled = False
                    'FrmProduction.TBOperation.Buttons(2).Enabled = False
                End If

                ' TBOperation.Enabled = True
                Return -1
            End If

            ThreadExecution = New Threading.Thread(AddressOf TheThreadExecution)
            ThreadExecution.Priority = Threading.ThreadPriority.Highest
            ThreadExecution.Start()

        End If
        Return 0
    End Function

    'run program
    Private Sub TBOperation_ButtonClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ToolBarButtonClickEventArgs) Handles TBOperation.ButtonClick
        Try

            If e.Button Is TBOperation.Buttons(0) Then 'run
                If Not IDS.Devices.DIO.DIO.DoorStatus() Then
                    LabelMessege.Text = "Please close the door first!"
                    Exit Sub
                End If
                m_OpBtnClick = 1
                RunPatternProg()
            ElseIf e.Button Is TBOperation.Buttons(1) Then 'pause
                m_Tri.m_TriCtrl.SetTable(21, 1, 1)
                m_OpBtnClick = 2
                TBOperation.Buttons(0).Enabled = True
                TBOperation.Buttons(2).Enabled = True
                TBOperation.Buttons(1).Enabled = False
            ElseIf e.Button Is TBOperation.Buttons(2) Then 'stop
                m_OpBtnClick = 3
                TBOperation.Buttons(0).Enabled = True
                TBOperation.Buttons(1).Enabled = False
                TBOperation.Buttons(2).Enabled = False
                m_Tri.m_TriCtrl.SetTable(21, 1, 0)
                m_Tri.m_TriCtrl.RapidStop()
                m_Tri.m_TriCtrl.Stop("Dispense")
                m_Tri.m_TriCtrl.Cancel(1, 0)
                m_Tri.m_TriCtrl.Cancel(1, 1)
                m_Tri.m_TriCtrl.Cancel(1, 2)
                m_Tri.m_TriCtrl.Cancel(0, 0)
                m_Tri.m_TriCtrl.Cancel(0, 1)
                m_Tri.m_TriCtrl.Cancel(0, 2)
                m_Tri.m_TriCtrl.Op(0)
                m_Tri.m_TriCtrl.Op(gIOZbreak, 1)
                'm_Tri.m_TriCtrl.Op(9, 0)
                'm_Tri.m_TriCtrl.Op(10, 0)
                'm_Tri.m_TriCtrl.Op(12, 0)
                'm_Tri.m_TriCtrl.Op(13, 0)
                'm_Tri.m_TriCtrl.Op(14, 0)
                'm_Tri.m_TriCtrl.Op(15, 0)
                If ThreadExecution Is Nothing Then
                Else
                    ThreadExecution.Abort()
                End If

            End If

        Catch ex As SystemException
            MessageBox.Show(ex.ToString)
        End Try
    End Sub

    Private Sub Trimotion_MoveTo(ByVal Pos() As Double)

        Dim spstr As String = gSystemSpeed.ToString
        Dim str As String
        str = "SPEED AXIS(0) =" + spstr
        m_Tri.m_TriCtrl.Execute(str)
        str = "SPEED AXIS(1) =" + spstr
        m_Tri.m_TriCtrl.Execute(str)
        str = "SPEED AXIS(2) =" + spstr
        m_Tri.m_TriCtrl.Execute(str)
        Dim z As Double = gSoftHome(2) + gSystemOrigin(2)   'park position wrf hardhome
        Dim ref(3), inmmPos(3) As Double
        ref(0) = m_ReferPt(0) * gmminchRatio
        ref(1) = m_ReferPt(1) * gmminchRatio
        ref(2) = m_ReferPt(2) * gmminchRatio
        ReferToSys(Pos, Pos, ref)
        SysToHard(Pos, Pos)
        m_Tri.m_TriCtrl.MoveAbs(1, z, 2)
        'LabelMessege.Text = "edit finished"
        m_Tri.WaitNextFree()

        m_Tri.m_TriCtrl.MoveAbs(2, Pos, 0)
        m_Tri.WaitForEndOfMove()
    End Sub

    'sj add
    Private Sub Trimotion_MoveTo(ByVal Pos() As Double, ByVal teachMode As Integer)

        Dim spstr As String = gSystemSpeed.ToString
        Dim str As String
        str = "SPEED AXIS(0) =" + spstr
        m_Tri.m_TriCtrl.Execute(str)
        str = "SPEED AXIS(1) =" + spstr
        m_Tri.m_TriCtrl.Execute(str)
        str = "SPEED AXIS(2) =" + spstr
        m_Tri.m_TriCtrl.Execute(str)
        Dim z As Double = gSoftHome(2) + gSystemOrigin(2)   'park position wrf hardhome
        Dim ref(3), inmmPos(3) As Double
        ref(0) = m_ReferPt(0) * gmminchRatio
        ref(1) = m_ReferPt(1) * gmminchRatio
        ref(2) = m_ReferPt(2) * gmminchRatio
        ReferToSys(Pos, Pos, ref)
        SysToHard(Pos, Pos)
        If teachMode = 1 Then  'left
            Pos(0) = Pos(0) - gLeftNeedleOffs(0)
            Pos(1) = Pos(1) - gLeftNeedleOffs(1)
        ElseIf teachMode = 2 Then 'right
            Pos(0) = Pos(0) - gRightNeedleOffs(0)
            Pos(1) = Pos(1) - gRightNeedleOffs(1)
        End If

        m_Tri.m_TriCtrl.MoveAbs(1, z, 2)
        'LabelMessege.Text = "edit finished"
        m_Tri.WaitNextFree()

        m_Tri.m_TriCtrl.MoveAbs(2, Pos, 0)
        m_Tri.WaitForEndOfMove()
    End Sub


    Private Sub Trimotion_MoveTo(ByVal Pos() As Double, ByVal DelayFlag As Boolean)

        Dim spstr As String = gSystemSpeed.ToString
        Dim str As String
        str = "SPEED AXIS(0) =" + spstr
        m_Tri.m_TriCtrl.Execute(str)
        str = "SPEED AXIS(1) =" + spstr
        m_Tri.m_TriCtrl.Execute(str)
        str = "SPEED AXIS(2) =" + spstr
        m_Tri.m_TriCtrl.Execute(str)
        Dim z As Double = gSoftHome(2) + gSystemOrigin(2)   'park position wrf hardhome
        Dim ref(3), inmmPos(3) As Double
        ref(0) = m_ReferPt(0) * gmminchRatio
        ref(1) = m_ReferPt(1) * gmminchRatio
        ref(2) = m_ReferPt(2) * gmminchRatio
        ReferToSys(Pos, Pos, ref)
        SysToHard(Pos, Pos)
        m_Tri.m_TriCtrl.MoveAbs(1, z, 2)
        'LabelMessege.Text = "edit finished"
        If DelayFlag Then
            m_Tri.WaitNextFree()
        End If
        m_Tri.m_TriCtrl.MoveAbs(2, Pos, 0)
        m_Tri.WaitForEndOfMove()
    End Sub


    Shared mouse_pos As New Point
    Shared cursor_hide As Boolean = False
    Shared isPress As Boolean
    Dim deadzone As Integer = 3
    Dim jogspeed As Integer = 0
    Const maxSpeed As Integer = 100
    Const maxMouseRangeX = 6000.0 '4500.0
    Const maxMouseRangeY = 2000.0
    Const ratioLB = 0.4
    Const ratioUB = 1.05
    Dim ratio As Double


    Private Sub MouseJogging(ByVal state As Object)
        Application.DoEvents()

        If m_RunState <> 0 Or IDS.Data.SystemAtLogin = True Or m_OpBtnClick <> 3 Then Exit Sub

        m_keyBoard.Poll()
        m_TrackBall.Poll()

        isPress = m_keyBoard.State.Item(Key.LeftControl)
        Dim x As Integer
        Dim y As Integer
        ' Dim CmdStr As String
        Dim VrData(3) As Single

        If isPress Then

            VrData(0) = 0
            VrData(1) = 0.0
            VrData(2) = 0.0
            'If Me.cursor_hide = False Then

            '    Me.Cursor.Hide()
            '    '  mouse_pos = Me.Cursor.Position
            '    cursor_hide = True
            'Else
            '    ' Me.Cursor.Position = mouse_pos
            'End If
            Dim isPressAlt As Boolean = m_keyBoard.State.Item(Key.LeftAlt)
            If isPressAlt Then
                Exit Sub
            End If
            x = m_TrackBall.MouseX()
            y = m_TrackBall.MouseY()
            'test
            'm_CamPos(0) = x
            'm_CamPos(1) = y
            'm_CamPos(2) = 0.0
            '''
            Dim ratio As Double
            'm_Tri.m_TriCtrl.Execute("SRAMP=6")

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

                        m_TriProgMemHandle.SetTable(5, 3, VrData(0))
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

                        m_TriProgMemHandle.SetTable(5, 3, VrData(0))
                        isJogON = True
                    Else   'X+
                        VrData(0) = 1
                        If m_Xlocked = True Then
                            VrData(1) = 0.0
                        Else
                            VrData(1) = jogspeed
                        End If
                        VrData(2) = 0

                        m_TriProgMemHandle.SetTable(5, 3, VrData(0))
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

                        m_TriProgMemHandle.SetTable(5, 3, VrData(0))
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

                        m_TriProgMemHandle.SetTable(5, 3, VrData(0))
                        isJogON = True
                    Else   'X-
                        VrData(0) = 1
                        If m_Xlocked = True Then
                            VrData(1) = 0.0
                        Else
                            VrData(1) = -jogspeed
                        End If
                        VrData(2) = 0

                        m_TriProgMemHandle.SetTable(5, 3, VrData(0))
                        isJogON = True
                    End If
                Else
                    VrData(0) = 2
                    VrData(1) = 0.0
                    VrData(2) = 0.0

                    m_TriProgMemHandle.SetTable(5, 3, VrData(0))
                    isJogON = True 'False
                End If
            Else      'If Math.Abs(x) < Math.Abs(y) Then
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

                        m_TriProgMemHandle.SetTable(5, 3, VrData(0))
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

                        m_TriProgMemHandle.SetTable(5, 3, VrData(0))
                        isJogON = True
                    Else   'Y+
                        VrData(0) = 1
                        VrData(1) = 0
                        If m_Ylocked = True Then
                            VrData(2) = 0.0
                        Else
                            VrData(2) = jogspeed
                        End If

                        m_TriProgMemHandle.SetTable(5, 3, VrData(0))
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

                        m_TriProgMemHandle.SetTable(5, 3, VrData(0))
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

                        m_TriProgMemHandle.SetTable(5, 3, VrData(0))
                        isJogON = True
                    Else   'Y-
                        VrData(0) = 1
                        VrData(1) = 0
                        If m_Ylocked = True Then
                            VrData(2) = 0.0
                        Else
                            VrData(2) = -jogspeed
                        End If

                        m_TriProgMemHandle.SetTable(5, 3, VrData(0))
                        isJogON = True
                    End If
                Else
                    VrData(0) = 2
                    VrData(1) = 0.0
                    VrData(2) = 0.0

                    m_TriProgMemHandle.SetTable(5, 3, VrData(0))
                    isJogON = True 'False
                End If
            End If
        Else
            If isJogON Then
                VrData(0) = 2
                VrData(1) = 0.0
                VrData(2) = 0.0

                m_TriProgMemHandle.SetTable(5, 3, VrData(0))
                isJogON = False
            End If
            'If cursor_hide = True Then

            '    'Me.Cursor.Position = mouse_pos
            'Me.Cursor.Show()
            '    cursor_hide = False
            'End If
        End If

    End Sub


    Dim m_prevEstop As Integer = 0
    Dim m_Estop As Integer = 0
    Dim m_PressedSwitchID As Integer = 0

    Private Sub DealSwitchControl(ByVal switchID As Integer)
        Select Case switchID
            Case 1  '1=Park/ChangeSyr
                If gExeMode = "Programmer" Then
                    ChangeSyringeCtrl(m_CurrentNeedle)
                ElseIf gExeMode = "Operator" Then
                    FrmProduction.ChangeSyringeCtrl(FrmProduction.m_CurrentNeedle)
                End If
            Case 2  '2=Purge
                If gExeMode = "Programmer" Then
                    PurgeCtrl(m_CurrentNeedle)
                ElseIf gExeMode = "Operator" Then
                    FrmProduction.PurgeCtrl(FrmProduction.m_CurrentNeedle)
                End If

            Case 3  '3=Clean
                If gExeMode = "Programmer" Then
                    CleanCtrl(m_CurrentNeedle)
                ElseIf gExeMode = "Operator" Then
                    FrmProduction.CleanCtrl(FrmProduction.m_CurrentNeedle)
                End If

            Case 4  '4=NeedleCal
                If gExeMode = "Programmer" Then
                    NeedleCalibCtrl(m_CurrentNeedle)
                ElseIf gExeMode = "Operator" Then
                    FrmProduction.NeedleCalibCtrl(FrmProduction.m_CurrentNeedle)
                End If

            Case 5  '5= VolumeCal
                If gExeMode = "Programmer" Then
                    VolumCalibCtrl(m_CurrentNeedle)
                ElseIf gExeMode = "Operator" Then
                    FrmProduction.VolumCalibCtrl(FrmProduction.m_CurrentNeedle)
                End If

            Case 6  '6=Run
                If gExeMode = "Programmer" Then
                    If Not IDS.Devices.DIO.DIO.DoorStatus() Then
                        LabelMessege.Text = "Please close the door first!"
                        Exit Sub
                    End If
                    RunPatternProg()
                    m_OpBtnClick = 1
                ElseIf gExeMode = "Operator" Then
                    FrmProduction.prodStart()
                End If
            Case 7  '7=Stop
                If gExeMode = "Programmer" Then
                    m_OpBtnClick = 3
                    TBOperation.Buttons(0).Enabled = True
                    TBOperation.Buttons(1).Enabled = False
                    TBOperation.Buttons(2).Enabled = False
                    m_Tri.m_TriCtrl.SetTable(21, 1, 0)
                    m_Tri.m_TriCtrl.RapidStop()
                    m_Tri.m_TriCtrl.Stop("Dispense")
                    m_Tri.m_TriCtrl.Cancel(1, 0)
                    m_Tri.m_TriCtrl.Cancel(1, 1)
                    m_Tri.m_TriCtrl.Cancel(1, 2)
                    m_Tri.m_TriCtrl.Cancel(0, 0)
                    m_Tri.m_TriCtrl.Cancel(0, 1)
                    m_Tri.m_TriCtrl.Cancel(0, 2)
                    m_Tri.m_TriCtrl.Op(0)
                    m_Tri.m_TriCtrl.Op(gIOZbreak, 1)
                    If ThreadExecution Is Nothing Then
                    Else
                        ThreadExecution.Abort()
                    End If
                ElseIf gExeMode = "Operator" Then
                    FrmProduction.prodStop()
                End If
            Case 8  '8=Pause
                If gExeMode = "Programmer" Then
                    m_Tri.m_TriCtrl.SetTable(21, 1, 1)
                    m_OpBtnClick = 2
                    TBOperation.Buttons(0).Enabled = True
                    TBOperation.Buttons(2).Enabled = True
                    TBOperation.Buttons(1).Enabled = False
                ElseIf gExeMode = "Operator" Then
                    FrmProduction.prodPause()
                End If
            Case 9  '9= Reset
            Case 10
                QCcheck()
            Case Else
        End Select
    End Sub


    Public Function QCcheck() As Integer
        Dim vParam As DLL_Export_Device_Vision.QC.QCParam
        Dim readbuf(40) As Double
        m_Tri.m_TriCtrl.GetTable(210, 37, readbuf)
        vParam._Binarized = readbuf(0)

        Dim BlackDot As Boolean
        If readbuf(1) = 1 Then
            BlackDot = True
        Else
            BlackDot = False
        End If
        vParam._BlackDot = BlackDot
        vParam._Brightness = readbuf(2)
        vParam._Close = readbuf(3)
        vParam._Compactness = readbuf(4)
        Dim Inside_out As Boolean
        If readbuf(5) = 1 Then
            Inside_out = True
        Else
            Inside_out = False
        End If
        vParam._Inside_out = Inside_out
        vParam._MaxArea = readbuf(6)
        vParam._MinArea = readbuf(7)
        vParam._MQC_OffX = readbuf(8)
        vParam._MQC_OffY = readbuf(9)
        vParam._MRegionX = readbuf(10)
        vParam._MRegionY = readbuf(11)
        vParam._MROIx = readbuf(12)
        vParam._MROIy = readbuf(13)
        vParam._Open = readbuf(14)
        vParam._PointX1 = readbuf(15)
        vParam._PointX2 = readbuf(16)
        vParam._PointX3 = readbuf(17)
        vParam._PointX4 = readbuf(18)
        vParam._PointX5 = readbuf(19)
        vParam._PointY1 = readbuf(20)
        vParam._PointY2 = readbuf(21)
        vParam._PointY3 = readbuf(22)
        vParam._PointY4 = readbuf(23)
        vParam._PointY5 = readbuf(24)
        vParam._Pos = readbuf(25)
        vParam._PosX = readbuf(26)
        vParam._PosY = readbuf(27)
        vParam._ROI = readbuf(28)
        vParam._Rot = readbuf(29)
        vParam._Roughness = readbuf(30)
        vParam._Size = readbuf(31)
        vParam._SizeX = readbuf(32)
        vParam._SizeY = readbuf(33)
        vParam._Threshold = readbuf(34)
        vParam._type = readbuf(35)
        Dim Vertical As Boolean
        If readbuf(36) = 1 Then
            Vertical = True
        Else
            Vertical = False
        End If
        vParam._Vertical = Vertical
        'Dim aaa As Integer
        'For aaa = 1 To 10000
        '    Application.DoEvents()
        'Next aaa
        'Application.DoEvents()

        'Thread.Sleep(200)

        'vParam._Binarized = 128

        'vParam._BlackDot = True
        'vParam._Brightness = 128
        'vParam._Close = 0
        'vParam._Compactness = 5
        'vParam._Inside_out = True
        'vParam._MaxArea = 8
        'vParam._MinArea = 0.3
        'vParam._MQC_OffX = 0
        'vParam._MQC_OffY = 0
        'vParam._MRegionX = 384
        'vParam._MRegionY = 288
        'vParam._MROIx = 100
        'vParam._MROIy = 100
        'vParam._Open = 0
        'vParam._PointX1 = 0
        'vParam._PointX2 = 0
        'vParam._PointX3 = 0
        'vParam._PointX4 = 0
        'vParam._PointX5 = 0
        'vParam._PointY1 = 0
        'vParam._PointY2 = 0
        'vParam._PointY3 = 0
        'vParam._PointY4 = 0
        'vParam._PointY5 = 0
        'vParam._Pos = 0
        'vParam._PosX = 0
        'vParam._PosY = 0
        'vParam._ROI = 2
        'vParam._Rot = 5
        'vParam._Roughness = 5
        'vParam._Size = 0
        'vParam._SizeX = 0
        'vParam._SizeY = 0
        'vParam._Threshold = 15
        'vParam._type = 1
        'vParam._Vertical = True

        Dim rtn As Boolean = IDS.Devices.Vision.IDSV_QC(vParam)
        If Not rtn Then
            '''''test 
            gQcAutoSkip = False
            '''''
            If gQcAutoSkip Then   'auto skip
                m_Tri.m_TriCtrl.SetTable(202, 1, 1) 'skip
            Else     'manu
                While IDS.__IDEXE.Text Mod 100 <> 0
                    Application.DoEvents()
                End While

                IDS.__IDEXE.Text = 1006204
                form_service.LogAnEvent(IDS.__IDEXE)

                While IDS.__IDEXE.Text > 1006200 And IDS.__IDEXE.Text < 1006300
                    Application.DoEvents()
                End While

                If IDS.__IDEXE.Text = 1026204 Then
                    'skip
                    'MsgBox("sub skip clicked")
                    IDS.__IDEXE.Text = 1006200
                    m_Tri.m_TriCtrl.SetTable(202, 1, 1) 'skip
                ElseIf IDS.__IDEXE.Text = 1066204 Then
                    'ignore
                    'MsgBox("sub Stop clicked") 
                    IDS.__IDEXE.Text = 1006200
                    m_Tri.m_TriCtrl.SetTable(202, 1, 0) 'ignore
                End If
            End If
            m_Tri.m_TriCtrl.SetTable(201, 1, 1)
            Return -1
        End If
        m_Tri.m_TriCtrl.SetTable(202, 1, 0) '
        While Not m_Tri.m_TriCtrl.SetTable(201, 1, 1) 'qc finished
            Thread.Sleep(100)
        End While
        Return 0
    End Function
    Private Sub machineStatus()
        Dim MPos(3), off(3), ref(3), readbuf(4) As Double
        Dim Spos(3) As Double
        Dim Rx, Ry, Rz As Double
        Dim HPos(3) As Double
        m_Tri.m_TriCtrl.GetTable(0, 4, readbuf)
        HPos(0) = readbuf(0)
        HPos(1) = readbuf(1)
        HPos(2) = readbuf(2)
        m_Estop = readbuf(3)

        If m_prevEstop <> m_Estop Then
            If m_Estop = 1 Then
                PictureBoxEstopRed.Show()
                If IDS.Estop_ok_clicked Then
                    IDS.EstopPressed = True
                End If
            Else
                PictureBoxEstopRed.Hide()
            End If
            m_prevEstop = m_Estop
        End If

        ' MPos.CopyTo(HPos, 0)
        HardToSys(HPos, Spos)
        ref(0) = m_ReferPt(0) * gmminchRatio
        ref(1) = m_ReferPt(1) * gmminchRatio
        ref(2) = m_ReferPt(2) * gmminchRatio
        SysToRefer(Spos, Spos, ref)
        If m_TeachMode = 1 Then 'LEFT
            off(0) = gLeftNeedleOffs(0)
            off(1) = gLeftNeedleOffs(1)
            off(2) = -gLeftNeedleOffs(2)
        ElseIf m_TeachMode = 2 Then 'Right
            off(0) = gRightNeedleOffs(0)
            off(1) = gRightNeedleOffs(1)
            off(2) = -gRightNeedleOffs(2)
        Else 'vision
            off(0) = 0.0
            off(1) = 0.0
            off(2) = 0.0
        End If
        Spos(0) = Spos(0) + off(0)
        Spos(1) = Spos(1) + off(1)
        Spos(2) = Spos(2) + off(2) + m_BoardnRefBlkDist * gmminchRatio

        Rx = CInt(Spos(0) * 1000 / gmminchRatio) / 1000.0
        Ry = CInt(Spos(1) * 1000 / gmminchRatio) / 1000.0
        Rz = CInt(Spos(2) * 1000 / gmminchRatio) / 1000.0
        m_MachinePos(0) = CInt(HPos(0) * 1000 / gmminchRatio) / 1000.0
        m_MachinePos(1) = CInt(HPos(1) * 1000 / gmminchRatio) / 1000.0
        m_MachinePos(2) = CInt(HPos(2) * 1000 / gmminchRatio) / 1000.0
        If gExeMode <> "Operator" Then
            TextBoxRobotX.Text = "X" & m_MachinePos(0).ToString
            TextBoxRobotY.Text = "Y" & m_MachinePos(1).ToString
            TextBoxRobotZ.Text = "Z" & m_MachinePos(2).ToString
        Else
            FrmProduction.TextBoxRobotPos.Text = "X" & m_MachinePos(0).ToString & ", Y" & m_MachinePos(1).ToString & ", Z" & m_MachinePos(2).ToString
        End If


        If m_PosUpdate Then
            'FrmImportVision.GetCameraPos(m_CamPos)
            Dim colmX, colmY, colmZ As Integer
            Try

                If (m_PosNo = 1) Then
                    colmX = gPos1XColumn
                    colmY = gPos1YColumn
                    colmZ = gPos1ZColumn
                ElseIf (m_PosNo = 2) Then
                    colmX = gPos2XColumn
                    colmY = gPos2YColumn
                    colmZ = gPos2ZColumn
                Else
                    colmX = gPos3XColumn
                    colmY = gPos3YColumn
                    colmZ = gPos3ZColumn
                End If

                AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, colmX) = Rx 'm_CamPos(0) '
                AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, colmY) = Ry 'm_CamPos(1) '
                If m_TeachMode = 1 Or m_TeachMode = 2 Then 'LEFT Right
                    AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, colmZ) = Rz 'm_CamPos(1) 
                Else
                    AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, colmZ) = 0 'm_CamPos(1) 
                End If
                'AxSpreadsheetProgramming.Update()
            Catch ex As Exception

            End Try
        End If



        If gExeMode <> "Operator" Then
            Dim runState As Integer
            m_Tri.m_TriCtrl.GetTable(21, 1, runState)
            m_RunState = runState
            'm_TriProgMemHandle.GetTable(21, 21, readBuff(0))
            'm_RunState = CInt(readBuff(0))
            If (m_PrevRunState <> m_RunState) Then
                If (m_RunState = 0) Then 'stop
                    PBRed.Show()
                    PBYellow.Hide()
                    PBGreen.Hide()
                    m_OpBtnClick = 3
                    TBOperation.Buttons(0).Enabled = True
                    TBOperation.Buttons(2).Enabled = False
                    TBOperation.Buttons(1).Enabled = False
                    'If ThreadExecution.IsAlive Then ThreadExecution.Abort()
                ElseIf (m_RunState = 1) Then 'pause
                    PBRed.Hide()
                    PBYellow.Show()
                    PBGreen.Hide()
                    TBOperation.Buttons(0).Enabled = True
                    TBOperation.Buttons(2).Enabled = True
                    TBOperation.Buttons(1).Enabled = False
                Else 'run
                    PBRed.Hide()
                    PBYellow.Hide()
                    PBGreen.Show()
                    TBOperation.Buttons(0).Enabled = False
                    TBOperation.Buttons(2).Enabled = True
                    TBOperation.Buttons(1).Enabled = True
                End If
                m_PrevRunState = m_RunState
            End If
        End If

        '     If m_RunState = 0 Then
        m_PressedSwitchID = IDS.Devices.DIO.DIO.CheckOperationSwitch()
        If m_PressedSwitchID <> 0 Then
            bswitchCtrl(m_PressedSwitchID)
        End If
        '    End If
    End Sub

    Public Sub SetRunMode(ByVal mode As Integer)
        m_RunMode = mode
    End Sub

    Public Function DispensingCtrl() As Integer
        Dim Dispensinglist As ArrayList
        If m_Execution.m_Command.CompileStatus > 0 Then
            Dispensinglist = m_Execution.m_Command.DispenseList
            If m_Tri.m_TriCtrl.IsOpen(0) Then
                m_Tri.m_TriCtrl.SetTable(20, 1, m_RunMode)
            Else
                If gExeMode <> "Operator" Then
                    LabelMessege.Text = "Controller not opened"
                    LabelMessege.Refresh()
                Else
                    FrmProduction.LabelMessege.Text = "Controller not opened"
                    FrmProduction.LabelMessege.Refresh()
                End If

                ' TBOperation.Enabled = True
                Return -1
            End If
            'test use only
            'Dim frm As New FormTmp(Dispensinglist)
            'frm.ShowDialog()
            'Exit Function
            If gExeMode <> "Operator" Then
                LabelMessege.Text = "Dispensing ..."
                LabelMessege.Refresh()
            Else
                FrmProduction.LabelMessege.Text = "Dispensing ..."
                FrmProduction.LabelMessege.Refresh()
            End If

            Dim cmdBurn As New CIDSPattnBurn(m_Tri, m_TriProgMemHandle)

            If (cmdBurn.BurnTable(Dispensinglist) < 0) Then
                If gExeMode <> "Operator" Then
                    LabelMessege.Text = "Download program error"
                    LabelMessege.Refresh()
                Else
                    FrmProduction.LabelMessege.Text = "Download program error"
                    FrmProduction.LabelMessege.Refresh()
                End If

                'TBOperation.Enabled = True
                Return -1
            Else
                'LabelMessege.Text = "Download finished"
            End If
            If gExeMode <> "Operator" Then
                LabelMessege.Text = "Download finished"
                LabelMessege.Refresh()
            Else
                FrmProduction.LabelMessege.Text = "Dowload finished"
                FrmProduction.LabelMessege.Refresh()
            End If

            ' TBOperation.Enabled = True
        End If
        Return 0
    End Function

    Dim hideON As Boolean = False
    Private Sub Reset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Reset.Click
        Dim buf(30) As Single
        Dim start As DateTime = DateTime.Now
        LabelMessege.Text = start.ToString
        LabelMessege.Update()
        Dim posX, posY, posZ As Double
        Dim no As Double = 1.2

        Dim vrdata(100) As Single
        vrdata(0) = 123
        'Dim rtn As Boolean = m_TriProgMemHandle.SetVr(100, 100, vrdata(0))
        Dim rtn As Boolean = m_TriProgMemHandle.GetVr(100, 100, vrdata(0))
        Dim tend As DateTime = DateTime.Now
        Dim span As Integer = (tend.Ticks - start.Ticks) * 0.0001

        LabelMessege.Text = span.ToString
        'Dim inValue As Integer
        'Dim rtn As Boolean = m_Tri.SetDIOs(0, 17, 1)

        Exit Sub
        If m_Tri.m_TriCtrl.IsOpen(0) Then
            m_Tri.m_TriCtrl.Run("STARTUP")
            m_Tri.m_TriCtrl.Run("LOGPOS", 4)
        Else
            LabelMessege.Text = "Controller not opened"
        End If
    End Sub


    'test online run
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        AxSpreadsheetProgramming.Update()
        'TBOperation.Enabled = False
        If m_Execution.m_Command.ReadPattern(AxSpreadsheetProgramming) < 0 Then
            LabelMessege.Text = "Empty sheet or Error"
            LabelMessege.Refresh()
            'TBOperation.Enabled = True
            Exit Sub
        End If

        If m_Execution.m_Command.Compile(m_RunMode) < 0 Then
            LabelMessege.Text = "Empty sheet or Error"
            LabelMessege.Refresh()
            'TBOperation.Enabled = True
            Exit Sub
        End If
        Dim tmplist As ArrayList
        If m_Execution.m_Command.CompileStatus > 0 Then
            tmplist = m_Execution.m_Command.DispenseList
            If m_Tri.m_TriCtrl.IsOpen(0) Then
                'm_Tri.m_TriCtrl.SetVr(20, m_RunMode)
                m_Tri.m_TriCtrl.SetTable(20, 1, m_RunMode)
            Else
                LabelMessege.Text = "Controller not opened"
                LabelMessege.Refresh()
            End If
            'test use only
            'Dim frm As New FormTmp(tmplist)
            'frm.ShowDialog()
            'Exit Sub
            LabelMessege.Text = "Dispensing ..."
            LabelMessege.Refresh()

            Dim cmdBurn As New CIDSPattnBurn(m_Tri, m_TriProgMemHandle)
            If (cmdBurn.OnlineRun(tmplist) = 0) Then
                LabelMessege.Text = "Dispensing finished"
            Else
                LabelMessege.Text = "Dispensing error"
            End If
            'TBOperation.Enabled = True
        End If


    End Sub

    Private Sub Download_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Download.Click
        LabelMessege.Text = "Downloading ..."
        LabelMessege.Refresh()
        Dim prevT As Long = Now.Ticks

        Dim currT As Long = Now.Ticks
        Dim pt As Long = (currT - prevT) * 0.0001

        prevT = Now.Ticks
        'TBOperation.Enabled = False
        If m_Execution.m_Command.ReadPattern(AxSpreadsheetProgramming) < 0 Then
            LabelMessege.Text = "Empty sheet or Error"
            LabelMessege.Refresh()
            'TBOperation.Enabled = True
            Exit Sub
        End If

        currT = Now.Ticks
        pt = (currT - prevT) * 0.0001
        LabelMessege.Text = "Compiling..."
        LabelMessege.Update()
        prevT = Now.Ticks

        If m_Execution.m_Command.Compile(m_RunMode) < 0 Then
            LabelMessege.Text = "Empty sheet or Error"
            LabelMessege.Refresh()
            ' TBOperation.Enabled = True
            Exit Sub
        End If
        currT = Now.Ticks
        pt = (currT - prevT) * 0.0001

        DispensingCtrl()

    End Sub

    Public Function Connect_Controller() As Boolean
        If Not m_TriProgMemHandle.isOpen() Then
            m_TriProgMemHandle.OpenDevice()
        End If
        If Not m_Tri.m_TriCtrl.IsOpen(0) Then
            If m_Tri.TrioConnectPci(0) Then
                m_Tri.m_TriCtrl.Run("STARTUP", 6)
                Threading.Thread.Sleep(500)
                m_Tri.m_TriCtrl.Run("JOGGING", 5)
                Threading.Thread.Sleep(500)
                m_Tri.m_TriCtrl.Run("LOGPOS", 4)
                'm_Tri.m_TriCtrl.Run("STEPPING", 3)
            Else
                MsgBox("Can't connect to controller")
            End If
        End If

    End Function

    Public Function Disconnect_Controller() As Boolean
        If m_Tri.m_TriCtrl.IsOpen(0) Then
            m_Tri.m_TriCtrl.Stop("LOGPOS")
            m_Tri.m_TriCtrl.Stop("JOGGING")
            'm_Tri.m_TriCtrl.Stop("STEPPING")
            m_Tri.m_TriCtrl.Run("EXITIDS")
            m_Tri.m_TriCtrl.Close(0)
        End If
        If m_TriProgMemHandle.isOpen() Then
            m_TriProgMemHandle.CloseDevice()
        End If

    End Function

    Private Sub ComboBox1_SelectedValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedValueChanged
        m_RunMode = Me.ComboBox1.SelectedIndex
    End Sub


    Private Sub AxSpreadsheetProgramming_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub


    Private Sub NumericUpDownSysSpeed_ValueChanged(ByVal sender As Object, _
        ByVal e As System.EventArgs) Handles NumericUpDownSysSpeed.ValueChanged
        ' gSystemSpeed = NumericUpDownSysSpeed.Value()
    End Sub


#Region " Gantry Step Move"
    Dim dStepVal(5) As Single
    Dim m_Step As Double = 0.1

    Private Sub BTStepXminus_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTStepXminus.Click
        LabelMessege.Text = ""
        If CheckBoxLockX.Checked = True Then
            LabelMessege.Text = "X axis is currrently locked!"
            Exit Sub
        End If
        dStepVal(0) = 1
        dStepVal(1) = gStepCtrlSpeed
        dStepVal(2) = -m_Step
        dStepVal(3) = 0
        dStepVal(4) = 0
        Try
            m_TriProgMemHandle.SetTable(50, 5, dStepVal(0))
        Catch ex As Exception

        End Try
    End Sub

    Private Sub BTStepXplus_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTStepXplus.Click
        LabelMessege.Text = ""
        If CheckBoxLockX.Checked = True Then
            LabelMessege.Text = "X axis is currrently locked!"
            Exit Sub
        End If
        dStepVal(0) = 1
        dStepVal(1) = gStepCtrlSpeed
        dStepVal(2) = m_Step
        dStepVal(3) = 0
        dStepVal(4) = 0
        Try
            m_TriProgMemHandle.SetTable(50, 5, dStepVal(0))
        Catch ex As Exception

        End Try
    End Sub

    Private Sub BTStepYminus_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTStepYminus.Click
        LabelMessege.Text = ""
        If CheckBoxLockY.Checked = True Then
            LabelMessege.Text = "Y axis is currrently locked!"
            Exit Sub
        End If
        dStepVal(0) = 1
        dStepVal(1) = gStepCtrlSpeed
        dStepVal(2) = 0
        dStepVal(3) = -m_Step
        dStepVal(4) = 0
        Try
            m_TriProgMemHandle.SetTable(50, 5, dStepVal(0))
        Catch ex As Exception

        End Try
    End Sub

    Private Sub BTStepYplus_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTStepYplus.Click
        LabelMessege.Text = ""
        If CheckBoxLockY.Checked = True Then
            LabelMessege.Text = "Y axis is currrently locked!"
            Exit Sub
        End If
        dStepVal(0) = 1
        dStepVal(1) = gStepCtrlSpeed
        dStepVal(2) = 0
        dStepVal(3) = m_Step
        dStepVal(4) = 0
        Try
            m_TriProgMemHandle.SetTable(50, 5, dStepVal(0))
        Catch ex As Exception

        End Try
    End Sub

    Private Sub BTStepZup_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTStepZup.Click
        LabelMessege.Text = ""
        If CheckBoxLockZ.Checked = True Then
            LabelMessege.Text = "Z axis is currrently locked!"
            Exit Sub
        End If
        dStepVal(0) = 1
        dStepVal(1) = gStepCtrlSpeed
        dStepVal(2) = 0
        dStepVal(3) = 0
        dStepVal(4) = m_Step
        Try
            m_TriProgMemHandle.SetTable(50, 5, dStepVal(0))
        Catch ex As Exception

        End Try
    End Sub

    Private Sub BTStepZdown_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTStepZdown.Click
        LabelMessege.Text = ""
        If CheckBoxLockZ.Checked = True Then
            LabelMessege.Text = "Z axis is currrently locked!"
            Exit Sub
        End If
        dStepVal(0) = 1
        dStepVal(1) = gStepCtrlSpeed
        dStepVal(2) = 0
        dStepVal(3) = 0
        dStepVal(4) = -m_Step
        Try
            m_TriProgMemHandle.SetTable(50, 5, dStepVal(0))
        Catch ex As Exception

        End Try
    End Sub

    Private Sub ComboBoxFineStep_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBoxFineStep.SelectedValueChanged
        Dim str As String = ComboBoxFineStep.SelectedItem
        Dim tempStep As Double = CDbl(str)
        If tempStep >= 0.001 And tempStep <= 10.0 Then
            m_Step = tempStep
        Else
            LabelMessege.Text = "Must be between 0.001 to 10."
        End If
    End Sub

    Private Sub ComboBoxFineStep_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBoxFineStep.TextChanged
        Dim dVal As Double = 0.0
        Dim str As String = ComboBoxFineStep.Text
        LabelMessege.Text = ""
        If ComboBoxFineStep.Text <> Nothing Then dVal = CDbl(str)
        If dVal >= 0.001 And dVal <= 10.0 Then
            m_Step = dVal
        Else
            LabelMessege.Text = "Must be between 0.001 to 10."
            ' ComboBoxFineStep.Text = m_Step.ToString
        End If
        'xu long 190707 /*
        Try
            If ComboBoxFineStep.Text <= 10 And ComboBoxFineStep.Text >= 0.001 Then
                m_Step = ComboBoxFineStep.Text
            ElseIf ComboBoxFineStep.Text < 0.001 And ComboBoxFineStep.Text >= 0 Then
                'ComboBoxFineStep.Text = ""
                LabelMessege.Text = "Must be between 0.001 to 10."
                m_Step = 0
            Else
                ComboBoxFineStep.Text = ""
                LabelMessege.Text = "Must be between 0.001 to 10."
                m_Step = 0
            End If

        Catch ex As Exception
            ComboBoxFineStep.Text = ""
            LabelMessege.Text = "Must be between 0.001 to 10."
            m_Step = 0
        End Try
        'xu long 190707 */
    End Sub

#End Region





#End Region

#Region "xu long"


    'Private Sub TBOperation_ButtonClick_1(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ToolBarButtonClickEventArgs) Handles TBOperation.ButtonClick
    '    Try

    '        If e.Button Is TBOperation.Buttons(0) Then 'run
    '            Dim mode As Integer = 0
    '            m_Tri.m_TriCtrl.GetTable(21, 1, mode)
    '            If mode = 2 Then  'is running
    '                Exit Sub
    '            ElseIf mode = 1 Then 'pause
    '                m_Tri.m_TriCtrl.SetTable(21, 1, 2)
    '            Else  'stop
    '                LabelMessege.Text = "Downloading ..."
    '                LabelMessege.Refresh()

    '                m_Tri.m_TriCtrl.SetTable(21, 1, 2)

    '                Dim prevT As Long = Now.Ticks

    '                Dim currT As Long = Now.Ticks
    '                Dim pt As Long = (currT - prevT) * 0.0001

    '                prevT = Now.Ticks
    '                'TBOperation.Enabled = False
    '                Dim rtn As Integer = m_Execution.m_Command.ReadPattern(AxSpreadsheetProgramming)
    '                If rtn < 0 Then
    '                    m_Tri.m_TriCtrl.SetTable(21, 1, 0)
    '                    LabelMessege.Text = "Empty sheet or Error"
    '                    LabelMessege.Refresh()
    '                    'TBOperation.Enabled = True
    '                    Exit Sub
    '                ElseIf rtn > 0 Then
    '                    m_Tri.m_TriCtrl.SetTable(21, 1, 0)
    '                    LabelMessege.Text = "Execution stopped"
    '                    LabelMessege.Refresh()
    '                    'TBOperation.Enabled = True
    '                    Exit Sub
    '                End If

    '                currT = Now.Ticks
    '                pt = (currT - prevT) * 0.0001
    '                LabelMessege.Text = "Compiling..."
    '                LabelMessege.Update()
    '                prevT = Now.Ticks

    '                If m_Execution.m_Command.Compile(m_RunMode) < 0 Then
    '                    LabelMessege.Text = "Empty sheet or Error"
    '                    LabelMessege.Refresh()
    '                    ' TBOperation.Enabled = True
    '                    Exit Sub
    '                End If
    '                currT = Now.Ticks
    '                pt = (currT - prevT) * 0.0001


    '                ThreadExecution = New Threading.Thread(AddressOf TheThreadExecution)
    '                ThreadExecution.Priority = Threading.ThreadPriority.Highest
    '                ThreadExecution.Start()

    '            End If
    '            m_OpBtnClick = 1
    '        ElseIf e.Button Is TBOperation.Buttons(1) Then 'pause
    '            m_Tri.m_TriCtrl.SetTable(21, 1, 1)
    '            m_OpBtnClick = 2
    '        ElseIf e.Button Is TBOperation.Buttons(2) Then 'stop
    '            m_OpBtnClick = 3
    '            m_Tri.m_TriCtrl.SetTable(21, 1, 0)
    '            m_Tri.m_TriCtrl.RapidStop()
    '            m_Tri.m_TriCtrl.Stop("Dispense")
    '            m_Tri.m_TriCtrl.Cancel(1, 0)
    '            m_Tri.m_TriCtrl.Cancel(1, 1)
    '            m_Tri.m_TriCtrl.Cancel(1, 2)
    '            m_Tri.m_TriCtrl.Cancel(0, 0)
    '            m_Tri.m_TriCtrl.Cancel(0, 1)
    '            m_Tri.m_TriCtrl.Cancel(0, 2)
    '            If ThreadExecution Is Nothing Then
    '            Else
    '                ThreadExecution.Abort()
    '            End If

    '        End If

    '    Catch ex As SystemException
    '        MessageBox.Show(ex.ToString)
    '    End Try
    'End Sub

    Private Sub BT_Pro_LDownload_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BT_Pro_LDownload.Click
        'IDS.Data.OpenData()
        If IDS.Data.Hardware.Dispenser.Left.HeadType = "Time Pressure" Then
            IDS.Devices.Regulator.FrmRegulator.Left_SetPressure(0, TB_Pro_LPressure.Text, TB_Pro_LSuck.Text)

            IDS.Data.Hardware.Dispenser.TimePressurePressure_Left = TB_Pro_LPressure.Text
            IDS.Data.Hardware.Dispenser.TimePressureSuckBack_Left = TB_Pro_LSuck.Text
            'IDS.Data.Hardware.Dispenser.PositiveDispRPM_Left()
            'IDS.Data.Hardware.Dispenser.PositiveDispReverseRev_Left()
            'IDS.Data.Hardware.Dispenser.PositiveDispPressure_Left()
            'IDS.Data.Hardware.Dispenser.PositiveDispSuckBack_Left()


        ElseIf IDS.Data.Hardware.Dispenser.Left.HeadType = "Positive Displacement" Then
            Try
                IDS.Devices.SPAV.FrmSPAV.Open_Port_SPAV(1)
            Catch
            End Try
            Sleep(400)

            Dim prog As Byte = TB_Pro_LProg.Text
            Dim valve As Byte = TB_Pro_LValve.Text
            Dim RPM As Short = TB_Pro_LRPM.Text
            Dim REV As Short = TB_Pro_LREV.Text
            IDS.Devices.SPAV.FrmSPAV.DLDataToAVC(prog, valve, 0, RPM, 1, REV)

            IDS.Devices.Regulator.FrmRegulator.Left_SetPressure(0, TB_Pro_LPressure.Text, TB_Pro_LSuck.Text)

            IDS.Data.Hardware.Dispenser.TimePressurePressure_Left = TB_Pro_LPressure.Text
            IDS.Data.Hardware.Dispenser.TimePressureSuckBack_Left = TB_Pro_LSuck.Text
            IDS.Data.Hardware.Dispenser.PositiveDispRPM_Left = RPM
            IDS.Data.Hardware.Dispenser.PositiveDispReverseRev_Left = REV
            IDS.Data.Hardware.Dispenser.PositiveDispPressure_Left = valve
            IDS.Data.Hardware.Dispenser.PositiveDispSuckBack_Left = prog

        ElseIf IDS.Data.Hardware.Dispenser.Left.HeadType = "Fixed Volumetric" Then
            IDS.Devices.Regulator.FrmRegulator.Left_SetPressure(0, TB_Pro_LPressure.Text, TB_Pro_LSuck.Text)
            IDS.Data.Hardware.Dispenser.TimePressurePressure_Left = TB_Pro_LPressure.Text
            IDS.Data.Hardware.Dispenser.TimePressureSuckBack_Left = TB_Pro_LSuck.Text
            'ElseIf CB_LeftDispenserType.Text = "Jetter" Then

        End If

        IDS.Data.SaveData()
    End Sub

    Private Sub BT_Pro_RDownload_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BT_Pro_RDownload.Click
        'IDS.Data.OpenData()
        If IDS.Data.Hardware.Dispenser.Right.HeadType = "Time Pressure" Then
            IDS.Devices.Regulator.FrmRegulator.Right_SetPressure(0, TB_Pro_RPressure.Text, TB_Pro_RSuck.Text)

            IDS.Data.Hardware.Dispenser.TimePressurePressure_Right = TB_Pro_RPressure.Text
            IDS.Data.Hardware.Dispenser.TimePressureSuckBack_Right = TB_Pro_RSuck.Text
        ElseIf IDS.Data.Hardware.Dispenser.Right.HeadType = "Positive Displacement" Then
            Try
                IDS.Devices.SPAV.FrmSPAV.Open_Port_SPAV(1)
            Catch
            End Try
            Sleep(400)

            Dim prog As Byte = TB_Pro_RProg.Text
            Dim valve As Byte = TB_Pro_RValve.Text
            Dim RPM As Short = TB_Pro_RRPM.Text
            Dim REV As Short = TB_Pro_RREV.Text
            IDS.Devices.SPAV.FrmSPAV.DLDataToAVC(prog, valve, 0, RPM, 1, REV)

            IDS.Devices.Regulator.FrmRegulator.Right_SetPressure(0, TB_Pro_RPressure.Text, TB_Pro_RSuck.Text)

            IDS.Data.Hardware.Dispenser.TimePressurePressure_Right = TB_Pro_RPressure.Text
            IDS.Data.Hardware.Dispenser.TimePressureSuckBack_Right = TB_Pro_RSuck.Text
            IDS.Data.Hardware.Dispenser.PositiveDispRPM_Right = RPM
            IDS.Data.Hardware.Dispenser.PositiveDispReverseRev_Right = REV
            IDS.Data.Hardware.Dispenser.PositiveDispPressure_Right = valve
            IDS.Data.Hardware.Dispenser.PositiveDispSuckBack_Right = prog

        ElseIf IDS.Data.Hardware.Dispenser.Right.HeadType = "Fixed Volumetric" Then
            IDS.Devices.Regulator.FrmRegulator.Right_SetPressure(0, TB_Pro_RPressure.Text, TB_Pro_RSuck.Text)
            IDS.Data.Hardware.Dispenser.TimePressurePressure_Right = TB_Pro_RPressure.Text
            IDS.Data.Hardware.Dispenser.TimePressureSuckBack_Right = TB_Pro_RSuck.Text
            'ElseIf CB_LeftDispenserType.Text = "Jetter" Then

        End If

        IDS.Data.SaveData()
    End Sub


    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Me.Controls.Add(FrmProduction.PanelProduction)
        FrmProduction.PanelProduction.BringToFront()
        FrmProduction.PanelProduction.Location = New Point(0, 0)
        Me.Menu = FrmProduction.MainMenuProduction
    End Sub
    Private Sub NumericUpDown_Brightness_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NumericUpDown_Brightness.ValueChanged
        If 0 < NumericUpDown_Brightness.Value < 255 Then
            IDS.Devices.Vision.IDSV_SetBrightness(NumericUpDown_Brightness.Value)
        Else
            NumericUpDown_Brightness.Value = 128
            IDS.Devices.Vision.IDSV_SetBrightness(NumericUpDown_Brightness.Value)
        End If
    End Sub
    'xu long 190707 /*
    Private Sub BT_CV_Retrieve_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BT_CV_Retrieve.Click
        IDS.T1.Stop()
        Sleep(40)
        IDS.Devices.Conveyor.FrmConveyor.Retrieve()
        IDS.T1.Start()
        Sleep(20)
    End Sub

    Private Sub BT_CV_Release_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BT_CV_Release.Click
        IDS.T1.Stop()
        Sleep(40)
        IDS.Devices.Conveyor.FrmConveyor.Release()
        IDS.T1.Start()
        Sleep(20)
    End Sub

    Private Sub TimerForUpdate_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TimerForUpdate.Tick
        TimerForUpdate.Stop()
        'IDS.Data.OpenData()

        LB_SyrHeater.Text = IDS.Devices.Thermal.FrmThermal.Temperature(0)
        FrmProduction.LB_Prod_SyrHeater.Text = IDS.Devices.Thermal.FrmThermal.Temperature(0)

        LB_NeedleHeater.Text = IDS.Devices.Thermal.FrmThermal.Temperature(1)
        FrmProduction.LB_Prod_NeedleHeater.Text = IDS.Devices.Thermal.FrmThermal.Temperature(1)

        LB_PreHeater.Text = IDS.Devices.Thermal.FrmThermal.Temperature(2)
        FrmProduction.LB_Prod_PreHeater.Text = IDS.Devices.Thermal.FrmThermal.Temperature(2)

        LB_DispHeater.Text = IDS.Devices.Thermal.FrmThermal.Temperature(3)
        FrmProduction.LB_Prod_DispHeater.Text = IDS.Devices.Thermal.FrmThermal.Temperature(3)

        LB_PostHeater.Text = IDS.Devices.Thermal.FrmThermal.Temperature(4)
        FrmProduction.LB_Prod_PostHeater.Text = IDS.Devices.Thermal.FrmThermal.Temperature(4)

        TimerForUpdate.Start()
    End Sub

    Private Sub BT_HeaterCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BT_HeaterCancel.Click
        IDS.Data.OpenData()
        TB_NeedleHeater.Text = IDS.Data.Hardware.Thermal.Needle.SetTemp
        TB_SyrHeater.Text = IDS.Data.Hardware.Thermal.Syringe.SetTemp
        TB_PreHeater.Text = IDS.Data.Hardware.Thermal.Station1.SetTemp
        TB_DispHeater.Text = IDS.Data.Hardware.Thermal.Station2.SetTemp
        TB_PostHeater.Text = IDS.Data.Hardware.Thermal.Station3.SetTemp
    End Sub

    Private Sub BT_HeaterUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BT_HeaterUpdate.Click
        IDS.T1.Stop()
        'IDS.Data.OpenData()
        If TB_NeedleHeater.Enabled Then
            IDS.Devices.Thermal.FrmThermal.Set_Temperature(0, TB_NeedleHeater.Text)
            IDS.Data.Hardware.Thermal.Needle.SetTemp = TB_NeedleHeater.Text
        End If
        If TB_SyrHeater.Enabled Then
            IDS.Devices.Thermal.FrmThermal.Set_Temperature(1, TB_SyrHeater.Text)
            IDS.Data.Hardware.Thermal.Syringe.SetTemp = TB_SyrHeater.Text
        End If
        If TB_PreHeater.Enabled Then
            IDS.Devices.Thermal.FrmThermal.Set_Temperature(2, TB_PreHeater.Text)
            IDS.Data.Hardware.Thermal.Station1.SetTemp = TB_PreHeater.Text
        End If
        If TB_DispHeater.Enabled Then
            IDS.Devices.Thermal.FrmThermal.Set_Temperature(3, TB_DispHeater.Text)
            IDS.Data.Hardware.Thermal.Station2.SetTemp = TB_DispHeater.Text
        End If
        If TB_PostHeater.Enabled Then
            IDS.Devices.Thermal.FrmThermal.Set_Temperature(4, TB_PostHeater.Text)
            IDS.Data.Hardware.Thermal.Station3.SetTemp = TB_PostHeater.Text
        End If
        IDS.Data.SaveData()
        IDS.T1.Start()
    End Sub

    Private Sub DispInforPanel() 'xu long
        'IDS.Data.OpenData()
        If IDS.Data.Hardware.Thermal.Needle.HeaterPresentFlag And IDS.Data.Hardware.Thermal.Needle.OnOffControl Then
            TB_NeedleHeater.Text = ""
            TB_NeedleHeater.Enabled = False
            LB_NeedleHeater.Visible = False
            FrmProduction.LB_Prod_NeedleHeater.Visible = False
        Else
            TB_NeedleHeater.Text = IDS.Data.Hardware.Thermal.Needle.SetTemp
            TB_NeedleHeater.Enabled = True
            LB_NeedleHeater.Visible = True
            FrmProduction.LB_Prod_NeedleHeater.Visible = True
        End If

        If IDS.Data.Hardware.Thermal.Syringe.HeaterPresentFlag And IDS.Data.Hardware.Thermal.Syringe.OnOffControl Then
            TB_SyrHeater.Text = ""
            TB_SyrHeater.Enabled = False
            LB_SyrHeater.Visible = False
            FrmProduction.LB_Prod_SyrHeater.Visible = False
        Else
            TB_SyrHeater.Text = IDS.Data.Hardware.Thermal.Syringe.SetTemp
            TB_SyrHeater.Enabled = True
            LB_SyrHeater.Visible = True
            FrmProduction.LB_Prod_SyrHeater.Visible = True
        End If

        If IDS.Data.Hardware.Thermal.Station1.HeaterPresentFlag And IDS.Data.Hardware.Thermal.Station1.OnOffControl Then
            TB_PreHeater.Text = ""
            TB_PreHeater.Enabled = False
            LB_PreHeater.Visible = False
            FrmProduction.LB_Prod_PreHeater.Visible = False
        Else
            TB_PreHeater.Text = IDS.Data.Hardware.Thermal.Station1.SetTemp
            TB_PreHeater.Enabled = True
            LB_PreHeater.Visible = True
            FrmProduction.LB_Prod_PreHeater.Visible = True
        End If

        If IDS.Data.Hardware.Thermal.Station2.HeaterPresentFlag And IDS.Data.Hardware.Thermal.Station2.OnOffControl Then
            TB_DispHeater.Text = ""
            TB_DispHeater.Enabled = False
            LB_DispHeater.Visible = False
            FrmProduction.LB_Prod_DispHeater.Visible = False
        Else
            TB_DispHeater.Text = IDS.Data.Hardware.Thermal.Station2.SetTemp
            TB_DispHeater.Enabled = True
            LB_DispHeater.Visible = True
            FrmProduction.LB_Prod_DispHeater.Visible = True
        End If

        If IDS.Data.Hardware.Thermal.Station3.HeaterPresentFlag And IDS.Data.Hardware.Thermal.Station3.OnOffControl Then
            TB_PostHeater.Text = ""
            TB_PostHeater.Enabled = False
            LB_PostHeater.Visible = False
            FrmProduction.LB_Prod_PostHeater.Visible = False
        Else
            TB_PostHeater.Text = IDS.Data.Hardware.Thermal.Station3.SetTemp
            TB_PostHeater.Enabled = True
            LB_PostHeater.Visible = True
            FrmProduction.LB_Prod_PostHeater.Visible = True
        End If

        'If IDS.Data.Hardware.Dispenser.Left.HeadType = "Time Pressure" Then
        '    GB_LTP.Visible = True
        '    FrmProduction.GB_LTP.Visible = True
        '    GB_LJetter.Visible = False
        '    FrmProduction.GB_LJetter.Visible = False

        '    LB_Pro_LRPM.Visible = False
        '    TB_Pro_LRPM.Visible = False
        '    LB_Pro_LREV.Visible = False
        '    TB_Pro_LREV.Visible = False
        '    LB_Pro_LValve.Visible = False
        '    TB_Pro_LValve.Visible = False
        '    LB_Pro_LProg.Visible = False
        '    TB_Pro_LProg.Visible = False

        '    FrmProduction.LB_Pro_LRPM.Visible = False
        '    FrmProduction.TB_Pro_LRPM.Visible = False
        '    FrmProduction.LB_Pro_LREV.Visible = False
        '    FrmProduction.TB_Pro_LREV.Visible = False
        '    FrmProduction.LB_Pro_LValve.Visible = False
        '    FrmProduction.TB_Pro_LValve.Visible = False
        '    FrmProduction.LB_Pro_LProg.Visible = False
        '    FrmProduction.TB_Pro_LProg.Visible = False



        'ElseIf IDS.Data.Hardware.Dispenser.Left.HeadType = "Positive Displacement" Then
        '    GB_LTP.Visible = True
        '    FrmProduction.GB_LTP.Visible = True
        '    GB_LJetter.Visible = False
        '    FrmProduction.GB_LJetter.Visible = False

        '    LB_Pro_LRPM.Visible = True
        '    TB_Pro_LRPM.Visible = True
        '    LB_Pro_LREV.Visible = True
        '    TB_Pro_LREV.Visible = True
        '    LB_Pro_LValve.Visible = True
        '    TB_Pro_LValve.Visible = True
        '    LB_Pro_LProg.Visible = True
        '    TB_Pro_LProg.Visible = True

        '    FrmProduction.LB_Pro_LRPM.Visible = True
        '    FrmProduction.TB_Pro_LRPM.Visible = True
        '    FrmProduction.LB_Pro_LREV.Visible = True
        '    FrmProduction.TB_Pro_LREV.Visible = True
        '    FrmProduction.LB_Pro_LValve.Visible = True
        '    FrmProduction.TB_Pro_LValve.Visible = True
        '    FrmProduction.LB_Pro_LProg.Visible = True
        '    FrmProduction.TB_Pro_LProg.Visible = True


        'ElseIf IDS.Data.Hardware.Dispenser.Left.HeadType = "Fixed Volumetric" Then
        '    GB_LTP.Visible = True
        '    FrmProduction.GB_LTP.Visible = True
        '    GB_LJetter.Visible = False
        '    FrmProduction.GB_LJetter.Visible = False

        '    LB_Pro_LRPM.Visible = False
        '    TB_Pro_LRPM.Visible = False
        '    LB_Pro_LREV.Visible = False
        '    TB_Pro_LREV.Visible = False
        '    LB_Pro_LValve.Visible = False
        '    TB_Pro_LValve.Visible = False
        '    LB_Pro_LProg.Visible = False
        '    TB_Pro_LProg.Visible = False

        '    FrmProduction.LB_Pro_LRPM.Visible = False
        '    FrmProduction.TB_Pro_LRPM.Visible = False
        '    FrmProduction.LB_Pro_LREV.Visible = False
        '    FrmProduction.TB_Pro_LREV.Visible = False
        '    FrmProduction.LB_Pro_LValve.Visible = False
        '    FrmProduction.TB_Pro_LValve.Visible = False
        '    FrmProduction.LB_Pro_LProg.Visible = False
        '    FrmProduction.TB_Pro_LProg.Visible = False


        'ElseIf IDS.Data.Hardware.Dispenser.Left.HeadType = "Jetter" Then
        '    GB_LTP.Visible = False
        '    FrmProduction.GB_LTP.Visible = False
        '    GB_LJetter.Visible = True
        '    FrmProduction.GB_LJetter.Visible = True
        'Else
        '    GB_LTP.Visible = False
        '    FrmProduction.GB_LTP.Visible = False
        '    GB_LJetter.Visible = False
        '    FrmProduction.GB_LJetter.Visible = False
        'End If

        'If IDS.Data.Hardware.Dispenser.Right.HeadType = "Time Pressure" Then
        '    GB_RTP.Visible = True
        '    FrmProduction.GB_RTP.Visible = True
        '    GB_RJetter.Visible = False
        '    FrmProduction.GB_RJetter.Visible = False
        'ElseIf IDS.Data.Hardware.Dispenser.Right.HeadType = "Positive Displacement" Then
        '    GB_RTP.Visible = True
        '    FrmProduction.GB_RTP.Visible = True
        '    GB_RJetter.Visible = False
        '    FrmProduction.GB_RJetter.Visible = False
        'ElseIf IDS.Data.Hardware.Dispenser.Right.HeadType = "Fixed Volumetric" Then
        '    GB_RTP.Visible = True
        '    FrmProduction.GB_RTP.Visible = True
        '    GB_RJetter.Visible = False
        '    FrmProduction.GB_RJetter.Visible = False
        'ElseIf IDS.Data.Hardware.Dispenser.Right.HeadType = "Jetter" Then
        '    GB_RTP.Visible = False
        '    FrmProduction.GB_RTP.Visible = False
        '    GB_RJetter.Visible = True
        '    FrmProduction.GB_RJetter.Visible = True
        'Else
        '    GB_RTP.Visible = False
        '    FrmProduction.GB_RTP.Visible = False
        '    GB_RJetter.Visible = False
        '    FrmProduction.GB_RJetter.Visible = False
        'End If


        'If IDS.Data.Hardware.Dispenser.Left.AlowEdit Then
        '    BT_Pro_LEdit.Checked = True
        '    BT_Pro_LUpdate.Enabled = True
        '    BT_Pro_LDownload.Enabled = True

        '    TB_Pro_LRPM.ReadOnly = False
        '    TB_Pro_LREV.ReadOnly = False
        '    TB_Pro_LValve.ReadOnly = False
        '    TB_Pro_LProg.ReadOnly = False
        '    TB_Pro_LPressure.ReadOnly = False
        '    TB_Pro_LSuck.ReadOnly = False

        '    FrmProduction.BT_Pro_LEdit.Checked = True
        '    FrmProduction.BT_Pro_LUpdate.Enabled = True
        '    FrmProduction.BT_Pro_LDownload.Enabled = True

        '    FrmProduction.TB_Pro_LRPM.ReadOnly = False
        '    FrmProduction.TB_Pro_LREV.ReadOnly = False
        '    FrmProduction.TB_Pro_LValve.ReadOnly = False
        '    FrmProduction.TB_Pro_LProg.ReadOnly = False
        '    FrmProduction.TB_Pro_LPressure.ReadOnly = False
        '    FrmProduction.TB_Pro_LSuck.ReadOnly = False

        '    BT_LJetter_Open.Enabled = True
        '    TB_LJetter_Path.ReadOnly = False
        '    BT_LJetter_Browse.Enabled = True
        '    BT_LJetter_Edit.Checked = True

        '    FrmProduction.BT_LJetter_Open.Enabled = True
        '    FrmProduction.TB_LJetter_Path.ReadOnly = False
        '    FrmProduction.BT_LJetter_Browse.Enabled = True
        '    FrmProduction.BT_LJetter_Edit.Checked = True
        'Else
        '    BT_Pro_LEdit.Checked = False
        '    BT_Pro_LUpdate.Enabled = False
        '    BT_Pro_LDownload.Enabled = False

        '    TB_Pro_LRPM.ReadOnly = True
        '    TB_Pro_LREV.ReadOnly = True
        '    TB_Pro_LValve.ReadOnly = True
        '    TB_Pro_LProg.ReadOnly = True
        '    TB_Pro_LPressure.ReadOnly = True
        '    TB_Pro_LSuck.ReadOnly = True

        '    FrmProduction.BT_Pro_LEdit.Checked = False
        '    FrmProduction.BT_Pro_LUpdate.Enabled = False
        '    FrmProduction.BT_Pro_LDownload.Enabled = False

        '    FrmProduction.TB_Pro_LRPM.ReadOnly = True
        '    FrmProduction.TB_Pro_LREV.ReadOnly = True
        '    FrmProduction.TB_Pro_LValve.ReadOnly = True
        '    FrmProduction.TB_Pro_LProg.ReadOnly = True
        '    FrmProduction.TB_Pro_LPressure.ReadOnly = True
        '    FrmProduction.TB_Pro_LSuck.ReadOnly = True

        '    BT_LJetter_Open.Enabled = False
        '    TB_LJetter_Path.ReadOnly = True
        '    BT_LJetter_Browse.Enabled = False
        '    BT_LJetter_Edit.Checked = False

        '    FrmProduction.BT_LJetter_Open.Enabled = False
        '    FrmProduction.TB_LJetter_Path.ReadOnly = True
        '    FrmProduction.BT_LJetter_Browse.Enabled = False
        '    FrmProduction.BT_LJetter_Edit.Checked = False
        'End If

        'If IDS.Data.Hardware.Dispenser.Right.AlowEdit Then
        '    BT_Pro_REdit.Checked = True
        '    BT_Pro_RUpdate.Enabled = True
        '    BT_Pro_RDownload.Enabled = True

        '    TB_Pro_RRPM.ReadOnly = False
        '    TB_Pro_RREV.ReadOnly = False
        '    TB_Pro_RValve.ReadOnly = False
        '    TB_Pro_RProg.ReadOnly = False
        '    TB_Pro_RPressure.ReadOnly = False
        '    TB_Pro_RSuck.ReadOnly = False

        '    FrmProduction.BT_Pro_REdit.Checked = True
        '    FrmProduction.BT_Pro_RUpdate.Enabled = True
        '    FrmProduction.BT_Pro_RDownload.Enabled = True

        '    FrmProduction.TB_Pro_RRPM.ReadOnly = False
        '    FrmProduction.TB_Pro_RREV.ReadOnly = False
        '    FrmProduction.TB_Pro_RValve.ReadOnly = False
        '    FrmProduction.TB_Pro_RProg.ReadOnly = False
        '    FrmProduction.TB_Pro_RPressure.ReadOnly = False
        '    FrmProduction.TB_Pro_RSuck.ReadOnly = False

        '    BT_RJetter_Open.Enabled = True
        '    TB_RJetter_Path.ReadOnly = False
        '    BT_RJetter_Browse.Enabled = True
        '    BT_RJetter_Edit.Checked = True

        '    FrmProduction.BT_RJetter_Open.Enabled = True
        '    FrmProduction.TB_RJetter_Path.ReadOnly = False
        '    FrmProduction.BT_RJetter_Browse.Enabled = True
        '    FrmProduction.BT_RJetter_Edit.Checked = True
        'Else
        '    BT_Pro_REdit.Checked = False
        '    BT_Pro_RUpdate.Enabled = False
        '    BT_Pro_RDownload.Enabled = False

        '    TB_Pro_RRPM.ReadOnly = True
        '    TB_Pro_RREV.ReadOnly = True
        '    TB_Pro_RValve.ReadOnly = True
        '    TB_Pro_RProg.ReadOnly = True
        '    TB_Pro_RPressure.ReadOnly = True
        '    TB_Pro_RSuck.ReadOnly = True

        '    FrmProduction.BT_Pro_REdit.Checked = False
        '    FrmProduction.BT_Pro_RUpdate.Enabled = False
        '    FrmProduction.BT_Pro_RDownload.Enabled = False

        '    FrmProduction.TB_Pro_RRPM.ReadOnly = True
        '    FrmProduction.TB_Pro_RREV.ReadOnly = True
        '    FrmProduction.TB_Pro_RValve.ReadOnly = True
        '    FrmProduction.TB_Pro_RProg.ReadOnly = True
        '    FrmProduction.TB_Pro_RPressure.ReadOnly = True
        '    FrmProduction.TB_Pro_RSuck.ReadOnly = True

        '    BT_RJetter_Open.Enabled = False
        '    TB_RJetter_Path.ReadOnly = True
        '    BT_RJetter_Browse.Enabled = False
        '    BT_RJetter_Edit.Checked = False

        '    FrmProduction.BT_RJetter_Open.Enabled = False
        '    FrmProduction.TB_RJetter_Path.ReadOnly = True
        '    FrmProduction.BT_RJetter_Browse.Enabled = False
        '    FrmProduction.BT_RJetter_Edit.Checked = False
        'End If

        'If IDS.Data.Hardware.Dispenser.Left.HeadExtCtrl Then
        '    LB_IDSPressure_LJetter.Visible = False
        '    LB_IDSPressure_LTP.Visible = False
        '    FrmProduction.LB_IDSPressure_LJetter.Visible = False
        '    FrmProduction.LB_IDSPressure_LTP.Visible = False
        'Else
        '    LB_IDSPressure_LJetter.Visible = True
        '    LB_IDSPressure_LTP.Visible = True
        '    FrmProduction.LB_IDSPressure_LJetter.Visible = True
        '    FrmProduction.LB_IDSPressure_LTP.Visible = True
        'End If

        'If IDS.Data.Hardware.Dispenser.Right.HeadExtCtrl Then
        '    LB_IDSPressure_RJetter.Visible = False
        '    LB_IDSPressure_RTP.Visible = False
        '    FrmProduction.LB_IDSPressure_RJetter.Visible = False
        '    FrmProduction.LB_IDSPressure_RTP.Visible = False
        'Else
        '    LB_IDSPressure_RJetter.Visible = True
        '    LB_IDSPressure_RTP.Visible = True
        '    FrmProduction.LB_IDSPressure_RJetter.Visible = True
        '    FrmProduction.LB_IDSPressure_RTP.Visible = True
        'End If
    End Sub
    'xu long 190707 */

    Private Sub BT_RJetter_Open_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BT_RJetter_Open.Click
        Try
            System.Diagnostics.Process.Start(TB_RJetter_Path.Text)
        Catch
        End Try
    End Sub

    Private Sub BT_LJetter_Open_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BT_LJetter_Open.Click
        Try
            System.Diagnostics.Process.Start(TB_LJetter_Path.Text)
        Catch
        End Try
    End Sub

    Private Sub BT_LJetter_Browse_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BT_LJetter_Browse.Click
        If MyDispenser.OpenFileDialog1.ShowDialog = DialogResult.OK Then
            TB_LJetter_Path.Text = MyDispenser.OpenFileDialog1.FileName
        End If
    End Sub

    Private Sub BT_RJetter_Browse_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BT_RJetter_Browse.Click
        If MyDispenser.OpenFileDialog1.ShowDialog = DialogResult.OK Then
            TB_RJetter_Path.Text = MyDispenser.OpenFileDialog1.FileName
        End If
    End Sub

    'Private Sub MenuItem89_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem89.Click
    ' IDS.Devices.DIO.DIO.TowerLight_Red(True)
    'End Sub

    'Private Sub MenuItem90_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem90.Click
    '    IDS.Devices.DIO.DIO.TowerLight_Red(False)
    'End Sub

    'Private Sub MenuItem91_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem91.Click
    '    IDS.Devices.DIO.DIO.TowerLight_Yellow(True)
    'End Sub

    'Private Sub MenuItem92_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem92.Click
    '    IDS.Devices.DIO.DIO.TowerLight_Yellow(False)
    'End Sub

    'Private Sub MenuItem93_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem93.Click
    '    IDS.Devices.DIO.DIO.TowerLight_Green(True)
    'End Sub

    'Private Sub MenuItem94_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem94.Click
    '    IDS.Devices.DIO.DIO.TowerLight_Green(False)
    'End Sub

    'Private Sub MenuItem95_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem95.Click
    '    IDS.Devices.DIO.DIO.TowerLight_Siren(True)
    'End Sub

    'Private Sub MenuItem96_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem96.Click
    '    IDS.Devices.DIO.DIO.TowerLight_Siren(False)
    'End Sub
#End Region

#Region "Jiang"
    Private Sub MenuItem88_Click(ByVal sender As System.Object, _
         ByVal e As System.EventArgs) Handles MenuItem88.Click
        Dim FrmEventViewer As New EventViewer
        FrmEventViewer.Show()
    End Sub

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'FlushSpreadsheet activeX spreadsheet pattern data           
    '   Axspreadsheet:  instance of activeX spreadsheet                                                                             
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Private Sub FlushSpreadsheet()

        Dim TotalSheets As Integer = AxSpreadsheetProgramming.ActiveWorkbook.Worksheets.Count()
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
        m_row = 1
        GC.Collect()
    End Sub



    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'menu: File --> New: Create a new project        
    '                                                                               
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Private Sub MenuFileNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuFileNew.Click
        Dim response As MsgBoxResult
        LabelMessege.Text = " "
        'Handling the case with a projct already exists.
        If "" <> gPatternFileName Then
            response = MsgBox("Do you want to create a new project?", MsgBoxStyle.Question + MsgBoxStyle.YesNo)

            If response = MsgBoxResult.No Then
                Return
            End If
        End If


        'Set defauly data for the DialogBox
        SavePatternFileDialog.InitialDirectory = "C:\IDS\Pattern_Dir"
        SavePatternFileDialog.AddExtension = True
        SavePatternFileDialog.Filter = "Create a new project folder|"
        SavePatternFileDialog.FileName() = ""

        SavePatternFileDialog.ShowDialog()
        Dim file As String
        Dim folder As String = SavePatternFileDialog.FileName()

        Try
            If folder <> Nothing Then

                If -1 <> folder.IndexOf(".") Then
                    'Project name cannot contain "."
                    MsgBox("Invalid project name!, Press Ok to return", _
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

                gFiducialFileName = gPatternFileName


                'Set ToolButton status
                EnableTBRefers(True)
                EnableTBElements(True)
                EnableTBElements(gOffsetCmdIndex, False)
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
                ''''''''''''''''''''''''''''''''''''''''''''''
                ' get default value from the default pat file
                ''''''''''''''''''''''''''''''''''''''''''''''

                IDS.Data.ParameterID.RecordID = "FactoryDefault"
                IDSData.Admin.Folder.FileExtension = "Pat"
                IDSData.Admin.Folder.PatternPath = "C:\IDS\Pattern_Dir"
                IDS.Data.OpenData()
                IDS.Data.SavePathFileData(gPatternFileName + ".pat")
                ''''''''''''''''''''''''''''''''''''''''''''''
                'lim
                IDS.Devices.Vision.SetFiducialFilename(gPatternFileName)

                m_EditStateFlag = False
                m_KeyinOK = False


                ' xu long. This function is in - Setup DLL - Module.vb -
                IniDeviceSetup()
                If IDS.Data.Hardware.Gantry.SystemUnit = False Then 'mm
                    LB_MM.Visible = True
                    LB_INCH.Visible = False
                Else
                    LB_MM.Visible = False
                    LB_INCH.Visible = True
                End If
                Disp_Dispenser_Unit_info()

                SystemSetupDataRetrieve() 'SJ add 


            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Private function to export Excel data file,  to be called by MenuFileExport_Click
    '   file:  filename name for the export                                                                                
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
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

        Dim strUnit As String
        If True = m_Execution.m_Pattern.m_ErrorChk.WorkingUnitFlag Then
            strUnit = "mm"
        Else
            strUnit = "inch"
        End If

        If file <> Nothing Then
            AxSpreadsheetProgramming.Worksheets("Main").Activate()
            AxSpreadsheetProgramming.ActiveSheet.Cells(1, gCmdColumn).Select()
            AxSpreadsheetProgramming.Selection.EntireRow.Insert()
            AxSpreadsheetProgramming.ActiveSheet.Cells(1, gCmdColumn).Value = "Unit"
            AxSpreadsheetProgramming.ActiveSheet.Cells(1, gDispensColumn).Value = strUnit

            AxSpreadsheetProgramming().Export(file, _
                OWC.SheetExportActionEnum.ssExportActionNone, OWC.SheetExportFormat.ssExportAsAppropriate)

            AxSpreadsheetProgramming.ActiveSheet.Rows(1).Delete()

        End If

        Me.Cursor = System.Windows.Forms.Cursors.Default

        Dim strUnitSeting As String = "Undefined"
        Dim filename As String
        Dim usedRange As Integer = AxSpreadsheetProgramming.Worksheets("Main").UsedRange.Rows.Count

        With AxSpreadsheetProgramming.Worksheets("Main")
            For i = 1 To usedRange
                If "SubPattern" = CStr(.Cells(i, gCmdColumn).value) Then
                    filename = CStr(.Cells(i, gSubnameColumn).value)
                    m_Execution.m_Pattern.LoadTxtPatternPara(AxSpreadsheetProgramming, filename, 1, 0, True, strUnitSeting)

                End If
            Next
        End With

        AxSpreadsheetProgramming.Worksheets("Main").Activate()
        GC.Collect()

    End Sub

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Private function to export Pattern data file, to be called by MenuFileExport_Click
    '   file:  filename name for the export                                                                                
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Private Sub ExportTxtPatternFile(ByVal file As String)

        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor

        Dim strUnit As String
        If True = m_Execution.m_Pattern.m_ErrorChk.WorkingUnitFlag Then
            strUnit = "mm"
        Else
            strUnit = "inch"
        End If

        Try
            If file <> Nothing Then
                m_Execution.m_Pattern.SavePatternPara(AxSpreadsheetProgramming, file, False, _
                    strUnit)
            End If
        Catch ex As Exception
        End Try

        Me.Cursor = System.Windows.Forms.Cursors.Default

    End Sub


    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'menu: File-->Export, Private function to export data file
    '   file:  filename name for the export                                                                                
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Private Sub MenuFileExport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuFileExport.Click
        Dim strTmp(4) As String
        LabelMessege.Text = " "
        'should get from data mamager 
        SavePatternFileDialog.InitialDirectory = "C:\IDS\Pattern_Dir"
        SavePatternFileDialog.AddExtension = True
        SavePatternFileDialog.Filter = "Excel Pattern files (*.Xls)|*.Xls|Txt Pattern files (*.ptp)|*.ptp"

        SavePatternFileDialog.ShowDialog()
        Dim file As String = SavePatternFileDialog.FileName()

        If file <> Nothing Then
            strTmp = file.Split(".")
            If "Xls" = strTmp(1) Then
                ExportExcelPatternFile(file)
            ElseIf "ptp" = strTmp(1) Then
                ExportTxtPatternFile(file)
            End If
        End If

        'AxSpreadsheetProgramming.Export(strTmp(0) + "." + "Xls", _
        '    OWC.SheetExportActionEnum.ssExportActionNone, OWC.SheetExportFormat.ssExportAsAppropriate)
    End Sub


    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Import Excel pattern file          
    '   file:  input filename                                                                             
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Private Sub ImportExcelPatternFile(ByVal file As String)    'Internal function
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor
        gPatternFileName = ""
        gFiducialFileName = ""

        Dim strUnit As String
        If True = m_Execution.m_Pattern.m_ErrorChk.WorkingUnitFlag Then
            strUnit = "mm"
        Else
            strUnit = "inch"
        End If

        Try
            AxSpreadsheetProgramming.Caption = file
            m_Execution.m_Pattern.LoadExcelPatternFile(AxSpreadsheetProgramming, file, 1, False, strUnit)
            m_row = 2

            'Error checking for all the Spreadsheet
            If 0 <> m_Execution.m_Pattern.m_ErrorChk.CheckAllError(AxSpreadsheetProgramming, ErrorSubSheet) Then
                MsgBox("Click Ok to flush all the data", MsgBoxStyle.OKOnly, "Error found")

                FlushSpreadsheet()
                gPatternFileName = ""
                gFiducialFileName = ""
                init_spreadsheet()
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

        Me.Cursor = System.Windows.Forms.Cursors.Default
        GC.Collect()
    End Sub


    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Import text pattern file          
    '   file:  input filename                                                                             
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Private Sub ImportTxtPatternFile(ByVal file As String)
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor
        gPatternFileName = ""
        gFiducialFileName = ""

        Dim strUnit As String
        If True = m_Execution.m_Pattern.m_ErrorChk.WorkingUnitFlag Then
            strUnit = "mm"
        Else
            strUnit = "inch"
        End If

        Try
            AxSpreadsheetProgramming.Caption = file
            If 0 = m_Execution.m_Pattern.LoadTxtPatternPara(AxSpreadsheetProgramming, file, 2, m_row, False, strUnit) Then
                m_row = 2

                'Error checking for all the Spreadsheet
                If 0 <> m_Execution.m_Pattern.m_ErrorChk.CheckAllError(AxSpreadsheetProgramming, ErrorSubSheet) Then
                    MsgBox("Click Ok to flush all the data", MsgBoxStyle.OKOnly, "Error found")

                    FlushSpreadsheet()
                    gPatternFileName = ""
                    gFiducialFileName = ""
                    init_spreadsheet()
                End If
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
        Me.Cursor = System.Windows.Forms.Cursors.Default
        GC.Collect()
    End Sub


    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'menu: File-->Import (from an excel/text file)
    '   file:                                                                                  
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Private Sub MenuFileImport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuFileImport.Click
        LabelMessege.Text = " "
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

        'FlushSpreadsheet()
        'init_spreadsheet()

        If "Xls" = strInfo(1) Then
            ImportExcelPatternFile(file)
        ElseIf "ptp" = strInfo(1) Then
            ImportTxtPatternFile(file)
        End If
        GC.Collect()
    End Sub


    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Load an encrypt text pattern file for production //SJ add
    '   
    '   Return: 0=success
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Public Function ProductionFileOpen() As Integer
        Dim ErrorMsg As String
        Dim Rtn As MsgBoxResult
        Dim strUnitSeting As String = "Undefined"

        Dim DialogPreview As New FormSelectPatternFile

        DialogPreview.ShowDialog()
        Dim file = DialogPreview.Path

        If Nothing = file Then
            FrmProduction.TextBoxFilename.Text = ""
            Return -1
        End If


        FlushSpreadsheet()
        gPatternFileName = ""
        gFiducialFileName = ""
        init_spreadsheet()
        m_row = 0
        m_EditStateFlag = False
        m_KeyinOK = False

        Try
            AxSpreadsheetProgramming.Caption = file
            m_Execution.m_Pattern.LoadTxtPatternPara(AxSpreadsheetProgramming, file, 0, 0, False, strUnitSeting) ' True)
            gPatternFileName = m_Execution.m_File.FolderWithNameFromFileName(file)
            gFiducialFileName = gPatternFileName
            m_row = 2

            IDS.Data.OpenPathFileData(gPatternFileName + ".pat")
            'IDS.Data.OpenData() ' xu long
            IniDeviceSetup() ' xu long. This function is in - Setup DLL - Module.vb -
            SystemSetupDataRetrieve() 'SJ add 
            'Example for change Unit setup,'1=mm; 0=inch
            m_Execution.m_Pattern.m_ErrorChk.SystemUnitSet(1, AxSpreadsheetProgramming, False)

            AxSpreadsheetProgramming.Worksheets("Main").Activate()
            'Error checking for all the Spreadsheet
            If 0 <> m_Execution.m_Pattern.m_ErrorChk.CheckAllError(AxSpreadsheetProgramming, ErrorSubSheet) Then

                Rtn = MsgBox("Click Ok to flush all the data", MsgBoxStyle.Question + MsgBoxStyle.YesNo, "Error found, flush all the data or not?")
                If MsgBoxResult.Yes = Rtn Then
                    FlushSpreadsheet()
                    gPatternFileName = ""
                    gFiducialFileName = ""
                    init_spreadsheet()
                    FrmProduction.TextBoxFilename.Text = ""
                    Return -1
                Else
                    FlushSpreadsheet()
                    gPatternFileName = ""
                    gFiducialFileName = ""
                    init_spreadsheet()
                    FrmProduction.TextBoxFilename.Text = ""
                    Return -1
                End If
            Else

            End If

            FrmProduction.TextBoxFilename.Text = file

            'set gd settings 'lim
            If gPatternFileName = Nothing Then
            Else
                IDS.Data.OpenData()
                FrmGD.LoadGDFile(gPatternFileName, IDS.Data.Hardware.Gantry.GraphicDisplayXYLT.X, IDS.Data.Hardware.Gantry.GraphicDisplayXYLT.Y)
                GDRefX = IDS.Data.Hardware.Gantry.GraphicDisplayXYLT.X
                GDRefY = IDS.Data.Hardware.Gantry.GraphicDisplayXYLT.Y

                'lim
                IDS.Devices.Vision.SetFiducialFilename(gPatternFileName)
            End If

        Catch ex As Exception
            MessageBox.Show("Load pattern data unknown error", "Error information", MessageBoxButtons.OK)

        End Try
        GC.Collect()
        Return 0
    End Function


    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'menu: File-->Load an encrypt text pattern file
    '   
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Private Sub MenuFileOpen_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuFileOpen.Click
        Dim ErrorMsg As String
        Dim Rtn As MsgBoxResult
        LabelMessege.Text = " "
        Dim DialogPreview As New FormSelectPatternFile

        DialogPreview.ShowDialog()
        Dim file = DialogPreview.Path


        If Nothing = file Then
            Return
        End If

        'AxSpreadsheetProgramming.ActiveSheet.UsedRange().Clear()
        FlushSpreadsheet()
        gPatternFileName = ""
        gFiducialFileName = ""
        init_spreadsheet()
        m_row = 0
        m_EditStateFlag = False
        m_KeyinOK = False

        Dim sysUnit As String = "Undefined"

        Try
            AxSpreadsheetProgramming.Caption = file
            m_Execution.m_Pattern.LoadTxtPatternPara(AxSpreadsheetProgramming, file, 0, 0, False, sysUnit)
            gPatternFileName = m_Execution.m_File.FolderWithNameFromFileName(file)
            gFiducialFileName = gPatternFileName
            m_row = 2
            EnableTBRefers(True)
            EnableTBElements(True)

            EnableTBElements(gOffsetCmdIndex, False)

            IDS.Data.OpenPathFileData(gPatternFileName + ".pat")
            'IDS.Data.OpenData() ' xu long.
            IniDeviceSetup() ' xu long. This function is in - Setup DLL - Module.vb -

            SystemSetupDataRetrieve() 'SJ add 

            'Example for change Unit setup,'1=mm; 0=inch
            m_Execution.m_Pattern.m_ErrorChk.SystemUnitSet(1, AxSpreadsheetProgramming, False)

            'Acitvate the "Main" page
            AxSpreadsheetProgramming.Worksheets("Main").Activate()

            'Error checking for all the Spreadsheet
            If 0 <> m_Execution.m_Pattern.m_ErrorChk.CheckAllError(AxSpreadsheetProgramming, ErrorSubSheet) Then

                Rtn = MsgBox("Click Ok to flush all the data", MsgBoxStyle.Question + MsgBoxStyle.YesNo, "Error found, flush all the data or not?")
                If MsgBoxResult.Yes = Rtn Then
                    FlushSpreadsheet()
                    gPatternFileName = ""
                    gFiducialFileName = ""
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
                FrmGD.LoadGDFile(gPatternFileName, IDS.Data.Hardware.Gantry.GraphicDisplayXYLT.X, IDS.Data.Hardware.Gantry.GraphicDisplayXYLT.Y)
                GDRefX = IDS.Data.Hardware.Gantry.GraphicDisplayXYLT.X
                GDRefY = IDS.Data.Hardware.Gantry.GraphicDisplayXYLT.Y

                'lim
                IDS.Devices.Vision.SetFiducialFilename(gPatternFileName)
            End If


        Catch ex As Exception
            MessageBox.Show("Loading data unknown error", "Error information", MessageBoxButtons.OK)

            MenuFileExport.Enabled = True
            MenuFileImport.Enabled = True
            MenuFileSave.Enabled = True
            MenuFileSaveAs.Enabled = True
        End Try
        PanelGraphDisp.Show()
        PanelGraphDisp.Controls.Add(FrmGD.Panel1)
        FrmGD.Panel1.Location = New Point(0, 0)
        PanelStepper.Hide()
        PanelDispenser.Hide()
        PanelThermal.Hide()
        PanelConveyor.Hide()
        TBMultiDisplayField.Buttons(0).Pushed = True
        TBMultiDisplayField.Buttons(1).Pushed = False
        TBMultiDisplayField.Buttons(2).Pushed = False
        TBMultiDisplayField.Buttons(3).Pushed = False
        TBMultiDisplayField.Buttons(4).Pushed = False

        If IDS.Data.Hardware.Gantry.SystemUnit = False Then 'mm
            LB_MM.Visible = True
            LB_INCH.Visible = False
        Else
            LB_MM.Visible = False
            LB_INCH.Visible = True
        End If

        Disp_Dispenser_Unit_info()

        GC.Collect()
    End Sub

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'menu: File-->Save encrypt text pattern file
    '   
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Private Sub MenuFileSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuFileSave.Click
        LabelMessege.Text = " "
        Dim strUnit As String
        If True = m_Execution.m_Pattern.m_ErrorChk.WorkingUnitFlag Then
            strUnit = "mm"
        Else
            strUnit = "inch"
        End If

        Try
            If gPatternFileName <> Nothing Then
                m_Execution.m_Pattern.SavePatternPara(AxSpreadsheetProgramming, gPatternFileName + ".txt", False, _
                    strUnit)
                FrmGD.SaveGDFile(gPatternFileName) 'save GD 'lim
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
                            MsgBox("Invalid project name!, Press Ok to return", _
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
                        gFiducialFileName = gPatternFileName
                        m_Execution.m_Pattern.SavePatternPara(AxSpreadsheetProgramming, gPatternFileName + ".txt", False, _
                            strUnit)
                        'lim
                        IDS.Devices.Vision.SetFiducialFilename(gPatternFileName)
                    End If
                End If
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
        GC.Collect()
    End Sub


    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'menu: File-->Save as encrypt text pattern file
    '   file:  filename                                                                                
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Private Sub MenuFileSaveAs_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuFileSaveAs.Click
        LabelMessege.Text = " "
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
                    MsgBox("Invalid project name!, Press Ok to return", _
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

                    System.IO.File.Copy(gPatternFileName_OldPath & ".bmp", gPatternFileName & ".bmp") 'GD
                    System.IO.File.Copy(gPatternFileName_OldPath & "1.mmo", gPatternFileName & "1.mmo") 'Fiducial File 1
                    System.IO.File.Copy(gPatternFileName_OldPath & "1.bmp", gPatternFileName & "1.bmp") 'Fiducial 1 image
                    If System.IO.File.Exists(gPatternFileName_OldPath & "2.mmo") Then 'if second fiducial exist
                        System.IO.File.Copy(gPatternFileName_OldPath & "2.mmo", gPatternFileName & "2.mmo") 'Fiducial File 2
                        System.IO.File.Copy(gPatternFileName_OldPath & "2.bmp", gPatternFileName & "2.bmp") 'Fiducial 2 image
                    End If
                End If

                gFiducialFileName = gPatternFileName

                Dim strUnit As String
                If True = m_Execution.m_Pattern.m_ErrorChk.WorkingUnitFlag Then
                    strUnit = "mm"
                Else
                    strUnit = "inch"
                End If

                m_Execution.m_Pattern.SavePatternPara(AxSpreadsheetProgramming, gPatternFileName + ".txt", _
                    False, strUnit)

                IDS.Data.SavePathFileData(gPatternFileName + ".pat")
                IDS.Devices.Vision.SetFiducialFilename(gPatternFileName)
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub


    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'menu: File-->Exit program
    '                                                                                  
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Private Sub MenuFileExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuFileExit.Click
        FormProgrammingClosed_Implementaton()
        Close()
    End Sub




    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Logging data for Undo operation
    '   UndoLevel: 0 or 1; 
    '               0 is slow, used for multiRow, sub or array
    '               1 is fast, used for single row without sub and array            
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Private Sub UndoData_Logging(ByVal UndoLevel As Integer)
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
            m_Execution.m_Undo.CurrentPageName_A = AxSpreadsheetProgramming.ActiveSheet.Name
            m_Execution.m_Undo.DataSaveFor_Undo(AxSpreadsheetProgramming)
        Else
            Dim sel As Microsoft.Office.Interop.OWC.Range = AxSpreadsheetProgramming.Selection

            Dim m_rowCount As Integer = sel.Rows.Count()
            Dim m_columnCount As Integer = sel.Columns.Count()
            Dim m_StartRow As Integer = sel.Row
            Dim m_columnStart As Integer = sel.Column

            m_Execution.m_Undo.UndoRow = m_StartRow

            Dim SheetName As String = AxSpreadsheetProgramming.ActiveSheet.Name
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
        GC.Collect()
    End Sub

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'menu: Edit-->Redo                                                                           
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Private Sub MenuEditRedo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuEditRedo.Click

        If 0 = m_Execution.m_Undo.UndoLevel Then
            m_Execution.m_Undo.UndoFilename = "C:\IDS\Pattern_Dir\SysSwapData\UndoStep_A.Xls"
            m_Execution.m_Undo.CurrentPageName_A = AxSpreadsheetProgramming.ActiveSheet.Name
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

            Dim SheetName As String = AxSpreadsheetProgramming.ActiveSheet.Name
            If 1 = m_rowCount Then
                m_Execution.m_Pattern.m_ErrorChk.ConvertToDataStruct(m_Execution.m_Undo.UndoPatternRec_A, _
                    AxSpreadsheetProgramming, SheetName, m_Execution.m_Undo.UndoRow)

                m_Execution.m_Pattern.m_ErrorChk.ConvertFromDataStruct(m_Execution.m_Undo.UndoPatternRec_B, _
                    AxSpreadsheetProgramming, m_Execution.m_Undo.UndoRow)
            End If
        End If

        MenuEditUndo.Enabled = True
        MenuEditRedo.Enabled = False
        GC.Collect()
    End Sub


    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'menu: Edit-->Undo                                                                            
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Private Sub MenuEditUndo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuEditUndo.Click

        If 0 = m_Execution.m_Undo.UndoLevel Then
            m_Execution.m_Undo.UndoFilename = "C:\IDS\Pattern_Dir\SysSwapData\UndoStep_B.Xls"
            m_Execution.m_Undo.CurrentPageName_B = AxSpreadsheetProgramming.ActiveSheet.Name
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

            Dim SheetName As String = AxSpreadsheetProgramming.ActiveSheet.Name
            If 1 = m_rowCount Then
                m_Execution.m_Pattern.m_ErrorChk.ConvertToDataStruct(m_Execution.m_Undo.UndoPatternRec_B, _
                    AxSpreadsheetProgramming, SheetName, m_Execution.m_Undo.UndoRow)

                m_Execution.m_Pattern.m_ErrorChk.ConvertFromDataStruct(m_Execution.m_Undo.UndoPatternRec_A, _
                    AxSpreadsheetProgramming, m_Execution.m_Undo.UndoRow)
            End If
        End If

        MenuEditUndo.Enabled = False
        MenuEditRedo.Enabled = True
        GC.Collect()
    End Sub


    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'menu: Edit-->Cut: Cut the selected range with Array/Sub attached                                                                            
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Private Sub MenuEditCut_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuEditCut.Click
        Dim sel As Microsoft.Office.Interop.OWC.Range = AxSpreadsheetProgramming.Selection
        Dim m_rowCount As Integer = sel.Rows.Count()
        Dim m_columnCount As Integer = sel.Columns.Count()
        Dim m_StartRow As Integer = sel.Row
        Dim m_columnStart As Integer = sel.Column

        Dim SheetName As String = AxSpreadsheetProgramming.ActiveSheet.Name
        If m_columnCount = gMaxColumns Then
            CopiedSheetName = AxSpreadsheetProgramming.ActiveSheet.Name

            'Copy selected data and also with the related Array/Sub data
            m_Execution.m_Pattern.SavePatternParaAllSheets(AxSpreadsheetProgramming, _
                "C:\IDS\Pattern_Dir\SysSwapData\CopyPaste.txt", 1, False)

            ConfirmOrCancel_Implementation(0)
            MenuEditPaste.Enabled = True
        Else
            MsgBox("Cut is not allowed inside LINK range", MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly, "Cut is not alowed")
        End If
        GC.Collect()
    End Sub

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'menu: Edit-->Copy: Copy the selected range with Array/Sub attached                                                                            
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Private Sub MenuEditCopy_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuEditCopy.Click
        Dim sel As Microsoft.Office.Interop.OWC.Range = AxSpreadsheetProgramming.Selection

        Dim m_rowCount As Integer = sel.Rows.Count()
        Dim m_columnCount As Integer = sel.Columns.Count()
        Dim m_StartRow As Integer = sel.Row
        Dim m_columnStart As Integer = sel.Column

        If m_columnCount = gMaxColumns Then
            CopiedSheetName = AxSpreadsheetProgramming.ActiveSheet.Name

            'Copy selected data and also with the related Array/Sub data
            m_Execution.m_Pattern.SavePatternParaAllSheets(AxSpreadsheetProgramming, _
                "C:\IDS\Pattern_Dir\SysSwapData\CopyPaste.txt", 1, False)
            MenuEditPaste.Enabled = True
        End If
        GC.Collect()
    End Sub

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'menu: Edit-->Paste: Paste the Copied range with Array/Sub attached                                                                            
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Private Sub MenuEditPaste_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuEditPaste.Click
        Dim sel As Microsoft.Office.Interop.OWC.Range = AxSpreadsheetProgramming.Selection

        Dim m_rowCount As Integer = sel.Rows.Count()
        Dim m_StartRow As Integer = sel.Row

        If CopiedSheetName <> AxSpreadsheetProgramming.ActiveSheet.Name Then
            MsgBox("Cannot be pasted in this Spreadsheet!", MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly, "Copy/Paste in Different Sheet")
            Return
        End If

        If "" = AxSpreadsheetProgramming.ActiveSheet.Name Then
            MsgBox("Copy contents cannot be used!", MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly, "Contents are invalid")
            Return
        End If

        If 1 = m_rowCount Then
            AxSpreadsheetProgramming_CheckForWithinLinkRange(False)
            If m_InLinkRange Then
                MsgBox("Paste is not allowed inside LINK range", MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly, "Paste is not alowed")
            Else
                UndoData_Logging(0)

                m_Execution.m_Pattern.LoadTxtPatternParaAllSheets(AxSpreadsheetProgramming, _
                    "C:\IDS\Pattern_Dir\SysSwapData\CopyPaste.txt", _
                    CopiedSheetName, 2, m_StartRow, False)
            End If
        End If
        GC.Collect()
    End Sub


    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'menu: Edit-->Delete: Delecte all select                                                                            
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Private Sub MenuEditDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuEditDelete.Click
        ConfirmOrCancel_Implementation(0)
        GC.Collect()
    End Sub


    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'menu: Edit-->SelectAll: Select all in the current spreadsheet                                                                            
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Private Sub MenuEditSelectAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuEditSelectAll.Click
        AxSpreadsheetProgramming.ActiveSheet.UsedRange.Select()
    End Sub


    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'ChangeUnitInSpredsheet: Change unit setting
    '   mmUnit:  unit flag, mm=True, inch=False                                                                                
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Public Sub ChangeUnitInSpredsheet(ByVal mmUnit As Boolean)
        m_Execution.m_Pattern.m_ErrorChk.SystemUnitSet(mmUnit, AxSpreadsheetProgramming, True)
        MenuEditPaste.Enabled = False
        CopiedSheetName = ""
        MenuEditUndo.Enabled = False
        MenuEditRedo.Enabled = False

    End Sub


    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'menu: Command, the following menu commands are self-explainatory
    '      Further information is not necessary                                                                            
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Private Sub MenuCommandsReference_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuCommandsReference.Click
        TBrefer_Implementation("Reference")
    End Sub

    Private Sub MenuCommandsFiducial_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuCommandsFiducial.Click
        TBrefer_Implementation("Fiducial")
    End Sub

    Private Sub MenuCommandsHeight_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuCommandsHeight.Click
        TBrefer_Implementation("Height")
    End Sub

    Private Sub MenuCommandsReject_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuCommandsReject.Click
        TBrefer_Implementation("Reject")
    End Sub



    Private Sub MenuCommandsDot_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuCommandsDot.Click
        TBElements_Implementation("Dot")
    End Sub

    Private Sub MenuCommandsLine_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuCommandsLine.Click
        TBElements_Implementation("Line")
    End Sub

    Private Sub MenuCommandsArc_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuCommandsArc.Click
        TBElements_Implementation("Arc")
    End Sub

    Private Sub MenuCommandsRectangle_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuCommandsRectangle.Click
        TBElements_Implementation("Rectangle")
    End Sub

    Private Sub MenuCommandsFillRectangle_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuCommandsFillRectangle.Click
        TBElements_Implementation("FillRectangle")
    End Sub

    Private Sub MenuCommandsCircle_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuCommandsCircle.Click
        TBElements_Implementation("Circle")
    End Sub

    Private Sub MenuComandsFillCircle_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuCommandsFillCircle.Click
        TBElements_Implementation("FillCircle")
    End Sub

    Private Sub MenuCommandsLink_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuCommandsLink.Click
        TBElements_Implementation("Link")
    End Sub

    Private Sub MenuCommandsChipEdge_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuCommandsChipEdge.Click
        TBElements_Implementation("ChipEdge")
    End Sub

    Private Sub MenuCommandsPurge_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuCommandsPurge.Click
        TBElements_Implementation("Purge")
    End Sub

    Private Sub MenuCommandsClean_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuCommandsClean.Click
        TBElements_Implementation("Clean")
    End Sub

    Private Sub MenuCommandsQC_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuCommandsQC.Click
        TBElements_Implementation("QC")
    End Sub

    Private Sub MenuCommandsArray_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuCommandsArray.Click
        TBElements_Implementation("Array")
    End Sub

    Private Sub MenuCommandsSubPattern_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuCommandsSubPattern.Click
        TBElements_Implementation("SubPattern")
    End Sub

    Private Sub MenuCommandsMove_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuCommandsMove.Click
        TBElements_Implementation("Move")
    End Sub

    Private Sub MenuCommandsWait_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuCommandsWait.Click
        TBElements_Implementation("Wait")
    End Sub

    Private Sub MenuCommandsSetIO_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuCommandsSetIO.Click
        TBElements_Implementation("SetIO")
    End Sub

    Private Sub MenuCommandsGetIO_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuCommandsGetIO.Click
        TBElements_Implementation("GetIO")
    End Sub

    Private Sub MenuCommandsOffset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuCommandsOffset.Click
        TBElements_Implementation("Offset")
    End Sub

    Private Sub MenuCommandsMeasurement_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuCommandsMeasurement.Click
        TBElements_Implementation("Measurement")
    End Sub


    Private Sub MenuCommandsEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuCommandsEdit.Click
        TBNextEdit_Implementation(1)
    End Sub




    Private Sub MenuItem23_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem23.Click
        AxSpreadsheetProgramming.Height = 912
        CBExpandSpreadsheet.Checked = True
    End Sub

    Private Sub MenuItem24_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem24.Click
        AxSpreadsheetProgramming.Height = 269
        CBExpandSpreadsheet.Checked = False
    End Sub

    Public Sub Disp_Dispenser_Unit_info()
        If IDS.Data.Hardware.Dispenser.Left.HeadEnable Then
            If IDS.Data.Hardware.Dispenser.Left.HeadType = "Time Pressure" Then
                LB_Pro_LPressure.Visible = True
                LB_Pro_LSuck.Visible = True
                LB_Pro_LRPM.Visible = False
                LB_Pro_LREV.Visible = False
                BT_Pro_LDownload.Visible = True
                'Bt_pro_LOpenJetterSoftware.Visible = False
                'BTLeftBrowseJetter.Visible = False
                'TBLeftJetterFile.Visible = False
                'LabelLeftJetter.Visible = False

                TB_Pro_LPressure.Visible = True
                TB_Pro_LSuck.Visible = True
                TB_Pro_LRPM.Visible = False
                TB_Pro_LREV.Visible = False

                LB_Pro_LValve.Visible = False
                LB_Pro_LProg.Visible = False
                TB_Pro_LValve.Visible = False
                TB_Pro_LProg.Visible = False

                FrmProduction.LB_Pro_LPressure.Visible = True
                FrmProduction.TB_Pro_LPressure.Visible = True
                FrmProduction.LB_Pro_LSuck.Visible = True
                FrmProduction.TB_Pro_LSuck.Visible = True
                FrmProduction.LB_Pro_LRPM.Visible = False
                FrmProduction.TB_Pro_LRPM.Visible = False
                FrmProduction.LB_Pro_LREV.Visible = False
                FrmProduction.TB_Pro_LREV.Visible = False
                FrmProduction.LB_Pro_LValve.Visible = False
                FrmProduction.TB_Pro_LValve.Visible = False
                FrmProduction.LB_Pro_LProg.Visible = False
                FrmProduction.TB_Pro_LProg.Visible = False

                FrmProduction.BT_Pro_LDownload.Visible = True

                GB_LTP.Visible = True
                GB_LJetter.Visible = False
                FrmProduction.GB_LJetter.Visible = False
                FrmProduction.GB_LTP.Visible = True
            ElseIf IDS.Data.Hardware.Dispenser.Left.HeadType = "Positive Displacement" Then
                LB_Pro_LPressure.Visible = True
                LB_Pro_LSuck.Visible = True
                LB_Pro_LRPM.Visible = True
                LB_Pro_LREV.Visible = True
                BT_Pro_LDownload.Visible = True
                'Bt_pro_LOpenJetterSoftware.Visible = False
                'BTLeftBrowseJetter.Visible = False
                'TBLeftJetterFile.Visible = False
                'LabelLeftJetter.Visible = False

                TB_Pro_LPressure.Visible = True
                TB_Pro_LSuck.Visible = True
                TB_Pro_LRPM.Visible = True
                TB_Pro_LREV.Visible = True

                LB_Pro_LValve.Visible = True
                LB_Pro_LProg.Visible = True
                TB_Pro_LValve.Visible = True
                TB_Pro_LProg.Visible = True

                FrmProduction.LB_Pro_LPressure.Visible = True
                FrmProduction.TB_Pro_LPressure.Visible = True
                FrmProduction.LB_Pro_LSuck.Visible = True
                FrmProduction.TB_Pro_LSuck.Visible = True
                FrmProduction.LB_Pro_LRPM.Visible = True
                FrmProduction.TB_Pro_LRPM.Visible = True
                FrmProduction.LB_Pro_LREV.Visible = True
                FrmProduction.TB_Pro_LREV.Visible = True
                FrmProduction.LB_Pro_LValve.Visible = True
                FrmProduction.TB_Pro_LValve.Visible = True
                FrmProduction.LB_Pro_LProg.Visible = True
                FrmProduction.TB_Pro_LProg.Visible = True

                FrmProduction.BT_Pro_LDownload.Visible = True

                GB_LTP.Visible = True
                GB_LJetter.Visible = False
                FrmProduction.GB_LJetter.Visible = False
                FrmProduction.GB_LTP.Visible = True
            ElseIf IDS.Data.Hardware.Dispenser.Left.HeadType = "Fixed Volumetric" Then
                LB_Pro_LPressure.Visible = True
                LB_Pro_LSuck.Visible = True
                LB_Pro_LRPM.Visible = False
                LB_Pro_LREV.Visible = False
                BT_Pro_LDownload.Visible = True
                'Bt_pro_LOpenJetterSoftware.Visible = False
                'BTLeftBrowseJetter.Visible = False
                'TBLeftJetterFile.Visible = False
                'LabelLeftJetter.Visible = False

                TB_Pro_LPressure.Visible = True
                TB_Pro_LSuck.Visible = True
                TB_Pro_LRPM.Visible = False
                TB_Pro_LREV.Visible = False

                LB_Pro_LValve.Visible = False
                LB_Pro_LProg.Visible = False
                TB_Pro_LValve.Visible = False
                TB_Pro_LProg.Visible = False

                FrmProduction.LB_Pro_LPressure.Visible = True
                FrmProduction.TB_Pro_LPressure.Visible = True
                FrmProduction.LB_Pro_LSuck.Visible = True
                FrmProduction.TB_Pro_LSuck.Visible = True
                FrmProduction.LB_Pro_LRPM.Visible = False
                FrmProduction.TB_Pro_LRPM.Visible = False
                FrmProduction.LB_Pro_LREV.Visible = False
                FrmProduction.TB_Pro_LREV.Visible = False
                FrmProduction.LB_Pro_LValve.Visible = False
                FrmProduction.TB_Pro_LValve.Visible = False
                FrmProduction.LB_Pro_LProg.Visible = False
                FrmProduction.TB_Pro_LProg.Visible = False

                FrmProduction.BT_Pro_LDownload.Visible = True


                GB_LTP.Visible = True
                GB_LJetter.Visible = False
                FrmProduction.GB_LJetter.Visible = False
                FrmProduction.GB_LTP.Visible = True
            ElseIf IDS.Data.Hardware.Dispenser.Left.HeadType = "Jetter" Then
                LB_Pro_LPressure.Visible = False
                LB_Pro_LSuck.Visible = False
                LB_Pro_LRPM.Visible = False
                LB_Pro_LREV.Visible = False
                BT_Pro_LDownload.Visible = False
                'Bt_pro_LOpenJetterSoftware.Visible = False
                'BTLeftBrowseJetter.Visible = False
                'TBLeftJetterFile.Visible = False
                'LabelLeftJetter.Visible = False

                TB_Pro_LPressure.Visible = False
                TB_Pro_LSuck.Visible = False
                TB_Pro_LRPM.Visible = False
                TB_Pro_LREV.Visible = False

                LB_Pro_LValve.Visible = False
                LB_Pro_LProg.Visible = False
                TB_Pro_LValve.Visible = False
                TB_Pro_LProg.Visible = False

                FrmProduction.LB_Pro_LPressure.Visible = False
                FrmProduction.TB_Pro_LPressure.Visible = False
                FrmProduction.LB_Pro_LSuck.Visible = False
                FrmProduction.TB_Pro_LSuck.Visible = False
                FrmProduction.LB_Pro_LRPM.Visible = False
                FrmProduction.TB_Pro_LRPM.Visible = False
                FrmProduction.LB_Pro_LREV.Visible = False
                FrmProduction.TB_Pro_LREV.Visible = False
                FrmProduction.LB_Pro_LValve.Visible = False
                FrmProduction.TB_Pro_LValve.Visible = False
                FrmProduction.LB_Pro_LProg.Visible = False
                FrmProduction.TB_Pro_LProg.Visible = False

                FrmProduction.BT_Pro_LDownload.Visible = False


                GB_LTP.Visible = False
                GB_LJetter.Visible = True
                FrmProduction.GB_LJetter.Visible = True
                FrmProduction.GB_LTP.Visible = False
            ElseIf IDS.Data.Hardware.Dispenser.Left.HeadType = "Others" Then
                LB_Pro_LPressure.Visible = False
                LB_Pro_LSuck.Visible = False
                LB_Pro_LRPM.Visible = False
                LB_Pro_LREV.Visible = False
                BT_Pro_LDownload.Visible = False
                'Bt_pro_LOpenJetterSoftware.Visible = False
                'BTLeftBrowseJetter.Visible = False
                'TBLeftJetterFile.Visible = False
                'LabelLeftJetter.Visible = False

                TB_Pro_LPressure.Visible = False
                TB_Pro_LSuck.Visible = False
                TB_Pro_LRPM.Visible = False
                TB_Pro_LREV.Visible = False

                LB_Pro_LValve.Visible = False
                LB_Pro_LProg.Visible = False
                TB_Pro_LValve.Visible = False
                TB_Pro_LProg.Visible = False


                FrmProduction.LB_Pro_LPressure.Visible = False
                FrmProduction.TB_Pro_LPressure.Visible = False
                FrmProduction.LB_Pro_LSuck.Visible = False
                FrmProduction.TB_Pro_LSuck.Visible = False
                FrmProduction.LB_Pro_LRPM.Visible = False
                FrmProduction.TB_Pro_LRPM.Visible = False
                FrmProduction.LB_Pro_LREV.Visible = False
                FrmProduction.TB_Pro_LREV.Visible = False
                FrmProduction.LB_Pro_LValve.Visible = False
                FrmProduction.TB_Pro_LValve.Visible = False
                FrmProduction.LB_Pro_LProg.Visible = False
                FrmProduction.TB_Pro_LProg.Visible = False

                FrmProduction.BT_Pro_LDownload.Visible = False

                GB_LTP.Visible = True
                GB_LJetter.Visible = False
                FrmProduction.GB_LJetter.Visible = False
                FrmProduction.GB_LTP.Visible = True
            End If

        Else

            GB_LTP.Visible = False
            GB_LJetter.Visible = False
            FrmProduction.GB_LJetter.Visible = False
            FrmProduction.GB_LTP.Visible = False
        End If

        If IDS.Data.Hardware.Dispenser.Right.HeadEnable Then
            If IDS.Data.Hardware.Dispenser.Right.HeadType = "Time Pressure" Then
                LB_Pro_RPressure.Visible = True
                LB_Pro_RSuck.Visible = True
                LB_Pro_RRPM.Visible = False
                LB_Pro_RREV.Visible = False
                BT_Pro_RDownload.Visible = True
                'Bt_pro_LOpenJetterSoftware.Visible = False
                'BTLeftBrowseJetter.Visible = False
                'TBLeftJetterFile.Visible = False
                'LabelLeftJetter.Visible = False

                TB_Pro_RPressure.Visible = True
                TB_Pro_RSuck.Visible = True
                TB_Pro_RRPM.Visible = False
                TB_Pro_RREV.Visible = False

                LB_Pro_RValve.Visible = False
                LB_Pro_RProg.Visible = False
                TB_Pro_RValve.Visible = False
                TB_Pro_RProg.Visible = False

                FrmProduction.LB_Pro_RPressure.Visible = True
                FrmProduction.TB_Pro_RPressure.Visible = True
                FrmProduction.LB_Pro_RSuck.Visible = True
                FrmProduction.TB_Pro_RSuck.Visible = True
                FrmProduction.LB_Pro_RRPM.Visible = False
                FrmProduction.TB_Pro_RRPM.Visible = False
                FrmProduction.LB_Pro_RREV.Visible = False
                FrmProduction.TB_Pro_RREV.Visible = False
                FrmProduction.LB_Pro_RValve.Visible = False
                FrmProduction.TB_Pro_RValve.Visible = False
                FrmProduction.LB_Pro_RProg.Visible = False
                FrmProduction.TB_Pro_RProg.Visible = False

                FrmProduction.BT_Pro_RDownload.Visible = True


                GB_RTP.Visible = True
                GB_RJetter.Visible = False
                FrmProduction.GB_RJetter.Visible = False
                FrmProduction.GB_RTP.Visible = True
            ElseIf IDS.Data.Hardware.Dispenser.Right.HeadType = "Positive Displacement" Then
                LB_Pro_RPressure.Visible = True
                LB_Pro_RSuck.Visible = True
                LB_Pro_RRPM.Visible = True
                LB_Pro_RREV.Visible = True
                BT_Pro_RDownload.Visible = True
                'Bt_pro_LOpenJetterSoftware.Visible = False
                'BTLeftBrowseJetter.Visible = False
                'TBLeftJetterFile.Visible = False
                'LabelLeftJetter.Visible = False

                TB_Pro_RPressure.Visible = True
                TB_Pro_RSuck.Visible = True
                TB_Pro_RRPM.Visible = True
                TB_Pro_RREV.Visible = True

                LB_Pro_RValve.Visible = True
                LB_Pro_RProg.Visible = True
                TB_Pro_RValve.Visible = True
                TB_Pro_RProg.Visible = True

                FrmProduction.LB_Pro_RPressure.Visible = True
                FrmProduction.TB_Pro_RPressure.Visible = True
                FrmProduction.LB_Pro_RSuck.Visible = True
                FrmProduction.TB_Pro_RSuck.Visible = True
                FrmProduction.LB_Pro_RRPM.Visible = True
                FrmProduction.TB_Pro_RRPM.Visible = True
                FrmProduction.LB_Pro_RREV.Visible = True
                FrmProduction.TB_Pro_RREV.Visible = True
                FrmProduction.LB_Pro_RValve.Visible = True
                FrmProduction.TB_Pro_RValve.Visible = True
                FrmProduction.LB_Pro_RProg.Visible = True
                FrmProduction.TB_Pro_RProg.Visible = True

                FrmProduction.BT_Pro_RDownload.Visible = True


                GB_RTP.Visible = True
                GB_RJetter.Visible = False
                FrmProduction.GB_RJetter.Visible = False
                FrmProduction.GB_RTP.Visible = True
            ElseIf IDS.Data.Hardware.Dispenser.Right.HeadType = "Fixed Volumetric" Then
                LB_Pro_RPressure.Visible = True
                LB_Pro_RSuck.Visible = True
                LB_Pro_RRPM.Visible = False
                LB_Pro_RREV.Visible = False
                BT_Pro_RDownload.Visible = True
                'Bt_pro_LOpenJetterSoftware.Visible = False
                'BTLeftBrowseJetter.Visible = False
                'TBLeftJetterFile.Visible = False
                'LabelLeftJetter.Visible = False

                TB_Pro_RPressure.Visible = True
                TB_Pro_RSuck.Visible = True
                TB_Pro_RRPM.Visible = False
                TB_Pro_RREV.Visible = False

                LB_Pro_RValve.Visible = False
                LB_Pro_RProg.Visible = False
                TB_Pro_RValve.Visible = False
                TB_Pro_RProg.Visible = False

                FrmProduction.LB_Pro_RPressure.Visible = True
                FrmProduction.TB_Pro_RPressure.Visible = True
                FrmProduction.LB_Pro_RSuck.Visible = True
                FrmProduction.TB_Pro_RSuck.Visible = True
                FrmProduction.LB_Pro_RRPM.Visible = False
                FrmProduction.TB_Pro_RRPM.Visible = False
                FrmProduction.LB_Pro_RREV.Visible = False
                FrmProduction.TB_Pro_RREV.Visible = False
                FrmProduction.LB_Pro_RValve.Visible = False
                FrmProduction.TB_Pro_RValve.Visible = False
                FrmProduction.LB_Pro_RProg.Visible = False
                FrmProduction.TB_Pro_RProg.Visible = False

                FrmProduction.BT_Pro_RDownload.Visible = True

                GB_RTP.Visible = True
                GB_RJetter.Visible = False
                FrmProduction.GB_RJetter.Visible = False
                FrmProduction.GB_RTP.Visible = True
            ElseIf IDS.Data.Hardware.Dispenser.Right.HeadType = "Jetter" Then
                LB_Pro_RPressure.Visible = False
                LB_Pro_RSuck.Visible = False
                LB_Pro_RRPM.Visible = False
                LB_Pro_RREV.Visible = False
                BT_Pro_RDownload.Visible = False
                'Bt_pro_LOpenJetterSoftware.Visible = False
                'BTLeftBrowseJetter.Visible = False
                'TBLeftJetterFile.Visible = False
                'LabelLeftJetter.Visible = False

                TB_Pro_RPressure.Visible = False
                TB_Pro_RSuck.Visible = False
                TB_Pro_RRPM.Visible = False
                TB_Pro_RREV.Visible = False

                LB_Pro_RValve.Visible = False
                LB_Pro_RProg.Visible = False
                TB_Pro_RValve.Visible = False
                TB_Pro_RProg.Visible = False

                FrmProduction.LB_Pro_RPressure.Visible = False
                FrmProduction.TB_Pro_RPressure.Visible = False
                FrmProduction.LB_Pro_RSuck.Visible = False
                FrmProduction.TB_Pro_RSuck.Visible = False
                FrmProduction.LB_Pro_RRPM.Visible = False
                FrmProduction.TB_Pro_RRPM.Visible = False
                FrmProduction.LB_Pro_RREV.Visible = False
                FrmProduction.TB_Pro_RREV.Visible = False
                FrmProduction.LB_Pro_RValve.Visible = False
                FrmProduction.TB_Pro_RValve.Visible = False
                FrmProduction.LB_Pro_RProg.Visible = False
                FrmProduction.TB_Pro_RProg.Visible = False

                FrmProduction.BT_Pro_RDownload.Visible = False


                GB_RTP.Visible = False
                GB_RJetter.Visible = True
                FrmProduction.GB_RJetter.Visible = True
                FrmProduction.GB_RTP.Visible = False
            ElseIf IDS.Data.Hardware.Dispenser.Right.HeadType = "Others" Then
                LB_Pro_RPressure.Visible = False
                LB_Pro_RSuck.Visible = False
                LB_Pro_RRPM.Visible = False
                LB_Pro_RREV.Visible = False
                BT_Pro_RDownload.Visible = False
                'Bt_pro_LOpenJetterSoftware.Visible = False
                'BTLeftBrowseJetter.Visible = False
                'TBLeftJetterFile.Visible = False
                'LabelLeftJetter.Visible = False

                TB_Pro_RPressure.Visible = False
                TB_Pro_RSuck.Visible = False
                TB_Pro_RRPM.Visible = False
                TB_Pro_RREV.Visible = False

                LB_Pro_RValve.Visible = False
                LB_Pro_RProg.Visible = False
                TB_Pro_RValve.Visible = False
                TB_Pro_RProg.Visible = False

                FrmProduction.LB_Pro_RPressure.Visible = False
                FrmProduction.TB_Pro_RPressure.Visible = False
                FrmProduction.LB_Pro_RSuck.Visible = False
                FrmProduction.TB_Pro_RSuck.Visible = False
                FrmProduction.LB_Pro_RRPM.Visible = False
                FrmProduction.TB_Pro_RRPM.Visible = False
                FrmProduction.LB_Pro_RREV.Visible = False
                FrmProduction.TB_Pro_RREV.Visible = False
                FrmProduction.LB_Pro_RValve.Visible = False
                FrmProduction.TB_Pro_RValve.Visible = False
                FrmProduction.LB_Pro_RProg.Visible = False
                FrmProduction.TB_Pro_RProg.Visible = False

                FrmProduction.BT_Pro_RDownload.Visible = False

                GB_RTP.Visible = True
                GB_RJetter.Visible = False
                FrmProduction.GB_RJetter.Visible = False
                FrmProduction.GB_RTP.Visible = True
            End If
        Else
            GB_RTP.Visible = False
            GB_RJetter.Visible = False
            FrmProduction.GB_RJetter.Visible = False
            FrmProduction.GB_RTP.Visible = False
        End If

        If IDS.Data.Hardware.Dispenser.Right.AlowEdit Then
            'FrmProduction.LB_Pro_RPressure.Visible = False
            FrmProduction.TB_Pro_RPressure.Enabled = True
            'FrmProduction.LB_Pro_RSuck.Visible = False
            FrmProduction.TB_Pro_RSuck.Enabled = True
            'FrmProduction.LB_Pro_RRPM.Visible = False
            FrmProduction.TB_Pro_RRPM.Enabled = True
            'FrmProduction.LB_Pro_RREV.Visible = False
            FrmProduction.TB_Pro_RREV.Enabled = True
            'FrmProduction.LB_Pro_RValve.Visible = False
            FrmProduction.TB_Pro_RValve.Enabled = True
            'FrmProduction.LB_Pro_RProg.Visible = False
            FrmProduction.TB_Pro_RProg.Enabled = True
            FrmProduction.BT_Pro_RDownload.Enabled = True
        Else
            'FrmProduction.LB_Pro_RPressure.Visible = False
            FrmProduction.TB_Pro_RPressure.Enabled = False
            'FrmProduction.LB_Pro_RSuck.Visible = False
            FrmProduction.TB_Pro_RSuck.Enabled = False
            'FrmProduction.LB_Pro_RRPM.Visible = False
            FrmProduction.TB_Pro_RRPM.Enabled = False
            'FrmProduction.LB_Pro_RREV.Visible = False
            FrmProduction.TB_Pro_RREV.Enabled = False
            'FrmProduction.LB_Pro_RValve.Visible = False
            FrmProduction.TB_Pro_RValve.Enabled = False
            'FrmProduction.LB_Pro_RProg.Visible = False
            FrmProduction.TB_Pro_RProg.Enabled = False
            FrmProduction.BT_Pro_RDownload.Enabled = False
        End If

        If IDS.Data.Hardware.Dispenser.Left.AlowEdit Then
            'FrmProduction.LB_Pro_RPressure.Visible = False
            FrmProduction.TB_Pro_LPressure.Enabled = True
            'FrmProduction.LB_Pro_RSuck.Visible = False
            FrmProduction.TB_Pro_LSuck.Enabled = True
            'FrmProduction.LB_Pro_RRPM.Visible = False
            FrmProduction.TB_Pro_LRPM.Enabled = True
            'FrmProduction.LB_Pro_RREV.Visible = False
            FrmProduction.TB_Pro_LREV.Enabled = True
            'FrmProduction.LB_Pro_RValve.Visible = False
            FrmProduction.TB_Pro_LValve.Enabled = True
            'FrmProduction.LB_Pro_RProg.Visible = False
            FrmProduction.TB_Pro_LProg.Enabled = True
            FrmProduction.BT_Pro_LDownload.Enabled = True
            FrmProduction.TB_LJetter_Path.Enabled = True
        Else
            'FrmProduction.LB_Pro_RPressure.Visible = False
            FrmProduction.TB_Pro_LPressure.Enabled = False
            'FrmProduction.LB_Pro_RSuck.Visible = False
            FrmProduction.TB_Pro_LSuck.Enabled = False
            'FrmProduction.LB_Pro_RRPM.Visible = False
            FrmProduction.TB_Pro_LRPM.Enabled = False
            'FrmProduction.LB_Pro_RREV.Visible = False
            FrmProduction.TB_Pro_LREV.Enabled = False
            'FrmProduction.LB_Pro_RValve.Visible = False
            FrmProduction.TB_Pro_LValve.Enabled = False
            'FrmProduction.LB_Pro_RProg.Visible = False
            FrmProduction.TB_Pro_LProg.Enabled = False
            FrmProduction.BT_Pro_LDownload.Enabled = False
            FrmProduction.TB_LJetter_Path.Enabled = False
        End If
    End Sub

    Private Sub RadioButton5_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RBProgramEditor.CheckedChanged
        If RBProgramEditor.Checked Then
            Disp_Dispenser_Unit_info()

            MyPSStationSettings.Timer1.Stop()
            CBExpandSpreadsheet.Visible = True
            If IDS.Data.Hardware.Gantry.SystemUnit = False Then 'mm
                LB_MM.Visible = True
                LB_INCH.Visible = False
            Else
                LB_MM.Visible = False
                LB_INCH.Visible = True
            End If
            'CBExpandSpreadsheet.Visible = True
            Me.Controls.Remove(MyBasicSetup.PanelRight)
            Me.Controls.Remove(MyBasicSetup.PanelLeft)
            Me.PanelVision.Location = New Point(84, 341)
            Me.PanelVisionCtrl.Location = New Point(84, 916)
            'Me.PanelVisionCtrl.SendToBack()
            'Me.PanelVision.SendToBack()
            'xu long 190707  /*
            IDS.Data.OpenData()

            If IDS.Data.Hardware.Gantry.SystemUnit = False Then 'mm lable show
                LB_MM.Visible = True
                LB_INCH.Visible = False
                'Example for change Unit setup,'1=mm; 0=inch
                ChangeUnitInSpredsheet(1)
            Else 'inch lable show
                LB_MM.Visible = False
                LB_INCH.Visible = True
                'Example for change Unit setup,'1=mm; 0=inch
                ChangeUnitInSpredsheet(0)
            End If
            'xu long 190707 */
            Try
                Dim GDsizeX, GDsizeY As Double
                GDsizeX = Abs(IDS.Data.Hardware.Gantry.GraphicDisplayXYRB.X - IDS.Data.Hardware.Gantry.GraphicDisplayXYLT.X)
                GDsizeY = Abs(IDS.Data.Hardware.Gantry.GraphicDisplayXYLT.Y - IDS.Data.Hardware.Gantry.GraphicDisplayXYRB.Y)
                GDRefX = IDS.Data.Hardware.Gantry.GraphicDisplayXYLT.X
                GDRefY = IDS.Data.Hardware.Gantry.GraphicDisplayXYLT.Y
                If GDsizeX = _GDSizeX And GDsizeY = _GDSizeY And GDRefX = _GDRefX And GDRefY = _GDRefY Then
                Else
                    FrmGD.SetGD(GDsizeX, GDsizeY, IDS.Data.Hardware.Gantry.GraphicDisplayXYLT.X, IDS.Data.Hardware.Gantry.GraphicDisplayXYLT.Y)
                    'have to refresh all the list to get the current GD display
                End If
            Catch ex As Exception

            End Try
            'SJ
            SystemSetupDataRetrieve()
        Else
            CBExpandSpreadsheet.Visible = False
            LB_INCH.Visible = False  'xu long 190707 
            LB_MM.Visible = False  'xu long 190707 
        End If
    End Sub
    Dim _GDSizeX, _GDSizeY As Double
    Dim _GDRefX, _GDRefY As Integer

    Private Sub RadioButton6_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RBBasicSetup.CheckedChanged
        If RBBasicSetup.Checked = True Then
            CBExpandSpreadsheet.Visible = False
            'Dim frm2 As New Form2
            'frm2.Instance(Me)
            'frm2.ShowDialog()
            MyBasicSetup.PanelRight.Location = New Point(768, 33)
            MyBasicSetup.PanelLeft.Location = New Point(0, 33)

            Me.PanelVision.Location = New Point(0, 341)
            Me.PanelVisionCtrl.Location = New Point(0, 916)
            'Me.PanelVisionCtrl.BringToFront()
            'Me.PanelVision.BringToFront()

            Me.Controls.Add(MyBasicSetup.PanelRight)
            Me.Controls.Add(MyBasicSetup.PanelLeft)
            MyBasicSetup.PanelRight.BringToFront()
            'MyBasicSetup.PanelRight.Show()
            MyBasicSetup.PanelLeft.BringToFront()
            MyBasicSetup.RichTextBox1.BringToFront()
            'MyBasicSetup.PanelLeft.Show()
            MyBasicSetup.SetTable()
            LB_INCH.Visible = False  'xu long 190707 
            LB_MM.Visible = False  'xu long 190707 

            If IDS.Data.Hardware.Dispenser.TimePressurePresent Then
                MyBasicSetup.CB_LeftHeadType.Items(0) = "Time Pressure"
                MyBasicSetup.CB_RightHeadType.Items(0) = "Time Pressure"
            End If
            If IDS.Data.Hardware.Dispenser.PositiveDispPresent Then
                MyBasicSetup.CB_LeftHeadType.Items(1) = "Positive Displacement"
                MyBasicSetup.CB_RightHeadType.Items(1) = "Positive Displacement"
            End If
            If IDS.Data.Hardware.Dispenser.FixedVolumetricPresent Then
                MyBasicSetup.CB_LeftHeadType.Items(2) = "Fixed Volumetric"
                MyBasicSetup.CB_RightHeadType.Items(2) = "Fixed Volumetric"
            End If
            If IDS.Data.Hardware.Dispenser.JetterPresent Then
                MyBasicSetup.CB_LeftHeadType.Items(3) = "Jetter"
                MyBasicSetup.CB_RightHeadType.Items(3) = "Jetter"
            End If

            If MyBasicSetup.RadioStationSettings.Checked Then
                'MyBasicSetup.PanelRight.Controls.Add(MyPSStationSettings.PanelToBeAdded)
                'MyPSStationSettings.PanelToBeAdded.BringToFront()
                'MyPSStationSettings.PanelToBeAdded.Location = New Point(0, 0)
                MyPSStationSettings.Timer1.Start()
                'MyPSStationSettings.SetTable()
            Else
                'MyBasicSetup.PanelRight.Controls.Remove(MyPSStationSettings.PanelToBeAdded)
                MyPSStationSettings.Timer1.Stop()
            End If

            'CBExpandSpreadsheet.Hide()


            'Lim For GD
            _GDSizeX = Abs(IDS.Data.Hardware.Gantry.GraphicDisplayXYRB.X - IDS.Data.Hardware.Gantry.GraphicDisplayXYLT.X)
            _GDSizeY = Abs(IDS.Data.Hardware.Gantry.GraphicDisplayXYLT.Y - IDS.Data.Hardware.Gantry.GraphicDisplayXYRB.Y)
            _GDRefX = IDS.Data.Hardware.Gantry.GraphicDisplayXYLT.X
            _GDRefY = IDS.Data.Hardware.Gantry.GraphicDisplayXYLT.Y
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


    ''SJ add
    'Private Sub CheckBox5_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CBDoorLock.CheckedChanged
    '    Dim status As Integer
    '    m_Tri.m_TriCtrl.GetTable(21, 1, status)
    '    If status <> 0 Then
    '        LabelMessege.Text = "Can't unlock the door when program running!"
    '        Exit Sub
    '    End If
    '    If CBDoorLock.Checked = True Then
    '        CBDoorLock.Text = "Unlock Door"
    '        CBDoorLock.ImageIndex = 6
    '        IDS.Devices.DIO.DIO.DoorLock(True)
    '    Else
    '        CBDoorLock.Text = "Lock Door"
    '        CBDoorLock.ImageIndex = 5
    '        IDS.Devices.DIO.DIO.DoorLock(False)
    '    End If
    'End Sub


    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Update link for next/previous row
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Private Sub UpdateLinkForNextRow()
        AxSpreadsheetProgramming_CheckForWithinLinkRange(True)
        With AxSpreadsheetProgramming.ActiveSheet
            If m_InLinkRange Then
                If "Line" = .Cells(m_row, gCmdColumn).value Or "Arc" = .Cells(m_row, gCmdColumn).value Then
                    If "Line" = .Cells(m_row, gCmdColumn).value Then
                        .Cells(m_row + 1, gPos1XColumn).value = .Cells(m_row, gPos2XColumn).value
                        .Cells(m_row + 1, gPos1YColumn).value = .Cells(m_row, gPos2YColumn).value
                    ElseIf "Arc" = .Cells(m_row, gCmdColumn).value Then
                        .Cells(m_row + 1, gPos1XColumn).value = .Cells(m_row, gPos3XColumn).value
                        .Cells(m_row + 1, gPos1YColumn).value = .Cells(m_row, gPos3YColumn).value
                    End If
                End If
            End If
        End With
    End Sub

    Private Sub UpdateLinkForPreviousRow()
        AxSpreadsheetProgramming_CheckForWithinLinkRange(False)
        With AxSpreadsheetProgramming.ActiveSheet
            If m_InLinkRange Then
                If "Line" = .Cells(m_row - 1, gCmdColumn).value Then
                    .Cells(m_row - 1, gPos2XColumn).value() = _
                        .Cells(m_row, gPos1XColumn).value
                    .Cells(m_row - 1, gPos2YColumn).value() = _
                        .Cells(m_row, gPos1YColumn).value
                ElseIf "Arc" = .Cells(m_row - 1, gCmdColumn).value Then
                    .Cells(m_row - 1, gPos3XColumn).value() = _
                        .Cells(m_row, gPos1XColumn).value
                    .Cells(m_row - 1, gPos3YColumn).value() = _
                        .Cells(m_row, gPos1YColumn).value
                End If
            End If
        End With
    End Sub



    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'TBNextEdit_ButtonClick: edit xyz coordinates using toolbar
    '                                                                                  
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Private Sub TBNextEdit_ButtonClick(ByVal sender As Object, ByVal e As System.Windows.Forms.ToolBarButtonClickEventArgs) Handles TBNextEdit.ButtonClick
        Dim idFlag As Integer = 0
        If e.Button Is TBNextEdit.Buttons(0) Then
            idFlag = 1
        End If

        TBNextEdit_Implementation(idFlag)
    End Sub


    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'TBNextEdit_ButtonClick: edit xyz coordinates using toolbar
    '    idFlag: 1=goto next point, 0=edit current point
    '
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Private Sub TBNextEdit_Implementation(ByVal idFlag As Integer)
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
        m_row = AxSpreadsheetProgramming.ActiveCell.Row
        CmdName = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gCmdColumn).value

        Dim sel As Microsoft.Office.Interop.OWC.Range = AxSpreadsheetProgramming.Selection
        Dim m_rowCount As Integer = sel.Rows.Count()
        Dim m_columnCount As Integer = sel.Columns.Count()
        Dim m_StartRow As Integer = sel.Row
        Dim m_StartColumn As Integer = sel.Column

        If 1 = idFlag Then             'Goto next point, next point button
            m_KeyinOK = False
            m_EditStateFlag = True
            m_PosUpdate = False
            LabelMessege.Text = "Edit or goto next Pt"
            If m_Execution.m_Pattern.m_ErrorChk.CheckValidPtXY(CmdName) Then

                If m_Execution.m_Pattern.m_ErrorChk.CheckRequiredPt3XY(CmdName) Then
                    posNoMax = 3
                ElseIf m_Execution.m_Pattern.m_ErrorChk.CheckRequiredPt2XY(CmdName) Then
                    posNoMax = 2
                Else
                    posNoMax = 1
                    EnableTBNextEdit(0, False)      'Goto next point. Position jump is not allowed
                End If

                If m_columnCount = gMaxColumns Then
                    m_PosNo = 1
                    posX = gPos1XColumn
                ElseIf gPos1XColumn <= m_StartColumn And gPos1ZColumn >= m_StartColumn Then
                    If posNoMax >= 2 Then
                        m_PosNo = 2
                        posX = gPos2XColumn
                    Else
                        m_PosNo = 1
                        posX = gPos1XColumn
                    End If

                    With AxSpreadsheetProgramming.ActiveSheet
                        cell1 = .Cells(m_row, gPos1XColumn)
                        cell2 = .Cells(m_row, gPos1XColumn + 2)
                    End With
                    m_Execution.m_Pattern.AxSpreadsheetProgramming_CellGrey(cell1, cell2, True, AxSpreadsheetProgramming)

                    'Previous confirmed data is the begaining of the row.  It will be used to update the end of previous row
                    UpdateLinkForPreviousRow()
                ElseIf gPos2XColumn <= m_StartColumn And gPos2ZColumn >= m_StartColumn Then
                    If posNoMax >= 3 Then
                        m_PosNo = 3
                        posX = gPos3XColumn
                    Else
                        m_PosNo = 1
                        posX = gPos1XColumn
                    End If

                    With AxSpreadsheetProgramming.ActiveSheet
                        cell1 = .Cells(m_row, gPos2XColumn)
                        cell2 = .Cells(m_row, gPos2XColumn + 2)
                    End With
                    m_Execution.m_Pattern.AxSpreadsheetProgramming_CellGrey(cell1, cell2, True, AxSpreadsheetProgramming)

                    'Previous confirmed data is the end of the row.  It will be used to update the begaining of next row
                    UpdateLinkForNextRow()

                ElseIf gPos3XColumn <= m_StartColumn And gPos3ZColumn >= m_StartColumn Then
                    m_PosNo = 1
                    posX = gPos1XColumn

                    With AxSpreadsheetProgramming.ActiveSheet
                        cell1 = .Cells(m_row, gPos3XColumn)
                        cell2 = .Cells(m_row, gPos3XColumn + 2)
                    End With
                    m_Execution.m_Pattern.AxSpreadsheetProgramming_CellGrey(cell1, cell2, True, AxSpreadsheetProgramming)

                    'Previous confirmed data is the end of the row.  It will be used to update the begaining of next row
                    UpdateLinkForNextRow()
                End If

                With AxSpreadsheetProgramming.ActiveSheet
                    cell1 = .Cells(m_row, posX)
                    cell2 = .Cells(m_row, posX + 2)
                    .Range(cell1, cell2).Select()

                    pos(0) = .Cells(m_row, posX).Value
                    pos(1) = .Cells(m_row, posX + 1).Value
                    pos(2) = .Cells(m_row, posX + 2).Value
                End With
                pos(0) = pos(0) * gmminchRatio
                pos(1) = pos(1) * gmminchRatio
                pos(2) = pos(2) * gmminchRatio
                ''''''SJ change
                'Trimotion_MoveTo(pos)
                If CmdName = "ChipEdge" Or CmdName = "QC" Or CmdName = "Height" Or CmdName = "Fiducial" Or CmdName = "Reject" Then
                    Trimotion_MoveTo(pos)
                Else
                    Trimotion_MoveTo(pos, m_TeachMode)
                End If

            End If
            TBYesNo.Buttons(0).Enabled = False
            TBYesNo.Buttons(1).Enabled = True
            EnableTBNextEdit(1, True) 'Edit start

            EnableTBElements(False)
            EnableTBRefers(False)
            UndoData_Logging(1)

        ElseIf 0 = idFlag Then         'Edit current point, pen button
            ToolBarSwitch("Edit")
            LabelMessege.Text = "Edit or confirm Pt"

            If m_PosNo = 1 Then
                posX = gPos1XColumn
            ElseIf m_PosNo = 2 Then
                posX = gPos2XColumn
            ElseIf m_PosNo = 3 Then
                posX = gPos3XColumn
            End If
            With AxSpreadsheetProgramming.ActiveSheet
                cell1 = .Cells(m_row, posX)
                cell2 = .Cells(m_row, posX + 2)
            End With
            m_Execution.m_Pattern.AxSpreadsheetProgramming_CellGrey(cell1, cell2, False, AxSpreadsheetProgramming)
            m_KeyinOK = True
            m_EditStateFlag = True

            EnableTBRefers(False)
            EnableTBElements(False)
            EnableTBNextEdit(1, False) 'Edit start
            TBYesNo.Buttons(1).Enabled = True
        End If
        GC.Collect()
    End Sub

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'TBEditOnly_ButtonClick: Not in use
    '                                                                                  
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Private Sub TBEditOnly_ButtonClick(ByVal sender As Object, _
        ByVal e As System.Windows.Forms.ToolBarButtonClickEventArgs) Handles TBEditOnly.ButtonClick
        If e.Button Is TBEditOnly.Buttons(0) Then
            TBYesNo.Show()
            TBNextEdit.Hide()
            TBEditOnly.Hide()
            'SJ
            EditOrCancel(1)
        End If
        GC.Collect()
    End Sub

    'Private Sub TBMachineBtn_ButtonClick(ByVal sender As System.Object, _
    '    ByVal e As System.Windows.Forms.ToolBarButtonClickEventArgs) Handles TBOperation.ButtonClick
    '    If e.Button Is TBOperation.Buttons(0) Then  'run
    '        PBRed.Hide()
    '        PBYellow.Hide()
    '        PBGreen.Show()
    '    ElseIf e.Button Is TBOperation.Buttons(1) Then 'pause
    '        PBRed.Hide()
    '        PBYellow.Show()
    '        PBGreen.Hide()
    '    ElseIf e.Button Is TBOperation.Buttons(2) Then 'stop
    '        PBRed.Show()
    '        PBYellow.Hide()
    '        PBGreen.Hide()
    '    End If
    'End Sub


    Private Sub TBMultiDisplayField_ButtonClick(ByVal sender As System.Object, _
        ByVal e As System.Windows.Forms.ToolBarButtonClickEventArgs) Handles TBMultiDisplayField.ButtonClick
        If e.Button Is TBMultiDisplayField.Buttons(0) Then
            PanelGraphDisp.Show()
            PanelGraphDisp.Controls.Add(FrmGD.Panel1)
            FrmGD.Panel1.Location = New Point(0, 0)
            PanelStepper.Hide()
            PanelDispenser.Hide()
            PanelThermal.Hide()
            PanelConveyor.Hide()
            TBMultiDisplayField.Buttons(0).Pushed = True
            TBMultiDisplayField.Buttons(1).Pushed = False
            TBMultiDisplayField.Buttons(2).Pushed = False
            TBMultiDisplayField.Buttons(3).Pushed = False
            TBMultiDisplayField.Buttons(4).Pushed = False
        ElseIf e.Button Is TBMultiDisplayField.Buttons(1) Then
            PanelGraphDisp.Hide()
            PanelStepper.Show()
            PanelDispenser.Hide()
            PanelThermal.Hide()
            PanelConveyor.Hide()
            TBMultiDisplayField.Buttons(0).Pushed = False
            TBMultiDisplayField.Buttons(1).Pushed = True
            TBMultiDisplayField.Buttons(2).Pushed = False
            TBMultiDisplayField.Buttons(3).Pushed = False
            TBMultiDisplayField.Buttons(4).Pushed = False
        ElseIf e.Button Is TBMultiDisplayField.Buttons(2) Then
            PanelGraphDisp.Hide()
            PanelStepper.Hide()
            PanelDispenser.Show()
            PanelThermal.Hide()
            PanelConveyor.Hide()
            TBMultiDisplayField.Buttons(0).Pushed = False
            TBMultiDisplayField.Buttons(1).Pushed = False
            TBMultiDisplayField.Buttons(2).Pushed = True
            TBMultiDisplayField.Buttons(3).Pushed = False
            TBMultiDisplayField.Buttons(4).Pushed = False
        ElseIf e.Button Is TBMultiDisplayField.Buttons(3) Then
            PanelGraphDisp.Hide()
            PanelStepper.Hide()
            PanelDispenser.Hide()
            PanelThermal.Show()
            PanelConveyor.Hide()
            TBMultiDisplayField.Buttons(0).Pushed = False
            TBMultiDisplayField.Buttons(1).Pushed = False
            TBMultiDisplayField.Buttons(2).Pushed = False
            TBMultiDisplayField.Buttons(3).Pushed = True
            TBMultiDisplayField.Buttons(4).Pushed = False
        ElseIf e.Button Is TBMultiDisplayField.Buttons(4) Then
            PanelGraphDisp.Hide()
            PanelStepper.Hide()
            PanelDispenser.Hide()
            PanelThermal.Hide()
            PanelConveyor.Show()
            TBMultiDisplayField.Buttons(0).Pushed = False
            TBMultiDisplayField.Buttons(1).Pushed = False
            TBMultiDisplayField.Buttons(2).Pushed = False
            TBMultiDisplayField.Buttons(3).Pushed = False
            TBMultiDisplayField.Buttons(4).Pushed = True
        End If
        GC.Collect()
    End Sub

    Private Sub InitialPatternParameterSetup()

        'SJ
        '     FrmMyExport.PanelMotor.Top = 0
        '    FrmMyExport.PanelMotor.Left = 0
        '   PanelMotor.Controls.Add(FrmMyExport.PanelMotor)
        '''''''''''''''
        ''''_IDS.Admin.User.Group.ID()
        Dim gpID As String = IDS.Data.Admin.User.Group.ID
        'SJ test only
        Dim recID As String
        m_Execution.m_Pattern.LoadDefaultPara(gpID, recID, 0)
        m_PosNo = 1
        m_CamPos(0) = 0.0
        m_CamPos(1) = 0.0
        m_CamPos(2) = 0.0
        GC.Collect()
    End Sub

    Public Sub AddPatternCommand(ByVal type As String)
        Dim cell1, cell2 As Object
        Dim SkippedFirstPt As Boolean = False
        Dim Cmd As CIDSData.CIDSPatternExecution.CIDSTemplate
        Cmd = m_Execution.m_Pattern.AddCommand(type, m_CamPos)


        AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, 1).Select()
        AxSpreadsheetProgramming.Selection.EntireRow.Insert()
        AxSpreadsheetProgramming.ActiveSheet.Rows(m_row).clear()

        With AxSpreadsheetProgramming.ActiveSheet
            .Cells(m_row, gCmdColumn) = type
            If (type.ToUpper <> "LINKSTART" And type.ToUpper <> "LINKEND") Then

                m_Execution.m_ItemLUT.Cmd2Index(type.ToUpper)
                If m_Execution.m_ItemLUT.GetFlag(gPos1XColumn) Then
                    If "Array" <> type Then
                        If m_InLinkRange Then
                            If "Line" = .Cells(m_row - 1, gCmdColumn).value Then
                                .Cells(m_row, gPos1XColumn) = .Cells(m_row - 1, gPos2XColumn)
                                .Cells(m_row, gPos1YColumn) = .Cells(m_row - 1, gPos2YColumn)
                                .Cells(m_row, gPos1ZColumn) = .Cells(m_row - 1, gPos2ZColumn)
                                m_PosNo = 2
                                cell1 = .Cells(m_row, gPos1XColumn)
                                cell2 = .Cells(m_row, gPos1ZColumn)
                                m_Execution.m_Pattern.AxSpreadsheetProgramming_CellGrey(cell1, cell2, True, AxSpreadsheetProgramming)
                                .Cells(m_row, gPos2XColumn).Select()
                                SkippedFirstPt = True

                            ElseIf "Arc" = .Cells(m_row - 1, gCmdColumn).value Then
                                .Cells(m_row, gPos1XColumn) = .Cells(m_row - 1, gPos3XColumn)
                                .Cells(m_row, gPos1YColumn) = .Cells(m_row - 1, gPos3YColumn)
                                .Cells(m_row, gPos1ZColumn) = .Cells(m_row - 1, gPos3ZColumn)
                                m_PosNo = 2
                                cell1 = .Cells(m_row, gPos1XColumn)
                                cell2 = .Cells(m_row, gPos1ZColumn)
                                m_Execution.m_Pattern.AxSpreadsheetProgramming_CellGrey(cell1, cell2, True, AxSpreadsheetProgramming)
                                .Cells(m_row, gPos2XColumn).Select()
                                SkippedFirstPt = True
                            Else
                                'If "Line" = .Cells(m_row + 1, gCmdColum).value Or _
                                '   "Arc" = .Cells(m_row + 1, gCmdColum).value Then


                                'Else
                                .Cells(m_row, gPos1XColumn) = Cmd.Pos1.X
                                .Cells(m_row, gPos1YColumn) = Cmd.Pos1.Y
                                .Cells(m_row, gPos1ZColumn) = Cmd.Pos1.Z

                                cell1 = .Cells(m_row, gPos1XColumn)
                                cell2 = .Cells(m_row, gPos1ZColumn)
                                .Range(cell1, cell2).Select()
                                'End If

                            End If
                        Else
                            .Cells(m_row, gPos1XColumn) = Cmd.Pos1.X
                            .Cells(m_row, gPos1YColumn) = Cmd.Pos1.Y
                            .Cells(m_row, gPos1ZColumn) = Cmd.Pos1.Z
                        End If
                    End If
                End If
                If m_Execution.m_ItemLUT.GetFlag(gPos2XColumn) Then
                    If "Array" <> type Then
                        .Cells(m_row, gPos2XColumn) = Cmd.Pos2.X
                        .Cells(m_row, gPos2YColumn) = Cmd.Pos2.Y
                        .Cells(m_row, gPos2ZColumn) = Cmd.Pos2.Z
                        If SkippedFirstPt = True Then
                            cell1 = .Cells(m_row, gPos2XColumn)
                            cell2 = .Cells(m_row, gPos2ZColumn)
                            .Range(cell1, cell2).Select()
                        End If
                    End If
                End If
                If m_Execution.m_ItemLUT.GetFlag(gPos3XColumn) Then
                    If "Array" <> type Then
                        .Cells(m_row, gPos3XColumn) = Cmd.pos3.X
                        .Cells(m_row, gPos3YColumn) = Cmd.pos3.Y
                        .Cells(m_row, gPos3ZColumn) = Cmd.pos3.Z
                    End If
                End If
                If m_Execution.m_ItemLUT.GetFlag(gNeedleColumn) Then
                    .Cells(m_row, gNeedleColumn) = "Right"  'Cmd.Needle
                End If
                If m_Execution.m_ItemLUT.GetFlag(gDispensColumn) Then
                    'gSubnameColumn use the same slot
                    If Cmd.DispenseFlag = "True" Then
                        Cmd.DispenseFlag = "On"
                    Else
                        Cmd.DispenseFlag = "Off"
                    End If
                    .Cells(m_row, gDispensColumn) = Cmd.DispenseFlag
                End If

                If m_Execution.m_ItemLUT.GetFlag(gTravelSpeedColumn) Then
                    .Cells(m_row, gTravelSpeedColumn) = Cmd.TravelSpeed
                End If
                If m_Execution.m_ItemLUT.GetFlag(gNeedleGapColumn) Then
                    .Cells(m_row, gNeedleGapColumn) = Cmd.NeedleGap
                End If
                If m_Execution.m_ItemLUT.GetFlag(gDurationColumn) Then
                    .Cells(m_row, gDurationColumn) = Cmd.DispenseDuration
                End If
                If m_Execution.m_ItemLUT.GetFlag(gTravelDelayColumn) Then
                    .Cells(m_row, gTravelDelayColumn) = Cmd.TravelDelay
                End If
                If m_Execution.m_ItemLUT.GetFlag(gRetractDelayColumn) Then
                    .Cells(m_row, gRetractDelayColumn) = Cmd.RetractDelay
                End If
                If m_Execution.m_ItemLUT.GetFlag(gApproachHtColumn) Then
                    .Cells(m_row, gApproachHtColumn) = Cmd.ApproachDispHeight
                End If
                If m_Execution.m_ItemLUT.GetFlag(gRetractSpeedColumn) Then
                    .Cells(m_row, gRetractSpeedColumn) = Cmd.RetractSpeed
                End If
                If m_Execution.m_ItemLUT.GetFlag(gRetractHtColumn) Then
                    .Cells(m_row, gRetractHtColumn) = Cmd.RetractHeight
                End If
                If m_Execution.m_ItemLUT.GetFlag(gClearanceHtColumn) Then
                    .Cells(m_row, gClearanceHtColumn) = Cmd.ClearanceHeight
                End If
                If m_Execution.m_ItemLUT.GetFlag(gDTaiDistColumn) Then
                    .Cells(m_row, gDTaiDistColumn) = Cmd.DetailingDistance
                End If
                If m_Execution.m_ItemLUT.GetFlag(gArcRadiusColumn) Then
                    .Cells(m_row, gArcRadiusColumn) = Cmd.ArcRadius
                End If
                If m_Execution.m_ItemLUT.GetFlag(gPitchColumn) Then
                    .Cells(m_row, gPitchColumn) = Cmd.Pitch
                End If

                If m_Execution.m_ItemLUT.GetFlag(gFillHeightColumn) Then
                    .Cells(m_row, gFillHeightColumn) = Cmd.FilledHeight
                End If
                If m_Execution.m_ItemLUT.GetFlag(gNoOfRunColumn) Then
                    .Cells(m_row, gNoOfRunColumn) = Cmd.NumofRun '1
                End If
                If m_Execution.m_ItemLUT.GetFlag(gSprialColumn) Then
                    .Cells(m_row, gSprialColumn) = Cmd.SpiralFlag
                End If
                If m_Execution.m_ItemLUT.GetFlag(gRtAngleColumn) Then
                    .Cells(m_row, gRtAngleColumn) = Cmd.RotatingAngle
                End If
                If m_Execution.m_ItemLUT.GetFlag(gEdgeClearColumn) Then
                    .Cells(m_row, gEdgeClearColumn) = Cmd.EdgeClearance
                End If
                If m_Execution.m_ItemLUT.GetFlag(gIONumberColumn) Then
                    .Cells(m_row, gIONumberColumn) = 1
                End If
            End If
        End With
        GC.Collect()
    End Sub

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'SetDefaultPositionToLastAbovePt: Not in use
    '                                                                                  
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Private Sub SetDefaultPositionToLastAbovePt(ByVal CmdStr As String)
        'Find the last valid point above current row.  We will set this as the last point used for programming.
        'This function is not in use currently.
        Dim pos(3) As Double
        Dim RefPos(3) As Double
        Dim ValidPtFound As Boolean = False

        Dim CurrentRow As Integer = AxSpreadsheetProgramming.ActiveCell.Row
        If m_Execution.m_Pattern.m_ErrorChk.CheckValidPtXY(CmdStr) Then

            If CurrentRow = 1 Then
                CurrentRow = 2
            End If

            ValidPtFound = m_Execution.m_Pattern.AxSpreadsheetProgramming_FindLastValidPoint( _
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
            m_Execution.m_Pattern.AxSpreadsheetProgramming_GetRowLocalReference( _
                AxSpreadsheetProgramming, CurrentRow, RefPos)
            AxSpreadsheetProgramming_SetMotionRef(RefPos)
            pos(0) = pos(0) * gmminchRatio
            pos(1) = pos(1) * gmminchRatio
            pos(2) = pos(2) * gmminchRatio
            Trimotion_MoveTo(pos)
        End If

    End Sub


    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Add Ref pattern command
    '                                                                                  
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Private Sub TBrefer_ButtonClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ToolBarButtonClickEventArgs) Handles TBrefer.ButtonClick
        Dim buttonText As String = e.Button.Text
        buttonText = buttonText.Trim(" ")
        m_EditStateFlag = False
        m_ProgrammingStateFlag = True
        UndoData_Logging(0)
        TBrefer_Implementation(buttonText)
        GC.Collect()
    End Sub


    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Add Ref pattern command implementation
    '                                                                                  
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Private Sub TBrefer_Implementation(ByVal ButtonText As String)
        Dim cell1, cell2 As Object

        SheetRowSelected = False
        EnableTBNextEdit(False)
        TBYesNo.Buttons(0).Enabled = True
        TBYesNo.Buttons(1).Enabled = True

        AxSpreadsheetProgramming.DisplayWorkbookTabs = False

        EnableTBRefers(False)
        EnableTBElements(False)

        'SetDefaultPositionToLastAbovePt(ButtonText)
        Dim sel As Microsoft.Office.Interop.OWC.Range = AxSpreadsheetProgramming.Selection
        Dim m_rowStart As Integer = sel.Row

        Dim RefPos(3) As Double
        If "Reference" = ButtonText Then
            RefPos(0) = 0 : RefPos(1) = 0 : RefPos(2) = 0
            AxSpreadsheetProgramming_SetMotionRef(RefPos)
        Else
            m_Execution.m_Pattern.AxSpreadsheetProgramming_GetRowLocalReference( _
                AxSpreadsheetProgramming, m_rowStart, RefPos)
            AxSpreadsheetProgramming_SetMotionRef(RefPos)
        End If

        Try

            Select Case (ButtonText)
                Case "Fiducial"
                    type = "Fiducial"
                    'IDS.Devices.Vision.IDSV_Form_FI()
                    m_PosNo = 1
                    m_cmdCompleted = 1
                    m_row = AxSpreadsheetProgramming.ActiveCell.Row
                    AddPatternCommand(type)
                    cell1 = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gPos1XColumn)
                    cell2 = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gPos1ZColumn)
                    AxSpreadsheetProgramming.ActiveSheet.Range(cell1, cell2).Select()
                    ToolBarSwitch("YesNo")
                Case "Reference"            'Absolute coordinate only
                    type = "Reference"
                    'SJ add
                    m_ReferPt(0) = 0.0
                    m_ReferPt(1) = 0.0
                    m_ReferPt(2) = 0.0
                    '
                    m_PosNo = 1
                    m_cmdCompleted = 1
                    m_row = AxSpreadsheetProgramming.ActiveCell.Row
                    AddPatternCommand(type)
                    cell1 = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gPos1XColumn)
                    cell2 = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gPos1ZColumn)
                    AxSpreadsheetProgramming.ActiveSheet.Range(cell1, cell2).Select()
                    ToolBarSwitch("YesNo")
                Case "Height"
                    type = "Height"
                    m_PosNo = 1
                    m_cmdCompleted = 1
                    m_row = AxSpreadsheetProgramming.ActiveCell.Row
                    AddPatternCommand(type)
                    cell1 = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gPos1XColumn)
                    cell2 = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gPos1ZColumn)
                    AxSpreadsheetProgramming.ActiveSheet.Range(cell1, cell2).Select()
                    ToolBarSwitch("YesNo")
                Case "Reject"
                    type = "Reject"
                    m_PosNo = 1
                    m_cmdCompleted = 0
                    m_PosUpdate = True
                    m_row = AxSpreadsheetProgramming.ActiveCell.Row
                    AddPatternCommand(type)
                    cell1 = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gPos1XColumn)
                    cell2 = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gPos1ZColumn)
                    AxSpreadsheetProgramming.ActiveSheet.Range(cell1, cell2).Select()

                    IDS.Devices.Vision.IDSV_Form_RM()
                    Dim status As Integer = 0
                    Dim x, y As Double

                    While status = 0
                        Application.DoEvents()
                        status = IDS.Devices.Vision.GetRMStatus
                    End While
                    If status = 2 Then          'Cancel
                        m_PosUpdate = False
                        LabelMessege.Text = ""
                        AxSpreadsheetProgramming.DisplayWorkbookTabs = True
                        AxSpreadsheetProgramming.ActiveSheet.Rows(m_row).delete()
                        AxSpreadsheetProgramming.Update()
                        AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gCmdColumn).Select()


                    ElseIf status = 1 Then      'Ok
                        Dim bb As Boolean
                        Dim vpara As DLL_Export_Device_Vision.RejectPoint.RMParam
                        IDS.Devices.Vision.GetRMParameters(vpara)
                        IDS.Devices.Vision.SetRMReset()
                        m_PosUpdate = False
                        With AxSpreadsheetProgramming.ActiveSheet
                            .Cells(m_row, gAcceptRatioCoulumn) = vpara._AcceptRatio
                            .Cells(m_row, gBinarizedColumn) = vpara._Binarized
                            .Cells(m_row, gBlackWithoutRMCoulumn) = vpara._BlackWithoutRM
                            .Cells(m_row, gBlackWithRMCoulumn) = vpara._BlackWithRM
                            .Cells(m_row, gBrightnessColumn) = vpara._Brightness
                            .Cells(m_row, gMRegionXColumn) = vpara._MRegionX
                            .Cells(m_row, gMRegionYColumn) = vpara._MRegionY
                            .Cells(m_row, gMROIxColumn) = vpara._MROIx
                            .Cells(m_row, gMROIyColumn) = vpara._MROIy
                            .Cells(m_row, gWhiteWithoutRMCoulumn) = vpara._WhiteWithoutRM
                            .Cells(m_row, gWhiteWithRMCoulumn) = vpara._WhiteWithRM
                            .Cells(m_row, gWoRMCoulumn) = vpara._WoRM
                            .Cells(m_row, gWRMCoulumn) = vpara._WRM

                            cell1 = .Cells(m_row, gPos1XColumn)
                            cell2 = .Cells(m_row, gPos1ZColumn)
                        End With

                        m_Execution.m_Pattern.AxSpreadsheetProgramming_CellGrey(cell1, cell2, True, AxSpreadsheetProgramming)

                        LabelMessege.Text = ""
                        AxSpreadsheetProgramming.DisplayWorkbookTabs = True
                        m_ProgrammingStateFlag = False
                        AxSpreadsheetProgramming.ActiveSheet.Cells(m_row + 1, gCmdColumn).Select()
                    End If
                    EnableTBRefers(True)
                    EnableTBElements(True)
                    EnableTBElements(gOffsetCmdIndex, False)
                    TBYesNo.Buttons(0).Enabled = False
                    TBYesNo.Buttons(1).Enabled = True
            End Select

        Catch ex As SystemException
            MessageBox.Show(ex.ToString)
        End Try

    End Sub

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Add Element pattern command
    '                                                                                  
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Friend Sub TBElements_ButtonClick(ByVal sender As System.Object, _
        ByVal e As System.Windows.Forms.ToolBarButtonClickEventArgs) Handles TBElements.ButtonClick

        Dim buttonText As String = e.Button.Text
        buttonText = buttonText.Trim(" ")
        m_EditStateFlag = False
        m_ProgrammingStateFlag = True

        Dim SheetName As String = AxSpreadsheetProgramming.ActiveSheet.Name
        If m_Execution.m_Pattern.AxSpreadsheetProgramming_IsAnArraySheet(AxSpreadsheetProgramming, SheetName) Then
            MessageBox.Show("Command is not allowed in an Array sheet", "Warnning!")
        Else
            If "Measure" <> buttonText Then
                UndoData_Logging(0)
            End If
            TBElements_Implementation(buttonText)
        End If

    End Sub


    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Add Element pattern command implementation
    '                                                                                  
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Friend Sub TBElements_Implementation(ByVal ButtonText As String)

        Dim i As Integer
        Dim Offset(3) As Double

        SheetRowSelected = False

        TBYesNo.Buttons(0).Enabled = True
        TBYesNo.Buttons(1).Enabled = True
        EnableTBRefers(False)
        EnableTBElements(False)
        EnableTBNextEdit(False)

        Dim pos(3), RefPos(3) As Double
        Dim sel As Microsoft.Office.Interop.OWC.Range = AxSpreadsheetProgramming.Selection
        Dim m_rowStart As Integer = sel.Row
        Dim m_rowEnd As Integer = (m_rowStart + sel.Rows.Count() - 1)

        'SetDefaultPositionToLastAbovePt(ButtonText)

        Try
            AxSpreadsheetProgramming.DisplayWorkbookTabs = False


            If "Offset" = ButtonText Then

            ElseIf "Reference" = ButtonText Then
                RefPos(0) = 0 : RefPos(1) = 0 : RefPos(2) = 0
                AxSpreadsheetProgramming_SetMotionRef(RefPos)
            Else
                m_Execution.m_Pattern.AxSpreadsheetProgramming_GetRowLocalReference( _
                    AxSpreadsheetProgramming, m_rowStart, RefPos)
                AxSpreadsheetProgramming_SetMotionRef(RefPos)
            End If


            Select Case (ButtonText)
                Case "Offset"
                    TBYesNo.Buttons(0).Enabled = False
                    Dim xyzOffsetDlg As New FormOffsetInput
                    If xyzOffsetDlg.ShowDialog() = DialogResult.OK Then
                        Offset(0) = CDbl(xyzOffsetDlg.TextBoxX.Text)
                        Offset(1) = CDbl(xyzOffsetDlg.TextBoxY.Text)
                        Offset(2) = CDbl(xyzOffsetDlg.TextBoxZ.Text)

                        Dim CmdStr As String


                        For i = m_rowStart To m_rowEnd
                            With AxSpreadsheetProgramming.ActiveSheet
                                CmdStr = .Cells(i, gCmdColumn).value
                                If m_Execution.m_Pattern.m_ErrorChk.CheckValidPtXY(CmdStr) Then
                                    If m_Execution.m_Pattern.m_ErrorChk.CheckRequiredPt3XY(CmdStr) Then
                                        .Cells(i, gPos1XColumn).value += Offset(0)
                                        .Cells(i, gPos1YColumn).value += Offset(1)
                                        .Cells(i, gPos1ZColumn).value += Offset(2)
                                        .Cells(i, gPos2XColumn).value += Offset(0)
                                        .Cells(i, gPos2YColumn).value += Offset(1)
                                        .Cells(i, gPos2ZColumn).value += Offset(2)
                                        .Cells(i, gPos3XColumn).value += Offset(0)
                                        .Cells(i, gPos3YColumn).value += Offset(1)
                                        .Cells(i, gPos3ZColumn).value += Offset(2)
                                    ElseIf m_Execution.m_Pattern.m_ErrorChk.CheckRequiredPt2XY(CmdStr) Then
                                        .Cells(i, gPos1XColumn).value += Offset(0)
                                        .Cells(i, gPos1YColumn).value += Offset(1)
                                        .Cells(i, gPos1ZColumn).value += Offset(2)
                                        .Cells(i, gPos2XColumn).value += Offset(0)
                                        .Cells(i, gPos2YColumn).value += Offset(1)
                                        .Cells(i, gPos2ZColumn).value += Offset(2)
                                    Else
                                        .Cells(i, gPos1XColumn).value += Offset(0)
                                        .Cells(i, gPos1YColumn).value += Offset(1)
                                        If m_Execution.m_Pattern.m_ErrorChk.CheckRequiredPtHeight(CmdStr) Then
                                            .Cells(i, gPos1ZColumn).value += Offset(2)
                                        End If
                                    End If
                                End If
                            End With
                        Next

                        'Move robot arm to the last offset position

                        With AxSpreadsheetProgramming.ActiveSheet
                            pos(0) = .Cells(m_rowEnd, gPos1XColumn).Value
                            pos(1) = .Cells(m_rowEnd, gPos1YColumn).Value
                            pos(2) = .Cells(m_rowEnd, gPos1ZColumn).Value
                        End With
                        m_Execution.m_Pattern.AxSpreadsheetProgramming_GetRowLocalReference(AxSpreadsheetProgramming, m_rowEnd, RefPos)
                        AxSpreadsheetProgramming_SetMotionRef(RefPos)
                        pos(0) = pos(0) * gmminchRatio
                        pos(1) = pos(1) * gmminchRatio
                        pos(2) = pos(2) * gmminchRatio
                        Trimotion_MoveTo(pos)
                    End If
                    m_ProgrammingStateFlag = False

                Case "Measure"
                    EnableTBRefers(False)
                    EnableTBElements(False)


                    IDS.Devices.Vision.IDSV_Measurement()
                    Dim status As Boolean
                    Dim x, y, Dist As Double
                    While status = False
                        Application.DoEvents()
                        status = IDS.Devices.Vision.IDSV_MeasurementFlag(x, y, Dist)
                    End While

                    Dist = CInt(Dist * 1000) / 1000
                    Dim DistInfo As String
                    DistInfo = "Measured dist x=" + CStr(x) + ", y=" + CStr(y) + ", d=" + CStr(Dist)
                    MsgBox(DistInfo, MsgBoxStyle.Information + MsgBoxStyle.OKOnly, "Measurement")
                    EnableTBRefers(True)
                    EnableTBElements(True)
                    EnableTBElements(gOffsetCmdIndex, False)

                    TBYesNo.Buttons(0).Enabled = False
                    TBYesNo.Buttons(1).Enabled = False

                    m_ProgrammingStateFlag = False

                Case "Dot", "Move", "Line", "Arc", "Circle", "Rectangle", _
                    "FillCircle", "FillRectangle", "ChipEdge", "QC", "SubPattern", "Wait"

                    type = ButtonText
                    m_PosNo = 1
                    m_row = AxSpreadsheetProgramming.ActiveCell.Row
                    AddPatternCommand(type)

                    Dim cell1, cell2 As Object
                    If m_InLinkRange Then

                    Else
                        cell1 = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gPos1XColumn)
                        cell2 = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gPos1ZColumn)
                        AxSpreadsheetProgramming.ActiveSheet.Range(cell1, cell2).Select()
                    End If

                    ToolBarSwitch("YesNo")



                    ''SJ add ****************************************************************************
                    If ButtonText = "ChipEdge" Then
                        IDS.Devices.Vision.IDSV_Form_CE()
                        Dim status As Integer = 0
                        Dim x, y As Double
                        Do
                            While Not (IDS.Devices.Vision.RobotMotionOffset(x, y) = True Or IDS.Devices.Vision.GetChipEdgeStatus = 2)
                                Application.DoEvents()
                            End While
                            If IDS.Devices.Vision.GetChipEdgeStatus = 2 Then
                                LabelMessege.Text = ""
                                TBElements.Enabled = True
                                TBrefer.Enabled = True
                                m_PosUpdate = False
                                AxSpreadsheetProgramming.DisplayWorkbookTabs = True
                                AxSpreadsheetProgramming.ActiveSheet.Rows(m_row).clear()
                                AxSpreadsheetProgramming.Update()
                                Return
                            End If

                            'moverobot
                            pos(0) = x
                            pos(1) = -y
                            pos(2) = 0
                            Dim spstr As String = gSystemSpeed.ToString
                            Dim str As String
                            str = "SPEED AXIS(0) =" + spstr
                            m_Tri.m_TriCtrl.Execute(str)
                            str = "SPEED AXIS(1) =" + spstr
                            m_Tri.m_TriCtrl.Execute(str)
                            str = "SPEED AXIS(2) =" + spstr
                            m_Tri.m_TriCtrl.Execute(str)
                            'Dim spstr As String = gSystemSpeed.ToString
                            'spstr = "SPEED AXIS(0) =" + spstr
                            'm_Tri.m_TriCtrl.Execute(spstr)
                            m_Tri.m_TriCtrl.MoveRel(2, pos, 0)
                            m_Tri.WaitForEndOfMove()
                            '??????????????????
                            If status = 3 Then
                                status = 0 'reset status after 5 points being reset
                            End If
                            While status = 0
                                Application.DoEvents()
                                status = IDS.Devices.Vision.GetChipEdgeStatus()
                            End While
                        Loop While status = 3 'status 3= reset 5 points

                        If status = 2 Then
                            LabelMessege.Text = ""
                            TBElements.Enabled = True
                            TBrefer.Enabled = True
                            m_PosUpdate = False
                            AxSpreadsheetProgramming.DisplayWorkbookTabs = True
                            AxSpreadsheetProgramming.ActiveSheet.Rows(m_row).clear()
                            AxSpreadsheetProgramming.Update()
                            Return
                        End If

                        If status = 1 Then 'chipedge finished settings
                            Dim bb As Boolean
                            Dim vpara As DLL_Export_Device_Vision.ChipEdgePoints.ChipEdgeParam
                            IDS.Devices.Vision.GetChipEdgeParameters(vpara)
                            IDS.Devices.Vision.SetCEReset()
                            IDS.Devices.Vision.FrmVision.DisplayIndicator()
                            With AxSpreadsheetProgramming.ActiveSheet
                                .Cells(m_row, gEdgeClearColumn) = vpara._EdgeClearance / gmminchRatio
                                .Cells(m_row, gBrightnessColumn) = vpara._Brightness
                                .Cells(m_row, gCheckBoxColumn) = vpara._CheckBox_ChipRec_Enable
                                .Cells(m_row, gCwCCwColumn) = vpara._Cw_CCw
                                .Cells(m_row, gDispModelColumn) = vpara._DispenseModel
                                .Cells(m_row, gInOutColumn) = vpara._Inside_out
                                .Cells(m_row, gMainEdgeColumn) = vpara._MainEdge
                                .Cells(m_row, gPointX1Column) = vpara._PointX1
                                .Cells(m_row, gPointX2Column) = vpara._PointX2
                                .Cells(m_row, gPointX3Column) = vpara._PointX3
                                .Cells(m_row, gPointX4Column) = vpara._PointX4
                                .Cells(m_row, gPointX5Column) = vpara._PointX5
                                .Cells(m_row, gPointY1Column) = vpara._PointY1
                                .Cells(m_row, gPointY2Column) = vpara._PointY2
                                .Cells(m_row, gPointY3Column) = vpara._PointY3
                                .Cells(m_row, gPointY4Column) = vpara._PointY4
                                .Cells(m_row, gPointY5Column) = vpara._PointY5
                                .Cells(m_row, gPosColumn) = vpara._Pos
                                .Cells(m_row, gPosXColumn) = vpara._PosX
                                .Cells(m_row, gPosYColumn) = vpara._PosY
                                .Cells(m_row, gROIColumn) = vpara._ROI
                                .Cells(m_row, gRotColumn) = vpara._Rot
                                .Cells(m_row, gSizeColumn) = vpara._Size
                                .Cells(m_row, gSizeXColumn) = vpara._SizeX
                                .Cells(m_row, gSizeYColumn) = vpara._SizeY
                                .Cells(m_row, gThresholdColumn) = vpara._Threshold
                                .Cells(m_row, gVerticalColumn) = vpara._Vertical

                                cell1 = .Cells(m_row, gPos1XColumn)
                                cell2 = .Cells(m_row, gPos1ZColumn)


                            End With


                            m_Execution.m_Pattern.AxSpreadsheetProgramming_CellGrey(cell1, cell2, True, AxSpreadsheetProgramming)
                            TBElements.Enabled = True
                            TBrefer.Enabled = True
                            m_PosUpdate = False
                            LabelMessege.Text = ""
                            AxSpreadsheetProgramming.DisplayWorkbookTabs = True
                            AxSpreadsheetProgramming.ActiveSheet.Cells(m_row + 1, 1).Select()
                        End If
                        EnableTBRefers(True)
                        EnableTBElements(True)
                        EnableTBElements(gOffsetCmdIndex, False)
                        TBYesNo.Buttons(0).Enabled = False
                        TBYesNo.Buttons(1).Enabled = True

                        m_ProgrammingStateFlag = False
                        'xiong==========

                    ElseIf ButtonText = "QC" Then

                        EnableTBRefers(False)
                        EnableTBElements(False)
                        EnableTBElements(gOffsetCmdIndex, False)
                        TBYesNo.Buttons(0).Enabled = False
                        TBYesNo.Buttons(1).Enabled = False

                        IDS.Devices.Vision.IDSV_Form_QC()
                        Dim status As Integer = 0
                        Dim x, y As Double
                        Dim _5dot As Boolean = False
                        While status = 0 Or status = 3
                            Do
                                Application.DoEvents()
                                If IDS.Devices.Vision.GetQCSelection = 1 Then 'dot

                                Else 'rect
                                    If status = 3 Then
                                        status = 0 'reset status after 5 points being reset
                                    End If
                                    While IDS.Devices.Vision.GetQCStatus <> 2 And IDS.Devices.Vision.GetQCSelection = 2
                                        If IDS.Devices.Vision.RobotMotionOffset(x, y) Then
                                            Application.DoEvents()
                                            _5dot = True
                                            Exit While
                                        Else
                                            Application.DoEvents()
                                        End If

                                    End While
                                    If _5dot Then
                                        'moverobot
                                        pos(0) = x
                                        pos(1) = -y
                                        pos(2) = 0
                                        Dim spstr As String = gSystemSpeed.ToString
                                        Dim str As String
                                        str = "SPEED AXIS(0) =" + spstr
                                        m_Tri.m_TriCtrl.Execute(str)
                                        str = "SPEED AXIS(1) =" + spstr
                                        m_Tri.m_TriCtrl.Execute(str)
                                        str = "SPEED AXIS(2) =" + spstr
                                        m_Tri.m_TriCtrl.Execute(str)
                                        'Dim spstr As String = gSystemSpeed.ToString
                                        'spstr = "SPEED AXIS(0) =" + spstr
                                        'm_Tri.m_TriCtrl.Execute(spstr)
                                        m_Tri.m_TriCtrl.MoveRel(2, pos, 0)
                                        m_Tri.WaitForEndOfMove()
                                        _5dot = False
                                    End If
                                End If
                                status = IDS.Devices.Vision.GetQCStatus()
                            Loop While status = 3
                        End While

                        If 2 = status Then  'Cancel
                            LabelMessege.Text = ""
                            TBElements.Enabled = True
                            TBrefer.Enabled = True
                            m_PosUpdate = False
                            AxSpreadsheetProgramming.DisplayWorkbookTabs = True
                            AxSpreadsheetProgramming.ActiveSheet.Rows(m_row).delete()
                            AxSpreadsheetProgramming.Update()
                        End If

                        If 1 = status Then                         'Ok
                            Dim bb As Boolean
                            Dim vpara As DLL_Export_Device_Vision.QC.QCParam
                            IDS.Devices.Vision.GetQCParameters(vpara)
                            IDS.Devices.Vision.SetQCReset()

                            With AxSpreadsheetProgramming.ActiveSheet
                                .Cells(m_row, gBrightnessColumn) = vpara._Brightness
                                .Cells(m_row, gBinarizedColumn) = vpara._Binarized
                                .Cells(m_row, gBlackDotColumn) = vpara._BlackDot
                                .Cells(m_row, gOpenColumn) = vpara._Open
                                .Cells(m_row, gInOutColumn) = vpara._Inside_out
                                .Cells(m_row, gCloseColumn) = vpara._Close
                                .Cells(m_row, gPointX1Column) = vpara._PointX1
                                .Cells(m_row, gPointX2Column) = vpara._PointX2
                                .Cells(m_row, gPointX3Column) = vpara._PointX3
                                .Cells(m_row, gPointX4Column) = vpara._PointX4
                                .Cells(m_row, gPointX5Column) = vpara._PointX5
                                .Cells(m_row, gPointY1Column) = vpara._PointY1
                                .Cells(m_row, gPointY2Column) = vpara._PointY2
                                .Cells(m_row, gPointY3Column) = vpara._PointY3
                                .Cells(m_row, gPointY4Column) = vpara._PointY4
                                .Cells(m_row, gPointY5Column) = vpara._PointY5
                                .Cells(m_row, gPosColumn) = vpara._Pos
                                .Cells(m_row, gPosXColumn) = vpara._PosX
                                .Cells(m_row, gPosYColumn) = vpara._PosY
                                .Cells(m_row, gROIColumn) = vpara._ROI
                                .Cells(m_row, gRotColumn) = vpara._Rot
                                .Cells(m_row, gSizeColumn) = vpara._Size
                                .Cells(m_row, gSizeXColumn) = vpara._SizeX
                                .Cells(m_row, gSizeYColumn) = vpara._SizeY
                                .Cells(m_row, gThresholdColumn) = vpara._Threshold
                                .Cells(m_row, gVerticalColumn) = vpara._Vertical
                                .Cells(m_row, gCompactnessColumn) = vpara._Compactness
                                .Cells(m_row, gMaxAreaColumn) = vpara._MaxArea
                                .Cells(m_row, gMinAreaColumn) = vpara._MinArea
                                .Cells(m_row, gMQC_OffXColumn) = vpara._MQC_OffX
                                .Cells(m_row, gMQC_OffYColumn) = vpara._MQC_OffY
                                .Cells(m_row, gMRegionXColumn) = vpara._MRegionX
                                .Cells(m_row, gMRegionYColumn) = vpara._MRegionY
                                .Cells(m_row, gMROIxColumn) = vpara._MROIx
                                .Cells(m_row, gMROIyColumn) = vpara._MROIy
                                .Cells(m_row, gRoughnessColumn) = vpara._Roughness
                                .Cells(m_row, gTypeColumn) = vpara._type
                                '.Cells(m_row, gBrightnessColumn)=vpara.

                                cell1 = .Cells(m_row, gPos1XColumn)
                                cell2 = .Cells(m_row, gPos1ZColumn)

                            End With
                            IDS.Devices.Vision.SetQCReset()
                            IDS.Devices.Vision.FrmVision.DisplayIndicator()

                            m_Execution.m_Pattern.AxSpreadsheetProgramming_CellGrey(cell1, cell2, True, AxSpreadsheetProgramming)
                            TBElements.Enabled = True
                            TBrefer.Enabled = True
                            m_PosUpdate = False
                            LabelMessege.Text = ""
                            AxSpreadsheetProgramming.DisplayWorkbookTabs = True
                            AxSpreadsheetProgramming.ActiveSheet.Cells(m_row + 1, 1).Select()
                            m_ProgrammingStateFlag = False
                        End If
                        EnableTBRefers(True)
                        EnableTBElements(True)
                        EnableTBElements(gOffsetCmdIndex, False)
                        TBYesNo.Buttons(0).Enabled = False
                        TBYesNo.Buttons(1).Enabled = True
                        'xiong end============



                    End If

                    'EnableTBrefer(True)
                    'EnableTBElements(True)
                    'EnableTBElements(gOffsetCmdIndex, False) 
                    'TBYesNo.Buttons(0).Enabled = False
                    'TBYesNo.Buttons(1).Enabled = True

                    ''SJ add end****************************************************************************
                Case "Link"

                    type = "LinkStart"
                    m_PosNo = 1
                    m_row = AxSpreadsheetProgramming.ActiveCell.Row
                    AddPatternCommand(type)
                    AxSpreadsheetProgramming.ActiveSheet.Cells(m_row + 1, 1).Select()
                    type = "LinkEnd"
                    m_row = AxSpreadsheetProgramming.ActiveCell.Row
                    AddPatternCommand(type)
                    ToolBarSwitch("YesNo")
                    m_row = AxSpreadsheetProgramming.ActiveCell.Row
                    m_PosUpdate = False

                    EnableTBElements(gLineCmdIndex, True)
                    EnableTBElements(gArcCmdIndex, True)
                    TBYesNo.Buttons(0).Enabled = False
                    TBYesNo.Buttons(1).Enabled = True
                    type = "Link"

                Case "Clean", "Purge", "GetIO", "SetIO"

                    type = ButtonText
                    m_PosNo = 1
                    m_row = AxSpreadsheetProgramming.ActiveCell.Row
                    AddPatternCommand(type)
                    'ToolBarSwitch("YesNo")
                    m_PosUpdate = False
                    EnableTBRefers(True)
                    EnableTBElements(True)
                    TBYesNo.Buttons(0).Enabled = False
                    TBYesNo.Buttons(1).Enabled = False
                    m_row = m_row + 1
                    AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, 1).Select()
                    m_ProgrammingStateFlag = False

                Case "Array"
                    TBYesNo.Buttons(0).Enabled = False
                    TBYesNo.Buttons(1).Enabled = False
                    type = ButtonText
                    AddPatternCommand(type)
                    m_row = AxSpreadsheetProgramming.ActiveCell.Row

                    m_PosUpdate = True
                    Dim ArrayDlg As New ArrayGenerate

                    ArrayDlg.SetDefaultPara(m_Execution.m_Pattern.m_CurrentDPara)
                    ArrayDlg.Show()
                    Dim DlgReturn = ArrayDlg.DialogResult()
                    Dim PointX, PointY, PointZ As Double

                    Do
                        Application.DoEvents()
                        PointX = CDbl(AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gPos1XColumn).value)
                        PointY = CDbl(AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gPos1YColumn).value)
                        PointZ = CDbl(AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gPos1ZColumn).value)

                        ArrayDlg.SetPoint(PointX, PointY, PointZ)
                        DlgReturn = ArrayDlg.DialogResult()
                    Loop While Nothing = DlgReturn

                    m_PosUpdate = False

                    If DialogResult.OK = DlgReturn Then
                        'Get "Array" data succesfully.  Data will be used to generate "Array"
                        LabelMessege.Text = "Array generating  starts"
                        Dim PatternLineRecord(3) As CIDSPattern.PatternRecord
                        Dim arraydata As ArrayData
                        arraydata.FirstX = 10.0
                        Dim dotarray As New FormArraySetting(arraydata)

                        If dotarray.ShowDialog() = DialogResult.OK Then
                            dotarray.GetArrayData(arraydata)

                            Dim cmdTmpString As String = ArrayDlg.ArrayPara1.Name

                            AxSpreadsheetProgramming_GetArrayRecord(PatternLineRecord(0), 1, ArrayDlg)
                            AxSpreadsheetProgramming_GetArrayRecord(PatternLineRecord(1), 2, ArrayDlg)
                            AxSpreadsheetProgramming_GetArrayRecord(PatternLineRecord(2), 3, ArrayDlg)

                            Dim ActSheetName As String = AxSpreadsheetProgramming.ActiveSheet.Name
                            AxSpreadsheetProgramming_GenerateArraySubSheetName(iSubSheetName, ActSheetName, cmdTmpString)

                            Dim file As New CIDSFileHandler
                            AxSpreadsheetProgramming.ActiveSheet.Rows(m_row).clear()
                            With AxSpreadsheetProgramming.ActiveSheet

                                .Cells(m_row, gCmdColumn) = "Array"
                                .Cells(m_row, gSubnameColumn) = file.ExtOnlyFromFullPath(iSubSheetName)
                                .Cells(m_row + 1, gCmdColumn).Select()
                            End With

                            AxSpreadsheetProgramming.Sheets.Add.Name = iSubSheetName

                            AxSpreadsheetProgramming.ActiveWindow.FreezePanes = False
                            AxSpreadsheetProgramming.Worksheets(iSubSheetName).Range("B1:B1").Select()
                            AxSpreadsheetProgramming.ActiveWindow.FreezePanes = True

                            'Build Sub array sheet
                            AxSpreadsheetProgramming_AddSheetRecord(PatternLineRecord(0), _
                                PatternLineRecord(1), PatternLineRecord(2), arraydata)

                            'Load sub and array data within the sub
                            AxSpreadsheetProgramming_AddSubandArrayInSub(PatternLineRecord(0).pPara.DispenseFlag)

                            EnableTBRefers(True)
                            EnableTBElements(True)
                            EnableTBElements(gOffsetCmdIndex, False)
                            TBYesNo.Buttons(0).Enabled = False
                            TBYesNo.Buttons(1).Enabled = True
                        Else
                            'Add here as delay
                            EnableTBRefers(True)
                            EnableTBElements(True)
                            EnableTBElements(gOffsetCmdIndex, False)
                            TBYesNo.Buttons(0).Enabled = False
                            TBYesNo.Buttons(1).Enabled = True
                            Application.DoEvents()
                            AxSpreadsheetProgramming.ActiveSheet.Rows(m_row).delete()
                        End If
                    Else
                        'Add here as delay
                        EnableTBRefers(True)
                        EnableTBElements(True)
                        EnableTBElements(gOffsetCmdIndex, False)
                        TBYesNo.Buttons(0).Enabled = False
                        TBYesNo.Buttons(1).Enabled = True
                        Application.DoEvents()
                        AxSpreadsheetProgramming.ActiveSheet.Rows(m_row).delete()
                    End If
                    m_ProgrammingStateFlag = False

            End Select
        Catch ex As SystemException
            MessageBox.Show(ex.ToString)
        End Try
        GC.Collect()
    End Sub


    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'spreadsheet operation ini
    '                                                                                  
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Public Sub init_spreadsheet()

        Dim Title(gMaxColumns) As String
        Title(gCmdColumn - 1) = "Command"
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
            For i = gCmdColumn To gIONumberColumn
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
            '.ActiveSheet.Range("A:AD").Select()
            '.ActiveSheet.Cells.AutoFit()
            .ActiveSheet.Cells(1, 1).Select()
        End With

        'With AxSpreadsheetProgramming.Selection
        '    .Interior.Color = RGB(0, 150, 0)
        'End With
        AxSpreadsheetProgramming.ActiveWindow.EnableResize = False
        AxSpreadsheetProgramming.DisplayWorkbookTabs = True

        'If "Main" = AxSpreadsheetProgramming.ActiveWindow.ActiveSheet.Name Then

        'AxSpreadsheetProgramming.ActiveWindow.ActiveSheet.Cells(1, 2) = "Unit"
        'AxSpreadsheetProgramming.ActiveWindow.ActiveSheet.Cells(1, 4) = "mm"
        'AxSpreadsheetProgramming.ActiveWindow.ActiveSheet.Cells(2, 1).Select()
        'End If

        ' AxSpreadsheetProgramming.ActiveSheet.Range("A" + m_row.ToString + ":" + "A" + m_row.ToString).Select()
        GC.Collect()
    End Sub


    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Initialize a sheet
    '   e: ActiveX component event handler
    '
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Private Sub AxSpreadsheetProgramming_Initialize(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AxSpreadsheetProgramming.Initialize
        init_spreadsheet()
        m_row = 1
        GC.Collect()
    End Sub

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Activiate a sheet
    '   e: ActiveX component event handler
    '
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Private Sub AxSpreadsheetProgramming_SheetActivate(ByVal sender As System.Object, ByVal e As AxOWC10.ISpreadsheetEventSink_SheetActivateEvent) Handles AxSpreadsheetProgramming.SheetActivate

        Dim sheetname As String = e.sh.Name
        init_spreadsheet()
        GC.Collect()
    End Sub


    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Check it can be copied or not
    '   AxSpreadsheet: ActiveX component
    '
    'Return: True=Can be copied, False=Cannot
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Private Function AxSpreadsheetProgramming_CanBeCopyCut(ByRef AxSpreadsheet As AxOWC10.AxSpreadsheet) As Boolean
        Dim Rtn As Boolean = False
        Dim sel As Microsoft.Office.Interop.OWC.Range = AxSpreadsheetProgramming.Selection

        Dim m_rowCount As Integer = sel.Rows.Count()
        Dim m_columnCount As Integer = sel.Columns.Count()
        Dim m_StartRow As Integer = sel.Row
        Dim m_EndRow As Integer = m_StartRow + m_rowCount - 1
        Dim SelectedPageName As String = AxSpreadsheet.ActiveWorkbook.ActiveSheet.Name()

        Dim bStartRowInLink As Boolean = AxSpreadsheetProgramming_WithinLinkRangeForCopyCut(m_StartRow)
        Dim bEndRowInLink As Boolean = AxSpreadsheetProgramming_WithinLinkRangeForCopyCut(m_EndRow)
        Dim bStartEndRowInTheSameLink As Boolean


        If 1 = m_rowCount Then
            If "LinkStart" = AxSpreadsheetProgramming.ActiveSheet.Cells(m_StartRow, gCmdColumn).value() Or _
                "LinkEnd" = AxSpreadsheetProgramming.ActiveSheet.Cells(m_StartRow, gCmdColumn).value() Then
                Return Rtn
            End If
        End If

        If gMaxColumns <> m_columnCount Then
            Return Rtn
        End If

        If bStartRowInLink And bEndRowInLink Then       'May possible be copied or cut, need to do further chk
            Rtn = AxSpreadsheetProgramming_WithinLinkRangeForCopyCut(m_StartRow, m_EndRow)  'Is a identity link

        ElseIf bStartRowInLink Then                     'may not be copied and cut
            If "LinkStart" <> AxSpreadsheetProgramming.ActiveSheet.Cells(m_StartRow, gCmdColumn).value() Then
                Rtn = False
            End If

        ElseIf bEndRowInLink Then                       'may not be copied and cut
            If "LinkEnd" <> AxSpreadsheetProgramming.ActiveSheet.Cells(m_EndRow, gCmdColumn).value() Then
                Rtn = False
            Else
                Rtn = True
            End If
        Else                                            'Can be copied or cut
            Rtn = True
        End If

        GC.Collect()
        Return Rtn

    End Function


    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Handle Del Key
    '   e: ActiveX component event handler
    '
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim cellString As String
    Private Sub AxSpreadsheetProgramming_BeforeButtonDelInput(ByVal sender As Object, _
           ByVal e As AxOWC10.ISpreadsheetEventSink_BeforeKeyDownEvent) Handles AxSpreadsheetProgramming.BeforeKeyDown

        Dim sel As Microsoft.Office.Interop.OWC.Range = AxSpreadsheetProgramming.Selection

        Dim m_rowCount As Integer = sel.Rows.Count()
        Dim m_columnCount As Integer = sel.Columns.Count()
        Dim row As Integer = AxSpreadsheetProgramming.ActiveCell.Row
        Dim colum As Integer = AxSpreadsheetProgramming.ActiveCell.Column

        If e.keyCode = 46 Then  'the delete key 
            If m_columnCount = gMaxColumns And m_rowCount = 1 Then
            Else
                cellString = AxSpreadsheetProgramming.ActiveSheet.Cells(row, gSubnameColumn).Value()
            End If
        End If
        GC.Collect()
    End Sub


    Private Sub AxSpreadsheetProgramming_AfterButtonDelInput(ByVal sender As Object, _
           ByVal e As AxOWC10.ISpreadsheetEventSink_KeyDownEvent) Handles AxSpreadsheetProgramming.KeyDownEvent

        Dim sel As Microsoft.Office.Interop.OWC.Range = AxSpreadsheetProgramming.Selection

        Dim m_rowCount As Integer = sel.Rows.Count()
        Dim m_columnCount As Integer = sel.Columns.Count()
        Dim row As Integer = AxSpreadsheetProgramming.ActiveCell.Row
        Dim colum As Integer = AxSpreadsheetProgramming.ActiveCell.Column
        Dim cellValue As String

        If e.keyCode = 46 Then  'the delete key 
            If m_columnCount = gMaxColumns And m_rowCount = 1 Then
                'ConfirmOrCancel_Implementation(0)
            Else
                AxSpreadsheetProgramming.ActiveSheet.Cells(row, gSubnameColumn).Value() = cellString
            End If
        End If
        GC.Collect()
    End Sub

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Start cell edit, tentative result may be rejected
    '   e: ActiveX component event handler
    '
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Private Sub AxSpreadsheetProgramming_StartEdit(ByVal sender As Object, _
           ByVal e As AxOWC10.ISpreadsheetEventSink_StartEditEvent) Handles AxSpreadsheetProgramming.StartEdit

        Dim row As Integer = AxSpreadsheetProgramming.ActiveCell.Row
        Dim colum As Integer = AxSpreadsheetProgramming.ActiveCell.Column
        Dim cellValue As Object = AxSpreadsheetProgramming.ActiveCell.Value
        'Dim InputStr As String
        m_PosUpdate = False                 ' (row <> m_row) Or _

        'InputStr = AxSpreadsheetProgramming.ActiveSheet.Cells(row, colum).value

        'If "" = InputStr Then
        'LabelMessege.Text = ""
        'AxSpreadsheetProgramming.ActiveSheet.Cells(row, colum).Value = cellValue
        If (colum >= gCmdColumn And colum <= gDispensColumn) Then
            If colum = gCmdColumn Then
                If "SubPattern" = cellValue Or "Array" = cellValue Then
                    Dim SubSheetName As String
                    SubSheetName = AxSpreadsheetProgramming.ActiveSheet.Cells(row, gSubnameColumn).Value
                    AxSpreadsheetProgramming.ActiveCell.Value = cellValue
                    AxSpreadsheetProgramming.Worksheets(SubSheetName).Activate()
                    Exit Sub
                Else
                    AxSpreadsheetProgramming.ActiveSheet.Cells(row, colum).Value = cellValue
                End If
            ElseIf (colum >= gNeedleColumn And colum <= gDispensColumn) Then
                LabelMessege.Text = "Can't change value directly"
                AxSpreadsheetProgramming.ActiveSheet.Cells(row, colum).Value = cellValue
            End If

        ElseIf ((colum >= gPos1XColumn And colum <= gPos3ZColumn) And (m_KeyinOK = False)) Then
            LabelMessege.Text = "Can't change value directly"
            '           System.Windows.Forms.MessageBox.Show("not allowed to edit", "Add Dot")
            AxSpreadsheetProgramming.ActiveSheet.Cells(row, colum).Value = cellValue
        Else
            LabelMessege.Text = "Please input a new value"
            If m_PosNo = 1 Then

                If (colum >= gPos1XColumn And colum <= gPos1ZColumn) Then


                Else
                    'LabelMessege.Text = "Can't change value here"
                    'AxSpreadsheetProgramming.ActiveSheet.Cells(row, colum).Value = cellValue
                End If

            ElseIf m_PosNo = 2 Then
                If (colum >= gPos2XColumn And colum <= gPos2ZColumn) Then


                Else
                    'LabelMessege.Text = "Can't change value here"
                    'AxSpreadsheetProgramming.ActiveSheet.Cells(row, colum).Value = cellValue
                End If
            ElseIf m_PosNo = 3 Then
                If (colum >= gPos3XColumn And colum <= gPos3ZColumn) Then


                Else
                    'LabelMessege.Text = "Can't change value here"
                    'AxSpreadsheetProgramming.ActiveSheet.Cells(row, colum).Value = cellValue
                End If
            End If
        End If

        GC.Collect()
    End Sub

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'End cell edit, result will be finalized
    '   e: ActiveX component event handler
    '
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Private Sub AxSpreadsheetProgramming_EndEdit(ByVal sender As Object, _
         ByVal e As AxOWC10.ISpreadsheetEventSink_EndEditEvent) Handles AxSpreadsheetProgramming.EndEdit
        Dim row As Integer = AxSpreadsheetProgramming.ActiveCell.Row
        Dim colum As Integer = AxSpreadsheetProgramming.ActiveCell.Column
        Dim dVal As Double

        Dim ErrorFound As Boolean = False

        Dim CmdStr As String = AxSpreadsheetProgramming.ActiveSheet.Cells(row, gCmdColumn).value
        Dim ActSheetName As String = AxSpreadsheetProgramming.ActiveSheet.Name
        Dim InputStr As String = e.finalValue.Value.ToString()

        If (colum >= gPos1XColumn And colum <= gPos1YColumn) Or _
            (colum >= gPos2XColumn And colum <= gPos2YColumn) Or _
            (colum >= gPos3XColumn And colum <= gPos3YColumn) Then ' And _
            '       (row > 0) And (m_KeyinOK = True) Then

            ErrorFound = m_Execution.m_Pattern.m_ErrorChk.CheckCellXYError( _
                ActSheetName, InputStr, row, colum, AxSpreadsheetProgramming, ErrorSubSheet)

            If ErrorFound Then
                MsgBox("Found edit error, click Ok to restore", MsgBoxStyle.Information + MsgBoxStyle.OKOnly, "Input error")
                e.cancel.Value = True
            Else
                m_Execution.m_Pattern.UpdateDispara(m_Execution.m_Pattern.m_CurrentDPara, InputStr, colum)
                'Move to the place
                Dim pos(3) As Double
                Try
                    Select Case colum
                        Case gPos1XColumn
                            pos(0) = CDbl(InputStr)
                            pos(1) = CDbl(AxSpreadsheetProgramming.ActiveSheet.Cells(row, gPos1YColumn).value)
                            pos(2) = CDbl(AxSpreadsheetProgramming.ActiveSheet.Cells(row, gPos1ZColumn).value)
                        Case gPos1YColumn
                            pos(0) = CDbl(AxSpreadsheetProgramming.ActiveSheet.Cells(row, gPos1XColumn).value)
                            pos(1) = CDbl(InputStr)
                            pos(2) = CDbl(AxSpreadsheetProgramming.ActiveSheet.Cells(row, gPos1ZColumn).value)
                        Case gPos2XColumn
                            pos(0) = CDbl(InputStr)
                            pos(1) = CDbl(AxSpreadsheetProgramming.ActiveSheet.Cells(row, gPos2YColumn).value)
                            pos(2) = CDbl(AxSpreadsheetProgramming.ActiveSheet.Cells(row, gPos2ZColumn).value)
                        Case gPos2YColumn
                            pos(0) = CDbl(AxSpreadsheetProgramming.ActiveSheet.Cells(row, gPos2XColumn).value)
                            pos(1) = CDbl(InputStr)
                            pos(2) = CDbl(AxSpreadsheetProgramming.ActiveSheet.Cells(row, gPos2ZColumn).value)
                        Case gPos3XColumn
                            pos(0) = CDbl(InputStr)
                            pos(1) = CDbl(AxSpreadsheetProgramming.ActiveSheet.Cells(row, gPos3YColumn).value)
                            pos(2) = CDbl(AxSpreadsheetProgramming.ActiveSheet.Cells(row, gPos3ZColumn).value)
                        Case gPos3YColumn
                            pos(0) = CDbl(AxSpreadsheetProgramming.ActiveSheet.Cells(row, gPos3XColumn).value)
                            pos(1) = CDbl(InputStr)
                            pos(2) = CDbl(AxSpreadsheetProgramming.ActiveSheet.Cells(row, gPos3ZColumn).value)
                    End Select
                    ''''''SJ change
                    'Trimotion_MoveTo(pos)
                    If CmdStr = "ChipEdge" Or CmdStr = "QC" Or CmdStr = "Height" Or CmdStr = "Fiducial" Or CmdStr = "Reject" Then
                        Trimotion_MoveTo(pos)
                    Else
                        Trimotion_MoveTo(pos, m_TeachMode)
                    End If
                Catch ex As Exception
                    LabelMessege.Text = "invalid value"
                    e.cancel.Value = True 'Cancel the edit
                End Try
            End If


            'Dim posX, posY, posZ As Double
            'm_Tri.m_TriCtrl.GetVr(1, posX)
            'm_Tri.m_TriCtrl.GetVr(2, posY)
            'm_Tri.m_TriCtrl.GetVr(3, posZ)
            ''''''''''''''temp use for z input
            'm_PosUpdate = True
            'm_PosUpdate = False
            'Dim z As Double = gSoftHome(2) + gSystemOrigin(2)
            'Select Case colum
            '    Case gPos1XColumn, gPos2XColumn, gPos3XColumn
            '        'retract to clearernce height
            '        m_Tri.m_TriCtrl.MoveAbs(1, z, 2)
            '        dVal = dVal * gmminchRatio
            '        m_Tri.m_TriCtrl.MoveAbs(1, dVal, 0)
            '    Case gPos1YColumn, gPos2YColumn, gPos3YColumn
            '        m_Tri.m_TriCtrl.MoveAbs(1, z, 2)
            '        dVal = dVal * gmminchRatio
            '        m_Tri.m_TriCtrl.MoveAbs(1, dVal, 1)
            'End Select


            'Try

            '    dVal = System.Double.Parse(e.finalValue.Value.ToString())
            '    'test
            '    If (dVal < -100) Or (dVal > 100) Then
            '        ' Value not between -100 and 100.
            '        LabelMessege.Text = "Must be between -100 to 100"
            '        e.cancel.Value = True 'Cancel the edit
            '        Exit Sub
            '    End If
            'Catch
            '    ' Cannot convert to a double.

            '    LabelMessege.Text = "invalid value"
            '    e.cancel.Value = True 'Cancel the edit
            'End Try

        Else

            ErrorFound = m_Execution.m_Pattern.m_ErrorChk.CheckOtherCellCharError(CmdStr, InputStr, colum)

            If Not ErrorFound Then
                If colum = gPos1ZColumn Or colum = gPos2ZColumn Or colum = gPos3ZColumn Or colum = gClearanceHtColumn Or _
                    colum = gRetractHtColumn Or colum = gNeedleGapColumn Then

                    If m_Execution.m_Pattern.m_ErrorChk.CheckRequiredPtHeight(CmdStr) Then
                        Dim CurrentHeight, fClearanceHeight, fRetractHeight, fNeedleGap As Double

                        With AxSpreadsheetProgramming.ActiveSheet
                            CurrentHeight = CDbl(.Cells(row, gApproachHtColumn).value)
                            fClearanceHeight = CDbl(.Cells(row, gClearanceHtColumn).value)
                            fRetractHeight = CDbl(.Cells(row, gRetractHtColumn).value)
                            fNeedleGap = CDbl(.Cells(row, gNeedleGapColumn).value)
                        End With

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
                MsgBox("Found edit error, click Ok to restore", MsgBoxStyle.Information + MsgBoxStyle.OKOnly, "Input error")
                e.cancel.Value = True
            Else
                m_Execution.m_Pattern.UpdateDispara(m_Execution.m_Pattern.m_CurrentDPara, InputStr, colum)
            End If
        End If

        GC.Collect()
    End Sub

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'set motion reference point
    '   RefPos: input ref point
    '
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Private Sub AxSpreadsheetProgramming_SetMotionRef(ByVal RefPos() As Double)
        m_ReferPt(0) = RefPos(0)
        m_ReferPt(1) = RefPos(1)
        m_ReferPt(2) = RefPos(2)
    End Sub


    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Check the keyboard input event, detail Buttons and flags will be reset accordingly
    '   e: ActiveX component event handler
    '
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Private Sub AxSpreadsheetProgramming_KeyUpEvent(ByVal sender As Object, _
    '        ByVal e As System.EventArgs) Handles AxSpreadsheetProgramming.KeyUpEvent



    'End Sub

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Check the click event, detail Buttons and flags will be reset accordingly
    '   e: ActiveX component event handler
    '
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Private Sub AxSpreadsheetProgramming_ClickEvent(ByVal sender As Object, _
            ByVal e As System.EventArgs) Handles AxSpreadsheetProgramming.ClickEvent

        Dim sel As Microsoft.Office.Interop.OWC.Range = AxSpreadsheetProgramming.Selection

        Dim m_rowCount As Integer = sel.Rows.Count()
        Dim m_columnCount As Integer = sel.Columns.Count()
        Dim m_StartRow As Integer = sel.Row
        Dim m_columnStart As Integer = sel.Column
        Dim m_EndRow As Integer = m_StartRow + m_rowCount - 1
        Dim RefPos() As Double = {0, 0, 0}

        If "" = gPatternFileName Or 1 = m_cmdCompleted Then
            Exit Sub
        End If

        Dim CellStr As String
        Dim CmdStr As String

        Dim cell1, cell2 As Object
        Dim inCurrentEditSlot As Boolean = True
        Dim ReselectCells As Boolean = False

        If True = m_EditStateFlag Then
            MenuEditSelectAll.Enabled = False
            CellStr = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, m_columnStart).Value
            CmdStr = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gCmdColumn).Value
            If m_PosNo = 1 Then
                If m_columnStart < gPos1XColumn Or m_columnStart > gPos1ZColumn Then
                    inCurrentEditSlot = False
                End If
                If m_Execution.m_Pattern.m_ErrorChk.CheckRequiredPt1XY(CmdStr) Then
                    cell1 = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gPos1XColumn)
                    cell2 = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gPos1ZColumn)
                    ReselectCells = True
                End If

            ElseIf m_PosNo = 2 Then
                If m_columnStart < gPos2XColumn Or m_columnStart > gPos2ZColumn Then
                    inCurrentEditSlot = False
                End If
                If m_Execution.m_Pattern.m_ErrorChk.CheckRequiredPt2XY(CmdStr) Then
                    cell1 = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gPos2XColumn)
                    cell2 = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gPos2ZColumn)
                    ReselectCells = True
                End If

            ElseIf m_PosNo = 3 Then
                If m_columnStart < gPos3XColumn Or m_columnStart > gPos3ZColumn Then
                    inCurrentEditSlot = False
                End If
                If m_Execution.m_Pattern.m_ErrorChk.CheckRequiredPt3XY(CmdStr) Then
                    cell1 = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gPos3XColumn)
                    cell2 = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gPos3ZColumn)
                    ReselectCells = True
                End If

            End If

            If m_row <> m_StartRow Or inCurrentEditSlot = False Then
                If ReselectCells = True Then
                    AxSpreadsheetProgramming.ActiveSheet.Range(cell1, cell2).Select()
                End If
                'AxSpreadsheetProgramming.DisplayWorkbookTabs = True

                Exit Sub
            End If

        ElseIf True = m_ProgrammingStateFlag Then
            MenuEditSelectAll.Enabled = False
            CmdStr = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gCmdColumn).Value
            If "LinkEnd" = CmdStr Then
                inCurrentEditSlot = False
                cell1 = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gCmdColumn)
                cell2 = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gCmdColumn)
                ReselectCells = True
            ElseIf m_PosNo = 1 Then
                If m_columnStart < gPos1XColumn Or m_columnStart > gPos1ZColumn Then
                    inCurrentEditSlot = False
                End If
                If m_Execution.m_Pattern.m_ErrorChk.CheckRequiredPt1XY(CmdStr) Then
                    cell1 = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gPos1XColumn)
                    cell2 = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gPos1ZColumn)
                    ReselectCells = True
                End If

            ElseIf m_PosNo = 2 Then
                If m_columnStart < gPos2XColumn Or m_columnStart > gPos2ZColumn Then
                    inCurrentEditSlot = False
                End If
                If m_Execution.m_Pattern.m_ErrorChk.CheckRequiredPt2XY(CmdStr) Then
                    cell1 = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gPos2XColumn)
                    cell2 = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gPos2ZColumn)
                    ReselectCells = True
                End If

            ElseIf m_PosNo = 3 Then
                If m_columnStart < gPos3XColumn Or m_columnStart > gPos3ZColumn Then
                    inCurrentEditSlot = False
                End If
                If m_Execution.m_Pattern.m_ErrorChk.CheckRequiredPt3XY(CmdStr) Then
                    cell1 = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gPos3XColumn)
                    cell2 = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gPos3ZColumn)
                    ReselectCells = True
                End If

            End If

            If m_row <> m_StartRow Or inCurrentEditSlot = False Then
                If m_row <> m_StartRow Or ReselectCells = True Then
                    AxSpreadsheetProgramming.ActiveSheet.Range(cell1, cell2).Select()
                End If
                'AxSpreadsheetProgramming.DisplayWorkbookTabs = True
                Exit Sub
            End If
        End If



        m_row = sel.Row
        SheetRowSelected = False
        m_PosUpdate = False
        MenuEditCopy.Enabled = False
        CellStr = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, m_columnStart).Value
        CmdStr = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gCmdColumn).Value

        'Set motion ref for any single row selected
        If 1 = m_rowCount Then
            m_Execution.m_Pattern.AxSpreadsheetProgramming_GetRowLocalReference(AxSpreadsheetProgramming, m_row, RefPos)
            AxSpreadsheetProgramming_SetMotionRef(RefPos)
        End If

        If gMaxRows = m_rowCount And gMaxColumns = m_columnCount Then
            Dim SubSheetName1 As String
            Dim SubSheetName2 As String
            SubSheetName1 = AxSpreadsheetProgramming.ActiveSheet.Name
            SubSheetName2 = m_Execution.m_Pattern.AxSpreadsheetProgramming_FindParrentSheet(SubSheetName1)

            AxSpreadsheetProgramming.Worksheets(SubSheetName2).Activate()
            AxSpreadsheetProgramming.DisplayWorkbookTabs = True
        ElseIf 1 = m_rowCount And 1 = m_columnCount Then
            Dim SubSheetName As String
            If "Array" = CellStr Or "SubPattern" = CellStr Then
                If "SubPattern" = CellStr Then
                    SubSheetName = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gSubnameColumn).Value
                    SubSheetName = m_Execution.m_File.NameOnlyFromFullPath(SubSheetName)
                ElseIf "Array" = CellStr Then
                    SubSheetName = AxSpreadsheetProgramming.ActiveSheet.Name + "."
                    SubSheetName = SubSheetName + AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gSubnameColumn).Value
                End If

                AxSpreadsheetProgramming.Worksheets(SubSheetName).Activate()

                EnableTBElements(gOffsetCmdIndex, False)
                EnableTBNextEdit(False)
            ElseIf m_Execution.m_Pattern.m_ErrorChk.CheckValidPtXY(CmdStr) And (m_KeyinOK = True) Then
                If m_Execution.m_Pattern.m_ErrorChk.CheckRequiredPt2XY(CmdStr) Then
                    EnableTBNextEdit(0, True)                   'Goto next
                Else
                    EnableTBNextEdit(0, False)
                End If

                EnableTBNextEdit(1, False)                      'Edit start
                TBYesNo.Buttons(1).Enabled = True
            Else
                EnableTBElements(gOffsetCmdIndex, False)
                EnableTBNextEdit(False)
            End If

        ElseIf gMaxColumns = m_columnCount Then
            Dim pos(3) As Double

            MenuEditDelete.Enabled = True
            TBYesNo.Buttons(0).Enabled = False
            TBYesNo.Buttons(1).Enabled = True

            'Check it can be copied or not
            If AxSpreadsheetProgramming_CanBeCopyCut(AxSpreadsheetProgramming) Then
                MenuEditCopy.Enabled = True
                MenuEditCut.Enabled = True
            Else
                MenuEditCopy.Enabled = False
                MenuEditCut.Enabled = False
            End If

            If 1 = m_rowCount Then
                SheetRowSelected = True
                AxSpreadsheetProgramming.DisplayWorkbookTabs = True

                If m_Execution.m_Pattern.m_ErrorChk.CheckValidPtXY(CmdStr) Then
                    m_PosNo = 1
                    pos(0) = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gPos1XColumn).Value
                    pos(1) = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gPos1YColumn).Value
                    pos(2) = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gPos1ZColumn).Value
                    pos(0) = pos(0) * gmminchRatio
                    pos(1) = pos(1) * gmminchRatio
                    pos(2) = pos(2) * gmminchRatio
                    ''''''SJ change
                    'Trimotion_MoveTo(pos)
                    If CmdStr = "ChipEdge" Or CmdStr = "QC" Or CmdStr = "Height" Or CmdStr = "Fiducial" Or CmdStr = "Reject" Then
                        Trimotion_MoveTo(pos)
                    Else
                        Trimotion_MoveTo(pos, m_TeachMode)
                    End If
                    AxSpreadsheetProgramming.DisplayWorkbookTabs = False

                    EnableTBElements(False)
                    EnableTBElements(gOffsetCmdIndex, True)         'Offset
                    EnableTBNextEdit(0, True)                       'Goto next
                    EnableTBNextEdit(1, False)                      'Edit start

                ElseIf "" = CmdStr Then
                    EnableTBElements(gOffsetCmdIndex, False)
                    EnableTBNextEdit(False)
                    AxSpreadsheetProgramming.DisplayWorkbookTabs = True

                Else
                    EnableTBElements(gOffsetCmdIndex, False)
                    EnableTBNextEdit(False)

                End If
                AxSpreadsheetProgramming_CheckForWithinLinkRange(True)
            Else
                'For multi-row OFFSET selection, we need to find the last valid point first.  
                EnableTBRefers(False)
                EnableTBElements(False)

                If m_Execution.m_Pattern.AxSpreadsheetProgramming_FindLastValidPoint(AxSpreadsheetProgramming, m_StartRow, m_EndRow, pos) Then
                    EnableTBElements(gOffsetCmdIndex, True)   'Offset enabled

                    'For the valid Offset selections, we get the motion ref for the last row
                    m_Execution.m_Pattern.AxSpreadsheetProgramming_GetRowLocalReference(AxSpreadsheetProgramming, m_EndRow, RefPos)
                    AxSpreadsheetProgramming_SetMotionRef(RefPos)

                    'Move robot arm to the point last valid point
                    pos(0) = pos(0) * gmminchRatio
                    pos(1) = pos(1) * gmminchRatio
                    pos(2) = pos(2) * gmminchRatio
                    Trimotion_MoveTo(pos, False)
                Else
                    'For the invalid Offset selections, we get the motion ref for the first row
                    m_Execution.m_Pattern.AxSpreadsheetProgramming_GetRowLocalReference(AxSpreadsheetProgramming, m_row, RefPos)
                    AxSpreadsheetProgramming_SetMotionRef(RefPos)

                    EnableTBElements(gOffsetCmdIndex, False)
                End If
                EnableTBNextEdit(False)
            End If

        ElseIf 1 = m_columnCount Then   'Update elements column-wise, need to be improved later to
            m_column = sel.Column       'exclude part of data
            TBYesNo.Buttons(0).Enabled = False
            TBYesNo.Buttons(1).Enabled = False
            EnableTBNextEdit(False)
            EnableTBElements(False)
            EnableTBRefers(False)

            If m_column < gTravelSpeedColumn Or m_column > gSprialColumn Or m_rowCount < 1 Then
                Exit Sub
            End If
            If "" = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gCmdColumn).Value Or _
                "" = CStr(AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, m_column).Value) Then
                Exit Sub
            End If

            Dim response As MsgBoxResult
            response = MsgBox("Update the column based on the first valid value?", MsgBoxStyle.YesNo)

            If response = MsgBoxResult.Yes Then
                If m_rowCount > 1 Then
                    AxSpreadsheetProgramming_UpdateColumn(m_row, m_column, m_rowCount, AxSpreadsheetProgramming)
                Else
                    MsgBox("Update column is not allowed", MsgBoxStyle.OKOnly)
                End If
            End If
        ElseIf 1 = m_rowCount Then
            m_Execution.m_Pattern.AxSpreadsheetProgramming_GetRowLocalReference(AxSpreadsheetProgramming, m_row, RefPos)
            AxSpreadsheetProgramming_SetMotionRef(RefPos)

            If 3 = m_columnCount Then
                If m_Execution.m_Pattern.m_ErrorChk.CheckValidPtXY(CmdStr) Then

                End If
            End If
        Else
            m_Execution.m_Pattern.AxSpreadsheetProgramming_GetRowLocalReference(AxSpreadsheetProgramming, m_row, RefPos)
            AxSpreadsheetProgramming_SetMotionRef(RefPos)

            EnableTBNextEdit(False)
            EnableTBElements(False)
            EnableTBRefers(False)
        End If
        GC.Collect()
    End Sub



    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Check the select row is valid for Copy/cut or not                
    '   Axspreadsheet:  instance of activeX spreadsheet               
    '   m_LocalRow:     The row which will be checked                 
    '                                                                 
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Private Function AxSpreadsheetProgramming_WithinLinkRangeForCopyCut(ByVal m_LocalRow As Integer) As Boolean
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
            m_LocalType = AxSpreadsheetProgramming.ActiveSheet.Cells(m_LocalRow, gCmdColumn).value()

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
                    m_LocalType = AxSpreadsheetProgramming.ActiveSheet.Cells(jj, gCmdColumn).value()

                    If "LinkEnd" = m_LocalType Then
                        Exit Try
                    ElseIf "LinkStart" = m_LocalType Then
                        Rtn = True
                        Exit For
                    End If
                Next
                'Forword search
                For jj = m_LocalRow To m_LocalRow + 250
                    m_LocalType = AxSpreadsheetProgramming.ActiveSheet.Cells(jj, gCmdColumn).value()
                    If "LinkStart" = m_LocalType Then
                        Exit Try
                    ElseIf "LinkEnd" = m_LocalType Then
                        Rtn = True
                        Exit For
                    End If
                Next
            End If

        Catch ex As SystemException
            MessageBox.Show(ex.ToString)
        End Try

        Return (Rtn)
        GC.Collect()
    End Function

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Check the select range is a link identity or not                 
    '   Axspreadsheet:  instance of activeX spreadsheet               
    '   m_LocalStartRow:  The start row of the range                  
    '   m_LocalEndRow:    The end row of the range                    
    '                                                                 
    '   Return Rtn=True if it is a identity                           
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Private Function AxSpreadsheetProgramming_WithinLinkRangeForCopyCut( _
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
                m_LocalType = AxSpreadsheetProgramming.ActiveSheet.Cells(jj, gCmdColumn).value()
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
            MessageBox.Show(ex.ToString)
        End Try

        Return (Rtn)
        GC.Collect()
    End Function


    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Check the current row is in the link range or not                                               
    '   UpdateToolBar:      To indicate the toolbar enabled/disabled will be updated or not
    '                       True=will be updated, False=not be updated
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Private Sub AxSpreadsheetProgramming_CheckForWithinLinkRange(ByVal UpdateToolBar As Boolean)
        'If selected row is within Link range, part of Element tool icon should be disabled.
        'For fast speed, only +- 250 line will be checked, agreed by EET on 21/11/2007

        Dim CurrentOffsetButtonStare As Boolean = TBElements.Buttons(gOffsetCmdIndex - 1).Enabled
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

            Dim m_LocalType As String = AxSpreadsheetProgramming.ActiveSheet.Cells(m_LocalRow, gCmdColumn).value()

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
                    m_LocalType = AxSpreadsheetProgramming.ActiveSheet.Cells(jj, gCmdColumn).value()

                    If "LinkEnd" = m_LocalType Then
                        If UpdateToolBar Then
                            EnableTBRefers(True)
                            EnableTBElements(True)
                        End If
                        Exit Sub
                    ElseIf "LinkStart" = m_LocalType Then
                        m_InLinkRange = True
                        Exit For
                    End If
                Next
                'Forword search
                For jj = m_LocalRow To m_LocalRow + 250
                    m_LocalType = AxSpreadsheetProgramming.ActiveSheet.Cells(jj, gCmdColumn).value()
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
                    EnableTBRefers(False)
                    EnableTBElements(False)
                    EnableTBElements(gLineCmdIndex, True)
                    EnableTBElements(gArcCmdIndex, True)
                Else
                    EnableTBRefers(True)
                    EnableTBElements(True)
                    EnableTBElements(gOffsetCmdIndex, CurrentOffsetButtonStare)
                End If
            End If
        Catch ex As SystemException
            MessageBox.Show(ex.ToString)
        End Try
        GC.Collect()
    End Sub


    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Update one selected column in spreadsheet                                               
    '   AxSpreadsheet:      ActiveX spreadsheet                                                    
    '   startRow:           Starting rows
    '   m_rowCount:         Rows to be updated  
    '   SelectedColumn:     Selected column 
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Private Sub AxSpreadsheetProgramming_UpdateColumn(ByVal startRow As Integer, ByVal SelectedColumn As Integer, _
      ByVal m_rowCount As Integer, ByRef AxSpreadsheet As AxOWC10.AxSpreadsheet)
        Dim emptyLinCount As Integer = 0
        Dim i, j As Integer
        Dim StrTmp As String
        UndoData_Logging(0)

        i = startRow
        With AxSpreadsheet.ActiveSheet
            Do
                StrTmp = .Cells(i, gCmdColumn).Value
                If "" = StrTmp Then
                    emptyLinCount = emptyLinCount + 1
                Else

                    m_Execution.m_ItemLUT.Cmd2Index(StrTmp.ToUpper)
                    If m_Execution.m_ItemLUT.GetFlag(SelectedColumn) Then
                        .Cells(i, SelectedColumn).Value = .Cells(startRow, SelectedColumn).Value
                    End If

                    emptyLinCount = 0
                End If
                i = i + 1
            Loop Until emptyLinCount > 20 Or i > m_rowCount + startRow - 1
        End With
        GC.Collect()
    End Sub

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Delete multi-row                                                
    '   AxSpreadsheet:      ActiveX spreadsheet                                                                                                              
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Public Sub AxSpreadsheetProgramming_DeleteMultiRow(ByRef AxSpreadsheet As AxOWC10.AxSpreadsheet)
        Dim sel As Microsoft.Office.Interop.OWC.Range = AxSpreadsheet.Selection
        Dim m_rowCount As Integer = sel.Rows.Count()
        Dim m_row As Integer = sel.Row
        Dim Type As String
        Dim DeltedRowNo As Integer = 0
        Dim DeltedRowExpected As Integer = -1

        Dim UsedRowNumber1 As Integer = AxSpreadsheet.ActiveSheet.UsedRange.Rows.Count
        Dim UsedRowNumber2 As Integer

        If m_rowCount < 1 Then Return

        Do
            AxSpreadsheetProgramming_DeleteOneRow(m_row, AxSpreadsheet)

            UsedRowNumber2 = AxSpreadsheet.ActiveSheet.UsedRange.Rows.Count
            DeltedRowNo = UsedRowNumber1 - UsedRowNumber2
            If UsedRowNumber2 = 1 Then
                DeltedRowExpected += 1
            End If
        Loop Until (DeltedRowNo >= m_rowCount Or DeltedRowExpected = 1)
    End Sub


    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Delete one row                                                
    '   AxSpreadsheet:      ActiveX spreadsheet                                                    
    '   row:                Row to be deleted                                                             
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Public Sub AxSpreadsheetProgramming_DeleteOneRow(ByVal row As Integer, ByRef AxSpreadsheet As AxOWC10.AxSpreadsheet)

        Dim m_type As String = AxSpreadsheet.ActiveSheet.Cells(m_row, gCmdColumn).Value
        If "" = m_type Then
            AxSpreadsheet.ActiveSheet.Rows(m_row).Delete()

        ElseIf "SUBPATTERN" = m_type.ToUpper Or "Array" = m_type Then
            Dim SubSheetName As String
            Dim CurrentSheetName As String

            SubSheetName = AxSpreadsheet.ActiveSheet.Cells(m_row, gSubnameColumn).Value
            CurrentSheetName = AxSpreadsheet.ActiveSheet.Name

            'If it is also used by others, the sub with other subs attached to it will not be deleted.
            'If 0 = AxSpreadsheetProgramming_QuickCountExistedSubsheet(AxSpreadsheetProgramming, SubSheetName) Then
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

            AxSpreadsheetProgramming_FindSubTotalUsage(AxSpreadsheetProgramming, SubSheetName)
            AxSpreadsheetProgramming_BuildRootConnectedSub(AxSpreadsheetProgramming, SubSheetName)
            AxSpreadsheetProgramming_FindSubSelfUsage(AxSpreadsheetProgramming, SubSheetName)
            AxSpreadsheetProgramming_FlatternTotalUsage(AxSpreadsheetProgramming, SubSheetName)
            AxSpreadsheetProgramming_DeleteSubSheet(AxSpreadsheetProgramming)

            AxSpreadsheet.ActiveSheet.Rows(m_row).Delete()
        ElseIf "Line" = m_type Or "Arc" = m_type Then
            AxSpreadsheetProgramming_CheckForWithinLinkRange(False)
            If m_InLinkRange Then
                With AxSpreadsheet.ActiveSheet
                    If "Link" = .Cells(m_row - 1, gCmdColumn).Value() Then

                    ElseIf "Line" = .Cells(m_row - 1, gCmdColumn).Value() Then
                        If "Line" = .Cells(m_row + 1, gCmdColumn).Value() Or _
                            "Arc" = .Cells(m_row + 1, gCmdColumn).Value() Then
                            .Cells(m_row + 1, gPos1XColumn) = .Cells(m_row - 1, gPos2XColumn)
                            .Cells(m_row + 1, gPos1YColumn) = .Cells(m_row - 1, gPos2YColumn)
                            .Cells(m_row + 1, gPos1ZColumn) = .Cells(m_row - 1, gPos2ZColumn)
                        End If
                    ElseIf "Arc" = .Cells(m_row - 1, gCmdColumn).Value() Then
                        If "Line" = .Cells(m_row + 1, gCmdColumn).Value() Or _
                            "Arc" = .Cells(m_row + 1, gCmdColumn).Value() Then
                            .Cells(m_row + 1, gPos1XColumn) = .Cells(m_row - 1, gPos3XColumn)
                            .Cells(m_row + 1, gPos1YColumn) = .Cells(m_row - 1, gPos3YColumn)
                            .Cells(m_row + 1, gPos1ZColumn) = .Cells(m_row - 1, gPos3ZColumn)
                        End If
                    End If
                End With
            End If
            AxSpreadsheet.ActiveSheet.Rows(m_row).Delete()
        ElseIf "LINKSTART" = m_type.ToUpper Or "LINKEND" = m_type.ToUpper Then
            Dim ii As Integer = 0
            If "LINKSTART" = m_type.ToUpper Then
                Do
                    type = AxSpreadsheet.ActiveSheet.Cells(m_row, gCmdColumn).Value()
                    AxSpreadsheet.ActiveSheet.Rows(m_row).Delete()
                    ii = ii + 1
                    If Nothing = type Then
                        type = "TMPSTRING"
                    End If
                Loop Until "LINKEND" = m_type.ToUpper Or ii > 500
                AxSpreadsheet.ActiveSheet.Cells(m_row, 1).Select()

            ElseIf "LINKEND" = m_type.ToUpper Then
                Do
                    If m_row - ii < 1 Then Exit Sub
                    type = AxSpreadsheet.ActiveSheet.Cells(m_row - ii, gCmdColumn).Value()
                    AxSpreadsheet.ActiveSheet.Rows(m_row - ii).Delete()
                    ii = ii + 1
                    If Nothing = m_type Then
                        m_type = "TMPSTRING"
                    End If
                Loop Until "LINKSTART" = m_type.ToUpper Or ii > 500

                AxSpreadsheet.ActiveSheet.Cells(m_row - ii + 1, 1).Select()
            End If
        Else
            AxSpreadsheet.ActiveSheet.Rows(m_row).Delete()
        End If
        GC.Collect()
    End Sub


    'Shen Jian's part
    Private Sub AxSpreadsheetProgramming_BeforeContextMenu(ByVal sender As System.Object, ByVal e As AxOWC10.ISpreadsheetEventSink_BeforeContextMenuEvent) _
            Handles AxSpreadsheetProgramming.BeforeContextMenu

        Dim oMenu1, oMenu2, oMenu3, oMenu3_1, oMenu4, oMenu5, oMenu6, oMenu7 As Object()
        Dim oMenu As Object()
        If AxSpreadsheetProgramming.ActiveCell.Column = gNeedleColumn Then
            oMenu1 = New Object() {"&Left", "Left"}
            oMenu2 = New Object() {"&Right", "Right"}
            'oMenu3 = New Object() {"____________", ""}
            'oMenu4 = New Object() {"&Undo", "Undo"}
            'oMenu5 = New Object() {"&Insert Row", "Insert"}
            'oMenu6 = New Object() {"&Delete Row", "Delete"}
            'oMenu = New Object() {oMenu1, oMenu2, oMenu3, oMenu4, oMenu5, oMenu6}
            oMenu = New Object() {oMenu1, oMenu2}

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

            'ElseIf AxSpreadsheetProgramming.ActiveCell.Column = gCmdColumn Then
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


    'Shen Jian's part
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
                                    "GETIO", "MOVE", "WAIT", "PURGE", "CLEAN"}
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

    Private Sub AxSpreadsheetProgramming_CommandExecute(ByVal sender As System.Object, ByVal e As AxOWC10.ISpreadsheetEventSink_CommandExecuteEvent) _
              Handles AxSpreadsheetProgramming.CommandExecute
        Dim row As Integer = AxSpreadsheetProgramming.ActiveCell.Row
        Dim colum As Integer = AxSpreadsheetProgramming.ActiveCell.Column
        Dim cmdType As String = AxSpreadsheetProgramming.ActiveSheet.Cells(row, colum).value
        Static m_rowCount As Integer

        UndoData_Logging(1)

        Select Case e.command
            Case "Left", "Right"
                If CheckNeedleChangeFlag(cmdType) Then
                    AxSpreadsheetProgramming.ActiveSheet.Cells(row, colum) = e.command
                Else
                    LabelMessege.Text = "Can't insert needle for this command!"
                    LabelMessege.Update()
                End If

            Case "On"
                If CheckDispenseONOFFFlag(cmdType) Then
                    AxSpreadsheetProgramming.ActiveSheet.Cells(row, colum) = "On"
                Else
                    LabelMessege.Text = "Can't insert ON/OFF for this command!"
                    LabelMessege.Update()
                End If

            Case "Off"
                If CheckDispenseONOFFFlag(cmdType) Then
                    AxSpreadsheetProgramming.ActiveSheet.Cells(row, colum) = "Off"
                Else
                    LabelMessege.Text = "Can't insert ON/OFF for this command!"
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
                AxSpreadsheetProgramming.ActiveSheet.Cells(row, 1).Select()
                AxSpreadsheetProgramming.Selection.Paste()
                AxSpreadsheetProgramming.ActiveSheet.Cells(row, 1).Select()

            Case "MInsert"
                AxSpreadsheetProgramming.ActiveSheet.Cells(row, 1).Select()
                Dim I As Integer = 0
                If m_rowCount > 0 Then
                    For I = 0 To m_rowCount - 1
                        AxSpreadsheetProgramming.Selection.EntireRow.Insert()
                    Next
                End If
                AxSpreadsheetProgramming.Selection.Paste()
                AxSpreadsheetProgramming.ActiveSheet.Cells(row, 1).Select()
                'end here

            Case "Undo"
                AxSpreadsheetProgramming.Undo()
            Case "Insert"
                AxSpreadsheetProgramming.ActiveSheet.Cells(row + 1, 1).Select()
                AxSpreadsheetProgramming.Selection.EntireRow.Insert()
                AxSpreadsheetProgramming.ActiveSheet.Rows(row + 1).clear()
                If (m_row > row) Then
                    m_row = m_row + 1
                End If

            Case "Delete"
                AxSpreadsheetProgramming.ActiveSheet.Cells(row, 1).Select()
                AxSpreadsheetProgramming.Selection.EntireRow.Delete()
                If (m_row > row) Then
                    m_row = m_row - 1
                ElseIf m_row = row Then
                    m_PosUpdate = False
                End If
                If (m_row < 2) Then
                    m_row = 2
                    m_PosUpdate = False
                End If
        End Select


        GC.Collect()
    End Sub




    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Automatically Generate Array or SubSheet Name
    '   SheetName: Sheet name to be generated 
    '   ActSheetName: Current actSheet name       
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Private Sub AxSpreadsheetProgramming_GenerateArraySubSheetName(ByRef SheetName, ByVal ActSheetName, ByVal MyType)

        Dim ArraySheetSerialNumber As Integer = 1

        If "Main" = ActSheetName Then
            Do
                SheetName = "Arr_" + MyType + "_" + CStr(ArraySheetSerialNumber)
                If 0 = m_Execution.m_Pattern.AxSpreadsheetProgramming_CheckSubsheetExist(AxSpreadsheetProgramming, SheetName) Then
                    Exit Do
                End If
                ArraySheetSerialNumber = ArraySheetSerialNumber + 1
            Loop
        Else
            Do
                SheetName = ActSheetName + "." + "Arr_" + MyType + "_" + CStr(ArraySheetSerialNumber)
                If 0 = m_Execution.m_Pattern.AxSpreadsheetProgramming_CheckSubsheetExist(AxSpreadsheetProgramming, SheetName) Then
                    Exit Do
                End If
                ArraySheetSerialNumber = ArraySheetSerialNumber + 1
            Loop
        End If

        GC.Collect()
    End Sub


    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Check the sub or array existance changes, not be used currently
    '   SheetName: Sheet to be checked
    '   
    '   Return: The changes of existance
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Private Function AxSpreadsheetProgramming_QuickCountExistedSubsheet(ByRef AxSpreadsheet As AxOWC10.AxSpreadsheet, _
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
                    strTmp = .Cells(j, gCmdColumn).Value

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

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Total usage of all sub/Array Spreadsheet                                                
    '   AxSpreadsheet:     ActiveX spreadsheet                                                    
    '   RootDelectSheetName: Root sheet name, not in use                                                             
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Private Sub AxSpreadsheetProgramming_FindSubTotalUsage(ByRef AxSpreadsheet As AxOWC10.AxSpreadsheet, _
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
                    strTmp = .Cells(j, gCmdColumn).Value

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
        GC.Collect()
    End Sub


    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Root connection of a sheet, itself is level 1                                              
    '   AxSpreadsheet:     ActiveX spreadsheet                                                    
    '   RootSheetName: Root sheet name                                                              
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Private Sub AxSpreadsheetProgramming_BuildRootConnectedSub(ByRef AxSpreadsheet As AxOWC10.AxSpreadsheet, _
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
                            strTmp = .Cells(j, gCmdColumn).Value

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
        GC.Collect()
    End Sub


    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Self usage of a sub/Array Spreadsheet                                              
    '   AxSpreadsheet:     ActiveX spreadsheet                                                    
    '   RootDelectSheetName: Root sheet name                                                              
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Private Sub AxSpreadsheetProgramming_FindSubSelfUsage(ByRef AxSpreadsheet As AxOWC10.AxSpreadsheet, _
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
                        strTmp = .Cells(j, gCmdColumn).Value

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
        GC.Collect()
    End Sub


    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Self usage of a sub/Array Spreadsheet to include external as as an add on                                               
    '   AxSpreadsheet:     ActiveX spreadsheet                                                    
    '   RootDelectSheetName: Root sheet name                                                              
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Private Sub AxSpreadsheetProgramming_FlatternTotalUsage(ByRef AxSpreadsheet As AxOWC10.AxSpreadsheet, _
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

        GC.Collect()
    End Sub

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Delete one page of Spreadsheet                                                
    '   AxSpreadsheet:     ActiveX spreadsheet                                                    
    '                                                                 
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Private Sub AxSpreadsheetProgramming_DeleteSubSheet(ByRef AxSpreadsheet As AxOWC10.AxSpreadsheet)
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
        GC.Collect()
    End Sub



    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Build Array data based on the 3 rows input                                                
    '   AxSpreadsheet:      ActiveX spreadsheet                                                    
    '   Record1:            The first row for the array   
    '   Record2:            The second row for the array  
    '   Record3:            The third row for the array  
    '   arraydata:          Array parameters 
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Private Sub AxSpreadsheetProgramming_AddSheetRecord( _
          ByRef Record1 As CIDSPattern.PatternRecord, _
          ByRef Record2 As CIDSPattern.PatternRecord, _
          ByRef Record3 As CIDSPattern.PatternRecord, _
          ByVal arraydata As ArrayData)
        Dim i, j, ij As Integer


        Dim dXh, dXv, dYh, dYv, Xij, Yij As Double
        Dim pt1(3), pt2(3), pt3(3), cen1(2), cen2(2), cen3(2), radius As Double
        Dim Record As CIDSPattern.PatternRecord

        If "Arr_Circle" = Record1.pPara.Name Or "Arr_FillCircle" = Record1.pPara.Name Then
            pt1(0) = Record1.pPara.Pos1.X : pt1(1) = Record1.pPara.Pos1.Y : pt1(2) = Record1.pPara.Pos1.Z
            pt2(0) = Record1.pPara.Pos2.X : pt2(1) = Record1.pPara.Pos2.Y : pt2(2) = Record1.pPara.Pos2.Z
            pt3(0) = Record1.pPara.Pos3.X : pt3(1) = Record1.pPara.Pos3.Y : pt3(2) = Record1.pPara.Pos3.Z
            MathFunction.Get3dCircleCentre(pt1, pt2, pt3, cen1, radius)

            pt1(0) = Record2.pPara.Pos1.X : pt1(1) = Record2.pPara.Pos1.Y : pt1(2) = Record2.pPara.Pos1.Z
            pt2(0) = Record2.pPara.Pos2.X : pt2(1) = Record2.pPara.Pos2.Y : pt2(2) = Record2.pPara.Pos2.Z
            pt3(0) = Record2.pPara.Pos3.X : pt3(1) = Record2.pPara.Pos3.Y : pt3(2) = Record2.pPara.Pos3.Z
            MathFunction.Get3dCircleCentre(pt1, pt2, pt3, cen2, radius)

            pt1(0) = Record3.pPara.Pos1.X : pt1(1) = Record3.pPara.Pos1.Y : pt1(2) = Record3.pPara.Pos1.Z
            pt2(0) = Record3.pPara.Pos2.X : pt2(1) = Record3.pPara.Pos2.Y : pt2(2) = Record3.pPara.Pos2.Z
            pt3(0) = Record3.pPara.Pos3.X : pt3(1) = Record3.pPara.Pos3.Y : pt3(2) = Record3.pPara.Pos3.Z
            MathFunction.Get3dCircleCentre(pt1, pt2, pt3, cen3, radius)

            dXh = (cen2(0) - cen1(0)) / (arraydata.ColumNo - 1)
            dYh = (cen2(1) - cen1(1)) / (arraydata.ColumNo - 1)
            dXv = (cen3(0) - cen1(0)) / (arraydata.RowNo - 1)
            dYv = (cen3(1) - cen1(1)) / (arraydata.RowNo - 1)
        Else
            dXh = (Record2.pPara.Pos1.X - Record1.pPara.Pos1.X) / (arraydata.ColumNo - 1)
            dYh = (Record2.pPara.Pos1.Y - Record1.pPara.Pos1.Y) / (arraydata.ColumNo - 1)
            dXv = (Record3.pPara.Pos1.X - Record2.pPara.Pos1.X) / (arraydata.RowNo - 1)
            dYv = (Record3.pPara.Pos1.Y - Record2.pPara.Pos1.Y) / (arraydata.RowNo - 1)
        End If

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

                With AxSpreadsheetProgramming.ActiveSheet

                    .Cells(ij, gCmdColumn).Value = Record1.pPara.Name
                    m_Execution.m_ItemLUT.Cmd2Index((Record1.pPara.Name).ToUpper)


                    If m_Execution.m_ItemLUT.GetFlag(gNeedleColumn) Then
                        .Cells(ij, gNeedleColumn).Value = Record1.pPara.Needle
                    End If
                    If m_Execution.m_ItemLUT.GetFlag(gDispensColumn) Then
                        .Cells(ij, gDispensColumn).Value = Record1.pPara.DispenseFlag
                    End If


                    If m_Execution.m_ItemLUT.GetFlag(gPos1XColumn) Then
                        .Cells(ij, gPos1XColumn).Value = Record1.pPara.Pos1.X + Xij
                        .Cells(ij, gPos1YColumn).Value = Record1.pPara.Pos1.Y + Yij
                        .Cells(ij, gPos1ZColumn).Value = Record1.pPara.Pos1.Z
                    End If
                    If m_Execution.m_ItemLUT.GetFlag(gPos2XColumn) Then
                        .Cells(ij, gPos2XColumn).Value = Record1.pPara.Pos2.X + Xij
                        .Cells(ij, gPos2YColumn).Value = Record1.pPara.Pos2.Y + Yij
                        .Cells(ij, gPos2ZColumn).Value = Record1.pPara.Pos2.Z
                    End If
                    If m_Execution.m_ItemLUT.GetFlag(gPos3XColumn) Then
                        .Cells(ij, gPos3XColumn).Value = Record1.pPara.Pos3.X + Xij
                        .Cells(ij, gPos3YColumn).Value = Record1.pPara.Pos3.Y + Yij
                        .Cells(ij, gPos3ZColumn).Value = Record1.pPara.Pos3.Z
                    End If


                    If m_Execution.m_ItemLUT.GetFlag(gTravelSpeedColumn) Then
                        .Cells(ij, gTravelSpeedColumn).Value = Record1.pPara.TravelSpeed
                    End If
                    If m_Execution.m_ItemLUT.GetFlag(gNeedleGapColumn) Then
                        .Cells(ij, gNeedleGapColumn).Value = Record1.pPara.NeedleGap
                    End If
                    If m_Execution.m_ItemLUT.GetFlag(gDurationColumn) Then
                        .Cells(ij, gDurationColumn).Value = Record1.pPara.DispenseDuration
                    End If
                    If m_Execution.m_ItemLUT.GetFlag(gTravelDelayColumn) Then
                        .Cells(ij, gTravelDelayColumn).Value = Record1.pPara.TravelDelay
                    End If
                    If m_Execution.m_ItemLUT.GetFlag(gRetractDelayColumn) Then
                        .Cells(ij, gRetractDelayColumn).Value = Record1.pPara.RetractDelay
                    End If
                    If m_Execution.m_ItemLUT.GetFlag(gApproachHtColumn) Then
                        .Cells(ij, gApproachHtColumn).Value = Record1.pPara.ApproachDispHeight
                    End If
                    If m_Execution.m_ItemLUT.GetFlag(gRetractSpeedColumn) Then
                        .Cells(ij, gRetractSpeedColumn).Value = Record1.pPara.RetractSpeed
                    End If
                    If m_Execution.m_ItemLUT.GetFlag(gRetractHtColumn) Then
                        .Cells(ij, gRetractHtColumn).Value = Record1.pPara.RetractHeight
                    End If
                    If m_Execution.m_ItemLUT.GetFlag(gClearanceHtColumn) Then
                        .Cells(ij, gClearanceHtColumn).Value = Record1.pPara.ClearanceHeight
                    End If
                    If m_Execution.m_ItemLUT.GetFlag(gDTaiDistColumn) Then
                        .Cells(ij, gDTaiDistColumn).Value = Record1.pPara.DetailingDistance
                    End If
                    If m_Execution.m_ItemLUT.GetFlag(gArcRadiusColumn) Then
                        .Cells(ij, gArcRadiusColumn).Value = Record1.pPara.ArcRadius
                    End If
                    If m_Execution.m_ItemLUT.GetFlag(gPitchColumn) Then
                        .Cells(ij, gPitchColumn).Value = Record1.pPara.Pitch
                    End If
                    If m_Execution.m_ItemLUT.GetFlag(gFillHeightColumn) Then
                        .Cells(ij, gFillHeightColumn).Value = Record1.pPara.FilledHeight
                    End If
                    If m_Execution.m_ItemLUT.GetFlag(gNoOfRunColumn) Then
                        .Cells(ij, gNoOfRunColumn).Value = Record1.pPara.NumberOfRun
                    End If
                    If m_Execution.m_ItemLUT.GetFlag(gSprialColumn) Then
                        .Cells(ij, gSprialColumn).Value = Record1.pPara.SpiralFlag
                    End If
                    If m_Execution.m_ItemLUT.GetFlag(gRtAngleColumn) Then
                        .Cells(ij, gRtAngleColumn).Value = Record1.pPara.RotatingAngle
                    End If
                    If m_Execution.m_ItemLUT.GetFlag(gEdgeClearColumn) Then
                        .Cells(ij, gEdgeClearColumn).Value = Record1.pPara.EdgeClear
                    End If
                    If m_Execution.m_ItemLUT.GetFlag(gIONumberColumn) Then
                        .Cells(ij, gIONumberColumn).Value = Record1.pPara.IONumber
                    End If
                End With
            Next j
        Next i
        GC.Collect()
    End Sub



    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Build SubSub and/or SubArray data based on the sub input                                                                                                   
    '   iSubSheetName:            The full sub sheet name   
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Private Sub AxSpreadsheetProgramming_AddSubandArrayInSub( _
          ByRef iSubSheetName As String)

        Dim strUnit As String
        If True = m_Execution.m_Pattern.m_ErrorChk.WorkingUnitFlag Then
            strUnit = "mm"
        Else
            strUnit = "inch"
        End If

        'Get current sheet name and row number
        Dim TmpSheetName As String = AxSpreadsheetProgramming.ActiveSheet.Name
        Dim TmpRowNumber As Integer = m_row

        If 0 = m_Execution.m_Pattern.AxSpreadsheetProgramming_CheckSubsheetExist( _
            AxSpreadsheetProgramming, m_Execution.m_File.NameOnlyFromFullPath(iSubSheetName)) Then
            m_Execution.m_Pattern.LoadTxtPatternPara(AxSpreadsheetProgramming, _
                iSubSheetName, 1, 0, False, strUnit)
        End If

        'Restore to the previous sheet and rows
        m_row = TmpRowNumber

        AxSpreadsheetProgramming.Worksheets(TmpSheetName).Activate()
        AxSpreadsheetProgramming.ActiveSheet.Cells(m_row + 1, gCmdColumn).select()

    End Sub



    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Get array data from the DialogBox                                                
    '   AxSpreadsheet:      ActiveX spreadsheet                                                    
    '   Record:             One row for the array   
    '   CurrentPt:          The number of position, such x1y1z1=1, x2y2z2=2....  
    '   ArrayDlg:           Array generation DialogBox 
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Private Sub AxSpreadsheetProgramming_GetArrayRecord( _
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

        GC.Collect()
    End Sub


    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Get array data from a row                                                
    '   AxSpreadsheetProgramming: ActiveX spreadsheet, used as global                                                   
    '   Record:             One row data record   
    '   CurrentRow:         The row in the Spreadsheet  
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Private Sub AxSpreadsheetProgramming_GetSheetRecord( _
            ByRef Record As CIDSPattern.PatternRecord, _
            ByVal CurrentRow As Integer)

        Dim strName As String = AxSpreadsheetProgramming.ActiveSheet.Cells(CurrentRow, gCmdColumn).Value()
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
        GC.Collect()
    End Sub

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Is the row empty?
    '   row:  The target row                                                                                
    '
    'Return: True=empty, False=Not empty
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Public Function isEmptyRow(ByVal row As Integer) As Boolean
        Dim i As Integer
        Dim str As String

        For i = 1 To gMaxColumns
            str = AxSpreadsheetProgramming.ActiveSheet.Cells(row, i).Value
            If str <> Nothing Then
                str = str.Trim(" ")
                If str <> "" Then
                    Return False
                End If
            End If
        Next
        Return True
    End Function

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Enable/disable TBrefer all buttons
    '   flag:  Ok=enable, Cancel=disable                                                                                
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Private Sub EnableTBRefers(ByVal flag As Boolean)
        Dim i As Integer
        For i = 0 To gMaxReferButtons
            TBrefer.Buttons(i).Enabled = flag
        Next

        MenuCommandsReference.Enabled = flag
        MenuCommandsFiducial.Enabled = flag
        MenuCommandsHeight.Enabled = flag
        MenuCommandsReject.Enabled = flag
        GC.Collect()
    End Sub

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Enable/disable TBrefer individual button
    '   CmdButton:  the target button    
    '   flag:  Ok=enable, Cancel=disable                                                                                
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Private Sub EnableTBRefers(ByVal CmdButton As Integer, ByVal flag As Boolean)

        TBrefer.Buttons(CmdButton - 1).Enabled = flag

        Select Case (CmdButton)
            Case gFiducialCmdIndex
                MenuCommandsReference.Enabled = flag
            Case gReferenceCmdIndex
                MenuCommandsFiducial.Enabled = flag
            Case gHeightCmdIndex
                MenuCommandsHeight.Enabled = flag
            Case gRejectCmdIndex
                MenuCommandsReject.Enabled = flag
        End Select
        GC.Collect()
    End Sub

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Enable/disable TBElements all buttons
    '   flag:  Ok=enable, Cancel=disable                                                                                
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Private Sub EnableTBElements(ByVal flag As Boolean)
        Dim i As Integer
        For i = 0 To gMaxElementButtons
            TBElements.Buttons(i).Enabled = flag
        Next

        MenuCommandsDot.Enabled = flag
        MenuCommandsLine.Enabled = flag
        MenuCommandsArc.Enabled = flag
        MenuCommandsRectangle.Enabled = flag
        MenuCommandsFillRectangle.Enabled = flag
        MenuCommandsCircle.Enabled = flag
        MenuCommandsFillCircle.Enabled = flag

        MenuCommandsLink.Enabled = flag
        MenuCommandsChipEdge.Enabled = flag
        MenuCommandsPurge.Enabled = flag
        MenuCommandsClean.Enabled = flag
        MenuCommandsQC.Enabled = flag
        MenuCommandsArray.Enabled = flag
        MenuCommandsSubPattern.Enabled = flag

        MenuCommandsMove.Enabled = flag
        MenuCommandsWait.Enabled = flag
        MenuCommandsSetIO.Enabled = flag
        MenuCommandsGetIO.Enabled = flag

        MenuCommandsOffset.Enabled = flag
        MenuCommandsMeasurement.Enabled = flag
        GC.Collect()
    End Sub

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Enable/disable TBElements individual button
    '   CmdButton:  the target button    
    '   flag:  Ok=enable, Cancel=disable                                                                                
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Private Sub EnableTBElements(ByVal CmdButton As Integer, ByVal flag As Boolean)
        TBElements.Buttons(CmdButton - 1).Enabled = flag

        Select Case (CmdButton)
            Case gDotCmdIndex
                MenuCommandsDot.Enabled = flag
            Case gLineCmdIndex
                MenuCommandsLine.Enabled = flag
            Case gArcCmdIndex
                MenuCommandsArc.Enabled = flag
            Case gLinkCmdIndex
                MenuCommandsLink.Enabled = flag

            Case gRectangleCmdIndex
                MenuCommandsRectangle.Enabled = flag
            Case gCircleCmdIndex
                MenuCommandsCircle.Enabled = flag
            Case gFillRectangleCmdIndex
                MenuCommandsFillRectangle.Enabled = flag
            Case gFillCircleCmdIndex
                MenuCommandsFillCircle.Enabled = flag

            Case gChipEdgeCmdIndex
                MenuCommandsChipEdge.Enabled = flag
            Case gMoveCmdIndex
                MenuCommandsMove.Enabled = flag
            Case gWaitCmdIndex
                MenuCommandsWait.Enabled = flag
            Case gPurgeCmdIndex
                MenuCommandsPurge.Enabled = flag
            Case gCleanCmdIndex
                MenuCommandsClean.Enabled = flag
            Case gQCCmdIndex
                MenuCommandsQC.Enabled = flag

            Case gSubPatternCmdIndex
                MenuCommandsSubPattern.Enabled = flag
            Case gArrayCmdIndex
                MenuCommandsArray.Enabled = flag

            Case gSetIOCmdIndex
                MenuCommandsSetIO.Enabled = flag
            Case gGetIOCmdIndex
                MenuCommandsGetIO.Enabled = flag
            Case gSeperatorCmdIndex

            Case gOffsetCmdIndex
                MenuCommandsOffset.Enabled = flag
            Case gMeasurementCmdIndex
                MenuCommandsMeasurement.Enabled = flag
        End Select
        GC.Collect()
    End Sub

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Enable/disable TBNextEdit all buttons
    '   flag:  Ok=enable, Cancel=disable                                                                                
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Private Sub EnableTBNextEdit(ByVal flag As Boolean)
        Dim i As Integer
        For i = 0 To 1
            TBNextEdit.Buttons(i).Enabled = flag
        Next

        MenuCommandsEdit.Enabled = flag
    End Sub

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Enable/disable TBNextEdit specified button
    '   CmdButton:  the target button    
    '   flag:  Ok=enable, Cancel=disable                                                                                
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Private Sub EnableTBNextEdit(ByVal CmdButton As Integer, ByVal flag As Boolean)
        TBNextEdit.Buttons(CmdButton).Enabled = flag

        Select Case (CmdButton)
            Case 1
                MenuCommandsEdit.Enabled = flag
        End Select
    End Sub




    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Click Yes/No button click
    '                                                                                  
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Private Sub ToolBarYesNo_ButtonClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ToolBarButtonClickEventArgs) Handles TBYesNo.ButtonClick
        If e.Button Is TBYesNo.Buttons(0) Then
            ConfirmOrCancel_Implementation(1)   'Add rows
        Else
            ConfirmOrCancel_Implementation(0)   'Cancel or delete rows
        End If
    End Sub

    Dim x1, x2, x3 As Double
    Dim y1, y2, y3 As Double

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Implement Ok/Cancel button
    '   yesnoFlag:  Ok=True, Cancel=False                                                                                
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Private Sub ConfirmOrCancel_Implementation(ByVal yesnoFlag As Integer)
        Dim i As Integer
        Dim PatternLineRecord(3) As CIDSPattern.PatternRecord
        Dim cell1, cell2 As Object

        Dim strUnit As String
        If True = m_Execution.m_Pattern.m_ErrorChk.WorkingUnitFlag Then
            strUnit = "mm"
        Else
            strUnit = "inch"
        End If

        m_row = AxSpreadsheetProgramming.ActiveCell.Row 'Get current row number

        'My testing for cell locking in other way
        'AxSpreadsheetProgramming.ActiveSheet.Cells(5, 5).Locked = True

        Try
            If yesnoFlag = 1 Then               'Programming for adding

                If type = Nothing Then Exit Sub
                Select Case (type.ToUpper)
                    Case "PURGE", "CLEAN"
                        m_PosUpdate = False
                        m_KeyinOK = False
                        AxSpreadsheetProgramming.DisplayWorkbookTabs = True
                        LabelMessege.Text = ""

                        AxSpreadsheetProgramming.ActiveSheet.Cells(m_row + 1, 1).Select()
                        m_row = AxSpreadsheetProgramming.ActiveCell.Row
                        TBYesNo.Buttons(0).Enabled = False
                        TBYesNo.Buttons(1).Enabled = False
                        m_ProgrammingStateFlag = False

                    Case "DOT", "SUBPATTERN", "REFERENCE", "HEIGHT", "QC", "MOVE", "CHIPEDGE", "REJECT", "WAIT"
                        m_PosUpdate = False
                        m_KeyinOK = False
                        AxSpreadsheetProgramming.DisplayWorkbookTabs = True
                        LabelMessege.Text = ""
                        cell1 = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gPos1XColumn)
                        cell2 = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gPos1ZColumn)
                        Dim x1 As Double = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gPos1XColumn).Value
                        Dim y1 As Double = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gPos1YColumn).Value


                        If type.ToUpper = "SUBPATTERN" Then
                            EnableTBRefers(False)
                            EnableTBElements(False)
                        Else
                            EnableTBRefers(True)
                            EnableTBElements(True)
                            TBYesNo.Buttons(0).Enabled = False
                            TBYesNo.Buttons(1).Enabled = False


                            If type.ToUpper = "DOT" Then
                                FrmGD.Dot(x1 + GDRefX, -(y1 - GDRefY), True)

                            ElseIf type.ToUpper = "MOVE" Then
                                'FrmGD.Move(x1, y1, True)

                            ElseIf type.ToUpper = "QC" Then
                                'FrmGD.QCCheck(...)

                            End If

                        End If
                        m_cmdCompleted = 0

                        m_Execution.m_Pattern.AxSpreadsheetProgramming_CellGrey(cell1, cell2, True, AxSpreadsheetProgramming)

                        'Select next row cells

                        If "Height" = type Then
                            Dim posBack(3), posTo(3), posOffset(3) As Double
                            With AxSpreadsheetProgramming.ActiveSheet
                                posBack(0) = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gPos1XColumn).value()
                                posBack(1) = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gPos1YColumn).value()
                                posBack(2) = 0
                            End With

                            posOffset(0) = posOffset(1) = posOffset(2) = 0

                            posTo(0) = posBack(0) * gmminchRatio - gLaserOffX 'IDS.Data.Hardware.HeightSensor.Laser.OffsetPos.X   'xu
                            posTo(1) = posBack(1) * gmminchRatio - gLaserOffY 'IDS.Data.Hardware.HeightSensor.Laser.OffsetPos.Y   'xu
                            posTo(2) = posBack(2) * gmminchRatio

                            Trimotion_MoveTo(posTo, False)

                            '/*      xu
                            ' wait move end 
                            Dim TestValue As Double
                            ' While Not MySSMain.m_Tri.m_TriCtrl.Ain(4, TestValue)
                            'End While
                            Dim rtn As Integer
                            If gLaserLVDTinstalled = 0 Then     'LASER
                                rtn = m_Tri.GetLaserRange(gAIOLaser, TestValue)
                            Else
                                rtn = m_Tri.GetLVDTRange(gAIOLaser, TestValue)
                            End If

                            If rtn = 0 Then                     'No laser readout error
                                posOffset(2) = TestValue - gLaserRefBlkDist 'IDS.Data.Hardware.HeightSensor.Laser.HeightReference
                                AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gPos1ZColumn).value() = CInt(posOffset(2) * 1000.0 / gmminchRatio) / 1000.0
                            Else
                                MessageBox.Show("laser sensor out of range")
                                m_PosNo = 0
                                TBYesNo.Buttons(0).Enabled = False
                                TBYesNo.Buttons(1).Enabled = False
                                EnableTBRefers(True)
                                EnableTBElements(True)
                                EnableTBElements(gOffsetCmdIndex, False)
                                AxSpreadsheetProgramming.ActiveSheet.Rows(m_row).Delete()
                            End If
                            'xu     */
                            posBack(0) = posBack(0) * gmminchRatio
                            posBack(1) = posBack(1) * gmminchRatio
                            posBack(2) = posBack(2) * gmminchRatio
                            Trimotion_MoveTo(posBack, False)
                            AxSpreadsheetProgramming.ActiveSheet.Cells(m_row + 1, gCmdColumn).Select()
                        Else
                            AxSpreadsheetProgramming.ActiveSheet.Cells(m_row + 1, gCmdColumn).Select()
                        End If
                        m_ProgrammingStateFlag = False

                    Case "LINE", "FIDUCIAL"
                        m_PosUpdate = True
                        m_KeyinOK = False
                        AxSpreadsheetProgramming.DisplayWorkbookTabs = False

                        If m_PosNo = 1 Then
                            LabelMessege.Text = "Please confirm the 2nd pt"
                            cell1 = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gPos1XColumn)
                            cell2 = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gPos1ZColumn)

                            x1 = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gPos1XColumn).Value
                            y1 = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gPos1YColumn).Value
                            m_Execution.m_Pattern.AxSpreadsheetProgramming_CellGrey(cell1, cell2, True, AxSpreadsheetProgramming)

                            cell1 = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gPos2XColumn)
                            cell2 = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gPos2ZColumn)
                            AxSpreadsheetProgramming.ActiveSheet.Range(cell1, cell2).Select()


                            AxSpreadsheetProgramming_CheckForWithinLinkRange(False)


                            If m_InLinkRange Then
                                UpdateLinkForPreviousRow()
                            ElseIf type.ToUpper = "FIDUCIAL" Then
                                LabelMessege.Text = "Please confirm the 2nd Fiducial pt"

                                IDS.Devices.Vision.IDSV_Form_FI(1)
                                Dim status As Integer = 0 'status: 1=ok, 2=cancel, 3=end with first fiducial

                                While status = 0
                                    Application.DoEvents()
                                    status = IDS.Devices.Vision.GetFiducialStatus()
                                End While
                                Dim fidname As String
                                If status = 1 Then
                                    fidname = IDS.Devices.Vision.GetFiducialFilename()
                                    Dim Brightness1 As Integer = IDS.Devices.Vision.GetFiducialBrightness
                                    AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gFid1Column) = fidname
                                    AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gBrightnessColumn) = Brightness1
                                ElseIf 2 = status Then  'Cancel command
                                    m_PosUpdate = False
                                    m_PosNo = 0
                                    TBYesNo.Buttons(0).Enabled = False
                                    TBYesNo.Buttons(1).Enabled = False
                                    EnableTBRefers(True)
                                    EnableTBElements(True)
                                    EnableTBElements(gOffsetCmdIndex, False)
                                    AxSpreadsheetProgramming.ActiveSheet.Rows(m_row).Delete()
                                    AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gCmdColumn).Select()

                                ElseIf 3 = status Then
                                    m_PosNo = 0
                                    AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gPos2XColumn).Value() = _
                                        AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gPos1XColumn).Value()
                                    AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gPos2YColumn).Value() = _
                                        AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gPos1YColumn).Value()

                                    cell1 = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gPos2XColumn)
                                    cell2 = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gPos2ZColumn)
                                    m_Execution.m_Pattern.AxSpreadsheetProgramming_CellGrey(cell1, cell2, True, AxSpreadsheetProgramming)

                                    TBYesNo.Buttons(0).Enabled = False
                                    TBYesNo.Buttons(1).Enabled = False
                                    EnableTBRefers(True)
                                    EnableTBElements(True)
                                    EnableTBElements(gOffsetCmdIndex, False)
                                End If

                                'IDS.Devices.Vision.IDSV_Form_FISetSheet(m_row, AxSpreadsheetProgramming)
                            End If

                        ElseIf m_PosNo = 2 Then
                            m_ProgrammingStateFlag = False
                            cell1 = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gPos2XColumn)
                            cell2 = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gPos2ZColumn)
                            x2 = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gPos2XColumn).Value
                            y2 = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gPos2YColumn).Value

                            If type.ToUpper = "LINE" Then
                                FrmGD.Line(x1 - GDRefX, x2 - GDRefX, -(y1 - GDRefY), -(y2 - GDRefY), True)
                            ElseIf type.ToUpper = "FIDUCIAL" Then
                                FrmGD.Fiducial(x1 - GDRefX, x2 - GDRefX, -(y1 - GDRefY), -(y2 - GDRefY), True)
                                IDS.Devices.Vision.IDSV_Form_FI(2)
                                Dim status As Integer = 0

                                While status = 0
                                    Application.DoEvents()
                                    status = IDS.Devices.Vision.GetFiducialStatus()
                                End While
                                Dim fidname As String
                                If status = 1 Then
                                    fidname = IDS.Devices.Vision.GetFiducialFilename()
                                    AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gFid2Column) = fidname
                                    Dim Brightness2 As Integer = IDS.Devices.Vision.GetFiducialBrightness
                                    AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gThresholdColumn) = Brightness2 'For brightness fiducial no.2
                                ElseIf 2 = status Then
                                    AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gPos2XColumn).Value() = _
                                        AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gPos1XColumn).Value()
                                    AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gPos2YColumn).Value() = _
                                        AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gPos1YColumn).Value()

                                    'cell1 = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gPos2XColum)
                                    'cell2 = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gPos2ZColum)
                                    'AxSpreadsheetProgramming_CellGrey(cell1, cell2, True, AxSpreadsheetProgramming)
                                End If
                                'IDS.Devices.Vision.IDSV_Form_FISetSheet(m_row, AxSpreadsheetProgramming)
                            End If

                            UpdateLinkForNextRow()

                            m_PosUpdate = False
                            m_cmdCompleted = 0

                            m_Execution.m_Pattern.AxSpreadsheetProgramming_CellGrey(cell1, cell2, True, AxSpreadsheetProgramming)
                            AxSpreadsheetProgramming.ActiveSheet.Cells(m_row + 1, 1).Select()

                        End If

                        If m_PosNo = 2 Then
                            m_PosUpdate = False
                            AxSpreadsheetProgramming.DisplayWorkbookTabs = True
                            LabelMessege.Text = ""
                            m_PosNo = 1
                            TBYesNo.Buttons(0).Enabled = False
                            TBYesNo.Buttons(1).Enabled = False
                        Else
                            m_PosNo = 2
                        End If

                    Case "ARC", "CIRCLE", "RECTANGLE", "FILLCIRCLE", "FILLRECTANGLE", "ARRAY"
                        m_PosUpdate = True
                        m_KeyinOK = False
                        AxSpreadsheetProgramming.DisplayWorkbookTabs = False

                        If (m_PosNo = 1) Then
                            cell1 = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gPos1XColumn)
                            cell2 = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gPos1ZColumn)
                            x1 = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gPos1XColumn).Value
                            y1 = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gPos1YColumn).Value
                            'SJ
                            'AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gPos1ZColum) = 0.0
                            m_Execution.m_Pattern.AxSpreadsheetProgramming_CellGrey(cell1, cell2, True, AxSpreadsheetProgramming)

                            cell1 = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gPos2XColumn)
                            cell2 = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gPos2ZColumn)
                            AxSpreadsheetProgramming.ActiveSheet.Range(cell1, cell2).Select()

                            m_PosNo = 2

                            UpdateLinkForPreviousRow()

                        ElseIf (m_PosNo = 2) Then
                            cell1 = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gPos2XColumn)
                            cell2 = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gPos2ZColumn)
                            x2 = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gPos2XColumn).Value
                            y2 = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gPos2YColumn).Value
                            'SJ
                            'AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gPos2ZColum) = 0.0
                            m_PosNo = 3
                            m_Execution.m_Pattern.AxSpreadsheetProgramming_CellGrey(cell1, cell2, True, AxSpreadsheetProgramming)
                            cell1 = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gPos3XColumn)
                            cell2 = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gPos3ZColumn)
                            AxSpreadsheetProgramming.ActiveSheet.Range(cell1, cell2).Select()
                        ElseIf (m_PosNo = 3) Then
                            m_ProgrammingStateFlag = False
                            cell1 = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gPos3XColumn)
                            cell2 = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gPos3ZColumn)
                            x3 = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gPos3XColumn).Value
                            y3 = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gPos3YColumn).Value
                            'SJ
                            'AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gPos3ZColum) = 0.0
                            m_PosNo = 1
                            m_cmdCompleted = 0

                            If type.ToUpper = "ARC" Then
                                FrmGD.Arc(x1 - GDRefX, x2 - GDRefX, x3 - GDRefX, -(y1 - GDRefY), -(y2 - GDRefY), -(y3 - GDRefY), True)
                            ElseIf type.ToUpper = "CIRCLE" Then
                                FrmGD.Circle(x1 - GDRefX, x2 - GDRefX, x3 - GDRefX, -(y1 - GDRefY), -(y2 - GDRefY), -(y3 - GDRefY), True)
                            ElseIf type.ToUpper = "RECTANGLE" Then
                                FrmGD.Rectangle(x1 - GDRefX, x2 - GDRefX, x3 - GDRefX, -(y1 - GDRefY), -(y2 - GDRefY), -(y3 - GDRefY), True)
                                'ElseIf type.ToUpper = "RECTANGLE" Then
                                'FrmGD.Rectangular(x1, x2, x3, y1, y2, y3, True)
                                'ElseIf type.ToUpper = "RECTANGLE" Then
                                'FrmGD.Rectangular(x1, x2, x3, y1, y2, y3, True)
                            ElseIf type.ToUpper = "FILLCIRCLE" Then
                                'FrmGD.CircleSpiral(x1, x2, x3, y1, y2, y3, True)
                            ElseIf type.ToUpper = "FILLRECTANGLE" Then
                                'FrmGD.FillRectangular(x1, x2, x3, y1, y2, y3, True)
                            ElseIf type.ToUpper = "CHIPEDGE" Then
                                'FrmGD.ChipEdge( x1, x2, x3, y1, y2, y3, True)
                            End If

                            TBYesNo.Buttons(0).Enabled = False
                            TBYesNo.Buttons(1).Enabled = False

                            UpdateLinkForNextRow()

                            m_Execution.m_Pattern.AxSpreadsheetProgramming_CellGrey(cell1, cell2, True, AxSpreadsheetProgramming)
                            AxSpreadsheetProgramming.ActiveSheet.Cells(m_row + 1, 1).Select()

                        End If

                        If m_PosNo = 1 Then
                            m_PosUpdate = False
                            AxSpreadsheetProgramming.DisplayWorkbookTabs = True
                            LabelMessege.Text = ""
                            If "Array" = type Then
                                EnableTBElements(True)
                            End If
                        End If
                End Select



                If ("SUBPATTERN" = type.ToUpper And m_PosNo = 1 And False = SheetRowSelected) Then

                    Dim DialogPreview As New FormSelectPatternFile

                    DialogPreview.ShowDialog()
                    Dim file As String = DialogPreview.Path


                    'Dim SubPatternDir As String = gPatternFileDir
                    'Dim subfileDlg As New ExplorerStyleViewer(SubPatternDir)

                    If Nothing = file Then
                        'Clear rows based on the rejected input.  Array of Sub needs to celar more rows.
                        With AxSpreadsheetProgramming.ActiveSheet

                            .Rows(m_row).Delete()
                            .Cells(m_row, gCmdColumn).Select()

                        End With
                        'Enable cmd button if it is rejected
                        EnableTBRefers(True)
                        EnableTBElements(True)

                        type = "Dot"
                        TBYesNo.Buttons(0).Enabled = False
                        TBYesNo.Buttons(1).Enabled = False

                    Else
                        AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gSubnameColumn) = file

                        'Get current sheet name and row number
                        Dim TmpSheetName As String = AxSpreadsheetProgramming.ActiveSheet.Name
                        Dim TmpRowNumber As Integer = m_row

                        If 0 = m_Execution.m_Pattern.AxSpreadsheetProgramming_CheckSubsheetExist( _
                            AxSpreadsheetProgramming, m_Execution.m_File.NameOnlyFromFullPath(file)) Then
                            m_Execution.m_Pattern.LoadTxtPatternPara(AxSpreadsheetProgramming, _
                                file, 1, 0, False, strUnit)
                        End If

                        'Restore to the previous sheet and rows
                        m_row = TmpRowNumber

                        AxSpreadsheetProgramming.Worksheets(TmpSheetName).Activate()
                        AxSpreadsheetProgramming.ActiveSheet.Cells(m_row + 1, gCmdColumn).select()
                        EnableTBRefers(True)
                        EnableTBElements(True)
                        TBYesNo.Buttons(0).Enabled = False
                        TBYesNo.Buttons(1).Enabled = False

                    End If
                End If
            ElseIf yesnoFlag = 0 Then   'Delete rows

                Dim sel As Microsoft.Office.Interop.OWC.Range = AxSpreadsheetProgramming.Selection
                Dim m_rowCount As Integer = sel.Rows.Count()

                If m_rowCount = 1 Then
                    type = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gCmdColumn).Value
                    UndoData_Logging(0)

                    If "" = type Then
                        AxSpreadsheetProgramming.ActiveSheet.Rows(m_row).Delete()
                    ElseIf m_EditStateFlag = True Then 'edit 
                        If m_Execution.m_Pattern.m_ErrorChk.CheckRequiredPt3XY(type) Then
                            cell1 = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gPos1XColumn)
                            cell2 = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gPos3ZColumn)
                            m_Execution.m_Pattern.AxSpreadsheetProgramming_CellGrey(cell1, cell2, True, AxSpreadsheetProgramming)

                            UpdateLinkForNextRow()
                            UpdateLinkForPreviousRow()

                        ElseIf m_Execution.m_Pattern.m_ErrorChk.CheckRequiredPt2XY(type) Then
                            cell1 = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gPos1XColumn)
                            cell2 = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gPos2ZColumn)
                            m_Execution.m_Pattern.AxSpreadsheetProgramming_CellGrey(cell1, cell2, True, AxSpreadsheetProgramming)

                            UpdateLinkForNextRow()
                            UpdateLinkForPreviousRow()

                    ElseIf m_Execution.m_Pattern.m_ErrorChk.CheckRequiredPt1XY(type) Then
                        cell1 = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gPos1XColumn)
                        cell2 = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gPos1ZColumn)
                        m_Execution.m_Pattern.AxSpreadsheetProgramming_CellGrey(cell1, cell2, True, AxSpreadsheetProgramming)

                    End If

                    'lim
                    Dim CmdName As String
                    m_row = AxSpreadsheetProgramming.ActiveCell.Row
                    CmdName = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gCmdColumn).value
                    If "FIDUCIAL" = CmdName.ToUpper Then
                        'For Lim's part
                        If m_PosNo = 1 Then
                            IDS.Devices.Vision.IDSV_Form_FI_Edit(1, gPatternFileName + "1.mmo")
                            Dim status As Integer = 0 'status: 1=ok, 2=cancel, 3=end with first fiducial

                            While status = 0
                                Application.DoEvents()
                                status = IDS.Devices.Vision.GetFiducialStatus()
                            End While
                            Dim fidname As String
                            If status = 1 Then
                                fidname = IDS.Devices.Vision.GetFiducialFilename()
                                Dim Brightness1 As Integer = IDS.Devices.Vision.GetFiducialBrightness
                                AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gFid1Column) = fidname
                                AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gBrightnessColumn) = Brightness1
                            ElseIf 2 = status Then  'Cancel command
                                m_PosUpdate = False
                                m_PosNo = 0
                                TBYesNo.Buttons(0).Enabled = False
                                TBYesNo.Buttons(1).Enabled = False
                                EnableTBRefers(True)
                                EnableTBElements(True)
                                EnableTBElements(gOffsetCmdIndex, False)
                                'AxSpreadsheetProgramming.ActiveSheet.Rows(m_row).Delete()
                                'AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gCmdColumn).Select()

                            ElseIf 3 = status Then
                                fidname = IDS.Devices.Vision.GetFiducialFilename()
                                AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gFid1Column) = fidname
                                'm_PosNo = 0
                                'AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gPos2XColumn).Value() = _
                                '    AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gPos1XColumn).Value()
                                'AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gPos2YColumn).Value() = _
                                '    AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gPos1YColumn).Value()

                                'cell1 = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gPos2XColumn)
                                'cell2 = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gPos2ZColumn)
                                'm_Execution.m_Pattern.AxSpreadsheetProgramming_CellGrey(cell1, cell2, True, AxSpreadsheetProgramming)

                                'TBYesNo.Buttons(0).Enabled = False
                                'TBYesNo.Buttons(1).Enabled = False
                                'EnableTBRefers(True)
                                'EnableTBElements(True)
                                'EnableTBElements(gOffsetCmdIndex, False)
                            End If
                        ElseIf m_PosNo = 2 Then
                            IDS.Devices.Vision.IDSV_Form_FI_Edit(2, gPatternFileName + "2.mmo")
                            Dim status As Integer = 0

                            While status = 0
                                Application.DoEvents()
                                status = IDS.Devices.Vision.GetFiducialStatus()
                            End While
                            Dim fidname As String
                            If status = 1 Then
                                fidname = IDS.Devices.Vision.GetFiducialFilename()
                                AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gFid2Column) = fidname
                                Dim Brightness2 As Integer = IDS.Devices.Vision.GetFiducialBrightness
                                AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gThresholdColumn) = Brightness2 'For brightness fiducial no.2
                            ElseIf 2 = status Then
                                '*Do nothing for editing... (unlike programming, it hold the original file)

                                'AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gPos2XColumn).Value() = _
                                'AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gPos1XColumn).Value()
                                'AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gPos2YColumn).Value() = _
                                'AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gPos1YColumn).Value()

                                'cell1 = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gPos2XColum)
                                'cell2 = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gPos2ZColum)
                                'AxSpreadsheetProgramming_CellGrey(cell1, cell2, True, AxSpreadsheetProgramming)
                            End If
                        End If
                    ElseIf "REJECT" = CmdName.ToUpper Then 'For Lim's part
                        Dim vPara As DLL_Export_Device_Vision.RejectPoint.RMParam
                        m_row = AxSpreadsheetProgramming.ActiveCell.Row
                        vpara._AcceptRatio = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gAcceptRatioCoulumn).value
                        vpara._Binarized = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gBinarizedColumn).value
                        vpara._BlackWithoutRM = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gBlackWithoutRMCoulumn).value
                        vpara._BlackWithRM = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gBlackWithRMCoulumn).value
                        vpara._Brightness = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gBrightnessColumn).value
                        vpara._MRegionX = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gMRegionXColumn).value
                        vpara._MRegionY = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gMRegionYColumn).value
                        vpara._MROIx = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gMROIxColumn).value
                        vpara._MROIy = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gMROIyColumn).value
                        vpara._WhiteWithoutRM = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gWhiteWithoutRMCoulumn).value
                        vpara._WhiteWithRM = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gWhiteWithRMCoulumn).value
                        vpara._WoRM = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gWoRMCoulumn).value
                        vpara._WRM = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gWRMCoulumn).value
                        IDS.Devices.Vision.IDSV_Form_RM_Edit(vpara)

                        Dim status As Integer = 0
                        Dim x, y As Double

                        While status = 0
                            Application.DoEvents()
                            status = IDS.Devices.Vision.GetRMStatus
                        End While
                        If status = 2 Then          'Cancel
                            m_PosUpdate = False
                            LabelMessege.Text = ""
                            AxSpreadsheetProgramming.DisplayWorkbookTabs = True
                            'AxSpreadsheetProgramming.ActiveSheet.Rows(m_row).delete()
                            AxSpreadsheetProgramming.Update()
                            AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gCmdColumn).Select()
                        ElseIf status = 1 Then      'Ok
                            Dim bb As Boolean
                            IDS.Devices.Vision.GetRMParameters(vpara)
                            IDS.Devices.Vision.SetRMReset()
                            m_PosUpdate = False
                            With AxSpreadsheetProgramming.ActiveSheet
                                .Cells(m_row, gAcceptRatioCoulumn) = vpara._AcceptRatio
                                .Cells(m_row, gBinarizedColumn) = vpara._Binarized
                                .Cells(m_row, gBlackWithoutRMCoulumn) = vpara._BlackWithoutRM
                                .Cells(m_row, gBlackWithRMCoulumn) = vpara._BlackWithRM
                                .Cells(m_row, gBrightnessColumn) = vpara._Brightness
                                .Cells(m_row, gMRegionXColumn) = vpara._MRegionX
                                .Cells(m_row, gMRegionYColumn) = vpara._MRegionY
                                .Cells(m_row, gMROIxColumn) = vpara._MROIx
                                .Cells(m_row, gMROIyColumn) = vpara._MROIy
                                .Cells(m_row, gWhiteWithoutRMCoulumn) = vpara._WhiteWithoutRM
                                .Cells(m_row, gWhiteWithRMCoulumn) = vpara._WhiteWithRM
                                .Cells(m_row, gWoRMCoulumn) = vpara._WoRM
                                .Cells(m_row, gWRMCoulumn) = vpara._WRM

                                cell1 = .Cells(m_row, gPos1XColumn)
                                cell2 = .Cells(m_row, gPos1ZColumn)
                            End With

                            m_Execution.m_Pattern.AxSpreadsheetProgramming_CellGrey(cell1, cell2, True, AxSpreadsheetProgramming)

                            LabelMessege.Text = ""
                            AxSpreadsheetProgramming.DisplayWorkbookTabs = True
                            AxSpreadsheetProgramming.ActiveSheet.Cells(m_row + 1, gCmdColumn).Select()
                        End If
                        EnableTBRefers(True)
                        EnableTBElements(True)
                        EnableTBElements(gOffsetCmdIndex, False)
                        TBYesNo.Buttons(0).Enabled = False
                        TBYesNo.Buttons(1).Enabled = True

                    ElseIf "CHIPEDGE" = CmdName.ToUpper Then 'lim
                        Dim vPara As DLL_Export_Device_Vision.ChipEdgePoints.ChipEdgeParam
                        m_row = AxSpreadsheetProgramming.ActiveCell.Row
                        vPara._Brightness = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gBrightnessColumn).value
                        vPara._EdgeClearance = (AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gEdgeClearColumn).value) * gmminchRatio
                        vPara._CheckBox_ChipRec_Enable = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gCheckBoxColumn).value
                        vPara._Cw_CCw = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gCwCCwColumn).value
                        vPara._DispenseModel = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gDispModelColumn).value
                        vPara._Inside_out = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gInOutColumn).value
                        vPara._MainEdge = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gMainEdgeColumn).value
                        vPara._PointX1 = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gPointX1Column).value
                        vPara._PointX2 = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gPointX2Column).value
                        vPara._PointX3 = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gPointX3Column).value
                        vPara._PointX4 = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gPointX4Column).value
                        vPara._PointX5 = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gPointX5Column).value
                        vPara._PointY1 = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gPointY1Column).value
                        vPara._PointY2 = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gPointY2Column).value
                        vPara._PointY3 = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gPointY3Column).value
                        vPara._PointY4 = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gPointY4Column).value
                        vPara._PointY5 = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gPointY5Column).value
                        vPara._Pos = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gPosColumn).value
                        vPara._PosX = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gPosXColumn).value
                        vPara._PosY = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gPosYColumn).value
                        vPara._ROI = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gROIColumn).value
                        vPara._Rot = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gRotColumn).value
                        vPara._Size = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gSizeColumn).value
                        vPara._SizeX = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gSizeXColumn).value
                        vPara._SizeY = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gSizeYColumn).value
                        vPara._Threshold = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gThresholdColumn).value
                        vPara._Vertical = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gVerticalColumn).value
                        IDS.Devices.Vision.IDSV_Form_CE_Edit(vPara)

                        Dim status As Integer = 0
                        Dim x, y As Double
                        Dim pos(3) As Double

                        Do
                                While Not (IDS.Devices.Vision.RobotMotionOffset(x, y) = True Or IDS.Devices.Vision.GetChipEdgeStatus = 2 Or IDS.Devices.Vision.GetChipEdgeStatus = 1)
                                    Application.DoEvents()
                                End While
                                If IDS.Devices.Vision.GetChipEdgeStatus = 2 Then
                                    LabelMessege.Text = ""
                                    TBElements.Enabled = True
                                    TBrefer.Enabled = True
                                    m_PosUpdate = False
                                    'AxSpreadsheetProgramming.DisplayWorkbookTabs = True
                                    'AxSpreadsheetProgramming.ActiveSheet.Rows(m_row).clear()
                                    'AxSpreadsheetProgramming.Update()
                                    'Return
                                End If

                                'moverobot
                                pos(0) = x
                                pos(1) = -y
                                pos(2) = 0
                                Dim spstr As String = gSystemSpeed.ToString
                                Dim str As String
                                str = "SPEED AXIS(0) =" + spstr
                                m_Tri.m_TriCtrl.Execute(str)
                                str = "SPEED AXIS(1) =" + spstr
                                m_Tri.m_TriCtrl.Execute(str)
                                str = "SPEED AXIS(2) =" + spstr
                                m_Tri.m_TriCtrl.Execute(str)
                                'Dim spstr As String = gSystemSpeed.ToString
                                'spstr = "SPEED AXIS(0) =" + spstr
                                'm_Tri.m_TriCtrl.Execute(spstr)
                                m_Tri.m_TriCtrl.MoveRel(2, pos, 0)
                                m_Tri.WaitForEndOfMove()
                                '??????????????????
                                If status = 3 Then
                                    status = 0 'reset status after 5 points being reset
                                End If
                                While status = 0
                                    Application.DoEvents()
                                    status = IDS.Devices.Vision.GetChipEdgeStatus()
                                End While
                            Loop While status = 3 'status 3= reset 5 points

                        If status = 2 Then
                            LabelMessege.Text = ""
                            TBElements.Enabled = True
                            TBrefer.Enabled = True
                            m_PosUpdate = False
                                'AxSpreadsheetProgramming.DisplayWorkbookTabs = True
                                'AxSpreadsheetProgramming.ActiveSheet.Rows(m_row).clear()
                                'AxSpreadsheetProgramming.Update()
                                'Return
                        End If

                        If status = 1 Then
                            Dim bb As Boolean
                            'Dim vpara As DLL_Export_Device_Vision.ChipEdgePoints.ChipEdgeParam
                            IDS.Devices.Vision.GetChipEdgeParameters(vPara)
                            IDS.Devices.Vision.SetCEReset()
                            With AxSpreadsheetProgramming.ActiveSheet
                                .Cells(m_row, gEdgeClearColumn) = vPara._EdgeClearance / gmminchRatio
                                .Cells(m_row, gBrightnessColumn) = vPara._Brightness
                                .Cells(m_row, gCheckBoxColumn) = vPara._CheckBox_ChipRec_Enable
                                .Cells(m_row, gCwCCwColumn) = vPara._Cw_CCw
                                .Cells(m_row, gDispModelColumn) = vPara._DispenseModel
                                .Cells(m_row, gInOutColumn) = vPara._Inside_out
                                .Cells(m_row, gMainEdgeColumn) = vPara._MainEdge
                                .Cells(m_row, gPointX1Column) = vPara._PointX1
                                .Cells(m_row, gPointX2Column) = vPara._PointX2
                                .Cells(m_row, gPointX3Column) = vPara._PointX3
                                .Cells(m_row, gPointX4Column) = vPara._PointX4
                                .Cells(m_row, gPointX5Column) = vPara._PointX5
                                .Cells(m_row, gPointY1Column) = vPara._PointY1
                                .Cells(m_row, gPointY2Column) = vPara._PointY2
                                .Cells(m_row, gPointY3Column) = vPara._PointY3
                                .Cells(m_row, gPointY4Column) = vPara._PointY4
                                .Cells(m_row, gPointY5Column) = vPara._PointY5
                                .Cells(m_row, gPosColumn) = vPara._Pos
                                .Cells(m_row, gPosXColumn) = vPara._PosX
                                .Cells(m_row, gPosYColumn) = vPara._PosY
                                .Cells(m_row, gROIColumn) = vPara._ROI
                                .Cells(m_row, gRotColumn) = vPara._Rot
                                .Cells(m_row, gSizeColumn) = vPara._Size
                                .Cells(m_row, gSizeXColumn) = vPara._SizeX
                                .Cells(m_row, gSizeYColumn) = vPara._SizeY
                                .Cells(m_row, gThresholdColumn) = vPara._Threshold
                                .Cells(m_row, gVerticalColumn) = vPara._Vertical

                                cell1 = .Cells(m_row, gPos1XColumn)
                                cell2 = .Cells(m_row, gPos1ZColumn)


                            End With


                            m_Execution.m_Pattern.AxSpreadsheetProgramming_CellGrey(cell1, cell2, True, AxSpreadsheetProgramming)
                            TBElements.Enabled = True
                            TBrefer.Enabled = True
                            m_PosUpdate = False
                            LabelMessege.Text = ""
                            AxSpreadsheetProgramming.DisplayWorkbookTabs = True
                                'm_row = m_row + 1
                                AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, 1).Select()

                            End If
                            EnableTBRefers(True)
                            EnableTBElements(True)
                            EnableTBElements(gOffsetCmdIndex, False)
                            TBYesNo.Buttons(0).Enabled = False
                            TBYesNo.Buttons(1).Enabled = True
                            m_ProgrammingStateFlag = False
                            '================================================================================================================
                          

                    ElseIf "QC" = CmdName.ToUpper Then 'lim
                        Dim vPara As DLL_Export_Device_Vision.QC.QCParam
                        m_row = AxSpreadsheetProgramming.ActiveCell.Row
                        vPara._type = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gTypeColumn).value
                        vPara._Brightness = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gBrightnessColumn).value
                        vPara._BlackDot = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gBlackDotColumn).value
                        vPara._Binarized = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gBinarizedColumn).value
                        vPara._MaxArea = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gMaxAreaColumn).value
                        vPara._MinArea = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gMinAreaColumn).value
                        vPara._Close = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gCloseColumn).value
                        vPara._Open = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gOpenColumn).value
                        vPara._Roughness = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gRoughnessColumn).value
                        vPara._Compactness = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gCompactnessColumn).value
                        vPara._MRegionX = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gMRegionXColumn).value
                        vPara._MRegionY = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gMRegionYColumn).value
                        vPara._MROIx = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gMROIxColumn).value
                        vPara._MROIy = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gMROIyColumn).value
                        vPara._SizeX = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gSizeXColumn).value
                        vPara._SizeY = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gSizeYColumn).value
                        vPara._PosX = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gPosXColumn).value
                        vPara._PosY = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gPosYColumn).value
                        vPara._Size = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gSizeColumn).value
                        vPara._Pos = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gPosColumn).value
                        vPara._Rot = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gRotColumn).value
                        vPara._Inside_out = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gInOutColumn).value
                        vPara._Threshold = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gThresholdColumn).value
                        vPara._ROI = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gROIColumn).value
                        vPara._Vertical = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gVerticalColumn).value
                        vPara._PointX1 = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gPointX1Column).value
                        vPara._PointY1 = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gPointY1Column).value
                        vPara._PointX2 = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gPointX2Column).value
                        vPara._PointY2 = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gPointY2Column).value
                        vPara._PointX3 = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gPointX3Column).value
                        vPara._PointY3 = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gPointY3Column).value
                        vPara._PointX4 = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gPointX4Column).value
                        vPara._PointY4 = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gPointY4Column).value
                        vPara._PointX5 = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gPointX5Column).value
                        vPara._PointY5 = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gPointY5Column).value
                        vPara._MQC_OffX = 0 'programming GUI
                        vPara._MQC_OffY = 0 'programming GUI
                        IDS.Devices.Vision.IDSV_Form_QC_Edit(vPara)

                        'EnableTBRefers(False)
                        'EnableTBElements(False)
                        'EnableTBElements(gOffsetCmdIndex, False)
                        'TBYesNo.Buttons(0).Enabled = False
                        'TBYesNo.Buttons(1).Enabled = False
                        'IDS.Devices.Vision.IDSV_Form_QC()
                        Dim status As Integer = 0
                        Dim x, y As Double
                        Dim _5dot As Boolean = False
                        Dim pos(3) As Double
                        While status = 0 Or status = 3
                            Do
                                Application.DoEvents()
                                If IDS.Devices.Vision.GetQCSelection = 1 Then 'dot

                                Else 'rect
                                    If status = 3 Then
                                        status = 0 'reset status after 5 points being reset
                                    End If
                                    While IDS.Devices.Vision.GetQCStatus <> 2 And IDS.Devices.Vision.GetQCSelection = 2
                                        If IDS.Devices.Vision.RobotMotionOffset(x, y) Then
                                            Application.DoEvents()
                                            _5dot = True
                                            Exit While
                                        Else
                                            Application.DoEvents()
                                        End If

                                    End While
                                    If _5dot Then
                                        'moverobot
                                        pos(0) = x
                                        pos(1) = -y
                                        pos(2) = 0
                                        Dim spstr As String = gSystemSpeed.ToString
                                        Dim str As String
                                        str = "SPEED AXIS(0) =" + spstr
                                        m_Tri.m_TriCtrl.Execute(str)
                                        str = "SPEED AXIS(1) =" + spstr
                                        m_Tri.m_TriCtrl.Execute(str)
                                        str = "SPEED AXIS(2) =" + spstr
                                        m_Tri.m_TriCtrl.Execute(str)
                                        'Dim spstr As String = gSystemSpeed.ToString
                                        'spstr = "SPEED AXIS(0) =" + spstr
                                        'm_Tri.m_TriCtrl.Execute(spstr)
                                        m_Tri.m_TriCtrl.MoveRel(2, pos, 0)
                                        m_Tri.WaitForEndOfMove()
                                        _5dot = False
                                    End If
                                End If
                                status = IDS.Devices.Vision.GetQCStatus()
                            Loop While status = 3
                        End While

                        If 2 = status Then  'Cancel
                            LabelMessege.Text = ""
                            TBElements.Enabled = True
                            TBrefer.Enabled = True
                            m_PosUpdate = False
                            AxSpreadsheetProgramming.DisplayWorkbookTabs = True
                            AxSpreadsheetProgramming.ActiveSheet.Rows(m_row).delete()
                            AxSpreadsheetProgramming.Update()
                        End If

                        If 1 = status Then                         'Ok
                            Dim bb As Boolean
                            IDS.Devices.Vision.GetQCParameters(vpara)
                            IDS.Devices.Vision.SetQCReset()

                            With AxSpreadsheetProgramming.ActiveSheet
                                .Cells(m_row, gBrightnessColumn) = vpara._Brightness
                                .Cells(m_row, gBinarizedColumn) = vpara._Binarized
                                .Cells(m_row, gBlackDotColumn) = vpara._BlackDot
                                .Cells(m_row, gOpenColumn) = vpara._Open
                                .Cells(m_row, gInOutColumn) = vpara._Inside_out
                                .Cells(m_row, gCloseColumn) = vpara._Close
                                .Cells(m_row, gPointX1Column) = vpara._PointX1
                                .Cells(m_row, gPointX2Column) = vpara._PointX2
                                .Cells(m_row, gPointX3Column) = vpara._PointX3
                                .Cells(m_row, gPointX4Column) = vpara._PointX4
                                .Cells(m_row, gPointX5Column) = vpara._PointX5
                                .Cells(m_row, gPointY1Column) = vpara._PointY1
                                .Cells(m_row, gPointY2Column) = vpara._PointY2
                                .Cells(m_row, gPointY3Column) = vpara._PointY3
                                .Cells(m_row, gPointY4Column) = vpara._PointY4
                                .Cells(m_row, gPointY5Column) = vpara._PointY5
                                .Cells(m_row, gPosColumn) = vpara._Pos
                                .Cells(m_row, gPosXColumn) = vpara._PosX
                                .Cells(m_row, gPosYColumn) = vpara._PosY
                                .Cells(m_row, gROIColumn) = vpara._ROI
                                .Cells(m_row, gRotColumn) = vpara._Rot
                                .Cells(m_row, gSizeColumn) = vpara._Size
                                .Cells(m_row, gSizeXColumn) = vpara._SizeX
                                .Cells(m_row, gSizeYColumn) = vpara._SizeY
                                .Cells(m_row, gThresholdColumn) = vpara._Threshold
                                .Cells(m_row, gVerticalColumn) = vpara._Vertical
                                .Cells(m_row, gCompactnessColumn) = vpara._Compactness
                                .Cells(m_row, gMaxAreaColumn) = vpara._MaxArea
                                .Cells(m_row, gMinAreaColumn) = vpara._MinArea
                                .Cells(m_row, gMQC_OffXColumn) = vpara._MQC_OffX
                                .Cells(m_row, gMQC_OffYColumn) = vpara._MQC_OffY
                                .Cells(m_row, gMRegionXColumn) = vpara._MRegionX
                                .Cells(m_row, gMRegionYColumn) = vpara._MRegionY
                                .Cells(m_row, gMROIxColumn) = vpara._MROIx
                                .Cells(m_row, gMROIyColumn) = vpara._MROIy
                                .Cells(m_row, gRoughnessColumn) = vpara._Roughness
                                .Cells(m_row, gTypeColumn) = vpara._type
                                '.Cells(m_row, gBrightnessColumn)=vpara.

                                cell1 = .Cells(m_row, gPos1XColumn)
                                cell2 = .Cells(m_row, gPos1ZColumn)

                            End With
                            IDS.Devices.Vision.SetQCReset()
                            IDS.Devices.Vision.FrmVision.DisplayIndicator()

                            m_Execution.m_Pattern.AxSpreadsheetProgramming_CellGrey(cell1, cell2, True, AxSpreadsheetProgramming)
                            TBElements.Enabled = True
                            TBrefer.Enabled = True
                            m_PosUpdate = False
                            LabelMessege.Text = ""
                            AxSpreadsheetProgramming.DisplayWorkbookTabs = True
                            m_row = m_row + 1
                            AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, 1).Select()
                        End If
                        EnableTBRefers(True)
                        EnableTBElements(True)
                        EnableTBElements(gOffsetCmdIndex, False)
                        TBYesNo.Buttons(0).Enabled = False
                        TBYesNo.Buttons(1).Enabled = True
                        'xiong end============                        
                        '=================================================================================
                    End If
                    MenuEditCopy.Enabled = True
                    If "" <> CopiedSheetName Then
                        MenuEditPaste.Enabled = True
                    End If
                    MenuEditCut.Enabled = True
                    MenuEditUndo.Enabled = True
                    MenuEditRedo.Enabled = False
                ElseIf m_ProgrammingStateFlag = True Then
                        AxSpreadsheetProgramming.ActiveSheet.Rows(m_row).Delete()
                        AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, 1).Select()
                        m_ProgrammingStateFlag = False
                    Else
                        AxSpreadsheetProgramming_DeleteMultiRow(AxSpreadsheetProgramming)
                        m_ProgrammingStateFlag = False
                    End If
                Else
                    UndoData_Logging(0)
                    AxSpreadsheetProgramming_DeleteMultiRow(AxSpreadsheetProgramming)

                End If
                m_PosUpdate = False
                m_KeyinOK = False
                m_EditStateFlag = False
                LabelMessege.Text = ""
                AxSpreadsheetProgramming.Update()
                AxSpreadsheetProgramming_CheckForWithinLinkRange(True)

                AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, gCmdColumn).select()
                TBYesNo.Buttons(0).Enabled = False
                TBYesNo.Buttons(1).Enabled = True
                EnableTBNextEdit(False)
                EnableTBElements(gOffsetCmdIndex, False)
                AxSpreadsheetProgramming.DisplayWorkbookTabs = True

            End If

        Catch ex As SystemException
            MessageBox.Show(ex.ToString)
        End Try
        GC.Collect()
    End Sub


    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Edit or cancel operation
    '   editcancelFlag: 0 or 1; 
    '               0 Start edit
    '               1 Edit finished   
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Private Sub EditOrCancel(ByVal editcancelFlag As Integer)
        Try
            Dim ran As String = "C" + m_row.ToString + ":E" + m_row.ToString
            Dim backup(3) As Object
            backup(0) = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, 3)
            backup(1) = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, 4)
            backup(2) = AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, 5)
            If editcancelFlag = 1 Then
                'm_row = AxSpreadsheetProgramming.ActiveCell.Row
                TBElements.Enabled = False
                TBrefer.Enabled = False
                m_PosUpdate = True
                m_KeyinOK = True
                LabelMessege.Text = "Move robot to teach point or keyin ponit"
                ToolBarSwitch("YesNo")
                AxSpreadsheetProgramming.ActiveSheet.Range(ran).Clear()
                AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, 3) = backup(0)
                AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, 4) = backup(1)
                AxSpreadsheetProgramming.ActiveSheet.Cells(m_row, 5) = backup(2)

            ElseIf editcancelFlag = 0 Then
                m_PosUpdate = False
                m_KeyinOK = False
                AxSpreadsheetProgramming.ActiveSheet.Rows(m_row).clear()

                LabelMessege.Text = ""
                AxSpreadsheetProgramming.Update()
            End If

        Catch ex As SystemException
            MessageBox.Show(ex.ToString)
        End Try
    End Sub


    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Switch toobar operation
    '   toggle_str: Switch between YesNo or Edit 
    '          
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Public Sub ToolBarSwitch(ByVal toggle_str As String)
        If toggle_str = "YesNo" Then
            LabelMessege.Text = "Move robot to teach point "
            m_KeyinOK = False

        Else
            LabelMessege.Text = "Move robot or keyin"
            m_KeyinOK = True
        End If

        m_PosUpdate = True

        ''''temp comment
        '        ToolBarYesNo.Enabled = True
        '       ToolBarYesNo.Buttons(0).Enabled = True
        '      ToolBarYesNo.Buttons(1).Enabled = True
        '     ToolBarYesNo.Visible = True
        '    ToolBarYesNo.Buttons(0).Visible = True
        '   ToolBarYesNo.Buttons(1).Visible = True
    End Sub


#End Region

#Region "Others"
    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Dim vParam As DLL_Export_Device_Vision.QC.QCParam
        vParam._Binarized = 128

        vParam._BlackDot = True
        vParam._Brightness = 128
        vParam._Close = 0
        vParam._Compactness = 5
        vParam._Inside_out = True
        vParam._MaxArea = 8
        vParam._MinArea = 0.3
        vParam._MQC_OffX = 0
        vParam._MQC_OffY = 0
        vParam._MRegionX = 384
        vParam._MRegionY = 288
        vParam._MROIx = 100
        vParam._MROIy = 100
        vParam._Open = 0
        vParam._PointX1 = 0
        vParam._PointX2 = 0
        vParam._PointX3 = 0
        vParam._PointX4 = 0
        vParam._PointX5 = 0
        vParam._PointY1 = 0
        vParam._PointY2 = 0
        vParam._PointY3 = 0
        vParam._PointY4 = 0
        vParam._PointY5 = 0
        vParam._Pos = 0
        vParam._PosX = 0
        vParam._PosY = 0
        vParam._ROI = 2
        vParam._Rot = 5
        vParam._Roughness = 5
        vParam._Size = 0
        vParam._SizeX = 0
        vParam._SizeY = 0
        vParam._Threshold = 15
        vParam._type = 1

        vParam._Vertical = True
        Dim result As Boolean = IDS.Devices.Vision.IDSV_QC(vParam)

    End Sub

    Private Sub PanelVision_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles PanelVision.Paint

    End Sub


    Private Sub OptimizePath_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OptimizePath.Click

        m_Execution.m_Command.SetOptimizFlag(1)
        Dim rtn As Integer = m_Execution.m_Command.ReadPattern(AxSpreadsheetProgramming)
        Dim optm As ArrayList
        Dim row As Integer

        If rtn = 0 Then
            optm = m_Execution.m_Command.PattenList
            row = optm.Count
        Else
            LabelMessege.Text = "Empty sheet or data error!"
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
        ''''''''''''''''''''''''''''''''''''''''''''''
        ' get default value from the default pat file
        ''''''''''''''''''''''''''''''''''''''''''''''

        IDS.Data.ParameterID.RecordID = "FactoryDefault"
        IDSData.Admin.Folder.FileExtension = "Pat"
        IDSData.Admin.Folder.PatternPath = "C:\IDS\Pattern_Dir"
        IDS.Data.OpenData()
        IDS.Data.SavePathFileData(FileName + ".pat")
        ''''''''''''''''''''''''''''''''''''''''''''''
        'lim
        'IDS.Devices.Vision.SetFiducialFilename(FileName)




        Dim strUnit As String
        If True = m_Execution.m_Pattern.m_ErrorChk.WorkingUnitFlag Then
            strUnit = "mm"
        Else
            strUnit = "inch"
        End If


        Try
            If FileName <> Nothing Then
                m_Execution.m_Pattern.SavePatternParaOpt(optm, FileName + ".txt", False, _
                     strUnit)
                FrmGD.SaveGDFile(FileName) 'save GD 'lim
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
        LabelMessege.Text = "Dot optimisation file saved!"
        GC.Collect()

    End Sub

    Private Sub BT_Pro_LEdit_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BT_Pro_LEdit.CheckedChanged
        IDS.Data.Hardware.Dispenser.Left.AlowEdit = BT_Pro_LEdit.Checked
        IDS.Data.SaveData()
    End Sub

    Private Sub BT_Pro_REdit_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BT_Pro_REdit.CheckedChanged
        IDS.Data.Hardware.Dispenser.Right.AlowEdit = BT_Pro_REdit.Checked
        IDS.Data.SaveData()
    End Sub

#End Region

   
End Class
