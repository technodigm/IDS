Imports System.Math
Imports System.io
Imports OpenFileDialogPreview.Form1
Imports System.Threading
Imports Matrox.ActiveMIL

'AxEcamera, AxEboard and AxEBWImage8 are ActiveX/COM components for communicating/interfacing with the Picolo frame grabber
'Since we are not using the eVision image processing library but MIL instead, we use pointer conversion to transmit the video
'information received from AxEcamera signal to a MIL pointer.

Public Class FormVision
    Inherits System.Windows.Forms.Form

    'Another method, not to fire "ValueChanged" event while constructing Form's components.
    Private Initializing As Boolean = True

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()
        'This call is required by the Windows Form Designer.
        MyVisionCounter += 1
        InitializeComponent()
        VisionInitialization()
        'Add any initialization after the InitializeComponent() call

        Initializing = False
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
    Public WithEvents PanelVision As System.Windows.Forms.Panel
    Friend WithEvents TBReference As System.Windows.Forms.Panel
    Friend WithEvents Pos_X As System.Windows.Forms.Label
    Friend WithEvents Pos_Y As System.Windows.Forms.Label
    Friend WithEvents Pos_Z As System.Windows.Forms.Label
    Friend WithEvents Label_Cursor As System.Windows.Forms.Label
    Friend WithEvents Label_Pos_X As System.Windows.Forms.Label
    Friend WithEvents Label_Pos_Y As System.Windows.Forms.Label
    Friend WithEvents Label_Pos_Z As System.Windows.Forms.Label
    Friend WithEvents Label_Contrast As System.Windows.Forms.Label
    Friend WithEvents Label_Brightness As System.Windows.Forms.Label
    Friend WithEvents ValueBrightness As System.Windows.Forms.NumericUpDown
    Friend WithEvents ValueContrast As System.Windows.Forms.NumericUpDown
    Friend WithEvents Button_LoadImage As System.Windows.Forms.Button
    Friend WithEvents Button_SaveImage As System.Windows.Forms.Button
    Friend WithEvents Button_Close As System.Windows.Forms.Button
    Friend WithEvents SaveFileDialog_Image As System.Windows.Forms.SaveFileDialog
    Friend WithEvents OpenFileDialog_Image As System.Windows.Forms.OpenFileDialog
    Friend WithEvents Timer_Circle As System.Windows.Forms.Timer
    Friend WithEvents Button_ChipEdge As System.Windows.Forms.Button
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents Button_ChipEdgePoints As System.Windows.Forms.Button
    Friend WithEvents Button_Measurement As System.Windows.Forms.Button
    Friend WithEvents Button5 As System.Windows.Forms.Button
    Friend WithEvents AxEBoard1 As AxMULTICAMLib.AxEBoard
    Friend WithEvents AxECamera1 As AxMULTICAMLib.AxECamera
    Friend WithEvents OpenFileDialog_RM As System.Windows.Forms.OpenFileDialog
    Friend WithEvents Button10 As System.Windows.Forms.Button
    Friend WithEvents Timer_Histo As System.Windows.Forms.Timer
    Friend WithEvents AxSystem1 As AxMatrox.ActiveMIL.AxMSystem
    Friend WithEvents AxGraphicContext1 As AxMatrox.ActiveMIL.AxMGraphicContext
    Friend WithEvents AxImage3 As AxMatrox.ActiveMIL.AxMImage
    Friend WithEvents AxGraphicContext2 As AxMatrox.ActiveMIL.AxMGraphicContext
    Friend WithEvents AxImageProcessing7 As AxMatrox.ActiveMIL.ImageProcessing.AxMImageProcessing
    Friend WithEvents AxImage16 As AxMatrox.ActiveMIL.AxMImage
    Friend WithEvents AxImageProcessing9 As AxMatrox.ActiveMIL.ImageProcessing.AxMImageProcessing
    Friend WithEvents AxImage17 As AxMatrox.ActiveMIL.AxMImage
    Friend WithEvents AxImageProcessing2 As AxMatrox.ActiveMIL.ImageProcessing.AxMImageProcessing
    Friend WithEvents AxImage13 As AxMatrox.ActiveMIL.AxMImage
    Friend WithEvents AxImageProcessing6 As AxMatrox.ActiveMIL.ImageProcessing.AxMImageProcessing
    Friend WithEvents AxImage12 As AxMatrox.ActiveMIL.AxMImage
    Friend WithEvents AxBlobAnalysis2 As AxMatrox.ActiveMIL.BlobAnalysis.AxMBlobAnalysis
    Friend WithEvents AxCalibration1 As AxMatrox.ActiveMIL.Calibration.AxMCalibration
    Friend WithEvents AxPatternMatching2 As AxMatrox.ActiveMIL.PatternMatching.AxMPatternMatching
    Friend WithEvents AxGraphicContext3 As AxMatrox.ActiveMIL.AxMGraphicContext
    Friend WithEvents AxImage5 As AxMatrox.ActiveMIL.AxMImage
    Friend WithEvents AxPatternMatching1 As AxMatrox.ActiveMIL.PatternMatching.AxMPatternMatching
    Friend WithEvents AxImage2 As AxMatrox.ActiveMIL.AxMImage
    Friend WithEvents AxGraphicContext5 As AxMatrox.ActiveMIL.AxMGraphicContext
    Friend WithEvents AxImageProcessing3 As AxMatrox.ActiveMIL.ImageProcessing.AxMImageProcessing
    Friend WithEvents AxImage8 As AxMatrox.ActiveMIL.AxMImage
    Friend WithEvents AxImageProcessing4 As AxMatrox.ActiveMIL.ImageProcessing.AxMImageProcessing
    Friend WithEvents AxImage9 As AxMatrox.ActiveMIL.AxMImage
    Friend WithEvents AxImageProcessing5 As AxMatrox.ActiveMIL.ImageProcessing.AxMImageProcessing
    Friend WithEvents AxMeasurement7 As AxMatrox.ActiveMIL.Measurement.AxMMeasurement
    Friend WithEvents AxMeasurement8 As AxMatrox.ActiveMIL.Measurement.AxMMeasurement
    Friend WithEvents AxMeasurement10 As AxMatrox.ActiveMIL.Measurement.AxMMeasurement
    Friend WithEvents AxMeasurement9 As AxMatrox.ActiveMIL.Measurement.AxMMeasurement
    Friend WithEvents AxMeasurement11 As AxMatrox.ActiveMIL.Measurement.AxMMeasurement
    Friend WithEvents AxMeasurement12 As AxMatrox.ActiveMIL.Measurement.AxMMeasurement
    Friend WithEvents AxMeasurement1 As AxMatrox.ActiveMIL.Measurement.AxMMeasurement
    Friend WithEvents AxMeasurement2 As AxMatrox.ActiveMIL.Measurement.AxMMeasurement
    Friend WithEvents AxMeasurement3 As AxMatrox.ActiveMIL.Measurement.AxMMeasurement
    Friend WithEvents AxMeasurement4 As AxMatrox.ActiveMIL.Measurement.AxMMeasurement
    Friend WithEvents AxMeasurement5 As AxMatrox.ActiveMIL.Measurement.AxMMeasurement
    Friend WithEvents AxMeasurement6 As AxMatrox.ActiveMIL.Measurement.AxMMeasurement
    Friend WithEvents AxImageProcessing1 As AxMatrox.ActiveMIL.ImageProcessing.AxMImageProcessing
    Friend WithEvents AxDisplay2 As AxMatrox.ActiveMIL.AxMDisplay
    Friend WithEvents AxImage7 As AxMatrox.ActiveMIL.AxMImage
    Friend WithEvents AxImage6 As AxMatrox.ActiveMIL.AxMImage
    Friend WithEvents AxPatternMatching3 As AxMatrox.ActiveMIL.PatternMatching.AxMPatternMatching
    Friend WithEvents AxImage10 As AxMatrox.ActiveMIL.AxMImage
    Friend WithEvents AxImage11 As AxMatrox.ActiveMIL.AxMImage
    Public WithEvents chkToggle As System.Windows.Forms.CheckBox
    Public WithEvents AxDisplay1 As AxMatrox.ActiveMIL.AxMDisplay
    Friend WithEvents AxMGraphicContext1 As AxMatrox.ActiveMIL.AxMGraphicContext
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents AxEBW8Image1 As AxeVision.AxEBW8Image


    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(FormVision))
        Me.PanelVision = New System.Windows.Forms.Panel
        Me.AxEBW8Image1 = New AxeVision.AxEBW8Image
        Me.TextBox1 = New System.Windows.Forms.TextBox
        Me.chkToggle = New System.Windows.Forms.CheckBox
        Me.AxImage11 = New AxMatrox.ActiveMIL.AxMImage
        Me.AxImage10 = New AxMatrox.ActiveMIL.AxMImage
        Me.AxPatternMatching3 = New AxMatrox.ActiveMIL.PatternMatching.AxMPatternMatching
        Me.AxImage6 = New AxMatrox.ActiveMIL.AxMImage
        Me.AxImage7 = New AxMatrox.ActiveMIL.AxMImage
        Me.AxDisplay2 = New AxMatrox.ActiveMIL.AxMDisplay
        Me.AxImageProcessing1 = New AxMatrox.ActiveMIL.ImageProcessing.AxMImageProcessing
        Me.AxMeasurement6 = New AxMatrox.ActiveMIL.Measurement.AxMMeasurement
        Me.AxMeasurement5 = New AxMatrox.ActiveMIL.Measurement.AxMMeasurement
        Me.AxMeasurement4 = New AxMatrox.ActiveMIL.Measurement.AxMMeasurement
        Me.AxMeasurement3 = New AxMatrox.ActiveMIL.Measurement.AxMMeasurement
        Me.AxMeasurement2 = New AxMatrox.ActiveMIL.Measurement.AxMMeasurement
        Me.AxMeasurement1 = New AxMatrox.ActiveMIL.Measurement.AxMMeasurement
        Me.AxMeasurement12 = New AxMatrox.ActiveMIL.Measurement.AxMMeasurement
        Me.AxMeasurement11 = New AxMatrox.ActiveMIL.Measurement.AxMMeasurement
        Me.AxMeasurement9 = New AxMatrox.ActiveMIL.Measurement.AxMMeasurement
        Me.AxMeasurement10 = New AxMatrox.ActiveMIL.Measurement.AxMMeasurement
        Me.AxMeasurement8 = New AxMatrox.ActiveMIL.Measurement.AxMMeasurement
        Me.AxMeasurement7 = New AxMatrox.ActiveMIL.Measurement.AxMMeasurement
        Me.AxImageProcessing5 = New AxMatrox.ActiveMIL.ImageProcessing.AxMImageProcessing
        Me.AxImage9 = New AxMatrox.ActiveMIL.AxMImage
        Me.AxImageProcessing4 = New AxMatrox.ActiveMIL.ImageProcessing.AxMImageProcessing
        Me.AxImage8 = New AxMatrox.ActiveMIL.AxMImage
        Me.AxImageProcessing3 = New AxMatrox.ActiveMIL.ImageProcessing.AxMImageProcessing
        Me.AxGraphicContext5 = New AxMatrox.ActiveMIL.AxMGraphicContext
        Me.AxImage2 = New AxMatrox.ActiveMIL.AxMImage
        Me.AxPatternMatching1 = New AxMatrox.ActiveMIL.PatternMatching.AxMPatternMatching
        Me.AxImage5 = New AxMatrox.ActiveMIL.AxMImage
        Me.AxGraphicContext3 = New AxMatrox.ActiveMIL.AxMGraphicContext
        Me.AxPatternMatching2 = New AxMatrox.ActiveMIL.PatternMatching.AxMPatternMatching
        Me.AxCalibration1 = New AxMatrox.ActiveMIL.Calibration.AxMCalibration
        Me.AxBlobAnalysis2 = New AxMatrox.ActiveMIL.BlobAnalysis.AxMBlobAnalysis
        Me.AxImage12 = New AxMatrox.ActiveMIL.AxMImage
        Me.AxImageProcessing6 = New AxMatrox.ActiveMIL.ImageProcessing.AxMImageProcessing
        Me.AxImage13 = New AxMatrox.ActiveMIL.AxMImage
        Me.AxImageProcessing2 = New AxMatrox.ActiveMIL.ImageProcessing.AxMImageProcessing
        Me.AxImage17 = New AxMatrox.ActiveMIL.AxMImage
        Me.AxImageProcessing9 = New AxMatrox.ActiveMIL.ImageProcessing.AxMImageProcessing
        Me.AxImage16 = New AxMatrox.ActiveMIL.AxMImage
        Me.AxImageProcessing7 = New AxMatrox.ActiveMIL.ImageProcessing.AxMImageProcessing
        Me.AxGraphicContext2 = New AxMatrox.ActiveMIL.AxMGraphicContext
        Me.AxGraphicContext1 = New AxMatrox.ActiveMIL.AxMGraphicContext
        Me.AxImage3 = New AxMatrox.ActiveMIL.AxMImage
        Me.AxSystem1 = New AxMatrox.ActiveMIL.AxMSystem
        Me.Button10 = New System.Windows.Forms.Button
        Me.Button_ChipEdgePoints = New System.Windows.Forms.Button
        Me.Button3 = New System.Windows.Forms.Button
        Me.Button5 = New System.Windows.Forms.Button
        Me.AxECamera1 = New AxMULTICAMLib.AxECamera
        Me.AxEBoard1 = New AxMULTICAMLib.AxEBoard
        Me.Button_Measurement = New System.Windows.Forms.Button
        Me.Button_ChipEdge = New System.Windows.Forms.Button
        Me.Button_Close = New System.Windows.Forms.Button
        Me.Button_LoadImage = New System.Windows.Forms.Button
        Me.Button_SaveImage = New System.Windows.Forms.Button
        Me.TBReference = New System.Windows.Forms.Panel
        Me.Label_Cursor = New System.Windows.Forms.Label
        Me.Label_Contrast = New System.Windows.Forms.Label
        Me.ValueContrast = New System.Windows.Forms.NumericUpDown
        Me.Label_Brightness = New System.Windows.Forms.Label
        Me.Pos_X = New System.Windows.Forms.Label
        Me.Label_Pos_X = New System.Windows.Forms.Label
        Me.Pos_Z = New System.Windows.Forms.Label
        Me.Pos_Y = New System.Windows.Forms.Label
        Me.Label_Pos_Y = New System.Windows.Forms.Label
        Me.Label_Pos_Z = New System.Windows.Forms.Label
        Me.ValueBrightness = New System.Windows.Forms.NumericUpDown
        Me.AxDisplay1 = New AxMatrox.ActiveMIL.AxMDisplay
        Me.AxMGraphicContext1 = New AxMatrox.ActiveMIL.AxMGraphicContext
        Me.Timer_Circle = New System.Windows.Forms.Timer(Me.components)
        Me.SaveFileDialog_Image = New System.Windows.Forms.SaveFileDialog
        Me.OpenFileDialog_Image = New System.Windows.Forms.OpenFileDialog
        Me.Timer_Histo = New System.Windows.Forms.Timer(Me.components)
        Me.OpenFileDialog_RM = New System.Windows.Forms.OpenFileDialog
        Me.PanelVision.SuspendLayout()
        CType(Me.AxEBW8Image1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.AxImage11, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.AxImage10, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.AxPatternMatching3, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.AxImage6, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.AxImage7, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.AxDisplay2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.AxImageProcessing1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.AxMeasurement6, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.AxMeasurement5, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.AxMeasurement4, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.AxMeasurement3, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.AxMeasurement2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.AxMeasurement1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.AxMeasurement12, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.AxMeasurement11, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.AxMeasurement9, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.AxMeasurement10, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.AxMeasurement8, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.AxMeasurement7, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.AxImageProcessing5, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.AxImage9, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.AxImageProcessing4, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.AxImage8, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.AxImageProcessing3, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.AxGraphicContext5, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.AxImage2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.AxPatternMatching1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.AxImage5, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.AxGraphicContext3, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.AxPatternMatching2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.AxCalibration1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.AxBlobAnalysis2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.AxImage12, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.AxImageProcessing6, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.AxImage13, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.AxImageProcessing2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.AxImage17, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.AxImageProcessing9, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.AxImage16, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.AxImageProcessing7, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.AxGraphicContext2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.AxGraphicContext1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.AxImage3, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.AxSystem1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.AxECamera1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.AxEBoard1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TBReference.SuspendLayout()
        CType(Me.ValueContrast, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.ValueBrightness, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.AxDisplay1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.AxMGraphicContext1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'PanelVision
        '
        Me.PanelVision.BackColor = System.Drawing.SystemColors.Control
        Me.PanelVision.Controls.Add(Me.AxEBW8Image1)
        Me.PanelVision.Controls.Add(Me.TextBox1)
        Me.PanelVision.Controls.Add(Me.chkToggle)
        Me.PanelVision.Controls.Add(Me.AxImage11)
        Me.PanelVision.Controls.Add(Me.AxImage10)
        Me.PanelVision.Controls.Add(Me.AxPatternMatching3)
        Me.PanelVision.Controls.Add(Me.AxImage6)
        Me.PanelVision.Controls.Add(Me.AxImage7)
        Me.PanelVision.Controls.Add(Me.AxDisplay2)
        Me.PanelVision.Controls.Add(Me.AxImageProcessing1)
        Me.PanelVision.Controls.Add(Me.AxMeasurement6)
        Me.PanelVision.Controls.Add(Me.AxMeasurement5)
        Me.PanelVision.Controls.Add(Me.AxMeasurement4)
        Me.PanelVision.Controls.Add(Me.AxMeasurement3)
        Me.PanelVision.Controls.Add(Me.AxMeasurement2)
        Me.PanelVision.Controls.Add(Me.AxMeasurement1)
        Me.PanelVision.Controls.Add(Me.AxMeasurement12)
        Me.PanelVision.Controls.Add(Me.AxMeasurement11)
        Me.PanelVision.Controls.Add(Me.AxMeasurement9)
        Me.PanelVision.Controls.Add(Me.AxMeasurement10)
        Me.PanelVision.Controls.Add(Me.AxMeasurement8)
        Me.PanelVision.Controls.Add(Me.AxMeasurement7)
        Me.PanelVision.Controls.Add(Me.AxImageProcessing5)
        Me.PanelVision.Controls.Add(Me.AxImage9)
        Me.PanelVision.Controls.Add(Me.AxImageProcessing4)
        Me.PanelVision.Controls.Add(Me.AxImage8)
        Me.PanelVision.Controls.Add(Me.AxImageProcessing3)
        Me.PanelVision.Controls.Add(Me.AxGraphicContext5)
        Me.PanelVision.Controls.Add(Me.AxImage2)
        Me.PanelVision.Controls.Add(Me.AxPatternMatching1)
        Me.PanelVision.Controls.Add(Me.AxImage5)
        Me.PanelVision.Controls.Add(Me.AxGraphicContext3)
        Me.PanelVision.Controls.Add(Me.AxPatternMatching2)
        Me.PanelVision.Controls.Add(Me.AxCalibration1)
        Me.PanelVision.Controls.Add(Me.AxBlobAnalysis2)
        Me.PanelVision.Controls.Add(Me.AxImage12)
        Me.PanelVision.Controls.Add(Me.AxImageProcessing6)
        Me.PanelVision.Controls.Add(Me.AxImage13)
        Me.PanelVision.Controls.Add(Me.AxImageProcessing2)
        Me.PanelVision.Controls.Add(Me.AxImage17)
        Me.PanelVision.Controls.Add(Me.AxImageProcessing9)
        Me.PanelVision.Controls.Add(Me.AxImage16)
        Me.PanelVision.Controls.Add(Me.AxImageProcessing7)
        Me.PanelVision.Controls.Add(Me.AxGraphicContext2)
        Me.PanelVision.Controls.Add(Me.AxGraphicContext1)
        Me.PanelVision.Controls.Add(Me.AxImage3)
        Me.PanelVision.Controls.Add(Me.AxSystem1)
        Me.PanelVision.Controls.Add(Me.Button10)
        Me.PanelVision.Controls.Add(Me.Button_ChipEdgePoints)
        Me.PanelVision.Controls.Add(Me.Button3)
        Me.PanelVision.Controls.Add(Me.Button5)
        Me.PanelVision.Controls.Add(Me.AxECamera1)
        Me.PanelVision.Controls.Add(Me.AxEBoard1)
        Me.PanelVision.Controls.Add(Me.Button_Measurement)
        Me.PanelVision.Controls.Add(Me.Button_ChipEdge)
        Me.PanelVision.Controls.Add(Me.Button_Close)
        Me.PanelVision.Controls.Add(Me.Button_LoadImage)
        Me.PanelVision.Controls.Add(Me.Button_SaveImage)
        Me.PanelVision.Controls.Add(Me.TBReference)
        Me.PanelVision.Controls.Add(Me.AxDisplay1)
        Me.PanelVision.Controls.Add(Me.AxMGraphicContext1)
        Me.PanelVision.Location = New System.Drawing.Point(0, 0)
        Me.PanelVision.Name = "PanelVision"
        Me.PanelVision.Size = New System.Drawing.Size(1280, 672)
        Me.PanelVision.TabIndex = 71
        '
        'AxEBW8Image1
        '
        Me.AxEBW8Image1.ContainingControl = Me
        Me.AxEBW8Image1.Enabled = True
        Me.AxEBW8Image1.Location = New System.Drawing.Point(864, 352)
        Me.AxEBW8Image1.Name = "AxEBW8Image1"
        Me.AxEBW8Image1.OcxState = CType(resources.GetObject("AxEBW8Image1.OcxState"), System.Windows.Forms.AxHost.State)
        Me.AxEBW8Image1.Size = New System.Drawing.Size(768, 576)
        Me.AxEBW8Image1.TabIndex = 268
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(512, 632)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(152, 20)
        Me.TextBox1.TabIndex = 267
        Me.TextBox1.Text = "Camera errors are written here"
        '
        'chkToggle
        '
        Me.chkToggle.Appearance = System.Windows.Forms.Appearance.Button
        Me.chkToggle.CheckAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.chkToggle.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkToggle.Location = New System.Drawing.Point(0, 576)
        Me.chkToggle.Name = "chkToggle"
        Me.chkToggle.TabIndex = 266
        Me.chkToggle.Text = "Teach Camera"
        Me.chkToggle.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'AxImage11
        '
        Me.AxImage11.ContainingControl = Me
        Me.AxImage11.Enabled = True
        Me.AxImage11.Location = New System.Drawing.Point(768, 352)
        Me.AxImage11.Name = "AxImage11"
        Me.AxImage11.OcxState = CType(resources.GetObject("AxImage11.OcxState"), System.Windows.Forms.AxHost.State)
        Me.AxImage11.Size = New System.Drawing.Size(32, 32)
        Me.AxImage11.TabIndex = 263
        '
        'AxImage10
        '
        Me.AxImage10.ContainingControl = Me
        Me.AxImage10.Enabled = True
        Me.AxImage10.Location = New System.Drawing.Point(800, 384)
        Me.AxImage10.Name = "AxImage10"
        Me.AxImage10.OcxState = CType(resources.GetObject("AxImage10.OcxState"), System.Windows.Forms.AxHost.State)
        Me.AxImage10.Size = New System.Drawing.Size(32, 32)
        Me.AxImage10.TabIndex = 262
        '
        'AxPatternMatching3
        '
        Me.AxPatternMatching3.ContainingControl = Me
        Me.AxPatternMatching3.Enabled = True
        Me.AxPatternMatching3.Location = New System.Drawing.Point(800, 160)
        Me.AxPatternMatching3.Name = "AxPatternMatching3"
        Me.AxPatternMatching3.OcxState = CType(resources.GetObject("AxPatternMatching3.OcxState"), System.Windows.Forms.AxHost.State)
        Me.AxPatternMatching3.Size = New System.Drawing.Size(32, 32)
        Me.AxPatternMatching3.TabIndex = 261
        '
        'AxImage6
        '
        Me.AxImage6.ContainingControl = Me
        Me.AxImage6.Enabled = True
        Me.AxImage6.Location = New System.Drawing.Point(768, 192)
        Me.AxImage6.Name = "AxImage6"
        Me.AxImage6.OcxState = CType(resources.GetObject("AxImage6.OcxState"), System.Windows.Forms.AxHost.State)
        Me.AxImage6.Size = New System.Drawing.Size(32, 32)
        Me.AxImage6.TabIndex = 260
        '
        'AxImage7
        '
        Me.AxImage7.ContainingControl = Me
        Me.AxImage7.Enabled = True
        Me.AxImage7.Location = New System.Drawing.Point(800, 224)
        Me.AxImage7.Name = "AxImage7"
        Me.AxImage7.OcxState = CType(resources.GetObject("AxImage7.OcxState"), System.Windows.Forms.AxHost.State)
        Me.AxImage7.Size = New System.Drawing.Size(32, 32)
        Me.AxImage7.TabIndex = 259
        '
        'AxDisplay2
        '
        Me.AxDisplay2.ContainingControl = Me
        Me.AxDisplay2.Enabled = True
        Me.AxDisplay2.Location = New System.Drawing.Point(800, 552)
        Me.AxDisplay2.Name = "AxDisplay2"
        Me.AxDisplay2.OcxState = CType(resources.GetObject("AxDisplay2.OcxState"), System.Windows.Forms.AxHost.State)
        Me.AxDisplay2.Size = New System.Drawing.Size(768, 576)
        Me.AxDisplay2.TabIndex = 258
        '
        'AxImageProcessing1
        '
        Me.AxImageProcessing1.ContainingControl = Me
        Me.AxImageProcessing1.Enabled = True
        Me.AxImageProcessing1.Location = New System.Drawing.Point(864, 96)
        Me.AxImageProcessing1.Name = "AxImageProcessing1"
        Me.AxImageProcessing1.OcxState = CType(resources.GetObject("AxImageProcessing1.OcxState"), System.Windows.Forms.AxHost.State)
        Me.AxImageProcessing1.Size = New System.Drawing.Size(32, 32)
        Me.AxImageProcessing1.TabIndex = 257
        '
        'AxMeasurement6
        '
        Me.AxMeasurement6.ContainingControl = Me
        Me.AxMeasurement6.Enabled = True
        Me.AxMeasurement6.Location = New System.Drawing.Point(864, 160)
        Me.AxMeasurement6.Name = "AxMeasurement6"
        Me.AxMeasurement6.OcxState = CType(resources.GetObject("AxMeasurement6.OcxState"), System.Windows.Forms.AxHost.State)
        Me.AxMeasurement6.Size = New System.Drawing.Size(32, 32)
        Me.AxMeasurement6.TabIndex = 256
        '
        'AxMeasurement5
        '
        Me.AxMeasurement5.ContainingControl = Me
        Me.AxMeasurement5.Enabled = True
        Me.AxMeasurement5.Location = New System.Drawing.Point(864, 128)
        Me.AxMeasurement5.Name = "AxMeasurement5"
        Me.AxMeasurement5.OcxState = CType(resources.GetObject("AxMeasurement5.OcxState"), System.Windows.Forms.AxHost.State)
        Me.AxMeasurement5.Size = New System.Drawing.Size(32, 32)
        Me.AxMeasurement5.TabIndex = 255
        '
        'AxMeasurement4
        '
        Me.AxMeasurement4.ContainingControl = Me
        Me.AxMeasurement4.Enabled = True
        Me.AxMeasurement4.Location = New System.Drawing.Point(832, 160)
        Me.AxMeasurement4.Name = "AxMeasurement4"
        Me.AxMeasurement4.OcxState = CType(resources.GetObject("AxMeasurement4.OcxState"), System.Windows.Forms.AxHost.State)
        Me.AxMeasurement4.Size = New System.Drawing.Size(32, 32)
        Me.AxMeasurement4.TabIndex = 254
        '
        'AxMeasurement3
        '
        Me.AxMeasurement3.ContainingControl = Me
        Me.AxMeasurement3.Enabled = True
        Me.AxMeasurement3.Location = New System.Drawing.Point(832, 128)
        Me.AxMeasurement3.Name = "AxMeasurement3"
        Me.AxMeasurement3.OcxState = CType(resources.GetObject("AxMeasurement3.OcxState"), System.Windows.Forms.AxHost.State)
        Me.AxMeasurement3.Size = New System.Drawing.Size(32, 32)
        Me.AxMeasurement3.TabIndex = 253
        '
        'AxMeasurement2
        '
        Me.AxMeasurement2.ContainingControl = Me
        Me.AxMeasurement2.Enabled = True
        Me.AxMeasurement2.Location = New System.Drawing.Point(832, 288)
        Me.AxMeasurement2.Name = "AxMeasurement2"
        Me.AxMeasurement2.OcxState = CType(resources.GetObject("AxMeasurement2.OcxState"), System.Windows.Forms.AxHost.State)
        Me.AxMeasurement2.Size = New System.Drawing.Size(32, 32)
        Me.AxMeasurement2.TabIndex = 252
        '
        'AxMeasurement1
        '
        Me.AxMeasurement1.ContainingControl = Me
        Me.AxMeasurement1.Enabled = True
        Me.AxMeasurement1.Location = New System.Drawing.Point(832, 256)
        Me.AxMeasurement1.Name = "AxMeasurement1"
        Me.AxMeasurement1.OcxState = CType(resources.GetObject("AxMeasurement1.OcxState"), System.Windows.Forms.AxHost.State)
        Me.AxMeasurement1.Size = New System.Drawing.Size(32, 32)
        Me.AxMeasurement1.TabIndex = 251
        '
        'AxMeasurement12
        '
        Me.AxMeasurement12.ContainingControl = Me
        Me.AxMeasurement12.Enabled = True
        Me.AxMeasurement12.Location = New System.Drawing.Point(864, 256)
        Me.AxMeasurement12.Name = "AxMeasurement12"
        Me.AxMeasurement12.OcxState = CType(resources.GetObject("AxMeasurement12.OcxState"), System.Windows.Forms.AxHost.State)
        Me.AxMeasurement12.Size = New System.Drawing.Size(32, 32)
        Me.AxMeasurement12.TabIndex = 250
        '
        'AxMeasurement11
        '
        Me.AxMeasurement11.ContainingControl = Me
        Me.AxMeasurement11.Enabled = True
        Me.AxMeasurement11.Location = New System.Drawing.Point(832, 192)
        Me.AxMeasurement11.Name = "AxMeasurement11"
        Me.AxMeasurement11.OcxState = CType(resources.GetObject("AxMeasurement11.OcxState"), System.Windows.Forms.AxHost.State)
        Me.AxMeasurement11.Size = New System.Drawing.Size(32, 32)
        Me.AxMeasurement11.TabIndex = 249
        '
        'AxMeasurement9
        '
        Me.AxMeasurement9.ContainingControl = Me
        Me.AxMeasurement9.Enabled = True
        Me.AxMeasurement9.Location = New System.Drawing.Point(864, 288)
        Me.AxMeasurement9.Name = "AxMeasurement9"
        Me.AxMeasurement9.OcxState = CType(resources.GetObject("AxMeasurement9.OcxState"), System.Windows.Forms.AxHost.State)
        Me.AxMeasurement9.Size = New System.Drawing.Size(32, 32)
        Me.AxMeasurement9.TabIndex = 248
        '
        'AxMeasurement10
        '
        Me.AxMeasurement10.ContainingControl = Me
        Me.AxMeasurement10.Enabled = True
        Me.AxMeasurement10.Location = New System.Drawing.Point(864, 192)
        Me.AxMeasurement10.Name = "AxMeasurement10"
        Me.AxMeasurement10.OcxState = CType(resources.GetObject("AxMeasurement10.OcxState"), System.Windows.Forms.AxHost.State)
        Me.AxMeasurement10.Size = New System.Drawing.Size(32, 32)
        Me.AxMeasurement10.TabIndex = 247
        '
        'AxMeasurement8
        '
        Me.AxMeasurement8.ContainingControl = Me
        Me.AxMeasurement8.Enabled = True
        Me.AxMeasurement8.Location = New System.Drawing.Point(864, 224)
        Me.AxMeasurement8.Name = "AxMeasurement8"
        Me.AxMeasurement8.OcxState = CType(resources.GetObject("AxMeasurement8.OcxState"), System.Windows.Forms.AxHost.State)
        Me.AxMeasurement8.Size = New System.Drawing.Size(32, 32)
        Me.AxMeasurement8.TabIndex = 246
        '
        'AxMeasurement7
        '
        Me.AxMeasurement7.ContainingControl = Me
        Me.AxMeasurement7.Enabled = True
        Me.AxMeasurement7.Location = New System.Drawing.Point(832, 224)
        Me.AxMeasurement7.Name = "AxMeasurement7"
        Me.AxMeasurement7.OcxState = CType(resources.GetObject("AxMeasurement7.OcxState"), System.Windows.Forms.AxHost.State)
        Me.AxMeasurement7.Size = New System.Drawing.Size(32, 32)
        Me.AxMeasurement7.TabIndex = 245
        '
        'AxImageProcessing5
        '
        Me.AxImageProcessing5.ContainingControl = Me
        Me.AxImageProcessing5.Enabled = True
        Me.AxImageProcessing5.Location = New System.Drawing.Point(800, 96)
        Me.AxImageProcessing5.Name = "AxImageProcessing5"
        Me.AxImageProcessing5.OcxState = CType(resources.GetObject("AxImageProcessing5.OcxState"), System.Windows.Forms.AxHost.State)
        Me.AxImageProcessing5.Size = New System.Drawing.Size(32, 32)
        Me.AxImageProcessing5.TabIndex = 244
        '
        'AxImage9
        '
        Me.AxImage9.ContainingControl = Me
        Me.AxImage9.Enabled = True
        Me.AxImage9.Location = New System.Drawing.Point(768, 384)
        Me.AxImage9.Name = "AxImage9"
        Me.AxImage9.OcxState = CType(resources.GetObject("AxImage9.OcxState"), System.Windows.Forms.AxHost.State)
        Me.AxImage9.Size = New System.Drawing.Size(32, 32)
        Me.AxImage9.TabIndex = 243
        '
        'AxImageProcessing4
        '
        Me.AxImageProcessing4.ContainingControl = Me
        Me.AxImageProcessing4.Enabled = True
        Me.AxImageProcessing4.Location = New System.Drawing.Point(768, 96)
        Me.AxImageProcessing4.Name = "AxImageProcessing4"
        Me.AxImageProcessing4.OcxState = CType(resources.GetObject("AxImageProcessing4.OcxState"), System.Windows.Forms.AxHost.State)
        Me.AxImageProcessing4.Size = New System.Drawing.Size(32, 32)
        Me.AxImageProcessing4.TabIndex = 242
        '
        'AxImage8
        '
        Me.AxImage8.ContainingControl = Me
        Me.AxImage8.Enabled = True
        Me.AxImage8.Location = New System.Drawing.Point(800, 352)
        Me.AxImage8.Name = "AxImage8"
        Me.AxImage8.OcxState = CType(resources.GetObject("AxImage8.OcxState"), System.Windows.Forms.AxHost.State)
        Me.AxImage8.Size = New System.Drawing.Size(32, 32)
        Me.AxImage8.TabIndex = 241
        '
        'AxImageProcessing3
        '
        Me.AxImageProcessing3.ContainingControl = Me
        Me.AxImageProcessing3.Enabled = True
        Me.AxImageProcessing3.Location = New System.Drawing.Point(864, 64)
        Me.AxImageProcessing3.Name = "AxImageProcessing3"
        Me.AxImageProcessing3.OcxState = CType(resources.GetObject("AxImageProcessing3.OcxState"), System.Windows.Forms.AxHost.State)
        Me.AxImageProcessing3.Size = New System.Drawing.Size(32, 32)
        Me.AxImageProcessing3.TabIndex = 240
        '
        'AxGraphicContext5
        '
        Me.AxGraphicContext5.ContainingControl = Me
        Me.AxGraphicContext5.Enabled = True
        Me.AxGraphicContext5.Location = New System.Drawing.Point(896, 0)
        Me.AxGraphicContext5.Name = "AxGraphicContext5"
        Me.AxGraphicContext5.OcxState = CType(resources.GetObject("AxGraphicContext5.OcxState"), System.Windows.Forms.AxHost.State)
        Me.AxGraphicContext5.Size = New System.Drawing.Size(32, 32)
        Me.AxGraphicContext5.TabIndex = 239
        '
        'AxImage2
        '
        Me.AxImage2.ContainingControl = Me
        Me.AxImage2.Enabled = True
        Me.AxImage2.Location = New System.Drawing.Point(800, 256)
        Me.AxImage2.Name = "AxImage2"
        Me.AxImage2.OcxState = CType(resources.GetObject("AxImage2.OcxState"), System.Windows.Forms.AxHost.State)
        Me.AxImage2.Size = New System.Drawing.Size(32, 32)
        Me.AxImage2.TabIndex = 238
        '
        'AxPatternMatching1
        '
        Me.AxPatternMatching1.ContainingControl = Me
        Me.AxPatternMatching1.Enabled = True
        Me.AxPatternMatching1.Location = New System.Drawing.Point(800, 192)
        Me.AxPatternMatching1.Name = "AxPatternMatching1"
        Me.AxPatternMatching1.OcxState = CType(resources.GetObject("AxPatternMatching1.OcxState"), System.Windows.Forms.AxHost.State)
        Me.AxPatternMatching1.Size = New System.Drawing.Size(32, 32)
        Me.AxPatternMatching1.TabIndex = 237
        '
        'AxImage5
        '
        Me.AxImage5.ContainingControl = Me
        Me.AxImage5.Enabled = True
        Me.AxImage5.Location = New System.Drawing.Point(768, 256)
        Me.AxImage5.Name = "AxImage5"
        Me.AxImage5.OcxState = CType(resources.GetObject("AxImage5.OcxState"), System.Windows.Forms.AxHost.State)
        Me.AxImage5.Size = New System.Drawing.Size(32, 32)
        Me.AxImage5.TabIndex = 236
        '
        'AxGraphicContext3
        '
        Me.AxGraphicContext3.ContainingControl = Me
        Me.AxGraphicContext3.Enabled = True
        Me.AxGraphicContext3.Location = New System.Drawing.Point(832, 0)
        Me.AxGraphicContext3.Name = "AxGraphicContext3"
        Me.AxGraphicContext3.OcxState = CType(resources.GetObject("AxGraphicContext3.OcxState"), System.Windows.Forms.AxHost.State)
        Me.AxGraphicContext3.Size = New System.Drawing.Size(32, 32)
        Me.AxGraphicContext3.TabIndex = 235
        '
        'AxPatternMatching2
        '
        Me.AxPatternMatching2.ContainingControl = Me
        Me.AxPatternMatching2.Enabled = True
        Me.AxPatternMatching2.Location = New System.Drawing.Point(768, 160)
        Me.AxPatternMatching2.Name = "AxPatternMatching2"
        Me.AxPatternMatching2.OcxState = CType(resources.GetObject("AxPatternMatching2.OcxState"), System.Windows.Forms.AxHost.State)
        Me.AxPatternMatching2.Size = New System.Drawing.Size(32, 32)
        Me.AxPatternMatching2.TabIndex = 234
        '
        'AxCalibration1
        '
        Me.AxCalibration1.ContainingControl = Me
        Me.AxCalibration1.Enabled = True
        Me.AxCalibration1.Location = New System.Drawing.Point(768, 128)
        Me.AxCalibration1.Name = "AxCalibration1"
        Me.AxCalibration1.OcxState = CType(resources.GetObject("AxCalibration1.OcxState"), System.Windows.Forms.AxHost.State)
        Me.AxCalibration1.Size = New System.Drawing.Size(32, 32)
        Me.AxCalibration1.TabIndex = 233
        '
        'AxBlobAnalysis2
        '
        Me.AxBlobAnalysis2.ContainingControl = Me
        Me.AxBlobAnalysis2.Enabled = True
        Me.AxBlobAnalysis2.Location = New System.Drawing.Point(800, 128)
        Me.AxBlobAnalysis2.Name = "AxBlobAnalysis2"
        Me.AxBlobAnalysis2.OcxState = CType(resources.GetObject("AxBlobAnalysis2.OcxState"), System.Windows.Forms.AxHost.State)
        Me.AxBlobAnalysis2.Size = New System.Drawing.Size(32, 32)
        Me.AxBlobAnalysis2.TabIndex = 232
        '
        'AxImage12
        '
        Me.AxImage12.ContainingControl = Me
        Me.AxImage12.Enabled = True
        Me.AxImage12.Location = New System.Drawing.Point(800, 320)
        Me.AxImage12.Name = "AxImage12"
        Me.AxImage12.OcxState = CType(resources.GetObject("AxImage12.OcxState"), System.Windows.Forms.AxHost.State)
        Me.AxImage12.Size = New System.Drawing.Size(32, 32)
        Me.AxImage12.TabIndex = 231
        '
        'AxImageProcessing6
        '
        Me.AxImageProcessing6.ContainingControl = Me
        Me.AxImageProcessing6.Enabled = True
        Me.AxImageProcessing6.Location = New System.Drawing.Point(832, 64)
        Me.AxImageProcessing6.Name = "AxImageProcessing6"
        Me.AxImageProcessing6.OcxState = CType(resources.GetObject("AxImageProcessing6.OcxState"), System.Windows.Forms.AxHost.State)
        Me.AxImageProcessing6.Size = New System.Drawing.Size(32, 32)
        Me.AxImageProcessing6.TabIndex = 230
        '
        'AxImage13
        '
        Me.AxImage13.ContainingControl = Me
        Me.AxImage13.Enabled = True
        Me.AxImage13.Location = New System.Drawing.Point(768, 288)
        Me.AxImage13.Name = "AxImage13"
        Me.AxImage13.OcxState = CType(resources.GetObject("AxImage13.OcxState"), System.Windows.Forms.AxHost.State)
        Me.AxImage13.Size = New System.Drawing.Size(32, 32)
        Me.AxImage13.TabIndex = 229
        '
        'AxImageProcessing2
        '
        Me.AxImageProcessing2.ContainingControl = Me
        Me.AxImageProcessing2.Enabled = True
        Me.AxImageProcessing2.Location = New System.Drawing.Point(800, 64)
        Me.AxImageProcessing2.Name = "AxImageProcessing2"
        Me.AxImageProcessing2.OcxState = CType(resources.GetObject("AxImageProcessing2.OcxState"), System.Windows.Forms.AxHost.State)
        Me.AxImageProcessing2.Size = New System.Drawing.Size(32, 32)
        Me.AxImageProcessing2.TabIndex = 228
        '
        'AxImage17
        '
        Me.AxImage17.ContainingControl = Me
        Me.AxImage17.Enabled = True
        Me.AxImage17.Location = New System.Drawing.Point(768, 320)
        Me.AxImage17.Name = "AxImage17"
        Me.AxImage17.OcxState = CType(resources.GetObject("AxImage17.OcxState"), System.Windows.Forms.AxHost.State)
        Me.AxImage17.Size = New System.Drawing.Size(32, 32)
        Me.AxImage17.TabIndex = 227
        '
        'AxImageProcessing9
        '
        Me.AxImageProcessing9.ContainingControl = Me
        Me.AxImageProcessing9.Enabled = True
        Me.AxImageProcessing9.Location = New System.Drawing.Point(768, 64)
        Me.AxImageProcessing9.Name = "AxImageProcessing9"
        Me.AxImageProcessing9.OcxState = CType(resources.GetObject("AxImageProcessing9.OcxState"), System.Windows.Forms.AxHost.State)
        Me.AxImageProcessing9.Size = New System.Drawing.Size(32, 32)
        Me.AxImageProcessing9.TabIndex = 226
        '
        'AxImage16
        '
        Me.AxImage16.ContainingControl = Me
        Me.AxImage16.Enabled = True
        Me.AxImage16.Location = New System.Drawing.Point(800, 288)
        Me.AxImage16.Name = "AxImage16"
        Me.AxImage16.OcxState = CType(resources.GetObject("AxImage16.OcxState"), System.Windows.Forms.AxHost.State)
        Me.AxImage16.Size = New System.Drawing.Size(32, 32)
        Me.AxImage16.TabIndex = 225
        '
        'AxImageProcessing7
        '
        Me.AxImageProcessing7.ContainingControl = Me
        Me.AxImageProcessing7.Enabled = True
        Me.AxImageProcessing7.Location = New System.Drawing.Point(832, 96)
        Me.AxImageProcessing7.Name = "AxImageProcessing7"
        Me.AxImageProcessing7.OcxState = CType(resources.GetObject("AxImageProcessing7.OcxState"), System.Windows.Forms.AxHost.State)
        Me.AxImageProcessing7.Size = New System.Drawing.Size(32, 32)
        Me.AxImageProcessing7.TabIndex = 224
        '
        'AxGraphicContext2
        '
        Me.AxGraphicContext2.ContainingControl = Me
        Me.AxGraphicContext2.Enabled = True
        Me.AxGraphicContext2.Location = New System.Drawing.Point(800, 0)
        Me.AxGraphicContext2.Name = "AxGraphicContext2"
        Me.AxGraphicContext2.OcxState = CType(resources.GetObject("AxGraphicContext2.OcxState"), System.Windows.Forms.AxHost.State)
        Me.AxGraphicContext2.Size = New System.Drawing.Size(32, 32)
        Me.AxGraphicContext2.TabIndex = 223
        '
        'AxGraphicContext1
        '
        Me.AxGraphicContext1.ContainingControl = Me
        Me.AxGraphicContext1.Enabled = True
        Me.AxGraphicContext1.Location = New System.Drawing.Point(768, 0)
        Me.AxGraphicContext1.Name = "AxGraphicContext1"
        Me.AxGraphicContext1.OcxState = CType(resources.GetObject("AxGraphicContext1.OcxState"), System.Windows.Forms.AxHost.State)
        Me.AxGraphicContext1.Size = New System.Drawing.Size(32, 32)
        Me.AxGraphicContext1.TabIndex = 221
        '
        'AxImage3
        '
        Me.AxImage3.ContainingControl = Me
        Me.AxImage3.Enabled = True
        Me.AxImage3.Location = New System.Drawing.Point(768, 224)
        Me.AxImage3.Name = "AxImage3"
        Me.AxImage3.OcxState = CType(resources.GetObject("AxImage3.OcxState"), System.Windows.Forms.AxHost.State)
        Me.AxImage3.Size = New System.Drawing.Size(32, 32)
        Me.AxImage3.TabIndex = 220
        '
        'AxSystem1
        '
        Me.AxSystem1.ContainingControl = Me
        Me.AxSystem1.Enabled = True
        Me.AxSystem1.Location = New System.Drawing.Point(768, 32)
        Me.AxSystem1.Name = "AxSystem1"
        Me.AxSystem1.OcxState = CType(resources.GetObject("AxSystem1.OcxState"), System.Windows.Forms.AxHost.State)
        Me.AxSystem1.Size = New System.Drawing.Size(32, 32)
        Me.AxSystem1.TabIndex = 219
        '
        'Button10
        '
        Me.Button10.Location = New System.Drawing.Point(120, 576)
        Me.Button10.Name = "Button10"
        Me.Button10.Size = New System.Drawing.Size(104, 23)
        Me.Button10.TabIndex = 212
        Me.Button10.Text = "NeedleCalibration"
        '
        'Button_ChipEdgePoints
        '
        Me.Button_ChipEdgePoints.Location = New System.Drawing.Point(232, 576)
        Me.Button_ChipEdgePoints.Name = "Button_ChipEdgePoints"
        Me.Button_ChipEdgePoints.Size = New System.Drawing.Size(104, 24)
        Me.Button_ChipEdgePoints.TabIndex = 166
        Me.Button_ChipEdgePoints.Text = "ChipEdgePoints"
        '
        'Button3
        '
        Me.Button3.Location = New System.Drawing.Point(0, 600)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(104, 24)
        Me.Button3.TabIndex = 160
        Me.Button3.Text = "RejectMark"
        '
        'Button5
        '
        Me.Button5.Location = New System.Drawing.Point(344, 576)
        Me.Button5.Name = "Button5"
        Me.Button5.Size = New System.Drawing.Size(104, 24)
        Me.Button5.TabIndex = 181
        Me.Button5.Text = "QC"
        '
        'AxECamera1
        '
        Me.AxECamera1.ContainingControl = Me
        Me.AxECamera1.Enabled = True
        Me.AxECamera1.Location = New System.Drawing.Point(832, 32)
        Me.AxECamera1.Name = "AxECamera1"
        Me.AxECamera1.OcxState = CType(resources.GetObject("AxECamera1.OcxState"), System.Windows.Forms.AxHost.State)
        Me.AxECamera1.Size = New System.Drawing.Size(32, 32)
        Me.AxECamera1.TabIndex = 203
        '
        'AxEBoard1
        '
        Me.AxEBoard1.ContainingControl = Me
        Me.AxEBoard1.Enabled = True
        Me.AxEBoard1.Location = New System.Drawing.Point(800, 32)
        Me.AxEBoard1.Name = "AxEBoard1"
        Me.AxEBoard1.OcxState = CType(resources.GetObject("AxEBoard1.OcxState"), System.Windows.Forms.AxHost.State)
        Me.AxEBoard1.Size = New System.Drawing.Size(32, 32)
        Me.AxEBoard1.TabIndex = 202
        '
        'Button_Measurement
        '
        Me.Button_Measurement.Location = New System.Drawing.Point(120, 600)
        Me.Button_Measurement.Name = "Button_Measurement"
        Me.Button_Measurement.Size = New System.Drawing.Size(104, 24)
        Me.Button_Measurement.TabIndex = 179
        Me.Button_Measurement.Text = "Measurement"
        '
        'Button_ChipEdge
        '
        Me.Button_ChipEdge.Location = New System.Drawing.Point(344, 600)
        Me.Button_ChipEdge.Name = "Button_ChipEdge"
        Me.Button_ChipEdge.Size = New System.Drawing.Size(104, 24)
        Me.Button_ChipEdge.TabIndex = 153
        Me.Button_ChipEdge.Text = "ChipEdge"
        '
        'Button_Close
        '
        Me.Button_Close.Location = New System.Drawing.Point(704, 592)
        Me.Button_Close.Name = "Button_Close"
        Me.Button_Close.Size = New System.Drawing.Size(56, 24)
        Me.Button_Close.TabIndex = 79
        Me.Button_Close.Text = "Close"
        '
        'Button_LoadImage
        '
        Me.Button_LoadImage.Location = New System.Drawing.Point(584, 592)
        Me.Button_LoadImage.Name = "Button_LoadImage"
        Me.Button_LoadImage.Size = New System.Drawing.Size(104, 24)
        Me.Button_LoadImage.TabIndex = 77
        Me.Button_LoadImage.Text = "LoadImage"
        '
        'Button_SaveImage
        '
        Me.Button_SaveImage.Location = New System.Drawing.Point(488, 592)
        Me.Button_SaveImage.Name = "Button_SaveImage"
        Me.Button_SaveImage.Size = New System.Drawing.Size(96, 24)
        Me.Button_SaveImage.TabIndex = 76
        Me.Button_SaveImage.Text = "SaveImage"
        '
        'TBReference
        '
        Me.TBReference.BackColor = System.Drawing.Color.Gainsboro
        Me.TBReference.Controls.Add(Me.Label_Cursor)
        Me.TBReference.Controls.Add(Me.Label_Contrast)
        Me.TBReference.Controls.Add(Me.ValueContrast)
        Me.TBReference.Controls.Add(Me.Label_Brightness)
        Me.TBReference.Controls.Add(Me.Pos_X)
        Me.TBReference.Controls.Add(Me.Label_Pos_X)
        Me.TBReference.Controls.Add(Me.Pos_Z)
        Me.TBReference.Controls.Add(Me.Pos_Y)
        Me.TBReference.Controls.Add(Me.Label_Pos_Y)
        Me.TBReference.Controls.Add(Me.Label_Pos_Z)
        Me.TBReference.Controls.Add(Me.ValueBrightness)
        Me.TBReference.Location = New System.Drawing.Point(0, 632)
        Me.TBReference.Name = "TBReference"
        Me.TBReference.Size = New System.Drawing.Size(480, 40)
        Me.TBReference.TabIndex = 61
        '
        'Label_Cursor
        '
        Me.Label_Cursor.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label_Cursor.ForeColor = System.Drawing.Color.Black
        Me.Label_Cursor.Location = New System.Drawing.Point(0, 8)
        Me.Label_Cursor.Name = "Label_Cursor"
        Me.Label_Cursor.Size = New System.Drawing.Size(48, 16)
        Me.Label_Cursor.TabIndex = 56
        Me.Label_Cursor.Text = "Cursor:"
        '
        'Label_Contrast
        '
        Me.Label_Contrast.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label_Contrast.ForeColor = System.Drawing.Color.Black
        Me.Label_Contrast.Location = New System.Drawing.Point(240, 8)
        Me.Label_Contrast.Name = "Label_Contrast"
        Me.Label_Contrast.Size = New System.Drawing.Size(56, 16)
        Me.Label_Contrast.TabIndex = 53
        Me.Label_Contrast.Text = "Contrast"
        '
        'ValueContrast
        '
        Me.ValueContrast.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ValueContrast.Location = New System.Drawing.Point(304, 8)
        Me.ValueContrast.Maximum = New Decimal(New Integer() {255, 0, 0, 0})
        Me.ValueContrast.Name = "ValueContrast"
        Me.ValueContrast.Size = New System.Drawing.Size(48, 20)
        Me.ValueContrast.TabIndex = 54
        Me.ValueContrast.Value = New Decimal(New Integer() {50, 0, 0, 0})
        '
        'Label_Brightness
        '
        Me.Label_Brightness.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label_Brightness.ForeColor = System.Drawing.Color.Black
        Me.Label_Brightness.Location = New System.Drawing.Point(352, 8)
        Me.Label_Brightness.Name = "Label_Brightness"
        Me.Label_Brightness.Size = New System.Drawing.Size(64, 16)
        Me.Label_Brightness.TabIndex = 52
        Me.Label_Brightness.Text = "Brightness"
        '
        'Pos_X
        '
        Me.Pos_X.BackColor = System.Drawing.Color.Gainsboro
        Me.Pos_X.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Pos_X.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Pos_X.ForeColor = System.Drawing.Color.Blue
        Me.Pos_X.Location = New System.Drawing.Point(64, 8)
        Me.Pos_X.Name = "Pos_X"
        Me.Pos_X.Size = New System.Drawing.Size(40, 16)
        Me.Pos_X.TabIndex = 49
        Me.Pos_X.Text = "1.000"
        '
        'Label_Pos_X
        '
        Me.Label_Pos_X.BackColor = System.Drawing.Color.Gainsboro
        Me.Label_Pos_X.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label_Pos_X.ForeColor = System.Drawing.Color.Black
        Me.Label_Pos_X.Location = New System.Drawing.Point(48, 8)
        Me.Label_Pos_X.Name = "Label_Pos_X"
        Me.Label_Pos_X.Size = New System.Drawing.Size(16, 16)
        Me.Label_Pos_X.TabIndex = 52
        Me.Label_Pos_X.Text = "X"
        '
        'Pos_Z
        '
        Me.Pos_Z.BackColor = System.Drawing.Color.Gainsboro
        Me.Pos_Z.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Pos_Z.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Pos_Z.ForeColor = System.Drawing.Color.Blue
        Me.Pos_Z.Location = New System.Drawing.Point(192, 8)
        Me.Pos_Z.Name = "Pos_Z"
        Me.Pos_Z.Size = New System.Drawing.Size(40, 16)
        Me.Pos_Z.TabIndex = 51
        Me.Pos_Z.Text = "1.000"
        '
        'Pos_Y
        '
        Me.Pos_Y.BackColor = System.Drawing.Color.Gainsboro
        Me.Pos_Y.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Pos_Y.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Pos_Y.ForeColor = System.Drawing.Color.Blue
        Me.Pos_Y.Location = New System.Drawing.Point(128, 8)
        Me.Pos_Y.Name = "Pos_Y"
        Me.Pos_Y.Size = New System.Drawing.Size(40, 16)
        Me.Pos_Y.TabIndex = 50
        Me.Pos_Y.Text = "1.000"
        '
        'Label_Pos_Y
        '
        Me.Label_Pos_Y.BackColor = System.Drawing.Color.Gainsboro
        Me.Label_Pos_Y.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label_Pos_Y.ForeColor = System.Drawing.Color.Black
        Me.Label_Pos_Y.Location = New System.Drawing.Point(112, 8)
        Me.Label_Pos_Y.Name = "Label_Pos_Y"
        Me.Label_Pos_Y.Size = New System.Drawing.Size(16, 16)
        Me.Label_Pos_Y.TabIndex = 53
        Me.Label_Pos_Y.Text = "Y"
        '
        'Label_Pos_Z
        '
        Me.Label_Pos_Z.BackColor = System.Drawing.Color.Gainsboro
        Me.Label_Pos_Z.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label_Pos_Z.ForeColor = System.Drawing.Color.Black
        Me.Label_Pos_Z.Location = New System.Drawing.Point(176, 8)
        Me.Label_Pos_Z.Name = "Label_Pos_Z"
        Me.Label_Pos_Z.Size = New System.Drawing.Size(16, 16)
        Me.Label_Pos_Z.TabIndex = 54
        Me.Label_Pos_Z.Text = "Z"
        '
        'ValueBrightness
        '
        Me.ValueBrightness.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ValueBrightness.Location = New System.Drawing.Point(416, 8)
        Me.ValueBrightness.Maximum = New Decimal(New Integer() {255, 0, 0, 0})
        Me.ValueBrightness.Name = "ValueBrightness"
        Me.ValueBrightness.Size = New System.Drawing.Size(48, 20)
        Me.ValueBrightness.TabIndex = 55
        Me.ValueBrightness.Value = New Decimal(New Integer() {51, 0, 0, 0})
        '
        'AxDisplay1
        '
        Me.AxDisplay1.ContainingControl = Me
        Me.AxDisplay1.Enabled = True
        Me.AxDisplay1.Location = New System.Drawing.Point(0, 0)
        Me.AxDisplay1.Name = "AxDisplay1"
        Me.AxDisplay1.OcxState = CType(resources.GetObject("AxDisplay1.OcxState"), System.Windows.Forms.AxHost.State)
        Me.AxDisplay1.Size = New System.Drawing.Size(768, 576)
        Me.AxDisplay1.TabIndex = 222
        '
        'AxMGraphicContext1
        '
        Me.AxMGraphicContext1.ContainingControl = Me
        Me.AxMGraphicContext1.Enabled = True
        Me.AxMGraphicContext1.Location = New System.Drawing.Point(864, 0)
        Me.AxMGraphicContext1.Name = "AxMGraphicContext1"
        Me.AxMGraphicContext1.OcxState = CType(resources.GetObject("AxMGraphicContext1.OcxState"), System.Windows.Forms.AxHost.State)
        Me.AxMGraphicContext1.Size = New System.Drawing.Size(32, 32)
        Me.AxMGraphicContext1.TabIndex = 223
        '
        'Timer_Circle
        '
        Me.Timer_Circle.Interval = 10
        '
        'SaveFileDialog_Image
        '
        Me.SaveFileDialog_Image.DefaultExt = "bmp"
        '
        'Timer_Histo
        '
        Me.Timer_Histo.Interval = 1000
        '
        'FormVision
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(968, 702)
        Me.Controls.Add(Me.PanelVision)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.SizableToolWindow
        Me.Name = "FormVision"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "FormVision"
        Me.PanelVision.ResumeLayout(False)
        CType(Me.AxEBW8Image1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.AxImage11, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.AxImage10, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.AxPatternMatching3, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.AxImage6, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.AxImage7, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.AxDisplay2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.AxImageProcessing1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.AxMeasurement6, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.AxMeasurement5, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.AxMeasurement4, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.AxMeasurement3, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.AxMeasurement2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.AxMeasurement1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.AxMeasurement12, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.AxMeasurement11, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.AxMeasurement9, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.AxMeasurement10, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.AxMeasurement8, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.AxMeasurement7, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.AxImageProcessing5, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.AxImage9, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.AxImageProcessing4, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.AxImage8, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.AxImageProcessing3, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.AxGraphicContext5, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.AxImage2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.AxPatternMatching1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.AxImage5, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.AxGraphicContext3, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.AxPatternMatching2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.AxCalibration1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.AxBlobAnalysis2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.AxImage12, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.AxImageProcessing6, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.AxImage13, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.AxImageProcessing2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.AxImage17, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.AxImageProcessing9, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.AxImage16, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.AxImageProcessing7, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.AxGraphicContext2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.AxGraphicContext1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.AxImage3, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.AxSystem1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.AxECamera1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.AxEBoard1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TBReference.ResumeLayout(False)
        CType(Me.ValueContrast, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.ValueBrightness, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.AxDisplay1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.AxMGraphicContext1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Shared m_instance As FormVision

    Public Shared ReadOnly Property Instance() As FormVision
        Get
            If m_instance Is Nothing Then
                m_instance = New FormVision
                Dim x As Integer
                x = 1 + 1
            Else
                If m_instance.IsDisposed Then
                    m_instance = New FormVision
                End If
            End If
            Return m_instance
        End Get

    End Property

    'Please Place Functions and attributes to be called by other DLL modules here

#Region "Initialization and termination"

    Dim hideDisplay As Integer = 0
    Dim PicoloPointer As Long
    Dim ImagePointer_PICOLO As New IntPtr
    Dim ImagePointer_MIL As New IntPtr
    Dim intermidiate(768 * 576) As Byte
    Dim intermidiate1(768 * 576) As Byte


#Region "Short functions"
    Private Sub InitializeFiducialForm()
        FiducialMark_form.Show()
        FiducialMark_form.ActiveForm.Visible = False
    End Sub
    Private Sub InitializeQCForm()
        QC_form.Show()
        QC_form.Visible = False
    End Sub
    Private Sub InitializeRejectMarkForm()
        RejectPoint_form.Show()
        RejectPoint_form.Visible = False
    End Sub
    Private Sub InitializeChipEdgeForm()
        ChipEdgePoints_form.Show()
        ChipEdgePoints_form.Visible = False
    End Sub
#End Region

    Public Sub VisionInitialization()
        InitializeCameraSettings()
        PicoloPointer = AxEBW8Image1.GetImagePointer(0, 0)
        ImagePointer_PICOLO = IntPtr.op_Explicit(PicoloPointer) 'convert long to ptr
        ImagePointer_MIL = IntPtr.op_Explicit(AxImage3.HostAddress) 'convert integer to ptr
        AxDisplay1.Visible = True
        AxDisplay1.OverlayKeyColor = System.Drawing.Color.Black
        AxDisplay1.OverlayVisible = True
        AxDisplay1.LUTEnabled = False
        ClearDisplay()

        'Initialize Forms
        'InitializeFiducialForm()
        'InitializeHistogramForm()
        'InitializeQCForm()
        'InitializeChipEdgeForm()
        'InitializeRejectMarkForm()
        'InitializeNCForm()
        'DisplayIndicator()
    End Sub

    Sub DisplayIndicator() 'Cross and Square
        Try
            ClearDisplay()
            Dim red As System.UInt32
            red = Convert.ToUInt32((RGB(255, 0, 0)))
            Dim green As System.UInt32
            green = Convert.ToUInt32((RGB(0, 255, 0)))

            With AxGraphicContext2.DrawingRegion()
                .StartX() = (DisplayCenterXPosition) - (2000 / 22)
                .StartY() = (DisplayCenterYPosition) - (2000 / 22)
                .EndX() = (DisplayCenterXPosition) + (2000 / 22)
                .EndY() = (DisplayCenterYPosition) + (2000 / 22)
            End With
            AxGraphicContext2.Rectangle(False, 0)

            With AxGraphicContext2.DrawingRegion()
                .StartX() = (DisplayCenterXPosition) - (3000 / 22)
                .StartY() = (DisplayCenterYPosition) - (3000 / 22)
                .EndX() = (DisplayCenterXPosition) + (3000 / 22)
                .EndY() = (DisplayCenterYPosition) + (3000 / 22)
            End With
            AxGraphicContext2.Cross(0)
        Catch ex As Exception
            ExceptionDisplay(ex)
        End Try
    End Sub

    Private Sub Button_Close_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_Close.Click
        If Button_Close.Text = "Close" Then
            Close()
        ElseIf Button_Close.Text = "Grab" Then
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        If hideDisplay = 1 Then
            Me.ClientSize = New System.Drawing.Size(DisplayWidth, 616)
            hideDisplay = 0
        Else
            Me.ClientSize = New System.Drawing.Size(1280, 616)
            hideDisplay = 1
        End If
    End Sub

    Private Sub FormVision_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        AxECamera1.SetParamNm("ChannelState", "IDLE")
        AxECamera1.SetParamNm("GrabCount", -1)
        Timer_Histo.Stop()
        Timer_Circle.Stop()
        AxSystem1.Free()

    End Sub
#End Region
#Region "Global Variables"
    Dim Folder As String = "C:\IDS\SRC\DLL Export Device Vision\Fiducial\" 'Nothing '"C:\IDS\SRC\DLL Export Device Vision\Fiducial\"
#End Region
#Region "Display Panel"
    Dim PanelPositionX As Integer = 0
    Dim PanelPositionY As Integer = 0
    Dim PanelLocationX As Integer = 0
    Dim PanelLocationY As Integer = 0

    Private Sub AxDisplay1_MouseMoveEvent1(ByVal sender As System.Object, ByVal e As AxMatrox.ActiveMIL.IDisplayEvents_MouseMoveEvent) Handles AxDisplay1.MouseMoveEvent
        PanelPositionX = e.x
        PanelPositionY = e.y
        Pos_X.Text = Convert.ToString(PanelPositionX)
        Pos_Y.Text = Convert.ToString(PanelPositionY)
        If FiducialMark_form.PictureBox10.BorderStyle = BorderStyle.Fixed3D Then
            If FiducialMark_form.Button_Teach.Enabled = True Then
                'If FiducialMark_form.Button_Teach.Text = "Next" Then
                ModelRegion_cursor()
                'ElseIf FiducialMark_form.Button_Teach.Text = "Finish" Then
                SearchRegion_cursor()
                'End If
                If mousedownID <> 0 Then
                    ClearDisplay()
                    'If FiducialMark_form.Button_Teach.Text = "Next" Then

                    If mousedownID = 10 Or mousedownID = 20 Or mousedownID = 30 Or mousedownID = 40 Then
                        ModelRegionSetting(False)
                    Else
                        SearchRegionSetting(False)
                    End If
                    modelregionDrawing()
                    SearchRegionDrawing()
                    'ElseIf FiducialMark_form.Button_Teach.Text = "Finish" Then

                    'End If
                End If
            End If
        ElseIf FiducialMark_form.PictureBox1.BorderStyle = BorderStyle.Fixed3D Or _
        FiducialMark_form.PictureBox2.BorderStyle = BorderStyle.Fixed3D Or _
        FiducialMark_form.PictureBox3.BorderStyle = BorderStyle.Fixed3D Or _
        FiducialMark_form.PictureBox4.BorderStyle = BorderStyle.Fixed3D Or _
        FiducialMark_form.PictureBox5.BorderStyle = BorderStyle.Fixed3D Or _
        FiducialMark_form.PictureBox6.BorderStyle = BorderStyle.Fixed3D Or _
        FiducialMark_form.PictureBox7.BorderStyle = BorderStyle.Fixed3D Or _
        FiducialMark_form.PictureBox8.BorderStyle = BorderStyle.Fixed3D Or _
        FiducialMark_form.PictureBox9.BorderStyle = BorderStyle.Fixed3D Or _
        FiducialMark_form.PictureBox11.BorderStyle = BorderStyle.Fixed3D Or _
        FiducialMark_form.PictureBox12.BorderStyle = BorderStyle.Fixed3D Or _
        FiducialMark_form.PictureBox13.BorderStyle = BorderStyle.Fixed3D Or _
        FiducialMark_form.PictureBox14.BorderStyle = BorderStyle.Fixed3D Or _
        FiducialMark_form.PictureBox15.BorderStyle = BorderStyle.Fixed3D Or _
        FiducialMark_form.PictureBox16.BorderStyle = BorderStyle.Fixed3D Or _
        FiducialMark_form.PictureBox17.BorderStyle = BorderStyle.Fixed3D Or _
        FiducialMark_form.PictureBox18.BorderStyle = BorderStyle.Fixed3D Or _
        FiducialMark_form.PictureBox19.BorderStyle = BorderStyle.Fixed3D Then
            SearchRegion_cursor()
            SearchRegionDrawing()
            If mousedownID <> 0 Then
                ClearDisplay()
                SearchRegionSetting(True)
                SearchRegionDrawing()
            End If
        End If
        If RejectPoint_form.TabControl1.SelectedIndex = 0 And RejectPoint_flag = True Then
            If RejectPoint_form.Button_Next.Enabled = True Then
                ModelRegion_cursor()
                If mousedownID <> 0 Then
                    ClearDisplay()
                    ModelRegionSetting(True)
                    modelregionDrawing()
                End If
            Else

            End If
        ElseIf RejectPoint_form.TabControl1.SelectedIndex = 1 Then
            If RejectPoint_form.Button_Test.Enabled = True And RejectPoint_form.Button_Load.Enabled = False Then
                ModelRegion_cursor()
                If mousedownID <> 0 Then
                    ClearDisplay()
                    ModelRegionSetting(True)
                    modelregionDrawing()
                End If
            Else
            End If
        End If
        If QC_form.Visible = True Then
            ModelRegion_cursor()
            If mousedownID <> 0 Then
                ClearDisplay()
                ModelRegionSetting(True)
                modelregionDrawing()
            End If
        End If
    End Sub
    Private Sub AxDisplay1_ClickEvent1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AxDisplay1.ClickEvent
        If ChipEdgeDrawingEnabled() Then 'by mouse
            x = PanelPositionX
            y = PanelPositionY
            ChipEdgePoints(x, y)
        End If
        If Measurement_flag = True Then
            x = PanelPositionX
            y = PanelPositionY
            Distance(x, y)
        End If
        '======================
    End Sub
    Private Sub AxDisplay1_MouseDownEvent1(ByVal sender As System.Object, ByVal e As AxMatrox.ActiveMIL.IDisplayEvents_MouseDownEvent) Handles AxDisplay1.MouseDownEvent
        If FiducialMark_form.Button_Teach.Enabled = True Then
            'If FiducialMark_form.Button_Teach.Text = "Next" Then
            MouseClickX = e.x
            MouseClickY = e.y
            ModelRegion()
            Drag_Flag = 0
            'ElseIf FiducialMark_form.Button_Teach.Text = "Finish" Then
            MouseClickX = e.x
            MouseClickY = e.y
            SearchRegion()
            Drag_Flag = 0
            'End If
        End If
        If FiducialMark_form.PictureBox1.BorderStyle = BorderStyle.Fixed3D Or _
        FiducialMark_form.PictureBox2.BorderStyle = BorderStyle.Fixed3D Or _
        FiducialMark_form.PictureBox3.BorderStyle = BorderStyle.Fixed3D Or _
        FiducialMark_form.PictureBox4.BorderStyle = BorderStyle.Fixed3D Or _
        FiducialMark_form.PictureBox5.BorderStyle = BorderStyle.Fixed3D Or _
        FiducialMark_form.PictureBox6.BorderStyle = BorderStyle.Fixed3D Or _
        FiducialMark_form.PictureBox7.BorderStyle = BorderStyle.Fixed3D Or _
        FiducialMark_form.PictureBox8.BorderStyle = BorderStyle.Fixed3D Or _
        FiducialMark_form.PictureBox9.BorderStyle = BorderStyle.Fixed3D Or _
        FiducialMark_form.PictureBox11.BorderStyle = BorderStyle.Fixed3D Or _
        FiducialMark_form.PictureBox12.BorderStyle = BorderStyle.Fixed3D Or _
        FiducialMark_form.PictureBox13.BorderStyle = BorderStyle.Fixed3D Or _
        FiducialMark_form.PictureBox14.BorderStyle = BorderStyle.Fixed3D Or _
        FiducialMark_form.PictureBox15.BorderStyle = BorderStyle.Fixed3D Or _
        FiducialMark_form.PictureBox16.BorderStyle = BorderStyle.Fixed3D Or _
        FiducialMark_form.PictureBox17.BorderStyle = BorderStyle.Fixed3D Or _
        FiducialMark_form.PictureBox18.BorderStyle = BorderStyle.Fixed3D Or _
        FiducialMark_form.PictureBox19.BorderStyle = BorderStyle.Fixed3D Then
            MouseClickX = e.x
            MouseClickY = e.y
            SearchRegion()
            Drag_Flag = 0
        End If
        If RejectPoint_form.TabControl1.SelectedIndex = 0 Then
            If RejectPoint_form.Button_Next.Enabled = False Then
            Else
                MouseClickX = e.x
                MouseClickY = e.y
                ModelRegion()
                Drag_Flag = 0
            End If
        ElseIf RejectPoint_form.TabControl1.SelectedIndex = 1 Then
            If RejectPoint_form.Button_Test.Enabled = True And RejectPoint_form.Button_Load.Enabled = False Then
                MouseClickX = e.x
                MouseClickY = e.y
                ModelRegion()
                Drag_Flag = 0
            Else
            End If
        End If
        If QC_form.Visible = True Then
            MouseClickX = e.x
            MouseClickY = e.y
            ModelRegion()
            Drag_Flag = 0
        End If
        'If SS_form.Btn_Start.Enabled = True And SS_form.Visible = True Then
        'MouseClickX = e.x
        'MouseClickY = e.y
        'ModelRegion()
        'Drag_Flag = 0
        'End If
    End Sub
    Private Sub AxDisplay1_MouseUpEvent1(ByVal sender As System.Object, ByVal e As AxMatrox.ActiveMIL.IDisplayEvents_MouseUpEvent) Handles AxDisplay1.MouseUpEvent
        If QC_form.Visible = True Then
            QC_form.GetVariables(DisplayCenterXPosition, DisplayCenterYPosition, MROIx, MROIy, 0, 0, 0)
        End If

        mousedownID = 0
    End Sub
#End Region
#Region "Brightness and Contrass"
    Private Sub ValueBrightness_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ValueBrightness.ValueChanged

        Dim a As Byte = (ValueBrightness.Value) 'Xu Long 20/02/2012
        If ValueBrightness.Focused = True Then
            _LightBox.SetBrightness(a)
        End If

        'If Initializing Then Return

        'Dim a As Byte = (ValueBrightness.Value)
        '_LightBox.SetBrightness(a)
    End Sub
    Private Sub ValueContrast_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ValueContrast.ValueChanged
        'AxDigitizer1.Contrast = ValueContrast.Value
    End Sub
    Function Brightness()
        Dim a As Byte = (ValueBrightness.Value)
        _LightBox.SetBrightness(a)
    End Function
    Function SetBrightness(ByVal SetBrightnessValue As Integer)
        'Dim a As Byte = (ValueBrightness.Value)
        _LightBox.SetBrightness(SetBrightnessValue)
    End Function
    Function SetLaser(ByVal a As Boolean) 'kr dec4
        'Dim a As Byte = (ValueBrightness.Value)
        _LightBox.SetLaser(a)
    End Function
#End Region
#Region "SaveImage"
    Sub SaveImage()
        If SaveFileDialog_Image.ShowDialog() = DialogResult.OK Then
            Dim extension As String = SaveFileDialog_Image.FileName
            extension = extension.Substring(extension.LastIndexOf(".") + 1).ToLower
            Select Case extension
                Case "bmp"
                    'AxImageProcessing1.Binarize(MILIMAGEPROCESSINGLib.ImpConditionConstants.impLessThan, 64)
                    'AxImage2 = AxImageProcessing1.Destination1()
                    AxImage3.Save(SaveFileDialog_Image.FileName)
                    'AxDisplay1.OverlayImage.Save(SaveFileDialog_Image.FileName)
                    'AxDisplay1.OverlayImage.Save(SaveFileDialog_Image.FileName, AxImage1.FileFormat.imBMP)
                Case "jpg", "jpeg"
                    AxImage3.Save(SaveFileDialog_Image.FileName)
                    'PanelVision.BackgroundImage.Save(SaveFileDialog_Image.FileName, Imaging.ImageFormat.Jpeg)
                Case "gif"
                    AxImage3.Save(SaveFileDialog_Image.FileName)
                    'PanelVision.BackgroundImage.Save(SaveFileDialog_Image.FileName, Imaging.ImageFormat.Gif)
                Case "ico"
                    AxImage3.Save(SaveFileDialog_Image.FileName)
                    'PanelVision.BackgroundImage.Save(SaveFileDialog_Image.FileName, Imaging.ImageFormat.Icon)
                Case "emf"
                    AxImage3.Save(SaveFileDialog_Image.FileName)
                    'PanelVision.BackgroundImage.Save(SaveFileDialog_Image.FileName, Imaging.ImageFormat.Emf)
                Case "wmf"
                    AxImage3.Save(SaveFileDialog_Image.FileName)
                    'PanelVision.BackgroundImage.Save(SaveFileDialog_Image.FileName, Imaging.ImageFormat.Wmf)
                Case "png"
                    AxImage3.Save(SaveFileDialog_Image.FileName)
                    'PanelVision.BackgroundImage.Save(SaveFileDialog_Image.FileName, Imaging.ImageFormat.Png)
                Case "tif", "tiff"
                    AxImage3.Save(SaveFileDialog_Image.FileName)
                    'PanelVision.BackgroundImage.Save(SaveFileDialog_Image.FileName, Imaging.ImageFormat.Tiff)
                Case "exif"
                    AxImage3.Save(SaveFileDialog_Image.FileName)
                    'PanelVision.BackgroundImage.Save(SaveFileDialog_Image.FileName, Imaging.ImageFormat.Exif)
            End Select
            'MsgBox("Image saved !!")
        End If
    End Sub
#End Region
#Region "Fiducial"
#Region "Fiducial-RunTime-PM"
    Shared ROIx As Decimal = 300 'Search ROI
    Shared ROIy As Decimal = 300
    Shared RegionX As Decimal = DisplayCenterXPosition
    Shared RegionY As Decimal = DisplayCenterYPosition
    Shared MROIx As Decimal = 100 'Model ROI
    Shared MROIy As Decimal = 100
    Shared DisplayCenterXPosition As Decimal = 768 / 2
    Shared DisplayCenterYPosition As Decimal = 576 / 2
    Shared IndexNo As Decimal = 0 'Model Index No
    Shared ModelNo As Decimal = 0
    Dim PM_indicator As Integer = 0
    Dim ClickPM As Integer = 0
    Dim MouseClickX, MouseClickY As Integer
    Shared mousedownID As Integer = 0
    Dim ClickRegion As Integer = 5
    Shared NewXX, NewYY, OldXX, OldYY As Decimal
    Shared Drag_Flag As Integer = 0
    Shared New_XX, New_YY, Old_XX, Old_YY As Decimal
    Sub CustomizeFiducial_PM(ByVal Reset As Boolean)
        Try
            If Reset = True Then
                ClickPM = 0
            Else
                If ClickPM = 0 Then 'First click
                    If IndexNo > 0 Or AxPatternMatching2.Models.Count > 0 Then
                        IndexNo = 0
                        If AxPatternMatching2.Models.Count <> 0 Then
                            For i As Integer = AxPatternMatching2.Models.Count To 1 Step -1
                                AxPatternMatching2.Models.Remove(i)
                            Next
                        End If
                    End If
                    'AxPatternMatching2.Free()
                    'AxPatternMatching2.Allocate()
                    'AxImageProcessing1.Binarize(MILIMAGEPROCESSINGLib.ImpConditionConstants.impGreaterThan, 128)
                    'Button_Teach.Text = "Select the SearchRegion"
                    IndexNo = AxPatternMatching2.Models.Add(AxDisplay1.Image, DisplayCenterXPosition - MROIx / 2, DisplayCenterYPosition - MROIy / 2, MROIx, MROIy, 1, Matrox.ActiveMIL.PatternMatching.PatModelTypeConstants.patNormalized)
                    ModelSaveImage()
                    ClickPM = 1
                ElseIf ClickPM > 0 Then
                    AxPatternMatching2.Models.Item(IndexNo).PositionAccuracy = Matrox.ActiveMIL.PatternMatching.PatHighLowConstants.patVeryHigh
                    AxPatternMatching2.Models.Item(IndexNo).AcceptanceThreshold = 70
                    AxPatternMatching2.Models.Item(IndexNo).CertaintyThreshold = 95
                    AxPatternMatching2.Models.Item(IndexNo).NumberOfOccurrences = 1
                    AxPatternMatching2.Models.Item(IndexNo).Speed = Matrox.ActiveMIL.PatternMatching.PatHighLowConstants.patMedium

                    AxPatternMatching2.Models.Item(IndexNo).SearchAngle.Enabled = False
                    AxPatternMatching2.Models.Item(IndexNo).SearchAngle.Value = 0
                    'AxPatternMatching2.Models.Item(IndexNo).SearchAngle.PositiveDelta = 10
                    'AxPatternMatching2.Models.Item(IndexNo).SearchAngle.NegativeDelta = 10
                    'AxPatternMatching2.Models.Item(IndexNo).SearchAngle.Accuracy = 0.1
                    'AxPatternMatching2.Models.Item(IndexNo).SearchAngle.Tolerance = 0.1

                    'AxPatternMatching2.Models.Item(IndexNo).SearchRegion.FullSize() = True
                    AxPatternMatching2.Models.Item(IndexNo).SearchRegion.StartX = RegionX - ROIx / 2 'Panels actual position== refer to display's origin
                    AxPatternMatching2.Models.Item(IndexNo).SearchRegion.StartY = RegionY - ROIy / 2 'Panels actual position== refer to display's origin
                    AxPatternMatching2.Models.Item(IndexNo).SearchRegion.SizeX = ROIx
                    AxPatternMatching2.Models.Item(IndexNo).SearchRegion.SizeY = ROIy

                    AxPatternMatching2.Models.Item(IndexNo).ReferenceX = MROIx / 2
                    AxPatternMatching2.Models.Item(IndexNo).ReferenceY = MROIy / 2

                    AxPatternMatching2.Models.Item(IndexNo).Preprocess()

                    ClearDisplay()
                    DisplayIndicator()
                    PM_indicator = 1
                    ClickPM = 0
                End If
            End If
        Catch ex As Exception
            ExceptionDisplay(ex)
        End Try
    End Sub 'called by MainGUI
    Sub ModelRegion()
        'MousePosition at ROI four corner
        If (((DisplayCenterXPosition - MROIx / 2) - ClickRegion) < PanelPositionX) And _
        (((DisplayCenterYPosition - MROIy / 2) - ClickRegion) < PanelPositionY) And _
        (((DisplayCenterXPosition - MROIx / 2) + ClickRegion) > PanelPositionX) And _
        (((DisplayCenterYPosition - MROIy / 2) + ClickRegion) > PanelPositionY) Then
            Cursor.Current = System.Windows.Forms.Cursors.Hand()
            mousedownID = 10
        ElseIf (((DisplayCenterXPosition + MROIx / 2) - ClickRegion) < PanelPositionX) And _
        (((DisplayCenterYPosition + MROIy / 2) - ClickRegion) < PanelPositionY) And _
        (((DisplayCenterXPosition + MROIx / 2) + ClickRegion) > PanelPositionX) And _
        (((DisplayCenterYPosition + MROIy / 2) + ClickRegion) > PanelPositionY) Then
            Cursor.Current = System.Windows.Forms.Cursors.Hand()
            mousedownID = 20
        ElseIf (((DisplayCenterXPosition - MROIx / 2) - ClickRegion) < PanelPositionX) And _
        (((DisplayCenterYPosition + MROIy / 2) - ClickRegion) < PanelPositionY) And _
        (((DisplayCenterXPosition - MROIx / 2) + ClickRegion) > PanelPositionX) And _
        (((DisplayCenterYPosition + MROIy / 2) + ClickRegion) > PanelPositionY) Then
            Cursor.Current = System.Windows.Forms.Cursors.Hand()
            mousedownID = 30
        ElseIf (((DisplayCenterXPosition + MROIx / 2) - ClickRegion) < PanelPositionX) And _
        (((DisplayCenterYPosition - MROIy / 2) - ClickRegion) < PanelPositionY) And _
        (((DisplayCenterXPosition + MROIx / 2) + ClickRegion) > PanelPositionX) And _
        (((DisplayCenterYPosition - MROIy / 2) + ClickRegion) > PanelPositionY) Then
            Cursor.Current = System.Windows.Forms.Cursors.Hand()
            mousedownID = 40
        ElseIf ((DisplayCenterXPosition - 0.4 * MROIx) < (PanelPositionX) And _
                (PanelPositionX) < (DisplayCenterXPosition + 0.4 * MROIx)) And _
                ((DisplayCenterYPosition - 0.4 * MROIy) < (PanelPositionY) And _
                (PanelPositionY) < (DisplayCenterYPosition + 0.4 * MROIy)) Then
            'Cursor.Current = System.Windows.Forms.Cursors.NoMove2D
            'mousedownID = 5000
        End If
    End Sub
    Sub ModelRegion_cursor()
        'MousePosition at ROI four corner
        If (((DisplayCenterXPosition - MROIx / 2) - ClickRegion) < PanelPositionX) And _
            (((DisplayCenterYPosition - MROIy / 2) - ClickRegion) < PanelPositionY) And _
            (((DisplayCenterXPosition - MROIx / 2) + ClickRegion) > PanelPositionX) And _
            (((DisplayCenterYPosition - MROIy / 2) + ClickRegion) > PanelPositionY) Then
            Cursor.Current = System.Windows.Forms.Cursors.Hand()
        ElseIf (((DisplayCenterXPosition + MROIx / 2) - ClickRegion) < PanelPositionX) And _
            (((DisplayCenterYPosition + MROIy / 2) - ClickRegion) < PanelPositionY) And _
            (((DisplayCenterXPosition + MROIx / 2) + ClickRegion) > PanelPositionX) And _
            (((DisplayCenterYPosition + MROIy / 2) + ClickRegion) > PanelPositionY) Then
            Cursor.Current = System.Windows.Forms.Cursors.Hand()
        ElseIf (((DisplayCenterXPosition - MROIx / 2) - ClickRegion) < PanelPositionX) And _
            (((DisplayCenterYPosition + MROIy / 2) - ClickRegion) < PanelPositionY) And _
            (((DisplayCenterXPosition - MROIx / 2) + ClickRegion) > PanelPositionX) And _
            (((DisplayCenterYPosition + MROIy / 2) + ClickRegion) > PanelPositionY) Then
            Cursor.Current = System.Windows.Forms.Cursors.Hand()
        ElseIf (((DisplayCenterXPosition + MROIx / 2) - ClickRegion) < PanelPositionX) And _
            (((DisplayCenterYPosition - MROIy / 2) - ClickRegion) < PanelPositionY) And _
            (((DisplayCenterXPosition + MROIx / 2) + ClickRegion) > PanelPositionX) And _
            (((DisplayCenterYPosition - MROIy / 2) + ClickRegion) > PanelPositionY) Then
            Cursor.Current = System.Windows.Forms.Cursors.Hand()
        ElseIf ((DisplayCenterXPosition - 0.4 * MROIx) < (PanelPositionX) And _
                (PanelPositionX) < (DisplayCenterXPosition + 0.4 * MROIx)) And _
                ((DisplayCenterYPosition - 0.4 * MROIy) < (PanelPositionY) And _
                (PanelPositionY) < (DisplayCenterYPosition + 0.4 * MROIy)) Then
            'Cursor.Current = System.Windows.Forms.Cursors.NoMove2D
        End If
    End Sub
    Sub modelregionDrawing()
        'ClearDisplay
        'DisplayIndicator()
        Try

            If PM_indicator = 1 Then 'To remove model after it clicked the 2nd time of fiducial button
                If IndexNo = 0 Then
                Else
                    AxPatternMatching2.Models.Remove(IndexNo)
                    IndexNo = 0
                    MROIx = 100 'Model ROI
                    MROIy = 100
                    DisplayCenterXPosition = 768 / 2
                    DisplayCenterYPosition = 576 / 2
                End If

            End If

            'Display model region
            With AxGraphicContext1.DrawingRegion
                .EndX = DisplayCenterXPosition + MROIx / 2
                .EndY = DisplayCenterYPosition + MROIy / 2
                .StartX = DisplayCenterXPosition - MROIx / 2
                .StartY = DisplayCenterYPosition - MROIy / 2
            End With
            AxGraphicContext1.Rectangle()
            With AxGraphicContext1.DrawingRegion
                .EndX = DisplayCenterXPosition - MROIx / 2 + ClickRegion
                .EndY = DisplayCenterYPosition - MROIy / 2 + ClickRegion
                .StartX = DisplayCenterXPosition - MROIx / 2 - ClickRegion
                .StartY = DisplayCenterYPosition - MROIy / 2 - ClickRegion
            End With
            AxGraphicContext1.Rectangle(True, 0)
            With AxGraphicContext1.DrawingRegion
                .EndX = DisplayCenterXPosition + MROIx / 2 + ClickRegion
                .EndY = DisplayCenterYPosition + MROIy / 2 + ClickRegion
                .StartX = DisplayCenterXPosition + MROIx / 2 - ClickRegion
                .StartY = DisplayCenterYPosition + MROIy / 2 - ClickRegion
            End With
            AxGraphicContext1.Rectangle(True, 0)
            With AxGraphicContext1.DrawingRegion
                .EndX = DisplayCenterXPosition - MROIx / 2 + ClickRegion
                .EndY = DisplayCenterYPosition + MROIy / 2 + ClickRegion
                .StartX = DisplayCenterXPosition - MROIx / 2 - ClickRegion
                .StartY = DisplayCenterYPosition + MROIy / 2 - ClickRegion
            End With
            AxGraphicContext1.Rectangle(True, 0)
            With AxGraphicContext1.DrawingRegion
                .EndX = DisplayCenterXPosition + MROIx / 2 + ClickRegion
                .EndY = DisplayCenterYPosition - MROIy / 2 + ClickRegion
                .StartX = DisplayCenterXPosition + MROIx / 2 - ClickRegion
                .StartY = DisplayCenterYPosition - MROIy / 2 - ClickRegion
            End With
            AxGraphicContext1.Rectangle(True, 0)
            With AxGraphicContext1.DrawingRegion
                .EndX = DisplayCenterXPosition + 10
                .EndY = DisplayCenterYPosition + 10
                .StartX = DisplayCenterXPosition - 10
                .StartY = DisplayCenterYPosition - 10
            End With
            AxGraphicContext1.Cross()
        Catch ex As Exception
            ExceptionDisplay(ex)
        End Try
    End Sub
    Sub ModelRegionSetting(ByVal SkipSearchRegion As Boolean)
        Dim OffsetX As Decimal
        Dim OffsetY As Decimal
        Dim Old_x, Old_y As Decimal
        Dim New_x, New_y As Decimal
        '        modelregionDrawing()
        If mousedownID = 10 Then
            If PanelPositionX > DisplayCenterXPosition - 10 And PanelPositionY > DisplayCenterYPosition - 10 Then
                PanelPositionX = DisplayCenterXPosition - 10
                PanelPositionY = DisplayCenterYPosition - 10
                'Windows.Forms.Cursor.Current.Position = New System.Drawing.Point(PanelPositionX, PanelPositionY)
            ElseIf PanelPositionX > DisplayCenterXPosition - 10 And PanelPositionY < DisplayCenterYPosition - 10 Then
                PanelPositionX = DisplayCenterXPosition - 10
                PanelPositionY = PanelPositionY
                'Windows.Forms.Cursor.Current.Position = New System.Drawing.Point(PanelPositionX, PanelPositionY)
            ElseIf PanelPositionX < DisplayCenterXPosition - 10 And PanelPositionY > DisplayCenterYPosition - 10 Then
                PanelPositionX = PanelPositionX
                PanelPositionY = DisplayCenterYPosition - 10
                'Windows.Forms.Cursor.Current.Position = New System.Drawing.Point(PanelPositionX, PanelPositionY)
            End If
            If SkipSearchRegion = False Then
                If PanelPositionX < (DisplayCenterXPosition - ROIx / 2) And PanelPositionY < DisplayCenterYPosition - ROIy / 2 Then
                    PanelPositionX = DisplayCenterXPosition - ROIx / 2 + 3
                    PanelPositionY = DisplayCenterYPosition - ROIy / 2 + 3
                    'Windows.Forms.Cursor.Current.Position = New System.Drawing.Point(PanelPositionX, PanelPositionY)
                ElseIf PanelPositionX < DisplayCenterXPosition - ROIx / 2 And PanelPositionY > DisplayCenterYPosition - ROIy / 2 Then
                    PanelPositionX = DisplayCenterXPosition - ROIx / 2 + 3
                    'PanelPositionY =DisplayCenterYPosition - ROIy / 2
                ElseIf PanelPositionX > DisplayCenterXPosition - ROIx / 2 And PanelPositionY < DisplayCenterYPosition - ROIy / 2 Then
                    'PanelPositionX = DisplayCenterXPosition - ROIx / 2
                    PanelPositionY = DisplayCenterYPosition - ROIy / 2 + 3
                End If
            End If

            Old_x = DisplayCenterXPosition - MROIx / 2
            Old_y = DisplayCenterYPosition - MROIy / 2
            New_x = PanelPositionX
            New_y = PanelPositionY
            OffsetX = New_x - Old_x
            OffsetY = New_y - Old_y
            MROIx = (MROIx - OffsetX)
            MROIy = (MROIy - OffsetY)
            DisplayCenterXPosition = 768 / 2 'New_x + 0.5 * MROIx
            DisplayCenterYPosition = 576 / 2 'New_y + 0.5 * MROIy
        ElseIf mousedownID = 20 Then
            If PanelPositionX < DisplayCenterXPosition + 10 And PanelPositionY < DisplayCenterYPosition + 10 Then
                PanelPositionX = DisplayCenterXPosition + 10
                PanelPositionY = DisplayCenterYPosition + 10
                'Windows.Forms.Cursor.Current.Position = New System.Drawing.Point(PanelPositionX, PanelPositionY)
            ElseIf PanelPositionX < DisplayCenterXPosition + 10 And PanelPositionY > DisplayCenterYPosition + 10 Then
                PanelPositionX = DisplayCenterXPosition + 10
                'PanelPositionY = DisplayCenterYPosition + 10
                'Windows.Forms.Cursor.Current.Position = New System.Drawing.Point(PanelPositionX, PanelPositionY)
            ElseIf PanelPositionX > DisplayCenterXPosition + 10 And PanelPositionY < DisplayCenterYPosition + 10 Then
                'PanelPositionX = DisplayCenterXPosition + 10
                PanelPositionY = DisplayCenterYPosition + 10
                'Windows.Forms.Cursor.Current.Position = New System.Drawing.Point(PanelPositionX, PanelPositionY)
            End If
            If SkipSearchRegion = False Then
                If PanelPositionX > (DisplayCenterXPosition + ROIx / 2) And PanelPositionY > DisplayCenterYPosition + ROIy / 2 Then
                    PanelPositionX = DisplayCenterXPosition + ROIx / 2 - 3
                    PanelPositionY = DisplayCenterYPosition + ROIy / 2 - 3
                    'Windows.Forms.Cursor.Current.Position = New System.Drawing.Point(PanelPositionX, PanelPositionY)
                ElseIf PanelPositionX > DisplayCenterXPosition + ROIx / 2 And PanelPositionY < DisplayCenterYPosition + ROIy / 2 Then
                    PanelPositionX = DisplayCenterXPosition + ROIx / 2 - 3
                    'PanelPositionY =DisplayCenterYPosition + ROIy / 2
                ElseIf PanelPositionX < DisplayCenterXPosition + ROIx / 2 And PanelPositionY > DisplayCenterYPosition + ROIy / 2 Then
                    'PanelPositionX = DisplayCenterXPosition + ROIx / 2
                    PanelPositionY = DisplayCenterYPosition + ROIy / 2 - 3
                End If
            End If
            Old_x = DisplayCenterXPosition + MROIx / 2
            Old_y = DisplayCenterYPosition + MROIy / 2
            New_x = PanelPositionX
            New_y = PanelPositionY
            OffsetX = New_x - Old_x
            OffsetY = New_y - Old_y
            MROIx = (MROIx + OffsetX)
            MROIy = (MROIy + OffsetY)
            DisplayCenterXPosition = 768 / 2 'New_x - 0.5 * MROIx
            DisplayCenterYPosition = 576 / 2 'New_y - 0.5 * MROIy
        ElseIf mousedownID = 30 Then
            If PanelPositionX > DisplayCenterXPosition - 10 And PanelPositionY < DisplayCenterYPosition + 10 Then
                PanelPositionX = DisplayCenterXPosition - 10
                PanelPositionY = DisplayCenterYPosition + 10
                'Windows.Forms.Cursor.Current.Position = New System.Drawing.Point(PanelPositionX, PanelPositionY)
            ElseIf PanelPositionX > DisplayCenterXPosition - 10 And PanelPositionY > DisplayCenterYPosition + 10 Then
                PanelPositionX = DisplayCenterXPosition - 10
                'PanelPositionY = DisplayCenterYPosition+ 10
                'Windows.Forms.Cursor.Current.Position = New System.Drawing.Point(PanelPositionX, PanelPositionY)
            ElseIf PanelPositionX < DisplayCenterXPosition - 10 And PanelPositionY < DisplayCenterYPosition + 10 Then
                'PanelPositionX = DisplayCenterXPosition - 10
                PanelPositionY = DisplayCenterYPosition + 10
                'Windows.Forms.Cursor.Current.Position = New System.Drawing.Point(PanelPositionX, PanelPositionY)
            End If
            If SkipSearchRegion = False Then
                If PanelPositionX < (DisplayCenterXPosition - ROIx / 2) And PanelPositionY > DisplayCenterYPosition + ROIy / 2 Then
                    PanelPositionX = DisplayCenterXPosition - ROIx / 2 + 3
                    PanelPositionY = DisplayCenterYPosition + ROIy / 2 - 3
                    'Windows.Forms.Cursor.Current.Position = New System.Drawing.Point(PanelPositionX, PanelPositionY)
                ElseIf PanelPositionX < DisplayCenterXPosition - ROIx / 2 And PanelPositionY < DisplayCenterYPosition + ROIy / 2 Then
                    PanelPositionX = DisplayCenterXPosition - ROIx / 2 + 3
                    'PanelPositionY =DisplayCenterYPosition + ROIy / 2
                ElseIf PanelPositionX > DisplayCenterXPosition - ROIx / 2 And PanelPositionY > DisplayCenterYPosition + ROIy / 2 Then
                    'PanelPositionX = DisplayCenterXPosition - ROIx / 2
                    PanelPositionY = DisplayCenterYPosition + ROIy / 2 - 3
                End If
            End If
            Old_x = DisplayCenterXPosition - MROIx / 2
            Old_y = DisplayCenterYPosition + MROIy / 2
            New_x = PanelPositionX
            New_y = PanelPositionY
            OffsetX = New_x - Old_x
            OffsetY = New_y - Old_y
            MROIx = (MROIx - OffsetX)
            MROIy = (MROIy + OffsetY)
            DisplayCenterXPosition = 768 / 2 'New_x + 0.5 * MROIx
            DisplayCenterYPosition = 576 / 2 'New_y - 0.5 * MROIy
        ElseIf mousedownID = 40 Then
            If PanelPositionX < DisplayCenterXPosition + 10 And PanelPositionY > DisplayCenterYPosition - 10 Then
                PanelPositionX = DisplayCenterXPosition + 10
                PanelPositionY = DisplayCenterYPosition + 10
                'Windows.Forms.Cursor.Current.Position = New System.Drawing.Point(PanelPositionX, PanelPositionY)
            ElseIf PanelPositionX < DisplayCenterXPosition + 10 And PanelPositionY < DisplayCenterYPosition - 10 Then
                PanelPositionX = DisplayCenterXPosition + 10
                'PanelPositionY = DisplayCenterYPosition - 10
                'Windows.Forms.Cursor.Current.Position = New System.Drawing.Point(PanelPositionX, PanelPositionY)
            ElseIf PanelPositionX > DisplayCenterXPosition + 10 And PanelPositionY > DisplayCenterYPosition - 10 Then
                'PanelPositionX = DisplayCenterXPosition+10 
                PanelPositionY = DisplayCenterYPosition - 10
                'Windows.Forms.Cursor.Current.Position = New System.Drawing.Point(PanelPositionX, PanelPositionY)
            End If
            If SkipSearchRegion = False Then
                If PanelPositionX > (DisplayCenterXPosition + ROIx / 2) And PanelPositionY < DisplayCenterYPosition - ROIy / 2 Then
                    PanelPositionX = DisplayCenterXPosition + ROIx / 2 - 3
                    PanelPositionY = DisplayCenterYPosition - ROIy / 2 + 3
                    'Windows.Forms.Cursor.Current.Position = New System.Drawing.Point(PanelPositionX, PanelPositionY)
                ElseIf PanelPositionX > DisplayCenterXPosition + ROIx / 2 And PanelPositionY > DisplayCenterYPosition - ROIy / 2 Then
                    PanelPositionX = DisplayCenterXPosition + ROIx / 2 - 3
                    'PanelPositionY =DisplayCenterYPosition - ROIy / 2
                ElseIf PanelPositionX < DisplayCenterXPosition + ROIx / 2 And PanelPositionY < DisplayCenterYPosition - ROIy / 2 Then
                    'PanelPositionX = DisplayCenterXPosition + ROIx / 2
                    PanelPositionY = DisplayCenterYPosition - ROIy / 2 + 3
                End If
            End If
            Old_x = DisplayCenterXPosition + MROIx / 2
            Old_y = DisplayCenterYPosition - MROIy / 2
            New_x = PanelPositionX
            New_y = PanelPositionY
            OffsetX = New_x - Old_x
            OffsetY = New_y - Old_y
            MROIx = (MROIx + OffsetX)
            MROIy = (MROIy - OffsetY)
            DisplayCenterXPosition = 768 / 2 'New_x - 0.5 * MROIx
            DisplayCenterYPosition = 576 / 2 'New_y + 0.5 * MROIy
        ElseIf mousedownID = 50 Then
            If Drag_Flag = 0 Then
                Old_XX = MouseClickX 'MousePosition.X() - Form1.ActiveForm.Location.X - FiducialMark_form.Panel2.Location.X - FiducialMark_form.GroupBox5.Location.X 'PanelPositionX 'RegionX
                Old_YY = MouseClickY 'MousePosition.Y() - Form1.ActiveForm.Location.Y - FiducialMark_form.Panel2.Location.Y - FiducialMark_form.GroupBox5.Location.Y 'PanelPositionY 'RegionY
                Drag_Flag = 1
            End If
            'Sleep(100)
            New_XX = PanelPositionX '- Form1.ActiveForm.Location.X
            New_YY = PanelPositionY '- Form1.ActiveForm.Location.Y
            OffsetX = (New_XX - Old_XX)
            OffsetY = (New_YY - Old_YY)
            'kr?
            DisplayCenterXPosition = 768 / 2 + OffsetX
            DisplayCenterYPosition = 576 / 2 + OffsetY
            Old_XX = New_XX
            Old_YY = New_YY
        End If
        If RejectPoint_form.TabControl1.SelectedIndex = 0 Then
            If RejectPoint_form.Button_Test.Enabled = True And RejectPoint_form.Button_Load.Enabled = False Then
                RejectPoint_form.GetVariables(DisplayCenterXPosition, DisplayCenterYPosition, MROIx, MROIy)
            Else
            End If
        End If
    End Sub
    Sub SearchRegion()
        'SearchRegionDrawing()
        'MousePosition at ROI four corner
        If (((RegionX - ROIx / 2) - ClickRegion) < PanelPositionX) And _
            (((RegionY - ROIy / 2) - ClickRegion) < PanelPositionY) And _
            (((RegionX - ROIx / 2) + ClickRegion) > PanelPositionX) And _
            (((RegionY - ROIy / 2) + ClickRegion) > PanelPositionY) Then
            Cursor.Current = System.Windows.Forms.Cursors.Hand()
            mousedownID = 1
        ElseIf (((RegionX + ROIx / 2) - ClickRegion) < PanelPositionX) And _
            (((RegionY + ROIy / 2) - ClickRegion) < PanelPositionY) And _
            (((RegionX + ROIx / 2) + ClickRegion) > PanelPositionX) And _
            (((RegionY + ROIy / 2) + ClickRegion) > PanelPositionY) Then
            Cursor.Current = System.Windows.Forms.Cursors.Hand()
            mousedownID = 2
        ElseIf (((RegionX - ROIx / 2) - ClickRegion) < PanelPositionX) And _
            (((RegionY + ROIy / 2) - ClickRegion) < PanelPositionY) And _
            (((RegionX - ROIx / 2) + ClickRegion) > PanelPositionX) And _
            (((RegionY + ROIy / 2) + ClickRegion) > PanelPositionY) Then
            Cursor.Current = System.Windows.Forms.Cursors.Hand()
            mousedownID = 3
        ElseIf (((RegionX + ROIx / 2) - ClickRegion) < PanelPositionX) And _
            (((RegionY - ROIy / 2) - ClickRegion) < PanelPositionY) And _
            (((RegionX + ROIx / 2) + ClickRegion) > PanelPositionX) And _
            (((RegionY - ROIy / 2) + ClickRegion) > PanelPositionY) Then
            Cursor.Current = System.Windows.Forms.Cursors.Hand()
            mousedownID = 4
        ElseIf ((RegionX - 0.4 * ROIx) < (PanelPositionX) And _
                (PanelPositionX) < (RegionX + 0.4 * ROIx)) And _
                ((RegionY - 0.4 * ROIy) < (PanelPositionY) And _
                (PanelPositionY) < (RegionY + 0.4 * ROIy)) Then
            'Cursor.Current = System.Windows.Forms.Cursors.NoMove2D
            'mousedownID = 5000
        Else
        End If
    End Sub
    Sub SearchRegion_cursor()
        'MousePosition at ROI four corner
        If (((RegionX - ROIx / 2) - ClickRegion) < PanelPositionX) And _
            (((RegionY - ROIy / 2) - ClickRegion) < PanelPositionY) And _
            (((RegionX - ROIx / 2) + ClickRegion) > PanelPositionX) And _
            (((RegionY - ROIy / 2) + ClickRegion) > PanelPositionY) Then
            Cursor.Current = System.Windows.Forms.Cursors.Hand()
        ElseIf (((RegionX + ROIx / 2) - ClickRegion) < PanelPositionX) And _
            (((RegionY + ROIy / 2) - ClickRegion) < PanelPositionY) And _
            (((RegionX + ROIx / 2) + ClickRegion) > PanelPositionX) And _
            (((RegionY + ROIy / 2) + ClickRegion) > PanelPositionY) Then
            Cursor.Current = System.Windows.Forms.Cursors.Hand()
        ElseIf (((RegionX - ROIx / 2) - ClickRegion) < PanelPositionX) And _
            (((RegionY + ROIy / 2) - ClickRegion) < PanelPositionY) And _
            (((RegionX - ROIx / 2) + ClickRegion) > PanelPositionX) And _
            (((RegionY + ROIy / 2) + ClickRegion) > PanelPositionY) Then
            Cursor.Current = System.Windows.Forms.Cursors.Hand()
        ElseIf (((RegionX + ROIx / 2) - ClickRegion) < PanelPositionX) And _
            (((RegionY - ROIy / 2) - ClickRegion) < PanelPositionY) And _
            (((RegionX + ROIx / 2) + ClickRegion) > PanelPositionX) And _
            (((RegionY - ROIy / 2) + ClickRegion) > PanelPositionY) Then
            Cursor.Current = System.Windows.Forms.Cursors.Hand()
        ElseIf ((RegionX - 0.4 * ROIx) < (PanelPositionX) And _
                (PanelPositionX) < (RegionX + 0.4 * ROIx)) And _
                ((RegionY - 0.4 * ROIy) < (PanelPositionY) And _
                (PanelPositionY) < (RegionY + 0.4 * ROIy)) Then
            'Cursor.Current = System.Windows.Forms.Cursors.NoMove2D
        End If
    End Sub
    Sub SearchRegionDrawing()
        'ClearDisplay
        'DisplayIndicator()
        'Display search region
        With AxGraphicContext3.DrawingRegion
            .EndX = RegionX - ROIx / 2 + ROIx
            .EndY = RegionY - ROIy / 2 + ROIy
            .StartX = RegionX - ROIx / 2
            .StartY = RegionY - ROIy / 2
        End With
        AxGraphicContext3.Rectangle()
        With AxGraphicContext3.DrawingRegion
            .EndX = RegionX - ROIx / 2 + ClickRegion
            .EndY = RegionY - ROIy / 2 + ClickRegion
            .StartX = RegionX - ROIx / 2 - ClickRegion
            .StartY = RegionY - ROIy / 2 - ClickRegion
        End With
        AxGraphicContext3.Rectangle(True, 0)
        With AxGraphicContext3.DrawingRegion
            .EndX = RegionX + ROIx / 2 + ClickRegion
            .EndY = RegionY + ROIy / 2 + ClickRegion
            .StartX = RegionX + ROIx / 2 - ClickRegion
            .StartY = RegionY + ROIy / 2 - ClickRegion
        End With
        AxGraphicContext3.Rectangle(True, 0)
        With AxGraphicContext3.DrawingRegion
            .EndX = RegionX - ROIx / 2 + ClickRegion
            .EndY = RegionY + ROIy / 2 + ClickRegion
            .StartX = RegionX - ROIx / 2 - ClickRegion
            .StartY = RegionY + ROIy / 2 - ClickRegion
        End With
        AxGraphicContext3.Rectangle(True, 0)
        With AxGraphicContext3.DrawingRegion
            .EndX = RegionX + ROIx / 2 + ClickRegion
            .EndY = RegionY - ROIy / 2 + ClickRegion
            .StartX = RegionX + ROIx / 2 - ClickRegion
            .StartY = RegionY - ROIy / 2 - ClickRegion
        End With
        AxGraphicContext3.Rectangle(True, 0)
        With AxGraphicContext3.DrawingRegion
            .EndX = (RegionX) + 10
            .EndY = (RegionY) + 10
            .StartX = (RegionX) - 10
            .StartY = (RegionY) - 10
        End With
        AxGraphicContext3.Cross()
    End Sub
    Sub SearchRegionSetting(ByVal SkipModelRegion As Boolean)
        Dim OffsetX As Decimal
        Dim OffsetY As Decimal
        Dim Oldx, Oldy As Decimal
        Dim newX, newY As Decimal
        If mousedownID = 1 Then
            If SkipModelRegion = False Then
                If PanelPositionX > (DisplayCenterXPosition - MROIx / 2) And PanelPositionY > DisplayCenterYPosition - MROIy / 2 Then
                    PanelPositionX = DisplayCenterXPosition - MROIx / 2
                    PanelPositionY = DisplayCenterYPosition - MROIy / 2
                    'Windows.Forms.Cursor.Current.Position = New System.Drawing.Point(PanelPositionX, PanelPositionY)
                ElseIf PanelPositionX > DisplayCenterXPosition - MROIx / 2 And PanelPositionY < DisplayCenterYPosition - MROIy / 2 Then
                    PanelPositionX = DisplayCenterXPosition - MROIx / 2
                    'PanelPositionY =DisplayCenterYPosition - MROIy / 2
                ElseIf PanelPositionX < DisplayCenterXPosition - MROIx / 2 And PanelPositionY > DisplayCenterYPosition - MROIy / 2 Then
                    'PanelPositionX = DisplayCenterXPosition - MROIx / 2
                    PanelPositionY = DisplayCenterYPosition - MROIy / 2
                End If
            Else
                If PanelPositionX > DisplayCenterXPosition - 10 And PanelPositionY > DisplayCenterYPosition - 10 Then
                    PanelPositionX = DisplayCenterXPosition - 10
                    PanelPositionY = DisplayCenterYPosition - 10
                    'Windows.Forms.Cursor.Current.Position = New System.Drawing.Point(PanelPositionX, PanelPositionY)
                ElseIf PanelPositionX > DisplayCenterXPosition - 10 And PanelPositionY < DisplayCenterYPosition - 10 Then
                    PanelPositionX = DisplayCenterXPosition - 10
                    PanelPositionY = PanelPositionY
                    'Windows.Forms.Cursor.Current.Position = New System.Drawing.Point(PanelPositionX, PanelPositionY)
                ElseIf PanelPositionX < DisplayCenterXPosition - 10 And PanelPositionY > DisplayCenterYPosition - 10 Then
                    PanelPositionX = PanelPositionX
                    PanelPositionY = DisplayCenterYPosition - 10
                    'Windows.Forms.Cursor.Current.Position = New System.Drawing.Point(PanelPositionX, PanelPositionY)
                End If
            End If
            If PanelPositionX < 0 And PanelPositionY < 0 Then
                PanelPositionX = 3
                PanelPositionY = 3
                'Windows.Forms.Cursor.Current.Position = New System.Drawing.Point(PanelPositionX, PanelPositionY)
            ElseIf PanelPositionX < 0 And PanelPositionY > 0 Then
                PanelPositionX = 3
                'PanelPositionY =0
            ElseIf PanelPositionX > 0 And PanelPositionY < 0 Then
                'PanelPositionX = 0
                PanelPositionY = 3
            End If

            Oldx = RegionX - ROIx / 2
            Oldy = RegionY - ROIy / 2
            newX = PanelPositionX
            newY = PanelPositionY
            OffsetX = newX - Oldx
            OffsetY = newY - Oldy
            ROIx = (ROIx - OffsetX)
            ROIy = (ROIy - OffsetY)
            RegionX = DisplayCenterXPosition 'newX + 0.5 * ROIx
            RegionY = DisplayCenterYPosition 'newY + 0.5 * ROIy
        ElseIf mousedownID = 2 Then
            If SkipModelRegion = False Then
                If PanelPositionX < (DisplayCenterXPosition + MROIx / 2) And PanelPositionY < DisplayCenterYPosition + MROIy / 2 Then
                    PanelPositionX = DisplayCenterXPosition + MROIx / 2
                    PanelPositionY = DisplayCenterYPosition + MROIy / 2
                    'Windows.Forms.Cursor.Current.Position = New System.Drawing.Point(PanelPositionX, PanelPositionY)
                ElseIf PanelPositionX < DisplayCenterXPosition + MROIx / 2 And PanelPositionY > DisplayCenterYPosition + MROIy / 2 Then
                    PanelPositionX = DisplayCenterXPosition + MROIx / 2
                    'PanelPositionY =DisplayCenterYPosition + MROIy / 2
                ElseIf PanelPositionX > DisplayCenterXPosition + MROIx / 2 And PanelPositionY < DisplayCenterYPosition + MROIy / 2 Then
                    'PanelPositionX = DisplayCenterXPosition + MROIx / 2
                    PanelPositionY = DisplayCenterYPosition + MROIy / 2
                End If
            Else
                If PanelPositionX < DisplayCenterXPosition + 10 And PanelPositionY < DisplayCenterYPosition + 10 Then
                    PanelPositionX = DisplayCenterXPosition + 10
                    PanelPositionY = DisplayCenterYPosition + 10
                    'Windows.Forms.Cursor.Current.Position = New System.Drawing.Point(PanelPositionX, PanelPositionY)
                ElseIf PanelPositionX < DisplayCenterXPosition + 10 And PanelPositionY > DisplayCenterYPosition + 10 Then
                    PanelPositionX = DisplayCenterXPosition + 10
                    'PanelPositionY = DisplayCenterYPosition + 10
                    'Windows.Forms.Cursor.Current.Position = New System.Drawing.Point(PanelPositionX, PanelPositionY)
                ElseIf PanelPositionX > DisplayCenterXPosition + 10 And PanelPositionY < DisplayCenterYPosition + 10 Then
                    'PanelPositionX = DisplayCenterXPosition + 10
                    PanelPositionY = DisplayCenterYPosition + 10
                    'Windows.Forms.Cursor.Current.Position = New System.Drawing.Point(PanelPositionX, PanelPositionY)
                End If
            End If
            If PanelPositionX > DisplayWidth And PanelPositionY > DisplayHeight Then
                PanelPositionX = DisplayWidth - 3
                PanelPositionY = DisplayHeight - 3
                'Windows.Forms.Cursor.Current.Position = New System.Drawing.Point(PanelPositionX, PanelPositionY)
            ElseIf PanelPositionX > DisplayWidth And PanelPositionY < DisplayHeight Then
                PanelPositionX = DisplayWidth - 3
                'PanelPositionY =0
            ElseIf PanelPositionX < DisplayWidth And PanelPositionY > DisplayHeight Then
                'PanelPositionX = 0
                PanelPositionY = DisplayHeight - 3
            End If
            Oldx = RegionX + ROIx / 2
            Oldy = RegionY + ROIy / 2
            newX = PanelPositionX
            newY = PanelPositionY
            OffsetX = newX - Oldx
            OffsetY = newY - Oldy
            ROIx = (ROIx + OffsetX)
            ROIy = (ROIy + OffsetY)
            RegionX = DisplayCenterXPosition 'newX - 0.5 * ROIx
            RegionY = DisplayCenterYPosition 'newY - 0.5 * ROIy
        ElseIf mousedownID = 3 Then
            If SkipModelRegion = False Then
                If PanelPositionX > (DisplayCenterXPosition - MROIx / 2) And PanelPositionY < DisplayCenterYPosition + MROIy / 2 Then
                    PanelPositionX = DisplayCenterXPosition - MROIx / 2
                    PanelPositionY = DisplayCenterYPosition + MROIy / 2
                    'Windows.Forms.Cursor.Current.Position = New System.Drawing.Point(PanelPositionX, PanelPositionY)
                ElseIf PanelPositionX > DisplayCenterXPosition - MROIx / 2 And PanelPositionY > DisplayCenterYPosition + MROIy / 2 Then
                    PanelPositionX = DisplayCenterXPosition - MROIx / 2
                    'PanelPositionY =DisplayCenterYPosition + MROIy / 2
                ElseIf PanelPositionX < DisplayCenterXPosition - MROIx / 2 And PanelPositionY < DisplayCenterYPosition + MROIy / 2 Then
                    'PanelPositionX = DisplayCenterXPosition - MROIx / 2
                    PanelPositionY = DisplayCenterYPosition + MROIy / 2
                End If
            Else
                If PanelPositionX > DisplayCenterXPosition - 10 And PanelPositionY < DisplayCenterYPosition + 10 Then
                    PanelPositionX = DisplayCenterXPosition - 10
                    PanelPositionY = DisplayCenterYPosition + 10
                    'Windows.Forms.Cursor.Current.Position = New System.Drawing.Point(PanelPositionX, PanelPositionY)
                ElseIf PanelPositionX > DisplayCenterXPosition - 10 And PanelPositionY > DisplayCenterYPosition + 10 Then
                    PanelPositionX = DisplayCenterXPosition - 10
                    'PanelPositionY = DisplayCenterYPosition+ 10
                    'Windows.Forms.Cursor.Current.Position = New System.Drawing.Point(PanelPositionX, PanelPositionY)
                ElseIf PanelPositionX < DisplayCenterXPosition - 10 And PanelPositionY < DisplayCenterYPosition + 10 Then
                    'PanelPositionX = DisplayCenterXPosition - 10
                    PanelPositionY = DisplayCenterYPosition + 10
                    'Windows.Forms.Cursor.Current.Position = New System.Drawing.Point(PanelPositionX, PanelPositionY)
                End If
            End If
            If PanelPositionX < 0 And PanelPositionY > DisplayHeight Then
                PanelPositionX = 3
                PanelPositionY = DisplayHeight - 3
                'Windows.Forms.Cursor.Current.Position = New System.Drawing.Point(PanelPositionX, PanelPositionY)
            ElseIf PanelPositionX < 0 And PanelPositionY < DisplayHeight Then
                PanelPositionX = 3
                'PanelPositionY =0
            ElseIf PanelPositionX > 0 And PanelPositionY > DisplayHeight Then
                'PanelPositionX = 0
                PanelPositionY = DisplayHeight - 3
            End If
            Oldx = RegionX - ROIx / 2
            Oldy = RegionY + ROIy / 2
            newX = PanelPositionX
            newY = PanelPositionY
            OffsetX = newX - Oldx
            OffsetY = newY - Oldy
            ROIx = (ROIx - OffsetX)
            ROIy = (ROIy + OffsetY)
            RegionX = DisplayCenterXPosition 'newX + 0.5 * ROIx
            RegionY = DisplayCenterYPosition 'newY - 0.5 * ROIy
        ElseIf mousedownID = 4 Then
            If SkipModelRegion = False Then
                If PanelPositionX < (DisplayCenterXPosition + MROIx / 2) And PanelPositionY > DisplayCenterYPosition - MROIy / 2 Then
                    PanelPositionX = DisplayCenterXPosition + MROIx / 2
                    PanelPositionY = DisplayCenterYPosition - MROIy / 2
                    'Windows.Forms.Cursor.Current.Position = New System.Drawing.Point(PanelPositionX, PanelPositionY)
                ElseIf PanelPositionX < DisplayCenterXPosition + MROIx / 2 And PanelPositionY < DisplayCenterYPosition - MROIy / 2 Then
                    PanelPositionX = DisplayCenterXPosition + MROIx / 2
                    'PanelPositionY =DisplayCenterYPosition - MROIy / 2
                ElseIf PanelPositionX > DisplayCenterXPosition + MROIx / 2 And PanelPositionY > DisplayCenterYPosition - MROIy / 2 Then
                    'PanelPositionX = DisplayCenterXPosition + MROIx / 2
                    PanelPositionY = DisplayCenterYPosition - MROIy / 2
                End If
            Else
                If PanelPositionX < DisplayCenterXPosition + 10 And PanelPositionY > DisplayCenterYPosition - 10 Then
                    PanelPositionX = DisplayCenterXPosition + 10
                    PanelPositionY = DisplayCenterYPosition + 10
                    'Windows.Forms.Cursor.Current.Position = New System.Drawing.Point(PanelPositionX, PanelPositionY)
                ElseIf PanelPositionX < DisplayCenterXPosition + 10 And PanelPositionY < DisplayCenterYPosition - 10 Then
                    PanelPositionX = DisplayCenterXPosition + 10
                    'PanelPositionY = DisplayCenterYPosition - 10
                    'Windows.Forms.Cursor.Current.Position = New System.Drawing.Point(PanelPositionX, PanelPositionY)
                ElseIf PanelPositionX > DisplayCenterXPosition + 10 And PanelPositionY > DisplayCenterYPosition - 10 Then
                    'PanelPositionX = DisplayCenterXPosition+10 
                    PanelPositionY = DisplayCenterYPosition - 10
                    'Windows.Forms.Cursor.Current.Position = New System.Drawing.Point(PanelPositionX, PanelPositionY)
                End If
            End If
            If PanelPositionX > DisplayWidth And PanelPositionY < 0 Then
                PanelPositionX = DisplayWidth - 3
                PanelPositionY = 3
                'Windows.Forms.Cursor.Current.Position = New System.Drawing.Point(PanelPositionX, PanelPositionY)
            ElseIf PanelPositionX > DisplayWidth And PanelPositionY > 0 Then
                PanelPositionX = DisplayWidth - 3
                'PanelPositionY =0
            ElseIf PanelPositionX < DisplayWidth And PanelPositionY < 0 Then
                'PanelPositionX = 0
                PanelPositionY = 3
            End If
            Oldx = RegionX + ROIx / 2
            Oldy = RegionY - ROIy / 2
            newX = PanelPositionX
            newY = PanelPositionY
            OffsetX = newX - Oldx
            OffsetY = newY - Oldy
            ROIx = (ROIx + OffsetX)
            ROIy = (ROIy - OffsetY)
            RegionX = DisplayCenterXPosition 'newX - 0.5 * ROIx
            RegionY = DisplayCenterYPosition 'newY + 0.5 * ROIy
        ElseIf mousedownID = 5 Then
            If Drag_Flag = 0 Then
                OldXX = MouseClickX
                OldYY = MouseClickY
                Drag_Flag = 1
            End If
            NewXX = PanelPositionX
            NewYY = PanelPositionY
            OffsetX = (NewXX - OldXX)
            OffsetY = (NewYY - OldYY)
            RegionX = RegionX + OffsetX
            RegionY = RegionY + OffsetY
            OldXX = NewXX
            OldYY = NewYY
        End If
        'SearchRegionDrawing()
    End Sub
    Sub ModelSaveImage()
        If AxImage5.IsAllocated = True Then
            AxImage5.Free()
        End If
        With AxImage5.ChildRegion
            .OffsetX = DisplayCenterXPosition - MROIx / 2
            .OffsetY = DisplayCenterYPosition - MROIy / 2
            .SizeX = MROIx
            .SizeY = MROIy
        End With
        AxImage5.Allocate()
        AxImage5.Save("C:\IDS\SRC\DLL Export Device Vision\Fiducial\temp" + ".bmp")
        'AxImage5.Save("c:\" & "b" & ".bmp") '**************Error: could not save if there is a file with same name exists
        If FiducialMark_form.PictureBox20.Image Is Nothing Then
        Else
            Dim image As Image = FiducialMark_form.PictureBox20.Image
            FiducialMark_form.PictureBox20.Image = Nothing
            image.Dispose()
        End If
        FiducialMark_form.PictureBox20.Image() = image.FromFile("C:\IDS\SRC\DLL Export Device Vision\Fiducial\temp" + ".bmp")
    End Sub
#End Region
#Region "Fiducial-PatternMatching"
    Sub PatternMatching_settings()
        Try
            If IndexNo > 0 Or AxPatternMatching2.Models.Count > 0 Then
                IndexNo = 0
                If AxPatternMatching2.Models.Count <> 0 Then
                    For i As Integer = AxPatternMatching2.Models.Count To 1 Step -1
                        AxPatternMatching2.Models.Remove(i)
                    Next
                End If
            End If
            AxPatternMatching2.Free()
            AxPatternMatching2.Allocate()
            'IndexNo = AxPatternMatching2.Models.Add(AxPatternMatching2.Image, 0, 0, AxImage4.SizeX, AxImage4.SizeY, 1, MILPATTERNMATCHINGLib.PatModelTypeConstants.patNormalized)
            IndexNo = AxPatternMatching2.Models.Add(AxPatternMatching1.Image, 0, 0, ModelsPM_sizeX, ModelsPM_sizeY, 1, Matrox.ActiveMIL.PatternMatching.PatModelTypeConstants.patNormalized)
            'IndexNo = AxPatternMatching2.Models.Add(AxImage2, 0, 0, AxImage2.SizeX, AxImage2.SizeY, 1, MILPATTERNMATCHINGLib.PatModelTypeConstants.patNormalized)
            AxPatternMatching2.Models.Item(IndexNo).PositionAccuracy = Matrox.ActiveMIL.PatternMatching.PatHighLowConstants.patVeryHigh
            AxPatternMatching2.Models.Item(IndexNo).AcceptanceThreshold = 40
            AxPatternMatching2.Models.Item(IndexNo).CertaintyThreshold = 95
            AxPatternMatching2.Models.Item(IndexNo).NumberOfOccurrences = 1
            AxPatternMatching2.Models.Item(IndexNo).Speed = Matrox.ActiveMIL.PatternMatching.PatHighLowConstants.patMedium
            AxPatternMatching2.Models.Item(IndexNo).SearchAngle.Enabled = False
            'AxPatternMatching2.Models.Item(IndexNo).SearchAngle.Value = 0
            'AxPatternMatching2.Models.Item(IndexNo).SearchAngle.PositiveDelta = 2
            'AxPatternMatching2.Models.Item(IndexNo).SearchAngle.NegativeDelta = 2
            'AxPatternMatching2.Models.Item(IndexNo).SearchAngle.Accuracy = 0.1
            'AxPatternMatching2.Models.Item(IndexNo).SearchAngle.Tolerance = 0.1
            'AxPatternMatching2.Models.Item(IndexNo).SearchRegion.FullSize() = True
            AxPatternMatching2.Models.Item(IndexNo).SearchRegion.StartX = RegionX - ROIx / 2 'Panels actual position== refer to display's origin
            AxPatternMatching2.Models.Item(IndexNo).SearchRegion.StartY = RegionY - ROIy / 2 'Panels actual position== refer to display's origin
            AxPatternMatching2.Models.Item(IndexNo).SearchRegion.SizeX = ROIx
            AxPatternMatching2.Models.Item(IndexNo).SearchRegion.SizeY = ROIy
            AxPatternMatching2.Models.Item(IndexNo).Preprocess()
            'ClearDisplay
            'DisplayIndicator()
        Catch ex As Exception
            ExceptionDisplay(ex)
        End Try
    End Sub
#End Region
#Region "FModels"
    Dim PixelSize As Double = PixelSizeX '22micron per pixel
    Dim imgNo As Integer = 0
    Dim imgNoOld As Integer = 0
    Dim ModelsPM_sizeX As Integer
    Dim ModelsPM_sizeY As Integer
    Dim ImgNoSQ, ImgNoOldSQ As Integer
    Function FCircle(ByVal FDiameter As Double) As Boolean
        Dim diameter_pixel As Double = 1
        PixelSize = PixelSizeX
        Dim diameter_mm As Double = FDiameter
        diameter_pixel = diameter_mm / PixelSizeX  'Convertion of mm --> pixel

        If FiducialMark_form.PictureBox1.BorderStyle = BorderStyle.Fixed3D Then
            diameter_pixel = diameter_mm / PixelSizeX  'Convertion of mm --> pixel
            Dim remainder As Integer = diameter_pixel Mod 2
            If AxImage2.IsAllocated = False Then
                With AxImage2
                    .SizeX = Convert.ToInt16(diameter_pixel)
                    .SizeY = Convert.ToInt16(diameter_pixel)
                    .CanProcess = True
                    .CanDisplay = True
                    .Allocate()
                End With
            Else
                AxImage2.Free()

                With AxImage2
                    If remainder = 1 Then
                        .SizeX = Convert.ToInt16(diameter_pixel)
                        .SizeY = Convert.ToInt16(diameter_pixel)
                    Else
                        .SizeX = Convert.ToInt16(diameter_pixel) + 1 'to centralize the circle
                        .SizeY = Convert.ToInt16(diameter_pixel) + 1
                    End If
                    .CanProcess = True
                    .CanDisplay = True
                    .Allocate()
                End With
                AxImage2.Clear(255) '255 for b/w, 0 for w/b
            End If
            ModelsPM_sizeX = AxImage2.SizeX
            ModelsPM_sizeY = AxImage2.SizeY

            With AxGraphicContext5.DrawingRegion
                .StartX = 0
                .StartY = 0
                If remainder = 1 Then
                    .EndX = Convert.ToInt16(diameter_pixel)
                    .EndY = Convert.ToInt16(diameter_pixel)
                Else
                    .EndX = Convert.ToInt16(diameter_pixel) + 1
                    .EndY = Convert.ToInt16(diameter_pixel) + 1
                End If
            End With
            AxGraphicContext5.BackgroundColor = AxGraphicContext5.BackgroundColor.White
            AxGraphicContext5.ForegroundColor = AxGraphicContext5.BackgroundColor.Black
            AxGraphicContext5.Arc(0, 360, True)

            '***********************************************************************
            AxImage2.Save("C:\IDS\SRC\DLL Export Device Vision\Fiducial\" + "temp.bmp")
            Return True

        ElseIf FiducialMark_form.PictureBox11.BorderStyle = BorderStyle.Fixed3D Then
            diameter_pixel = diameter_mm / PixelSizeX  'Convertion of mm --> pixel
            Dim remainder As Integer = diameter_pixel Mod 2
            If AxImage2.IsAllocated = False Then
                With AxImage2
                    .SizeX = Convert.ToInt16(diameter_pixel)
                    .SizeY = Convert.ToInt16(diameter_pixel)
                    .CanProcess = True
                    .CanDisplay = True
                    .Allocate()
                End With
            Else
                AxImage2.Free()

                With AxImage2
                    If remainder = 1 Then
                        .SizeX = Convert.ToInt16(diameter_pixel)
                        .SizeY = Convert.ToInt16(diameter_pixel)
                    Else
                        .SizeX = Convert.ToInt16(diameter_pixel) + 1 'to centralize the circle
                        .SizeY = Convert.ToInt16(diameter_pixel) + 1
                    End If
                    .CanProcess = True
                    .CanDisplay = True
                    .Allocate()
                End With
                AxImage2.Clear(0) '255 for b/w, 0 for w/b
            End If
            ModelsPM_sizeX = AxImage2.SizeX
            ModelsPM_sizeY = AxImage2.SizeY

            With AxGraphicContext5.DrawingRegion
                .StartX = 0
                .StartY = 0
                If remainder = 1 Then
                    .EndX = Convert.ToInt16(diameter_pixel)
                    .EndY = Convert.ToInt16(diameter_pixel)
                Else
                    .EndX = Convert.ToInt16(diameter_pixel) + 1
                    .EndY = Convert.ToInt16(diameter_pixel) + 1
                End If
            End With
            AxGraphicContext5.BackgroundColor = AxGraphicContext5.BackgroundColor.Black
            AxGraphicContext5.ForegroundColor = AxGraphicContext5.BackgroundColor.White
            AxGraphicContext5.Arc(0, 360, True)

            '***********************************************************************
            AxImage2.Save("C:\IDS\SRC\DLL Export Device Vision\Fiducial\" + "temp.bmp")
            Return True
        End If
    End Function
    Function FSquare(ByVal FX As Double, ByVal FY As Double) As Integer
        Dim FX_pixel, FY_pixel As Double
        Dim X_mm As Double = FX
        Dim Y_mm As Double = FY

        If FiducialMark_form.PictureBox14.BorderStyle = BorderStyle.Fixed3D Then
            Try
                FX_pixel = FX / PixelSizeX  'Convertion of mm --> pixel
                FY_pixel = FY / PixelSizeX
                Dim RemainderX As Integer = FX_pixel Mod 2
                Dim RemainderY As Integer = FY_pixel Mod 2
                If AxImage2.IsAllocated = False Then
                    With AxImage2
                        .SizeX = Convert.ToInt16(FX_pixel)
                        .SizeY = Convert.ToInt16(FY_pixel)
                        .CanProcess = True
                        .CanDisplay = True
                        .Allocate()
                    End With
                Else
                    AxImage2.Free()
                    With AxImage2
                        If RemainderX = 1 Then
                            .SizeX = Convert.ToInt16(FX_pixel) + 2
                        Else
                            .SizeX = Convert.ToInt16(FX_pixel) + 3
                        End If
                        If RemainderY = 1 Then
                            .SizeY = Convert.ToInt16(FY_pixel) + 2
                        Else
                            .SizeY = Convert.ToInt16(FY_pixel) + 3
                        End If
                        .CanProcess = True
                        .CanDisplay = True
                        .Allocate()
                    End With
                    AxImage2.Clear(0)
                End If
                ModelsPM_sizeX = AxImage2.SizeX
                ModelsPM_sizeY = AxImage2.SizeY
                With AxGraphicContext5.DrawingRegion
                    .StartX = 1
                    .StartY = 1
                    If RemainderX = 1 Then
                        .EndX = Convert.ToInt16(FX_pixel)
                    Else
                        .EndX = Convert.ToInt16(FX_pixel) + 1
                    End If
                    If RemainderY = 1 Then
                        .EndY = Convert.ToInt16(FY_pixel)
                    Else
                        .EndY = Convert.ToInt16(FY_pixel) + 1
                    End If
                End With
                AxGraphicContext5.BackgroundColor = AxGraphicContext5.BackgroundColor.Black
                AxGraphicContext5.ForegroundColor = AxGraphicContext5.BackgroundColor.White
                AxGraphicContext5.Rectangle(True, 0)

                ImgNoSQ = 1 'ImgNoSQ + 1
                'If ImgNoSQ <> ImgNoOldSQ Then
                ImgNoOldSQ = ImgNoSQ
                AxImage2.Save("C:\IDS\SRC\DLL Export Device Vision\Fiducial\sq" & ImgNoSQ.ToString & ".bmp")
                Return ImgNoSQ
                'End If
            Catch ex As Exception
                ExceptionDisplay(ex)
            End Try
        ElseIf FiducialMark_form.PictureBox4.BorderStyle = BorderStyle.Fixed3D Then
            Try
                FX_pixel = FX / PixelSizeX  'Convertion of mm --> pixel
                FY_pixel = FY / PixelSizeX
                Dim RemainderX As Integer = FX_pixel Mod 2
                Dim RemainderY As Integer = FY_pixel Mod 2
                If AxImage2.IsAllocated = False Then
                    With AxImage2
                        .SizeX = Convert.ToInt16(FX_pixel)
                        .SizeY = Convert.ToInt16(FY_pixel)
                        .CanProcess = True
                        .CanDisplay = True
                        .Allocate()
                    End With
                Else
                    AxImage2.Free()
                    With AxImage2
                        If RemainderX = 1 Then
                            .SizeX = Convert.ToInt16(FX_pixel) + 2
                        Else
                            .SizeX = Convert.ToInt16(FX_pixel) + 3
                        End If
                        If RemainderY = 1 Then
                            .SizeY = Convert.ToInt16(FY_pixel) + 2
                        Else
                            .SizeY = Convert.ToInt16(FY_pixel) + 3
                        End If
                        .CanProcess = True
                        .CanDisplay = True
                        .Allocate()
                    End With
                    AxImage2.Clear(255)
                End If
                ModelsPM_sizeX = AxImage2.SizeX
                ModelsPM_sizeY = AxImage2.SizeY
                With AxGraphicContext5.DrawingRegion
                    .StartX = 1
                    .StartY = 1
                    If RemainderX = 1 Then
                        .EndX = Convert.ToInt16(FX_pixel)
                    Else
                        .EndX = Convert.ToInt16(FX_pixel) + 1
                    End If
                    If RemainderY = 1 Then
                        .EndY = Convert.ToInt16(FY_pixel)
                    Else
                        .EndY = Convert.ToInt16(FY_pixel) + 1
                    End If
                End With
                AxGraphicContext5.BackgroundColor = AxGraphicContext5.BackgroundColor.White
                AxGraphicContext5.ForegroundColor = AxGraphicContext5.BackgroundColor.Black
                AxGraphicContext5.Rectangle(True, 0)

                ImgNoSQ = 1 'ImgNoSQ + 1
                'If ImgNoSQ <> ImgNoOldSQ Then
                ImgNoOldSQ = ImgNoSQ
                AxImage2.Save("C:\IDS\SRC\DLL Export Device Vision\Fiducial\sq" & ImgNoSQ.ToString & ".bmp")
                Return ImgNoSQ
                'End If
            Catch ex As Exception
                ExceptionDisplay(ex)
            End Try
        End If
    End Function
    Function FLFlag(ByVal FX As Double)
        Dim FX_pixel As Double
        'Dim FY_pixel As Double
        Dim X_mm As Double = FX
        'Dim Y_mm As Double = FY

        If FiducialMark_form.PictureBox15.BorderStyle = BorderStyle.Fixed3D Then
            Try
                FX_pixel = FX / PixelSizeX  'Convertion of mm --> pixel
                'FY_pixel = FY / PixelSizeX
                Dim RemainderX As Integer = FX_pixel Mod 2
                'Dim RemainderY As Integer = FY_pixel Mod 2
                If AxImage2.IsAllocated = False Then
                    With AxImage2
                        .SizeX = Convert.ToInt16(FX_pixel)
                        .SizeY = Convert.ToInt16(FX_pixel)
                        .CanProcess = True
                        .CanDisplay = True
                        .Allocate()
                    End With
                Else
                    AxImage2.Free()
                    With AxImage2
                        If RemainderX = 0 Or RemainderX = 2 Then
                            .SizeX = Convert.ToInt16(FX_pixel)
                        Else
                            .SizeX = Convert.ToInt16(FX_pixel) + 1
                        End If
                        If RemainderX = 0 Or RemainderX = 2 Then
                            .SizeY = Convert.ToInt16(FX_pixel)
                        Else
                            .SizeY = Convert.ToInt16(FX_pixel) + 1
                        End If
                        .CanProcess = True
                        .CanDisplay = True
                        .Allocate()
                    End With
                    AxImage2.Clear(255)
                End If
                ModelsPM_sizeX = AxImage2.SizeX
                ModelsPM_sizeY = AxImage2.SizeY
                With AxGraphicContext5.DrawingRegion
                    .StartX = 0
                    .StartY = 0
                    If RemainderX = 0 Or RemainderX = 2 Then
                        .EndX = (Convert.ToInt16(FX_pixel)) / 2 - 1
                    Else
                        .EndX = (Convert.ToInt16(FX_pixel) + 1) / 2 - 1
                    End If
                    If RemainderX = 0 Or RemainderX = 2 Then
                        .EndY = (Convert.ToInt16(FX_pixel)) / 2 - 1
                    Else
                        .EndY = (Convert.ToInt16(FX_pixel) + 1) / 2 - 1
                    End If
                End With
                AxGraphicContext5.BackgroundColor = AxGraphicContext5.BackgroundColor.White
                AxGraphicContext5.ForegroundColor = AxGraphicContext5.BackgroundColor.Black
                AxGraphicContext5.Rectangle(True, 0)
                With AxGraphicContext5.DrawingRegion
                    If RemainderX = 0 Or RemainderX = 2 Then
                        .StartX = (Convert.ToInt16(FX_pixel)) / 2
                    Else
                        .StartX = (Convert.ToInt16(FX_pixel) + 1) / 2
                    End If
                    If RemainderX = 0 Or RemainderX = 2 Then
                        .StartY = (Convert.ToInt16(FX_pixel)) / 2
                    Else
                        .StartY = (Convert.ToInt16(FX_pixel) + 1) / 2
                    End If
                    If RemainderX = 0 Or RemainderX = 2 Then
                        .EndX = Convert.ToInt16(FX_pixel)
                    Else
                        .EndX = Convert.ToInt16(FX_pixel) + 1
                    End If
                    If RemainderX = 0 Or RemainderX = 2 Then
                        .EndY = Convert.ToInt16(FX_pixel)
                    Else
                        .EndY = Convert.ToInt16(FX_pixel) + 1
                    End If
                End With
                AxGraphicContext5.BackgroundColor = AxGraphicContext5.BackgroundColor.White
                AxGraphicContext5.ForegroundColor = AxGraphicContext5.BackgroundColor.Black
                AxGraphicContext5.Rectangle(True, 0)

                ImgNoSQ = 1 'ImgNoSQ + 1
                'If ImgNoSQ <> ImgNoOldSQ Then
                ImgNoOldSQ = ImgNoSQ
                AxImage2.Save("C:\IDS\SRC\DLL Export Device Vision\Fiducial\sq" & ImgNoSQ.ToString & ".bmp")
                Return ImgNoSQ
                'End If
            Catch ex As Exception
                ExceptionDisplay(ex)
            End Try
        ElseIf FiducialMark_form.PictureBox5.BorderStyle = BorderStyle.Fixed3D Then
            Try
                FX_pixel = FX / PixelSizeX  'Convertion of mm --> pixel
                'FY_pixel = FY / PixelSizeX
                Dim RemainderX As Integer = FX_pixel Mod 2
                'Dim RemainderY As Integer = FY_pixel Mod 2
                If AxImage2.IsAllocated = False Then
                    With AxImage2
                        .SizeX = Convert.ToInt16(FX_pixel)
                        .SizeY = Convert.ToInt16(FX_pixel)
                        .CanProcess = True
                        .CanDisplay = True
                        .Allocate()
                    End With
                Else
                    AxImage2.Free()
                    With AxImage2
                        If RemainderX = 0 Or RemainderX = 2 Then
                            .SizeX = Convert.ToInt16(FX_pixel)
                        Else
                            .SizeX = Convert.ToInt16(FX_pixel) + 1
                        End If
                        If RemainderX = 0 Or RemainderX = 2 Then
                            .SizeY = Convert.ToInt16(FX_pixel)
                        Else
                            .SizeY = Convert.ToInt16(FX_pixel) + 1
                        End If
                        .CanProcess = True
                        .CanDisplay = True
                        .Allocate()
                    End With
                    AxImage2.Clear(0)
                End If
                ModelsPM_sizeX = AxImage2.SizeX
                ModelsPM_sizeY = AxImage2.SizeY
                With AxGraphicContext5.DrawingRegion
                    .StartX = 0
                    .StartY = 0
                    If RemainderX = 0 Or RemainderX = 2 Then
                        .EndX = (Convert.ToInt16(FX_pixel)) / 2 - 1
                    Else
                        .EndX = (Convert.ToInt16(FX_pixel) + 1) / 2 - 1
                    End If
                    If RemainderX = 0 Or RemainderX = 2 Then
                        .EndY = (Convert.ToInt16(FX_pixel)) / 2 - 1
                    Else
                        .EndY = (Convert.ToInt16(FX_pixel) + 1) / 2 - 1
                    End If
                End With
                AxGraphicContext5.BackgroundColor = AxGraphicContext5.BackgroundColor.Black
                AxGraphicContext5.ForegroundColor = AxGraphicContext5.BackgroundColor.White
                AxGraphicContext5.Rectangle(True, 0)
                With AxGraphicContext5.DrawingRegion
                    If RemainderX = 0 Or RemainderX = 2 Then
                        .StartX = (Convert.ToInt16(FX_pixel)) / 2
                    Else
                        .StartX = (Convert.ToInt16(FX_pixel) + 1) / 2
                    End If
                    If RemainderX = 0 Or RemainderX = 2 Then
                        .StartY = (Convert.ToInt16(FX_pixel)) / 2
                    Else
                        .StartY = (Convert.ToInt16(FX_pixel) + 1) / 2
                    End If
                    If RemainderX = 0 Or RemainderX = 2 Then
                        .EndX = Convert.ToInt16(FX_pixel)
                    Else
                        .EndX = Convert.ToInt16(FX_pixel) + 1
                    End If
                    If RemainderX = 0 Or RemainderX = 2 Then
                        .EndY = Convert.ToInt16(FX_pixel)
                    Else
                        .EndY = Convert.ToInt16(FX_pixel) + 1
                    End If
                End With
                AxGraphicContext5.BackgroundColor = AxGraphicContext5.BackgroundColor.Black
                AxGraphicContext5.ForegroundColor = AxGraphicContext5.BackgroundColor.White
                AxGraphicContext5.Rectangle(True, 0)

                ImgNoSQ = 1 'ImgNoSQ + 1
                'If ImgNoSQ <> ImgNoOldSQ Then
                ImgNoOldSQ = ImgNoSQ
                AxImage2.Save("C:\IDS\SRC\DLL Export Device Vision\Fiducial\sq" & ImgNoSQ.ToString & ".bmp")
                Return ImgNoSQ
                'End If
            Catch ex As Exception
                ExceptionDisplay(ex)
            End Try
        End If
    End Function
    Function FRFlag(ByVal FX As Double)
        Dim FX_pixel As Double
        'Dim FY_pixel As Double
        Dim X_mm As Double = FX
        'Dim Y_mm As Double = FY

        If FiducialMark_form.PictureBox16.BorderStyle = BorderStyle.Fixed3D Then
            Try
                FX_pixel = FX / PixelSizeX  'Convertion of mm --> pixel
                'FY_pixel = FY / PixelSizeX
                Dim RemainderX As Integer = FX_pixel Mod 2
                'Dim RemainderY As Integer = FY_pixel Mod 2
                If AxImage2.IsAllocated = False Then
                    With AxImage2
                        .SizeX = Convert.ToInt16(FX_pixel)
                        .SizeY = Convert.ToInt16(FX_pixel)
                        .CanProcess = True
                        .CanDisplay = True
                        .Allocate()
                    End With
                Else
                    AxImage2.Free()
                    With AxImage2
                        If RemainderX = 0 Or RemainderX = 2 Then
                            .SizeX = Convert.ToInt16(FX_pixel)
                        Else
                            .SizeX = Convert.ToInt16(FX_pixel) + 1
                        End If
                        If RemainderX = 0 Or RemainderX = 2 Then
                            .SizeY = Convert.ToInt16(FX_pixel)
                        Else
                            .SizeY = Convert.ToInt16(FX_pixel) + 1
                        End If
                        .CanProcess = True
                        .CanDisplay = True
                        .Allocate()
                    End With
                    AxImage2.Clear(0)
                End If
                ModelsPM_sizeX = AxImage2.SizeX
                ModelsPM_sizeY = AxImage2.SizeY
                With AxGraphicContext5.DrawingRegion
                    .StartX = 0
                    .StartY = 0
                    If RemainderX = 0 Or RemainderX = 2 Then
                        .EndX = (Convert.ToInt16(FX_pixel)) / 2 - 1
                    Else
                        .EndX = (Convert.ToInt16(FX_pixel) + 1) / 2 - 1
                    End If
                    If RemainderX = 0 Or RemainderX = 2 Then
                        .EndY = (Convert.ToInt16(FX_pixel)) / 2 - 1
                    Else
                        .EndY = (Convert.ToInt16(FX_pixel) + 1) / 2 - 1
                    End If
                End With
                AxGraphicContext5.BackgroundColor = AxGraphicContext5.BackgroundColor.Black
                AxGraphicContext5.ForegroundColor = AxGraphicContext5.BackgroundColor.White
                AxGraphicContext5.Rectangle(True, 0)
                With AxGraphicContext5.DrawingRegion
                    If RemainderX = 0 Or RemainderX = 2 Then
                        .StartX = (Convert.ToInt16(FX_pixel)) / 2
                    Else
                        .StartX = (Convert.ToInt16(FX_pixel) + 1) / 2
                    End If
                    If RemainderX = 0 Or RemainderX = 2 Then
                        .StartY = (Convert.ToInt16(FX_pixel)) / 2
                    Else
                        .StartY = (Convert.ToInt16(FX_pixel) + 1) / 2
                    End If
                    If RemainderX = 0 Or RemainderX = 2 Then
                        .EndX = Convert.ToInt16(FX_pixel)
                    Else
                        .EndX = Convert.ToInt16(FX_pixel) + 1
                    End If
                    If RemainderX = 0 Or RemainderX = 2 Then
                        .EndY = Convert.ToInt16(FX_pixel)
                    Else
                        .EndY = Convert.ToInt16(FX_pixel) + 1
                    End If
                End With
                AxGraphicContext5.BackgroundColor = AxGraphicContext5.BackgroundColor.Black
                AxGraphicContext5.ForegroundColor = AxGraphicContext5.BackgroundColor.White
                AxGraphicContext5.Rectangle(True, 0)

                ImgNoSQ = 1 'ImgNoSQ + 1
                'If ImgNoSQ <> ImgNoOldSQ Then
                ImgNoOldSQ = ImgNoSQ
                AxImage2.Save("C:\IDS\SRC\DLL Export Device Vision\Fiducial\sq" & ImgNoSQ.ToString & ".bmp")
                Return ImgNoSQ
                'End If
            Catch ex As Exception
                ExceptionDisplay(ex)
            End Try
        ElseIf FiducialMark_form.PictureBox6.BorderStyle = BorderStyle.Fixed3D Then
            Try
                FX_pixel = FX / PixelSizeX  'Convertion of mm --> pixel
                'FY_pixel = FY / PixelSizeX
                Dim RemainderX As Integer = FX_pixel Mod 2
                'Dim RemainderY As Integer = FY_pixel Mod 2
                If AxImage2.IsAllocated = False Then
                    With AxImage2
                        .SizeX = Convert.ToInt16(FX_pixel)
                        .SizeY = Convert.ToInt16(FX_pixel)
                        .CanProcess = True
                        .CanDisplay = True
                        .Allocate()
                    End With
                Else
                    AxImage2.Free()
                    With AxImage2
                        If RemainderX = 0 Or RemainderX = 2 Then
                            .SizeX = Convert.ToInt16(FX_pixel)
                        Else
                            .SizeX = Convert.ToInt16(FX_pixel) + 1
                        End If
                        If RemainderX = 0 Or RemainderX = 2 Then
                            .SizeY = Convert.ToInt16(FX_pixel)
                        Else
                            .SizeY = Convert.ToInt16(FX_pixel) + 1
                        End If
                        .CanProcess = True
                        .CanDisplay = True
                        .Allocate()
                    End With
                    AxImage2.Clear(255)
                End If
                ModelsPM_sizeX = AxImage2.SizeX
                ModelsPM_sizeY = AxImage2.SizeY
                With AxGraphicContext5.DrawingRegion
                    .StartX = 0
                    .StartY = 0
                    If RemainderX = 0 Or RemainderX = 2 Then
                        .EndX = (Convert.ToInt16(FX_pixel)) / 2 - 1
                    Else
                        .EndX = (Convert.ToInt16(FX_pixel) + 1) / 2 - 1
                    End If
                    If RemainderX = 0 Or RemainderX = 2 Then
                        .EndY = (Convert.ToInt16(FX_pixel)) / 2 - 1
                    Else
                        .EndY = (Convert.ToInt16(FX_pixel) + 1) / 2 - 1
                    End If
                End With
                AxGraphicContext5.BackgroundColor = AxGraphicContext5.BackgroundColor.White
                AxGraphicContext5.ForegroundColor = AxGraphicContext5.BackgroundColor.Black
                AxGraphicContext5.Rectangle(True, 0)
                With AxGraphicContext5.DrawingRegion
                    If RemainderX = 0 Or RemainderX = 2 Then
                        .StartX = (Convert.ToInt16(FX_pixel)) / 2
                    Else
                        .StartX = (Convert.ToInt16(FX_pixel) + 1) / 2
                    End If
                    If RemainderX = 0 Or RemainderX = 2 Then
                        .StartY = (Convert.ToInt16(FX_pixel)) / 2
                    Else
                        .StartY = (Convert.ToInt16(FX_pixel) + 1) / 2
                    End If
                    If RemainderX = 0 Or RemainderX = 2 Then
                        .EndX = Convert.ToInt16(FX_pixel)
                    Else
                        .EndX = Convert.ToInt16(FX_pixel) + 1
                    End If
                    If RemainderX = 0 Or RemainderX = 2 Then
                        .EndY = Convert.ToInt16(FX_pixel)
                    Else
                        .EndY = Convert.ToInt16(FX_pixel) + 1
                    End If
                End With
                AxGraphicContext5.BackgroundColor = AxGraphicContext5.BackgroundColor.White
                AxGraphicContext5.ForegroundColor = AxGraphicContext5.BackgroundColor.Black
                AxGraphicContext5.Rectangle(True, 0)

                ImgNoSQ = 1 'ImgNoSQ + 1
                'If ImgNoSQ <> ImgNoOldSQ Then
                ImgNoOldSQ = ImgNoSQ
                AxImage2.Save("C:\IDS\SRC\DLL Export Device Vision\Fiducial\sq" & ImgNoSQ.ToString & ".bmp")
                Return ImgNoSQ
                'End If
            Catch ex As Exception
                ExceptionDisplay(ex)
            End Try
        End If
    End Function
    Function FDiamond1(ByVal FX As Double)
        Dim FX_pixel As Double
        Dim X_mm As Double = FX

        If FiducialMark_form.PictureBox17.BorderStyle = BorderStyle.Fixed3D Then
            Try
                FX_pixel = FX / PixelSizeX  'Convertion of mm --> pixel
                Dim RemainderX As Integer = FX_pixel Mod 2
                If AxImage2.IsAllocated = False Then
                    With AxImage2
                        .SizeX = Convert.ToInt16(FX_pixel)
                        .SizeY = Convert.ToInt16(FX_pixel)
                        .CanProcess = True
                        .CanDisplay = True
                        .Allocate()
                    End With
                Else
                    AxImage2.Free()
                    With AxImage2
                        If RemainderX = 1 Then
                            .SizeX = Convert.ToInt16(FX_pixel)
                        Else
                            .SizeX = Convert.ToInt16(FX_pixel) + 1
                        End If
                        If RemainderX = 1 Then
                            .SizeY = Convert.ToInt16(FX_pixel)
                        Else
                            .SizeY = Convert.ToInt16(FX_pixel) + 1
                        End If
                        .CanProcess = True
                        .CanDisplay = True
                        .Allocate()
                    End With
                    AxImage2.Clear(0)
                End If
                ModelsPM_sizeX = AxImage2.SizeX
                ModelsPM_sizeY = AxImage2.SizeY

                Dim Diagonal As Double = Sqrt(AxImage2.SizeX ^ 2 + AxImage2.SizeY ^ 2)
                With AxGraphicContext5.DrawingRegion
                    If RemainderX = 1 Then
                        .StartX = AxImage2.SizeX / 2 - (Convert.ToInt16(FX_pixel - 1) * AxImage2.SizeX / Diagonal) / 2
                    Else
                        .StartX = AxImage2.SizeX / 2 - ((Convert.ToInt16(FX_pixel)) * AxImage2.SizeX / Diagonal) / 2
                    End If
                    If RemainderX = 1 Then
                        .StartY = AxImage2.SizeY / 2 - (Convert.ToInt16(FX_pixel - 1) * AxImage2.SizeY / Diagonal) / 2
                    Else
                        .StartY = AxImage2.SizeY / 2 - ((Convert.ToInt16(FX_pixel)) * AxImage2.SizeY / Diagonal) / 2
                    End If
                    If RemainderX = 1 Then
                        .EndX = AxImage2.SizeX / 2 + (Convert.ToInt16(FX_pixel - 2) * AxImage2.SizeX / Diagonal) / 2
                    Else
                        .EndX = AxImage2.SizeX / 2 + ((Convert.ToInt16(FX_pixel) - 1) * AxImage2.SizeX / Diagonal) / 2
                    End If
                    If RemainderX = 1 Then
                        .EndY = AxImage2.SizeX / 2 + (Convert.ToInt16(FX_pixel - 2) * AxImage2.SizeY / Diagonal) / 2
                    Else
                        .EndY = AxImage2.SizeX / 2 + ((Convert.ToInt16(FX_pixel) - 1) * AxImage2.SizeY / Diagonal) / 2
                    End If
                End With
                AxGraphicContext5.BackgroundColor = AxGraphicContext5.BackgroundColor.Black
                AxGraphicContext5.ForegroundColor = AxGraphicContext5.BackgroundColor.White
                AxGraphicContext5.Rectangle(True, -45)

                ImgNoSQ = 1 'ImgNoSQ + 1
                'If ImgNoSQ <> ImgNoOldSQ Then
                ImgNoOldSQ = ImgNoSQ
                AxImage2.Save("C:\IDS\SRC\DLL Export Device Vision\Fiducial\sq" & ImgNoSQ.ToString & ".bmp")
                Return ImgNoSQ
                'End If
            Catch ex As Exception
                ExceptionDisplay(ex)
            End Try
        ElseIf FiducialMark_form.PictureBox7.BorderStyle = BorderStyle.Fixed3D Then
            Try
                FX_pixel = FX / PixelSizeX  'Convertion of mm --> pixel
                Dim RemainderX As Integer = FX_pixel Mod 2
                If AxImage2.IsAllocated = False Then
                    With AxImage2
                        .SizeX = Convert.ToInt16(FX_pixel)
                        .SizeY = Convert.ToInt16(FX_pixel)
                        .CanProcess = True
                        .CanDisplay = True
                        .Allocate()
                    End With
                Else
                    AxImage2.Free()
                    With AxImage2
                        If RemainderX = 1 Then
                            .SizeX = Convert.ToInt16(FX_pixel)
                        Else
                            .SizeX = Convert.ToInt16(FX_pixel) + 1
                        End If
                        If RemainderX = 1 Then
                            .SizeY = Convert.ToInt16(FX_pixel)
                        Else
                            .SizeY = Convert.ToInt16(FX_pixel) + 1
                        End If
                        .CanProcess = True
                        .CanDisplay = True
                        .Allocate()
                    End With
                    AxImage2.Clear(255)
                End If
                ModelsPM_sizeX = AxImage2.SizeX
                ModelsPM_sizeY = AxImage2.SizeY

                Dim Diagonal As Double = Sqrt(AxImage2.SizeX ^ 2 + AxImage2.SizeY ^ 2)
                With AxGraphicContext5.DrawingRegion
                    If RemainderX = 1 Then
                        .StartX = AxImage2.SizeX / 2 - (Convert.ToInt16(FX_pixel - 1) * AxImage2.SizeX / Diagonal) / 2
                    Else
                        .StartX = AxImage2.SizeX / 2 - ((Convert.ToInt16(FX_pixel)) * AxImage2.SizeX / Diagonal) / 2
                    End If
                    If RemainderX = 1 Then
                        .StartY = AxImage2.SizeY / 2 - (Convert.ToInt16(FX_pixel - 1) * AxImage2.SizeY / Diagonal) / 2
                    Else
                        .StartY = AxImage2.SizeY / 2 - ((Convert.ToInt16(FX_pixel)) * AxImage2.SizeY / Diagonal) / 2
                    End If
                    If RemainderX = 1 Then
                        .EndX = AxImage2.SizeX / 2 + (Convert.ToInt16(FX_pixel - 2) * AxImage2.SizeX / Diagonal) / 2
                    Else
                        .EndX = AxImage2.SizeX / 2 + ((Convert.ToInt16(FX_pixel) - 1) * AxImage2.SizeX / Diagonal) / 2
                    End If
                    If RemainderX = 1 Then
                        .EndY = AxImage2.SizeX / 2 + (Convert.ToInt16(FX_pixel - 2) * AxImage2.SizeY / Diagonal) / 2
                    Else
                        .EndY = AxImage2.SizeX / 2 + ((Convert.ToInt16(FX_pixel) - 1) * AxImage2.SizeY / Diagonal) / 2
                    End If
                End With
                AxGraphicContext5.BackgroundColor = AxGraphicContext5.BackgroundColor.White
                AxGraphicContext5.ForegroundColor = AxGraphicContext5.BackgroundColor.Black
                AxGraphicContext5.Rectangle(True, -45)

                ImgNoSQ = 1 'ImgNoSQ + 1
                'If ImgNoSQ <> ImgNoOldSQ Then
                ImgNoOldSQ = ImgNoSQ
                AxImage2.Save("C:\IDS\SRC\DLL Export Device Vision\Fiducial\sq" & ImgNoSQ.ToString & ".bmp")
                Return ImgNoSQ
                'End If
            Catch ex As Exception
                ExceptionDisplay(ex)
            End Try
        End If
    End Function 'Failure
    Function FDiamond(ByVal FX As Double, ByVal FY As Double)
        Dim FX_pixel, FY_pixel As Double
        Dim X_mm As Double = FX
        Dim Y_mm As Double = FY

        If FiducialMark_form.PictureBox7.BorderStyle = BorderStyle.Fixed3D Then
            Try
                FX_pixel = FX / PixelSizeX  'Convertion of mm --> pixel
                FY_pixel = FY / PixelSizeX
                Dim RemainderX As Integer = FX_pixel Mod 2
                Dim RemainderY As Integer = FY_pixel Mod 2
                If AxImage2.IsAllocated = False Then
                    With AxImage2
                        .SizeX = Convert.ToInt16(FX_pixel)
                        .SizeY = Convert.ToInt16(FY_pixel)
                        .CanProcess = True
                        .CanDisplay = True
                        .Allocate()
                    End With
                Else
                    AxImage2.Free()
                    With AxImage2
                        If RemainderX = 1 Then
                            .SizeX = Convert.ToInt16(FX_pixel)
                        Else
                            .SizeX = Convert.ToInt16(FX_pixel) + 1
                        End If
                        If RemainderY = 1 Then
                            .SizeY = Convert.ToInt16(FY_pixel)
                        Else
                            .SizeY = Convert.ToInt16(FY_pixel) + 1
                        End If
                        .CanProcess = True
                        .CanDisplay = True
                        .Allocate()
                    End With
                    AxImage2.Clear(255)
                End If
                ModelsPM_sizeX = AxImage2.SizeX
                ModelsPM_sizeY = AxImage2.SizeY

                With AxGraphicContext5.DrawingRegion
                    .StartX = 0
                    .StartY = (AxImage2.SizeY - 1) / 2
                    .EndX = (AxImage2.SizeX - 1) / 2
                    .EndY = 0
                End With
                AxGraphicContext5.BackgroundColor = AxGraphicContext5.BackgroundColor.White
                AxGraphicContext5.ForegroundColor = AxGraphicContext5.BackgroundColor.Black
                AxGraphicContext5.LineSegment()
                With AxGraphicContext5.DrawingRegion
                    .StartX = (AxImage2.SizeX - 1) / 2
                    .StartY = 0
                    .EndX = AxImage2.SizeX - 1
                    .EndY = (AxImage2.SizeY - 1) / 2
                End With
                AxGraphicContext5.BackgroundColor = AxGraphicContext5.BackgroundColor.White
                AxGraphicContext5.ForegroundColor = AxGraphicContext5.BackgroundColor.Black
                AxGraphicContext5.LineSegment()
                With AxGraphicContext5.DrawingRegion
                    .StartX = AxImage2.SizeX - 1
                    .StartY = (AxImage2.SizeY - 1) / 2
                    .EndX = (AxImage2.SizeX - 1) / 2
                    .EndY = AxImage2.SizeY - 1
                End With
                AxGraphicContext5.BackgroundColor = AxGraphicContext5.BackgroundColor.White
                AxGraphicContext5.ForegroundColor = AxGraphicContext5.BackgroundColor.Black
                AxGraphicContext5.LineSegment()
                With AxGraphicContext5.DrawingRegion
                    .StartX = (AxImage2.SizeX - 1) / 2
                    .StartY = AxImage2.SizeY - 1
                    .EndX = 0
                    .EndY = (AxImage2.SizeY - 1) / 2
                End With
                AxGraphicContext5.BackgroundColor = AxGraphicContext5.BackgroundColor.White
                AxGraphicContext5.ForegroundColor = AxGraphicContext5.BackgroundColor.Black
                AxGraphicContext5.LineSegment()
                With AxGraphicContext5.DrawingRegion
                    .StartX = (AxImage2.SizeX - 1) / 2
                    .StartY = (AxImage2.SizeY - 1) / 2
                End With
                'AxGraphicContext5.BackgroundColor = AxGraphicContext5.BackgroundColor.White
                AxGraphicContext5.ForegroundColor = AxGraphicContext5.BackgroundColor.Black
                AxGraphicContext5.Fill()

                ImgNoSQ = 1 'ImgNoSQ + 1
                'If ImgNoSQ <> ImgNoOldSQ Then
                ImgNoOldSQ = ImgNoSQ
                AxImage2.Save("C:\IDS\SRC\DLL Export Device Vision\Fiducial\sq" & ImgNoSQ.ToString & ".bmp")
                Return ImgNoSQ
                'End If
            Catch ex As Exception
                ExceptionDisplay(ex)
            End Try
        ElseIf FiducialMark_form.PictureBox17.BorderStyle = BorderStyle.Fixed3D Then
            Try
                FX_pixel = FX / PixelSizeX  'Convertion of mm --> pixel
                FY_pixel = FY / PixelSizeX
                Dim RemainderX As Integer = FX_pixel Mod 2
                Dim RemainderY As Integer = FY_pixel Mod 2
                If AxImage2.IsAllocated = False Then
                    With AxImage2
                        .SizeX = Convert.ToInt16(FX_pixel)
                        .SizeY = Convert.ToInt16(FY_pixel)
                        .CanProcess = True
                        .CanDisplay = True
                        .Allocate()
                    End With
                Else
                    AxImage2.Free()
                    With AxImage2
                        If RemainderX = 1 Then
                            .SizeX = Convert.ToInt16(FX_pixel)
                        Else
                            .SizeX = Convert.ToInt16(FX_pixel) + 1
                        End If
                        If RemainderY = 1 Then
                            .SizeY = Convert.ToInt16(FY_pixel)
                        Else
                            .SizeY = Convert.ToInt16(FY_pixel) + 1
                        End If
                        .CanProcess = True
                        .CanDisplay = True
                        .Allocate()
                    End With
                    AxImage2.Clear(0)
                End If
                ModelsPM_sizeX = AxImage2.SizeX
                ModelsPM_sizeY = AxImage2.SizeY

                With AxGraphicContext5.DrawingRegion
                    .StartX = 0
                    .StartY = (AxImage2.SizeY - 1) / 2
                    .EndX = (AxImage2.SizeX - 1) / 2
                    .EndY = 0
                End With
                AxGraphicContext5.BackgroundColor = AxGraphicContext5.BackgroundColor.Black
                AxGraphicContext5.ForegroundColor = AxGraphicContext5.BackgroundColor.White
                AxGraphicContext5.LineSegment()
                With AxGraphicContext5.DrawingRegion
                    .StartX = (AxImage2.SizeX - 1) / 2
                    .StartY = 0
                    .EndX = AxImage2.SizeX - 1
                    .EndY = (AxImage2.SizeY - 1) / 2
                End With
                AxGraphicContext5.BackgroundColor = AxGraphicContext5.BackgroundColor.Black
                AxGraphicContext5.ForegroundColor = AxGraphicContext5.BackgroundColor.White
                AxGraphicContext5.LineSegment()
                With AxGraphicContext5.DrawingRegion
                    .StartX = AxImage2.SizeX - 1
                    .StartY = (AxImage2.SizeY - 1) / 2
                    .EndX = (AxImage2.SizeX - 1) / 2
                    .EndY = AxImage2.SizeY - 1
                End With
                AxGraphicContext5.BackgroundColor = AxGraphicContext5.BackgroundColor.Black
                AxGraphicContext5.ForegroundColor = AxGraphicContext5.BackgroundColor.White
                AxGraphicContext5.LineSegment()
                With AxGraphicContext5.DrawingRegion
                    .StartX = (AxImage2.SizeX - 1) / 2
                    .StartY = AxImage2.SizeY - 1
                    .EndX = 0
                    .EndY = (AxImage2.SizeY - 1) / 2
                End With
                AxGraphicContext5.BackgroundColor = AxGraphicContext5.BackgroundColor.Black
                AxGraphicContext5.ForegroundColor = AxGraphicContext5.BackgroundColor.White
                AxGraphicContext5.LineSegment()
                With AxGraphicContext5.DrawingRegion
                    .StartX = (AxImage2.SizeX - 1) / 2
                    .StartY = (AxImage2.SizeY - 1) / 2
                End With
                'AxGraphicContext5.BackgroundColor = AxGraphicContext5.BackgroundColor.White
                AxGraphicContext5.ForegroundColor = AxGraphicContext5.BackgroundColor.White
                AxGraphicContext5.Fill()

                ImgNoSQ = 1 'ImgNoSQ + 1
                'If ImgNoSQ <> ImgNoOldSQ Then
                ImgNoOldSQ = ImgNoSQ
                AxImage2.Save("C:\IDS\SRC\DLL Export Device Vision\Fiducial\sq" & ImgNoSQ.ToString & ".bmp")
                Return ImgNoSQ
                'End If
            Catch ex As Exception
                ExceptionDisplay(ex)
            End Try
        End If
    End Function
    Function FCross(ByVal Fx As Double, ByVal Thickness As Double)
        Dim FX_pixel As Double
        Dim X_mm As Double = Fx
        Thickness = Thickness / PixelSizeX / 2
        FX_pixel = Fx / PixelSizeX  'Convertion of mm --> pixel
        Dim RemainderX As Integer = FX_pixel Mod 2

        If FiducialMark_form.PictureBox13.BorderStyle = BorderStyle.Fixed3D Then
            Try

                If AxImage2.IsAllocated = False Then
                    With AxImage2
                        .SizeX = Convert.ToInt16(FX_pixel)
                        .SizeY = Convert.ToInt16(FX_pixel)
                        .CanProcess = True
                        .CanDisplay = True
                        .Allocate()
                    End With
                Else
                    AxImage2.Free()
                    With AxImage2
                        If RemainderX = 1 Then
                            .SizeX = Convert.ToInt16(FX_pixel)
                        Else
                            .SizeX = Convert.ToInt16(FX_pixel) + 1
                        End If
                        If RemainderX = 1 Then
                            .SizeY = Convert.ToInt16(FX_pixel)
                        Else
                            .SizeY = Convert.ToInt16(FX_pixel) + 1
                        End If
                        .CanProcess = True
                        .CanDisplay = True
                        .Allocate()
                    End With
                    AxImage2.Clear(255)
                End If
                ModelsPM_sizeX = AxImage2.SizeX
                ModelsPM_sizeY = AxImage2.SizeY

                With AxGraphicContext5.DrawingRegion
                    If RemainderX = 1 Then
                        .StartX = 0
                    Else
                        .StartX = 0
                    End If
                    If RemainderX = 1 Then
                        .StartY = 0
                    Else
                        .StartY = 0
                    End If
                    If RemainderX = 1 Then
                        .EndX = AxImage2.SizeX / 2 - Thickness - 1
                    Else
                        .EndX = AxImage2.SizeX / 2 - Thickness - 1
                    End If
                    If RemainderX = 1 Then
                        .EndY = AxImage2.SizeX / 2 - Thickness - 1
                    Else
                        .EndY = AxImage2.SizeX / 2 - Thickness - 1
                    End If
                End With
                AxGraphicContext5.BackgroundColor = AxGraphicContext5.BackgroundColor.White
                AxGraphicContext5.ForegroundColor = AxGraphicContext5.BackgroundColor.Black
                AxGraphicContext5.Rectangle(True, 0)

                With AxGraphicContext5.DrawingRegion
                    If RemainderX = 1 Then
                        .StartX = AxImage2.SizeX / 2 + Thickness
                    Else
                        .StartX = AxImage2.SizeX / 2 + Thickness
                    End If
                    If RemainderX = 1 Then
                        .StartY = 0
                    Else
                        .StartY = 0
                    End If
                    If RemainderX = 1 Then
                        .EndX = AxImage2.SizeX
                    Else
                        .EndX = AxImage2.SizeX
                    End If
                    If RemainderX = 1 Then
                        .EndY = AxImage2.SizeX / 2 - Thickness - 1
                    Else
                        .EndY = AxImage2.SizeX / 2 - Thickness - 1
                    End If
                End With
                AxGraphicContext5.BackgroundColor = AxGraphicContext5.BackgroundColor.White
                AxGraphicContext5.ForegroundColor = AxGraphicContext5.BackgroundColor.Black
                AxGraphicContext5.Rectangle(True, 0)

                With AxGraphicContext5.DrawingRegion
                    If RemainderX = 1 Then
                        .StartX = 0
                    Else
                        .StartX = 0
                    End If
                    If RemainderX = 1 Then
                        .StartY = AxImage2.SizeX / 2 + Thickness
                    Else
                        .StartY = AxImage2.SizeX / 2 + Thickness
                    End If
                    If RemainderX = 1 Then
                        .EndX = AxImage2.SizeX / 2 - Thickness - 1
                    Else
                        .EndX = AxImage2.SizeX / 2 - Thickness - 1
                    End If
                    If RemainderX = 1 Then
                        .EndY = AxImage2.SizeX
                    Else
                        .EndY = AxImage2.SizeX
                    End If
                End With
                AxGraphicContext5.BackgroundColor = AxGraphicContext5.BackgroundColor.White
                AxGraphicContext5.ForegroundColor = AxGraphicContext5.BackgroundColor.Black
                AxGraphicContext5.Rectangle(True, 0)

                With AxGraphicContext5.DrawingRegion
                    If RemainderX = 1 Then
                        .StartX = AxImage2.SizeX / 2 + Thickness
                    Else
                        .StartX = AxImage2.SizeX / 2 + Thickness
                    End If
                    If RemainderX = 1 Then
                        .StartY = AxImage2.SizeX / 2 + Thickness
                    Else
                        .StartY = AxImage2.SizeX / 2 + Thickness
                    End If
                    If RemainderX = 1 Then
                        .EndX = AxImage2.SizeX
                    Else
                        .EndX = AxImage2.SizeX
                    End If
                    If RemainderX = 1 Then
                        .EndY = AxImage2.SizeX
                    Else
                        .EndY = AxImage2.SizeX
                    End If
                End With
                AxGraphicContext5.BackgroundColor = AxGraphicContext5.BackgroundColor.White
                AxGraphicContext5.ForegroundColor = AxGraphicContext5.BackgroundColor.Black
                AxGraphicContext5.Rectangle(True, 0)

                ImgNoSQ = 1 'ImgNoSQ + 1
                'If ImgNoSQ <> ImgNoOldSQ Then
                ImgNoOldSQ = ImgNoSQ
                AxImage2.Save("C:\IDS\SRC\DLL Export Device Vision\Fiducial\sq" & ImgNoSQ.ToString & ".bmp")
                Return ImgNoSQ
                'End If
            Catch ex As Exception
                ExceptionDisplay(ex)
            End Try

        ElseIf FiducialMark_form.PictureBox3.BorderStyle = BorderStyle.Fixed3D Then
            Try
                If AxImage2.IsAllocated = False Then
                    With AxImage2
                        .SizeX = Convert.ToInt16(FX_pixel)
                        .SizeY = Convert.ToInt16(FX_pixel)
                        .CanProcess = True
                        .CanDisplay = True
                        .Allocate()
                    End With
                Else
                    AxImage2.Free()
                    With AxImage2
                        If RemainderX = 1 Then
                            .SizeX = Convert.ToInt16(FX_pixel)
                        Else
                            .SizeX = Convert.ToInt16(FX_pixel) + 1
                        End If
                        If RemainderX = 1 Then
                            .SizeY = Convert.ToInt16(FX_pixel)
                        Else
                            .SizeY = Convert.ToInt16(FX_pixel) + 1
                        End If
                        .CanProcess = True
                        .CanDisplay = True
                        .Allocate()
                    End With
                    AxImage2.Clear(0)
                End If
                ModelsPM_sizeX = AxImage2.SizeX
                ModelsPM_sizeY = AxImage2.SizeY

                With AxGraphicContext5.DrawingRegion
                    If RemainderX = 1 Then
                        .StartX = 0
                    Else
                        .StartX = 0
                    End If
                    If RemainderX = 1 Then
                        .StartY = 0
                    Else
                        .StartY = 0
                    End If
                    If RemainderX = 1 Then
                        .EndX = AxImage2.SizeX / 2 - Thickness - 1
                    Else
                        .EndX = AxImage2.SizeX / 2 - Thickness - 1
                    End If
                    If RemainderX = 1 Then
                        .EndY = AxImage2.SizeX / 2 - Thickness - 1
                    Else
                        .EndY = AxImage2.SizeX / 2 - Thickness - 1
                    End If
                End With
                AxGraphicContext5.BackgroundColor = AxGraphicContext5.BackgroundColor.Black
                AxGraphicContext5.ForegroundColor = AxGraphicContext5.BackgroundColor.White
                AxGraphicContext5.Rectangle(True, 0)

                With AxGraphicContext5.DrawingRegion
                    If RemainderX = 1 Then
                        .StartX = AxImage2.SizeX / 2 + Thickness
                    Else
                        .StartX = AxImage2.SizeX / 2 + Thickness
                    End If
                    If RemainderX = 1 Then
                        .StartY = 0
                    Else
                        .StartY = 0
                    End If
                    If RemainderX = 1 Then
                        .EndX = AxImage2.SizeX
                    Else
                        .EndX = AxImage2.SizeX
                    End If
                    If RemainderX = 1 Then
                        .EndY = AxImage2.SizeX / 2 - Thickness - 1
                    Else
                        .EndY = AxImage2.SizeX / 2 - Thickness - 1
                    End If
                End With
                AxGraphicContext5.BackgroundColor = AxGraphicContext5.BackgroundColor.Black
                AxGraphicContext5.ForegroundColor = AxGraphicContext5.BackgroundColor.White
                AxGraphicContext5.Rectangle(True, 0)

                With AxGraphicContext5.DrawingRegion
                    If RemainderX = 1 Then
                        .StartX = 0
                    Else
                        .StartX = 0
                    End If
                    If RemainderX = 1 Then
                        .StartY = AxImage2.SizeX / 2 + Thickness
                    Else
                        .StartY = AxImage2.SizeX / 2 + Thickness
                    End If
                    If RemainderX = 1 Then
                        .EndX = AxImage2.SizeX / 2 - Thickness - 1
                    Else
                        .EndX = AxImage2.SizeX / 2 - Thickness - 1
                    End If
                    If RemainderX = 1 Then
                        .EndY = AxImage2.SizeX
                    Else
                        .EndY = AxImage2.SizeX
                    End If
                End With
                AxGraphicContext5.BackgroundColor = AxGraphicContext5.BackgroundColor.Black
                AxGraphicContext5.ForegroundColor = AxGraphicContext5.BackgroundColor.White
                AxGraphicContext5.Rectangle(True, 0)

                With AxGraphicContext5.DrawingRegion
                    If RemainderX = 1 Then
                        .StartX = AxImage2.SizeX / 2 + Thickness
                    Else
                        .StartX = AxImage2.SizeX / 2 + Thickness
                    End If
                    If RemainderX = 1 Then
                        .StartY = AxImage2.SizeX / 2 + Thickness
                    Else
                        .StartY = AxImage2.SizeX / 2 + Thickness
                    End If
                    If RemainderX = 1 Then
                        .EndX = AxImage2.SizeX
                    Else
                        .EndX = AxImage2.SizeX
                    End If
                    If RemainderX = 1 Then
                        .EndY = AxImage2.SizeX
                    Else
                        .EndY = AxImage2.SizeX
                    End If
                End With
                AxGraphicContext5.BackgroundColor = AxGraphicContext5.BackgroundColor.Black
                AxGraphicContext5.ForegroundColor = AxGraphicContext5.BackgroundColor.White
                AxGraphicContext5.Rectangle(True, 0)

                ImgNoSQ = 1 'ImgNoSQ + 1
                'If ImgNoSQ <> ImgNoOldSQ Then
                ImgNoOldSQ = ImgNoSQ
                AxImage2.Save("C:\IDS\SRC\DLL Export Device Vision\Fiducial\sq" & ImgNoSQ.ToString & ".bmp")
                Return ImgNoSQ
                'End If
            Catch ex As Exception
                ExceptionDisplay(ex)
            End Try
        End If
    End Function
    Function FRing(ByVal FOutDiameter As Double, ByVal FInDiameter As Double) As Boolean
        Dim OuterDiameter_pixel As Double = 1
        Dim InnerDiameter_pixel As Double = 1
        PixelSize = PixelSizeX
        Dim OuterDiameter_mm As Double = FOutDiameter
        OuterDiameter_pixel = OuterDiameter_mm / PixelSizeX  'Convertion of mm --> pixel
        Dim InnerDiameter_mm As Double = FInDiameter
        InnerDiameter_pixel = InnerDiameter_mm / PixelSizeX  'Convertion of mm --> pixel

        If FiducialMark_form.PictureBox2.BorderStyle = BorderStyle.Fixed3D Then
            Dim OutRemainder As Integer = OuterDiameter_pixel Mod 2
            Dim InRemainder As Integer = InnerDiameter_pixel Mod 2
            If AxImage2.IsAllocated = False Then
                With AxImage2
                    .SizeX = Convert.ToInt16(OuterDiameter_pixel)
                    .SizeY = Convert.ToInt16(OuterDiameter_pixel)
                    .CanProcess = True
                    .CanDisplay = True
                    .Allocate()
                End With
            Else
                AxImage2.Free()

                With AxImage2
                    If OutRemainder = 1 Then
                        .SizeX = Convert.ToInt16(OuterDiameter_pixel)
                        .SizeY = Convert.ToInt16(OuterDiameter_pixel)
                    Else
                        .SizeX = Convert.ToInt16(OuterDiameter_pixel) + 1 'to centralize the circle
                        .SizeY = Convert.ToInt16(OuterDiameter_pixel) + 1
                    End If
                    .CanProcess = True
                    .CanDisplay = True
                    .Allocate()
                End With
                AxImage2.Clear(255) '255 for b/w, 0 for w/b
            End If
            ModelsPM_sizeX = AxImage2.SizeX
            ModelsPM_sizeY = AxImage2.SizeY

            With AxGraphicContext5.DrawingRegion
                .StartX = 0
                .StartY = 0
                If OutRemainder = 1 Then
                    .EndX = Convert.ToInt16(OuterDiameter_pixel)
                    .EndY = Convert.ToInt16(OuterDiameter_pixel)
                Else
                    .EndX = Convert.ToInt16(OuterDiameter_pixel) + 1
                    .EndY = Convert.ToInt16(OuterDiameter_pixel) + 1
                End If
            End With
            AxGraphicContext5.BackgroundColor = AxGraphicContext5.BackgroundColor.White
            AxGraphicContext5.ForegroundColor = AxGraphicContext5.BackgroundColor.Black
            AxGraphicContext5.Arc(0, 360, True)
            With AxGraphicContext5.DrawingRegion
                If OutRemainder = 1 Then
                    If InRemainder = 1 Then
                        .StartX = Convert.ToInt16(OuterDiameter_pixel) / 2 - InnerDiameter_pixel / 2
                        .StartY = Convert.ToInt16(OuterDiameter_pixel) / 2 - InnerDiameter_pixel / 2
                    Else
                        .StartX = Convert.ToInt16(OuterDiameter_pixel) / 2 - (InnerDiameter_pixel + 1) / 2
                        .StartY = Convert.ToInt16(OuterDiameter_pixel) / 2 - (InnerDiameter_pixel + 1) / 2
                    End If
                Else
                    If InRemainder = 1 Then
                        .StartX = (Convert.ToInt16(OuterDiameter_pixel) + 1) / 2 - InnerDiameter_pixel / 2
                        .StartY = (Convert.ToInt16(OuterDiameter_pixel) + 1) / 2 - InnerDiameter_pixel / 2
                    Else
                        .StartX = (Convert.ToInt16(OuterDiameter_pixel) + 1) / 2 - (InnerDiameter_pixel + 1) / 2
                        .StartY = (Convert.ToInt16(OuterDiameter_pixel) + 1) / 2 - (InnerDiameter_pixel + 1) / 2
                    End If
                End If
                If OutRemainder = 1 Then
                    If InRemainder = 1 Then
                        .EndX = (Convert.ToInt16(OuterDiameter_pixel)) / 2 + Convert.ToInt16(InnerDiameter_pixel) / 2
                        .EndY = (Convert.ToInt16(OuterDiameter_pixel)) / 2 + Convert.ToInt16(InnerDiameter_pixel) / 2
                    Else
                        .EndX = (Convert.ToInt16(OuterDiameter_pixel)) / 2 + Convert.ToInt16((InnerDiameter_pixel + 1)) / 2
                        .EndY = (Convert.ToInt16(OuterDiameter_pixel)) / 2 + Convert.ToInt16((InnerDiameter_pixel + 1)) / 2
                    End If
                Else
                    If InRemainder = 1 Then
                        .EndX = (Convert.ToInt16(OuterDiameter_pixel) + 1) / 2 + Convert.ToInt16(InnerDiameter_pixel) / 2
                        .EndY = (Convert.ToInt16(OuterDiameter_pixel) + 1) / 2 + Convert.ToInt16(InnerDiameter_pixel) / 2
                    Else
                        .EndX = (Convert.ToInt16(OuterDiameter_pixel) + 1) / 2 + Convert.ToInt16((InnerDiameter_pixel + 1)) / 2
                        .EndY = (Convert.ToInt16(OuterDiameter_pixel) + 1) / 2 + Convert.ToInt16((InnerDiameter_pixel + 1)) / 2
                    End If
                End If
            End With
            AxGraphicContext5.BackgroundColor = AxGraphicContext5.BackgroundColor.Black
            AxGraphicContext5.ForegroundColor = AxGraphicContext5.BackgroundColor.White
            AxGraphicContext5.Arc(0, 360, True)

            '***********************************************************************
            AxImage2.Save("C:\IDS\SRC\DLL Export Device Vision\Fiducial\" + "temp.bmp")
            Return True
        ElseIf FiducialMark_form.PictureBox12.BorderStyle = BorderStyle.Fixed3D Then
            Dim OutRemainder As Integer = OuterDiameter_pixel Mod 2
            Dim InRemainder As Integer = InnerDiameter_pixel Mod 2
            If AxImage2.IsAllocated = False Then
                With AxImage2
                    .SizeX = Convert.ToInt16(OuterDiameter_pixel)
                    .SizeY = Convert.ToInt16(OuterDiameter_pixel)
                    .CanProcess = True
                    .CanDisplay = True
                    .Allocate()
                End With
            Else
                AxImage2.Free()

                With AxImage2
                    If OutRemainder = 1 Then
                        .SizeX = Convert.ToInt16(OuterDiameter_pixel)
                        .SizeY = Convert.ToInt16(OuterDiameter_pixel)
                    Else
                        .SizeX = Convert.ToInt16(OuterDiameter_pixel) + 1 'to centralize the circle
                        .SizeY = Convert.ToInt16(OuterDiameter_pixel) + 1
                    End If
                    .CanProcess = True
                    .CanDisplay = True
                    .Allocate()
                End With
                AxImage2.Clear(0) '255 for b/w, 0 for w/b
            End If
            ModelsPM_sizeX = AxImage2.SizeX
            ModelsPM_sizeY = AxImage2.SizeY

            With AxGraphicContext5.DrawingRegion
                .StartX = 0
                .StartY = 0
                If OutRemainder = 1 Then
                    .EndX = Convert.ToInt16(OuterDiameter_pixel)
                    .EndY = Convert.ToInt16(OuterDiameter_pixel)
                Else
                    .EndX = Convert.ToInt16(OuterDiameter_pixel) + 1
                    .EndY = Convert.ToInt16(OuterDiameter_pixel) + 1
                End If
            End With
            AxGraphicContext5.BackgroundColor = AxGraphicContext5.BackgroundColor.Black
            AxGraphicContext5.ForegroundColor = AxGraphicContext5.BackgroundColor.White
            AxGraphicContext5.Arc(0, 360, True)
            With AxGraphicContext5.DrawingRegion
                If OutRemainder = 1 Then
                    If InRemainder = 1 Then
                        .StartX = Convert.ToInt16(OuterDiameter_pixel) / 2 - InnerDiameter_pixel / 2
                        .StartY = Convert.ToInt16(OuterDiameter_pixel) / 2 - InnerDiameter_pixel / 2
                    Else
                        .StartX = Convert.ToInt16(OuterDiameter_pixel) / 2 - (InnerDiameter_pixel + 1) / 2
                        .StartY = Convert.ToInt16(OuterDiameter_pixel) / 2 - (InnerDiameter_pixel + 1) / 2
                    End If
                Else
                    If InRemainder = 1 Then
                        .StartX = (Convert.ToInt16(OuterDiameter_pixel) + 1) / 2 - InnerDiameter_pixel / 2
                        .StartY = (Convert.ToInt16(OuterDiameter_pixel) + 1) / 2 - InnerDiameter_pixel / 2
                    Else
                        .StartX = (Convert.ToInt16(OuterDiameter_pixel) + 1) / 2 - (InnerDiameter_pixel + 1) / 2
                        .StartY = (Convert.ToInt16(OuterDiameter_pixel) + 1) / 2 - (InnerDiameter_pixel + 1) / 2
                    End If
                End If
                If OutRemainder = 1 Then
                    If InRemainder = 1 Then
                        .EndX = (Convert.ToInt16(OuterDiameter_pixel)) / 2 + Convert.ToInt16(InnerDiameter_pixel) / 2
                        .EndY = (Convert.ToInt16(OuterDiameter_pixel)) / 2 + Convert.ToInt16(InnerDiameter_pixel) / 2
                    Else
                        .EndX = (Convert.ToInt16(OuterDiameter_pixel)) / 2 + Convert.ToInt16((InnerDiameter_pixel + 1)) / 2
                        .EndY = (Convert.ToInt16(OuterDiameter_pixel)) / 2 + Convert.ToInt16((InnerDiameter_pixel + 1)) / 2
                    End If
                Else
                    If InRemainder = 1 Then
                        .EndX = (Convert.ToInt16(OuterDiameter_pixel) + 1) / 2 + Convert.ToInt16(InnerDiameter_pixel) / 2
                        .EndY = (Convert.ToInt16(OuterDiameter_pixel) + 1) / 2 + Convert.ToInt16(InnerDiameter_pixel) / 2
                    Else
                        .EndX = (Convert.ToInt16(OuterDiameter_pixel) + 1) / 2 + Convert.ToInt16((InnerDiameter_pixel + 1)) / 2
                        .EndY = (Convert.ToInt16(OuterDiameter_pixel) + 1) / 2 + Convert.ToInt16((InnerDiameter_pixel + 1)) / 2
                    End If
                End If
            End With
            AxGraphicContext5.BackgroundColor = AxGraphicContext5.BackgroundColor.White
            AxGraphicContext5.ForegroundColor = AxGraphicContext5.BackgroundColor.Black
            AxGraphicContext5.Arc(0, 360, True)

            '***********************************************************************
            AxImage2.Save("C:\IDS\SRC\DLL Export Device Vision\Fiducial\" + "temp.bmp")
            Return True
        End If
    End Function
#End Region
    Dim PositX, PositY, Rotat As Double
    Shared BinarizeThresholdNo As Integer = 128
    Function BinarizeThreshold(ByVal BinarizeThresholdNo)
        BinarizeThresholdNo = BinarizeThresholdNo
    End Function
    Function Test_Fiducial(ByVal FID As Integer, ByRef FI_OffX As Double, ByRef FI_OffY As Double) As Boolean
        FI_OffX = 0
        FI_OffY = 0
        PositX = 0
        PositY = 0
        Rotat = 0

        If FID <> 10 And FID <> 20 And FID <> 21 Then
            For i As Integer = 1 To 5
                Try
                    AxPatternMatching2.FindModel(IndexNo)
                    If AxPatternMatching2.Results.Count > 0 Then
                        'Score()
                        PositX = PositX + AxPatternMatching2.Results.Item(IndexNo).PositionX
                        PositY = PositY + AxPatternMatching2.Results.Item(IndexNo).PositionY
                        Rotat = Rotat + AxPatternMatching2.Results.Item(IndexNo).Angle
                        Dim del_x As Double = AxPatternMatching2.Results.Item(1).PositionX - DisplayCenterXPosition
                        Dim del_y As Double = AxPatternMatching2.Results.Item(1).PositionY - DisplayCenterYPosition
                        FiducialMark_form.TextBox_RPosX.Text = CvrtDoubleToString(del_x * PixelSizeX)
                        FiducialMark_form.TextBox_RPosY.Text = CvrtDoubleToString(del_y * PixelSizeY)
                        AxPatternMatching2.Results.Item(IndexNo).Draw(AxDisplay1.OverlayImage, Matrox.ActiveMIL.PatternMatching.PatDrawOperationConstants.patDrawPosition + Matrox.ActiveMIL.PatternMatching.PatDrawOperationConstants.patDrawBox)
                    Else
                        FiducialMark_form.TextBox_RPosX.Text = 0
                        FiducialMark_form.TextBox_RPosY.Text = 0
                    End If
                Catch ex As Exception
                    ExceptionDisplay(ex)
                End Try
            Next i
        ElseIf FID = 10 Then 'no binarized
            For i As Integer = 1 To 5
                Try
                    AxPatternMatching2.FindModel(IndexNo)
                    If AxPatternMatching2.Results.Count > 0 Then
                        'Score()
                        PositX = PositX + AxPatternMatching2.Results.Item(IndexNo).PositionX
                        PositY = PositY + AxPatternMatching2.Results.Item(IndexNo).PositionY
                        Rotat = Rotat + AxPatternMatching2.Results.Item(IndexNo).Angle
                        Dim del_x As Double = AxPatternMatching2.Results.Item(1).PositionX - DisplayCenterXPosition
                        Dim del_y As Double = AxPatternMatching2.Results.Item(1).PositionY - DisplayCenterYPosition
                        FiducialMark_form.TextBox_RPosX.Text = CvrtDoubleToString(del_x * PixelSizeX)
                        FiducialMark_form.TextBox_RPosY.Text = CvrtDoubleToString(del_y * PixelSizeY)
                        AxPatternMatching2.Results.Item(IndexNo).Draw(AxDisplay1.OverlayImage, Matrox.ActiveMIL.PatternMatching.PatDrawOperationConstants.patDrawPosition + Matrox.ActiveMIL.PatternMatching.PatDrawOperationConstants.patDrawBox)
                    Else
                        FiducialMark_form.TextBox_RPosX.Text = 0
                        FiducialMark_form.TextBox_RPosY.Text = 0
                    End If
                Catch ex As Exception
                    ExceptionDisplay(ex)
                End Try
            Next i
        ElseIf FID = 20 Then 'For FileLoad in Programming
            For i As Integer = 1 To 5
                Try
                    AxPatternMatching2.FindModel(1)
                    If AxPatternMatching2.Results.Count > 0 Then
                        'Score()
                        PositX = PositX + AxPatternMatching2.Results.Item(1).PositionX
                        PositY = PositY + AxPatternMatching2.Results.Item(1).PositionY
                        Rotat = Rotat + AxPatternMatching2.Results.Item(1).Angle
                        Dim del_x As Double = AxPatternMatching2.Results.Item(1).PositionX - DisplayCenterXPosition
                        Dim del_y As Double = AxPatternMatching2.Results.Item(1).PositionY - DisplayCenterYPosition
                        FiducialMark_form.TextBox_RPosX.Text = CvrtDoubleToString(del_x * PixelSizeX)
                        FiducialMark_form.TextBox_RPosY.Text = CvrtDoubleToString(del_y * PixelSizeY)
                        AxPatternMatching2.Results.Item(1).Draw(AxDisplay1.OverlayImage, Matrox.ActiveMIL.PatternMatching.PatDrawOperationConstants.patDrawPosition + Matrox.ActiveMIL.PatternMatching.PatDrawOperationConstants.patDrawBox)
                        FiducialMark_form.FiducialOutput(del_x * PixelSizeX, del_y * PixelSizeY)
                    Else
                        FiducialMark_form.TextBox_RPosX.Text = 0
                        FiducialMark_form.TextBox_RPosY.Text = 0
                        MsgBox("Failed")
                        'FiducialMark_form.RemoveImage20()
                    End If
                Catch ex As Exception
                    ExceptionDisplay(ex)
                End Try
            Next i
            Score(0)
        ElseIf FID = 21 Then 'For Production
            For i As Integer = 1 To 5
                Try
                    Dim test As Integer = AxPatternMatching2.Models.Count
                    AxPatternMatching2.FindModel(1)
                    If AxPatternMatching2.Results.Count > 0 Then
                        'Score()
                        AxImage3.Save("C:\fiducial-succeed.bmp")
                        PositX = PositX + AxPatternMatching2.Results.Item(1).PositionX
                        PositY = PositY + AxPatternMatching2.Results.Item(1).PositionY
                        Rotat = Rotat + AxPatternMatching2.Results.Item(1).Angle
                        Dim del_x As Double = AxPatternMatching2.Results.Item(1).PositionX - DisplayCenterXPosition
                        Dim del_y As Double = AxPatternMatching2.Results.Item(1).PositionY - DisplayCenterYPosition
                        FiducialMark_form.TextBox_RPosX.Text = del_x * PixelSizeX
                        FiducialMark_form.TextBox_RPosY.Text = del_y * PixelSizeY
                        AxPatternMatching2.Results.Item(1).Draw(AxDisplay1.OverlayImage, Matrox.ActiveMIL.PatternMatching.PatDrawOperationConstants.patDrawPosition + Matrox.ActiveMIL.PatternMatching.PatDrawOperationConstants.patDrawBox)
                        FI_OffX = del_x * PixelSizeX
                        FI_OffY = del_y * PixelSizeY
                        Return True
                    Else
                        WaitForImageToStabilize()
                    End If
                    If i = 5 Then
                        AxImage3.Save("C:\fiducial-fail.bmp")
                        FiducialMark_form.TextBox_RPosX.Text = 0
                        FiducialMark_form.TextBox_RPosY.Text = 0
                        FI_OffX = 0
                        FI_OffY = 0
                        Return False
                    End If
                Catch ex As Exception
                    ExceptionDisplay(ex)
                End Try
            Next i
            'Score(0)
        End If

    End Function 'called by FiducialMark_Form
    Public Function Score(ByVal PM As Integer) As String
        If AxPatternMatching2.Results.Count > 0 Then
            If PM = 0 Then
                Dim WriterPM As New StreamWriter("c:\ids\src\DLL Export Device Vision\Test\PM.xls", True, System.Text.Encoding.Default)
                'WriterPM.WriteLine(AxPatternMatching2.Results.Item(IndexNo).PositionX.ToString & "	" & AxPatternMatching2.Results.Item(IndexNo).PositionY.ToString & "	" & AxPatternMatching2.Results.Item(IndexNo).Angle.ToString)
                WriterPM.WriteLine((PositX / 5).ToString & "	" & (PositY / 5).ToString & "	" & (Rotat / 5).ToString)
                WriterPM.Close()
            End If
            Return CvrtDoubleToString(AxPatternMatching2.Results.Item(1).Score())
        End If
    End Function 'called by FiducialMark_Form
    Sub SaveFiducial(ByVal Fiducial_no As Integer, ByVal Fid As Integer)
        If Fid <> 20 Then
            If IndexNo = 1 Then 'Got model
                If Fiducial_no = 1 Then
                    AxPatternMatching2.Models.Item(IndexNo).Save(Folder & Fiducial_no.ToString & ".mmo")
                    FiducialMark_form.PictureBox20.Image.Save(Folder & Fiducial_no.ToString & ".bmp")
                Else
                    AxPatternMatching2.Models.Item(IndexNo).Save(Folder & Fiducial_no.ToString & ".mmo")
                    FiducialMark_form.PictureBox20.Image.Save(Folder & Fiducial_no.ToString & ".bmp")
                End If
            End If
        Else
            If Fiducial_no = 1 Then
                AxPatternMatching2.Models.Item(1).Save(Folder & Fiducial_no.ToString & ".mmo")
            Else
                AxPatternMatching2.Models.Item(1).Save(Folder & Fiducial_no.ToString & ".mmo")
            End If
        End If
    End Sub
    Function LoadFiducial(ByVal Filename As String) As Boolean
        'Dim LoadFile As New OpenFileDialogPreview.Form1
        'Dim PathRaw As String = LoadFile.Loadfile()
        Dim PathRaw As String = Filename
        If PathRaw = Nothing Then
            Return False
        Else
            Dim Path As String = PathRaw.Remove(PathRaw.LastIndexOf(".") + 1, PathRaw.Length - PathRaw.LastIndexOf(".") - 1)
            Dim FPath As String = Path & "mmo"
            If CheckFileName(FPath) = True Then 'To check file existence
                If AxPatternMatching2.Models.Count <> 0 Then 'To remove all the models
                    For i As Integer = AxPatternMatching2.Models.Count To 1 Step -1
                        AxPatternMatching2.Models.Remove(i)
                    Next
                End If
                AxPatternMatching2.Models.Load(FPath)
                'SJ
                'FiducialMark_form.PictureBox20.Image = Image.FromFile(PathRaw)
            Else
                MsgBox("Selected file doesn't exist or lost of .mmo file. Please create a new Fiducial Point")
                Return False
            End If
            Return True
        End If
    End Function
    Function LoadFiducial() As Boolean
        Dim LoadFile As New OpenFileDialogPreview.Form1
        FiducialMark_form.TopMost = False
        LoadFile.TopMost = True
        Dim PathRaw As String = LoadFile.Loadfile()
        LoadFile.TopMost = False
        FiducialMark_form.TopMost = True
        'Dim PathRaw As String = Filename
        If PathRaw = Nothing Then
            Return False
        Else
            Dim Path As String = PathRaw.Remove(PathRaw.LastIndexOf(".") + 1, PathRaw.Length - PathRaw.LastIndexOf(".") - 1)
            Dim FPath As String = Path & "mmo"
            If CheckFileName(FPath) = True Then 'To check file existence
                If AxPatternMatching2.Models.Count <> 0 Then 'To remove all the models
                    For i As Integer = AxPatternMatching2.Models.Count To 1 Step -1
                        AxPatternMatching2.Models.Remove(i)
                    Next
                End If
                AxPatternMatching2.Models.Load(FPath)
                FiducialMark_form.PictureBox20.Image() = Image.FromFile(PathRaw)
            Else
                MsgBox("Selected file doesn't exist or lost of .mmo file. Please create a new Fiducial Point")
                Return False
            End If
            Return True
        End If
    End Function
#End Region
#Region "Calibration"
    Dim NoOfColumns As Integer = 10
    Dim NoOfRows As Integer = 8
    Dim Spacing As Double '= 1.5 * 1000
    Dim PixelSizeX As Double = 0.0210769251512467
    Dim PixelSizeY As Double = 0.0211540014212895
    Dim PixelRatio As Double
    Function PixelXSize() As Double
        Return PixelSizeX
    End Function
    Function PixelYSize() As Double
        Return PixelSizeY
    End Function
    Function Pixel_Ratio() As Double
        Return PixelRatio
    End Function
    Function CameraCalibration(ByVal ColumnsNo As Integer, ByVal RowsNo As Integer, ByVal PitchSpacing As Double) As Integer
        Try
            NoOfColumns = ColumnsNo
            NoOfRows = RowsNo
            Spacing = PitchSpacing
            AxCalibration1.AutomaticAllocation = True
            If AxCalibration1.IsAllocated = True Then
                'AxCalibration1.Free()
                'AxCalibration1.CalibrationMode() = MILCALIBRATIONLib.CalCalibrationModeConstants.calCalibrationModeDefault
                AxCalibration1.CalibrationPoints.RemoveAll()
                AxCalibration1.RelativeOrigin.Reset()
                AxCalibration1.Grid.NumberOfColumns = NoOfColumns
                AxCalibration1.Grid.NumberOfRows = NoOfRows
                AxCalibration1.Grid.ColumnSpacing = Spacing
                AxCalibration1.Grid.RowSpacing = Spacing
                AxCalibration1.Grid.OriginX = 0
                AxCalibration1.Grid.OriginY = 0
                AxCalibration1.CameraPosition.X = 0
                AxCalibration1.CameraPosition.Y = 0
                AxCalibration1.RelativeOrigin.X = 0
                AxCalibration1.RelativeOrigin.Y = 0
                AxCalibration1.RelativeOrigin.Angle = 0
                AxCalibration1.OutputCoordinateSystem = Matrox.ActiveMIL.Calibration.CalCoordinateSystemConstants.calPixel
                AxCalibration1.TransformCacheEnabled = True
            End If
            AxCalibration1.CalibrateGrid(AxDisplay1.Image)
            'AxCalibration1.CalibrateGrid(AxImage3)
            If AxCalibration1.IsCalibrated = True Then
                PixelSizeX = AxCalibration1.Results.PixelSizeX
                PixelSizeY = AxCalibration1.Results.PixelSizeY
                PixelRatio = AxCalibration1.Results.PixelAspectRatio
            Else
                MsgBox("Calibration Failed!")
            End If
        Catch ex As Exception
            MsgBox("error " + ex.Message.ToString)
        End Try
    End Function
    Function CameraCalibration_Test(ByVal ColumnsNo As Integer, ByVal RowsNo As Integer, ByVal PitchSpacing As Double) As Integer
        Try
            NoOfColumns = ColumnsNo
            NoOfRows = RowsNo
            Spacing = PitchSpacing
            AxCalibration1.AutomaticAllocation = True
            If AxCalibration1.IsAllocated = True Then
                'AxCalibration1.Free()
                'AxCalibration1.CalibrationMode() = MILCALIBRATIONLib.CalCalibrationModeConstants.calCalibrationModeDefault
                AxCalibration1.CalibrationPoints.RemoveAll()
                AxCalibration1.RelativeOrigin.Reset()
                AxCalibration1.Grid.NumberOfColumns = NoOfColumns
                AxCalibration1.Grid.NumberOfRows = NoOfRows
                AxCalibration1.Grid.ColumnSpacing = Spacing
                AxCalibration1.Grid.RowSpacing = Spacing
                AxCalibration1.Grid.OriginX = 0
                AxCalibration1.Grid.OriginY = 0
                AxCalibration1.CameraPosition.X = 0
                AxCalibration1.CameraPosition.Y = 0
                AxCalibration1.RelativeOrigin.X = 0
                AxCalibration1.RelativeOrigin.Y = 0
                AxCalibration1.RelativeOrigin.Angle = 0
                AxCalibration1.OutputCoordinateSystem = Matrox.ActiveMIL.Calibration.CalCoordinateSystemConstants.calPixel
                AxCalibration1.TransformCacheEnabled = True
            End If
            AxCalibration1.CalibrateGrid(AxDisplay1.Image)
            'AxCalibration1.CalibrateGrid(AxImage3)
            If AxCalibration1.IsCalibrated = True Then
                PixelSizeX = AxCalibration1.Results.PixelSizeX
                PixelSizeY = AxCalibration1.Results.PixelSizeY
                PixelRatio = AxCalibration1.Results.PixelAspectRatio
            Else
                MsgBox("Calibration Failed!")
            End If
        Catch ex As Exception
            MsgBox("error " + ex.Message.ToString)
        End Try
    End Function
#End Region
#Region "RejectPoint"
    Public RejectPoint_flag As Boolean = False

    Function RejectPoints()
        If RejectPoint_flag = True Then
            RejectPoint_form.Show()
            RejectPoint_form.Location = New Point(0, 50)
            RejectPoint_form.Button_Next.Enabled = True
            RejectPoint_form.Button_Load.Enabled = True
            'RejectPoint_flag = False
        End If
        If RejectPoint_form.Visible = False Then
            RejectPoint_form.Visible = True
        End If
    End Function
    Function RejectPointROI()
        Dim HistogramRejectPoint(255) As Integer
        Dim min, max As Integer
        If AxImage8.IsAllocated = True Then
            AxImage8.Free()
            'Else
            '   AxImage8.AutomaticAllocation = True
        End If

        With AxImage8.ChildRegion
            .OffsetX = DisplayCenterXPosition - MROIx / 2
            .OffsetY = DisplayCenterYPosition - MROIy / 2
            .SizeX = MROIx
            .SizeY = MROIy
        End With
        AxImage8.Allocate()
        'AxImage8.ChildRegion.Move(DisplayCenterXPosition - MROIx / 2, DisplayCenterYPosition - MROIy / 2, MROIx, MROIy)

        RejectPoint_form.GetVariables(DisplayCenterXPosition, DisplayCenterYPosition, MROIx, MROIy)
        AxImageProcessing4.Binarize(Matrox.ActiveMIL.ImageProcessing.ImpConditionConstants.impGreaterOrEqualTo, RejectPoint_form.ValueBinarizedThreshold.Value)
        'If RejectPoint_form.PictureBox1.Image Is Nothing Then
        'Else
        'RejectPoint_form.PictureBox1.Image.Dispose()
        'End If
        If RejectPoint_form.PictureBox1.Image Is Nothing Then
        Else
            Dim image As Image = RejectPoint_form.PictureBox1.Image
            RejectPoint_form.PictureBox1.Image = Nothing
            image.Dispose()
        End If
        If RejectPoint_form.Button_Next.Enabled = True Then 'Or RejectPoint_form.Button_Load.Enabled = False Then
            AxImage9.Save(Folder + "RM.bmp")
        Else

        End If

        If AxImageProcessing5.IsAllocated = False Then
            AxImageProcessing5.ResultSize = 2
            AxImageProcessing5.ResultType = Matrox.ActiveMIL.ImageProcessing.ImpResultTypeConstants.impExtremeList
        End If
        AxImageProcessing5.Histogram()
        AxImageProcessing5.Results.Get(Matrox.ActiveMIL.ImageProcessing.ImpResultFieldConstants.impValues, HistogramRejectPoint)
        TextBox1.Text = HistogramRejectPoint(0) & " " & HistogramRejectPoint(255) & " " & HistogramRejectPoint(1)
        Dim Black = HistogramRejectPoint(0)
        Dim White = HistogramRejectPoint(255)
        RejectPoint_form.BlacknWhite(Black, White)
        SetRMFilename(Folder)
    End Function
    Function SetRMFilename(ByVal pathname As String)
        RejectPoint_form.SetRMFilename(pathname)
    End Function
    Function RunReject(ByVal MROIx, ByVal MROIy)
        Dim HistogramRejectPoint(255) As Integer
        Dim min, max As Integer
        If AxImage8.IsAllocated = True Then
            AxImage8.Free()
        End If
        With AxImage8.ChildRegion
            .OffsetX = DisplayCenterXPosition - MROIx / 2
            .OffsetY = DisplayCenterYPosition - MROIy / 2
            .SizeX = MROIx
            .SizeY = MROIy
        End With
        AxImage8.Allocate()
        'RejectPoint_form.GetVariables(DisplayCenterXPosition, DisplayCenterYPosition, MROIx, MROIy)
        AxImageProcessing4.Binarize(Matrox.ActiveMIL.ImageProcessing.ImpConditionConstants.impGreaterOrEqualTo, RejectPoint_form.ValueBinarized.Value)
        'If RejectPoint_form.PictureBox1.Image Is Nothing Then
        'Else
        'RejectPoint_form.PictureBox1.Image.Dispose()
        'End If
        If RejectPoint_form.PictureBox2.Image Is Nothing Then
            ' no image inside (kangren)
        Else
            'Dim image As Image = RejectPoint_form.PictureBox2.Image
            'RejectPoint_form.PictureBox2.Image = Nothing
            'image.Dispose()
        End If
        'If RejectPoint_form.Button5.Enabled = False Or RejectPoint_form.Button1.Enabled = False Then
        'Else
        'AxImage9.Save("C:\IDS\SRC\DLL Export Device Vision\RejectPoint\1.bmp")
        'End If

        If AxImageProcessing5.IsAllocated = False Then
            AxImageProcessing5.ResultSize = 2
            AxImageProcessing5.ResultType = Matrox.ActiveMIL.ImageProcessing.ImpResultTypeConstants.impExtremeList
        End If
        AxImageProcessing5.Histogram()
        AxImageProcessing5.Results.Get(Matrox.ActiveMIL.ImageProcessing.ImpResultFieldConstants.impValues, HistogramRejectPoint)
        TextBox1.Text = HistogramRejectPoint(0) & " " & HistogramRejectPoint(255) & " " & HistogramRejectPoint(1)
        Dim Black = HistogramRejectPoint(0)
        Dim White = HistogramRejectPoint(255)
        RejectPoint_form.BlacknWhite(Black, White)
    End Function
    Public Function FileNameWithoutExtension(ByVal FullPath As String) As String
        Return System.IO.Path.GetDirectoryName(FullPath)
    End Function

#End Region
#Region "ChipEdgePoints"
    Dim ChipEdgeDrawingFlag As Boolean = False
    Dim ClickNoPoints As Integer
    Dim PosXPoint As Double
    Dim PosYPoint As Double
    Dim PointX1, PointY1, PointX2, PointY2, PointX3, PointY3, PointX4, PointY4, PointX5, PointY5 As Double
    Dim SizeXPoint As Double
    Dim SizeYPoint As Double
    Dim Edgelimit As Integer = 5
    Dim MainEdge As Integer = 1
    Dim MarkersPointNo1 As Integer = 0
    Dim MarkersPointNo2 As Integer = 0
    Dim MarkersPointNo3 As Integer = 0
    Dim MarkersPointNo4 As Integer = 0
    Dim MarkersPointNo5 As Integer = 0
    Dim slope1, slope2, slope3, slope4 As Double
    Dim Px1, Py1, Px2, Py2, Px3, Py3, Px4, Py4 As Double
    Dim IntC1, IntC2, IntC3, IntC4 As Double
    Dim P1x, P1y, P2x, P2y, P3x, P3y, P4x, P4y, P5x, P5y As Double
    Dim ChipEdgePoint_Score As Double
    Dim WidthPointX, WidthPointY, ChipPointPosX, ChipPointPosY, ChipPointRot As Double
    Dim Output(7) As Double
    Dim ChipFirst As Integer = 0

    Function ChipEdgeDrawingEnabled()
        Return ChipEdgeDrawingFlag
    End Function
    Function ChipEdgeDrawingDisabled()
        Return Not ChipEdgeDrawingFlag
    End Function
    Sub EnableChipEdgeDrawing()
        ChipEdgeDrawingFlag = True
    End Sub
    Sub DisableChipEdgeDrawing()
        ChipEdgeDrawingFlag = False
    End Sub
    Function ResetChipEdgePoint()
        ClearDisplay()
        PointX1 = PointY1 = 0
        PointX2 = PointY2 = 0
        PointX3 = PointY3 = 0
        PointX4 = PointY4 = 0
        PointX5 = PointY5 = 0
        ClickNoPoints = 0
        Vertical_Horizontal = True
    End Function
    Function PointNo() As Integer
        Return ClickNoPoints
    End Function
    Function ResetChipPoint()
        ClearDisplay()
        DisplayIndicator()
        PointX1 = PointY1 = 0
        PointX2 = PointY2 = 0
        PointX3 = PointY3 = 0
        PointX4 = PointY4 = 0
        PointX5 = PointY5 = 0
        ClickNoPoints = 0
    End Function
    Dim MoveOffsetXmm As Double
    Dim MoveOffsetYmm As Double
    Sub ChipEdgePoints(ByVal xPoint As Double, ByVal yPoint As Double)
        Try
            Dim vertical As Boolean
            MoveOffsetXmm = 0
            MoveOffsetYmm = 0
            If ChipEdgeDrawingEnabled() Then
                If ChipEdgeDrawingEnabled() Then
                    vertical = ChipEdgePoints_form.Vertical_Horizontal
                End If
                If ClickNoPoints = 0 Then
                    PointX1 = xPoint
                    PointY1 = yPoint
                    ClickNoPoints = ClickNoPoints + 1
                    With AxGraphicContext1.DrawingRegion
                        .SizeX = 10
                        .SizeY = 10
                        .CenterX = PointX1
                        .CenterY = PointY1
                    End With
                    AxGraphicContext1.Cross()
                    ChipEdgePoints_form.GroupBox_Vertical_Horizontal.Enabled = True
                    ChipEdgePoints_form.Button_Ok.Enabled = False
                ElseIf ClickNoPoints = 1 Then
                    PointX2 = xPoint
                    PointY2 = yPoint

                    If vertical = True Then
                        If Abs(PointX1 - PointX2) > Edgelimit Then
                            MsgBox("Wrong direction! Please re-program second point")
                        Else
                            ClickNoPoints = ClickNoPoints + 1
                            With AxGraphicContext1.DrawingRegion
                                .SizeX = 10
                                .SizeY = 10
                                .CenterX = PointX2
                                .CenterY = PointY2
                            End With
                            AxGraphicContext1.Cross()
                            ChipEdgePoints_form.GroupBox_Vertical_Horizontal.Enabled = False
                        End If
                    Else
                        If Abs(PointY1 - PointY2) > Edgelimit Then
                            MsgBox("Wrong direction! Please re-program second point")
                        Else
                            ClickNoPoints = ClickNoPoints + 1
                            With AxGraphicContext1.DrawingRegion
                                .SizeX = 10
                                .SizeY = 10
                                .CenterX = PointX2
                                .CenterY = PointY2
                            End With
                            AxGraphicContext1.Cross()
                            ChipEdgePoints_form.GroupBox_Vertical_Horizontal.Enabled = False
                        End If
                    End If
                ElseIf ClickNoPoints = 2 Then
                    PointX3 = xPoint
                    PointY3 = yPoint
                    ClickNoPoints = ClickNoPoints + 1
                    With AxGraphicContext1.DrawingRegion
                        .SizeX = 10
                        .SizeY = 10
                        .CenterX = PointX3
                        .CenterY = PointY3
                    End With
                    AxGraphicContext1.Cross()
                ElseIf ClickNoPoints = 3 Then
                    PointX4 = xPoint
                    PointY4 = yPoint
                    ClickNoPoints = ClickNoPoints + 1
                    With AxGraphicContext1.DrawingRegion
                        .SizeX = 10
                        .SizeY = 10
                        .CenterX = PointX4
                        .CenterY = PointY4
                    End With
                    AxGraphicContext1.Cross()
                ElseIf ClickNoPoints = 4 Then
                    PointX5 = xPoint
                    PointY5 = yPoint
                    ClickNoPoints = ClickNoPoints + 1
                    With AxGraphicContext1.DrawingRegion
                        .SizeX = 10
                        .SizeY = 10
                        .CenterX = PointX5
                        .CenterY = PointY5
                    End With
                    AxGraphicContext1.Cross()
                End If
                If ClickNoPoints = 5 Then

                    Dim s1 As Double = (PointY2 - PointY1) / (PointX2 - PointX1)
                    Dim s2 As Double = -1 / s1
                    Dim in1 As Double = PointY1 - PointX1 * s1
                    Dim Distance1, Distance2, Distance3 As Double

                    If s1 > 10 ^ 50 Or s1 < -10 ^ 50 Then
                        s2 = 0
                        Distance1 = Abs(PointX3 - PointX1)
                        Distance2 = Abs(PointX4 - PointX1)
                        Distance3 = Abs(PointX5 - PointX1)
                    ElseIf s2 > 10 ^ 50 Or s2 < -10 ^ 50 Then
                        s1 = 0
                        Distance1 = Abs(PointY3 - PointY1)
                        Distance2 = Abs(PointY4 - PointY1)
                        Distance3 = Abs(PointY5 - PointY1)
                    Else
                        Dim in2 As Double = PointY3 - s2 * PointX3
                        Dim in3 As Double = PointY4 - s2 * PointX4
                        Dim in4 As Double = PointY5 - s2 * PointX5
                        Dim intercept1x As Double = (in2 - in1) / (s1 - s2)
                        Dim intercept1y As Double = s2 * intercept1x + in2
                        Dim intercept2x As Double = (in3 - in1) / (s1 - s2)
                        Dim intercept2y As Double = s2 * intercept2x + in3
                        Dim intercept3x As Double = (in4 - in1) / (s1 - s2)
                        Dim intercept3y As Double = s2 * intercept3x + in4
                        Distance1 = Sqrt((PointX3 - intercept1x) ^ 2 + (PointY3 - intercept1y) ^ 2)
                        Distance2 = Sqrt((PointX4 - intercept2x) ^ 2 + (PointY4 - intercept2y) ^ 2)
                        Distance3 = Sqrt((PointX5 - intercept3x) ^ 2 + (PointY5 - intercept3y) ^ 2)
                    End If
                    Dim TempPointX, TempPointY
                    If Distance1 > Distance2 Then
                        If Distance1 > Distance3 Then
                            'distance1 furthest
                            TempPointX = PointX4
                            TempPointY = PointY4
                            PointX4 = PointX3
                            PointY4 = PointY3
                            PointX3 = TempPointX
                            PointY3 = TempPointY
                            'MsgBox("1")
                        Else
                            'distance3 furthest
                            TempPointX = PointX4
                            TempPointY = PointY4
                            PointX4 = PointX5
                            PointY4 = PointY5
                            PointX5 = TempPointX
                            PointY5 = TempPointY
                            'MsgBox("3")
                        End If
                    ElseIf Distance2 > Distance3 Then
                        'distance2 furthest
                        'MsgBox("2")
                    Else
                        'distance3 furthest
                        TempPointX = PointX4
                        TempPointY = PointY4
                        PointX4 = PointX5
                        PointY4 = PointY5
                        PointX5 = TempPointX
                        PointY5 = TempPointY
                        'MsgBox("3")
                    End If
                    If vertical = False Then ' Shift to chip center
                        SizeXPoint = Abs(PointX5 - PointX3) * PixelSizeX
                        SizeYPoint = Abs(PointY1 - PointY4) * PixelSizeY
                        PosXPoint = ((PointX5 + PointX3) / 2) * PixelSizeX
                        PosYPoint = ((PointY1 + PointY4) / 2) * PixelSizeY
                        'Ask Robot to move to the centre of the chip
                        MoveOffsetXmm = (PosXPoint / PixelSizeX - DisplayCenterXPosition) * PixelSizeX
                        MoveOffsetYmm = (PosYPoint / PixelSizeY - DisplayCenterYPosition) * PixelSizeY
                        Dim MoveOffsetXpix = (PosXPoint / PixelSizeX - DisplayCenterXPosition)
                        Dim MoveOffsetYpix = (PosYPoint / PixelSizeY - DisplayCenterYPosition)
                        PointX1 = PointX1 - MoveOffsetXpix
                        PointX2 = PointX2 - MoveOffsetXpix
                        PointX3 = PointX3 - MoveOffsetXpix
                        PointX4 = PointX4 - MoveOffsetXpix
                        PointX5 = PointX5 - MoveOffsetXpix
                        PointY1 = PointY1 - MoveOffsetYpix
                        PointY2 = PointY2 - MoveOffsetYpix
                        PointY3 = PointY3 - MoveOffsetYpix
                        PointY4 = PointY4 - MoveOffsetYpix
                        PointY5 = PointY5 - MoveOffsetYpix
                        MotionFlag = True
                        'RobotMotionOffset(MoveOffsetXmm, MoveOffsetYmm)

                        ClearDisplay()
                        SearchRegionPoints(1, ChipEdgePoints_form.ValueROI.Value)
                    Else
                        SizeXPoint = Abs(PointX1 - PointX4) * PixelSizeX
                        SizeYPoint = Abs(PointY3 - PointY5) * PixelSizeY
                        PosXPoint = ((PointX1 + PointX4) / 2) * PixelSizeX
                        PosYPoint = ((PointY3 + PointY5) / 2) * PixelSizeY
                        'Ask Robot to move to the centre of the chip
                        MoveOffsetXmm = (PosXPoint / PixelSizeX - DisplayCenterXPosition) * PixelSizeX
                        MoveOffsetYmm = (PosYPoint / PixelSizeY - DisplayCenterYPosition) * PixelSizeY
                        Dim MoveOffsetXpix = (PosXPoint / PixelSizeX - DisplayCenterXPosition)
                        Dim MoveOffsetYpix = (PosYPoint / PixelSizeY - DisplayCenterYPosition)
                        PointX1 = PointX1 - MoveOffsetXpix
                        PointX2 = PointX2 - MoveOffsetXpix
                        PointX3 = PointX3 - MoveOffsetXpix
                        PointX4 = PointX4 - MoveOffsetXpix
                        PointX5 = PointX5 - MoveOffsetXpix
                        PointY1 = PointY1 - MoveOffsetYpix
                        PointY2 = PointY2 - MoveOffsetYpix
                        PointY3 = PointY3 - MoveOffsetYpix
                        PointY4 = PointY4 - MoveOffsetYpix
                        PointY5 = PointY5 - MoveOffsetYpix
                        PosXPoint = (DisplayCenterXPosition) * PixelSizeX
                        PosYPoint = (DisplayCenterYPosition) * PixelSizeY
                        MotionFlag = True
                        'RobotMotionOffset(MoveOffsetXmm, MoveOffsetYmm)
                        ClearDisplay()
                        SearchRegionPoints(1, ChipEdgePoints_form.ValueROI.Value)
                    End If
                    If vertical = False Then
                        If PointY4 > PointY1 Then
                            '1-2
                            MainEdge = 1
                        Else
                            '3-4
                            MainEdge = 3
                        End If
                        ChipEdgePoints_form.TextBox_SizeX.Text = Format(SizeXPoint, "##0.000")
                        ChipEdgePoints_form.TextBox_SizeY.Text = Format(SizeYPoint, "##0.000")
                        ChipEdgePoints_form.TextBox_PosX.Text = Format(PosXPoint, "##0.000")
                        ChipEdgePoints_form.TextBox_PosY.Text = Format(PosYPoint, "##0.000")
                        MeasurementPoint_Settings()
                        DisableChipEdgeDrawing()
                        ChipEdgePoints_form.Button_Test.Enabled = True
                        ChipEdgePoints_form.Button_Ok.Enabled = True
                        ChipEdgePoints_form.ChipEdgePointsCoordinate(PointX1, PointY1, PointX2, PointY2, PointX3, PointY3, PointX4, PointY4, PointX5, PointY5)
                    Else
                        If PointX4 > PointX1 Then
                            '1-4
                            MainEdge = 4
                        Else
                            '2-3
                            MainEdge = 2
                        End If
                        ChipEdgePoints_form.TextBox_SizeX.Text = Format(SizeXPoint, "##0.000")
                        ChipEdgePoints_form.TextBox_SizeY.Text = Format(SizeYPoint, "##0.000")
                        ChipEdgePoints_form.TextBox_PosX.Text = Format(PosXPoint, "##0.000")
                        ChipEdgePoints_form.TextBox_PosY.Text = Format(PosYPoint, "##0.000")
                        MeasurementPoint_Settings()
                        DisableChipEdgeDrawing()
                        ChipEdgePoints_form.Button_Test.Enabled = True
                        ChipEdgePoints_form.Button_Ok.Enabled = True
                        ChipEdgePoints_form.ChipEdgePointsCoordinate(PointX1, PointY1, PointX2, PointY2, PointX3, PointY3, PointX4, PointY4, PointX5, PointY5)
                    End If
                End If
            End If
            'output offset to sj
        Catch ex As Exception
            ExceptionDisplay(ex)
        End Try
    End Sub
#Region "SS"
    Dim TextboxSizeX As New TextBox
    Dim TextboxSizeY As New TextBox
    Dim TextboxPosX As New TextBox
    Dim TextboxPosY As New TextBox
    Dim TextboxRSizeX As New TextBox
    Dim TextboxRSizeY As New TextBox
    Dim TextboxRPosX As New TextBox
    Dim TextboxRPosY As New TextBox
    Dim TextboxRRot As New TextBox
    Dim TextboxScore As New TextBox
    Dim ValueSize As New NumericUpDown
    Dim ValuePos As New NumericUpDown
    Dim ValueRot As New NumericUpDown
    Dim Vertical_Horizontal As Boolean
    Dim GroupBox_Vertical_Horizontal As New GroupBox

    Sub SS_Vertical_Horizontal(ByVal _Vertical_Horizontal As Boolean)
        Vertical_Horizontal = _Vertical_Horizontal
    End Sub
    Public Sub SetupVariable(ByRef _TextboxSizeX As TextBox, ByRef _TextboxSizeY As TextBox, ByRef _TextboxPosX As TextBox, ByRef _TextboxPosY As TextBox, ByRef _ValueSize As NumericUpDown, ByRef _ValuePos As NumericUpDown, ByRef _ValueRot As NumericUpDown)
        TextboxSizeX = _TextboxSizeX
        TextboxSizeY = _TextboxSizeY
        TextboxPosX = _TextboxPosX
        TextboxPosY = _TextboxPosY
        ValueSize = _ValueSize
        ValuePos = _ValuePos
        ValueRot = _ValueRot
    End Sub
    Public Sub SetupRVariable(ByRef _TextboxRSizeX As TextBox, ByRef _TextboxRSizeY As TextBox, ByRef _TextboxRPosX As TextBox, ByRef _TextboxRPosY As TextBox, ByRef _TextboxRRot As TextBox, ByRef _TextboxScore As TextBox)
        TextboxRSizeX = _TextboxRSizeX
        TextboxRSizeY = _TextboxRSizeY
        TextboxRPosX = _TextboxRPosX
        TextboxRPosY = _TextboxRPosY
        TextboxRRot = _TextboxRRot
        TextboxScore = _TextboxScore
    End Sub
    Function SS_ChipParrameter() As Array
        Dim SS(10) As Double
        SS(0) = SizeXPoint
        SS(1) = SizeYPoint
        SS(2) = PosXPoint
        SS(3) = PosYPoint
        Return SS
    End Function
    Function SS_Results() As Array
        Dim SS(10) As Double
        SS(0) = ChipPointPosX
        SS(1) = ChipPointPosY
        SS(2) = WidthPointX
        SS(3) = WidthPointY
        SS(4) = ChipPointRot
        Return SS
    End Function
    Function SS_Dispose()
        TextboxSizeX.Dispose()
        TextboxSizeY.Dispose()
        TextboxPosX.Dispose()
        TextboxPosY.Dispose()
        ValueSize.Dispose()
        ValuePos.Dispose()
        ValueRot.Dispose()
        TextboxRSizeX.Dispose()
        TextboxRSizeY.Dispose()
        TextboxRPosX.Dispose()
        TextboxRPosY.Dispose()
        TextboxRRot.Dispose()
        TextboxScore.Dispose()
    End Function
#End Region
    Dim MotionFlag As Boolean = False
    Function RobotMotionOffset(ByRef _MoveOffsetXmm As Double, ByRef _MoveOffsetYmm As Double) As Boolean 'to SJ
        _MoveOffsetXmm = 0
        _MoveOffsetYmm = 0
        If ClickNoPoints = 5 Then
            If MotionFlag = True Then
                _MoveOffsetXmm = MoveOffsetXmm
                _MoveOffsetYmm = MoveOffsetYmm
                MotionFlag = False
                Return True
            End If
            Return True
        Else
            _MoveOffsetXmm = 0
            _MoveOffsetYmm = 0
            Return False
        End If
        ChipEdgePoints_form.ResetChipEdgeStatus()
    End Function
    Function MeasurementPoint_Settings()
        AxMeasurement7.MeasurementList.SelectAll()
        AxMeasurement7.PixelAspectRatio.Value = 1
        AxMeasurement8.MeasurementList.SelectAll()
        AxMeasurement8.PixelAspectRatio.Value = 1
        AxMeasurement9.MeasurementList.SelectAll()
        AxMeasurement9.PixelAspectRatio.Value = 1
        AxMeasurement10.MeasurementList.SelectAll()
        AxMeasurement10.PixelAspectRatio.Value = 1
        AxMeasurement11.MeasurementList.SelectAll()
        AxMeasurement11.PixelAspectRatio.Value = 1
    End Function
    Function SearchRegionPoints(ByVal Chip_QC_SS As Integer, ByVal ROI As Double)

        Dim SizeROI As Double
        If Chip_QC_SS = 1 Then
            SizeROI = (ROI / 0.001 * PixelSizeX)
        ElseIf Chip_QC_SS = 3 Then
            SizeROI = ROI / 0.001 * PixelSizeX
        End If

        With AxGraphicContext1.DrawingRegion
            .SizeX = SizeROI
            .SizeY = SizeROI
            .CenterX = PointX1
            .CenterY = PointY1
        End With
        AxGraphicContext1.Rectangle(False, 0)
        AxGraphicContext1.CtlText("1")
        With AxGraphicContext1.DrawingRegion
            .SizeX = 10
            .SizeY = 10
            .CenterX = PointX1
            .CenterY = PointY1
        End With
        AxGraphicContext1.Cross()

        With AxGraphicContext1.DrawingRegion
            .SizeX = SizeROI
            .SizeY = SizeROI
            .CenterX = PointX2
            .CenterY = PointY2
        End With
        AxGraphicContext1.Rectangle(False, 0)
        AxGraphicContext1.CtlText("2")
        With AxGraphicContext1.DrawingRegion
            .SizeX = 10
            .SizeY = 10
            .CenterX = PointX2
            .CenterY = PointY2
        End With
        AxGraphicContext1.Cross()

        With AxGraphicContext1.DrawingRegion
            .SizeX = SizeROI
            .SizeY = SizeROI
            .CenterX = PointX3
            .CenterY = PointY3
        End With
        AxGraphicContext1.Rectangle(False, 0)
        AxGraphicContext1.CtlText("3")
        With AxGraphicContext1.DrawingRegion
            .SizeX = 10
            .SizeY = 10
            .CenterX = PointX3
            .CenterY = PointY3
        End With
        AxGraphicContext1.Cross()

        With AxGraphicContext1.DrawingRegion
            .SizeX = SizeROI
            .SizeY = SizeROI
            .CenterX = PointX4
            .CenterY = PointY4
        End With
        AxGraphicContext1.Rectangle(False, 0)
        AxGraphicContext1.CtlText("4")
        With AxGraphicContext1.DrawingRegion
            .SizeX = 10
            .SizeY = 10
            .CenterX = PointX4
            .CenterY = PointY4
        End With
        AxGraphicContext1.Cross()

        With AxGraphicContext1.DrawingRegion
            .SizeX = SizeROI
            .SizeY = SizeROI
            .CenterX = PointX5
            .CenterY = PointY5
        End With
        AxGraphicContext1.Rectangle(False, 0)
        AxGraphicContext1.CtlText("5")
        With AxGraphicContext1.DrawingRegion
            .SizeX = 10
            .SizeY = 10
            .CenterX = PointX5
            .CenterY = PointY5
        End With
        AxGraphicContext1.Cross()
    End Function

    Function ClearanceRectangle()
        ChipPointPosX = ChipPointPosX / PixelSizeX 'mm
        ChipPointPosY = ChipPointPosY / PixelSizeY

        If ChipEdgePoints_form.CheckBox_ChipRec_Enable.Checked = True Then
            Dim Output_Temp(7) As Double
            If Px1 < ChipPointPosX Then
                If Py1 < ChipPointPosY Then
                    'top left
                    Output_Temp(0) = Px1
                    Output_Temp(1) = Py1
                Else
                    'bottom left
                    Output_Temp(6) = Px1
                    Output_Temp(7) = Py1
                End If
            Else
                If Py1 < ChipPointPosY Then
                    'top right
                    Output_Temp(2) = Px1
                    Output_Temp(3) = Py1
                Else
                    'bottom right
                    Output_Temp(4) = Px1
                    Output_Temp(5) = Py1
                End If
            End If

            If Px2 < ChipPointPosX Then
                If Py2 < ChipPointPosY Then
                    'top left
                    Output_Temp(0) = Px2
                    Output_Temp(1) = Py2
                Else
                    'bottom left
                    Output_Temp(6) = Px2
                    Output_Temp(7) = Py2
                End If
            Else
                If Py2 < ChipPointPosY Then
                    'top right
                    Output_Temp(2) = Px2
                    Output_Temp(3) = Py2
                Else
                    'bottom right
                    Output_Temp(4) = Px2
                    Output_Temp(5) = Py2
                End If
            End If

            If Px3 < ChipPointPosX Then
                If Py3 < ChipPointPosY Then
                    'top left
                    Output_Temp(0) = Px3
                    Output_Temp(1) = Py3
                Else
                    'bottom left
                    Output_Temp(6) = Px3
                    Output_Temp(7) = Py3
                End If
            Else
                If Py3 < ChipPointPosY Then
                    'top right
                    Output_Temp(2) = Px3
                    Output_Temp(3) = Py3
                Else
                    'bottom right
                    Output_Temp(4) = Px3
                    Output_Temp(5) = Py3
                End If
            End If

            If Px4 < ChipPointPosX Then
                If Py4 < ChipPointPosY Then
                    'top left
                    Output_Temp(0) = Px4
                    Output_Temp(1) = Py4
                Else
                    'bottom left
                    Output_Temp(6) = Px4
                    Output_Temp(7) = Py4
                End If
            Else
                If Py4 < ChipPointPosY Then
                    'top right
                    Output_Temp(2) = Px4
                    Output_Temp(3) = Py4
                Else
                    'bottom right
                    Output_Temp(4) = Px4
                    Output_Temp(5) = Py4
                End If
            End If
            'Two vertical line
            ChipPointPosX = ChipPointPosX * PixelSizeX 'mm
            ChipPointPosY = ChipPointPosY * PixelSizeY
            If MainEdge = 1 Then
                With AxGraphicContext1.DrawingRegion
                    .StartX = Output_Temp(0) - ChipEdgePoints_form.TextBox_EdgeClearance.Text / PixelSizeX
                    .StartY = Output_Temp(1) - ChipEdgePoints_form.TextBox_EdgeClearance.Text / PixelSizeY
                    .EndX = Output_Temp(2) + ChipEdgePoints_form.TextBox_EdgeClearance.Text / PixelSizeX
                    .EndY = Output_Temp(3) - ChipEdgePoints_form.TextBox_EdgeClearance.Text / PixelSizeY
                End With
                AxGraphicContext1.LineSegment()
            Else
                With AxGraphicContext2.DrawingRegion
                    .StartX = Output_Temp(0) - ChipEdgePoints_form.TextBox_EdgeClearance.Text / PixelSizeX
                    .StartY = Output_Temp(1) - ChipEdgePoints_form.TextBox_EdgeClearance.Text / PixelSizeY
                    .EndX = Output_Temp(2) + ChipEdgePoints_form.TextBox_EdgeClearance.Text / PixelSizeX
                    .EndY = Output_Temp(3) - ChipEdgePoints_form.TextBox_EdgeClearance.Text / PixelSizeY
                End With
                AxGraphicContext2.LineSegment()
            End If

            If MainEdge = 3 Then
                With AxGraphicContext1.DrawingRegion
                    .StartX = Output_Temp(4) + ChipEdgePoints_form.TextBox_EdgeClearance.Text / PixelSizeX
                    .StartY = Output_Temp(5) + ChipEdgePoints_form.TextBox_EdgeClearance.Text / PixelSizeY
                    .EndX = Output_Temp(6) - ChipEdgePoints_form.TextBox_EdgeClearance.Text / PixelSizeX
                    .EndY = Output_Temp(7) + ChipEdgePoints_form.TextBox_EdgeClearance.Text / PixelSizeY
                End With
                AxGraphicContext1.LineSegment()
            Else
                With AxGraphicContext2.DrawingRegion
                    .StartX = Output_Temp(4) + ChipEdgePoints_form.TextBox_EdgeClearance.Text / PixelSizeX
                    .StartY = Output_Temp(5) + ChipEdgePoints_form.TextBox_EdgeClearance.Text / PixelSizeY
                    .EndX = Output_Temp(6) - ChipEdgePoints_form.TextBox_EdgeClearance.Text / PixelSizeX
                    .EndY = Output_Temp(7) + ChipEdgePoints_form.TextBox_EdgeClearance.Text / PixelSizeY
                End With
                AxGraphicContext2.LineSegment()
            End If

            'Two horizontal line
            If MainEdge = 4 Then
                With AxGraphicContext1.DrawingRegion
                    .StartX = Output_Temp(0) - ChipEdgePoints_form.TextBox_EdgeClearance.Text / PixelSizeX
                    .StartY = Output_Temp(1) - ChipEdgePoints_form.TextBox_EdgeClearance.Text / PixelSizeY
                    .EndX = Output_Temp(6) - ChipEdgePoints_form.TextBox_EdgeClearance.Text / PixelSizeX
                    .EndY = Output_Temp(7) + ChipEdgePoints_form.TextBox_EdgeClearance.Text / PixelSizeY
                End With
                AxGraphicContext1.LineSegment()
            Else
                With AxGraphicContext2.DrawingRegion
                    .StartX = Output_Temp(0) - ChipEdgePoints_form.TextBox_EdgeClearance.Text / PixelSizeX
                    .StartY = Output_Temp(1) - ChipEdgePoints_form.TextBox_EdgeClearance.Text / PixelSizeY
                    .EndX = Output_Temp(6) - ChipEdgePoints_form.TextBox_EdgeClearance.Text / PixelSizeX
                    .EndY = Output_Temp(7) + ChipEdgePoints_form.TextBox_EdgeClearance.Text / PixelSizeY
                End With
                AxGraphicContext2.LineSegment()
            End If

            If MainEdge = 2 Then
                With AxGraphicContext1.DrawingRegion
                    .StartX = Output_Temp(2) + ChipEdgePoints_form.TextBox_EdgeClearance.Text / PixelSizeX
                    .StartY = Output_Temp(3) - ChipEdgePoints_form.TextBox_EdgeClearance.Text / PixelSizeY
                    .EndX = Output_Temp(4) + ChipEdgePoints_form.TextBox_EdgeClearance.Text / PixelSizeX
                    .EndY = Output_Temp(5) + ChipEdgePoints_form.TextBox_EdgeClearance.Text / PixelSizeY
                End With
                AxGraphicContext1.LineSegment()
            Else
                With AxGraphicContext2.DrawingRegion
                    .StartX = Output_Temp(2) + ChipEdgePoints_form.TextBox_EdgeClearance.Text / PixelSizeX
                    .StartY = Output_Temp(3) - ChipEdgePoints_form.TextBox_EdgeClearance.Text / PixelSizeY
                    .EndX = Output_Temp(4) + ChipEdgePoints_form.TextBox_EdgeClearance.Text / PixelSizeX
                    .EndY = Output_Temp(5) + ChipEdgePoints_form.TextBox_EdgeClearance.Text / PixelSizeY
                End With
                AxGraphicContext2.LineSegment()
            End If
        Else
            'Two vertical line
            With AxGraphicContext2.DrawingRegion
                .StartX = (ChipEdgePoints_form.TextBox_PosX.Text - ChipEdgePoints_form.TextBox_SizeX.Text / 2 - ChipEdgePoints_form.TextBox_EdgeClearance.Text) / PixelSizeX
                .StartY = (ChipEdgePoints_form.TextBox_PosY.Text - ChipEdgePoints_form.TextBox_SizeY.Text / 2 - ChipEdgePoints_form.TextBox_EdgeClearance.Text) / PixelSizeX
                .EndX = (ChipEdgePoints_form.TextBox_PosX.Text + ChipEdgePoints_form.TextBox_SizeX.Text / 2 + ChipEdgePoints_form.TextBox_EdgeClearance.Text) / PixelSizeX
                .EndY = (ChipEdgePoints_form.TextBox_PosY.Text - ChipEdgePoints_form.TextBox_SizeY.Text / 2 - ChipEdgePoints_form.TextBox_EdgeClearance.Text) / PixelSizeX
            End With
            AxGraphicContext2.LineSegment()
            With AxGraphicContext2.DrawingRegion
                .StartX = (ChipEdgePoints_form.TextBox_PosX.Text + ChipEdgePoints_form.TextBox_SizeX.Text / 2 + ChipEdgePoints_form.TextBox_EdgeClearance.Text) / PixelSizeX
                .StartY = (ChipEdgePoints_form.TextBox_PosY.Text + ChipEdgePoints_form.TextBox_SizeY.Text / 2 + ChipEdgePoints_form.TextBox_EdgeClearance.Text) / PixelSizeX
                .EndX = (ChipEdgePoints_form.TextBox_PosX.Text - ChipEdgePoints_form.TextBox_SizeX.Text / 2 - ChipEdgePoints_form.TextBox_EdgeClearance.Text) / PixelSizeX
                .EndY = (ChipEdgePoints_form.TextBox_PosY.Text + ChipEdgePoints_form.TextBox_SizeY.Text / 2 + ChipEdgePoints_form.TextBox_EdgeClearance.Text) / PixelSizeX
            End With
            AxGraphicContext2.LineSegment()
            'Two horizontal line
            With AxGraphicContext2.DrawingRegion
                .StartX = (ChipEdgePoints_form.TextBox_PosX.Text - ChipEdgePoints_form.TextBox_SizeX.Text / 2 - ChipEdgePoints_form.TextBox_EdgeClearance.Text) / PixelSizeX
                .StartY = (ChipEdgePoints_form.TextBox_PosY.Text - ChipEdgePoints_form.TextBox_SizeY.Text / 2 - ChipEdgePoints_form.TextBox_EdgeClearance.Text) / PixelSizeX
                .EndX = (ChipEdgePoints_form.TextBox_PosX.Text - ChipEdgePoints_form.TextBox_SizeX.Text / 2 - ChipEdgePoints_form.TextBox_EdgeClearance.Text) / PixelSizeX
                .EndY = (ChipEdgePoints_form.TextBox_PosY.Text + ChipEdgePoints_form.TextBox_SizeY.Text / 2 + ChipEdgePoints_form.TextBox_EdgeClearance.Text) / PixelSizeX
            End With
            AxGraphicContext2.LineSegment()
            With AxGraphicContext2.DrawingRegion
                .StartX = (ChipEdgePoints_form.TextBox_PosX.Text + ChipEdgePoints_form.TextBox_SizeX.Text / 2 + ChipEdgePoints_form.TextBox_EdgeClearance.Text) / PixelSizeX
                .StartY = (ChipEdgePoints_form.TextBox_PosY.Text - ChipEdgePoints_form.TextBox_SizeY.Text / 2 - ChipEdgePoints_form.TextBox_EdgeClearance.Text) / PixelSizeX
                .EndX = (ChipEdgePoints_form.TextBox_PosX.Text + ChipEdgePoints_form.TextBox_SizeX.Text / 2 + ChipEdgePoints_form.TextBox_EdgeClearance.Text) / PixelSizeX
                .EndY = (ChipEdgePoints_form.TextBox_PosY.Text + ChipEdgePoints_form.TextBox_SizeY.Text / 2 + ChipEdgePoints_form.TextBox_EdgeClearance.Text) / PixelSizeX
            End With
            AxGraphicContext2.LineSegment()
        End If
    End Function
    Function MeasurementPoint(ByVal Contrast As Double, ByVal Threshold As Double, ByVal RotationAngle As Double, ByVal Inside_Out As Boolean, ByVal vertical As Boolean, ByVal ROI As Double, ByVal Chip_QC_SS As Integer) As Boolean
        Dim SizeROI As Double = (ChipEdgePoints_form.ValueROI.Value / 0.001 * PixelSizeX)
        If Chip_QC_SS = 1 Then
            SizeROI = ROI / 0.001 * PixelSizeX
        ElseIf Chip_QC_SS = 3 Then
            SizeROI = ROI / 0.001 * PixelSizeX
        End If
        Try
            If MarkersPointNo1 = 1 Or MarkersPointNo2 = 1 Or MarkersPointNo3 = 1 Or MarkersPointNo4 = 1 Or MarkersPointNo5 = 1 Then
                AxMeasurement7.Markers.Remove(1)
                AxMeasurement8.Markers.Remove(1)
                AxMeasurement9.Markers.Remove(1)
                AxMeasurement10.Markers.Remove(1)
                AxMeasurement11.Markers.Remove(1)
                MarkersPointNo1 = 0
                MarkersPointNo2 = 0
                MarkersPointNo3 = 0
                MarkersPointNo4 = 0
                MarkersPointNo5 = 0
            End If
            MarkersPointNo1 = AxMeasurement7.Markers.Add(Measurement.MeasMarkerTypeConstants.measEdge)
            With AxMeasurement7.Markers.Item(MarkersPointNo1)
                .Position.X = PointX1
                .Position.Y = PointY1
                .ReferenceX = 0
                .ReferenceY = 0

                'brightness
                .Contrast.Edge1 = Contrast
                .EdgeStrength.Edge1 = 50
                .EdgeThreshold = Threshold
                .NumberOfOccurrences = 1

                If vertical = False Then
                    .Orientation = Measurement.MeasMarkerOrientationConstants.measHorizontal
                    .Polarity.Edge1 = Measurement.MeasMarkerPolarityConstants.measAnyPolarity
                Else
                    .Orientation = Measurement.MeasMarkerOrientationConstants.measVertical
                    .Polarity.Edge1 = Measurement.MeasMarkerPolarityConstants.measAnyPolarity
                End If

                .SearchRegion.SizeX = Convert.ToInt16(SizeROI)
                .SearchRegion.SizeY = Convert.ToInt16(SizeROI)
                .SearchRegion.CenterX = PointX1
                .SearchRegion.CenterY = PointY1

                .SearchRegion.Angle.Enabled = True
                .SearchRegion.Angle.Accuracy = 0.1
                .SearchRegion.Angle.InterpolationMode = AngleInterpolationModeConstants.angleBilinear
                .SearchRegion.Angle.CenterOfRotation = AngleCenterOfRotationConstants.angleCenter

                If Inside_Out = True Then
                    .SearchRegion.Angle.Value = 180
                Else
                    .SearchRegion.Angle.Value = 0
                End If

                .SearchRegion.Angle.NegativeDelta = RotationAngle
                .SearchRegion.Angle.PositiveDelta = RotationAngle
            End With
            MarkersPointNo2 = AxMeasurement8.Markers.Add(Matrox.ActiveMIL.Measurement.MeasMarkerTypeConstants.measEdge)
            AxMeasurement8.Markers.Item(MarkersPointNo2).Position.X = PointX2
            AxMeasurement8.Markers.Item(MarkersPointNo2).Position.Y = PointY2
            AxMeasurement8.Markers.Item(MarkersPointNo2).ReferenceX = 0
            AxMeasurement8.Markers.Item(MarkersPointNo2).ReferenceY = 0

            'brightness
            AxMeasurement8.Markers.Item(MarkersPointNo2).Contrast.Edge1 = Contrast
            AxMeasurement8.Markers.Item(MarkersPointNo2).EdgeStrength.Edge1 = 50
            AxMeasurement8.Markers.Item(MarkersPointNo2).EdgeThreshold = Threshold
            AxMeasurement8.Markers.Item(MarkersPointNo2).NumberOfOccurrences = 1

            If vertical = False Then
                AxMeasurement8.Markers.Item(MarkersPointNo2).Orientation = Matrox.ActiveMIL.Measurement.MeasMarkerOrientationConstants.measHorizontal
                AxMeasurement8.Markers.Item(MarkersPointNo2).Polarity.Edge1 = Matrox.ActiveMIL.Measurement.MeasMarkerPolarityConstants.measAnyPolarity
            Else
                AxMeasurement8.Markers.Item(MarkersPointNo2).Orientation = Matrox.ActiveMIL.Measurement.MeasMarkerOrientationConstants.measVertical
                AxMeasurement8.Markers.Item(MarkersPointNo2).Polarity.Edge1 = Matrox.ActiveMIL.Measurement.MeasMarkerPolarityConstants.measAnyPolarity
            End If

            AxMeasurement8.Markers.Item(MarkersPointNo2).SearchRegion.SizeX = Convert.ToInt16(SizeROI)
            AxMeasurement8.Markers.Item(MarkersPointNo2).SearchRegion.SizeY = Convert.ToInt16(SizeROI)
            AxMeasurement8.Markers.Item(MarkersPointNo2).SearchRegion.CenterX = PointX2
            AxMeasurement8.Markers.Item(MarkersPointNo2).SearchRegion.CenterY = PointY2

            AxMeasurement8.Markers.Item(MarkersPointNo2).SearchRegion.Angle.Enabled = True
            AxMeasurement8.Markers.Item(MarkersPointNo2).SearchRegion.Angle.Accuracy = 0.1
            AxMeasurement8.Markers.Item(MarkersPointNo2).SearchRegion.Angle.InterpolationMode = Matrox.ActiveMIL.AngleInterpolationModeConstants.angleBilinear
            AxMeasurement8.Markers.Item(MarkersPointNo2).SearchRegion.Angle.CenterOfRotation = Matrox.ActiveMIL.AngleCenterOfRotationConstants.angleCenter

            If Inside_Out = True Then
                AxMeasurement8.Markers.Item(MarkersPointNo2).SearchRegion.Angle.Value = 0
            Else
                AxMeasurement8.Markers.Item(MarkersPointNo2).SearchRegion.Angle.Value = 180
            End If

            AxMeasurement8.Markers.Item(MarkersPointNo2).SearchRegion.Angle.NegativeDelta = RotationAngle
            AxMeasurement8.Markers.Item(MarkersPointNo2).SearchRegion.Angle.PositiveDelta = RotationAngle
            MarkersPointNo3 = AxMeasurement9.Markers.Add(Matrox.ActiveMIL.Measurement.MeasMarkerTypeConstants.measEdge)
            AxMeasurement9.Markers.Item(MarkersPointNo3).Position.X = PointX3
            AxMeasurement9.Markers.Item(MarkersPointNo3).Position.Y = PointY3
            AxMeasurement9.Markers.Item(MarkersPointNo3).ReferenceX = 0
            AxMeasurement9.Markers.Item(MarkersPointNo3).ReferenceY = 0

            'brightness
            AxMeasurement9.Markers.Item(MarkersPointNo3).Contrast.Edge1 = Contrast
            AxMeasurement9.Markers.Item(MarkersPointNo3).EdgeStrength.Edge1 = 50
            AxMeasurement9.Markers.Item(MarkersPointNo3).EdgeThreshold = Threshold
            AxMeasurement9.Markers.Item(MarkersPointNo3).NumberOfOccurrences = 1

            If vertical = False Then
                AxMeasurement9.Markers.Item(MarkersPointNo3).Orientation = Matrox.ActiveMIL.Measurement.MeasMarkerOrientationConstants.measVertical
                AxMeasurement9.Markers.Item(MarkersPointNo3).Polarity.Edge1 = Matrox.ActiveMIL.Measurement.MeasMarkerPolarityConstants.measAnyPolarity
            Else
                AxMeasurement9.Markers.Item(MarkersPointNo3).Orientation = Matrox.ActiveMIL.Measurement.MeasMarkerOrientationConstants.measHorizontal
                AxMeasurement9.Markers.Item(MarkersPointNo3).Polarity.Edge1 = Matrox.ActiveMIL.Measurement.MeasMarkerPolarityConstants.measAnyPolarity
            End If

            AxMeasurement9.Markers.Item(MarkersPointNo3).SearchRegion.SizeX = Convert.ToInt16(SizeROI)
            AxMeasurement9.Markers.Item(MarkersPointNo3).SearchRegion.SizeY = Convert.ToInt16(SizeROI)
            AxMeasurement9.Markers.Item(MarkersPointNo3).SearchRegion.CenterX = PointX3
            AxMeasurement9.Markers.Item(MarkersPointNo3).SearchRegion.CenterY = PointY3

            AxMeasurement9.Markers.Item(MarkersPointNo3).SearchRegion.Angle.Enabled = True
            AxMeasurement9.Markers.Item(MarkersPointNo3).SearchRegion.Angle.Accuracy = 0.1
            AxMeasurement9.Markers.Item(MarkersPointNo3).SearchRegion.Angle.InterpolationMode = Matrox.ActiveMIL.AngleInterpolationModeConstants.angleBilinear
            AxMeasurement9.Markers.Item(MarkersPointNo3).SearchRegion.Angle.CenterOfRotation = Matrox.ActiveMIL.AngleCenterOfRotationConstants.angleCenter

            If Inside_Out = True Then
                AxMeasurement9.Markers.Item(MarkersPointNo3).SearchRegion.Angle.Value = 0
            Else
                AxMeasurement9.Markers.Item(MarkersPointNo3).SearchRegion.Angle.Value = 180
            End If

            AxMeasurement9.Markers.Item(MarkersPointNo3).SearchRegion.Angle.NegativeDelta = RotationAngle
            AxMeasurement9.Markers.Item(MarkersPointNo3).SearchRegion.Angle.PositiveDelta = RotationAngle
            MarkersPointNo4 = AxMeasurement10.Markers.Add(Matrox.ActiveMIL.Measurement.MeasMarkerTypeConstants.measEdge)
            AxMeasurement10.Markers.Item(MarkersPointNo4).Position.X = PointX4
            AxMeasurement10.Markers.Item(MarkersPointNo4).Position.Y = PointY4
            AxMeasurement10.Markers.Item(MarkersPointNo4).ReferenceX = 0
            AxMeasurement10.Markers.Item(MarkersPointNo4).ReferenceY = 0

            'brightness
            AxMeasurement10.Markers.Item(MarkersPointNo4).Contrast.Edge1 = Contrast
            AxMeasurement10.Markers.Item(MarkersPointNo4).EdgeStrength.Edge1 = 50
            AxMeasurement10.Markers.Item(MarkersPointNo4).EdgeThreshold = Threshold
            AxMeasurement10.Markers.Item(MarkersPointNo4).NumberOfOccurrences = 1

            If vertical = False Then
                AxMeasurement10.Markers.Item(MarkersPointNo4).Orientation = Matrox.ActiveMIL.Measurement.MeasMarkerOrientationConstants.measHorizontal
                AxMeasurement10.Markers.Item(MarkersPointNo4).Polarity.Edge1 = Matrox.ActiveMIL.Measurement.MeasMarkerPolarityConstants.measAnyPolarity
            Else
                AxMeasurement10.Markers.Item(MarkersPointNo4).Orientation = Matrox.ActiveMIL.Measurement.MeasMarkerOrientationConstants.measVertical
                AxMeasurement10.Markers.Item(MarkersPointNo4).Polarity.Edge1 = Matrox.ActiveMIL.Measurement.MeasMarkerPolarityConstants.measAnyPolarity
            End If

            AxMeasurement10.Markers.Item(MarkersPointNo4).SearchRegion.SizeX = Convert.ToInt16(SizeROI)
            AxMeasurement10.Markers.Item(MarkersPointNo4).SearchRegion.SizeY = Convert.ToInt16(SizeROI)
            AxMeasurement10.Markers.Item(MarkersPointNo4).SearchRegion.CenterX = PointX4
            AxMeasurement10.Markers.Item(MarkersPointNo4).SearchRegion.CenterY = PointY4

            AxMeasurement10.Markers.Item(MarkersPointNo4).SearchRegion.Angle.Enabled = True
            AxMeasurement10.Markers.Item(MarkersPointNo4).SearchRegion.Angle.Accuracy = 0.1
            AxMeasurement10.Markers.Item(MarkersPointNo4).SearchRegion.Angle.InterpolationMode = Matrox.ActiveMIL.AngleInterpolationModeConstants.angleBilinear
            AxMeasurement10.Markers.Item(MarkersPointNo4).SearchRegion.Angle.CenterOfRotation = Matrox.ActiveMIL.AngleCenterOfRotationConstants.angleCenter

            If Inside_Out = True Then
                AxMeasurement10.Markers.Item(MarkersPointNo4).SearchRegion.Angle.Value = 180
            Else
                AxMeasurement10.Markers.Item(MarkersPointNo4).SearchRegion.Angle.Value = 0
            End If

            AxMeasurement10.Markers.Item(MarkersPointNo4).SearchRegion.Angle.NegativeDelta = RotationAngle
            AxMeasurement10.Markers.Item(MarkersPointNo4).SearchRegion.Angle.PositiveDelta = RotationAngle
            MarkersPointNo5 = AxMeasurement11.Markers.Add(Matrox.ActiveMIL.Measurement.MeasMarkerTypeConstants.measEdge)
            AxMeasurement11.Markers.Item(MarkersPointNo5).Position.X = PointX5
            AxMeasurement11.Markers.Item(MarkersPointNo5).Position.Y = PointY5
            AxMeasurement11.Markers.Item(MarkersPointNo5).ReferenceX = 0
            AxMeasurement11.Markers.Item(MarkersPointNo5).ReferenceY = 0

            'brightness
            AxMeasurement11.Markers.Item(MarkersPointNo5).Contrast.Edge1 = Contrast
            AxMeasurement11.Markers.Item(MarkersPointNo5).EdgeStrength.Edge1 = 50
            AxMeasurement11.Markers.Item(MarkersPointNo5).EdgeThreshold = Threshold
            AxMeasurement11.Markers.Item(MarkersPointNo5).NumberOfOccurrences = 1

            If vertical = False Then
                AxMeasurement11.Markers.Item(MarkersPointNo5).Orientation = Matrox.ActiveMIL.Measurement.MeasMarkerOrientationConstants.measVertical
                AxMeasurement11.Markers.Item(MarkersPointNo5).Polarity.Edge1 = Matrox.ActiveMIL.Measurement.MeasMarkerPolarityConstants.measAnyPolarity
            Else
                AxMeasurement11.Markers.Item(MarkersPointNo5).Orientation = Matrox.ActiveMIL.Measurement.MeasMarkerOrientationConstants.measHorizontal
                AxMeasurement11.Markers.Item(MarkersPointNo5).Polarity.Edge1 = Matrox.ActiveMIL.Measurement.MeasMarkerPolarityConstants.measAnyPolarity
            End If

            AxMeasurement11.Markers.Item(MarkersPointNo5).SearchRegion.SizeX = Convert.ToInt16(SizeROI)
            AxMeasurement11.Markers.Item(MarkersPointNo5).SearchRegion.SizeY = Convert.ToInt16(SizeROI)
            AxMeasurement11.Markers.Item(MarkersPointNo5).SearchRegion.CenterX = PointX5
            AxMeasurement11.Markers.Item(MarkersPointNo5).SearchRegion.CenterY = PointY5
            AxMeasurement11.Markers.Item(MarkersPointNo5).SearchRegion.Angle.Enabled = True
            AxMeasurement11.Markers.Item(MarkersPointNo5).SearchRegion.Angle.Accuracy = 0.1
            AxMeasurement11.Markers.Item(MarkersPointNo5).SearchRegion.Angle.InterpolationMode = Matrox.ActiveMIL.AngleInterpolationModeConstants.angleBilinear
            AxMeasurement11.Markers.Item(MarkersPointNo5).SearchRegion.Angle.CenterOfRotation = Matrox.ActiveMIL.AngleCenterOfRotationConstants.angleCenter

            If Inside_Out = True Then
                AxMeasurement11.Markers.Item(MarkersPointNo5).SearchRegion.Angle.Value = 180
            Else
                AxMeasurement11.Markers.Item(MarkersPointNo5).SearchRegion.Angle.Value = 0
            End If

            AxMeasurement11.Markers.Item(MarkersPointNo5).SearchRegion.Angle.NegativeDelta = RotationAngle
            AxMeasurement11.Markers.Item(MarkersPointNo5).SearchRegion.Angle.PositiveDelta = RotationAngle

            Try
                Try
                    AxMeasurement7.FindMarker(MarkersPointNo1)
                    AxMeasurement8.FindMarker(MarkersPointNo2)
                    AxMeasurement9.FindMarker(MarkersPointNo3)
                    AxMeasurement10.FindMarker(MarkersPointNo4)
                    AxMeasurement11.FindMarker(MarkersPointNo5)

                Catch ex As Exception
                    ExceptionDisplay(ex)
                End Try

                If AxMeasurement7.Results.Count > 0 Then
                    If AxMeasurement7.Markers.Item(MarkersPointNo1).IsFound = True Then
                        P1x = AxMeasurement7.Results.Item(MarkersPointNo1).PositionX
                        P1y = AxMeasurement7.Results.Item(MarkersPointNo1).PositionY
                        With AxGraphicContext2.DrawingRegion
                            .SizeX = SizeROI
                            .SizeY = SizeROI
                            .CenterX = P1x
                            .CenterY = P1y
                        End With
                        AxGraphicContext2.Rectangle(False, 0)
                        AxGraphicContext2.CtlText("1")
                        With AxGraphicContext2.DrawingRegion
                            .SizeX = 10
                            .SizeY = 10
                            .CenterX = P1x
                            .CenterY = P1y
                        End With
                        AxGraphicContext2.Cross()

                    End If
                End If
                If AxMeasurement8.Results.Count > 0 Then
                    If AxMeasurement8.Markers.Item(MarkersPointNo2).IsFound = True Then
                        P2x = AxMeasurement8.Results.Item(MarkersPointNo2).PositionX
                        P2y = AxMeasurement8.Results.Item(MarkersPointNo2).PositionY
                        With AxGraphicContext2.DrawingRegion
                            .SizeX = SizeROI
                            .SizeY = SizeROI
                            .CenterX = P2x
                            .CenterY = P2y
                        End With
                        AxGraphicContext2.Rectangle(False, 0)
                        AxGraphicContext2.CtlText("2")
                        With AxGraphicContext2.DrawingRegion
                            .SizeX = 10
                            .SizeY = 10
                            .CenterX = P2x
                            .CenterY = P2y
                        End With
                        AxGraphicContext2.Cross()
                    End If
                End If
                If AxMeasurement9.Results.Count > 0 Then
                    If AxMeasurement9.Markers.Item(MarkersPointNo3).IsFound = True Then
                        P3x = AxMeasurement9.Results.Item(MarkersPointNo3).PositionX
                        P3y = AxMeasurement9.Results.Item(MarkersPointNo3).PositionY
                        With AxGraphicContext2.DrawingRegion
                            .SizeX = SizeROI
                            .SizeY = SizeROI
                            .CenterX = P3x
                            .CenterY = P3y
                        End With
                        AxGraphicContext2.Rectangle(False, 0)
                        AxGraphicContext2.CtlText("3")
                        With AxGraphicContext2.DrawingRegion
                            .SizeX = 10
                            .SizeY = 10
                            .CenterX = P3x
                            .CenterY = P3y
                        End With
                        AxGraphicContext2.Cross()
                    End If

                End If
                If AxMeasurement10.Results.Count > 0 Then
                    If AxMeasurement10.Markers.Item(MarkersPointNo4).IsFound = True Then
                        P4x = AxMeasurement10.Results.Item(MarkersPointNo4).PositionX
                        P4y = AxMeasurement10.Results.Item(MarkersPointNo4).PositionY
                        With AxGraphicContext2.DrawingRegion
                            .SizeX = SizeROI
                            .SizeY = SizeROI
                            .CenterX = P4x
                            .CenterY = P4y
                        End With
                        AxGraphicContext2.Rectangle(False, 0)
                        AxGraphicContext2.CtlText("4")
                        With AxGraphicContext2.DrawingRegion
                            .SizeX = 10
                            .SizeY = 10
                            .CenterX = P4x
                            .CenterY = P4y
                        End With
                        AxGraphicContext2.Cross()
                    End If
                End If
                If AxMeasurement11.Results.Count > 0 Then
                    If AxMeasurement11.Markers.Item(MarkersPointNo5).IsFound = True Then
                        P5x = AxMeasurement11.Results.Item(MarkersPointNo5).PositionX
                        P5y = AxMeasurement11.Results.Item(MarkersPointNo5).PositionY

                        With AxGraphicContext2.DrawingRegion
                            .SizeX = SizeROI
                            .SizeY = SizeROI
                            .CenterX = P5x
                            .CenterY = P5y
                        End With
                        AxGraphicContext2.Rectangle(False, 0)
                        AxGraphicContext2.CtlText("5")
                        With AxGraphicContext2.DrawingRegion
                            .SizeX = 10
                            .SizeY = 10
                            .CenterX = P5x
                            .CenterY = P5y
                        End With
                        AxGraphicContext2.Cross()
                    End If
                End If

                If AxMeasurement7.Results.Count > 0 And AxMeasurement8.Results.Count > 0 And AxMeasurement9.Results.Count > 0 And AxMeasurement10.Results.Count > 0 And AxMeasurement11.Results.Count > 0 Then

                    slope1 = (P1y - P2y) / (P1x - P2x)
                    IntC1 = P1y - slope1 * P1x
                    slope2 = (-1) / slope1
                    IntC2 = P3y - slope2 * P3x
                    slope3 = (-1) / slope2
                    IntC3 = P4y - slope3 * P4x
                    slope4 = (-1) / slope1
                    IntC4 = P5y - slope4 * P5x
                    If slope1 > 10 ^ 50 Or slope1 < (-10 ^ 50) Then
                        Px1 = P1x
                        Py1 = P2y
                        Px2 = P4x
                        Py2 = P2y
                        Px3 = P4x
                        Py3 = P5y
                        Px4 = P1x
                        Py4 = P5y
                    Else
                        'line Vertical_1 intesect with Horizontal_1
                        Px1 = (IntC1 - IntC2) / (slope2 - slope1)
                        Py1 = slope1 * Px1 + IntC1
                        'line Vertical_1 intesect with Horizontal_2
                        Px2 = (IntC2 - IntC3) / (slope3 - slope2)
                        Py2 = slope2 * Px2 + IntC2

                        'line Vertical_2 intesect with Horizontal_1
                        Px3 = (IntC3 - IntC4) / (slope4 - slope3)
                        Py3 = slope3 * Px3 + IntC3

                        'line Vertical_2 intesect with Horizontal_2
                        Px4 = (IntC1 - IntC4) / (slope4 - slope1)
                        Py4 = slope1 * Px4 + IntC1
                    End If

                    If vertical = False Then
                        If Chip_QC_SS = 1 Then
                            WidthPointY = Sqrt((Py1 - Py2) ^ 2 + (Px1 - Px2) ^ 2) * PixelSizeY
                            WidthPointX = Sqrt((Py1 - Py4) ^ 2 + (Px1 - Px4) ^ 2) * PixelSizeX
                            ChipPointPosX = (Px1 + Px2 + Px3 + Px4) / 4 * PixelSizeX
                            ChipPointPosY = (Py1 + Py2 + Py3 + Py4) / 4 * PixelSizeY
                            ChipPointRot = Atan(-(Py4 - Py1) / (Px4 - Px1)) / PI * 180
                            ChipEdgePoints_form.TextBox_RPosX.Text = ChipPointPosX
                            ChipEdgePoints_form.TextBox_RPosY.Text = ChipPointPosY
                            ChipEdgePoints_form.TextBox_RSizeX.Text = WidthPointX
                            ChipEdgePoints_form.TextBox_RSizeY.Text = WidthPointY
                            ChipEdgePoints_form.TextBox_RRot.Text = ChipPointRot
                            ChipPointDrawing(Chip_QC_SS)
                        ElseIf Chip_QC_SS = 3 Then
                            WidthPointY = Sqrt((Py1 - Py2) ^ 2 + (Px1 - Px2) ^ 2) * PixelSizeY
                            WidthPointX = Sqrt((Py1 - Py4) ^ 2 + (Px1 - Px4) ^ 2) * PixelSizeX
                            ChipPointPosX = (Px1 + Px2 + Px3 + Px4) / 4 * PixelSizeX
                            ChipPointPosY = (Py1 + Py2 + Py3 + Py4) / 4 * PixelSizeY
                            ChipPointRot = Atan(-(Py4 - Py1) / (Px4 - Px1)) / PI * 180
                            TextboxRPosX.Text = ChipPointPosX
                            TextboxRPosY.Text = ChipPointPosY
                            TextboxRSizeX.Text = WidthPointX
                            TextboxRSizeY.Text = WidthPointY
                            TextboxRRot.Text = ChipPointRot
                            ChipPointDrawing(Chip_QC_SS)
                        End If
                    Else
                        If Chip_QC_SS = 1 Then
                            WidthPointX = Sqrt((Py1 - Py2) ^ 2 + (Px1 - Px2) ^ 2) * PixelSizeX
                            WidthPointY = Sqrt((Py1 - Py4) ^ 2 + (Px1 - Px4) ^ 2) * PixelSizeY
                            ChipPointPosX = (Px1 + Px2 + Px3 + Px4) / 4 * PixelSizeX
                            ChipPointPosY = (Py1 + Py2 + Py3 + Py4) / 4 * PixelSizeY
                            If Atan((Py4 - Py1) / (Px4 - Px1)) / PI * 180 > 0 Then
                                ChipPointRot = 90 - Atan((Py4 - Py1) / (Px4 - Px1)) / PI * 180
                            Else
                                ChipPointRot = -90 - Atan((Py4 - Py1) / (Px4 - Px1)) / PI * 180
                            End If
                            ChipEdgePoints_form.TextBox_RPosX.Text = ChipPointPosX
                            ChipEdgePoints_form.TextBox_RPosY.Text = ChipPointPosY
                            ChipEdgePoints_form.TextBox_RSizeX.Text = WidthPointX
                            ChipEdgePoints_form.TextBox_RSizeY.Text = WidthPointY
                            ChipEdgePoints_form.TextBox_RRot.Text = ChipPointRot
                            ChipPointDrawing(Chip_QC_SS)
                        ElseIf Chip_QC_SS = 3 Then
                            WidthPointX = Sqrt((Py1 - Py2) ^ 2 + (Px1 - Px2) ^ 2) * PixelSizeX
                            WidthPointY = Sqrt((Py1 - Py4) ^ 2 + (Px1 - Px4) ^ 2) * PixelSizeY
                            ChipPointPosX = (Px1 + Px2 + Px3 + Px4) / 4 * PixelSizeX
                            ChipPointPosY = (Py1 + Py2 + Py3 + Py4) / 4 * PixelSizeY
                            If Atan((Py4 - Py1) / (Px4 - Px1)) / PI * 180 > 0 Then
                                ChipPointRot = 90 - Atan((Py4 - Py1) / (Px4 - Px1)) / PI * 180
                            Else
                                ChipPointRot = -90 - Atan((Py4 - Py1) / (Px4 - Px1)) / PI * 180
                            End If
                            'SS_form.TextBox_RPosX.Text = ChipPointPosX
                            'SS_form.TextBox_RPosY.Text = ChipPointPosY
                            'SS_form.TextBox_RSizeX.Text = WidthPointX
                            'SS_form.TextBox_RSizeY.Text = WidthPointY
                            'SS_form.TextBox_RRot.Text = ChipPointRot
                            TextboxRPosX.Text = ChipPointPosX
                            TextboxRPosY.Text = ChipPointPosY
                            TextboxRSizeX.Text = WidthPointX
                            TextboxRSizeY.Text = WidthPointY
                            TextboxRRot.Text = ChipPointRot
                            ChipPointDrawing(Chip_QC_SS)
                        End If
                    End If
                    Return True
                Else
                    Dim Point1, Point2, Point3, Point4, Point5 As String
                    If AxMeasurement7.Results.Count = 0 Then
                        Point1 = "1,"
                    Else
                        Point1 = ""
                    End If
                    If AxMeasurement8.Results.Count = 0 Then
                        Point2 = "2,"
                    Else
                        Point2 = ""
                    End If
                    If AxMeasurement9.Results.Count = 0 Then
                        Point3 = "3,"
                    Else
                        Point3 = ""
                    End If
                    If AxMeasurement10.Results.Count = 0 Then
                        Point4 = "4,"
                    Else
                        Point4 = ""
                    End If
                    If AxMeasurement11.Results.Count = 0 Then
                        Point5 = 5
                    Else
                        Point5 = ""
                    End If

                    With AxGraphicContext1
                        .FontScaleX = 1
                        .FontScaleY = 1
                        .DrawingRegion.StartX = DisplayCenterXPosition - 50
                        .DrawingRegion.StartY = DisplayCenterYPosition + 25
                        .DrawingRegion.EndX = DisplayWidth
                        .DrawingRegion.EndY = DisplayHeight
                    End With
                    Dim text1 As String = ("")
                    AxGraphicContext1.CtlText(text1)
                    Dim text As String = ("Can't find Edge:" + Point1 + Point2 + Point3 + Point4 + Point5)
                    AxGraphicContext1.CtlText(text)

                    If Chip_QC_SS = 1 Then
                        ChipEdgePoints_form.TextBox_RPosX.Text = 0
                        ChipEdgePoints_form.TextBox_RPosY.Text = 0
                        ChipEdgePoints_form.TextBox_RSizeX.Text = 0
                        ChipEdgePoints_form.TextBox_RSizeY.Text = 0
                        ChipEdgePoints_form.TextBox_RRot.Text = 0
                        Return False
                    ElseIf Chip_QC_SS = 3 Then
                        TextboxRPosX.Text = 0
                        TextboxRPosY.Text = 0
                        TextboxRSizeX.Text = 0
                        TextboxRSizeY.Text = 0
                        TextboxRRot.Text = 0
                        ChipPointPosX = 0
                        ChipPointPosY = 0
                        WidthPointX = 0
                        WidthPointY = 0
                        ChipPointRot = 0
                    End If

                End If

            Catch ex As Exception
                ExceptionDisplay(ex)
            End Try
        Catch ex As Exception
            ExceptionDisplay(ex)
        End Try

    End Function
    Function ChipPointDrawing(ByVal Chip_QC_SS As Integer)
        If AxMeasurement7.Results.Count > 0 And AxMeasurement8.Results.Count > 0 And AxMeasurement9.Results.Count > 0 And AxMeasurement10.Results.Count > 0 And AxMeasurement11.Results.Count > 0 Then
            If AxMeasurement7.Markers.Item(MarkersPointNo1).IsFound And AxMeasurement8.Markers.Item(MarkersPointNo2).IsFound And AxMeasurement9.Markers.Item(MarkersPointNo3).IsFound And AxMeasurement10.Markers.Item(MarkersPointNo4).IsFound And AxMeasurement11.Markers.Item(MarkersPointNo5).IsFound Then
                If Chip_QC_SS = 1 Then
                    If _
                    ((ChipEdgePoints_form.TextBox_SizeX.Text) * (1 - ChipEdgePoints_form.ValueSize.Value / 100)) < WidthPointX And _
                    WidthPointX < ((ChipEdgePoints_form.TextBox_SizeX.Text) * (1 + ChipEdgePoints_form.ValueSize.Value / 100)) And _
                    ((ChipEdgePoints_form.TextBox_SizeY.Text) * (1 - ChipEdgePoints_form.ValueSize.Value / 100)) < WidthPointY And _
                    WidthPointY < ((ChipEdgePoints_form.TextBox_SizeY.Text) * (1 + ChipEdgePoints_form.ValueSize.Value / 100)) And _
                    ((ChipEdgePoints_form.TextBox_PosX.Text) * (1 - ChipEdgePoints_form.ValuePos.Value / 100)) < ChipPointPosX And _
                    ChipPointPosX < ((ChipEdgePoints_form.TextBox_PosX.Text) * (1 + ChipEdgePoints_form.ValuePos.Value / 100)) And _
                    ((ChipEdgePoints_form.TextBox_PosY.Text) * (1 - ChipEdgePoints_form.ValuePos.Value / 100)) < ChipPointPosY And _
                    ChipPointPosY < ((ChipEdgePoints_form.TextBox_PosY.Text) * (1 + ChipEdgePoints_form.ValuePos.Value / 100)) And _
                    ChipPointRot < ((ChipEdgePoints_form.ValueRot.Value)) And _
                    (-(ChipEdgePoints_form.ValueRot.Value) < ChipPointRot) Then
                        Try
                            'Two vertical line
                            With AxGraphicContext1.DrawingRegion
                                .StartX = Px1
                                .StartY = Py1
                                .EndX = Px2
                                .EndY = Py2
                            End With
                            AxGraphicContext1.LineSegment()
                            With AxGraphicContext1.DrawingRegion
                                .StartX = Px3
                                .StartY = Py3
                                .EndX = Px4
                                .EndY = Py4
                            End With
                            AxGraphicContext1.LineSegment()

                            'Two horizontal line
                            With AxGraphicContext1.DrawingRegion
                                .StartX = Px1
                                .StartY = Py1
                                .EndX = Px4
                                .EndY = Py4
                            End With
                            AxGraphicContext1.LineSegment()
                            With AxGraphicContext1.DrawingRegion
                                .StartX = Px2
                                .StartY = Py2
                                .EndX = Px3
                                .EndY = Py3
                            End With
                            AxGraphicContext1.LineSegment()
                            ClearanceRectangle()
                            ChipEdgePoints_form.TextBox_Score.Text = (AxMeasurement7.Results.Item(1).Score + AxMeasurement8.Results.Item(1).Score + AxMeasurement9.Results.Item(1).Score + AxMeasurement10.Results.Item(1).Score + AxMeasurement11.Results.Item(1).Score) / 5
                        Catch ex As Exception
                            ExceptionDisplay(ex)
                        End Try
                    ElseIf _
                        ((ChipEdgePoints_form.TextBox_SizeX.Text) * (1 - ChipEdgePoints_form.ValueSize.Value / 100)) < WidthPointX And _
                        WidthPointX < ((ChipEdgePoints_form.TextBox_SizeX.Text) * (1 + ChipEdgePoints_form.ValueSize.Value / 100)) And _
                        ((ChipEdgePoints_form.TextBox_SizeY.Text) * (1 - ChipEdgePoints_form.ValueSize.Value / 100)) < WidthPointY And _
                        WidthPointY < ((ChipEdgePoints_form.TextBox_SizeY.Text) * (1 + ChipEdgePoints_form.ValueSize.Value / 100)) And _
                        ChipPointRot < ((ChipEdgePoints_form.ValueRot.Value)) And _
                        (-(ChipEdgePoints_form.ValueRot.Value) < ChipPointRot) Then
                        With AxGraphicContext1
                            .FontScaleX = 1
                            .FontScaleY = 1
                            .DrawingRegion.StartX = ChipPointPosX / PixelSizeX - 100
                            .DrawingRegion.StartY = ChipPointPosY / PixelSizeY - 100
                            .DrawingRegion.EndX = DisplayWidth
                            .DrawingRegion.EndY = DisplayHeight
                            .CtlText("Position Out of Range!!")
                        End With

                    ElseIf _
                    ((ChipEdgePoints_form.TextBox_PosX.Text) * (1 - ChipEdgePoints_form.ValuePos.Value / 100)) < ChipPointPosX And _
                    ChipPointPosX < ((ChipEdgePoints_form.TextBox_PosX.Text) * (1 + ChipEdgePoints_form.ValuePos.Value / 100)) And _
                    ((ChipEdgePoints_form.TextBox_PosY.Text) * (1 - ChipEdgePoints_form.ValuePos.Value / 100)) < ChipPointPosY And _
                    ChipPointPosY < ((ChipEdgePoints_form.TextBox_PosY.Text) * (1 + ChipEdgePoints_form.ValuePos.Value / 100)) And _
                    ChipPointRot < ((ChipEdgePoints_form.ValueRot.Value)) And _
                    (-(ChipEdgePoints_form.ValueRot.Value) < ChipPointRot) Then
                        With AxGraphicContext1
                            .FontScaleX = 1
                            .FontScaleY = 1
                            .DrawingRegion.StartX = ChipPointPosX / PixelSizeX - 100
                            .DrawingRegion.StartY = ChipPointPosY / PixelSizeY - 100
                            .DrawingRegion.EndX = DisplayWidth
                            .DrawingRegion.EndY = DisplayHeight
                        End With
                        AxGraphicContext1.CtlText("Size Out of Range!!")
                    ElseIf _
                    ((ChipEdgePoints_form.TextBox_SizeX.Text) * (1 - ChipEdgePoints_form.ValueSize.Value / 100)) < WidthPointX And _
                    WidthPointX < ((ChipEdgePoints_form.TextBox_SizeX.Text) * (1 + ChipEdgePoints_form.ValueSize.Value / 100)) And _
                    ((ChipEdgePoints_form.TextBox_SizeY.Text) * (1 - ChipEdgePoints_form.ValueSize.Value / 100)) < WidthPointY And _
                    WidthPointY < ((ChipEdgePoints_form.TextBox_SizeY.Text) * (1 + ChipEdgePoints_form.ValueSize.Value / 100)) And _
                    ((ChipEdgePoints_form.TextBox_PosX.Text) * (1 - ChipEdgePoints_form.ValuePos.Value / 100)) < ChipPointPosX And _
                    ChipPointPosX < ((ChipEdgePoints_form.TextBox_PosX.Text) * (1 + ChipEdgePoints_form.ValuePos.Value / 100)) And _
                    ((ChipEdgePoints_form.TextBox_PosY.Text) * (1 - ChipEdgePoints_form.ValuePos.Value / 100)) < ChipPointPosY And _
                    ChipPointPosY < ((ChipEdgePoints_form.TextBox_PosY.Text) * (1 + ChipEdgePoints_form.ValuePos.Value / 100)) Then
                        With AxGraphicContext1
                            .FontScaleX = 1
                            .FontScaleY = 1
                            .DrawingRegion.StartX = ChipPointPosX / PixelSizeX - 100
                            .DrawingRegion.StartY = ChipPointPosY / PixelSizeY - 100
                            .DrawingRegion.EndX = DisplayWidth
                            .DrawingRegion.EndY = DisplayHeight
                        End With
                        Dim text As String = ("Chip Rotated" + " " & ChipPointRot.ToString & " deg")
                        AxGraphicContext1.CtlText(text)
                    ElseIf _
                    ((ChipEdgePoints_form.TextBox_SizeX.Text) * (1 - ChipEdgePoints_form.ValueSize.Value / 100)) < WidthPointX And _
                    WidthPointX < ((ChipEdgePoints_form.TextBox_SizeX.Text) * (1 + ChipEdgePoints_form.ValueSize.Value / 100)) And _
                    ((ChipEdgePoints_form.TextBox_SizeY.Text) * (1 - ChipEdgePoints_form.ValueSize.Value / 100)) < WidthPointY And _
                    WidthPointY < ((ChipEdgePoints_form.TextBox_SizeY.Text) * (1 + ChipEdgePoints_form.ValueSize.Value / 100)) Then
                        With AxGraphicContext1
                            .FontScaleX = 1
                            .FontScaleY = 1
                            .DrawingRegion.StartX = ChipPointPosX / PixelSizeX - 100
                            .DrawingRegion.StartY = ChipPointPosY / PixelSizeY - 100
                            .DrawingRegion.EndX = DisplayWidth
                            .DrawingRegion.EndY = DisplayHeight
                            .CtlText("Position and Rotation Angle Out of Range!!")
                        End With
                    ElseIf _
                        ChipPointRot < ((ChipEdgePoints_form.ValueRot.Value)) And _
                        (-(ChipEdgePoints_form.ValueRot.Value) < ChipPointRot) Then
                        With AxGraphicContext1
                            .FontScaleX = 1
                            .FontScaleY = 1
                            .DrawingRegion.StartX = ChipPointPosX / PixelSizeX - 100
                            .DrawingRegion.StartY = ChipPointPosY / PixelSizeY - 100
                            .DrawingRegion.EndX = DisplayWidth
                            .DrawingRegion.EndY = DisplayHeight
                            .CtlText("Position and Size Out of Range!!")
                        End With
                    ElseIf _
                    ((ChipEdgePoints_form.TextBox_PosX.Text) * (1 - ChipEdgePoints_form.ValuePos.Value / 100)) < ChipPointPosX And _
                    ChipPointPosX < ((ChipEdgePoints_form.TextBox_PosX.Text) * (1 + ChipEdgePoints_form.ValuePos.Value / 100)) And _
                    ((ChipEdgePoints_form.TextBox_PosY.Text) * (1 - ChipEdgePoints_form.ValuePos.Value / 100)) < ChipPointPosY And _
                    ChipPointPosY < ((ChipEdgePoints_form.TextBox_PosY.Text) * (1 + ChipEdgePoints_form.ValuePos.Value / 100)) Then
                        With AxGraphicContext1
                            .FontScaleX = 1
                            .FontScaleY = 1
                            .DrawingRegion.StartX = ChipPointPosX / PixelSizeX - 100
                            .DrawingRegion.StartY = ChipPointPosY / PixelSizeY - 100
                            .DrawingRegion.EndX = DisplayWidth
                            .DrawingRegion.EndY = DisplayHeight
                        End With
                        Dim text As String = ("Size and Rotation Angle Out of Range!!")
                        AxGraphicContext1.CtlText(text)
                    Else
                        With AxGraphicContext1
                            .FontScaleX = 1
                            .FontScaleY = 1
                            .DrawingRegion.StartX = ChipPointPosX / PixelSizeX - 100
                            .DrawingRegion.StartY = ChipPointPosY / PixelSizeY - 100
                            .DrawingRegion.EndX = DisplayWidth
                            .DrawingRegion.EndY = DisplayHeight
                        End With
                        Dim text As String = ("Chip's Edge not found")
                        AxGraphicContext1.CtlText(text)
                    End If
                ElseIf Chip_QC_SS = 3 Then
                    If _
                    ((TextboxSizeX.Text) * (1 - ValueSize.Value / 100)) < WidthPointX And _
                    WidthPointX < ((TextboxSizeX.Text) * (1 + ValueSize.Value / 100)) And _
                    ((TextboxSizeY.Text) * (1 - ValueSize.Value / 100)) < WidthPointY And _
                    WidthPointY < ((TextboxSizeY.Text) * (1 + ValueSize.Value / 100)) And _
                    ((TextboxPosX.Text) * (1 - ValuePos.Value / 100)) < ChipPointPosX And _
                    ChipPointPosX < ((TextboxPosX.Text) * (1 + ValuePos.Value / 100)) And _
                    ((TextboxPosY.Text) * (1 - ValuePos.Value / 100)) < ChipPointPosY And _
                    ChipPointPosY < ((TextboxPosY.Text) * (1 + ValuePos.Value / 100)) And _
                    ChipPointRot < ((ValueRot.Value)) And _
                    (-(ValueRot.Value) < ChipPointRot) Then
                        Try
                            'Two vertical line
                            With AxGraphicContext1.DrawingRegion
                                .StartX = Px1
                                .StartY = Py1
                                .EndX = Px2
                                .EndY = Py2
                            End With
                            AxGraphicContext1.LineSegment()
                            With AxGraphicContext1.DrawingRegion
                                .StartX = Px3
                                .StartY = Py3
                                .EndX = Px4
                                .EndY = Py4
                            End With
                            AxGraphicContext1.LineSegment()

                            'Two horizontal line
                            With AxGraphicContext1.DrawingRegion
                                .StartX = Px1
                                .StartY = Py1
                                .EndX = Px4
                                .EndY = Py4
                            End With
                            AxGraphicContext1.LineSegment()
                            With AxGraphicContext1.DrawingRegion
                                .StartX = Px2
                                .StartY = Py2
                                .EndX = Px3
                                .EndY = Py3
                            End With
                            AxGraphicContext1.LineSegment()
                            TextboxScore.Text = (AxMeasurement7.Results.Item(1).Score + AxMeasurement8.Results.Item(1).Score + AxMeasurement9.Results.Item(1).Score + AxMeasurement10.Results.Item(1).Score + AxMeasurement11.Results.Item(1).Score) / 5
                        Catch ex As Exception
                            ExceptionDisplay(ex)
                        End Try
                    ElseIf _
                        ((TextboxSizeX.Text) * (1 - ValueSize.Value / 100)) < WidthPointX And _
                        WidthPointX < ((TextboxSizeX.Text) * (1 + ValueSize.Value / 100)) And _
                        ((TextboxSizeY.Text) * (1 - ValueSize.Value / 100)) < WidthPointY And _
                        WidthPointY < ((TextboxSizeY.Text) * (1 + ValueSize.Value / 100)) And _
                        ChipPointRot < ((ValueRot.Value)) And _
                        (-(ValueRot.Value) < ChipPointRot) Then
                        With AxGraphicContext1
                            .FontScaleX = 1
                            .FontScaleY = 1
                            .DrawingRegion.StartX = ChipPointPosX / PixelSizeX - 100
                            .DrawingRegion.StartY = ChipPointPosY / PixelSizeY - 100
                            .DrawingRegion.EndX = DisplayWidth
                            .DrawingRegion.EndY = DisplayHeight
                            .CtlText("Position Out of Range!!")
                        End With

                    ElseIf _
                    ((TextboxPosX.Text) * (1 - ValuePos.Value / 100)) < ChipPointPosX And _
                    ChipPointPosX < ((TextboxPosX.Text) * (1 + ValuePos.Value / 100)) And _
                    ((TextboxPosY.Text) * (1 - ValuePos.Value / 100)) < ChipPointPosY And _
                    ChipPointPosY < ((TextboxPosY.Text) * (1 + ValuePos.Value / 100)) And _
                    ChipPointRot < ((ValueRot.Value)) And _
                    (-(ValueRot.Value) < ChipPointRot) Then
                        With AxGraphicContext1
                            .FontScaleX = 1
                            .FontScaleY = 1
                            .DrawingRegion.StartX = ChipPointPosX / PixelSizeX - 100
                            .DrawingRegion.StartY = ChipPointPosY / PixelSizeY - 100
                            .DrawingRegion.EndX = DisplayWidth
                            .DrawingRegion.EndY = DisplayHeight
                        End With
                        AxGraphicContext1.CtlText("Size Out of Range!!")
                    ElseIf _
                    ((TextboxSizeX.Text) * (1 - ValueSize.Value / 100)) < WidthPointX And _
                    WidthPointX < ((TextboxSizeX.Text) * (1 + ValueSize.Value / 100)) And _
                    ((TextboxSizeY.Text) * (1 - ValueSize.Value / 100)) < WidthPointY And _
                    WidthPointY < ((TextboxSizeY.Text) * (1 + ValueSize.Value / 100)) And _
                    ((TextboxPosX.Text) * (1 - ValuePos.Value / 100)) < ChipPointPosX And _
                    ChipPointPosX < ((TextboxPosX.Text) * (1 + ValuePos.Value / 100)) And _
                    ((TextboxPosY.Text) * (1 - ValuePos.Value / 100)) < ChipPointPosY And _
                    ChipPointPosY < ((TextboxPosY.Text) * (1 + ValuePos.Value / 100)) Then
                        With AxGraphicContext1
                            .FontScaleX = 1
                            .FontScaleY = 1
                            .DrawingRegion.StartX = ChipPointPosX / PixelSizeX - 100
                            .DrawingRegion.StartY = ChipPointPosY / PixelSizeY - 100
                            .DrawingRegion.EndX = DisplayWidth
                            .DrawingRegion.EndY = DisplayHeight
                        End With
                        Dim text As String = ("Chip Rotated" + " " & ChipPointRot.ToString & " deg")
                        AxGraphicContext1.CtlText(text)
                    ElseIf _
                    ((TextboxSizeX.Text) * (1 - ValueSize.Value / 100)) < WidthPointX And _
                    WidthPointX < ((TextboxSizeX.Text) * (1 + ValueSize.Value / 100)) And _
                    ((TextboxSizeY.Text) * (1 - ValueSize.Value / 100)) < WidthPointY And _
                    WidthPointY < ((TextboxSizeY.Text) * (1 + ValueSize.Value / 100)) Then
                        With AxGraphicContext1
                            .FontScaleX = 1
                            .FontScaleY = 1
                            .DrawingRegion.StartX = ChipPointPosX / PixelSizeX - 100
                            .DrawingRegion.StartY = ChipPointPosY / PixelSizeY - 100
                            .DrawingRegion.EndX = DisplayWidth
                            .DrawingRegion.EndY = DisplayHeight
                            .CtlText("Position and Rotation Angle Out of Range!!")
                        End With
                    ElseIf _
                        ChipPointRot < ((ValueRot.Value)) And _
                        (-(ValueRot.Value) < ChipPointRot) Then
                        With AxGraphicContext1
                            .FontScaleX = 1
                            .FontScaleY = 1
                            .DrawingRegion.StartX = ChipPointPosX / PixelSizeX - 100
                            .DrawingRegion.StartY = ChipPointPosY / PixelSizeY - 100
                            .DrawingRegion.EndX = DisplayWidth
                            .DrawingRegion.EndY = DisplayHeight
                            .CtlText("Position and Size Out of Range!!")
                        End With
                    ElseIf _
                    ((TextboxPosX.Text) * (1 - ValuePos.Value / 100)) < ChipPointPosX And _
                    ChipPointPosX < ((TextboxPosX.Text) * (1 + ValuePos.Value / 100)) And _
                    ((TextboxPosY.Text) * (1 - ValuePos.Value / 100)) < ChipPointPosY And _
                    ChipPointPosY < ((TextboxPosY.Text) * (1 + ValuePos.Value / 100)) Then
                        With AxGraphicContext1
                            .FontScaleX = 1
                            .FontScaleY = 1
                            .DrawingRegion.StartX = ChipPointPosX / PixelSizeX - 100
                            .DrawingRegion.StartY = ChipPointPosY / PixelSizeY - 100
                            .DrawingRegion.EndX = DisplayWidth
                            .DrawingRegion.EndY = DisplayHeight
                        End With
                        Dim text As String = ("Size and Rotation Angle Out of Range!!")
                        AxGraphicContext1.CtlText(text)
                    Else
                        With AxGraphicContext1
                            .FontScaleX = 1
                            .FontScaleY = 1
                            .DrawingRegion.StartX = ChipPointPosX / PixelSizeX - 100
                            .DrawingRegion.StartY = ChipPointPosY / PixelSizeY - 100
                            .DrawingRegion.EndX = DisplayWidth
                            .DrawingRegion.EndY = DisplayHeight
                        End With
                        Dim text As String = ("Chip's Edge not found")
                        AxGraphicContext1.CtlText(text)
                    End If
                End If
            End If
        End If
    End Function
    Function ChipPointInput(ByVal PointX1, ByVal PointY1, ByVal PointX2, ByVal PointY2, ByVal PointX3, ByVal PointY3, ByVal PointX4, ByVal PointY4, ByVal PointX5, ByVal PointY5)
        SizeXPoint = Abs(PointX5 - PointX3)
        SizeYPoint = Abs(PointY1 - PointY4)
        PosXPoint = ((PointX5 + PointX3) / 2)
        PosYPoint = ((PointY1 + PointY4) / 2)
        ChipEdgePoints_form.TextBox_SizeX.Text = Format(SizeXPoint * PixelSizeX, "##0.000")
        ChipEdgePoints_form.TextBox_SizeY.Text = Format(SizeYPoint * PixelSizeY, "##0.000")
        ChipEdgePoints_form.TextBox_PosX.Text = Format(PosXPoint * PixelSizeX, "##0.000")
        ChipEdgePoints_form.TextBox_PosY.Text = Format(PosYPoint * PixelSizeY, "##0.000")

        'Ask Robot to move to the centre of the chip
        MeasurementPoint_Settings()
        If ChipEdgeDrawingEnabled() Then
            ChipEdgePoints_form.Show()
        End If
        DisableChipEdgeDrawing()
    End Function
    Function ChipPointOutPut(ByVal Chip_QC_SS As Integer)
        Px1 = Px1 * PixelSizeX 'mm
        Py1 = Py1 * PixelSizeY
        Px2 = Px2 * PixelSizeX
        Py2 = Py2 * PixelSizeY
        Px3 = Px3 * PixelSizeX
        Py3 = Py3 * PixelSizeY
        Px4 = Px4 * PixelSizeX
        Py4 = Py4 * PixelSizeY

        If Px1 < ChipPointPosX Then
            If Py1 < ChipPointPosY Then
                'top left
                Output(0) = Px1
                Output(1) = Py1
            Else
                'bottom left
                Output(6) = Px1
                Output(7) = Py1
            End If
        Else
            If Py1 < ChipPointPosY Then
                'top right
                Output(2) = Px1
                Output(3) = Py1
            Else
                'bottom right
                Output(4) = Px1
                Output(5) = Py1
            End If
        End If

        If Px2 < ChipPointPosX Then
            If Py2 < ChipPointPosY Then
                'top left
                Output(0) = Px2
                Output(1) = Py2
            Else
                'bottom left
                Output(6) = Px2
                Output(7) = Py2
            End If
        Else
            If Py2 < ChipPointPosY Then
                'top right
                Output(2) = Px2
                Output(3) = Py2
            Else
                'bottom right
                Output(4) = Px2
                Output(5) = Py2
            End If
        End If

        If Px3 < ChipPointPosX Then
            If Py3 < ChipPointPosY Then
                'top left
                Output(0) = Px3
                Output(1) = Py3
            Else
                'bottom left
                Output(6) = Px3
                Output(7) = Py3
            End If
        Else
            If Py3 < ChipPointPosY Then
                'top right
                Output(2) = Px3
                Output(3) = Py3
            Else
                'bottom right
                Output(4) = Px3
                Output(5) = Py3
            End If
        End If

        If Px4 < ChipPointPosX Then
            If Py4 < ChipPointPosY Then
                'top left
                Output(0) = Px4
                Output(1) = Py4
            Else
                'bottom left
                Output(6) = Px4
                Output(7) = Py4
            End If
        Else
            If Py4 < ChipPointPosY Then
                'top right
                Output(2) = Px4
                Output(3) = Py4
            Else
                'bottom right
                Output(4) = Px4
                Output(5) = Py4
            End If
        End If

        Output(0) = Output(0) - DisplayCenterXPosition * PixelSizeX
        Output(1) = Output(1) - DisplayCenterYPosition * PixelSizeY
        Output(2) = Output(2) - DisplayCenterXPosition * PixelSizeX
        Output(3) = Output(3) - DisplayCenterYPosition * PixelSizeY
        Output(4) = Output(4) - DisplayCenterXPosition * PixelSizeX
        Output(5) = Output(5) - DisplayCenterYPosition * PixelSizeY
        Output(6) = Output(6) - DisplayCenterXPosition * PixelSizeX
        Output(7) = Output(7) - DisplayCenterYPosition * PixelSizeY

        If Chip_QC_SS = 1 Then
            ChipEdgePoints_form.ChipEdgeOutput(Output(0), Output(1), Output(2), Output(3), Output(4), Output(5), Output(6), Output(7), MainEdge)
        ElseIf Chip_QC_SS = 3 Then
            Dim Center_X = (Output(0) + Output(2) + Output(4) + Output(6)) / 4
            Dim Center_Y = (Output(1) + Output(3) + Output(5) + Output(7)) / 4
            MsgBox(Center_X & " " & Center_Y) 'output offset
        End If
    End Function
#End Region
#Region "Measurement"
    Dim FirstPoint, SecondPoint, DistanceX, DistanceY, MAngle As Double
    Dim Measurement_flag As Boolean = False
    Dim FirstClick As Integer = 0
    Dim SecondClick As Integer = 0
    Dim Mclick As Boolean = True
    Dim timer As Integer = 0
    Function Distance(ByVal x, ByVal y)
        Try
            TextBox1.Text = 0

            AxMeasurement12.MeasurementList.SelectAll()
            AxMeasurement12.PixelAspectRatio.Value = 1

            If Mclick = True Then
                FirstClick = 1
                Mclick = False
            ElseIf Mclick = False Then
                SecondClick = 1
                Mclick = True
            End If
            If FirstClick = 1 Then
                FirstPoint = AxMeasurement12.Markers.Add(Matrox.ActiveMIL.Measurement.MeasMarkerTypeConstants.measPoint) 'MILMEASUREMENTLib.MeasMarkerTypeConstants.measPoint)
                AxMeasurement12.Markers.Item(FirstPoint).Position.X = x
                AxMeasurement12.Markers.Item(FirstPoint).Position.Y = y
                With AxGraphicContext1.DrawingRegion
                    .SizeX = 10
                    .SizeY = 10
                    .CenterX = x
                    .CenterY = y
                End With
                AxGraphicContext1.Cross()
                FirstClick = 0

            ElseIf SecondClick = 1 Then
                SecondPoint = AxMeasurement12.Markers.Add(Matrox.ActiveMIL.Measurement.MeasMarkerTypeConstants.measPoint) 'MILMEASUREMENTLib.MeasMarkerTypeConstants.measPoint)
                AxMeasurement12.Markers.Item(SecondPoint).Position.X = x
                AxMeasurement12.Markers.Item(SecondPoint).Position.Y = y
                With AxGraphicContext1.DrawingRegion
                    .SizeX = 10
                    .SizeY = 10
                    .CenterX = x
                    .CenterY = y
                End With
                AxGraphicContext1.Cross()
                With AxGraphicContext1.DrawingRegion
                    .StartX = AxMeasurement12.Markers.Item(FirstPoint).Position.X
                    .StartY = AxMeasurement12.Markers.Item(FirstPoint).Position.Y
                    .EndX = x
                    .EndY = y
                End With
                AxGraphicContext1.LineSegment()
                SecondClick = 0

                If AxMeasurement12.Markers.Item(1).IsFound = True And AxMeasurement12.Markers.Item(2).IsFound = True Then
                    AxMeasurement12.Calculate(FirstPoint, SecondPoint)
                    DistanceX = AxMeasurement12.Results.Item(1).DistanceX()
                    DistanceY = AxMeasurement12.Results.Item(1).DistanceY()
                    MAngle = AxMeasurement12.Results.Item(1).Angle
                End If

                Timer_Histo.Start()
                Measurement_flag = False

            End If
        Catch ex As Exception
            ExceptionDisplay(ex)
        End Try
    End Function
    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer_Histo.Tick
        timer = timer + 1
        If timer = 3 Then
            DisplayIndicator()
            Timer_Histo.Stop()
            timer = 0
        End If
    End Sub

#End Region
#Region "QC"
    Dim QC_firstTime As Integer = 1
    Dim OldRegionX As Double = DisplayCenterXPosition
    Dim OldRegionY As Double = DisplayCenterYPosition
    Dim OldMRoix As Double = MROIx
    Dim OldMRoiy As Double = MROIy
    Function ClearImage16()
        If AxImage16.IsAllocated = True Then
            AxImage16.Free()
        End If
    End Function

    Function QC(ByVal ValueBinarized As Integer, ByVal ValueClose As Integer, ByVal ValueOpen As Integer)
        Try

            FreeIfAllocated(AxImage12)
            With AxImage12
                .SizeX = MROIx
                .SizeY = MROIy
            End With
            AxImage12.Allocate()

            FreeIfAllocated(AxImage13)
            With AxImage13
                .SizeX = MROIx
                .SizeY = MROIy
            End With
            AxImage13.Allocate()

            FreeIfAllocated(AxImage16)
            With AxImage16.ChildRegion
                .OffsetX = DisplayCenterXPosition - MROIx / 2
                .OffsetY = DisplayCenterYPosition - MROIy / 2
                .SizeX = MROIx
                .SizeY = MROIy
            End With
            AxImage16.Allocate()

            FreeIfAllocated(AxImage17)
            With AxImage17
                .SizeX = MROIx
                .SizeY = MROIy
            End With
            AxImage17.Allocate()

            'FreeIfAllocated(AxImageProcessing2)
            'AxImageProcessing2.Allocate()

            'FreeIfAllocated(AxImageProcessing6)
            'AxImageProcessing6.Allocate()

            'FreeIfAllocated(AxImageProcessing9)
            'AxImageProcessing9.Allocate()

            AxImageProcessing9.Binarize(Matrox.ActiveMIL.ImageProcessing.ImpConditionConstants.impGreaterOrEqualTo, ValueBinarized)
            If ValueClose = 0 Then
                AxImageProcessing2.Binarize(Matrox.ActiveMIL.ImageProcessing.ImpConditionConstants.impGreaterOrEqualTo, ValueBinarized)
            Else
                AxImageProcessing2.Close(ValueClose, Matrox.ActiveMIL.ImageProcessing.ImpProcessingModeConstants.impBinary)
            End If
            If ValueOpen = 0 Then
                AxImageProcessing6.Binarize(Matrox.ActiveMIL.ImageProcessing.ImpConditionConstants.impGreaterOrEqualTo, ValueBinarized)
            Else
                AxImageProcessing6.Open(ValueOpen, Matrox.ActiveMIL.ImageProcessing.ImpProcessingModeConstants.impBinary)
            End If
            AxDisplay1.OverlayImage.CopyRegion(AxImageProcessing6.Destination1, Matrox.ActiveMIL.ImBandConstants.imAllBands, 0, 0, Matrox.ActiveMIL.ImBandConstants.imAllBands, CInt(DisplayCenterXPosition - MROIx / 2), CInt(DisplayCenterYPosition - MROIy / 2), CInt(MROIx), CInt(MROIy))

        Catch ex As Exception
            ExceptionDisplay(ex) 'copyregion error
        End Try

    End Function

    Sub ApplyBlobAnalysis_Setting(ByVal BW)
        Try
            With AxBlobAnalysis2
                .IdentifierBlobType = BlobAnalysis.BlobIdentifierBlobTypeConstants.blobIndividual
                .IdentifierPixelType = BlobAnalysis.BlobPixelTypeConstants.blobBinary
                .Lattice = BlobAnalysis.BlobLatticeConstants.blob8Connected
                .NumberOfFeretAngles() = 8
                .PixelAspectRatio = 1
                If BW = True Then
                    .ForegroundPixelValue = BlobAnalysis.BlobForegroundPixelValueConstants.blobZero
                Else
                    .ForegroundPixelValue = BlobAnalysis.BlobForegroundPixelValueConstants.blobNonZero
                End If
                .FeatureList.Area = True
                .FeatureList.Compactness = True
                .FeatureList.Roughness = True
                .FeatureList.Box.All = True
                .FeatureList.CenterOfGravity.All = BlobAnalysis.BlobTrueFalseBinGrayConstants.blobTrue
                .FeatureList.ContactPoints.All = True
                .FeatureList.ConvexPerimeter = True
                .FeatureList.FeretElongation = True
                .FeatureList.MaximumFeretAngle = True
                .FeatureList.MinimumFeretAngle = True
                .FeatureList.MaximumFeretDiameter = True
                .FeatureList.MinimumFeretDiameter = True
                .FeatureList.MeanFeretDiameter = True
                .FeatureList.NumberOfHoles = True
                .FeatureList.Perimeter = True
            End With
        Catch ex As Exception
            ExceptionDisplay(ex)
        End Try
    End Sub

    Function BlobAnalysisCondition(ByVal i As Integer, ByVal ValueCompactness As Double, ByVal ValueRoughness As Double)
        Return AxBlobAnalysis2.Results.Item(i).Roughness < ValueRoughness And AxBlobAnalysis2.Results.Item(i).Compactness < ValueCompactness And Not (AxBlobAnalysis2.Results.Item(i).BoxYMaximum + DisplayCenterYPosition - MROIy / 2 < DisplayCenterYPosition + MROIy / 2 + 5 And _
                                    AxBlobAnalysis2.Results.Item(i).BoxYMaximum + DisplayCenterYPosition - MROIy / 2 > DisplayCenterYPosition + MROIy / 2 - 5) _
                                    Or (AxBlobAnalysis2.Results.Item(i).BoxYMinimum + DisplayCenterYPosition - MROIy / 2 > DisplayCenterYPosition - MROIy / 2 - 5 And _
                                    AxBlobAnalysis2.Results.Item(i).BoxYMinimum + DisplayCenterYPosition - MROIy / 2 < DisplayCenterYPosition - MROIy / 2 + 5) _
                                    Or (AxBlobAnalysis2.Results.Item(i).BoxXMinimum + DisplayCenterXPosition - MROIx / 2 < DisplayCenterXPosition - MROIx / 2 + 5 And _
                                    AxBlobAnalysis2.Results.Item(i).BoxXMinimum + DisplayCenterXPosition - MROIx / 2 > DisplayCenterXPosition - MROIx / 2 - 5) _
                                    Or (AxBlobAnalysis2.Results.Item(i).BoxXMaximum + DisplayCenterXPosition - MROIx / 2 > DisplayCenterXPosition + MROIx / 2 - 5 And _
                                    AxBlobAnalysis2.Results.Item(i).BoxXMaximum + DisplayCenterXPosition - MROIx / 2 < DisplayCenterXPosition + MROIx / 2 + 5)
    End Function

    Function BlobAnalysisNC(ByVal BlackDot As Boolean, ByVal ValueMinArea As Integer, ByVal ValueMaxArea As Integer, ByVal ValueRoughness As Double, ByVal ValueCompactness As Double, ByRef NC_OffX As Double, ByRef NC_OffY As Double) As Boolean
        Try

            Dim nTotalBlobs, BlobAreaFilter, BlobCompactFilter As Integer
            'FreeIfAllocated(AxBlobAnalysis2)
            'AxBlobAnalysis2.Allocate()
            AxBlobAnalysis2.Calculate()

            If AxBlobAnalysis2.Filters.Count > 0 Then
                For i As Integer = 1 To AxBlobAnalysis2.Filters.Count
                    AxBlobAnalysis2.Filters.Remove(i)
                Next
            End If

            BlobAreaFilter = AxBlobAnalysis2.Filters.Add(BlobAnalysis.BlobFilterTypeConstants.blobInclude, BlobAnalysis.BlobFeatureConstants.blobArea, BlobAnalysis.BlobFilterConditionConstants.blobInRange, ValueMinArea, ValueMaxArea)
            AxBlobAnalysis2.ApplyFilter(BlobAreaFilter, False)
            nTotalBlobs = AxBlobAnalysis2.Results.Count

            Dim j As Integer = 0
            Dim QC_X, QC_Y, QC_Dia As Double

            If nTotalBlobs = 1 Then
                If BlobAnalysisCondition(1, ValueCompactness, ValueRoughness) Then
                    j = j + 1
                    With AxMGraphicContext1.DrawingRegion
                        .CenterX = AxBlobAnalysis2.Results.Item(1).CenterOfGravityX + DisplayCenterXPosition - MROIx / 2
                        .CenterY = AxBlobAnalysis2.Results.Item(1).CenterOfGravityY + DisplayCenterYPosition - MROIy / 2
                        .SizeX = 10
                        .SizeY = 10
                    End With
                    AxMGraphicContext1.Cross()
                    With AxMGraphicContext1.DrawingRegion
                        .CenterX = AxBlobAnalysis2.Results.Item(1).CenterOfGravityX + DisplayCenterXPosition - MROIx / 2
                        .CenterY = AxBlobAnalysis2.Results.Item(1).CenterOfGravityY + DisplayCenterYPosition - MROIy / 2
                        .SizeX = AxBlobAnalysis2.Results.Item(1).MeanFeretDiameter
                        .SizeY = AxBlobAnalysis2.Results.Item(1).MeanFeretDiameter
                    End With
                    AxMGraphicContext1.Arc(0, 360, False)
                    QC_X = AxBlobAnalysis2.Results.Item(1).CenterOfGravityX + DisplayCenterXPosition - MROIx / 2
                    QC_Y = AxBlobAnalysis2.Results.Item(1).CenterOfGravityY + DisplayCenterYPosition - MROIy / 2
                    QC_Dia = AxBlobAnalysis2.Results.Item(1).MeanFeretDiameter
                    With AxMGraphicContext1.DrawingRegion
                        .CenterX = AxBlobAnalysis2.Results.Item(1).CenterOfGravityX + DisplayCenterXPosition - MROIx / 2
                        .CenterY = AxBlobAnalysis2.Results.Item(1).CenterOfGravityY + DisplayCenterYPosition - MROIy / 2 - MROIx / 4
                        .SizeX = 200
                        .SizeY = 20
                    End With
                    AxMGraphicContext1.CtlText("Dot diameter is: " + (QC_Dia * PixelSizeX).ToString + "mm")
                    DelayClearingImage()
                End If
            Else
                Dim i As Integer = 1
                For i = 1 To nTotalBlobs
                    If BlobAnalysisCondition(i, ValueCompactness, ValueRoughness) Then
                        j = j + 1
                        With AxMGraphicContext1.DrawingRegion
                            .CenterX = AxBlobAnalysis2.Results.Item(i).CenterOfGravityX + DisplayCenterXPosition - MROIx / 2
                            .CenterY = AxBlobAnalysis2.Results.Item(i).CenterOfGravityY + DisplayCenterYPosition - MROIy / 2
                            .SizeX = 10
                            .SizeY = 10
                        End With
                        AxMGraphicContext1.Cross()
                        With AxMGraphicContext1.DrawingRegion
                            .CenterX = AxBlobAnalysis2.Results.Item(i).CenterOfGravityX + DisplayCenterXPosition - MROIx / 2
                            .CenterY = AxBlobAnalysis2.Results.Item(i).CenterOfGravityY + DisplayCenterYPosition - MROIy / 2
                            .SizeX = AxBlobAnalysis2.Results.Item(i).MeanFeretDiameter
                            .SizeY = AxBlobAnalysis2.Results.Item(i).MeanFeretDiameter
                        End With
                        AxMGraphicContext1.Arc(0, 360, False)
                        QC_X = AxBlobAnalysis2.Results.Item(i).CenterOfGravityX + DisplayCenterXPosition - MROIx / 2
                        QC_Y = AxBlobAnalysis2.Results.Item(i).CenterOfGravityY + DisplayCenterYPosition - MROIy / 2
                        QC_Dia = AxBlobAnalysis2.Results.Item(i).MeanFeretDiameter
                        With AxMGraphicContext1.DrawingRegion
                            .CenterX = AxBlobAnalysis2.Results.Item(i).CenterOfGravityX + DisplayCenterXPosition - MROIx / 2
                            .CenterY = AxBlobAnalysis2.Results.Item(i).CenterOfGravityY + DisplayCenterYPosition - MROIy / 2 - 20
                            .SizeX = 200
                            .SizeY = 20
                        End With
                        AxMGraphicContext1.CtlText("Dot diameter is: " + (QC_Dia * PixelSizeX).ToString + "mm")
                    End If
                Next i
                DelayClearingImage()
            End If

            AxBlobAnalysis2.Filters.Remove(BlobAreaFilter)
            AxImage12.Clear(30)
            If j = 1 Then  'output
                Dim Del_X = (QC_X - DisplayCenterXPosition) * PixelSizeX
                Dim Del_Y = (QC_Y - DisplayCenterYPosition) * PixelSizeY
                Dim DistanceXY = Sqrt(Del_X ^ 2 + Del_Y ^ 2)
                NC_OffX = Del_X
                NC_OffY = Del_Y
                Return True
            ElseIf j > 1 Then
                Return False
            ElseIf j = 0 Then
                Return False
            End If
        Catch ex As Exception
            ExceptionDisplay(ex)
        End Try
    End Function

    Function BlobAnalysisQC(ByVal BlackDot As Boolean, ByVal ValueMinArea As Integer, ByVal ValueMaxArea As Integer, ByVal ValueRoughness As Double, ByVal ValueCompactness As Double) As Boolean
        Try
            Dim nTotalBlobs, BlobAreaFilter, BlobCompactFilter As Integer

            'FreeIfAllocated(AxBlobAnalysis2)
            'AxBlobAnalysis2.Allocate()
            AxBlobAnalysis2.Calculate()

            If AxBlobAnalysis2.Filters.Count > 0 Then
                For j As Integer = 1 To AxBlobAnalysis2.Filters.Count
                    AxBlobAnalysis2.Filters.Remove(j)
                Next
            End If

            BlobAreaFilter = AxBlobAnalysis2.Filters.Add(BlobAnalysis.BlobFilterTypeConstants.blobInclude, BlobAnalysis.BlobFeatureConstants.blobArea, BlobAnalysis.BlobFilterConditionConstants.blobInRange, ValueMinArea, ValueMaxArea)
            AxBlobAnalysis2.ApplyFilter(BlobAreaFilter, False)
            nTotalBlobs = AxBlobAnalysis2.Results.Count
            nTotalBlobs = AxBlobAnalysis2.Results.Count
            Dim QC_X, QC_Y, QC_Dia As Double
            Dim i As Integer = 1
            For i = 1 To nTotalBlobs
                QC_X = AxBlobAnalysis2.Results.Item(i).CenterOfGravityX + DisplayCenterXPosition - MROIx / 2
                QC_Y = AxBlobAnalysis2.Results.Item(i).CenterOfGravityY + DisplayCenterYPosition - MROIy / 2
                QC_Dia = AxBlobAnalysis2.Results.Item(i).MeanFeretDiameter
                'check if dot corresponds within compact, roughness values
                If BlobAnalysisCondition(1, ValueCompactness, ValueRoughness) Then
                    With AxMGraphicContext1.DrawingRegion
                        .CenterX = QC_X
                        .CenterY = QC_Y
                        .SizeX = 10
                        .SizeY = 10
                    End With
                    AxMGraphicContext1.Cross()
                    With AxMGraphicContext1.DrawingRegion
                        .CenterX = QC_X
                        .CenterY = QC_Y
                        .SizeX = QC_Dia
                        .SizeY = QC_Dia
                    End With
                    AxMGraphicContext1.Arc(0, 360, False)
                    With AxMGraphicContext1.DrawingRegion
                        .CenterX = QC_X
                        .CenterY = QC_Y - MROIx / 4
                        .SizeX = 200
                        .SizeY = 20
                    End With
                    If nTotalBlobs = 1 Then
                        DelayClearingImage()
                        AxMGraphicContext1.CtlText("Dot diameter is: " + (QC_Dia * PixelSizeX).ToString + "mm")
                    End If
                End If
            Next i

            AxBlobAnalysis2.Filters.Remove(BlobAreaFilter)
            AxImage12.Clear(30)
            If nTotalBlobs = 1 Then  'output
                QC_form.TextBox1.Text = "Found one object. Click <Ok> to save the settings."
                QC_form.GetVariables(DisplayCenterXPosition, DisplayCenterYPosition, MROIx, MROIy, QC_X, QC_Y, QC_Dia)
                QC_form.ValueDetectedDiameter.Text = QC_Dia * PixelSizeX
                QC_form.DotOutput((QC_X - DisplayCenterXPosition) * PixelSizeX, (QC_Y - DisplayCenterYPosition) * PixelSizeY, QC_Dia * PixelSizeX)
                QC_form.QCResults(True)
                Return True
            ElseIf nTotalBlobs > 1 Then
                QC_form.QCResults(False)
                QC_form.TextBox1.Text = "More than one object is found. Please re-adjust the settings or ROI"
                QC_form.ValueDetectedDiameter.Text = 0
                Return False
            ElseIf nTotalBlobs = 0 Then
                QC_form.QCResults(False)
                QC_form.TextBox1.Text = "No object is found. Please re-adjust the settings or ROI"
                QC_form.ValueDetectedDiameter.Text = 0
                Return False
            End If
        Catch ex As Exception
            ExceptionDisplay(ex)
        End Try
    End Function

    Function BlobAnalysisQC(ByVal BlackDot As Boolean, ByVal ValueMinArea As Integer, ByVal ValueMaxArea As Integer, ByVal ValueRoughness As Double, ByVal ValueCompactness As Double, ByVal Tolerance As Double, ByVal Diameter As Double) As Boolean
        Dim nTotalBlobs, BlobAreaFilter, BlobCompactFilter As Integer

        'FreeIfAllocated(AxBlobAnalysis2)
        'AxBlobAnalysis2.Allocate()
        AxBlobAnalysis2.Calculate()

        If AxBlobAnalysis2.Filters.Count > 0 Then
            For j As Integer = 1 To AxBlobAnalysis2.Filters.Count
                AxBlobAnalysis2.Filters.Remove(j)
            Next
        End If

        Try
            BlobAreaFilter = AxBlobAnalysis2.Filters.Add(BlobAnalysis.BlobFilterTypeConstants.blobInclude, BlobAnalysis.BlobFeatureConstants.blobArea, BlobAnalysis.BlobFilterConditionConstants.blobInRange, ValueMinArea, ValueMaxArea)
        Catch ex As Exception
            ExceptionDisplay(ex)
        End Try

        Try
            AxBlobAnalysis2.ApplyFilter(BlobAreaFilter, False)
        Catch ex As Exception
            ExceptionDisplay(ex)
        End Try

        nTotalBlobs = AxBlobAnalysis2.Results.Count
        Dim QC_X, QC_Y, QC_Dia As Double
        Dim i As Integer = 1
        For i = 1 To nTotalBlobs
            QC_X = AxBlobAnalysis2.Results.Item(i).CenterOfGravityX + DisplayCenterXPosition - MROIx / 2
            QC_Y = AxBlobAnalysis2.Results.Item(i).CenterOfGravityY + DisplayCenterYPosition - MROIy / 2
            QC_Dia = AxBlobAnalysis2.Results.Item(i).MeanFeretDiameter
            'check if dot corresponds within compact, roughness values and if the diameter is within the previously detected value with an error margin of the value of tolerance
            If BlobAnalysisCondition(1, ValueCompactness, ValueRoughness) And QC_Dia * PixelSizeX < Diameter + Tolerance And QC_Dia * PixelSizeX > Diameter - Tolerance Then
                With AxMGraphicContext1.DrawingRegion
                    .CenterX = QC_X
                    .CenterY = QC_Y
                    .SizeX = 10
                    .SizeY = 10
                End With
                AxMGraphicContext1.Cross()
                With AxMGraphicContext1.DrawingRegion
                    .CenterX = QC_X
                    .CenterY = QC_Y
                    .SizeX = QC_Dia
                    .SizeY = QC_Dia
                End With
                AxMGraphicContext1.Arc(0, 360, False)
                With AxMGraphicContext1.DrawingRegion
                    .CenterX = QC_X
                    .CenterY = QC_Y - MROIx / 4
                    .SizeX = 200
                    .SizeY = 20
                End With
                If nTotalBlobs = 1 Then
                    DelayClearingImage()
                    AxMGraphicContext1.CtlText("Dot diameter is: " + (QC_Dia * PixelSizeX).ToString + "mm")
                End If
            End If
        Next i

        AxBlobAnalysis2.Filters.Remove(BlobAreaFilter)
        AxImage12.Clear(30)
        Dim QCsize3decimalpoints As Double = CInt(QC_Dia * PixelSizeX * 1000.0) / 1000.0
        If nTotalBlobs = 1 Then  'output
            If QC_Dia * PixelSizeX < Diameter + Tolerance And QC_Dia * PixelSizeX > Diameter - Tolerance Then
                QC_form.QC_success = ": " + CStr(QCsize3decimalpoints) + " mm"
                QC_form.TextBox1.Text = "Found one object. Click <Ok> to save the settings."
                QC_form.GetVariables(DisplayCenterXPosition, DisplayCenterYPosition, MROIx, MROIy, QC_X, QC_Y, QC_Dia)
                QC_form.ValueDetectedDiameter.Text = QC_Dia * PixelSizeX
                QC_form.DotOutput((QC_X - DisplayCenterXPosition) * PixelSizeX, (QC_Y - DisplayCenterYPosition) * PixelSizeY, QC_Dia * PixelSizeX)
                QC_form.QCResults(True)
                Return True
            Else
                QC_form.QC_failure = "Dot size detected: " + CStr(QCsize3decimalpoints) + " mm" + ControlChars.CrLf + "		Desired diameter: " + CStr(Diameter) + " +/- " + CStr(Tolerance) + " mm."
                QC_form.QCResults(False)
                Return False
            End If
        ElseIf nTotalBlobs > 1 Then
            QC_form.QCResults(False)
            QC_form.QC_failure = "Noisy image. More than one blob found."
            QC_form.TextBox1.Text = "More than one object is found. Please re-adjust the settings or ROI"
            QC_form.ValueDetectedDiameter.Text = 0
            Return False
        ElseIf nTotalBlobs = 0 Then
            QC_form.QCResults(False)
            QC_form.QC_failure = "No dot detected."
            QC_form.TextBox1.Text = "No object is found. Please re-adjust the settings or ROI"
            QC_form.ValueDetectedDiameter.Text = 0
            Return False
        End If
    End Function

    Function DotNC(ByVal BlackDot As Boolean, ByVal ValueBinarized As Integer, ByVal ValueClose As Integer, ByVal ValueOpen As Integer, ByVal ValueMinArea As Double, ByVal ValueMaxArea As Double, ByVal ValueRoughness As Double, ByVal ValueCompactness As Double, ByRef NC_OffX As Double, ByRef NC_Offy As Double) As Boolean

        ValueMaxArea = PI * (ValueMaxArea ^ 2) / (PixelSizeX ^ 2)
        ValueMinArea = PI * (ValueMinArea ^ 2) / (PixelSizeX ^ 2)
        QC(ValueBinarized, ValueClose, ValueOpen)
        ApplyBlobAnalysis_Setting(BlackDot)
        AxImage16.Save("C:\nc-1.bmp")
        AxImage17.Save("C:\nc-2.bmp")
        AxImage13.Save("C:\nc-3.bmp")
        AxImage12.Save("C:\nc-4.bmp")
        Dim result As Boolean = BlobAnalysisNC(BlackDot, ValueMinArea, ValueMaxArea, ValueRoughness, ValueCompactness, NC_OffX, NC_Offy)
        AxImage16.Clear()
        AxImage17.Clear()
        AxImage13.Clear()
        AxImage12.Clear()
        FreeIfAllocated(AxImage16)
        FreeIfAllocated(AxImage17)
        FreeIfAllocated(AxImage13)
        FreeIfAllocated(AxImage12)
        Return result

    End Function
    Function DotQC(ByVal BlackDot As Boolean, ByVal ValueBinarized As Integer, ByVal ValueClose As Integer, ByVal ValueOpen As Integer, ByVal ValueMinArea As Double, ByVal ValueMaxArea As Double, ByVal ValueRoughness As Double, ByVal ValueCompactness As Double, ByVal Tolerance As Double, ByVal Diameter As Double) As Boolean

        ValueMaxArea = PI * (ValueMaxArea ^ 2) / (PixelSizeX ^ 2)
        ValueMinArea = PI * (ValueMinArea ^ 2) / (PixelSizeX ^ 2)
        QC(ValueBinarized, ValueClose, ValueOpen)
        ApplyBlobAnalysis_Setting(BlackDot)

        AxImage16.Save("C:\run-1.bmp")
        AxImage17.Save("C:\run-2.bmp")
        AxImage13.Save("C:\run-3.bmp")
        AxImage12.Save("C:\run-4.bmp")
        Dim result As Boolean = BlobAnalysisQC(BlackDot, ValueMinArea, ValueMaxArea, ValueRoughness, ValueCompactness, Tolerance, Diameter)
        AxImage16.Clear()
        AxImage17.Clear()
        AxImage13.Clear()
        AxImage12.Clear()
        FreeIfAllocated(AxImage16)
        FreeIfAllocated(AxImage17)
        FreeIfAllocated(AxImage13)
        FreeIfAllocated(AxImage12)
        Return result

    End Function
    Function DotQC(ByVal BlackDot As Boolean, ByVal ValueBinarized As Integer, ByVal ValueClose As Integer, ByVal ValueOpen As Integer, ByVal ValueMinArea As Double, ByVal ValueMaxArea As Double, ByVal ValueRoughness As Double, ByVal ValueCompactness As Double) As Boolean

        ValueMaxArea = PI * (ValueMaxArea ^ 2) / (PixelSizeX ^ 2)
        ValueMinArea = PI * (ValueMinArea ^ 2) / (PixelSizeX ^ 2)
        QC(ValueBinarized, ValueClose, ValueOpen)
        ApplyBlobAnalysis_Setting(BlackDot)

        AxImage16.Save("C:\IDS\SRC\DLL Export Device Vision\QC\test1.bmp")
        AxImage17.Save("C:\IDS\SRC\DLL Export Device Vision\QC\test2.bmp")
        AxImage13.Save("C:\IDS\SRC\DLL Export Device Vision\QC\test3.bmp")
        AxImage12.Save("C:\IDS\SRC\DLL Export Device Vision\QC\test4.bmp")
        Dim result As Boolean = BlobAnalysisQC(BlackDot, ValueMinArea, ValueMaxArea, ValueRoughness, ValueCompactness)
        AxImage16.Clear()
        AxImage17.Clear()
        AxImage13.Clear()
        AxImage12.Clear()
        FreeIfAllocated(AxImage16)
        FreeIfAllocated(AxImage17)
        FreeIfAllocated(AxImage13)
        FreeIfAllocated(AxImage12)
        Return result

    End Function
#End Region
#Region "RESET Variables"
    Sub Reset()
        DisplayCenterXPosition = 768 / 2
        DisplayCenterYPosition = 576 / 2
        MROIx = 100
        MROIy = 100
        ROIx = 300
        ROIy = 300
        RegionX = DisplayCenterXPosition
        RegionY = DisplayCenterYPosition
    End Sub
    Function ClearDisplay()
        AxDisplay1.OverlayImage.Clear()
    End Function
#End Region

    Function CheckFileName(ByVal Path) As Boolean
        'This function is used when saving a file to check there is not already a file with the same name so that you don't overwrite it.
        'It adds numbers to the filename e.g. file.gif becomes file1.gif becomes file2.gif and so on.
        'It keeps going until it returns a filename that does not exist.
        'You could just create a filename from the ID field but that means writing the record - and it still might exist!
        'N.B. Requires strSaveToPath variable to be available - and containing the path to save to
        Dim Counter
        Dim Flag
        Dim strTempFileName
        Dim FileExt As String
        Dim NewFullPath
        Dim objFSO As Object

        objFSO = CreateObject("Scripting.FileSystemObject")
        Counter = 0
        FileExt = "txt" 'GetExt(FileName)
        strTempFileName = "Reject" 'StripExt(FileName)
        'NewFullPath = "c:\iDS\SRC\DLL Export Device Vision\RejectPoint" & "\" & FileName & "." & FileExt
        NewFullPath = Path
        Flag = False

        'Do Until Flag = True
        If objFSO.FileExists(NewFullPath) = True Then
            'Flag = True
            Dim FileName_checked = Mid(NewFullPath, InStrRev(NewFullPath, "\") + 1)
            Return True
        Else
            'Counter = Counter + 1
            Return False
            'NewFullPath = "c:\iDS\SRC\DLL Export Device Vision\RejectPoint" & "\" & strTempFileName & Counter & "." & FileExt
        End If
        'Loop
    End Function ' check text file

    Function CheckExistFileName(ByVal FullPath) As Boolean
        Dim Counter
        Dim Flag
        Dim strTempFileName
        Dim FileExt As String
        Dim NewFullPath
        Dim objFSO As Object

        objFSO = CreateObject("Scripting.FileSystemObject")
        Counter = 0
        FileExt = "mmo" 'GetExt(FileName)
        strTempFileName = "" 'StripExt(FileName)
        NewFullPath = FullPath & "." & FileExt
        'NewFullPath = Path
        Flag = False

        'Do Until Flag = True
        If objFSO.FileExists(NewFullPath) = True Then
            Flag = True
            Dim FileName_checked = Mid(NewFullPath, InStrRev(NewFullPath, "\") + 1)
            Return True
        Else
            'Counter = Counter + 1
            NewFullPath = FullPath & "\" & strTempFileName & Counter & "." & FileExt
            Return False
        End If
        'Loop
    End Function
    Function CvrtDoubleToString(ByVal value As Double) As String
        If value >= 10 And value < 100 Then
            If (Convert.ToString((Fix(1000 * value)) / 1000)).Length = 5 Then
                Return ((Fix(1000 * value)) / 1000) & "0"
            ElseIf (Convert.ToString((Fix(1000 * value)) / 1000)).Length = 4 Then
                Return ((Fix(1000 * value)) / 1000) & "00"
            Else
                Return ((Fix(1000 * value)) / 1000)
            End If
        ElseIf value >= 100 And value < 1000 Then
            If (Convert.ToString((Fix(1000 * value)) / 1000)).Length = 6 Then
                Return ((Fix(1000 * value)) / 1000) & "0"
            ElseIf (Convert.ToString((Fix(1000 * value)) / 1000)).Length = 5 Then
                Return ((Fix(1000 * value)) / 1000) & "00"
            Else
                Return ((Fix(1000 * value)) / 1000)
            End If
        Else
            If (Convert.ToString((Fix(1000 * value)) / 1000)).Length = 4 Then
                Return ((Fix(1000 * value)) / 1000) & "0"
            ElseIf (Convert.ToString((Fix(1000 * value)) / 1000)).Length = 3 Then
                Return ((Fix(1000 * value)) / 1000) & "00"
            Else
                Return ((Fix(1000 * value)) / 1000)
            End If
        End If

    End Function
#Region "Programming"
    Sub ShowFiducialForm(ByVal Fi_No As Integer, ByVal brightness As Double)
        Try
            SetBrightness(brightness)
            FiducialMark_form.BrightnessValue.Value = brightness
            FiducialMark_form.BrightnessValue.Text = brightness
            FiducialMark_form.Show()
            FiducialMark_form.Location = New Point(0, 50)
            If FiducialMark_form.Visible = False Then
                FiducialMark_form.Visible = True
            End If
            FiducialMark_form.ResetFiducialStatus()
            FiducialMark_form.Fi_No(Fi_No)
        Catch ex As Exception
            ExceptionDisplay(ex)
        End Try
    End Sub
    Function SetFiducialFilename(ByVal pathname As String)
        FiducialMark_form.SetFiducialFilename(pathname)
        Folder = pathname
    End Function
    Sub ResetReject_Flag()
        RejectPoint_flag = False
    End Sub
    Function GetRMStatus() As Integer
        Return RejectPoint_form.GetRMStatus
    End Function
    Function GetRMParameters(ByRef param As DLL_Export_Device_Vision.RejectPoint.RMParam)
        RejectPoint_form.GetRMParameters(param)
    End Function
    Function SetRMReset()
        RejectPoint_form.SetRMReset()
    End Function
    Sub Form_QC(ByVal brightness As Double)
        SetBrightness(brightness)
        QC_form.ValueDotBrightness.Value = brightness
        QC_form.ValueDotBrightness.Text = brightness
        QC_form.Show()
        QC_form.DialogResult = DialogResult.OK
        QC_form.Location = New Point(0, 50)
        QC_form.SetQCReset()
        ClearDisplay()
        modelregionDrawing()
        If QC_form.Visible = False Then
            QC_form.Visible = True
        End If
        QC_form.ResetQCStatus()
    End Sub
    Sub Form_CE(ByVal brightness)
        SetBrightness(brightness)
        ChipEdgePoints_form.ValueBrightness.Value = brightness
        ChipEdgePoints_form.ValueBrightness.Text = brightness
        EnableChipEdgeDrawing()
        ChipEdgePoints_form.Show()
        ChipEdgePoints_form.Location = New Point(0, 50)
        If ChipEdgePoints_form.Visible = False Then
            ChipEdgePoints_form.Visible = True
        End If
        ChipEdgePoints_form.ResetChipEdgeStatus()
    End Sub
    Function GetChipEdgeParameters(ByRef param As DLL_Export_Device_Vision.ChipEdgePoints.ChipEdgeParam)
        ChipEdgePoints_form.GetChipEdgeParameters(param)
    End Function

#End Region
#Region "Edit"
    Function Form_FI_Edit(ByVal Fi_No As Integer, ByVal Filename As String, ByVal brightness As Integer) As Boolean
        Try
            SetBrightness(brightness)
            FiducialMark_form.BrightnessValue.Value = brightness
            FiducialMark_form.BrightnessValue.Text = brightness
            FiducialMark_form.Show()
            FiducialMark_form.Location = New Point(0, 50)
            If FiducialMark_form.Visible = False Then
                FiducialMark_form.Visible = True
            End If
            FiducialMark_form.ResetFiducialStatus()
            FiducialMark_form.Fi_No(Fi_No)
            Dim IDSError As Boolean = LoadFiducial(Filename)
            If IDSError = True Then
                Return True
            Else
                MsgBox("Load File Failed!")
                Return False
            End If
        Catch ex As Exception
            ExceptionDisplay(ex)
        End Try
    End Function
    Function Form_CE_Edit(ByVal vParam As ChipEdgePoints.ChipEdgeParam) As Boolean
        EnableChipEdgeDrawing()
        ChipEdgePoints_form.Show()
        ChipEdgePoints_form.Location = New Point(0, 50)
        If ChipEdgePoints_form.Visible = False Then
            ChipEdgePoints_form.Visible = True
        End If
        ChipEdgePoints_form.ResetChipEdgeStatus()

        ChipEdgePoints_form.TextBox_SizeX.Text = vParam._SizeX
        ChipEdgePoints_form.TextBox_SizeY.Text = vParam._SizeY
        ChipEdgePoints_form.TextBox_PosX.Text = vParam._PosX
        ChipEdgePoints_form.TextBox_PosY.Text = vParam._PosY
        ChipEdgePoints_form.ValueSize.Value() = vParam._Size
        ChipEdgePoints_form.ValuePos.Value() = vParam._Pos
        ChipEdgePoints_form.ValueRot.Value() = vParam._Rot
        ChipEdgePoints_form.ValueThreshold.Value = vParam._Threshold
        ChipEdgePoints_form.ValueROI.Value = vParam._ROI
        ChipEdgePoints_form.ValueBrightness.Value = vParam._Brightness
        ChipEdgePoints_form.ValueBrightness.Text = vParam._Brightness

        Dim Cw_CCw As Boolean = vParam._Cw_CCw
        Dim Vertical As Boolean = vParam._Vertical
        Dim Inside_out As Boolean = vParam._Inside_out
        If Inside_out = True Then
            ChipEdgePoints_form.RadioButton_Inside_out.Checked = True
        Else
            ChipEdgePoints_form.RadioButton_Outside_In.Checked = True
        End If

        If Cw_CCw = True Then
            ChipEdgePoints_form.RadioButton_CW.Checked = True
        Else
            ChipEdgePoints_form.RadioButton_CCW.Checked = True
        End If

        If Vertical = True Then
            ChipEdgePoints_form.RadioButton_Vertical.Checked = True
        Else
            ChipEdgePoints_form.RadioButton_Horizontal.Checked = True
        End If

        MainEdge = vParam._MainEdge
        PointX1 = vParam._PointX1
        PointY1 = vParam._PointY1
        PointX2 = vParam._PointX2
        PointY2 = vParam._PointY2
        PointX3 = vParam._PointX3
        PointY3 = vParam._PointY3
        PointX4 = vParam._PointX4
        PointY4 = vParam._PointY4
        PointX5 = vParam._PointX5
        PointY5 = vParam._PointY5
        Dim DispenseModel As Integer = vParam._DispenseModel
        If DispenseModel = 2 Then
            ChipEdgePoints_form.RadioButton_TwoEdges.Checked = True
        ElseIf DispenseModel = 1 Then
            ChipEdgePoints_form.RadioButton_OneEdge.Checked = True
        Else
            ChipEdgePoints_form.RadioButton_Dot.Checked = True
        End If

        ChipEdgePoints_form.TextBox_EdgeClearance.Text = vParam._EdgeClearance
        ChipEdgePoints_form.CheckBox_ChipRec_Enable.Checked = vParam._CheckBox_ChipRec_Enable
        ChipEdgePoints_form.Button_Test.Enabled = True
        ChipEdgePoints_form.Button_Ok.Enabled = True
        ChipEdgePoints_form.ChipEdgePointsCoordinate(PointX1, PointY1, PointX2, PointY2, PointX3, PointY3, PointX4, PointY4, PointX5, PointY5)
        If ChipEdgePoints_form.CheckBox_ChipRec_Enable.Checked = False Then
            'Do nth
        Else
            ChipEdgePoints_form.GetVariables(Inside_out, Vertical, Cw_CCw, DispenseModel)
        End If
    End Function
    Function Form_QC_Edit(ByVal vParam As QC.QCParam) As Boolean
        Try
            QC_form.Show()
            QC_form.DialogResult = DialogResult.OK
            QC_form.Location = New Point(0, 50)
            QC_form.SetQCReset()
            ClearDisplay()
            If QC_form.Visible = False Then
                QC_form.Visible = True
            End If
            QC_form.ResetQCStatus()
            If vParam._BlackDot = True Then
                QC_form.RadioButton_BlackDot.Checked = vParam._BlackDot
            Else
                QC_form.RadioButton_WhiteDot.Checked = True
            End If
            _LightBox.SetBrightness(vParam._Brightness)
            QC_form.ValueBinarized.Value = vParam._Binarized
            QC_form.ValueMaxArea.Value = vParam._MaxArea
            QC_form.ValueMinArea.Value = vParam._MinArea
            QC_form.ValueClose.Value = vParam._Close
            QC_form.ValueOpen.Value = vParam._Open
            QC_form.ValueRoughness.Value = vParam._Roughness
            QC_form.ValueCompactness.Value = vParam._Compactness
            QC_form.ValueDotBrightness.Value = vParam._Brightness
            QC_form.ValueDesiredDiameter.Text = vParam._Diameter
            QC_form.ValueTolerance.Value = vParam._Tolerance
            DisplayCenterXPosition = vParam._MRegionX
            DisplayCenterYPosition = vParam._MRegionY
            MROIx = vParam._MROIx
            MROIy = vParam._MROIy
            modelregionDrawing()
        Catch ex As Exception
        End Try

    End Function
    Function Form_RM_Edit(ByVal vParam As RejectPoint.RMParam) As Boolean
        RejectPoint_flag = True
        RejectPoint_form.ResetRMStatus()
        RejectPoints()
        SetRMReset()
        RejectPoint_form.Button_Next.Enabled = True
        If RejectPoint_form.Visible = False Then
            RejectPoint_form.Visible = True
        End If

        Dim BlackWithoutRM, WhiteWithoutRM, BlackWithRM, WhiteWithRM As Boolean
        RejectPoint_form.ValueBrightness.Value = vParam._Brightness
        RejectPoint_form.ValueBinarized.Value = vParam._Binarized
        RejectPoint_form.RadioButton_WoRM.Checked = vParam._WoRM
        RejectPoint_form.RadioButton_WRM.Checked = vParam._WRM
        RejectPoint_form.ValueAcceptRatio.Value = vParam._AcceptRatio
        BlackWithoutRM = vParam._BlackWithoutRM
        WhiteWithoutRM = vParam._WhiteWithoutRM
        BlackWithRM = vParam._BlackWithRM
        WhiteWithRM = vParam._WhiteWithRM
        DisplayCenterXPosition = vParam._MRegionX
        DisplayCenterYPosition = vParam._MRegionY
        MROIx = vParam._MROIx
        MROIy = vParam._MROIy
        RejectPoint_form.BlackWhite(BlackWithoutRM, WhiteWithoutRM, BlackWithRM, WhiteWithRM)
        RejectPoint_form.Button_Test.Enabled = True
        RejectPoint_form.Button_Ok.Enabled = True
        RejectPoint_form.PictureBox1.Image = Image.FromFile(Folder + "RM.bmp")

        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '   Xue Wen                                                                                     '
        '   Set folder name for "RejectPoint.vb".                                                       '
        '   If not, "RejectPoint.vb" will take its default value "c:\". So, it can cause some errors.   '
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        SetRMFilename(Folder)
    End Function
#End Region
#Region "Production_Input"
    Function IDSV_CE(ByVal vParam As ChipEdgePoints.ChipEdgeParam) As Boolean
        Dim Inside_out As Boolean = vParam._Inside_out
        Dim Cw_CCw As Boolean = vParam._Cw_CCw
        Dim Vertical As Boolean = vParam._Vertical
        Dim DispenseModel As Integer = vParam._DispenseModel
        ChipEdgePoints_form.ValueBrightness.Value = vParam._Brightness
        ChipEdgePoints_form.TextBox_SizeX.Text = vParam._SizeX
        ChipEdgePoints_form.TextBox_SizeY.Text = vParam._SizeY
        ChipEdgePoints_form.TextBox_PosX.Text = vParam._PosX
        ChipEdgePoints_form.TextBox_PosY.Text = vParam._PosY
        ChipEdgePoints_form.ValueSize.Value() = vParam._Size
        ChipEdgePoints_form.ValuePos.Value() = vParam._Pos
        ChipEdgePoints_form.ValueRot.Value() = vParam._Rot
        ChipEdgePoints_form.ValueThreshold.Value = vParam._Threshold
        ChipEdgePoints_form.ValueROI.Value = vParam._ROI
        MainEdge = vParam._MainEdge
        PointX1 = vParam._PointX1
        PointY1 = vParam._PointY1
        PointX2 = vParam._PointX2
        PointY2 = vParam._PointY2
        PointX3 = vParam._PointX3
        PointY3 = vParam._PointY3
        PointX4 = vParam._PointX4
        PointY4 = vParam._PointY4
        PointX5 = vParam._PointX5
        PointY5 = vParam._PointY5
        ChipEdgePoints_form.TextBox_EdgeClearance.Text = vParam._EdgeClearance
        ChipEdgePoints_form.CheckBox_ChipRec_Enable.Checked = vParam._CheckBox_ChipRec_Enable
        If ChipEdgePoints_form.CheckBox_ChipRec_Enable.Checked = False Then
            'Do nth
        Else
            ChipEdgePoints_form.GetVariables(Inside_out, Vertical, Cw_CCw, DispenseModel)
            WaitForImageToStabilize()
            Dim result As Boolean = ChipEdgePoints_form.ChipEdgePointExe()
            'ClearDisplay()
            'DisplayIndicator()
            Return result
        End If
    End Function
    Function IDSV_NC(ByVal BlackDot As Boolean, ByVal Binarized As Integer, ByVal MaxArea As Double, ByVal MinArea As Double, ByVal Close As Integer, ByVal Open As Integer, ByVal Roughness As Double, ByVal Compactness As Double, ByVal _DisplayCenterXPosition As Decimal, ByVal _DisplayCenterYPosition As Decimal, ByVal _MRoiX As Decimal, ByVal _MRoiY As Decimal, ByRef NC_OffX As Double, ByRef NC_OffY As Double) As Boolean
        DisplayCenterXPosition = _DisplayCenterXPosition
        DisplayCenterYPosition = _DisplayCenterYPosition
        MROIx = _MRoiX
        MROIy = _MRoiY
        NC_OffX = 0
        NC_OffY = 0
        modelregionDrawing()
        WaitForImageToStabilize()
        Dim result As Boolean = DotNC(BlackDot, Binarized, Close, Open, MinArea, MaxArea, Roughness, Compactness, NC_OffX, NC_OffY)
        FreeIfAllocated(AxImage16)
        Return result
    End Function
    Function IDSV_QC(ByVal vParam As QC.QCParam) As Boolean
        QC_form.ValueDotBrightness.Value = vParam._Brightness
        QC_form.RadioButton_BlackDot.Checked = vParam._BlackDot
        QC_form.ValueBinarized.Value = vParam._Binarized
        QC_form.ValueMaxArea.Value = vParam._MaxArea
        QC_form.ValueMinArea.Value = vParam._MinArea
        QC_form.ValueClose.Value = vParam._Close
        QC_form.ValueOpen.Value = vParam._Open
        QC_form.ValueRoughness.Value = vParam._Roughness
        QC_form.ValueCompactness.Value = vParam._Compactness
        QC_form.ValueDesiredDiameter.Text = vParam._Diameter
        QC_form.ValueTolerance.Value = vParam._Tolerance
        DisplayCenterXPosition = vParam._MRegionX
        DisplayCenterYPosition = vParam._MRegionY
        MROIx = vParam._MROIx
        MROIy = vParam._MROIy
        modelregionDrawing()
        WaitForImageToStabilize()
        Dim result As Boolean = DotQC(vParam._BlackDot, vParam._Binarized, vParam._Close, vParam._Open, vParam._MinArea, vParam._MaxArea, vParam._Roughness, vParam._Compactness, vParam._Tolerance, vParam._Diameter)
        FreeIfAllocated(AxImage16)
        Return result
    End Function
    Function IDSV_RM(ByVal vParam As RejectPoint.RMParam) As Boolean
        Dim BlackWithoutRM, WhiteWithoutRM, BlackWithRM, WhiteWithRM As Boolean
        RejectPoint_form.ValueBrightness.Value = vParam._Brightness
        RejectPoint_form.ValueBinarized.Value = vParam._Binarized
        RejectPoint_form.RadioButton_WoRM.Checked = vParam._WoRM
        RejectPoint_form.RadioButton_WRM.Checked = vParam._WRM
        RejectPoint_form.ValueAcceptRatio.Value = vParam._AcceptRatio
        BlackWithoutRM = vParam._BlackWithoutRM
        WhiteWithoutRM = vParam._WhiteWithoutRM
        BlackWithRM = vParam._BlackWithRM
        WhiteWithRM = vParam._WhiteWithRM
        DisplayCenterXPosition = vParam._MRegionX
        DisplayCenterYPosition = vParam._MRegionY
        MROIx = vParam._MROIx
        MROIy = vParam._MROIy
        RejectPoint_form.BlackWhite(BlackWithoutRM, WhiteWithoutRM, BlackWithRM, WhiteWithRM)
        WaitForImageToStabilize()
        Dim result As Boolean = RejectPoint_form.ProductionRUN()
        Return result
    End Function
    Function IDSV_FI(ByVal FileName As String, ByRef FI_OffX As Double, ByRef FI_OffY As Double) As Boolean
        Dim IDSError As Boolean = LoadFiducial(FileName)
        If IDSError = True Then
            WaitForImageToStabilize()
            Dim Result As Boolean = Test_Fiducial(21, FI_OffX, FI_OffY)
            'CameraResume()
            'ClearDisplay()
            Return Result
        Else
        End If
    End Function
#End Region
#Region "Production_Output"
    Function IDSV_CE(ByRef PointX1 As Double, ByRef PointY1 As Double, ByRef PointX2 As Double, ByRef PointY2 As Double, ByRef PointX3 As Double, ByRef PointY3 As Double, ByRef PointX4 As Double, ByRef PointY4 As Double)
        PointX1 = Output(0)
        PointX2 = Output(2)
        PointX3 = Output(4)
        PointX4 = Output(6)
        PointY1 = Output(1)
        PointY2 = Output(3)
        PointY3 = Output(5)
        PointY4 = Output(7)
    End Function
    Function IDSV_FI(ByRef DelX As Double, ByRef DelY As Double)
        DelX = FiducialMark_form.TextBox_RPosX.Text
        DelY = FiducialMark_form.TextBox_RPosY.Text
    End Function
#End Region

#Region "Camera Functionality"

    Private CurrentMode As String = "Teach Camera"
    Public Sub SwitchCamera(ByVal command As String)
        If CurrentMode = command Then
            Exit Sub
        Else
            CurrentMode = command
        End If
        If command = "View Camera" Then
            AxECamera1.SetParamNm("ChannelState", "IDLE")
            AxECamera1.SetParamNm("Connector", "VID3")
            AxECamera1.SetParamNm("GrabCount", -1)
            AxECamera1.SetParamNm("ChannelState", "ACTIVE")
            SetBrightness(0) 'change to view
        ElseIf command = "Teach Camera" Then
            AxECamera1.SetParamNm("ChannelState", "IDLE")
            AxECamera1.SetParamNm("Connector", "VID2")
            AxECamera1.SetParamNm("GrabCount", -1)
            AxECamera1.SetParamNm("ChannelState", "ACTIVE")
        End If
    End Sub

    Sub InitializeCameraSettings()
        Dim PicoloIndex, a As Integer
        Dim Camera1 As String
        Dim ImageSizeX As Long
        Dim ImageSizeY As Long
        AxECamera1.Mpf = "Channel"
        AxECamera1.SetParamNm("DriverIndex", 0)
        AxECamera1.SetParamNm("Connector", "VID2")
        AxECamera1.SetParamNm("CamFile", "PAL")
        AxECamera1.SetParamNm("ColorFormat", "Y8")
        AxECamera1.SetParamNm("AcquisitionMode", "VIDEO")
        AxECamera1.SetParamNm("TrigMode", "IMMEDIATE")
        AxECamera1.SetParamNm("NextTrigMode", "REPEAT")
        AxECamera1.SetParamNm("SeqLength_fr", -1)
        AxECamera1.SetParamNm("ImageFlipY", "ON") 'Can use
        ImageSizeX = AxECamera1.GetParamNm("ImageSizeX")
        ImageSizeY = AxECamera1.GetParamNm("ImageSizeY")
        AxEBW8Image1.SetSize(ImageSizeX, ImageSizeY)
        AxECamera1.Cluster = AxEBW8Image1.Obj
        AxECamera1.SetParamNm("SignalEnable:2", "ON")
        AxECamera1.SetParamNm("SignalEnable:7", "ON")
        AxECamera1.SetParamNm("GrabCount", -1)
        CameraResume()
    End Sub

    Public Sub CameraResume()
        AxECamera1.SetParamNm("ChannelState", "ACTIVE")
    End Sub

    Public Sub CameraIdle()
        AxECamera1.SetParamNm("ChannelState", "IDLE")
    End Sub

    Private Shared Marshal As System.Runtime.InteropServices.Marshal

    Private Sub AxECamera1_Signal(ByVal sender As Object, ByVal e As AxMULTICAMLib._IECameraEvents_SignalEvent) Handles AxECamera1.Signal
        If e.signalType = MULTICAMLib.enumMcEventType.eMcSigAcquisitionFailure Then
            TextBox1.Text = "Acquisition Failure"
        Else
            If e.signalType = MULTICAMLib.enumMcEventType.eMcSigSurfaceFilled Then

                Try
                    Marshal.Copy(ImagePointer_PICOLO, intermidiate, 0, DisplayWidth * DisplayHeight)
                    Marshal.Copy(ImagePointer_MIL, intermidiate1, 0, DisplayWidth * DisplayHeight)
                    Array.Copy(intermidiate, intermidiate1, DisplayWidth * DisplayHeight)
                    Marshal.Copy(intermidiate1, 0, ImagePointer_MIL, DisplayWidth * DisplayHeight)
                    AxImageProcessing7.Flip(ImageProcessing.ImpFlipConstants.impHorizontal)  'slow and flicker
                    WaitRefresh = True
                Catch ex As Exception
                    ExceptionDisplay(ex)
                End Try
            End If
        End If
    End Sub

    Public Function IsThereCameraError() As Boolean
        If TextBox1.Text = "Acquisition Failure" Then
            Return True
        End If
        Return False
    End Function

#End Region


    Dim ClickNo As Integer
    Dim x As Double
    Dim y As Double
    Dim PosX As Double
    Dim PosY As Double
    Dim ChipX1, ChipY1, ChipX2, ChipY2, ChipX3, ChipY3 As Double
    Dim SizeX As Double
    Dim SizeY As Double

    Dim ClickNoROI As Integer
    Dim xROI As Double
    Dim yROI As Double
    Dim PosXROI As Double
    Dim PosYROI As Double
    Dim ROIX1, ROIY1, ROIX2, ROIY2, ROIX3, ROIY3 As Double
    Dim SizeXROI As Double
    Dim SizeYROI As Double

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Form_QC(ValueBrightness.Value)
    End Sub

    'this function is being called from other subs
    Public Sub WaitForImageToStabilize()
        CameraIdle()
        Sleep(250)
        CameraResume()
        TraceDoEvents()
    End Sub

    Private Sub Button_ChipEdgePoints_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_ChipEdgePoints.Click

        EnableChipEdgeDrawing()
        ChipEdgePoints_form.Show()
        ChipEdgePoints_form.Location = New Point(0, 50)
    End Sub

    Private Sub AxDisplay1_KeyDownEvent(ByVal sender As Object, ByVal e As AxMatrox.ActiveMIL.IDisplayEvents_KeyDownEvent) Handles AxDisplay1.KeyDownEvent
        Console.WriteLine("KeyDown in vision")
    End Sub
End Class
