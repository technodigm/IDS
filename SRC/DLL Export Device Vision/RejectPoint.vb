Imports System.IO

Public Class RejectPoint
    Inherits System.Windows.Forms.Form

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '   Xue Wen                                                                                     '
    '   Another method, Not to fire "ValueChanged" event while constructing Form's components.      '
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Private Initializing As Boolean = True

#Region " Windows Form Designer generated code "
    Public Sub New()

        MyBase.New()
        'This call is required by the Windows Form Designer.
        InitializeComponent()

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
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents SelectRejectMode As System.Windows.Forms.TabPage
    Friend WithEvents CreateRejectMode As System.Windows.Forms.TabPage
    Friend WithEvents Button_Test As System.Windows.Forms.Button
    Friend WithEvents Button_Ok As System.Windows.Forms.Button
    Friend WithEvents Button_Cancel As System.Windows.Forms.Button
    Friend WithEvents RadioButton_WithoutRejectMark As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton_RejectMark As System.Windows.Forms.RadioButton
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents ValueRatio As System.Windows.Forms.NumericUpDown
    Friend WithEvents Timer1 As System.Windows.Forms.Timer
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents PictureBox2 As System.Windows.Forms.PictureBox
    Friend WithEvents GroupBox4 As System.Windows.Forms.GroupBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents ValueBrightness As System.Windows.Forms.NumericUpDown
    Friend WithEvents ValueBinarized As System.Windows.Forms.NumericUpDown
    Friend WithEvents ValueAcceptRatio As System.Windows.Forms.NumericUpDown
    Friend WithEvents ValueBrightness1 As System.Windows.Forms.NumericUpDown
    Friend WithEvents ValueBinarizedThreshold As System.Windows.Forms.NumericUpDown
    Friend WithEvents RadioButton_WoRM As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton_WRM As System.Windows.Forms.RadioButton
    Friend WithEvents Button_Load As System.Windows.Forms.Button
    Friend WithEvents Button_Next As System.Windows.Forms.Button
    Friend WithEvents Button_RESET As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Me.Button_Load = New System.Windows.Forms.Button
        Me.Button_Test = New System.Windows.Forms.Button
        Me.Button_Ok = New System.Windows.Forms.Button
        Me.Button_Cancel = New System.Windows.Forms.Button
        Me.TabControl1 = New System.Windows.Forms.TabControl
        Me.CreateRejectMode = New System.Windows.Forms.TabPage
        Me.GroupBox3 = New System.Windows.Forms.GroupBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.ValueRatio = New System.Windows.Forms.NumericUpDown
        Me.Button_Next = New System.Windows.Forms.Button
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.RadioButton_RejectMark = New System.Windows.Forms.RadioButton
        Me.RadioButton_WithoutRejectMark = New System.Windows.Forms.RadioButton
        Me.ValueBrightness1 = New System.Windows.Forms.NumericUpDown
        Me.Label3 = New System.Windows.Forms.Label
        Me.ValueBinarizedThreshold = New System.Windows.Forms.NumericUpDown
        Me.Label4 = New System.Windows.Forms.Label
        Me.PictureBox1 = New System.Windows.Forms.PictureBox
        Me.SelectRejectMode = New System.Windows.Forms.TabPage
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.ValueAcceptRatio = New System.Windows.Forms.NumericUpDown
        Me.PictureBox2 = New System.Windows.Forms.PictureBox
        Me.GroupBox4 = New System.Windows.Forms.GroupBox
        Me.RadioButton_WRM = New System.Windows.Forms.RadioButton
        Me.RadioButton_WoRM = New System.Windows.Forms.RadioButton
        Me.ValueBrightness = New System.Windows.Forms.NumericUpDown
        Me.Label2 = New System.Windows.Forms.Label
        Me.ValueBinarized = New System.Windows.Forms.NumericUpDown
        Me.Label6 = New System.Windows.Forms.Label
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog
        Me.TextBox1 = New System.Windows.Forms.TextBox
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.Button_RESET = New System.Windows.Forms.Button
        Me.TabControl1.SuspendLayout()
        Me.CreateRejectMode.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        CType(Me.ValueRatio, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox2.SuspendLayout()
        CType(Me.ValueBrightness1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.ValueBinarizedThreshold, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SelectRejectMode.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        CType(Me.ValueAcceptRatio, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox4.SuspendLayout()
        CType(Me.ValueBrightness, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.ValueBinarized, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Button_Load
        '
        Me.Button_Load.Enabled = False
        Me.Button_Load.Location = New System.Drawing.Point(592, 8)
        Me.Button_Load.Name = "Button_Load"
        Me.Button_Load.Size = New System.Drawing.Size(232, 248)
        Me.Button_Load.TabIndex = 0
        Me.Button_Load.Text = "Load Model"
        '
        'Button_Test
        '
        Me.Button_Test.Location = New System.Drawing.Point(1096, 56)
        Me.Button_Test.Name = "Button_Test"
        Me.Button_Test.Size = New System.Drawing.Size(168, 240)
        Me.Button_Test.TabIndex = 1
        Me.Button_Test.Text = "Test"
        '
        'Button_Ok
        '
        Me.Button_Ok.Enabled = False
        Me.Button_Ok.Location = New System.Drawing.Point(1184, 304)
        Me.Button_Ok.Name = "Button_Ok"
        Me.Button_Ok.Size = New System.Drawing.Size(80, 25)
        Me.Button_Ok.TabIndex = 5
        Me.Button_Ok.Text = "Ok"
        '
        'Button_Cancel
        '
        Me.Button_Cancel.Location = New System.Drawing.Point(1096, 304)
        Me.Button_Cancel.Name = "Button_Cancel"
        Me.Button_Cancel.Size = New System.Drawing.Size(80, 25)
        Me.Button_Cancel.TabIndex = 6
        Me.Button_Cancel.Text = "Cancel"
        '
        'TabControl1
        '
        Me.TabControl1.Controls.Add(Me.CreateRejectMode)
        Me.TabControl1.Controls.Add(Me.SelectRejectMode)
        Me.TabControl1.Location = New System.Drawing.Point(8, 8)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(1080, 288)
        Me.TabControl1.TabIndex = 7
        '
        'CreateRejectMode
        '
        Me.CreateRejectMode.Controls.Add(Me.GroupBox3)
        Me.CreateRejectMode.Controls.Add(Me.Button_Next)
        Me.CreateRejectMode.Controls.Add(Me.GroupBox2)
        Me.CreateRejectMode.Controls.Add(Me.PictureBox1)
        Me.CreateRejectMode.Location = New System.Drawing.Point(4, 22)
        Me.CreateRejectMode.Name = "CreateRejectMode"
        Me.CreateRejectMode.Size = New System.Drawing.Size(1072, 262)
        Me.CreateRejectMode.TabIndex = 1
        Me.CreateRejectMode.Text = "Create Reject Model"
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.Label5)
        Me.GroupBox3.Controls.Add(Me.ValueRatio)
        Me.GroupBox3.Location = New System.Drawing.Point(464, 8)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(120, 248)
        Me.GroupBox3.TabIndex = 9
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Acceptance Ratio"
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(88, 44)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(18, 12)
        Me.Label5.TabIndex = 3
        Me.Label5.Text = "%"
        '
        'ValueRatio
        '
        Me.ValueRatio.Location = New System.Drawing.Point(16, 40)
        Me.ValueRatio.Name = "ValueRatio"
        Me.ValueRatio.Size = New System.Drawing.Size(64, 20)
        Me.ValueRatio.TabIndex = 0
        Me.ValueRatio.Value = New Decimal(New Integer() {5, 0, 0, 0})
        '
        'Button_Next
        '
        Me.Button_Next.Location = New System.Drawing.Point(592, 8)
        Me.Button_Next.Name = "Button_Next"
        Me.Button_Next.Size = New System.Drawing.Size(232, 248)
        Me.Button_Next.TabIndex = 7
        Me.Button_Next.Text = "Next"
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.RadioButton_RejectMark)
        Me.GroupBox2.Controls.Add(Me.RadioButton_WithoutRejectMark)
        Me.GroupBox2.Controls.Add(Me.ValueBrightness1)
        Me.GroupBox2.Controls.Add(Me.Label3)
        Me.GroupBox2.Controls.Add(Me.ValueBinarizedThreshold)
        Me.GroupBox2.Controls.Add(Me.Label4)
        Me.GroupBox2.Location = New System.Drawing.Point(8, 8)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(448, 248)
        Me.GroupBox2.TabIndex = 6
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Settings"
        '
        'RadioButton_RejectMark
        '
        Me.RadioButton_RejectMark.Location = New System.Drawing.Point(192, 56)
        Me.RadioButton_RejectMark.Name = "RadioButton_RejectMark"
        Me.RadioButton_RejectMark.Size = New System.Drawing.Size(120, 24)
        Me.RadioButton_RejectMark.TabIndex = 6
        Me.RadioButton_RejectMark.Text = "With RejectMark"
        '
        'RadioButton_WithoutRejectMark
        '
        Me.RadioButton_WithoutRejectMark.Checked = True
        Me.RadioButton_WithoutRejectMark.Location = New System.Drawing.Point(192, 24)
        Me.RadioButton_WithoutRejectMark.Name = "RadioButton_WithoutRejectMark"
        Me.RadioButton_WithoutRejectMark.Size = New System.Drawing.Size(128, 24)
        Me.RadioButton_WithoutRejectMark.TabIndex = 5
        Me.RadioButton_WithoutRejectMark.TabStop = True
        Me.RadioButton_WithoutRejectMark.Text = "Without RejectMark"
        '
        'ValueBrightness1
        '
        Me.ValueBrightness1.Location = New System.Drawing.Point(120, 24)
        Me.ValueBrightness1.Maximum = New Decimal(New Integer() {255, 0, 0, 0})
        Me.ValueBrightness1.Name = "ValueBrightness1"
        Me.ValueBrightness1.Size = New System.Drawing.Size(56, 20)
        Me.ValueBrightness1.TabIndex = 1
        Me.ValueBrightness1.Value = New Decimal(New Integer() {128, 0, 0, 0})
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(8, 24)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(80, 24)
        Me.Label3.TabIndex = 2
        Me.Label3.Text = "Brightness"
        '
        'ValueBinarizedThreshold
        '
        Me.ValueBinarizedThreshold.Location = New System.Drawing.Point(120, 56)
        Me.ValueBinarizedThreshold.Maximum = New Decimal(New Integer() {255, 0, 0, 0})
        Me.ValueBinarizedThreshold.Name = "ValueBinarizedThreshold"
        Me.ValueBinarizedThreshold.Size = New System.Drawing.Size(56, 20)
        Me.ValueBinarizedThreshold.TabIndex = 3
        Me.ValueBinarizedThreshold.Value = New Decimal(New Integer() {128, 0, 0, 0})
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(8, 56)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(104, 24)
        Me.Label4.TabIndex = 4
        Me.Label4.Text = "BinarizedThreshold"
        '
        'PictureBox1
        '
        Me.PictureBox1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.PictureBox1.Location = New System.Drawing.Point(832, 8)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(232, 248)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage
        Me.PictureBox1.TabIndex = 8
        Me.PictureBox1.TabStop = False
        '
        'SelectRejectMode
        '
        Me.SelectRejectMode.Controls.Add(Me.GroupBox1)
        Me.SelectRejectMode.Controls.Add(Me.PictureBox2)
        Me.SelectRejectMode.Controls.Add(Me.GroupBox4)
        Me.SelectRejectMode.Controls.Add(Me.Button_Load)
        Me.SelectRejectMode.Location = New System.Drawing.Point(4, 22)
        Me.SelectRejectMode.Name = "SelectRejectMode"
        Me.SelectRejectMode.Size = New System.Drawing.Size(1072, 262)
        Me.SelectRejectMode.TabIndex = 0
        Me.SelectRejectMode.Text = "Select Reject Model"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Controls.Add(Me.ValueAcceptRatio)
        Me.GroupBox1.Location = New System.Drawing.Point(464, 8)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(120, 248)
        Me.GroupBox1.TabIndex = 13
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Acceptance Ratio"
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(88, 44)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(18, 12)
        Me.Label1.TabIndex = 3
        Me.Label1.Text = "%"
        '
        'ValueAcceptRatio
        '
        Me.ValueAcceptRatio.Enabled = False
        Me.ValueAcceptRatio.Location = New System.Drawing.Point(16, 40)
        Me.ValueAcceptRatio.Name = "ValueAcceptRatio"
        Me.ValueAcceptRatio.Size = New System.Drawing.Size(64, 20)
        Me.ValueAcceptRatio.TabIndex = 0
        Me.ValueAcceptRatio.Value = New Decimal(New Integer() {5, 0, 0, 0})
        '
        'PictureBox2
        '
        Me.PictureBox2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.PictureBox2.Location = New System.Drawing.Point(832, 8)
        Me.PictureBox2.Name = "PictureBox2"
        Me.PictureBox2.Size = New System.Drawing.Size(232, 248)
        Me.PictureBox2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage
        Me.PictureBox2.TabIndex = 12
        Me.PictureBox2.TabStop = False
        '
        'GroupBox4
        '
        Me.GroupBox4.Controls.Add(Me.RadioButton_WRM)
        Me.GroupBox4.Controls.Add(Me.RadioButton_WoRM)
        Me.GroupBox4.Controls.Add(Me.ValueBrightness)
        Me.GroupBox4.Controls.Add(Me.Label2)
        Me.GroupBox4.Controls.Add(Me.ValueBinarized)
        Me.GroupBox4.Controls.Add(Me.Label6)
        Me.GroupBox4.Location = New System.Drawing.Point(8, 8)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Size = New System.Drawing.Size(448, 248)
        Me.GroupBox4.TabIndex = 10
        Me.GroupBox4.TabStop = False
        Me.GroupBox4.Text = "Settings"
        '
        'RadioButton_WRM
        '
        Me.RadioButton_WRM.Enabled = False
        Me.RadioButton_WRM.Location = New System.Drawing.Point(192, 56)
        Me.RadioButton_WRM.Name = "RadioButton_WRM"
        Me.RadioButton_WRM.Size = New System.Drawing.Size(120, 24)
        Me.RadioButton_WRM.TabIndex = 6
        Me.RadioButton_WRM.Text = "With RejectMark"
        '
        'RadioButton_WoRM
        '
        Me.RadioButton_WoRM.Checked = True
        Me.RadioButton_WoRM.Enabled = False
        Me.RadioButton_WoRM.Location = New System.Drawing.Point(192, 24)
        Me.RadioButton_WoRM.Name = "RadioButton_WoRM"
        Me.RadioButton_WoRM.Size = New System.Drawing.Size(128, 24)
        Me.RadioButton_WoRM.TabIndex = 5
        Me.RadioButton_WoRM.TabStop = True
        Me.RadioButton_WoRM.Text = "Without RejectMark"
        '
        'ValueBrightness
        '
        Me.ValueBrightness.Enabled = False
        Me.ValueBrightness.Location = New System.Drawing.Point(120, 24)
        Me.ValueBrightness.Maximum = New Decimal(New Integer() {255, 0, 0, 0})
        Me.ValueBrightness.Name = "ValueBrightness"
        Me.ValueBrightness.Size = New System.Drawing.Size(56, 20)
        Me.ValueBrightness.TabIndex = 1
        Me.ValueBrightness.Value = New Decimal(New Integer() {128, 0, 0, 0})
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(8, 24)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(80, 24)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "Brightness"
        '
        'ValueBinarized
        '
        Me.ValueBinarized.Enabled = False
        Me.ValueBinarized.Location = New System.Drawing.Point(120, 56)
        Me.ValueBinarized.Maximum = New Decimal(New Integer() {255, 0, 0, 0})
        Me.ValueBinarized.Name = "ValueBinarized"
        Me.ValueBinarized.Size = New System.Drawing.Size(56, 20)
        Me.ValueBinarized.TabIndex = 3
        Me.ValueBinarized.Value = New Decimal(New Integer() {128, 0, 0, 0})
        '
        'Label6
        '
        Me.Label6.Location = New System.Drawing.Point(8, 56)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(104, 24)
        Me.Label6.TabIndex = 4
        Me.Label6.Text = "BinarizedThreshold"
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(8, 304)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(1072, 20)
        Me.TextBox1.TabIndex = 8
        Me.TextBox1.Text = "Load a model and Test it"
        '
        'Button_RESET
        '
        Me.Button_RESET.Location = New System.Drawing.Point(1096, 32)
        Me.Button_RESET.Name = "Button_RESET"
        Me.Button_RESET.Size = New System.Drawing.Size(168, 23)
        Me.Button_RESET.TabIndex = 15
        Me.Button_RESET.Text = "Reset"
        '
        'RejectPoint
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(1232, 336)
        Me.Controls.Add(Me.Button_RESET)
        Me.Controls.Add(Me.TextBox1)
        Me.Controls.Add(Me.TabControl1)
        Me.Controls.Add(Me.Button_Cancel)
        Me.Controls.Add(Me.Button_Ok)
        Me.Controls.Add(Me.Button_Test)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Name = "RejectPoint"
        Me.Text = "RejectPoint"
        Me.TopMost = True
        Me.TabControl1.ResumeLayout(False)
        Me.CreateRejectMode.ResumeLayout(False)
        Me.GroupBox3.ResumeLayout(False)
        CType(Me.ValueRatio, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox2.ResumeLayout(False)
        CType(Me.ValueBrightness1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.ValueBinarizedThreshold, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SelectRejectMode.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        CType(Me.ValueAcceptRatio, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox4.ResumeLayout(False)
        CType(Me.ValueBrightness, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.ValueBinarized, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region
#Region "Global Variables"
    Dim BlackWithoutRM As Boolean = False
    Dim WhiteWithoutRM As Boolean = False
    Dim BlackWithRM As Boolean = False
    Dim WhiteWithRM As Boolean = False
    Dim BlackNo, WhiteNo As Long
    Dim DisplayCenterXPosition, DisplayCenterYPosition, MROIx, MROIy As Double
    Dim Folder As String = "c:\"

    '''''''''''''''''
    '   Xue Wen     '
    '''''''''''''''''
    Dim picture1Enabled As Boolean = False
#End Region

    Private Shared m_instance As RejectPoint

    Public Shared ReadOnly Property Instance() As RejectPoint
        Get
            If m_instance Is Nothing Then
                m_instance = New RejectPoint
            Else
                If m_instance.IsDisposed Then
                    m_instance = New RejectPoint
                End If
            End If
            Return m_instance
        End Get

    End Property

    Public Structure RMParam
        Public _Brightness As Integer
        Public _Binarized As Integer
        Public _WoRM As Boolean
        Public _WRM As Boolean
        Public _AcceptRatio As Double
        Public _BlackWithoutRM As Boolean
        Public _WhiteWithoutRM As Boolean
        Public _BlackWithRM As Boolean
        Public _WhiteWithRM As Boolean
        Public _MRegionX As Double
        Public _MRegionY As Double
        Public _MROIx As Double
        Public _MROIy As Double
    End Structure
    Sub SetRMFilename(ByVal PathName As String)
        Folder = PathName
    End Sub
    Dim Status = 0
    Function GetRMStatus() As Integer '0=yet done, 1=done, 2=cancel
        Return Status
    End Function
    Sub ResetRMStatus()
        Status = 0
    End Sub
    Function GetRMParameters(ByRef RMParam As RMParam)
        RMParam._Brightness = ValueBrightness1.Value
        RMParam._Binarized = ValueBinarizedThreshold.Value
        RMParam._WoRM = RadioButton_WithoutRejectMark.Checked
        RMParam._WRM = RadioButton_RejectMark.Checked
        RMParam._AcceptRatio = ValueRatio.Value
        RMParam._BlackWithoutRM = BlackWithoutRM
        RMParam._WhiteWithoutRM = WhiteWithoutRM
        RMParam._BlackWithRM = BlackWithRM
        RMParam._WhiteWithRM = WhiteWithRM
        RMParam._MRegionX = DisplayCenterXPosition
        RMParam._MRegionY = DisplayCenterYPosition
        RMParam._MROIx = MROIx
        RMParam._MROIy = MROIy
    End Function

    Sub TextboxDisplay()
        If TabControl1.SelectedIndex = 0 Then
            TextBox1.Text = "Select ROI and Test it"
        ElseIf TabControl1.SelectedIndex = 1 Then
            TextBox1.Text = "Load a model and Test it"
        End If
    End Sub
    Sub BlacknWhite(ByVal Black As Long, ByVal White As Long)
        BlackNo = Black
        WhiteNo = White
    End Sub
    Sub BlackWhite(ByVal FrmBlackWithoutRM As Boolean, ByVal FrmWhiteWithoutRM As Boolean, ByVal FrmBlackWithRM As Boolean, ByVal FrmWhiteWithRM As Boolean)
        BlackWithoutRM = FrmBlackWithoutRM
        WhiteWithoutRM = FrmWhiteWithoutRM
        BlackWithRM = FrmBlackWithRM
        WhiteWithRM = FrmWhiteWithRM
    End Sub
    Sub GetVariables(ByVal RDisplayCenterXPosition, ByVal RDisplayCenterYPosition, ByVal RMroix, ByVal RMroiy)
        DisplayCenterXPosition = RDisplayCenterXPosition
        DisplayCenterYPosition = RDisplayCenterYPosition
        MROIx = RMroix
        MROIy = RMroiy
    End Sub
    Function CheckFileName(ByVal FileName) As Boolean
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
        NewFullPath = Folder & "." & FileExt
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
    End Function

    Private Sub TabControl1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabControl1.SelectedIndexChanged
        If TabControl1.SelectedIndex = 0 Then
            TextboxDisplay()
            FrmVision.Reset()
            FrmVision.modelregionDrawing()
            If PictureBox2.Image Is Nothing Then
            Else
                Dim image As Image = PictureBox2.Image
                PictureBox2.Image = Nothing
                image.Dispose()
            End If
            RadioButton_WithoutRejectMark.Enabled = True
            RadioButton_RejectMark.Enabled = True
            ValueBrightness1.Value = 128
            ValueBinarizedThreshold.Value = 128
            RadioButton_WithoutRejectMark.Checked = True
            RadioButton_RejectMark.Checked = False
            ValueRatio.Value = 5
            BlackWithoutRM = False
            WhiteWithoutRM = False
            BlackWithRM = False
            WhiteWithRM = False
            WhiteNo = 0
            BlackNo = 0
            Button_Next.Enabled = True

            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '   Xue Wen                                                     '
            '   Upload previous picture and update the related button.      '
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            If (picture1Enabled = True) Then
                PictureBox1.Image = Image.FromFile(Folder + "RM.bmp")
                Button_Test.Enabled = True
                picture1Enabled = False
                Button_Ok.Enabled = True
            Else
                Button_Test.Enabled = False
                Button_Ok.Enabled = False
            End If
        ElseIf TabControl1.SelectedIndex = 1 Then
            Try
                If FrmVision.AxDisplay1.IsAllocated = False Then
                Else
                    FrmVision.ClearDisplay()
                    TextboxDisplay()
                End If
                If PictureBox1.Image Is Nothing Then
                Else
                    Dim image As Image = PictureBox1.Image
                    PictureBox1.Image = Nothing
                    image.Dispose()

                    '''''''''''''''''''''''''''''''''''''''''
                    '   Xue Wen                             '
                    '   Flag for uploading old picture.     '
                    '''''''''''''''''''''''''''''''''''''''''
                    picture1Enabled = True
                End If
                'RadioButton_WoRM.Enabled = True
                'RadioButton_WRM.Enabled = True
                ValueBrightness.Value = 128
                ValueBinarized.Value = 128
                'RadioButton_WoRM.Checked = True
                'RadioButton_WRM.Checked = False
                ValueAcceptRatio.Value = 5
                BlackWithoutRM = False
                WhiteWithoutRM = False
                BlackWithRM = False
                WhiteWithRM = False
                WhiteNo = 0
                BlackNo = 0
                'Button5.Enabled = False
                Button_Load.Enabled = False 'True
                Button_Test.Enabled = False
                FrmVision.Reset()
                Button_Ok.Enabled = False
            Catch ex As Exception
                ExceptionDisplay(ex)
            End Try
        End If
    End Sub
    Private Sub Button_Next_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_Next.Click
        If BlackWithoutRM = True Or WhiteWithoutRM = True Or BlackWithRM = True Or WhiteWithRM = True Then
            BlackWithoutRM = False
            WhiteWithoutRM = False
            BlackWithRM = False
            WhiteWithRM = False
        End If
        FrmVision.RejectPOintROI()
        If RadioButton_WithoutRejectMark.Checked = True Then
            If BlackNo <> 0 And WhiteNo <> 0 Then
                TextBox1.Text = "Please adjust Threshold values."
            ElseIf BlackNo = 0 And WhiteNo = 0 Then
                '''''''''''''''''
                '   Origin      '
                '''''''''''''''''
                'TextBox1.Text = "Please contact EET for help"

                TextBox1.Text = "Please contact the technical person for help."
            ElseIf BlackNo > 0 Then
                BlackWithoutRM = True
                Button_Next.Enabled = False
                RadioButton_WithoutRejectMark.Enabled = False
                RadioButton_RejectMark.Enabled = False
                Button_Test.Enabled = True
            ElseIf WhiteNo > 0 Then
                WhiteWithoutRM = True
                RadioButton_WithoutRejectMark.Enabled = False
                RadioButton_RejectMark.Enabled = False
                Button_Next.Enabled = False
                Button_Test.Enabled = True
            End If
        ElseIf RadioButton_RejectMark.Checked = True Then
            If BlackNo <> 0 And WhiteNo <> 0 Then
                If BlackNo > WhiteNo Then
                    BlackWithRM = True
                    RadioButton_WithoutRejectMark.Enabled = False
                    RadioButton_RejectMark.Enabled = False
                    Button_Next.Enabled = False
                    Button_Test.Enabled = True
                ElseIf WhiteNo > BlackNo Then
                    WhiteWithRM = True
                    RadioButton_WithoutRejectMark.Enabled = False
                    RadioButton_RejectMark.Enabled = False
                    Button_Next.Enabled = False
                    Button_Test.Enabled = True
                End If
            ElseIf BlackNo = 0 And WhiteNo = 0 Then
                TextBox1.Text = "Please contact the technical person for help."
            ElseIf BlackNo = 0 Or WhiteNo = 0 Then
                TextBox1.Text = "Please adjust Threshold values."
            End If
        End If
        PictureBox1.Image = Image.FromFile(Folder + "RM.bmp")
        Button_Ok.Enabled = True
    End Sub
    Private Sub Button_Test_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_Test.Click
        ' BlackWithRM refers to an image taken with a black pixel majority
        Dim WhiteAcceptanceRatio As Double
        Dim BlackAcceptanceRatio As Double
        WhiteAcceptanceRatio = (WhiteNo / (BlackNo + WhiteNo) * 100)
        BlackAcceptanceRatio = (BlackNo / (BlackNo + WhiteNo) * 100)
        If TabControl1.SelectedIndex = 0 Then
            FrmVision.RunReject(MROIx, MROIy)
            MsgBox("White ratio = " & WhiteAcceptanceRatio & " Black ratio = " & BlackAcceptanceRatio & " RejectMark(W/B) = " & WhiteWithRM & " " & BlackWithRM)
            'PictureBox2.Image = Image.FromFile("C:\IDS\SRC\DLL Export Device Vision\RejectPoint\1.bmp")
            If BlackWithoutRM = True Then 'black on white
                If (WhiteNo / (BlackNo + WhiteNo) * 100) > ValueRatio.Value Then
                    TextBox1.Text = "reject black w/o1"
                Else
                    TextBox1.Text = "Accept black w/o1"
                End If
            ElseIf WhiteWithoutRM = True Then 'white on black
                If (BlackNo / (BlackNo + WhiteNo) * 100) > ValueRatio.Value Then
                    TextBox1.Text = "reject white w/o2"
                Else
                    TextBox1.Text = "Accept white w/o2"
                End If
            ElseIf BlackWithRM = True Then
                If (WhiteNo / (BlackNo + WhiteNo) * 100) < ValueRatio.Value Then
                    TextBox1.Text = "Accept black w1"
                Else
                    TextBox1.Text = "reject black w1"
                End If
            ElseIf WhiteWithRM = True Then
                If (BlackNo / (BlackNo + WhiteNo) * 100) < ValueRatio.Value Then
                    TextBox1.Text = "Accept white w2"
                Else
                    TextBox1.Text = "reject white w2"
                End If
            End If
        Else
            FrmVision.RejectPOintROI()
            PictureBox1.Image = Image.FromFile(Folder + "RM.bmp")
            If BlackWithoutRM = True Then
                If (WhiteNo / (BlackNo + WhiteNo) * 100) > ValueAcceptRatio.Value Then
                    TextBox1.Text = "reject"
                Else
                    TextBox1.Text = "Accept"
                End If
            ElseIf WhiteWithoutRM = True Then
                If (BlackNo / (BlackNo + WhiteNo) * 100) > ValueAcceptRatio.Value Then
                    TextBox1.Text = "reject"
                Else
                    TextBox1.Text = "Accept"
                End If
            ElseIf BlackWithRM = True Then
                If (WhiteNo / (BlackNo + WhiteNo) * 100) < ValueAcceptRatio.Value Then
                    TextBox1.Text = "Accept"
                Else
                    TextBox1.Text = "reject"
                End If
            ElseIf WhiteWithRM = True Then
                If (BlackNo / (BlackNo + WhiteNo) * 100) < ValueAcceptRatio.Value Then
                    TextBox1.Text = "Accept"
                Else
                    TextBox1.Text = "reject"
                End If
            End If
        End If

    End Sub
    Private Sub RejectPoint_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.Visible = True
        'Button5.Enabled = True
        Button_Test.Enabled = False
        If TabControl1.SelectedIndex = 0 Then
            TextboxDisplay()
            FrmVision.ClearDisplay()
            FrmVision.modelregionDrawing()
        ElseIf TabControl1.SelectedIndex = 1 Then
            Try
                'FrmVision.ClearDisplay
                'TextboxDisplay()
            Catch ex As Exception
                ExceptionDisplay(ex)
            End Try
        End If
    End Sub
    Private Sub Button_Cancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button_Cancel.Click
        Status = 2
        ResetVariables()
        FrmVision.DisplayIndicator()
        FrmVision.ResetReject_Flag()
        Me.Visible = False
    End Sub

    Sub SetRMReset()
        ResetVariables()
    End Sub
    Sub ResetVariables()
        Button_Load.Enabled = True
        ValueBrightness1.Value = 128
        ValueBinarizedThreshold.Value = 128
        RadioButton_WithoutRejectMark.Enabled = True
        RadioButton_RejectMark.Enabled = True
        RadioButton_WithoutRejectMark.Checked = True
        RadioButton_RejectMark.Checked = False
        ValueRatio.Value = 5
        BlackWithoutRM = False
        WhiteWithoutRM = False
        BlackWithRM = False
        WhiteWithRM = False
        WhiteNo = 0
        BlackNo = 0
        TabControl1.SelectedIndex = 0
        Button_Next.Enabled = False
        Button_Test.Enabled = False
        If PictureBox1.Image Is Nothing Then
        Else
            Dim image As Image = PictureBox1.Image
            PictureBox1.Image = Nothing
            image.Dispose()
        End If
        'FrmVision.DisplayIndicator()
        FrmVision.Reset()
    End Sub

    Private Sub Button_Ok_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_Ok.Click
        'save variables
        'WriteFileInText() 'For testing
        Status = 1
        Dim RMParam As RMParam
        FrmVision.GetRMParameters(RMParam)
        FrmVision.DisplayIndicator()
        FrmVision.ResetReject_Flag()
        'ResetVariables()
        Me.Visible = False
    End Sub
    Public Function FileNameWithoutExtension(ByVal FullPath As String) As String
        Return System.IO.Path.GetDirectoryName(FullPath)
    End Function

#Region "Production"
    Function ProductionRUN() As Boolean
        FrmVision.RunReject(MROIx, MROIy)
        'PictureBox2.Image = Image.FromFile("C:\IDS\SRC\DLL Export Device Vision\RejectPoint\1.bmp")
        FrmVision.DisplayIndicator()
        If BlackWithoutRM = True Then
            If (WhiteNo / (BlackNo + WhiteNo) * 100) > ValueAcceptRatio.Value Then
                TextBox1.Text = "reject"
                'MsgBox("Reject")
                Return False
            Else
                TextBox1.Text = "Accept"
                'MsgBox("Accept")
                Return True
            End If

        ElseIf WhiteWithoutRM = True Then
            If (BlackNo / (BlackNo + WhiteNo) * 100) > ValueAcceptRatio.Value Then
                TextBox1.Text = "reject"
                'MsgBox("Reject")
                Return False
            Else
                TextBox1.Text = "Accept"
                'MsgBox("Accept")
                Return True
            End If
        ElseIf BlackWithRM = True Then
            If (WhiteNo / (BlackNo + WhiteNo) * 100) < ValueAcceptRatio.Value Then
                TextBox1.Text = "Accept"
                'MsgBox("Accept")
                Return True
            Else
                TextBox1.Text = "reject"
                'MsgBox("Reject")
                Return False
            End If
        ElseIf WhiteWithRM = True Then
            If (BlackNo / (BlackNo + WhiteNo) * 100) < ValueAcceptRatio.Value Then
                TextBox1.Text = "Accept"
                'MsgBox("Accept")
                Return True
            Else
                TextBox1.Text = "allfalse" 'kr changed, original is "reject"
                'MsgBox("Reject")
                Return False
            End If
        End If
    End Function
#End Region

    Private Sub ValueBrightness_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ValueBrightness.ValueChanged
        If ValueBrightness.Focused = True Then
            ValueBrightness1.Value = ValueBrightness.Value
        End If
    End Sub
    Private Sub ValueBinarized_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ValueBinarized.ValueChanged
        ValueBinarizedThreshold.Value = ValueBinarized.Value
    End Sub
    Private Sub ValueAcceptRatio_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ValueAcceptRatio.ValueChanged
        ValueRatio.Value = ValueAcceptRatio.Value
    End Sub
    Private Sub Button_RESET_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_RESET.Click
        'ResetVariables()
        TextboxDisplay()
        FrmVision.Reset()
        FrmVision.ClearDisplay()
        FrmVision.modelregionDrawing()

        '''''''''''''''''
        '   Origin      '
        '''''''''''''''''
        'If PictureBox2.Image Is Nothing Then
        'Else
        '    Dim image As Image = PictureBox2.Image
        '    PictureBox2.Image = Nothing
        '    image.Dispose()
        'End If

        If PictureBox1.Image Is Nothing Then
        Else
            Dim image As Image = PictureBox1.Image
            PictureBox1.Image = Nothing
            image.Dispose()
        End If

        If PictureBox2.Image Is Nothing Then
        Else
            Dim image2 As Image = PictureBox2.Image
            PictureBox2.Image = Nothing
            image2.Dispose()
        End If

        RadioButton_WithoutRejectMark.Enabled = True
        RadioButton_RejectMark.Enabled = True
        ValueBrightness1.Value = 128
        ValueBinarizedThreshold.Value = 128
        RadioButton_WithoutRejectMark.Checked = True
        RadioButton_RejectMark.Checked = False
        ValueRatio.Value = 5
        BlackWithoutRM = False
        WhiteWithoutRM = False
        BlackWithRM = False
        WhiteWithRM = False
        WhiteNo = 0
        BlackNo = 0
        Button_Next.Enabled = True
        Button_Test.Enabled = False
        Button_Ok.Enabled = False
    End Sub

    Private Sub ValueBrightness1_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ValueBrightness1.ValueChanged
        If ValueBrightness1.Focused = True Then   'Xu Long 20/02/2012
            FrmVision.SetBrightness(ValueBrightness1.Value)
        End If

        'If Initializing Then Return

        'FrmVision.SetBrightness(ValueBrightness1.Value)
    End Sub
End Class
