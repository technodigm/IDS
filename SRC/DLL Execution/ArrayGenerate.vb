Imports DLL_DataManager
Imports DLL_Export_Service

Public Class ArrayGenerate
    Inherits System.Windows.Forms.Form
    Dim FormIni As Boolean = False
    Dim CurrentRow As Integer = 0
    Public ArrayPara1 As New CIDSData.CIDSPatternExecution.CIDSTemplate
    Public ArrayPara2 As New CIDSData.CIDSPatternExecution.CIDSTemplate
    Public ArrayPara3 As New CIDSData.CIDSPatternExecution.CIDSTemplate
    Public isVisionMode As Boolean = True

#Region " Windows Form Designer generated code "
    Dim m_CurrentPara As New CIDSData.CIDSPatternExecution.CIDSTemplate

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        Button_OnOK.Enabled = False
        Button_OnCancel.Enabled = True
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
    Friend WithEvents Button_OnOK As System.Windows.Forms.Button
    Friend WithEvents Button_OnCancel As System.Windows.Forms.Button
    Friend WithEvents CombElementType As System.Windows.Forms.ComboBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents Label19 As System.Windows.Forms.Label
    Friend WithEvents Label22 As System.Windows.Forms.Label
    Friend WithEvents Label23 As System.Windows.Forms.Label
    Friend WithEvents Label24 As System.Windows.Forms.Label
    Friend WithEvents Label25 As System.Windows.Forms.Label
    Friend WithEvents Label26 As System.Windows.Forms.Label
    Friend WithEvents Label27 As System.Windows.Forms.Label
    Friend WithEvents Label28 As System.Windows.Forms.Label
    Friend WithEvents Label29 As System.Windows.Forms.Label
    Friend WithEvents Label30 As System.Windows.Forms.Label
    Friend WithEvents Label31 As System.Windows.Forms.Label
    Friend WithEvents Label32 As System.Windows.Forms.Label
    Friend WithEvents Label33 As System.Windows.Forms.Label
    Friend WithEvents Label34 As System.Windows.Forms.Label
    Friend WithEvents Label35 As System.Windows.Forms.Label
    Friend WithEvents Label36 As System.Windows.Forms.Label
    Friend WithEvents Label37 As System.Windows.Forms.Label
    Friend WithEvents Label38 As System.Windows.Forms.Label
    Friend WithEvents Label39 As System.Windows.Forms.Label
    Friend WithEvents TextBox_P1X As System.Windows.Forms.TextBox
    Friend WithEvents TextBox_P1Y As System.Windows.Forms.TextBox
    Friend WithEvents TextBox_P1Z As System.Windows.Forms.TextBox
    Friend WithEvents TextBox_P2X As System.Windows.Forms.TextBox
    Friend WithEvents TextBox_P2Y As System.Windows.Forms.TextBox
    Friend WithEvents TextBox_P3Y As System.Windows.Forms.TextBox
    Friend WithEvents TextBox_P3X As System.Windows.Forms.TextBox
    Friend WithEvents TextBox_P4Y As System.Windows.Forms.TextBox
    Friend WithEvents TextBox_P4X As System.Windows.Forms.TextBox
    Friend WithEvents RadioButton_P1 As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton_P2 As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton_P3 As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton_P4 As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton_P7 As System.Windows.Forms.RadioButton
    Friend WithEvents TextBox_P7Y As System.Windows.Forms.TextBox
    Friend WithEvents TextBox_P7X As System.Windows.Forms.TextBox
    Friend WithEvents TextBox_Needle As System.Windows.Forms.TextBox
    Friend WithEvents TextBox_Dispense As System.Windows.Forms.TextBox
    Friend WithEvents TextBox_TravelSpeed As System.Windows.Forms.TextBox
    Friend WithEvents TextBox_NeedleGap As System.Windows.Forms.TextBox
    Friend WithEvents TextBox_ApproachHeight As System.Windows.Forms.TextBox
    Friend WithEvents TextBox_RetractDelay As System.Windows.Forms.TextBox
    Friend WithEvents TextBox_TravelDelay As System.Windows.Forms.TextBox
    Friend WithEvents TextBox_Duration As System.Windows.Forms.TextBox
    Friend WithEvents TextBox_NoOfRun As System.Windows.Forms.TextBox
    Friend WithEvents TextBox_FillHeight As System.Windows.Forms.TextBox
    Friend WithEvents TextBox_Pitch As System.Windows.Forms.TextBox
    Friend WithEvents TextBox_ArcRadius As System.Windows.Forms.TextBox
    Friend WithEvents TextBox_DTailDist As System.Windows.Forms.TextBox
    Friend WithEvents TextBox_ClearanceHt As System.Windows.Forms.TextBox
    Friend WithEvents TextBox_RetractHeight As System.Windows.Forms.TextBox
    Friend WithEvents TextBox_RetractSpeed As System.Windows.Forms.TextBox
    Friend WithEvents TextBox_IONumber As System.Windows.Forms.TextBox
    Friend WithEvents TextBox_EdgeClear As System.Windows.Forms.TextBox
    Friend WithEvents TextBox_RtAngle As System.Windows.Forms.TextBox
    Friend WithEvents TextBox_Sprial As System.Windows.Forms.TextBox
    Friend WithEvents lbDispenseMode As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.Button_OnOK = New System.Windows.Forms.Button
        Me.Button_OnCancel = New System.Windows.Forms.Button
        Me.CombElementType = New System.Windows.Forms.ComboBox
        Me.TextBox_Needle = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.lbDispenseMode = New System.Windows.Forms.Label
        Me.TextBox_Dispense = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.TextBox_TravelSpeed = New System.Windows.Forms.TextBox
        Me.TextBox_NeedleGap = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.TextBox_P1X = New System.Windows.Forms.TextBox
        Me.TextBox_P1Y = New System.Windows.Forms.TextBox
        Me.TextBox_P1Z = New System.Windows.Forms.TextBox
        Me.TextBox_P2X = New System.Windows.Forms.TextBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.TextBox_P2Y = New System.Windows.Forms.TextBox
        Me.TextBox_P3Y = New System.Windows.Forms.TextBox
        Me.Label7 = New System.Windows.Forms.Label
        Me.TextBox_P3X = New System.Windows.Forms.TextBox
        Me.Label8 = New System.Windows.Forms.Label
        Me.Label9 = New System.Windows.Forms.Label
        Me.Label10 = New System.Windows.Forms.Label
        Me.Label11 = New System.Windows.Forms.Label
        Me.Label13 = New System.Windows.Forms.Label
        Me.Label14 = New System.Windows.Forms.Label
        Me.Label16 = New System.Windows.Forms.Label
        Me.TextBox_P4Y = New System.Windows.Forms.TextBox
        Me.TextBox_P4X = New System.Windows.Forms.TextBox
        Me.Label19 = New System.Windows.Forms.Label
        Me.TextBox_P7Y = New System.Windows.Forms.TextBox
        Me.TextBox_P7X = New System.Windows.Forms.TextBox
        Me.Label22 = New System.Windows.Forms.Label
        Me.Label23 = New System.Windows.Forms.Label
        Me.TextBox_ApproachHeight = New System.Windows.Forms.TextBox
        Me.TextBox_RetractDelay = New System.Windows.Forms.TextBox
        Me.Label24 = New System.Windows.Forms.Label
        Me.TextBox_TravelDelay = New System.Windows.Forms.TextBox
        Me.Label25 = New System.Windows.Forms.Label
        Me.Label26 = New System.Windows.Forms.Label
        Me.TextBox_Duration = New System.Windows.Forms.TextBox
        Me.Label27 = New System.Windows.Forms.Label
        Me.Label28 = New System.Windows.Forms.Label
        Me.TextBox_NoOfRun = New System.Windows.Forms.TextBox
        Me.TextBox_FillHeight = New System.Windows.Forms.TextBox
        Me.Label29 = New System.Windows.Forms.Label
        Me.TextBox_Pitch = New System.Windows.Forms.TextBox
        Me.Label30 = New System.Windows.Forms.Label
        Me.Label31 = New System.Windows.Forms.Label
        Me.TextBox_ArcRadius = New System.Windows.Forms.TextBox
        Me.Label32 = New System.Windows.Forms.Label
        Me.TextBox_DTailDist = New System.Windows.Forms.TextBox
        Me.TextBox_ClearanceHt = New System.Windows.Forms.TextBox
        Me.Label33 = New System.Windows.Forms.Label
        Me.TextBox_RetractHeight = New System.Windows.Forms.TextBox
        Me.Label34 = New System.Windows.Forms.Label
        Me.Label35 = New System.Windows.Forms.Label
        Me.TextBox_RetractSpeed = New System.Windows.Forms.TextBox
        Me.TextBox_IONumber = New System.Windows.Forms.TextBox
        Me.Label36 = New System.Windows.Forms.Label
        Me.Label37 = New System.Windows.Forms.Label
        Me.TextBox_EdgeClear = New System.Windows.Forms.TextBox
        Me.Label38 = New System.Windows.Forms.Label
        Me.TextBox_RtAngle = New System.Windows.Forms.TextBox
        Me.TextBox_Sprial = New System.Windows.Forms.TextBox
        Me.Label39 = New System.Windows.Forms.Label
        Me.RadioButton_P1 = New System.Windows.Forms.RadioButton
        Me.RadioButton_P2 = New System.Windows.Forms.RadioButton
        Me.RadioButton_P3 = New System.Windows.Forms.RadioButton
        Me.RadioButton_P4 = New System.Windows.Forms.RadioButton
        Me.RadioButton_P7 = New System.Windows.Forms.RadioButton
        Me.SuspendLayout()
        '
        'Button_OnOK
        '
        Me.Button_OnOK.Location = New System.Drawing.Point(1152, 184)
        Me.Button_OnOK.Name = "Button_OnOK"
        Me.Button_OnOK.Size = New System.Drawing.Size(88, 32)
        Me.Button_OnOK.TabIndex = 0
        Me.Button_OnOK.Text = "Ok"
        '
        'Button_OnCancel
        '
        Me.Button_OnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Button_OnCancel.Location = New System.Drawing.Point(1032, 184)
        Me.Button_OnCancel.Name = "Button_OnCancel"
        Me.Button_OnCancel.Size = New System.Drawing.Size(96, 32)
        Me.Button_OnCancel.TabIndex = 1
        Me.Button_OnCancel.Text = "Cancel"
        '
        'CombElementType
        '
        Me.CombElementType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CombElementType.Items.AddRange(New Object() {"Dot", "Line", "Arc", "Circle", "FillCircle", "Rectangle", "FillRectangle", "SubPattern"})
        Me.CombElementType.Location = New System.Drawing.Point(120, 48)
        Me.CombElementType.Name = "CombElementType"
        Me.CombElementType.Size = New System.Drawing.Size(104, 21)
        Me.CombElementType.TabIndex = 2
        '
        'TextBox_Needle
        '
        Me.TextBox_Needle.Location = New System.Drawing.Point(272, 48)
        Me.TextBox_Needle.Name = "TextBox_Needle"
        Me.TextBox_Needle.Size = New System.Drawing.Size(72, 20)
        Me.TextBox_Needle.TabIndex = 3
        Me.TextBox_Needle.Text = "Right"
        Me.TextBox_Needle.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(288, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(48, 24)
        Me.Label1.TabIndex = 4
        Me.Label1.Text = "Needle"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lbDispenseMode
        '
        Me.lbDispenseMode.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbDispenseMode.Location = New System.Drawing.Point(448, 16)
        Me.lbDispenseMode.Name = "lbDispenseMode"
        Me.lbDispenseMode.Size = New System.Drawing.Size(248, 24)
        Me.lbDispenseMode.TabIndex = 5
        Me.lbDispenseMode.Text = "Dispense Mode / SubPattern name"
        Me.lbDispenseMode.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'TextBox_Dispense
        '
        Me.TextBox_Dispense.Location = New System.Drawing.Point(360, 48)
        Me.TextBox_Dispense.Name = "TextBox_Dispense"
        Me.TextBox_Dispense.Size = New System.Drawing.Size(392, 20)
        Me.TextBox_Dispense.TabIndex = 6
        Me.TextBox_Dispense.Text = "On"
        '
        'Label3
        '
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(760, 16)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(56, 24)
        Me.Label3.TabIndex = 7
        Me.Label3.Text = "Travel Speed"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'TextBox_TravelSpeed
        '
        Me.TextBox_TravelSpeed.Location = New System.Drawing.Point(760, 48)
        Me.TextBox_TravelSpeed.Name = "TextBox_TravelSpeed"
        Me.TextBox_TravelSpeed.Size = New System.Drawing.Size(56, 20)
        Me.TextBox_TravelSpeed.TabIndex = 8
        Me.TextBox_TravelSpeed.Text = ""
        '
        'TextBox_NeedleGap
        '
        Me.TextBox_NeedleGap.Location = New System.Drawing.Point(824, 48)
        Me.TextBox_NeedleGap.Name = "TextBox_NeedleGap"
        Me.TextBox_NeedleGap.Size = New System.Drawing.Size(56, 20)
        Me.TextBox_NeedleGap.TabIndex = 9
        Me.TextBox_NeedleGap.Text = ""
        '
        'Label4
        '
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(824, 16)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(56, 24)
        Me.Label4.TabIndex = 10
        Me.Label4.Text = "Needle Gap"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(8, 128)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(32, 16)
        Me.Label5.TabIndex = 11
        Me.Label5.Text = "Pt1"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'TextBox_P1X
        '
        Me.TextBox_P1X.Location = New System.Drawing.Point(48, 128)
        Me.TextBox_P1X.Name = "TextBox_P1X"
        Me.TextBox_P1X.ReadOnly = True
        Me.TextBox_P1X.Size = New System.Drawing.Size(56, 20)
        Me.TextBox_P1X.TabIndex = 12
        Me.TextBox_P1X.Text = "0"
        Me.TextBox_P1X.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'TextBox_P1Y
        '
        Me.TextBox_P1Y.Location = New System.Drawing.Point(112, 128)
        Me.TextBox_P1Y.Name = "TextBox_P1Y"
        Me.TextBox_P1Y.ReadOnly = True
        Me.TextBox_P1Y.Size = New System.Drawing.Size(56, 20)
        Me.TextBox_P1Y.TabIndex = 13
        Me.TextBox_P1Y.Text = "0"
        Me.TextBox_P1Y.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'TextBox_P1Z
        '
        Me.TextBox_P1Z.Location = New System.Drawing.Point(176, 128)
        Me.TextBox_P1Z.Name = "TextBox_P1Z"
        Me.TextBox_P1Z.ReadOnly = True
        Me.TextBox_P1Z.Size = New System.Drawing.Size(56, 20)
        Me.TextBox_P1Z.TabIndex = 14
        Me.TextBox_P1Z.Text = "0"
        Me.TextBox_P1Z.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'TextBox_P2X
        '
        Me.TextBox_P2X.Location = New System.Drawing.Point(296, 128)
        Me.TextBox_P2X.Name = "TextBox_P2X"
        Me.TextBox_P2X.ReadOnly = True
        Me.TextBox_P2X.Size = New System.Drawing.Size(56, 20)
        Me.TextBox_P2X.TabIndex = 15
        Me.TextBox_P2X.Text = "0"
        Me.TextBox_P2X.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label6
        '
        Me.Label6.Location = New System.Drawing.Point(248, 128)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(32, 16)
        Me.Label6.TabIndex = 16
        Me.Label6.Text = "Pt2"
        Me.Label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'TextBox_P2Y
        '
        Me.TextBox_P2Y.Location = New System.Drawing.Point(360, 128)
        Me.TextBox_P2Y.Name = "TextBox_P2Y"
        Me.TextBox_P2Y.ReadOnly = True
        Me.TextBox_P2Y.Size = New System.Drawing.Size(56, 20)
        Me.TextBox_P2Y.TabIndex = 17
        Me.TextBox_P2Y.Text = "0"
        Me.TextBox_P2Y.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'TextBox_P3Y
        '
        Me.TextBox_P3Y.Location = New System.Drawing.Point(544, 128)
        Me.TextBox_P3Y.Name = "TextBox_P3Y"
        Me.TextBox_P3Y.ReadOnly = True
        Me.TextBox_P3Y.Size = New System.Drawing.Size(56, 20)
        Me.TextBox_P3Y.TabIndex = 21
        Me.TextBox_P3Y.Text = "0"
        Me.TextBox_P3Y.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label7
        '
        Me.Label7.Location = New System.Drawing.Point(432, 128)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(32, 16)
        Me.Label7.TabIndex = 20
        Me.Label7.Text = "Pt3"
        Me.Label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'TextBox_P3X
        '
        Me.TextBox_P3X.Location = New System.Drawing.Point(480, 128)
        Me.TextBox_P3X.Name = "TextBox_P3X"
        Me.TextBox_P3X.ReadOnly = True
        Me.TextBox_P3X.Size = New System.Drawing.Size(56, 20)
        Me.TextBox_P3X.TabIndex = 19
        Me.TextBox_P3X.Text = "0"
        Me.TextBox_P3X.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label8
        '
        Me.Label8.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label8.ForeColor = System.Drawing.Color.RoyalBlue
        Me.Label8.Location = New System.Drawing.Point(64, 104)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(32, 16)
        Me.Label8.TabIndex = 23
        Me.Label8.Text = "X"
        Me.Label8.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label9
        '
        Me.Label9.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label9.ForeColor = System.Drawing.Color.RoyalBlue
        Me.Label9.Location = New System.Drawing.Point(192, 104)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(32, 16)
        Me.Label9.TabIndex = 24
        Me.Label9.Text = "Z"
        Me.Label9.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label10
        '
        Me.Label10.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label10.ForeColor = System.Drawing.Color.RoyalBlue
        Me.Label10.Location = New System.Drawing.Point(128, 104)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(32, 16)
        Me.Label10.TabIndex = 25
        Me.Label10.Text = "Y"
        Me.Label10.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label11
        '
        Me.Label11.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label11.ForeColor = System.Drawing.Color.RoyalBlue
        Me.Label11.Location = New System.Drawing.Point(376, 104)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(32, 16)
        Me.Label11.TabIndex = 28
        Me.Label11.Text = "Y"
        Me.Label11.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label13
        '
        Me.Label13.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label13.ForeColor = System.Drawing.Color.RoyalBlue
        Me.Label13.Location = New System.Drawing.Point(312, 104)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(32, 16)
        Me.Label13.TabIndex = 26
        Me.Label13.Text = "X"
        Me.Label13.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label14
        '
        Me.Label14.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label14.ForeColor = System.Drawing.Color.RoyalBlue
        Me.Label14.Location = New System.Drawing.Point(560, 104)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(32, 16)
        Me.Label14.TabIndex = 31
        Me.Label14.Text = "Y"
        Me.Label14.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label16
        '
        Me.Label16.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label16.ForeColor = System.Drawing.Color.RoyalBlue
        Me.Label16.Location = New System.Drawing.Point(496, 104)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(32, 16)
        Me.Label16.TabIndex = 29
        Me.Label16.Text = "X"
        Me.Label16.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'TextBox_P4Y
        '
        Me.TextBox_P4Y.Location = New System.Drawing.Point(112, 160)
        Me.TextBox_P4Y.Name = "TextBox_P4Y"
        Me.TextBox_P4Y.ReadOnly = True
        Me.TextBox_P4Y.Size = New System.Drawing.Size(56, 20)
        Me.TextBox_P4Y.TabIndex = 34
        Me.TextBox_P4Y.Text = "0"
        Me.TextBox_P4Y.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'TextBox_P4X
        '
        Me.TextBox_P4X.Location = New System.Drawing.Point(48, 160)
        Me.TextBox_P4X.Name = "TextBox_P4X"
        Me.TextBox_P4X.ReadOnly = True
        Me.TextBox_P4X.Size = New System.Drawing.Size(56, 20)
        Me.TextBox_P4X.TabIndex = 33
        Me.TextBox_P4X.Text = "0"
        Me.TextBox_P4X.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label19
        '
        Me.Label19.Location = New System.Drawing.Point(8, 160)
        Me.Label19.Name = "Label19"
        Me.Label19.Size = New System.Drawing.Size(32, 16)
        Me.Label19.TabIndex = 32
        Me.Label19.Text = "Pt4"
        Me.Label19.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'TextBox_P7Y
        '
        Me.TextBox_P7Y.Location = New System.Drawing.Point(112, 192)
        Me.TextBox_P7Y.Name = "TextBox_P7Y"
        Me.TextBox_P7Y.ReadOnly = True
        Me.TextBox_P7Y.Size = New System.Drawing.Size(56, 20)
        Me.TextBox_P7Y.TabIndex = 46
        Me.TextBox_P7Y.Text = "0"
        Me.TextBox_P7Y.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'TextBox_P7X
        '
        Me.TextBox_P7X.Location = New System.Drawing.Point(48, 192)
        Me.TextBox_P7X.Name = "TextBox_P7X"
        Me.TextBox_P7X.ReadOnly = True
        Me.TextBox_P7X.Size = New System.Drawing.Size(56, 20)
        Me.TextBox_P7X.TabIndex = 45
        Me.TextBox_P7X.Text = "0"
        Me.TextBox_P7X.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label22
        '
        Me.Label22.Location = New System.Drawing.Point(8, 192)
        Me.Label22.Name = "Label22"
        Me.Label22.Size = New System.Drawing.Size(32, 16)
        Me.Label22.TabIndex = 44
        Me.Label22.Text = "Pt7"
        Me.Label22.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label23
        '
        Me.Label23.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label23.Location = New System.Drawing.Point(1016, 16)
        Me.Label23.Name = "Label23"
        Me.Label23.Size = New System.Drawing.Size(64, 24)
        Me.Label23.TabIndex = 63
        Me.Label23.Text = "Retract Delay"
        Me.Label23.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'TextBox_ApproachHeight
        '
        Me.TextBox_ApproachHeight.Location = New System.Drawing.Point(1080, 48)
        Me.TextBox_ApproachHeight.Name = "TextBox_ApproachHeight"
        Me.TextBox_ApproachHeight.Size = New System.Drawing.Size(56, 20)
        Me.TextBox_ApproachHeight.TabIndex = 62
        Me.TextBox_ApproachHeight.Text = ""
        '
        'TextBox_RetractDelay
        '
        Me.TextBox_RetractDelay.Location = New System.Drawing.Point(1016, 48)
        Me.TextBox_RetractDelay.Name = "TextBox_RetractDelay"
        Me.TextBox_RetractDelay.Size = New System.Drawing.Size(56, 20)
        Me.TextBox_RetractDelay.TabIndex = 61
        Me.TextBox_RetractDelay.Text = ""
        '
        'Label24
        '
        Me.Label24.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label24.Location = New System.Drawing.Point(960, 16)
        Me.Label24.Name = "Label24"
        Me.Label24.Size = New System.Drawing.Size(40, 24)
        Me.Label24.TabIndex = 60
        Me.Label24.Text = "Travel Delay"
        Me.Label24.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'TextBox_TravelDelay
        '
        Me.TextBox_TravelDelay.Location = New System.Drawing.Point(952, 48)
        Me.TextBox_TravelDelay.Name = "TextBox_TravelDelay"
        Me.TextBox_TravelDelay.Size = New System.Drawing.Size(56, 20)
        Me.TextBox_TravelDelay.TabIndex = 59
        Me.TextBox_TravelDelay.Text = ""
        '
        'Label25
        '
        Me.Label25.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label25.Location = New System.Drawing.Point(896, 16)
        Me.Label25.Name = "Label25"
        Me.Label25.Size = New System.Drawing.Size(48, 24)
        Me.Label25.TabIndex = 58
        Me.Label25.Text = "Duration"
        Me.Label25.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label26
        '
        Me.Label26.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label26.Location = New System.Drawing.Point(1080, 16)
        Me.Label26.Name = "Label26"
        Me.Label26.Size = New System.Drawing.Size(56, 24)
        Me.Label26.TabIndex = 57
        Me.Label26.Text = "Approch Ht"
        Me.Label26.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'TextBox_Duration
        '
        Me.TextBox_Duration.Location = New System.Drawing.Point(888, 48)
        Me.TextBox_Duration.Name = "TextBox_Duration"
        Me.TextBox_Duration.Size = New System.Drawing.Size(56, 20)
        Me.TextBox_Duration.TabIndex = 56
        Me.TextBox_Duration.Text = ""
        '
        'Label27
        '
        Me.Label27.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label27.ForeColor = System.Drawing.Color.FromArgb(CType(192, Byte), CType(0, Byte), CType(0, Byte))
        Me.Label27.Location = New System.Drawing.Point(112, 16)
        Me.Label27.Name = "Label27"
        Me.Label27.Size = New System.Drawing.Size(120, 16)
        Me.Label27.TabIndex = 64
        Me.Label27.Text = "Select array type"
        '
        'Label28
        '
        Me.Label28.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label28.Location = New System.Drawing.Point(888, 104)
        Me.Label28.Name = "Label28"
        Me.Label28.Size = New System.Drawing.Size(56, 16)
        Me.Label28.TabIndex = 80
        Me.Label28.Text = "Fill Height"
        Me.Label28.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'TextBox_NoOfRun
        '
        Me.TextBox_NoOfRun.Location = New System.Drawing.Point(952, 128)
        Me.TextBox_NoOfRun.Name = "TextBox_NoOfRun"
        Me.TextBox_NoOfRun.Size = New System.Drawing.Size(56, 20)
        Me.TextBox_NoOfRun.TabIndex = 79
        Me.TextBox_NoOfRun.Text = ""
        '
        'TextBox_FillHeight
        '
        Me.TextBox_FillHeight.Location = New System.Drawing.Point(888, 128)
        Me.TextBox_FillHeight.Name = "TextBox_FillHeight"
        Me.TextBox_FillHeight.Size = New System.Drawing.Size(56, 20)
        Me.TextBox_FillHeight.TabIndex = 78
        Me.TextBox_FillHeight.Text = ""
        '
        'Label29
        '
        Me.Label29.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label29.Location = New System.Drawing.Point(832, 104)
        Me.Label29.Name = "Label29"
        Me.Label29.Size = New System.Drawing.Size(40, 16)
        Me.Label29.TabIndex = 77
        Me.Label29.Text = "Pitch"
        Me.Label29.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'TextBox_Pitch
        '
        Me.TextBox_Pitch.Location = New System.Drawing.Point(824, 128)
        Me.TextBox_Pitch.Name = "TextBox_Pitch"
        Me.TextBox_Pitch.Size = New System.Drawing.Size(56, 20)
        Me.TextBox_Pitch.TabIndex = 76
        Me.TextBox_Pitch.Text = ""
        '
        'Label30
        '
        Me.Label30.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label30.Location = New System.Drawing.Point(752, 100)
        Me.Label30.Name = "Label30"
        Me.Label30.Size = New System.Drawing.Size(72, 24)
        Me.Label30.TabIndex = 75
        Me.Label30.Text = "Arc Radius"
        Me.Label30.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label31
        '
        Me.Label31.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label31.Location = New System.Drawing.Point(952, 104)
        Me.Label31.Name = "Label31"
        Me.Label31.Size = New System.Drawing.Size(64, 16)
        Me.Label31.TabIndex = 74
        Me.Label31.Text = "No Of Run"
        Me.Label31.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'TextBox_ArcRadius
        '
        Me.TextBox_ArcRadius.Location = New System.Drawing.Point(760, 128)
        Me.TextBox_ArcRadius.Name = "TextBox_ArcRadius"
        Me.TextBox_ArcRadius.Size = New System.Drawing.Size(56, 20)
        Me.TextBox_ArcRadius.TabIndex = 73
        Me.TextBox_ArcRadius.Text = ""
        '
        'Label32
        '
        Me.Label32.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label32.Location = New System.Drawing.Point(680, 100)
        Me.Label32.Name = "Label32"
        Me.Label32.Size = New System.Drawing.Size(72, 24)
        Me.Label32.TabIndex = 72
        Me.Label32.Text = "DTaiDist"
        Me.Label32.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'TextBox_DTailDist
        '
        Me.TextBox_DTailDist.Location = New System.Drawing.Point(688, 128)
        Me.TextBox_DTailDist.Name = "TextBox_DTailDist"
        Me.TextBox_DTailDist.Size = New System.Drawing.Size(64, 20)
        Me.TextBox_DTailDist.TabIndex = 71
        Me.TextBox_DTailDist.Text = ""
        '
        'TextBox_ClearanceHt
        '
        Me.TextBox_ClearanceHt.Location = New System.Drawing.Point(616, 128)
        Me.TextBox_ClearanceHt.Name = "TextBox_ClearanceHt"
        Me.TextBox_ClearanceHt.Size = New System.Drawing.Size(64, 20)
        Me.TextBox_ClearanceHt.TabIndex = 70
        Me.TextBox_ClearanceHt.Text = ""
        '
        'Label33
        '
        Me.Label33.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label33.Location = New System.Drawing.Point(608, 100)
        Me.Label33.Name = "Label33"
        Me.Label33.Size = New System.Drawing.Size(80, 24)
        Me.Label33.TabIndex = 69
        Me.Label33.Text = "Clearance Ht"
        Me.Label33.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'TextBox_RetractHeight
        '
        Me.TextBox_RetractHeight.Location = New System.Drawing.Point(1208, 48)
        Me.TextBox_RetractHeight.Name = "TextBox_RetractHeight"
        Me.TextBox_RetractHeight.Size = New System.Drawing.Size(56, 20)
        Me.TextBox_RetractHeight.TabIndex = 68
        Me.TextBox_RetractHeight.Text = ""
        '
        'Label34
        '
        Me.Label34.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label34.Location = New System.Drawing.Point(1216, 16)
        Me.Label34.Name = "Label34"
        Me.Label34.Size = New System.Drawing.Size(48, 24)
        Me.Label34.TabIndex = 67
        Me.Label34.Text = "Retract Ht"
        Me.Label34.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label35
        '
        Me.Label35.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label35.Location = New System.Drawing.Point(1136, 16)
        Me.Label35.Name = "Label35"
        Me.Label35.Size = New System.Drawing.Size(64, 24)
        Me.Label35.TabIndex = 66
        Me.Label35.Text = "Retract Speed"
        Me.Label35.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'TextBox_RetractSpeed
        '
        Me.TextBox_RetractSpeed.Location = New System.Drawing.Point(1144, 48)
        Me.TextBox_RetractSpeed.Name = "TextBox_RetractSpeed"
        Me.TextBox_RetractSpeed.Size = New System.Drawing.Size(56, 20)
        Me.TextBox_RetractSpeed.TabIndex = 65
        Me.TextBox_RetractSpeed.Text = ""
        '
        'TextBox_IONumber
        '
        Me.TextBox_IONumber.Location = New System.Drawing.Point(1208, 128)
        Me.TextBox_IONumber.Name = "TextBox_IONumber"
        Me.TextBox_IONumber.Size = New System.Drawing.Size(56, 20)
        Me.TextBox_IONumber.TabIndex = 88
        Me.TextBox_IONumber.Text = ""
        '
        'Label36
        '
        Me.Label36.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label36.Location = New System.Drawing.Point(1200, 104)
        Me.Label36.Name = "Label36"
        Me.Label36.Size = New System.Drawing.Size(64, 16)
        Me.Label36.TabIndex = 87
        Me.Label36.Text = "IO Number"
        Me.Label36.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label37
        '
        Me.Label37.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label37.Location = New System.Drawing.Point(1136, 104)
        Me.Label37.Name = "Label37"
        Me.Label37.Size = New System.Drawing.Size(64, 16)
        Me.Label37.TabIndex = 86
        Me.Label37.Text = "Edge Clear"
        Me.Label37.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'TextBox_EdgeClear
        '
        Me.TextBox_EdgeClear.Location = New System.Drawing.Point(1144, 128)
        Me.TextBox_EdgeClear.Name = "TextBox_EdgeClear"
        Me.TextBox_EdgeClear.Size = New System.Drawing.Size(56, 20)
        Me.TextBox_EdgeClear.TabIndex = 85
        Me.TextBox_EdgeClear.Text = ""
        '
        'Label38
        '
        Me.Label38.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label38.Location = New System.Drawing.Point(1016, 104)
        Me.Label38.Name = "Label38"
        Me.Label38.Size = New System.Drawing.Size(56, 16)
        Me.Label38.TabIndex = 84
        Me.Label38.Text = "Sprial"
        Me.Label38.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'TextBox_RtAngle
        '
        Me.TextBox_RtAngle.Location = New System.Drawing.Point(1080, 128)
        Me.TextBox_RtAngle.Name = "TextBox_RtAngle"
        Me.TextBox_RtAngle.Size = New System.Drawing.Size(56, 20)
        Me.TextBox_RtAngle.TabIndex = 83
        Me.TextBox_RtAngle.Text = ""
        '
        'TextBox_Sprial
        '
        Me.TextBox_Sprial.Location = New System.Drawing.Point(1016, 128)
        Me.TextBox_Sprial.Name = "TextBox_Sprial"
        Me.TextBox_Sprial.Size = New System.Drawing.Size(56, 20)
        Me.TextBox_Sprial.TabIndex = 82
        Me.TextBox_Sprial.Text = ""
        '
        'Label39
        '
        Me.Label39.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label39.Location = New System.Drawing.Point(1088, 104)
        Me.Label39.Name = "Label39"
        Me.Label39.Size = New System.Drawing.Size(48, 16)
        Me.Label39.TabIndex = 81
        Me.Label39.Text = "Rt Angle"
        Me.Label39.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'RadioButton_P1
        '
        Me.RadioButton_P1.Checked = True
        Me.RadioButton_P1.Location = New System.Drawing.Point(32, 128)
        Me.RadioButton_P1.Name = "RadioButton_P1"
        Me.RadioButton_P1.Size = New System.Drawing.Size(16, 24)
        Me.RadioButton_P1.TabIndex = 89
        Me.RadioButton_P1.TabStop = True
        '
        'RadioButton_P2
        '
        Me.RadioButton_P2.Location = New System.Drawing.Point(280, 128)
        Me.RadioButton_P2.Name = "RadioButton_P2"
        Me.RadioButton_P2.Size = New System.Drawing.Size(16, 24)
        Me.RadioButton_P2.TabIndex = 90
        '
        'RadioButton_P3
        '
        Me.RadioButton_P3.Location = New System.Drawing.Point(464, 128)
        Me.RadioButton_P3.Name = "RadioButton_P3"
        Me.RadioButton_P3.Size = New System.Drawing.Size(16, 24)
        Me.RadioButton_P3.TabIndex = 91
        '
        'RadioButton_P4
        '
        Me.RadioButton_P4.Location = New System.Drawing.Point(32, 160)
        Me.RadioButton_P4.Name = "RadioButton_P4"
        Me.RadioButton_P4.Size = New System.Drawing.Size(16, 24)
        Me.RadioButton_P4.TabIndex = 92
        '
        'RadioButton_P7
        '
        Me.RadioButton_P7.Location = New System.Drawing.Point(32, 192)
        Me.RadioButton_P7.Name = "RadioButton_P7"
        Me.RadioButton_P7.Size = New System.Drawing.Size(16, 24)
        Me.RadioButton_P7.TabIndex = 95
        '
        'ArrayGenerate
        '
        Me.AcceptButton = Me.Button_OnOK
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.CancelButton = Me.Button_OnCancel
        Me.ClientSize = New System.Drawing.Size(1285, 262)
        Me.ControlBox = False
        Me.Controls.Add(Me.RadioButton_P7)
        Me.Controls.Add(Me.RadioButton_P4)
        Me.Controls.Add(Me.RadioButton_P3)
        Me.Controls.Add(Me.RadioButton_P2)
        Me.Controls.Add(Me.RadioButton_P1)
        Me.Controls.Add(Me.TextBox_IONumber)
        Me.Controls.Add(Me.TextBox_EdgeClear)
        Me.Controls.Add(Me.TextBox_RtAngle)
        Me.Controls.Add(Me.TextBox_Sprial)
        Me.Controls.Add(Me.TextBox_NoOfRun)
        Me.Controls.Add(Me.TextBox_FillHeight)
        Me.Controls.Add(Me.TextBox_Pitch)
        Me.Controls.Add(Me.TextBox_ArcRadius)
        Me.Controls.Add(Me.TextBox_DTailDist)
        Me.Controls.Add(Me.TextBox_ClearanceHt)
        Me.Controls.Add(Me.TextBox_RetractHeight)
        Me.Controls.Add(Me.TextBox_RetractSpeed)
        Me.Controls.Add(Me.TextBox_ApproachHeight)
        Me.Controls.Add(Me.TextBox_RetractDelay)
        Me.Controls.Add(Me.TextBox_TravelDelay)
        Me.Controls.Add(Me.TextBox_Duration)
        Me.Controls.Add(Me.TextBox_P7Y)
        Me.Controls.Add(Me.TextBox_P7X)
        Me.Controls.Add(Me.TextBox_P4Y)
        Me.Controls.Add(Me.TextBox_P4X)
        Me.Controls.Add(Me.TextBox_P3Y)
        Me.Controls.Add(Me.TextBox_P3X)
        Me.Controls.Add(Me.TextBox_P2Y)
        Me.Controls.Add(Me.TextBox_P2X)
        Me.Controls.Add(Me.TextBox_P1Z)
        Me.Controls.Add(Me.TextBox_P1Y)
        Me.Controls.Add(Me.TextBox_P1X)
        Me.Controls.Add(Me.TextBox_NeedleGap)
        Me.Controls.Add(Me.TextBox_TravelSpeed)
        Me.Controls.Add(Me.TextBox_Dispense)
        Me.Controls.Add(Me.TextBox_Needle)
        Me.Controls.Add(Me.Label36)
        Me.Controls.Add(Me.Label37)
        Me.Controls.Add(Me.Label38)
        Me.Controls.Add(Me.Label39)
        Me.Controls.Add(Me.Label28)
        Me.Controls.Add(Me.Label29)
        Me.Controls.Add(Me.Label30)
        Me.Controls.Add(Me.Label31)
        Me.Controls.Add(Me.Label32)
        Me.Controls.Add(Me.Label33)
        Me.Controls.Add(Me.Label34)
        Me.Controls.Add(Me.Label35)
        Me.Controls.Add(Me.Label27)
        Me.Controls.Add(Me.Label23)
        Me.Controls.Add(Me.Label24)
        Me.Controls.Add(Me.Label25)
        Me.Controls.Add(Me.Label26)
        Me.Controls.Add(Me.Label22)
        Me.Controls.Add(Me.Label19)
        Me.Controls.Add(Me.Label14)
        Me.Controls.Add(Me.Label16)
        Me.Controls.Add(Me.Label11)
        Me.Controls.Add(Me.Label13)
        Me.Controls.Add(Me.Label10)
        Me.Controls.Add(Me.Label9)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.lbDispenseMode)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.CombElementType)
        Me.Controls.Add(Me.Button_OnCancel)
        Me.Controls.Add(Me.Button_OnOK)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Location = New System.Drawing.Point(0, 78)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "ArrayGenerate"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Array Generation"
        Me.TopMost = True
        Me.ResumeLayout(False)

    End Sub

#End Region

    Public Sub SetDefaultPara(ByVal Para As CIDSData.CIDSPatternExecution.CIDSTemplate)
        m_CurrentPara = Para
    End Sub

    Public Sub SetPoint(ByVal PointX As Double, ByVal PointY As Double, ByVal PointZ As Double)
        If RadioButton_P1.Checked Then
            TextBox_P1X.Text = CStr(PointX)
            TextBox_P1Y.Text = CStr(PointY)
            TextBox_P1Z.Text = CStr(PointZ)
        ElseIf RadioButton_P2.Checked Then
            TextBox_P2X.Text = CStr(PointX)
            TextBox_P2Y.Text = CStr(PointY)
        ElseIf RadioButton_P3.Checked Then
            TextBox_P3X.Text = CStr(PointX)
            TextBox_P3Y.Text = CStr(PointY)
        ElseIf RadioButton_P4.Checked Then
            TextBox_P4X.Text = CStr(PointX)
            TextBox_P4Y.Text = CStr(PointY)
        ElseIf RadioButton_P7.Checked Then
            TextBox_P7X.Text = CStr(PointX)
            TextBox_P7Y.Text = CStr(PointY)
        End If
        TraceGCCollect()
    End Sub

    Private Sub CombElementType_SelectedTextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CombElementType.TextChanged
        Dim typeSelected As String = CombElementType.Text
        RadioButton_P1.Enabled = True
        ArrayButtonEnable(typeSelected)
        ArrayTextBoxEnable(typeSelected)

        If FormIni = True Then
            ArrayTextBoxFillIn(typeSelected)
            Button_OnOK.Enabled = True
            Button_OnCancel.Enabled = True
        End If
        FormIni = True

        TraceGCCollect()
    End Sub

    Private Sub ArrayButtonEnable(ByVal ElementType As String)

        Select Case ElementType
            Case "Dot"
                RadioButton_P1.Enabled = True
                RadioButton_P2.Enabled = False
                RadioButton_P3.Enabled = False
                RadioButton_P4.Enabled = True
                RadioButton_P7.Enabled = True

            Case "Line"
                RadioButton_P1.Enabled = True
                RadioButton_P2.Enabled = True
                RadioButton_P3.Enabled = False
                RadioButton_P4.Enabled = True
                RadioButton_P7.Enabled = True

            Case "Arc"
                RadioButton_P1.Enabled = True
                RadioButton_P2.Enabled = True
                RadioButton_P3.Enabled = True
                RadioButton_P4.Enabled = True
                RadioButton_P7.Enabled = True
            Case "Circle"
                RadioButton_P1.Enabled = True
                RadioButton_P2.Enabled = True
                RadioButton_P3.Enabled = True
                RadioButton_P4.Enabled = True
                RadioButton_P7.Enabled = True

            Case "FillCircle"
                RadioButton_P1.Enabled = True
                RadioButton_P2.Enabled = True
                RadioButton_P3.Enabled = True
                RadioButton_P4.Enabled = True
                RadioButton_P7.Enabled = True


            Case "Rectangle"
                RadioButton_P1.Enabled = True
                RadioButton_P2.Enabled = True
                RadioButton_P3.Enabled = True
                RadioButton_P4.Enabled = True
                RadioButton_P7.Enabled = True

            Case "FillRectangle"
                RadioButton_P1.Enabled = True
                RadioButton_P2.Enabled = True
                RadioButton_P3.Enabled = True
                RadioButton_P4.Enabled = True
                RadioButton_P7.Enabled = True

            Case "SubPattern"
                RadioButton_P1.Enabled = True
                RadioButton_P2.Enabled = False
                RadioButton_P3.Enabled = False
                RadioButton_P4.Enabled = True
                RadioButton_P7.Enabled = True

        End Select

        TraceGCCollect()
    End Sub

    Private Sub ArrayTextBoxEnable(ByVal ElementType As String)
        Select Case ElementType
            Case "Dot"
                TextBox_Needle.Enabled = True
                TextBox_Dispense.Enabled = True
                TextBox_TravelSpeed.Enabled = False
                TextBox_NeedleGap.Enabled = True
                TextBox_Duration.Enabled = True
                TextBox_TravelDelay.Enabled = False
                TextBox_RetractDelay.Enabled = True
                TextBox_ApproachHeight.Enabled = True
                TextBox_RetractSpeed.Enabled = True
                TextBox_RetractHeight.Enabled = True
                TextBox_ClearanceHt.Enabled = True
                TextBox_DTailDist.Enabled = False
                TextBox_ArcRadius.Enabled = True
                TextBox_Pitch.Enabled = False
                TextBox_FillHeight.Enabled = False
                TextBox_NoOfRun.Enabled = False
                TextBox_Sprial.Enabled = False
                TextBox_RtAngle.Enabled = False
                TextBox_EdgeClear.Enabled = False
                TextBox_IONumber.Enabled = False


            Case "Line"
                TextBox_Needle.Enabled = True
                TextBox_Dispense.Enabled = True
                TextBox_TravelSpeed.Enabled = True
                TextBox_NeedleGap.Enabled = True
                TextBox_Duration.Enabled = False
                TextBox_TravelDelay.Enabled = True
                TextBox_RetractDelay.Enabled = True
                TextBox_ApproachHeight.Enabled = True
                TextBox_RetractSpeed.Enabled = True
                TextBox_RetractHeight.Enabled = True
                TextBox_ClearanceHt.Enabled = True
                TextBox_DTailDist.Enabled = True
                TextBox_ArcRadius.Enabled = True
                TextBox_Pitch.Enabled = False
                TextBox_FillHeight.Enabled = False
                TextBox_NoOfRun.Enabled = False
                TextBox_Sprial.Enabled = False
                TextBox_RtAngle.Enabled = False
                TextBox_EdgeClear.Enabled = False
                TextBox_IONumber.Enabled = False

            Case "Arc"
                TextBox_Needle.Enabled = True
                TextBox_Dispense.Enabled = True
                TextBox_TravelSpeed.Enabled = True
                TextBox_NeedleGap.Enabled = True
                TextBox_Duration.Enabled = True
                TextBox_TravelDelay.Enabled = True
                TextBox_RetractDelay.Enabled = True
                TextBox_ApproachHeight.Enabled = True
                TextBox_RetractSpeed.Enabled = True
                TextBox_RetractHeight.Enabled = True
                TextBox_ClearanceHt.Enabled = True
                TextBox_DTailDist.Enabled = True
                TextBox_ArcRadius.Enabled = True
                TextBox_Pitch.Enabled = False
                TextBox_FillHeight.Enabled = False
                TextBox_NoOfRun.Enabled = False
                TextBox_Sprial.Enabled = False
                TextBox_RtAngle.Enabled = False
                TextBox_EdgeClear.Enabled = False
                TextBox_IONumber.Enabled = False

            Case "Circle"
                TextBox_Needle.Enabled = True
                TextBox_Dispense.Enabled = True
                TextBox_TravelSpeed.Enabled = True
                TextBox_NeedleGap.Enabled = True
                TextBox_Duration.Enabled = True
                TextBox_TravelDelay.Enabled = True
                TextBox_RetractDelay.Enabled = True
                TextBox_ApproachHeight.Enabled = True
                TextBox_RetractSpeed.Enabled = True
                TextBox_RetractHeight.Enabled = True
                TextBox_ClearanceHt.Enabled = True
                TextBox_DTailDist.Enabled = True
                TextBox_ArcRadius.Enabled = True
                TextBox_Pitch.Enabled = False
                TextBox_FillHeight.Enabled = False
                TextBox_NoOfRun.Enabled = False
                TextBox_Sprial.Enabled = False
                TextBox_RtAngle.Enabled = False
                TextBox_EdgeClear.Enabled = False
                TextBox_IONumber.Enabled = False

            Case "FillCircle"
                TextBox_Needle.Enabled = True
                TextBox_Dispense.Enabled = True
                TextBox_TravelSpeed.Enabled = True
                TextBox_NeedleGap.Enabled = True
                TextBox_Duration.Enabled = True
                TextBox_TravelDelay.Enabled = True
                TextBox_RetractDelay.Enabled = True
                TextBox_ApproachHeight.Enabled = True
                TextBox_RetractSpeed.Enabled = True
                TextBox_RetractHeight.Enabled = True
                TextBox_ClearanceHt.Enabled = True
                TextBox_DTailDist.Enabled = True
                TextBox_ArcRadius.Enabled = True
                TextBox_Pitch.Enabled = True
                TextBox_FillHeight.Enabled = True
                TextBox_NoOfRun.Enabled = True
                TextBox_Sprial.Enabled = True
                TextBox_RtAngle.Enabled = False
                TextBox_EdgeClear.Enabled = False
                TextBox_IONumber.Enabled = False

            Case "Rectangle"
                TextBox_Needle.Enabled = True
                TextBox_Dispense.Enabled = True
                TextBox_TravelSpeed.Enabled = True
                TextBox_NeedleGap.Enabled = True
                TextBox_Duration.Enabled = True
                TextBox_TravelDelay.Enabled = True
                TextBox_RetractDelay.Enabled = True
                TextBox_ApproachHeight.Enabled = True
                TextBox_RetractSpeed.Enabled = True
                TextBox_RetractHeight.Enabled = True
                TextBox_ClearanceHt.Enabled = True
                TextBox_DTailDist.Enabled = True
                TextBox_ArcRadius.Enabled = True
                TextBox_Pitch.Enabled = False
                TextBox_FillHeight.Enabled = False
                TextBox_NoOfRun.Enabled = False
                TextBox_Sprial.Enabled = False
                TextBox_RtAngle.Enabled = False
                TextBox_EdgeClear.Enabled = False
                TextBox_IONumber.Enabled = False

            Case "FillRectangle"
                TextBox_Needle.Enabled = True
                TextBox_Dispense.Enabled = True
                TextBox_TravelSpeed.Enabled = True
                TextBox_NeedleGap.Enabled = True
                TextBox_Duration.Enabled = True
                TextBox_TravelDelay.Enabled = True
                TextBox_RetractDelay.Enabled = True
                TextBox_ApproachHeight.Enabled = True
                TextBox_RetractSpeed.Enabled = True
                TextBox_RetractHeight.Enabled = True
                TextBox_ClearanceHt.Enabled = True
                TextBox_DTailDist.Enabled = True
                TextBox_ArcRadius.Enabled = True
                TextBox_Pitch.Enabled = True
                TextBox_FillHeight.Enabled = True
                TextBox_NoOfRun.Enabled = True
                TextBox_Sprial.Enabled = True
                TextBox_RtAngle.Enabled = False
                TextBox_EdgeClear.Enabled = False
                TextBox_IONumber.Enabled = False

            Case "SubPattern"
                TextBox_Needle.Enabled = False
                TextBox_Dispense.Enabled = False
                TextBox_TravelSpeed.Enabled = False
                TextBox_NeedleGap.Enabled = False
                TextBox_Duration.Enabled = False
                TextBox_TravelDelay.Enabled = False
                TextBox_RetractDelay.Enabled = False
                TextBox_ApproachHeight.Enabled = False
                TextBox_RetractSpeed.Enabled = False
                TextBox_RetractHeight.Enabled = False
                TextBox_ClearanceHt.Enabled = False
                TextBox_DTailDist.Enabled = False
                TextBox_ArcRadius.Enabled = False
                TextBox_Pitch.Enabled = False
                TextBox_FillHeight.Enabled = False
                TextBox_NoOfRun.Enabled = False
                TextBox_Sprial.Enabled = False
                TextBox_RtAngle.Enabled = False
                TextBox_EdgeClear.Enabled = False
                TextBox_IONumber.Enabled = False

        End Select

        TraceGCCollect()
    End Sub

    Private Sub ArrayTextBoxFillIn(ByVal ElementType As String)
        If (CStr(m_CurrentPara.DispenseFlag) = "True" Or CStr(m_CurrentPara.DispenseFlag).ToUpper = "ON") Then
            m_CurrentPara.DispenseFlag = "On"
        ElseIf (CStr(m_CurrentPara.DispenseFlag) = "False") Then
            m_CurrentPara.DispenseFlag = "Off"
        End If
        If Not (ElementType = "SubPattern") Then
            m_CurrentPara.DispenseFlag = "On"
        End If
        lbDispenseMode.Text = "Dispensing Status"
        Select Case ElementType
            Case "Dot"
                TextBox_Needle.Text = CStr(m_CurrentPara.Needle)
                TextBox_Dispense.Text = CStr(m_CurrentPara.DispenseFlag)
                TextBox_TravelSpeed.Text = ""
                TextBox_NeedleGap.Text = CStr(m_CurrentPara.NeedleGap)
                TextBox_Duration.Text = CStr(m_CurrentPara.DispenseDuration)
                TextBox_TravelDelay.Text = ""
                TextBox_RetractDelay.Text = CStr(m_CurrentPara.RetractDelay)
                TextBox_ApproachHeight.Text = CStr(m_CurrentPara.ApproachDispHeight)
                TextBox_RetractSpeed.Text = CStr(m_CurrentPara.RetractSpeed)
                TextBox_RetractHeight.Text = CStr(m_CurrentPara.RetractHeight)
                TextBox_ClearanceHt.Text = CStr(m_CurrentPara.ClearanceHeight)
                TextBox_DTailDist.Text = ""
                TextBox_ArcRadius.Text = CStr(m_CurrentPara.ArcRadius)
                TextBox_Pitch.Text = ""
                TextBox_FillHeight.Text = ""
                TextBox_NoOfRun.Text = ""
                TextBox_Sprial.Text = ""
                TextBox_RtAngle.Text = ""
                TextBox_EdgeClear.Text = ""
                TextBox_IONumber.Text = ""
                TextBox_P2X.Text = ""
                TextBox_P2Y.Text = ""
                TextBox_P3X.Text = ""
                TextBox_P3Y.Text = ""
                TextBox_P4X.Text = "0"
                TextBox_P4Y.Text = "0"
                TextBox_P7X.Text = "0"
                TextBox_P7Y.Text = "0"

            Case "Line"
                TextBox_Needle.Text = CStr(m_CurrentPara.Needle)
                TextBox_Dispense.Text = CStr(m_CurrentPara.DispenseFlag)
                TextBox_TravelSpeed.Text = CStr(m_CurrentPara.TravelSpeed)
                TextBox_NeedleGap.Text = CStr(m_CurrentPara.NeedleGap)
                TextBox_Duration.Text = ""
                TextBox_TravelDelay.Text = CStr(m_CurrentPara.TravelDelay)
                TextBox_RetractDelay.Text = CStr(m_CurrentPara.RetractDelay)
                TextBox_ApproachHeight.Text = CStr(m_CurrentPara.ApproachDispHeight)
                TextBox_RetractSpeed.Text = CStr(m_CurrentPara.RetractSpeed)
                TextBox_RetractHeight.Text = CStr(m_CurrentPara.RetractHeight)
                TextBox_ClearanceHt.Text = CStr(m_CurrentPara.ClearanceHeight)
                TextBox_DTailDist.Text = CStr(m_CurrentPara.DetailingDistance)
                TextBox_ArcRadius.Text = CStr(m_CurrentPara.ArcRadius)
                TextBox_Pitch.Text = ""
                TextBox_FillHeight.Text = ""
                TextBox_NoOfRun.Text = ""
                TextBox_Sprial.Text = ""
                TextBox_RtAngle.Text = ""
                TextBox_EdgeClear.Text = ""
                TextBox_IONumber.Text = ""
                TextBox_P3X.Text = ""
                TextBox_P3Y.Text = ""
                TextBox_P4X.Text = "0"
                TextBox_P4Y.Text = "0"
                TextBox_P7X.Text = "0"
                TextBox_P7Y.Text = "0"

            Case "Arc"
                TextBox_Needle.Text = CStr(m_CurrentPara.Needle)
                TextBox_Dispense.Text = CStr(m_CurrentPara.DispenseFlag)
                TextBox_TravelSpeed.Text = CStr(m_CurrentPara.TravelSpeed)
                TextBox_NeedleGap.Text = CStr(m_CurrentPara.NeedleGap)
                TextBox_Duration.Text = ""
                TextBox_TravelDelay.Text = CStr(m_CurrentPara.TravelDelay)
                TextBox_RetractDelay.Text = CStr(m_CurrentPara.RetractDelay)
                TextBox_ApproachHeight.Text = CStr(m_CurrentPara.ApproachDispHeight)
                TextBox_RetractSpeed.Text = CStr(m_CurrentPara.RetractSpeed)
                TextBox_RetractHeight.Text = CStr(m_CurrentPara.RetractHeight)
                TextBox_ClearanceHt.Text = CStr(m_CurrentPara.ClearanceHeight)
                TextBox_DTailDist.Text = CStr(m_CurrentPara.DetailingDistance)
                TextBox_ArcRadius.Text = CStr(m_CurrentPara.ArcRadius)
                TextBox_Pitch.Text = ""
                TextBox_FillHeight.Text = ""
                TextBox_NoOfRun.Text = ""
                TextBox_Sprial.Text = ""
                TextBox_RtAngle.Text = ""
                TextBox_EdgeClear.Text = ""
                TextBox_IONumber.Text = ""
                TextBox_P2X.Text = "0"
                TextBox_P2Y.Text = "0"
                TextBox_P3X.Text = "0"
                TextBox_P3Y.Text = "0"
                TextBox_P4X.Text = "0"
                TextBox_P4Y.Text = "0"
                TextBox_P7X.Text = "0"
                TextBox_P7Y.Text = "0"


            Case "Circle"
                TextBox_Needle.Text = CStr(m_CurrentPara.Needle)
                TextBox_Dispense.Text = CStr(m_CurrentPara.DispenseFlag)
                TextBox_TravelSpeed.Text = CStr(m_CurrentPara.TravelSpeed)
                TextBox_NeedleGap.Text = CStr(m_CurrentPara.NeedleGap)
                TextBox_Duration.Text = ""
                TextBox_TravelDelay.Text = CStr(m_CurrentPara.TravelDelay)
                TextBox_RetractDelay.Text = CStr(m_CurrentPara.RetractDelay)
                TextBox_ApproachHeight.Text = CStr(m_CurrentPara.ApproachDispHeight)
                TextBox_RetractSpeed.Text = CStr(m_CurrentPara.RetractSpeed)
                TextBox_RetractHeight.Text = CStr(m_CurrentPara.RetractHeight)
                TextBox_ClearanceHt.Text = CStr(m_CurrentPara.ClearanceHeight)
                TextBox_DTailDist.Text = CStr(m_CurrentPara.DetailingDistance)
                TextBox_ArcRadius.Text = CStr(m_CurrentPara.ArcRadius)
                TextBox_Pitch.Text = ""
                TextBox_FillHeight.Text = ""
                TextBox_NoOfRun.Text = ""
                TextBox_Sprial.Text = ""
                TextBox_RtAngle.Text = ""
                TextBox_EdgeClear.Text = ""
                TextBox_IONumber.Text = ""
                TextBox_P2X.Text = "0"
                TextBox_P2Y.Text = "0"
                TextBox_P3X.Text = "0"
                TextBox_P3Y.Text = "0"
                TextBox_P4X.Text = "0"
                TextBox_P4Y.Text = "0"
                TextBox_P7X.Text = "0"
                TextBox_P7Y.Text = "0"
            Case "FillCircle"
                TextBox_Needle.Text = CStr(m_CurrentPara.Needle)
                TextBox_Dispense.Text = CStr(m_CurrentPara.DispenseFlag)
                TextBox_TravelSpeed.Text = CStr(m_CurrentPara.TravelSpeed)
                TextBox_NeedleGap.Text = CStr(m_CurrentPara.NeedleGap)
                TextBox_Duration.Text = ""
                TextBox_TravelDelay.Text = CStr(m_CurrentPara.TravelDelay)
                TextBox_RetractDelay.Text = CStr(m_CurrentPara.RetractDelay)
                TextBox_ApproachHeight.Text = CStr(m_CurrentPara.ApproachDispHeight)
                TextBox_RetractSpeed.Text = CStr(m_CurrentPara.RetractSpeed)
                TextBox_RetractHeight.Text = CStr(m_CurrentPara.RetractHeight)
                TextBox_ClearanceHt.Text = CStr(m_CurrentPara.ClearanceHeight)
                TextBox_DTailDist.Text = CStr(m_CurrentPara.DetailingDistance)
                TextBox_ArcRadius.Text = CStr(m_CurrentPara.ArcRadius)
                TextBox_Pitch.Text = CStr(m_CurrentPara.Pitch)
                TextBox_FillHeight.Text = CStr(m_CurrentPara.FilledHeight)
                TextBox_NoOfRun.Text = CStr(m_CurrentPara.NumofRun)
                TextBox_Sprial.Text = CStr(m_CurrentPara.SpiralFlag)
                TextBox_RtAngle.Text = ""
                TextBox_EdgeClear.Text = ""
                TextBox_IONumber.Text = ""
                TextBox_P2X.Text = "0"
                TextBox_P2Y.Text = "0"
                TextBox_P3X.Text = "0"
                TextBox_P3Y.Text = "0"
                TextBox_P4X.Text = "0"
                TextBox_P4Y.Text = "0"
                TextBox_P7X.Text = "0"
                TextBox_P7Y.Text = "0"

            Case "Rectangle"
                TextBox_Needle.Text = CStr(m_CurrentPara.Needle)
                TextBox_Dispense.Text = CStr(m_CurrentPara.DispenseFlag)
                TextBox_TravelSpeed.Text = CStr(m_CurrentPara.TravelSpeed)
                TextBox_NeedleGap.Text = CStr(m_CurrentPara.NeedleGap)
                TextBox_Duration.Text = ""
                TextBox_TravelDelay.Text = CStr(m_CurrentPara.TravelDelay)
                TextBox_RetractDelay.Text = CStr(m_CurrentPara.RetractDelay)
                TextBox_ApproachHeight.Text = CStr(m_CurrentPara.ApproachDispHeight)
                TextBox_RetractSpeed.Text = CStr(m_CurrentPara.RetractSpeed)
                TextBox_RetractHeight.Text = CStr(m_CurrentPara.RetractHeight)
                TextBox_ClearanceHt.Text = CStr(m_CurrentPara.ClearanceHeight)
                TextBox_DTailDist.Text = CStr(m_CurrentPara.DetailingDistance)
                TextBox_ArcRadius.Text = CStr(m_CurrentPara.ArcRadius)
                TextBox_Pitch.Text = ""
                TextBox_FillHeight.Text = ""
                TextBox_NoOfRun.Text = ""
                TextBox_Sprial.Text = ""
                TextBox_RtAngle.Text = ""
                TextBox_EdgeClear.Text = ""
                TextBox_IONumber.Text = ""
                TextBox_P2X.Text = "0"
                TextBox_P2Y.Text = "0"
                TextBox_P3X.Text = "0"
                TextBox_P3Y.Text = "0"
                TextBox_P4X.Text = "0"
                TextBox_P4Y.Text = "0"
                TextBox_P7X.Text = "0"
                TextBox_P7Y.Text = "0"


            Case "FillRectangle"
                TextBox_Needle.Text = CStr(m_CurrentPara.Needle)
                TextBox_Dispense.Text = CStr(m_CurrentPara.DispenseFlag)
                TextBox_TravelSpeed.Text = CStr(m_CurrentPara.TravelSpeed)
                TextBox_NeedleGap.Text = CStr(m_CurrentPara.NeedleGap)
                TextBox_Duration.Text = ""
                TextBox_TravelDelay.Text = CStr(m_CurrentPara.TravelDelay)
                TextBox_RetractDelay.Text = CStr(m_CurrentPara.RetractDelay)
                TextBox_ApproachHeight.Text = CStr(m_CurrentPara.ApproachDispHeight)
                TextBox_RetractSpeed.Text = CStr(m_CurrentPara.RetractSpeed)
                TextBox_RetractHeight.Text = CStr(m_CurrentPara.RetractHeight)
                TextBox_ClearanceHt.Text = CStr(m_CurrentPara.ClearanceHeight)
                TextBox_DTailDist.Text = CStr(m_CurrentPara.DetailingDistance)
                TextBox_ArcRadius.Text = CStr(m_CurrentPara.ArcRadius)
                TextBox_Pitch.Text = CStr(m_CurrentPara.Pitch)
                TextBox_FillHeight.Text = CStr(m_CurrentPara.FilledHeight)
                TextBox_NoOfRun.Text = CStr(m_CurrentPara.NumofRun)
                TextBox_Sprial.Text = CStr(m_CurrentPara.SpiralFlag)
                TextBox_RtAngle.Text = ""
                TextBox_EdgeClear.Text = ""
                TextBox_IONumber.Text = ""
                TextBox_P2X.Text = "0"
                TextBox_P2Y.Text = "0"
                TextBox_P3X.Text = "0"
                TextBox_P3Y.Text = "0"
                TextBox_P4X.Text = "0"
                TextBox_P4Y.Text = "0"
                TextBox_P7X.Text = "0"
                TextBox_P7Y.Text = "0"

            Case "SubPattern"
                TextBox_Needle.Text = ""
                lbDispenseMode.Text = "Subpattern file path"
                Dim DialogPreview As New FormSelectPatternFile
                DialogPreview.Location.Offset(0, 200)

                DialogPreview.ShowDialog()
                Dim file As String = DialogPreview.Path
                If Nothing = file Then
                    TextBox_Dispense.Text = ""
                Else
                    TextBox_Dispense.Text = file
                End If

                TextBox_TravelSpeed.Text = ""
                TextBox_NeedleGap.Text = ""
                TextBox_Duration.Text = ""
                TextBox_TravelDelay.Text = ""
                TextBox_RetractDelay.Text = ""
                TextBox_ApproachHeight.Text = ""
                TextBox_RetractSpeed.Text = ""
                TextBox_RetractHeight.Text = ""
                TextBox_ClearanceHt.Text = ""
                TextBox_DTailDist.Text = ""
                TextBox_ArcRadius.Text = ""
                TextBox_Pitch.Text = ""
                TextBox_FillHeight.Text = ""
                TextBox_NoOfRun.Text = ""
                TextBox_Sprial.Text = ""
                TextBox_RtAngle.Text = ""
                TextBox_EdgeClear.Text = ""
                TextBox_IONumber.Text = ""
                TextBox_P2X.Text = ""
                TextBox_P2Y.Text = ""
                TextBox_P3X.Text = ""
                TextBox_P3Y.Text = ""
                TextBox_P4X.Text = "0"
                TextBox_P4Y.Text = "0"
                TextBox_P7X.Text = "0"
                TextBox_P7Y.Text = "0"

        End Select

        TraceGCCollect()
    End Sub

    Private Sub ArrayTextBoxStore(ByVal ElementType As String)
        'Preserve data for external use

        ArrayPara1.Name = ElementType
        Select Case ElementType
            Case "Dot"
                ArrayPara1.Needle = TextBox_Needle.Text
                ArrayPara1.DispenseFlag = TextBox_Dispense.Text
                ArrayPara1.NeedleGap = CDbl(TextBox_NeedleGap.Text)
                ArrayPara1.DispenseDuration = CDbl(TextBox_Duration.Text)

                ArrayPara1.RetractDelay = CDbl(TextBox_RetractDelay.Text)
                ArrayPara1.ApproachDispHeight = CDbl(TextBox_ApproachHeight.Text)
                ArrayPara1.RetractSpeed = CDbl(TextBox_RetractSpeed.Text)
                ArrayPara1.RetractHeight = CDbl(TextBox_RetractHeight.Text)
                ArrayPara1.ClearanceHeight = CDbl(TextBox_ClearanceHt.Text)
                ArrayPara1.ArcRadius = CDbl(TextBox_ArcRadius.Text)

                ArrayPara1.Pos1.X = CDbl(TextBox_P1X.Text)
                ArrayPara1.Pos1.Y = CDbl(TextBox_P1Y.Text)
                ArrayPara1.Pos1.Z = CDbl(TextBox_P1Z.Text)

                ArrayPara2.Pos1.X = CDbl(TextBox_P4X.Text)
                ArrayPara2.Pos1.Y = CDbl(TextBox_P4Y.Text)

                ArrayPara3.Pos1.X = CDbl(TextBox_P7X.Text)
                ArrayPara3.Pos1.Y = CDbl(TextBox_P7Y.Text)


            Case "Line"
                ArrayPara1.Needle = TextBox_Needle.Text
                ArrayPara1.DispenseFlag = TextBox_Dispense.Text
                ArrayPara1.TravelSpeed = CDbl(TextBox_TravelSpeed.Text)
                ArrayPara1.NeedleGap = CDbl(TextBox_NeedleGap.Text)
                ArrayPara1.TravelDelay = CDbl(TextBox_TravelDelay.Text)
                ArrayPara1.RetractDelay = CDbl(TextBox_RetractDelay.Text)
                ArrayPara1.ApproachDispHeight = CDbl(TextBox_ApproachHeight.Text)
                ArrayPara1.RetractSpeed = CDbl(TextBox_RetractSpeed.Text)
                ArrayPara1.RetractHeight = CDbl(TextBox_RetractHeight.Text)
                ArrayPara1.ClearanceHeight = CDbl(TextBox_ClearanceHt.Text)
                ArrayPara1.DetailingDistance = CDbl(TextBox_DTailDist.Text)
                ArrayPara1.ArcRadius = CDbl(TextBox_ArcRadius.Text)

                ArrayPara1.Pos1.X = CDbl(TextBox_P1X.Text)
                ArrayPara1.Pos1.Y = CDbl(TextBox_P1Y.Text)
                ArrayPara1.Pos1.Z = CDbl(TextBox_P1Z.Text)
                ArrayPara1.Pos2.X = CDbl(TextBox_P2X.Text)
                ArrayPara1.Pos2.Y = CDbl(TextBox_P2Y.Text)

                ArrayPara2.Pos1.X = CDbl(TextBox_P4X.Text)
                ArrayPara2.Pos1.Y = CDbl(TextBox_P4Y.Text)

                ArrayPara3.Pos1.X = CDbl(TextBox_P7X.Text)
                ArrayPara3.Pos1.Y = CDbl(TextBox_P7Y.Text)


            Case "Arc"
                ArrayPara1.Needle = TextBox_Needle.Text
                ArrayPara1.DispenseFlag = TextBox_Dispense.Text
                ArrayPara1.TravelSpeed = CDbl(TextBox_TravelSpeed.Text)
                ArrayPara1.NeedleGap = CDbl(TextBox_NeedleGap.Text)
                ArrayPara1.TravelDelay = CDbl(TextBox_TravelDelay.Text)
                ArrayPara1.RetractDelay = CDbl(TextBox_RetractDelay.Text)
                ArrayPara1.ApproachDispHeight = CDbl(TextBox_ApproachHeight.Text)
                ArrayPara1.RetractSpeed = CDbl(TextBox_RetractSpeed.Text)
                ArrayPara1.RetractHeight = CDbl(TextBox_RetractHeight.Text)
                ArrayPara1.ClearanceHeight = CDbl(TextBox_ClearanceHt.Text)
                ArrayPara1.DetailingDistance = CDbl(TextBox_DTailDist.Text)
                ArrayPara1.ArcRadius = CDbl(TextBox_ArcRadius.Text)


                ArrayPara1.Pos1.X = CDbl(TextBox_P1X.Text)
                ArrayPara1.Pos1.Y = CDbl(TextBox_P1Y.Text)
                ArrayPara1.Pos1.Z = CDbl(TextBox_P1Z.Text)
                ArrayPara1.Pos2.X = CDbl(TextBox_P2X.Text)
                ArrayPara1.Pos2.Y = CDbl(TextBox_P2Y.Text)
                ArrayPara1.pos3.X = CDbl(TextBox_P3X.Text)
                ArrayPara1.pos3.Y = CDbl(TextBox_P3Y.Text)

                ArrayPara2.Pos1.X = CDbl(TextBox_P4X.Text)
                ArrayPara2.Pos1.Y = CDbl(TextBox_P4Y.Text)

                ArrayPara3.Pos1.X = CDbl(TextBox_P7X.Text)
                ArrayPara3.Pos1.Y = CDbl(TextBox_P7Y.Text)


            Case "Circle"
                ArrayPara1.Needle = TextBox_Needle.Text
                ArrayPara1.DispenseFlag = TextBox_Dispense.Text
                ArrayPara1.TravelSpeed = CDbl(TextBox_TravelSpeed.Text)
                ArrayPara1.NeedleGap = CDbl(TextBox_NeedleGap.Text)
                ArrayPara1.TravelDelay = CDbl(TextBox_TravelDelay.Text)
                ArrayPara1.RetractDelay = CDbl(TextBox_RetractDelay.Text)
                ArrayPara1.ApproachDispHeight = CDbl(TextBox_ApproachHeight.Text)
                ArrayPara1.RetractSpeed = CDbl(TextBox_RetractSpeed.Text)
                ArrayPara1.RetractHeight = CDbl(TextBox_RetractHeight.Text)
                ArrayPara1.ClearanceHeight = CDbl(TextBox_ClearanceHt.Text)
                ArrayPara1.DetailingDistance = CDbl(TextBox_DTailDist.Text)
                ArrayPara1.ArcRadius = CDbl(TextBox_ArcRadius.Text)


                ArrayPara1.Pos1.X = CDbl(TextBox_P1X.Text)
                ArrayPara1.Pos1.Y = CDbl(TextBox_P1Y.Text)
                ArrayPara1.Pos1.Z = CDbl(TextBox_P1Z.Text)
                ArrayPara1.Pos2.X = CDbl(TextBox_P2X.Text)
                ArrayPara1.Pos2.Y = CDbl(TextBox_P2Y.Text)
                ArrayPara1.pos3.X = CDbl(TextBox_P3X.Text)
                ArrayPara1.pos3.Y = CDbl(TextBox_P3Y.Text)

                ArrayPara2.Pos1.X = CDbl(TextBox_P4X.Text)
                ArrayPara2.Pos1.Y = CDbl(TextBox_P4Y.Text)

                ArrayPara3.Pos1.X = CDbl(TextBox_P7X.Text)
                ArrayPara3.Pos1.Y = CDbl(TextBox_P7Y.Text)


            Case "FillCircle"
                ArrayPara1.Needle = TextBox_Needle.Text
                ArrayPara1.DispenseFlag = TextBox_Dispense.Text
                ArrayPara1.TravelSpeed = CDbl(TextBox_TravelSpeed.Text)
                ArrayPara1.NeedleGap = CDbl(TextBox_NeedleGap.Text)

                ArrayPara1.TravelDelay = CDbl(TextBox_TravelDelay.Text)
                ArrayPara1.RetractDelay = CDbl(TextBox_RetractDelay.Text)
                ArrayPara1.ApproachDispHeight = CDbl(TextBox_ApproachHeight.Text)
                ArrayPara1.RetractSpeed = CDbl(TextBox_RetractSpeed.Text)
                ArrayPara1.RetractHeight = CDbl(TextBox_RetractHeight.Text)
                ArrayPara1.ClearanceHeight = CDbl(TextBox_ClearanceHt.Text)
                ArrayPara1.DetailingDistance = CDbl(TextBox_DTailDist.Text)
                ArrayPara1.ArcRadius = CDbl(TextBox_ArcRadius.Text)

                ArrayPara1.Pitch = CDbl(TextBox_Pitch.Text)
                ArrayPara1.FilledHeight = CDbl(TextBox_FillHeight.Text)
                'SJ modify 28/09/09
                ArrayPara1.NumofRun = CInt(TextBox_NoOfRun.Text)

                '''''''''''''''''''''''''''''''''''''''''
                '   Xue Wen                             '
                '   Save this parameter in arrayList.   '
                '''''''''''''''''''''''''''''''''''''''''
                ArrayPara1.SpiralFlag = TextBox_Sprial.Text.ToUpper

                ArrayPara1.Pos1.X = CDbl(TextBox_P1X.Text)
                ArrayPara1.Pos1.Y = CDbl(TextBox_P1Y.Text)
                ArrayPara1.Pos1.Z = CDbl(TextBox_P1Z.Text)
                ArrayPara1.Pos2.X = CDbl(TextBox_P2X.Text)
                ArrayPara1.Pos2.Y = CDbl(TextBox_P2Y.Text)
                ArrayPara1.pos3.X = CDbl(TextBox_P3X.Text)
                ArrayPara1.pos3.Y = CDbl(TextBox_P3Y.Text)

                ArrayPara2.Pos1.X = CDbl(TextBox_P4X.Text)
                ArrayPara2.Pos1.Y = CDbl(TextBox_P4Y.Text)

                ArrayPara3.Pos1.X = CDbl(TextBox_P7X.Text)
                ArrayPara3.Pos1.Y = CDbl(TextBox_P7Y.Text)

            Case "Rectangle"
                ArrayPara1.Needle = TextBox_Needle.Text
                ArrayPara1.DispenseFlag = TextBox_Dispense.Text
                ArrayPara1.TravelSpeed = CDbl(TextBox_TravelSpeed.Text)
                ArrayPara1.NeedleGap = CDbl(TextBox_NeedleGap.Text)
                ArrayPara1.TravelDelay = CDbl(TextBox_TravelDelay.Text)
                ArrayPara1.RetractDelay = CDbl(TextBox_RetractDelay.Text)
                ArrayPara1.ApproachDispHeight = CDbl(TextBox_ApproachHeight.Text)
                ArrayPara1.RetractSpeed = CDbl(TextBox_RetractSpeed.Text)
                ArrayPara1.RetractHeight = CDbl(TextBox_RetractHeight.Text)
                ArrayPara1.ClearanceHeight = CDbl(TextBox_ClearanceHt.Text)
                ArrayPara1.DetailingDistance = CDbl(TextBox_DTailDist.Text)
                ArrayPara1.ArcRadius = CDbl(TextBox_ArcRadius.Text)


                ArrayPara1.Pos1.X = CDbl(TextBox_P1X.Text)
                ArrayPara1.Pos1.Y = CDbl(TextBox_P1Y.Text)
                ArrayPara1.Pos1.Z = CDbl(TextBox_P1Z.Text)
                ArrayPara1.Pos2.X = CDbl(TextBox_P2X.Text)
                ArrayPara1.Pos2.Y = CDbl(TextBox_P2Y.Text)
                ArrayPara1.pos3.X = CDbl(TextBox_P3X.Text)
                ArrayPara1.pos3.Y = CDbl(TextBox_P3Y.Text)

                ArrayPara2.Pos1.X = CDbl(TextBox_P4X.Text)
                ArrayPara2.Pos1.Y = CDbl(TextBox_P4Y.Text)

                ArrayPara3.Pos1.X = CDbl(TextBox_P7X.Text)
                ArrayPara3.Pos1.Y = CDbl(TextBox_P7Y.Text)


            Case "FillRectangle"
                ArrayPara1.Needle = TextBox_Needle.Text
                ArrayPara1.DispenseFlag = TextBox_Dispense.Text
                ArrayPara1.TravelSpeed = CDbl(TextBox_TravelSpeed.Text)
                ArrayPara1.NeedleGap = CDbl(TextBox_NeedleGap.Text)

                ArrayPara1.TravelDelay = CDbl(TextBox_TravelDelay.Text)
                ArrayPara1.RetractDelay = CDbl(TextBox_RetractDelay.Text)
                ArrayPara1.ApproachDispHeight = CDbl(TextBox_ApproachHeight.Text)
                ArrayPara1.RetractSpeed = CDbl(TextBox_RetractSpeed.Text)
                ArrayPara1.RetractHeight = CDbl(TextBox_RetractHeight.Text)
                ArrayPara1.ClearanceHeight = CDbl(TextBox_ClearanceHt.Text)
                ArrayPara1.DetailingDistance = CDbl(TextBox_DTailDist.Text)
                ArrayPara1.ArcRadius = CDbl(TextBox_ArcRadius.Text)

                ArrayPara1.Pitch = CDbl(TextBox_Pitch.Text)
                ArrayPara1.FilledHeight = CDbl(TextBox_FillHeight.Text)
                'SJ modify 28/09/09
                ArrayPara1.NumofRun = CInt(TextBox_NoOfRun.Text)

                '''''''''''''''''''''''''''''''''''''''''
                '   Xue Wen                             '
                '   Save this parameter in arrayList.   '
                '''''''''''''''''''''''''''''''''''''''''
                ArrayPara1.SpiralFlag = TextBox_Sprial.Text.ToUpper

                ArrayPara1.Pos1.X = CDbl(TextBox_P1X.Text)
                ArrayPara1.Pos1.Y = CDbl(TextBox_P1Y.Text)
                ArrayPara1.Pos1.Z = CDbl(TextBox_P1Z.Text)
                ArrayPara1.Pos2.X = CDbl(TextBox_P2X.Text)
                ArrayPara1.Pos2.Y = CDbl(TextBox_P2Y.Text)
                ArrayPara1.pos3.X = CDbl(TextBox_P3X.Text)
                ArrayPara1.pos3.Y = CDbl(TextBox_P3Y.Text)

                ArrayPara2.Pos1.X = CDbl(TextBox_P4X.Text)
                ArrayPara2.Pos1.Y = CDbl(TextBox_P4Y.Text)

                ArrayPara3.Pos1.X = CDbl(TextBox_P7X.Text)
                ArrayPara3.Pos1.Y = CDbl(TextBox_P7Y.Text)

            Case "SubPattern"

                ArrayPara1.DispenseFlag = TextBox_Dispense.Text

                ArrayPara1.Pos1.X = CDbl(TextBox_P1X.Text)
                ArrayPara1.Pos1.Y = CDbl(TextBox_P1Y.Text)
                ArrayPara1.Pos1.Z = CDbl(TextBox_P1Z.Text)

                ArrayPara2.Pos1.X = CDbl(TextBox_P4X.Text)
                ArrayPara2.Pos1.Y = CDbl(TextBox_P4Y.Text)

                ArrayPara3.Pos1.X = CDbl(TextBox_P7X.Text)
                ArrayPara3.Pos1.Y = CDbl(TextBox_P7Y.Text)
        End Select

        TraceGCCollect()
    End Sub

    Private Function CheckValidate(ByVal ElementType As String) As Boolean
        Dim Rtn As Boolean = True

        Select Case ElementType
            Case "Dot"
                TextBox_RetractSpeed.Text = CStr(m_CurrentPara.RetractSpeed)
                TextBox_RetractHeight.Text = CStr(m_CurrentPara.RetractHeight)
                TextBox_ClearanceHt.Text = CStr(m_CurrentPara.ClearanceHeight)
                TextBox_ArcRadius.Text = CStr(m_CurrentPara.ArcRadius)


                If "Left" <> TextBox_Needle.Text And "Right" <> TextBox_Needle.Text Then
                    MyMsgBox("Please input valid Needle information", MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly, "Input number invalid")
                    Rtn = False
                End If
                If "On" <> TextBox_Dispense.Text And "Off" <> TextBox_Dispense.Text Then
                    MyMsgBox("Please input valid Dispense information", MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly, "Input number invalid")
                    Rtn = False
                End If


                If Not IsNumeric(TextBox_NeedleGap.Text) Then
                    MyMsgBox("Please input valid number for NeedleGap", MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly, "Input number invalid")
                    Rtn = False
                End If
                If Not IsNumeric(TextBox_Duration.Text) Then
                    MyMsgBox("Please input valid number for Duration", MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly, "Input number invalid")
                    Rtn = False
                End If
                If Not IsNumeric(TextBox_RetractDelay.Text) Then
                    MyMsgBox("Please input valid number for RetractDelay", MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly, "Input number invalid")
                    Rtn = False
                End If
                If Not IsNumeric(TextBox_ApproachHeight.Text) Then
                    MyMsgBox("Please input valid number for ApproachHeight", MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly, "Input number invalid")
                    Rtn = False
                End If
                If Not IsNumeric(TextBox_RetractSpeed.Text) Then
                    MyMsgBox("Please input valid number for RetractSpeed", MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly, "Input number invalid")
                    Rtn = False
                End If
                If Not IsNumeric(TextBox_RetractHeight.Text) Then
                    MyMsgBox("Please input valid number for RetractHeight", MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly, "Input number invalid")
                    Rtn = False
                End If
                If Not IsNumeric(TextBox_ClearanceHt.Text) Then
                    MyMsgBox("Please input valid number for ClearanceHt", MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly, "Input number invalid")
                    Rtn = False
                End If
                If Not IsNumeric(TextBox_ArcRadius.Text) Then
                    MyMsgBox("Please input valid number for ArcRadius", MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly, "Input number invalid")
                    Rtn = False
                End If


                If IsNumeric(TextBox_P1X.Text) And IsNumeric(TextBox_P1Y.Text) And IsNumeric(TextBox_P1Z.Text) Then

                Else
                    MyMsgBox("Please input valid number for Pt1", MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly, "Input number invalid")
                    Rtn = False
                End If
                If IsNumeric(TextBox_P4X.Text) And IsNumeric(TextBox_P4Y.Text) Then

                Else
                    MyMsgBox("Please input valid number for Pt4", MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly, "Input number invalid")
                    Rtn = False
                End If
                If IsNumeric(TextBox_P7X.Text) And IsNumeric(TextBox_P7Y.Text) Then

                Else
                    MyMsgBox("Please input valid number for Pt7", MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly, "Input number invalid")
                    Rtn = False
                End If

            Case "Line"

                If "Left" <> TextBox_Needle.Text And "Right" <> TextBox_Needle.Text Then
                    MyMsgBox("Please input valid Needle information", MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly, "Input number invalid")
                    Rtn = False
                End If

                If "On" <> TextBox_Dispense.Text And "Off" <> TextBox_Dispense.Text Then
                    MyMsgBox("Please input valid Dispense information", MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly, "Input number invalid")
                    Rtn = False
                End If

                If Not IsNumeric(TextBox_TravelSpeed.Text) Then
                    MyMsgBox("Please input valid number for TravelSpeed", MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly, "Input number invalid")
                    Rtn = False
                End If

                If Not IsNumeric(TextBox_NeedleGap.Text) Then
                    MyMsgBox("Please input valid number for NeedleGap", MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly, "Input number invalid")
                    Rtn = False
                End If

                If Not IsNumeric(TextBox_TravelDelay.Text) Then
                    MyMsgBox("Please input valid number for TravelDelay", MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly, "Input number invalid")
                    Rtn = False
                End If
                If Not IsNumeric(TextBox_RetractDelay.Text) Then
                    MyMsgBox("Please input valid number for RetractDelay", MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly, "Input number invalid")
                    Rtn = False
                End If
                If Not IsNumeric(TextBox_ApproachHeight.Text) Then
                    MyMsgBox("Please input valid number for ApproachHeight", MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly, "Input number invalid")
                    Rtn = False
                End If
                If Not IsNumeric(TextBox_RetractSpeed.Text) Then
                    MyMsgBox("Please input valid number for RetractSpeed", MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly, "Input number invalid")
                    Rtn = False
                End If
                If Not IsNumeric(TextBox_RetractHeight.Text) Then
                    MyMsgBox("Please input valid number for RetractHeight", MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly, "Input number invalid")
                    Rtn = False
                End If
                If Not IsNumeric(TextBox_ClearanceHt.Text) Then
                    MyMsgBox("Please input valid number for ClearanceHeight", MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly, "Input number invalid")
                    Rtn = False
                End If
                If Not IsNumeric(TextBox_DTailDist.Text) Then
                    MyMsgBox("Please input valid number for DetailingDistance", MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly, "Input number invalid")
                    Rtn = False
                End If
                If Not IsNumeric(TextBox_ArcRadius.Text) Then
                    MyMsgBox("Please input valid number for ArcRadius", MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly, "Input number invalid")
                    Rtn = False
                End If

                If Not (IsNumeric(TextBox_P4X.Text) And IsNumeric(TextBox_P4Y.Text)) Then
                    MyMsgBox("Please input valid number for P4X,Y", MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly, "Input number invalid")
                    Rtn = False
                End If

                If Not (IsNumeric(TextBox_P7X.Text) And IsNumeric(TextBox_P7Y.Text)) Then
                    MyMsgBox("Please input valid number for P7X,Y", MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly, "Input number invalid")
                    Rtn = False
                End If


            Case "Arc"
                If "Left" <> TextBox_Needle.Text And "Right" <> TextBox_Needle.Text Then
                    MyMsgBox("Please input valid Needle information", MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly, "Input number invalid")
                    Rtn = False
                End If

                If "On" <> TextBox_Dispense.Text And "Off" <> TextBox_Dispense.Text Then
                    MyMsgBox("Please input valid Dispense information", MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly, "Input number invalid")
                    Rtn = False
                End If

                If Not IsNumeric(TextBox_TravelSpeed.Text) Then
                    MyMsgBox("Please input valid number for TravelSpeed", MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly, "Input number invalid")
                    Rtn = False
                End If
                If Not IsNumeric(TextBox_NeedleGap.Text) Then
                    MyMsgBox("Please input valid number for NeedleGap", MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly, "Input number invalid")
                    Rtn = False
                End If

                If Not IsNumeric(TextBox_TravelDelay.Text) Then
                    MyMsgBox("Please input valid number for TravelDelay", MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly, "Input number invalid")
                    Rtn = False
                End If
                If Not IsNumeric(TextBox_RetractDelay.Text) Then
                    MyMsgBox("Please input valid number for RetractDelay", MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly, "Input number invalid")
                    Rtn = False
                End If
                If Not IsNumeric(TextBox_ApproachHeight.Text) Then
                    MyMsgBox("Please input valid number for ApproachHeight", MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly, "Input number invalid")
                    Rtn = False
                End If
                If Not IsNumeric(TextBox_RetractSpeed.Text) Then
                    MyMsgBox("Please input valid number for RetractSpeed", MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly, "Input number invalid")
                    Rtn = False
                End If

                If Not IsNumeric(TextBox_RetractHeight.Text) Then
                    MyMsgBox("Please input valid number for RetractHeight", MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly, "Input number invalid")
                    Rtn = False
                End If
                If Not IsNumeric(TextBox_ClearanceHt.Text) Then
                    MyMsgBox("Please input valid number for ClearanceHeight", MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly, "Input number invalid")
                    Rtn = False
                End If
                If Not IsNumeric(TextBox_DTailDist.Text) Then
                    MyMsgBox("Please input valid number for DetailingDistance", MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly, "Input number invalid")
                    Rtn = False
                End If
                If Not IsNumeric(TextBox_ArcRadius.Text) Then
                    MyMsgBox("Please input valid number for ArcRadius", MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly, "Input number invalid")
                    Rtn = False
                End If


                If Not (IsNumeric(TextBox_P2X.Text) And IsNumeric(TextBox_P2Y.Text)) Then
                    MyMsgBox("Please input valid number for P2X,Y", MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly, "Input number invalid")
                    Rtn = False
                End If

                If Not (IsNumeric(TextBox_P3X.Text) And IsNumeric(TextBox_P3Y.Text)) Then
                    MyMsgBox("Please input valid number for P3X,Y", MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly, "Input number invalid")
                    Rtn = False
                End If

                If Not (IsNumeric(TextBox_P4X.Text) And IsNumeric(TextBox_P4Y.Text)) Then
                    MyMsgBox("Please input valid number for P4X,Y", MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly, "Input number invalid")
                    Rtn = False
                End If

                If Not (IsNumeric(TextBox_P7X.Text) And IsNumeric(TextBox_P7Y.Text)) Then
                    MyMsgBox("Please input valid number for P7X,Y", MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly, "Input number invalid")
                    Rtn = False
                End If



            Case "Circle", "Rectangle"
                If "Left" <> TextBox_Needle.Text And "Right" <> TextBox_Needle.Text Then
                    MyMsgBox("Please input valid Needle information", MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly, "Input number invalid")
                    Rtn = False
                End If
                If "On" <> TextBox_Dispense.Text And "Off" <> TextBox_Dispense.Text Then
                    MyMsgBox("Please input valid Dispense information", MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly, "Input number invalid")
                    Rtn = False
                End If
                If Not IsNumeric(TextBox_TravelSpeed.Text) Then
                    MyMsgBox("Please input valid number for TravelSpeed", MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly, "Input number invalid")
                    Rtn = False
                End If
                If Not IsNumeric(TextBox_NeedleGap.Text) Then
                    MyMsgBox("Please input valid number for NeedleGap", MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly, "Input number invalid")
                    Rtn = False
                End If

                If Not IsNumeric(TextBox_TravelDelay.Text) Then
                    MyMsgBox("Please input valid number for TravelDelay", MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly, "Input number invalid")
                    Rtn = False
                End If
                If Not IsNumeric(TextBox_RetractDelay.Text) Then
                    MyMsgBox("Please input valid number for RetractDelay", MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly, "Input number invalid")
                    Rtn = False
                End If
                If Not IsNumeric(TextBox_ApproachHeight.Text) Then
                    MyMsgBox("Please input valid number for ApproachHeight", MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly, "Input number invalid")
                    Rtn = False
                End If
                If Not IsNumeric(TextBox_RetractSpeed.Text) Then
                    MyMsgBox("Please input valid number for RetractSpeed", MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly, "Input number invalid")
                    Rtn = False
                End If


                If Not IsNumeric(TextBox_RetractHeight.Text) Then
                    MyMsgBox("Please input valid number for RetractHeight", MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly, "Input number invalid")
                    Rtn = False
                End If
                If Not IsNumeric(TextBox_ClearanceHt.Text) Then
                    MyMsgBox("Please input valid number for ClearanceHeight", MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly, "Input number invalid")
                    Rtn = False
                End If
                If Not IsNumeric(TextBox_DTailDist.Text) Then
                    MyMsgBox("Please input valid number for DetailingDistance", MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly, "Input number invalid")
                    Rtn = False
                End If
                If Not IsNumeric(TextBox_ArcRadius.Text) Then
                    MyMsgBox("Please input valid number for ArcRadius", MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly, "Input number invalid")
                    Rtn = False
                End If

                If Not (IsNumeric(TextBox_P1X.Text) And IsNumeric(TextBox_P1Y.Text)) Then
                    MyMsgBox("Please input valid number for P1X,Y", MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly, "Input number invalid")
                    Rtn = False
                End If
                If Not (IsNumeric(TextBox_P2X.Text) And IsNumeric(TextBox_P2Y.Text)) Then
                    MyMsgBox("Please input valid number for P2X,Y", MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly, "Input number invalid")
                    Rtn = False
                End If
                If Not (IsNumeric(TextBox_P3X.Text) And IsNumeric(TextBox_P3Y.Text)) Then
                    MyMsgBox("Please input valid number for P3X,Y", MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly, "Input number invalid")
                    Rtn = False
                End If
                If Not (IsNumeric(TextBox_P4X.Text) And IsNumeric(TextBox_P4Y.Text)) Then
                    MyMsgBox("Please input valid number for P4X,Y", MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly, "Input number invalid")
                    Rtn = False
                End If

                If Not (IsNumeric(TextBox_P7X.Text) And IsNumeric(TextBox_P7Y.Text)) Then
                    MyMsgBox("Please input valid number for P7X,Y", MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly, "Input number invalid")
                    Rtn = False
                End If

            Case "FillCircle", "FillRectangle"
                If "Left" <> TextBox_Needle.Text And "Right" <> TextBox_Needle.Text Then
                    MyMsgBox("Please input valid Needle information", MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly, "Input number invalid")
                    Rtn = False
                End If
                If "On" <> TextBox_Dispense.Text And "Off" <> TextBox_Dispense.Text Then
                    MyMsgBox("Please input valid Dispense information", MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly, "Input number invalid")
                    Rtn = False
                End If
                If Not IsNumeric(TextBox_TravelSpeed.Text) Then
                    MyMsgBox("Please input valid number for TravelSpeed", MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly, "Input number invalid")
                    Rtn = False
                End If
                If Not IsNumeric(TextBox_NeedleGap.Text) Then
                    MyMsgBox("Please input valid number for NeedleGap", MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly, "Input number invalid")
                    Rtn = False
                End If

                If Not IsNumeric(TextBox_TravelDelay.Text) Then
                    MyMsgBox("Please input valid number for TravelDelay", MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly, "Input number invalid")
                    Rtn = False
                End If
                If Not IsNumeric(TextBox_RetractDelay.Text) Then
                    MyMsgBox("Please input valid number for RetractDelay", MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly, "Input number invalid")
                    Rtn = False
                End If
                If Not IsNumeric(TextBox_ApproachHeight.Text) Then
                    MyMsgBox("Please input valid number for ApproachHeight", MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly, "Input number invalid")
                    Rtn = False
                End If
                If Not IsNumeric(TextBox_RetractSpeed.Text) Then
                    MyMsgBox("Please input valid number for RetractSpeed", MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly, "Input number invalid")
                    Rtn = False
                End If

                If Not IsNumeric(TextBox_RetractHeight.Text) Then
                    MyMsgBox("Please input valid number for RetractHeight", MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly, "Input number invalid")
                    Rtn = False
                End If
                If Not IsNumeric(TextBox_ClearanceHt.Text) Then
                    MyMsgBox("Please input valid number for ClearanceHeight", MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly, "Input number invalid")
                    Rtn = False
                End If
                If Not IsNumeric(TextBox_DTailDist.Text) Then
                    MyMsgBox("Please input valid number for DetailingDistance", MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly, "Input number invalid")
                    Rtn = False
                End If
                If Not IsNumeric(TextBox_ArcRadius.Text) Then
                    MyMsgBox("Please input valid number for ArcRadius", MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly, "Input number invalid")
                    Rtn = False
                End If


                If Not IsNumeric(TextBox_Pitch.Text) Then
                    MyMsgBox("Please input valid number for Pitch", MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly, "Input number invalid")
                    Rtn = False
                End If
                If Not IsNumeric(TextBox_FillHeight.Text) Then
                    MyMsgBox("Please input valid number for FilledHeight", MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly, "Input number invalid")
                    Rtn = False
                End If
                If Not IsNumeric(TextBox_NoOfRun.Text) Then
                    MyMsgBox("Please input valid number for NumofRun", MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly, "Input number invalid")
                    Rtn = False
                End If

                If "IN" <> TextBox_Sprial.Text.ToUpper And "OUT" <> TextBox_Sprial.Text.ToUpper Then
                    MyMsgBox("Please input valid SpiralFlag information", MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly, "Input number invalid")
                    Rtn = False
                End If


                If Not (IsNumeric(TextBox_P1X.Text) And IsNumeric(TextBox_P1Y.Text)) Then
                    MyMsgBox("Please input valid number for P1X,Y", MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly, "Input number invalid")
                    Rtn = False
                End If
                If Not (IsNumeric(TextBox_P2X.Text) And IsNumeric(TextBox_P2Y.Text)) Then
                    MyMsgBox("Please input valid number for P2X,Y", MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly, "Input number invalid")
                    Rtn = False
                End If
                If Not (IsNumeric(TextBox_P3X.Text) And IsNumeric(TextBox_P3Y.Text)) Then
                    MyMsgBox("Please input valid number for P3X,Y", MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly, "Input number invalid")
                    Rtn = False
                End If
                If Not (IsNumeric(TextBox_P4X.Text) And IsNumeric(TextBox_P4Y.Text)) Then
                    MyMsgBox("Please input valid number for P4X,Y", MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly, "Input number invalid")
                    Rtn = False
                End If

                If Not (IsNumeric(TextBox_P7X.Text) And IsNumeric(TextBox_P7Y.Text)) Then
                    MyMsgBox("Please input valid number for P7X,Y", MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly, "Input number invalid")
                    Rtn = False
                End If

            Case "SubPattern"
                If Nothing = TextBox_Dispense.Text Then
                    MyMsgBox("Please input valid filename information", MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly, "Input number invalid")


                    Dim DialogPreview As New FormSelectPatternFile
                    DialogPreview.Location.Offset(0, 200)

                    DialogPreview.ShowDialog()
                    Dim file As String = DialogPreview.Path
                    TextBox_Dispense.Text = file

                    If Nothing = TextBox_Dispense.Text Then
                        MyMsgBox("Invalid filename", MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly, "Input invalid")
                        Rtn = False
                    End If

                End If

                If IsNumeric(TextBox_P1X.Text) And IsNumeric(TextBox_P1Y.Text) And IsNumeric(TextBox_P1Z.Text) Then

                Else
                    MyMsgBox("Please input valid number for Pt1", MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly, "Input number invalid")
                    Rtn = False
                End If
                If IsNumeric(TextBox_P4X.Text) And IsNumeric(TextBox_P4Y.Text) Then

                Else
                    MyMsgBox("Please input valid number for Pt4", MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly, "Input number invalid")
                    Rtn = False
                End If
                If IsNumeric(TextBox_P7X.Text) And IsNumeric(TextBox_P7Y.Text) Then

                Else
                    MyMsgBox("Please input valid number for Pt7", MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly, "Input number invalid")
                    Rtn = False
                End If
                Return Rtn
        End Select

        'Check if the approach z, needle gap, clearance height, retract z is within the z range.
        Dim ApproachZ As Double = 0
        Dim NeedleGapZ As Double = 0
        Dim ClearanceZ As Double = 0
        Dim RetractZ As Double = 0
        Dim needleTipReferenceZ As Double = 0
        If Not (isVisionMode) Then
            needleTipReferenceZ = IDS.Data.Hardware.Needle.Left.NeedleCalibrationPosition.Z
        End If
        ApproachZ = Convert.ToDouble(TextBox_P1Z.Text) + Convert.ToDouble(TextBox_ApproachHeight.Text)
        ApproachZ = ApproachZ + gSystemOrigin(2) + needleTipReferenceZ 'Get the hardware coordinate
        If (WorkAreaErrorCheckZ(ApproachZ) = False) Then
            MessageBox.Show("Approach Z Height Error: " + ErrorMessage())
            Return False
        End If
        NeedleGapZ = Convert.ToDouble(TextBox_P1Z.Text) + Convert.ToDouble(TextBox_NeedleGap.Text)
        NeedleGapZ = NeedleGapZ + gSystemOrigin(2) + needleTipReferenceZ 'Get the hardware coordinate
        If (WorkAreaErrorCheckZ(NeedleGapZ) = False) Then
            MessageBox.Show("Needle Gap Z Height Error: " + ErrorMessage())
            Return False
        End If
        ClearanceZ = Convert.ToDouble(TextBox_P1Z.Text) + Convert.ToDouble(TextBox_ClearanceHt.Text)
        ClearanceZ = ClearanceZ + gSystemOrigin(2) + needleTipReferenceZ 'Get the hardware coordinate
        If (WorkAreaErrorCheckZ(ClearanceZ) = False) Then
            MessageBox.Show("Clearance Z Height Error: " + ErrorMessage())
            Return False
        End If
        RetractZ = Convert.ToDouble(TextBox_P1Z.Text) + Convert.ToDouble(TextBox_RetractHeight.Text)
        RetractZ = RetractZ + gSystemOrigin(2) + needleTipReferenceZ 'Get the hardware coordinate
        If (WorkAreaErrorCheckZ(RetractZ) = False) Then
            MessageBox.Show("Retract Z Height Error: " + ErrorMessage())
            Return False
        End If

        Return Rtn
        TraceGCCollect()
    End Function
    Public Sub SetDispenseDuration(ByVal dispenseTime As String)
        TextBox_Duration.Text = dispenseTime
        m_CurrentPara.DispenseDuration = Convert.ToDouble(dispenseTime)
    End Sub
    Public Sub SetNeedleGap(ByVal needleGap As String)
        TextBox_NeedleGap.Text = needleGap
        m_CurrentPara.NeedleGap = Convert.ToDouble(needleGap)
    End Sub
    Private Sub Button_OnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_OnOK.Click
        Dim typeSelected As String = CombElementType.Text

        If CheckValidate(CombElementType.Text) Then
            Me.DialogResult = DialogResult.OK
            ArrayTextBoxStore(typeSelected)
            Me.Close()
        End If
    End Sub

    Private Sub Button_OnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_OnCancel.Click
        Me.DialogResult = DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub CombElementType_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles CombElementType.SelectedIndexChanged
        Dim typeSelected As String = CombElementType.SelectedItem
        RadioButton_P1.Enabled = True
        ArrayButtonEnable(typeSelected)
        ArrayTextBoxEnable(typeSelected)

        ArrayTextBoxFillIn(typeSelected)
        Button_OnOK.Enabled = True
        Button_OnCancel.Enabled = True
        TraceGCCollect()
    End Sub
End Class
