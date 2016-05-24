Imports DLL_Export_Service

Public Class SPCLogging
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

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
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents CB_EnableSPCLog As System.Windows.Forms.CheckBox
    Friend WithEvents Label17 As System.Windows.Forms.Label
    Friend WithEvents PanelToBeAdded As System.Windows.Forms.Panel
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents CB_ReportDirectory As System.Windows.Forms.ComboBox
    Friend WithEvents CB_SPCPassNum As System.Windows.Forms.CheckBox
    Friend WithEvents CB_SPCOutgoNum As System.Windows.Forms.CheckBox
    Friend WithEvents CB_SPCIncomeNum As System.Windows.Forms.CheckBox
    Friend WithEvents CB_SPCMaterialBatch As System.Windows.Forms.CheckBox
    Friend WithEvents CB_SPCPartFailedNum As System.Windows.Forms.CheckBox
    Friend WithEvents CB_SPCNeedleCalNum As System.Windows.Forms.CheckBox
    Friend WithEvents CB_SPCVolCalNum As System.Windows.Forms.CheckBox
    Friend WithEvents CB_SPCPauseNum As System.Windows.Forms.CheckBox
    Friend WithEvents CB_SPCDownTime As System.Windows.Forms.CheckBox
    Friend WithEvents CB_SPCRunTime As System.Windows.Forms.CheckBox
    Friend WithEvents CB_SPCTotalFailedNum As System.Windows.Forms.CheckBox
    Friend WithEvents ButtonRevert As System.Windows.Forms.Button
    Friend WithEvents ButtonSave As System.Windows.Forms.Button
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Nud_EventCleanInterval As System.Windows.Forms.NumericUpDown
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents CB_SPCBoardPerHour As System.Windows.Forms.CheckBox
    Friend WithEvents CB_SPCTimePerBoard As System.Windows.Forms.CheckBox
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents ButtonExit As System.Windows.Forms.Button
    Friend WithEvents tbReportDirectory As System.Windows.Forms.TextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(SPCLogging))
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.CB_SPCBoardPerHour = New System.Windows.Forms.CheckBox
        Me.Label10 = New System.Windows.Forms.Label
        Me.Label9 = New System.Windows.Forms.Label
        Me.Label8 = New System.Windows.Forms.Label
        Me.Label7 = New System.Windows.Forms.Label
        Me.CB_SPCPassNum = New System.Windows.Forms.CheckBox
        Me.CB_SPCOutgoNum = New System.Windows.Forms.CheckBox
        Me.CB_SPCIncomeNum = New System.Windows.Forms.CheckBox
        Me.CB_SPCMaterialBatch = New System.Windows.Forms.CheckBox
        Me.CB_SPCPartFailedNum = New System.Windows.Forms.CheckBox
        Me.CB_SPCNeedleCalNum = New System.Windows.Forms.CheckBox
        Me.CB_SPCVolCalNum = New System.Windows.Forms.CheckBox
        Me.CB_SPCPauseNum = New System.Windows.Forms.CheckBox
        Me.CB_SPCDownTime = New System.Windows.Forms.CheckBox
        Me.CB_SPCRunTime = New System.Windows.Forms.CheckBox
        Me.CB_SPCTimePerBoard = New System.Windows.Forms.CheckBox
        Me.CB_SPCTotalFailedNum = New System.Windows.Forms.CheckBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.Button1 = New System.Windows.Forms.Button
        Me.CB_ReportDirectory = New System.Windows.Forms.ComboBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.CB_EnableSPCLog = New System.Windows.Forms.CheckBox
        Me.PanelToBeAdded = New System.Windows.Forms.Panel
        Me.Label5 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.Nud_EventCleanInterval = New System.Windows.Forms.NumericUpDown
        Me.Label2 = New System.Windows.Forms.Label
        Me.ButtonRevert = New System.Windows.Forms.Button
        Me.ButtonSave = New System.Windows.Forms.Button
        Me.Label17 = New System.Windows.Forms.Label
        Me.ButtonExit = New System.Windows.Forms.Button
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog
        Me.tbReportDirectory = New System.Windows.Forms.TextBox
        Me.GroupBox2.SuspendLayout()
        Me.PanelToBeAdded.SuspendLayout()
        CType(Me.Nud_EventCleanInterval, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.CB_SPCBoardPerHour)
        Me.GroupBox2.Controls.Add(Me.Label10)
        Me.GroupBox2.Controls.Add(Me.Label9)
        Me.GroupBox2.Controls.Add(Me.Label8)
        Me.GroupBox2.Controls.Add(Me.Label7)
        Me.GroupBox2.Controls.Add(Me.CB_SPCPassNum)
        Me.GroupBox2.Controls.Add(Me.CB_SPCOutgoNum)
        Me.GroupBox2.Controls.Add(Me.CB_SPCIncomeNum)
        Me.GroupBox2.Controls.Add(Me.CB_SPCMaterialBatch)
        Me.GroupBox2.Controls.Add(Me.CB_SPCPartFailedNum)
        Me.GroupBox2.Controls.Add(Me.CB_SPCNeedleCalNum)
        Me.GroupBox2.Controls.Add(Me.CB_SPCVolCalNum)
        Me.GroupBox2.Controls.Add(Me.CB_SPCPauseNum)
        Me.GroupBox2.Controls.Add(Me.CB_SPCDownTime)
        Me.GroupBox2.Controls.Add(Me.CB_SPCRunTime)
        Me.GroupBox2.Controls.Add(Me.CB_SPCTimePerBoard)
        Me.GroupBox2.Controls.Add(Me.CB_SPCTotalFailedNum)
        Me.GroupBox2.Controls.Add(Me.Label6)
        Me.GroupBox2.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.GroupBox2.Location = New System.Drawing.Point(8, 192)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(496, 480)
        Me.GroupBox2.TabIndex = 19
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Items to be reported"
        '
        'CB_SPCBoardPerHour
        '
        Me.CB_SPCBoardPerHour.Location = New System.Drawing.Point(288, 192)
        Me.CB_SPCBoardPerHour.Name = "CB_SPCBoardPerHour"
        Me.CB_SPCBoardPerHour.Size = New System.Drawing.Size(184, 24)
        Me.CB_SPCBoardPerHour.TabIndex = 43
        Me.CB_SPCBoardPerHour.Text = "No. of Boards/Hour"
        '
        'Label10
        '
        Me.Label10.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label10.Location = New System.Drawing.Point(40, 240)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(184, 24)
        Me.Label10.TabIndex = 42
        Me.Label10.Text = "Pattern File Name"
        Me.Label10.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label9
        '
        Me.Label9.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label9.Location = New System.Drawing.Point(40, 192)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(184, 24)
        Me.Label9.TabIndex = 41
        Me.Label9.Text = "Operator ID"
        Me.Label9.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label8
        '
        Me.Label8.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label8.Location = New System.Drawing.Point(40, 144)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(184, 24)
        Me.Label8.TabIndex = 40
        Me.Label8.Text = "Duration"
        Me.Label8.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label7
        '
        Me.Label7.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label7.Location = New System.Drawing.Point(40, 96)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(184, 24)
        Me.Label7.TabIndex = 39
        Me.Label7.Text = "Production End Time"
        Me.Label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'CB_SPCPassNum
        '
        Me.CB_SPCPassNum.Location = New System.Drawing.Point(24, 432)
        Me.CB_SPCPassNum.Name = "CB_SPCPassNum"
        Me.CB_SPCPassNum.Size = New System.Drawing.Size(240, 24)
        Me.CB_SPCPassNum.TabIndex = 17
        Me.CB_SPCPassNum.Text = "No. of Passed Boards"
        '
        'CB_SPCOutgoNum
        '
        Me.CB_SPCOutgoNum.Location = New System.Drawing.Point(24, 384)
        Me.CB_SPCOutgoNum.Name = "CB_SPCOutgoNum"
        Me.CB_SPCOutgoNum.Size = New System.Drawing.Size(248, 24)
        Me.CB_SPCOutgoNum.TabIndex = 16
        Me.CB_SPCOutgoNum.Text = "No. of Outgoing Boards"
        '
        'CB_SPCIncomeNum
        '
        Me.CB_SPCIncomeNum.Location = New System.Drawing.Point(24, 336)
        Me.CB_SPCIncomeNum.Name = "CB_SPCIncomeNum"
        Me.CB_SPCIncomeNum.Size = New System.Drawing.Size(248, 24)
        Me.CB_SPCIncomeNum.TabIndex = 15
        Me.CB_SPCIncomeNum.Text = "No. of Incoming Boards"
        '
        'CB_SPCMaterialBatch
        '
        Me.CB_SPCMaterialBatch.Location = New System.Drawing.Point(24, 288)
        Me.CB_SPCMaterialBatch.Name = "CB_SPCMaterialBatch"
        Me.CB_SPCMaterialBatch.Size = New System.Drawing.Size(176, 24)
        Me.CB_SPCMaterialBatch.TabIndex = 14
        Me.CB_SPCMaterialBatch.Text = "Material Batch"
        '
        'CB_SPCPartFailedNum
        '
        Me.CB_SPCPartFailedNum.Location = New System.Drawing.Point(288, 48)
        Me.CB_SPCPartFailedNum.Name = "CB_SPCPartFailedNum"
        Me.CB_SPCPartFailedNum.Size = New System.Drawing.Size(200, 24)
        Me.CB_SPCPartFailedNum.TabIndex = 13
        Me.CB_SPCPartFailedNum.Text = "No. of Partially Failed"
        '
        'CB_SPCNeedleCalNum
        '
        Me.CB_SPCNeedleCalNum.Location = New System.Drawing.Point(288, 432)
        Me.CB_SPCNeedleCalNum.Name = "CB_SPCNeedleCalNum"
        Me.CB_SPCNeedleCalNum.Size = New System.Drawing.Size(176, 24)
        Me.CB_SPCNeedleCalNum.TabIndex = 11
        Me.CB_SPCNeedleCalNum.Text = "No. of Needle Cal."
        '
        'CB_SPCVolCalNum
        '
        Me.CB_SPCVolCalNum.Location = New System.Drawing.Point(288, 384)
        Me.CB_SPCVolCalNum.Name = "CB_SPCVolCalNum"
        Me.CB_SPCVolCalNum.Size = New System.Drawing.Size(176, 24)
        Me.CB_SPCVolCalNum.TabIndex = 10
        Me.CB_SPCVolCalNum.Text = "No. of Volume Cal."
        '
        'CB_SPCPauseNum
        '
        Me.CB_SPCPauseNum.Location = New System.Drawing.Point(288, 336)
        Me.CB_SPCPauseNum.Name = "CB_SPCPauseNum"
        Me.CB_SPCPauseNum.Size = New System.Drawing.Size(160, 24)
        Me.CB_SPCPauseNum.TabIndex = 9
        Me.CB_SPCPauseNum.Text = "No. of Pauses"
        '
        'CB_SPCDownTime
        '
        Me.CB_SPCDownTime.Location = New System.Drawing.Point(288, 288)
        Me.CB_SPCDownTime.Name = "CB_SPCDownTime"
        Me.CB_SPCDownTime.Size = New System.Drawing.Size(168, 24)
        Me.CB_SPCDownTime.TabIndex = 8
        Me.CB_SPCDownTime.Text = "Total Down Time"
        '
        'CB_SPCRunTime
        '
        Me.CB_SPCRunTime.Location = New System.Drawing.Point(288, 240)
        Me.CB_SPCRunTime.Name = "CB_SPCRunTime"
        Me.CB_SPCRunTime.Size = New System.Drawing.Size(152, 24)
        Me.CB_SPCRunTime.TabIndex = 7
        Me.CB_SPCRunTime.Text = "Total Run Time"
        '
        'CB_SPCTimePerBoard
        '
        Me.CB_SPCTimePerBoard.Location = New System.Drawing.Point(288, 144)
        Me.CB_SPCTimePerBoard.Name = "CB_SPCTimePerBoard"
        Me.CB_SPCTimePerBoard.Size = New System.Drawing.Size(190, 24)
        Me.CB_SPCTimePerBoard.TabIndex = 6
        Me.CB_SPCTimePerBoard.Text = "Average Time/Board"
        '
        'CB_SPCTotalFailedNum
        '
        Me.CB_SPCTotalFailedNum.Location = New System.Drawing.Point(288, 96)
        Me.CB_SPCTotalFailedNum.Name = "CB_SPCTotalFailedNum"
        Me.CB_SPCTotalFailedNum.Size = New System.Drawing.Size(184, 24)
        Me.CB_SPCTotalFailedNum.TabIndex = 5
        Me.CB_SPCTotalFailedNum.Text = "No. of Totally Failed"
        '
        'Label6
        '
        Me.Label6.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label6.Location = New System.Drawing.Point(40, 48)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(184, 24)
        Me.Label6.TabIndex = 38
        Me.Label6.Text = "Production Start Time"
        Me.Label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Button1
        '
        Me.Button1.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Button1.Location = New System.Drawing.Point(408, 130)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(75, 32)
        Me.Button1.TabIndex = 14
        Me.Button1.Text = "Browse"
        '
        'CB_ReportDirectory
        '
        Me.CB_ReportDirectory.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CB_ReportDirectory.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.CB_ReportDirectory.Location = New System.Drawing.Point(256, 0)
        Me.CB_ReportDirectory.Name = "CB_ReportDirectory"
        Me.CB_ReportDirectory.Size = New System.Drawing.Size(352, 28)
        Me.CB_ReportDirectory.TabIndex = 13
        Me.CB_ReportDirectory.Visible = False
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label1.Location = New System.Drawing.Point(40, 104)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(144, 23)
        Me.Label1.TabIndex = 12
        Me.Label1.Text = "Report Directory"
        '
        'CB_EnableSPCLog
        '
        Me.CB_EnableSPCLog.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.CB_EnableSPCLog.Location = New System.Drawing.Point(16, 56)
        Me.CB_EnableSPCLog.Name = "CB_EnableSPCLog"
        Me.CB_EnableSPCLog.Size = New System.Drawing.Size(200, 24)
        Me.CB_EnableSPCLog.TabIndex = 11
        Me.CB_EnableSPCLog.Text = "Enable Auto Report"
        '
        'PanelToBeAdded
        '
        Me.PanelToBeAdded.BackColor = System.Drawing.SystemColors.Control
        Me.PanelToBeAdded.Controls.Add(Me.tbReportDirectory)
        Me.PanelToBeAdded.Controls.Add(Me.Label5)
        Me.PanelToBeAdded.Controls.Add(Me.Label3)
        Me.PanelToBeAdded.Controls.Add(Me.Nud_EventCleanInterval)
        Me.PanelToBeAdded.Controls.Add(Me.Label2)
        Me.PanelToBeAdded.Controls.Add(Me.ButtonRevert)
        Me.PanelToBeAdded.Controls.Add(Me.ButtonSave)
        Me.PanelToBeAdded.Controls.Add(Me.Label17)
        Me.PanelToBeAdded.Controls.Add(Me.CB_ReportDirectory)
        Me.PanelToBeAdded.Controls.Add(Me.GroupBox2)
        Me.PanelToBeAdded.Controls.Add(Me.CB_EnableSPCLog)
        Me.PanelToBeAdded.Controls.Add(Me.Label1)
        Me.PanelToBeAdded.Controls.Add(Me.Button1)
        Me.PanelToBeAdded.Controls.Add(Me.ButtonExit)
        Me.PanelToBeAdded.Location = New System.Drawing.Point(64, 8)
        Me.PanelToBeAdded.Name = "PanelToBeAdded"
        Me.PanelToBeAdded.Size = New System.Drawing.Size(512, 944)
        Me.PanelToBeAdded.TabIndex = 20
        '
        'Label5
        '
        Me.Label5.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label5.Location = New System.Drawing.Point(56, 744)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(376, 32)
        Me.Label5.TabIndex = 42
        Me.Label5.Text = "(Default is 10 days, by Maintenance Engineer)"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label3
        '
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label3.Location = New System.Drawing.Point(336, 712)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(48, 32)
        Me.Label3.TabIndex = 41
        Me.Label3.Text = "Days"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Nud_EventCleanInterval
        '
        Me.Nud_EventCleanInterval.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Nud_EventCleanInterval.Location = New System.Drawing.Point(264, 712)
        Me.Nud_EventCleanInterval.Name = "Nud_EventCleanInterval"
        Me.Nud_EventCleanInterval.Size = New System.Drawing.Size(64, 27)
        Me.Nud_EventCleanInterval.TabIndex = 40
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label2.Location = New System.Drawing.Point(56, 712)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(200, 32)
        Me.Label2.TabIndex = 39
        Me.Label2.Text = "Event Clean Up Interval"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'ButtonRevert
        '
        Me.ButtonRevert.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonRevert.Location = New System.Drawing.Point(392, 816)
        Me.ButtonRevert.Name = "ButtonRevert"
        Me.ButtonRevert.Size = New System.Drawing.Size(88, 48)
        Me.ButtonRevert.TabIndex = 38
        Me.ButtonRevert.Text = "Revert"
        '
        'ButtonSave
        '
        Me.ButtonSave.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonSave.Location = New System.Drawing.Point(272, 816)
        Me.ButtonSave.Name = "ButtonSave"
        Me.ButtonSave.Size = New System.Drawing.Size(88, 48)
        Me.ButtonSave.TabIndex = 37
        Me.ButtonSave.Text = "Save"
        '
        'Label17
        '
        Me.Label17.Font = New System.Drawing.Font("Microsoft Sans Serif", 16.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label17.ForeColor = System.Drawing.SystemColors.ActiveCaption
        Me.Label17.Location = New System.Drawing.Point(0, 0)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(144, 32)
        Me.Label17.TabIndex = 34
        Me.Label17.Text = "SPC Logging"
        '
        'ButtonExit
        '
        Me.ButtonExit.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonExit.Image = CType(resources.GetObject("ButtonExit.Image"), System.Drawing.Image)
        Me.ButtonExit.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.ButtonExit.Location = New System.Drawing.Point(416, 24)
        Me.ButtonExit.Name = "ButtonExit"
        Me.ButtonExit.Size = New System.Drawing.Size(75, 50)
        Me.ButtonExit.TabIndex = 46
        Me.ButtonExit.TabStop = False
        Me.ButtonExit.Text = "Exit"
        Me.ButtonExit.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'tbReportDirectory
        '
        Me.tbReportDirectory.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tbReportDirectory.Location = New System.Drawing.Point(32, 136)
        Me.tbReportDirectory.Name = "tbReportDirectory"
        Me.tbReportDirectory.ReadOnly = True
        Me.tbReportDirectory.Size = New System.Drawing.Size(368, 26)
        Me.tbReportDirectory.TabIndex = 47
        Me.tbReportDirectory.Text = ""
        '
        'SPCLogging
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(680, 912)
        Me.Controls.Add(Me.PanelToBeAdded)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Name = "SPCLogging"
        Me.Text = "SPC_Logging"
        Me.GroupBox2.ResumeLayout(False)
        Me.PanelToBeAdded.ResumeLayout(False)
        CType(Me.Nud_EventCleanInterval, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region
     

    Private Sub ButtonSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonSave.Click
        'IDS.Data.OpenData()

        IDS.Data.Hardware.SPC.CleanUpInterval = Nud_EventCleanInterval.Value
        IDS.Data.Hardware.SPC.EnableSPCReport = CB_EnableSPCLog.Checked
        IDS.Data.Hardware.SPC.ReportFileName = CB_ReportDirectory.Text
        IDS.Data.Hardware.SPC.ItemsToBeReported = Convert.ToByte(CB_SPCMaterialBatch.Checked) & _
                                                    Convert.ToByte(CB_SPCIncomeNum.Checked) & _
                                                    Convert.ToByte(CB_SPCOutgoNum.Checked) & _
                                                    Convert.ToByte(CB_SPCPassNum.Checked) & _
                                                    Convert.ToByte(CB_SPCPartFailedNum.Checked) & _
                                                    Convert.ToByte(CB_SPCTotalFailedNum.Checked) & _
                                                    Convert.ToByte(CB_SPCTimePerBoard.Checked) & _
                                                    Convert.ToByte(CB_SPCBoardPerHour.Checked) & _
                                                    Convert.ToByte(CB_SPCRunTime.Checked) & _
                                                    Convert.ToByte(CB_SPCDownTime.Checked) & _
                                                    Convert.ToByte(CB_SPCPauseNum.Checked) & _
                                                    Convert.ToByte(CB_SPCVolCalNum.Checked) & _
                                                    Convert.ToByte(CB_SPCNeedleCalNum.Checked)

        IDS.Data.SaveData()
    End Sub

    Private Sub ButtonRevert_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonRevert.Click
        RevertData()
    End Sub

    Public Sub RevertData()
        'IDS.Data.OpenData()
        If IDS.Data.Hardware.SPC.ItemsToBeReported.Length < 2 Then
            IDS.Data.Hardware.SPC.ItemsToBeReported = "0000000000000"
        End If

        Nud_EventCleanInterval.Text = IDS.Data.Hardware.SPC.CleanUpInterval
        CB_EnableSPCLog.Checked = IDS.Data.Hardware.SPC.EnableSPCReport
        'CB_ReportDirectory.Text = IDS.Data.Hardware.SPC.ReportFileName
        tbReportDirectory.Text = IDS.Data.Hardware.SPC.ReportFileName
        CB_SPCMaterialBatch.Checked = Convert.ToByte(IDS.Data.Hardware.SPC.ItemsToBeReported.Chars(0)) - 48
        CB_SPCIncomeNum.Checked = Convert.ToByte(IDS.Data.Hardware.SPC.ItemsToBeReported.Chars(1)) - 48
        CB_SPCOutgoNum.Checked = Convert.ToByte(IDS.Data.Hardware.SPC.ItemsToBeReported.Chars(2)) - 48
        CB_SPCPassNum.Checked = Convert.ToByte(IDS.Data.Hardware.SPC.ItemsToBeReported.Chars(3)) - 48
        CB_SPCPartFailedNum.Checked = Convert.ToByte(IDS.Data.Hardware.SPC.ItemsToBeReported.Chars(4)) - 48
        CB_SPCTotalFailedNum.Checked = Convert.ToByte(IDS.Data.Hardware.SPC.ItemsToBeReported.Chars(5)) - 48
        CB_SPCTimePerBoard.Checked = Convert.ToByte(IDS.Data.Hardware.SPC.ItemsToBeReported.Chars(6)) - 48
        CB_SPCBoardPerHour.Checked = Convert.ToByte(IDS.Data.Hardware.SPC.ItemsToBeReported.Chars(7)) - 48
        CB_SPCRunTime.Checked = Convert.ToByte(IDS.Data.Hardware.SPC.ItemsToBeReported.Chars(8)) - 48
        CB_SPCDownTime.Checked = Convert.ToByte(IDS.Data.Hardware.SPC.ItemsToBeReported.Chars(9)) - 48
        CB_SPCPauseNum.Checked = Convert.ToByte(IDS.Data.Hardware.SPC.ItemsToBeReported.Chars(10)) - 48
        CB_SPCVolCalNum.Checked = Convert.ToByte(IDS.Data.Hardware.SPC.ItemsToBeReported.Chars(11)) - 48
        CB_SPCNeedleCalNum.Checked = Convert.ToByte(IDS.Data.Hardware.SPC.ItemsToBeReported.Chars(12)) - 48

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        OpenFileDialog1.InitialDirectory = "c:\"
        OpenFileDialog1.Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*"
        OpenFileDialog1.FilterIndex = 0
        If OpenFileDialog1.ShowDialog = DialogResult.OK Then
            'CB_ReportDirectory.Text = OpenFileDialog1.FileName
            tbReportDirectory.Text = IDS.Data.Hardware.SPC.ReportFileName
        End If
    End Sub

    Private Sub ButtonExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonExit.Click
        RemovePanel(CurrentControl)
    End Sub
End Class
