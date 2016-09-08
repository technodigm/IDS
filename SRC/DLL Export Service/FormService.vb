'Imports events_execution
Imports System.IO

Public Class FormService
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
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox2 As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents TextBox3 As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents CheckBox1 As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox2 As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox3 As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox4 As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox5 As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox6 As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox7 As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox8 As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox9 As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox10 As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox11 As System.Windows.Forms.CheckBox
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents Button4 As System.Windows.Forms.Button
    Friend WithEvents TextBox4 As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents CheckBox12 As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox13 As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox14 As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox15 As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox16 As System.Windows.Forms.CheckBox
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents Timer1 As System.Windows.Forms.Timer
    Friend WithEvents CheckBox17 As System.Windows.Forms.CheckBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Me.Button1 = New System.Windows.Forms.Button
        Me.TextBox1 = New System.Windows.Forms.TextBox
        Me.TextBox2 = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.TextBox3 = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.CheckBox1 = New System.Windows.Forms.CheckBox
        Me.CheckBox2 = New System.Windows.Forms.CheckBox
        Me.CheckBox3 = New System.Windows.Forms.CheckBox
        Me.CheckBox4 = New System.Windows.Forms.CheckBox
        Me.CheckBox5 = New System.Windows.Forms.CheckBox
        Me.CheckBox6 = New System.Windows.Forms.CheckBox
        Me.CheckBox7 = New System.Windows.Forms.CheckBox
        Me.CheckBox8 = New System.Windows.Forms.CheckBox
        Me.CheckBox9 = New System.Windows.Forms.CheckBox
        Me.CheckBox10 = New System.Windows.Forms.CheckBox
        Me.CheckBox11 = New System.Windows.Forms.CheckBox
        Me.Button2 = New System.Windows.Forms.Button
        Me.Button4 = New System.Windows.Forms.Button
        Me.TextBox4 = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.CheckBox17 = New System.Windows.Forms.CheckBox
        Me.CheckBox12 = New System.Windows.Forms.CheckBox
        Me.CheckBox13 = New System.Windows.Forms.CheckBox
        Me.CheckBox14 = New System.Windows.Forms.CheckBox
        Me.CheckBox15 = New System.Windows.Forms.CheckBox
        Me.CheckBox16 = New System.Windows.Forms.CheckBox
        Me.Button3 = New System.Windows.Forms.Button
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(69, 229)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(72, 40)
        Me.Button1.TabIndex = 0
        Me.Button1.Text = "Input Event (Integer)"
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(24, 238)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(34, 20)
        Me.TextBox1.TabIndex = 1
        Me.TextBox1.Text = "1"
        '
        'TextBox2
        '
        Me.TextBox2.Location = New System.Drawing.Point(273, 37)
        Me.TextBox2.Name = "TextBox2"
        Me.TextBox2.Size = New System.Drawing.Size(47, 26)
        Me.TextBox2.TabIndex = 2
        Me.TextBox2.Text = ""
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("SimSun", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label1.Location = New System.Drawing.Point(208, 40)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(64, 15)
        Me.Label1.TabIndex = 3
        Me.Label1.Text = "Returned:"
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("SimSun", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label2.Location = New System.Drawing.Point(13, 7)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(435, 16)
        Me.Label2.TabIndex = 4
        Me.Label2.Text = "SPC Report Setting for The Current Pattern File:"
        '
        'TextBox3
        '
        Me.TextBox3.Location = New System.Drawing.Point(107, 67)
        Me.TextBox3.Name = "TextBox3"
        Me.TextBox3.Size = New System.Drawing.Size(280, 20)
        Me.TextBox3.TabIndex = 5
        Me.TextBox3.Text = "e:/IDS/event_db/event_db/SPC_Report_1.txt"
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(13, 37)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(99, 16)
        Me.Label3.TabIndex = 6
        Me.Label3.Text = "Pattern File Name"
        '
        'CheckBox1
        '
        Me.CheckBox1.Checked = True
        Me.CheckBox1.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CheckBox1.Location = New System.Drawing.Point(393, 52)
        Me.CheckBox1.Name = "CheckBox1"
        Me.CheckBox1.Size = New System.Drawing.Size(119, 45)
        Me.CheckBox1.TabIndex = 7
        Me.CheckBox1.Text = "Append new report to this SPC file"
        '
        'CheckBox2
        '
        Me.CheckBox2.Checked = True
        Me.CheckBox2.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CheckBox2.Location = New System.Drawing.Point(13, 119)
        Me.CheckBox2.Name = "CheckBox2"
        Me.CheckBox2.Size = New System.Drawing.Size(60, 23)
        Me.CheckBox2.TabIndex = 8
        Me.CheckBox2.Text = "UserID"
        '
        'CheckBox3
        '
        Me.CheckBox3.Checked = True
        Me.CheckBox3.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CheckBox3.Location = New System.Drawing.Point(93, 119)
        Me.CheckBox3.Name = "CheckBox3"
        Me.CheckBox3.Size = New System.Drawing.Size(99, 23)
        Me.CheckBox3.TabIndex = 9
        Me.CheckBox3.Text = "Pattn File"
        '
        'CheckBox4
        '
        Me.CheckBox4.Checked = True
        Me.CheckBox4.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CheckBox4.Location = New System.Drawing.Point(180, 119)
        Me.CheckBox4.Name = "CheckBox4"
        Me.CheckBox4.Size = New System.Drawing.Size(87, 23)
        Me.CheckBox4.TabIndex = 10
        Me.CheckBox4.Text = "Material"
        '
        'CheckBox5
        '
        Me.CheckBox5.Checked = True
        Me.CheckBox5.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CheckBox5.Location = New System.Drawing.Point(267, 119)
        Me.CheckBox5.Name = "CheckBox5"
        Me.CheckBox5.Size = New System.Drawing.Size(80, 23)
        Me.CheckBox5.TabIndex = 11
        Me.CheckBox5.Text = "#Incoming"
        '
        'CheckBox6
        '
        Me.CheckBox6.Checked = True
        Me.CheckBox6.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CheckBox6.Location = New System.Drawing.Point(347, 119)
        Me.CheckBox6.Name = "CheckBox6"
        Me.CheckBox6.Size = New System.Drawing.Size(80, 23)
        Me.CheckBox6.TabIndex = 12
        Me.CheckBox6.Text = "#Outgoing"
        '
        'CheckBox7
        '
        Me.CheckBox7.Checked = True
        Me.CheckBox7.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CheckBox7.Location = New System.Drawing.Point(13, 141)
        Me.CheckBox7.Name = "CheckBox7"
        Me.CheckBox7.Size = New System.Drawing.Size(75, 24)
        Me.CheckBox7.TabIndex = 13
        Me.CheckBox7.Text = "#Pass"
        '
        'CheckBox8
        '
        Me.CheckBox8.Checked = True
        Me.CheckBox8.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CheckBox8.Location = New System.Drawing.Point(93, 141)
        Me.CheckBox8.Name = "CheckBox8"
        Me.CheckBox8.Size = New System.Drawing.Size(87, 24)
        Me.CheckBox8.TabIndex = 14
        Me.CheckBox8.Text = "#Failed(P)"
        '
        'CheckBox9
        '
        Me.CheckBox9.Checked = True
        Me.CheckBox9.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CheckBox9.Location = New System.Drawing.Point(180, 141)
        Me.CheckBox9.Name = "CheckBox9"
        Me.CheckBox9.Size = New System.Drawing.Size(87, 24)
        Me.CheckBox9.TabIndex = 15
        Me.CheckBox9.Text = "#Failed(T)"
        '
        'CheckBox10
        '
        Me.CheckBox10.Checked = True
        Me.CheckBox10.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CheckBox10.Location = New System.Drawing.Point(267, 141)
        Me.CheckBox10.Name = "CheckBox10"
        Me.CheckBox10.Size = New System.Drawing.Size(73, 24)
        Me.CheckBox10.TabIndex = 16
        Me.CheckBox10.Text = "Time/B"
        '
        'CheckBox11
        '
        Me.CheckBox11.Checked = True
        Me.CheckBox11.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CheckBox11.Location = New System.Drawing.Point(347, 141)
        Me.CheckBox11.Name = "CheckBox11"
        Me.CheckBox11.Size = New System.Drawing.Size(69, 24)
        Me.CheckBox11.TabIndex = 17
        Me.CheckBox11.Text = "Boards/h"
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(440, 149)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(67, 37)
        Me.Button2.TabIndex = 18
        Me.Button2.Text = "Setting OK"
        '
        'Button4
        '
        Me.Button4.Location = New System.Drawing.Point(440, 232)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(67, 42)
        Me.Button4.TabIndex = 22
        Me.Button4.Text = "Event Viewer"
        '
        'TextBox4
        '
        Me.TextBox4.Location = New System.Drawing.Point(107, 37)
        Me.TextBox4.Name = "TextBox4"
        Me.TextBox4.Size = New System.Drawing.Size(133, 20)
        Me.TextBox4.TabIndex = 23
        Me.TextBox4.Text = "Pattern_1.pat"
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(13, 67)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(91, 16)
        Me.Label4.TabIndex = 24
        Me.Label4.Text = "SPC Path, Name"
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(13, 97)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(180, 15)
        Me.Label5.TabIndex = 25
        Me.Label5.Text = "Items to be reported in the SPC:"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.CheckBox17)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Controls.Add(Me.TextBox2)
        Me.GroupBox1.Font = New System.Drawing.Font("SimSun", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.GroupBox1.Location = New System.Drawing.Point(16, 201)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(328, 74)
        Me.GroupBox1.TabIndex = 26
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Emulate An Event"
        '
        'CheckBox17
        '
        Me.CheckBox17.Font = New System.Drawing.Font("SimSun", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.CheckBox17.Location = New System.Drawing.Point(136, 40)
        Me.CheckBox17.Name = "CheckBox17"
        Me.CheckBox17.Size = New System.Drawing.Size(48, 16)
        Me.CheckBox17.TabIndex = 4
        Me.CheckBox17.Text = "Auto"
        '
        'CheckBox12
        '
        Me.CheckBox12.Checked = True
        Me.CheckBox12.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CheckBox12.Location = New System.Drawing.Point(347, 163)
        Me.CheckBox12.Name = "CheckBox12"
        Me.CheckBox12.Size = New System.Drawing.Size(86, 25)
        Me.CheckBox12.TabIndex = 31
        Me.CheckBox12.Text = "#Needle.Cal"
        '
        'CheckBox13
        '
        Me.CheckBox13.Checked = True
        Me.CheckBox13.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CheckBox13.Location = New System.Drawing.Point(267, 163)
        Me.CheckBox13.Name = "CheckBox13"
        Me.CheckBox13.Size = New System.Drawing.Size(73, 25)
        Me.CheckBox13.TabIndex = 30
        Me.CheckBox13.Text = "#Vol.Cal"
        '
        'CheckBox14
        '
        Me.CheckBox14.Checked = True
        Me.CheckBox14.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CheckBox14.Location = New System.Drawing.Point(180, 163)
        Me.CheckBox14.Name = "CheckBox14"
        Me.CheckBox14.Size = New System.Drawing.Size(67, 25)
        Me.CheckBox14.TabIndex = 29
        Me.CheckBox14.Text = "#pause"
        '
        'CheckBox15
        '
        Me.CheckBox15.Checked = True
        Me.CheckBox15.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CheckBox15.Location = New System.Drawing.Point(93, 163)
        Me.CheckBox15.Name = "CheckBox15"
        Me.CheckBox15.Size = New System.Drawing.Size(80, 25)
        Me.CheckBox15.TabIndex = 28
        Me.CheckBox15.Text = "down time"
        '
        'CheckBox16
        '
        Me.CheckBox16.Checked = True
        Me.CheckBox16.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CheckBox16.Location = New System.Drawing.Point(13, 163)
        Me.CheckBox16.Name = "CheckBox16"
        Me.CheckBox16.Size = New System.Drawing.Size(75, 25)
        Me.CheckBox16.TabIndex = 27
        Me.CheckBox16.Text = "Run time"
        '
        'Button3
        '
        Me.Button3.Location = New System.Drawing.Point(360, 232)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(67, 42)
        Me.Button3.TabIndex = 32
        Me.Button3.Text = "Generate SPC Report"
        '
        'Timer1
        '
        '
        'FormService
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(519, 285)
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.CheckBox12)
        Me.Controls.Add(Me.CheckBox13)
        Me.Controls.Add(Me.CheckBox14)
        Me.Controls.Add(Me.CheckBox15)
        Me.Controls.Add(Me.CheckBox16)
        Me.Controls.Add(Me.TextBox4)
        Me.Controls.Add(Me.CheckBox11)
        Me.Controls.Add(Me.CheckBox10)
        Me.Controls.Add(Me.CheckBox9)
        Me.Controls.Add(Me.CheckBox8)
        Me.Controls.Add(Me.CheckBox7)
        Me.Controls.Add(Me.CheckBox6)
        Me.Controls.Add(Me.CheckBox5)
        Me.Controls.Add(Me.CheckBox4)
        Me.Controls.Add(Me.CheckBox3)
        Me.Controls.Add(Me.CheckBox2)
        Me.Controls.Add(Me.CheckBox1)
        Me.Controls.Add(Me.TextBox3)
        Me.Controls.Add(Me.TextBox1)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Button4)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.GroupBox1)
        Me.Name = "FormService"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "EventComes"
        Me.GroupBox1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Public Function NoActionToExecute()
        If CInt(IDS.__ID) Mod 100 = 0 Then Return True
        Return False
    End Function

    Public Sub ResetEventCode()
        IDS.__ID = "0"
    End Sub

    'only for this project. inaccessible outside other project, keep it this way so that any mutation to ids.__id is within this project only
    Friend Sub SetEventCode(ByVal ID As String)
        IDS.__ID = ID
    End Sub

    Public Function VolCalPressedContinue() As Boolean
        If IDS.__ID = "1076206" Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Function VolCalPressedStop() As Boolean
        If IDS.__ID = "1036206" Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Function VolCalPressedRedo() As Boolean
        If IDS.__ID = "1156206" Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Function EStopPressedOk() As Boolean
        If IDS.__ID = "1016205" Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Function DotSizeCheckContinue() As Boolean
        If IDS.__ID = "1076204" Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Function DotSizeCheckStop() As Boolean
        If IDS.__ID = "1036204" Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Function HeightSkipped() As Boolean
        If IDS.__ID = "1026202" Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Function ChipEdgeSkipped() As Boolean
        If IDS.__ID = "1056203" Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Function FiducialSkipped() As Boolean
        If IDS.__ID = "1026201" Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Sub DisplayErrorMessage(ByVal ErrorType As String)
        Try
            'Form_Execution.Hide()
            Select Case ErrorType
                Case "Width Adjustment Failed"
                    IDS.FrmExecution.DisplayErrorPopup("1003201")
                Case "Lifter 1 Blocked"
                    IDS.FrmExecution.DisplayErrorPopup("1003202")
                Case "Lifter 2 Blocked"
                    IDS.FrmExecution.DisplayErrorPopup("1003203")
                Case "Lifter 3 Blocked"
                    IDS.FrmExecution.DisplayErrorPopup("1003204")
                Case "Retrieve Timeout"
                    IDS.FrmExecution.DisplayErrorPopup("1003205")
                Case "Stage 1 Travel Timeout"
                    IDS.FrmExecution.DisplayErrorPopup("1003206")
                Case "Release Travel Timeout"
                    IDS.FrmExecution.DisplayErrorPopup("1003207")
                Case "Stage 3 Travel Timeout"
                    IDS.FrmExecution.DisplayErrorPopup("1003208")
                Case "Board Not Aligned"
                    IDS.FrmExecution.DisplayErrorPopup("1003209")
                Case "PLC Communication Error"
                    IDS.FrmExecution.DisplayErrorPopup("1003210")
                Case "PLC No Signal"
                    IDS.FrmExecution.DisplayErrorPopup("1003211")
                Case "Low Incoming Air Pressure"
                    IDS.FrmExecution.DisplayErrorPopup("1003212")
                Case "Fiducial Failed"
                    IDS.FrmExecution.DisplayErrorPopup("1006201")
                Case "Height Point Failed"
                    IDS.FrmExecution.DisplayErrorPopup("1006202")
                Case "Chip Finding Failed"
                    IDS.FrmExecution.DisplayErrorPopup("1006203")
                Case "Dot Size Check Failed"
                    IDS.FrmExecution.DisplayErrorPopup("1006204")
                Case "E-Stop"
                    IDS.FrmExecution.DisplayErrorPopup("1006205")
                Case "Camera Signal Failed"
                    IDS.FrmExecution.DisplayErrorPopup("1006299")
                Case "Volume Calibration Failed"
                    IDS.FrmExecution.DisplayErrorPopup("1006206")
            End Select
        Catch ex As Exception
            ExceptionDisplay(ex)
        End Try
    End Sub

    Public Sub LogEventInSPCReport(ByVal EventName As String)
        Dim ID As String
        Select Case EventName
            Case "Motion controller error"

            Case "Production Start"
                IDS.FrmExecution.LogAnEvent("1007401")
            Case "Board Comes"
                IDS.FrmExecution.LogAnEvent("1007402")
            Case "Board Goes"
                IDS.FrmExecution.LogAnEvent("1007403")
            Case "Pause"
                IDS.FrmExecution.LogAnEvent("1007404")
            Case "Resume"
                IDS.FrmExecution.LogAnEvent("1007405")
            Case "Needle Calibration"
                IDS.FrmExecution.LogAnEvent("1007406")
            Case "Volume Calibration"
                'IDS.FrmExecution.LogAnEvent("1007407")
            Case "Board Partial Failure"
                IDS.FrmExecution.LogAnEvent("1007408")
            Case "Board Total Failure"
                IDS.FrmExecution.LogAnEvent("1007409")
            Case "Production End"
                IDS.FrmExecution.LogAnEvent("1007410")
            Case "Purge And Clean"
                IDS.FrmExecution.LogAnEvent("1007411")
            Case "Dot Size Check Passed"
                IDS.FrmExecution.LogAnEvent("1007412")
            Case "Volume Calibration Passed"
                IDS.FrmExecution.LogAnEvent("1007413")
            Case "Volume Calibration Failed"
                IDS.FrmExecution.LogAnEvent("1006206")

            Case "Fiducial Failed"
                IDS.FrmExecution.LogAnEvent("1006201")
            Case "Height Point Failed"
                IDS.FrmExecution.LogAnEvent("1006202")
            Case "Chip Finding Failed"
                IDS.FrmExecution.LogAnEvent("1006203")
            Case "Dot Size Check Failed"
                IDS.FrmExecution.LogAnEvent("1006204")
            Case "E-Stop"
                IDS.FrmExecution.LogAnEvent("1006205")
            Case "Camera Signal Failed"
                IDS.FrmExecution.LogAnEvent("1006299")
        End Select
    End Sub

    Public Function GenerateSPCReport()
        Dim Report As New SPCReport
        Report.generate(0)  'mode 0, to auto generate SPC for the last production
    End Function

    'Later, these settings should be gotten from pattern file. No need to write them here.
#Region " Write Log Settings "
    Public Sub WriteLogSettings()
        Dim Writer As New StreamWriter("c:\IDS\BIN\LogSettings.txt", False, System.Text.Encoding.Default)
        Dim content As String

        If CheckBox2.Checked = True Then
            content = "\1"
        Else
            content = "\-1"
        End If
        If CheckBox3.Checked = True Then
            content &= "	\1"
        Else
            content &= "	\-1"
        End If
        If CheckBox4.Checked = True Then
            content &= "	\1"
        Else
            content &= "	\-1"
        End If
        If CheckBox5.Checked = True Then
            content &= "	\1"
        Else
            content &= "	\-1"
        End If
        If CheckBox6.Checked = True Then
            content &= "	\1"
        Else
            content &= "	\-1"
        End If
        If CheckBox7.Checked = True Then
            content &= "	\1"
        Else
            content &= "	\-1"
        End If
        If CheckBox8.Checked = True Then
            content &= "	\1"
        Else
            content &= "	\-1"
        End If
        If CheckBox9.Checked = True Then
            content &= "	\1"
        Else
            content &= "	\-1"
        End If
        If CheckBox10.Checked = True Then
            content &= "	\1"
        Else
            content &= "	\-1"
        End If
        If CheckBox11.Checked = True Then
            content &= "	\1"
        Else
            content &= "	\-1"
        End If
        If CheckBox16.Checked = True Then
            content &= "	\1"
        Else
            content &= "	\-1"
        End If
        If CheckBox15.Checked = True Then
            content &= "	\1"
        Else
            content &= "	\-1"
        End If
        If CheckBox14.Checked = True Then
            content &= "	\1"
        Else
            content &= "	\-1"
        End If
        If CheckBox13.Checked = True Then
            content &= "	\1"
        Else
            content &= "	\-1"
        End If
        If CheckBox12.Checked = True Then
            content &= "	\1"
        Else
            content &= "	\-1"
        End If
        If CheckBox1.Checked = True Then
            content &= "	\a	" & TextBox3.Text & "	" & TextBox4.Text
        Else
            content &= "	\-a	" & TextBox3.Text & "	" & TextBox4.Text
        End If
        Writer.Write(content)
        Writer.Close()
    End Sub
#End Region

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        'TextBox2.Text = LogAnEvent(cint(TextBox1.Text))
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        WriteLogSettings()
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        'DataGrid1.Show()
        Dim evtviewer As New EventViewer
        evtviewer.Show()
    End Sub

#Region "Old form_load, adapter.fill method"
    'Private Sub FormService_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
    'DataGrid1.Hide()
    'OleDbDataAdapter2.Fill(DataSet21)
    'End Sub
#End Region
    'Dim cont As Integer
#Region "old button3, database operations"
    'Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    'Dim OleDbCom As OleDbCommand
    ''Dim Conn As OleDbConnection
    ''OleDbConnection1.Execute('Select * Into abcd From [Text;Database=c:\temp].aaaa.txt');
    'OleDbCom = New OleDbCommand("Select * Into OleDbDataAdapter1 From [Text;Database=e:/].LogSettings.txt", OleDbConnection1)

    'OleDbDataAdapter2.Fill(DataSet21)


    ''cont += 1

    'Dim command1 As OleDbCommand = New OleDbCommand("SELECT * FROM Log")
    'command1.CommandType = CommandType.Text
    'OleDbConnection1.Open()
    'command1.Connection = OleDbConnection1
    'OleDbDataAdapter2.SelectCommand = command1


    'Dim Builder As OleDbCommandBuilder = New OleDbCommandBuilder(OleDbDataAdapter2)

    ''Dim drow As System.data.DataRow = DataSet11.Table1.NewRow()

    ''drow(1) = cstr(cont)
    ''drow(0) = cstr(cont + 1)

    'If DataSet11.Log.Rows(0)("Field10") = "1" Then
    'DataSet11.Log.Columns.Remove("Field1")
    'End If

    ''DataSet11.Table1.Rows.Add(drow)

    ''OleDbDataAdapter1.Update(DataSet11)
    'OleDbConnection1.Close()

    'UpdateData("SELECT * FROM DBTable", "DBTable")

    'End Sub

    '    Friend Function UpdateData(ByRef SQLStr As String, ByRef TableStr As String)
    '       Try
    '          'Conn.Open()
    '
    '           'Dim OleDbCom As OleDbCommand = New OleDbCommand(SQLStr, Conn)
    '          'OleDbCom.CommandTimeout = TIMEOUT_VALUE
    '         'Adapter.SelectCommand = OleDbCom
    '
    '           Builder = New OleDbCommandBuilder(Adapter)
    '          Dim DsChanges As DataSet = DS.GetChanges()
    '
    '           If Not DsChanges Is Nothing Then
    '              Adapter.Update(DsChanges, TableStr)
    '             DS.AcceptChanges()
    '        End If
    '
    '       Catch ex As Exception
    '          MessageBox.Show(ex.ToString)
    '     End Try
    '    Conn.Close()
    '
    '   End Function
#End Region

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        'to generate SPC Report
        GenerateSPCReport()
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        'TextBox2.Text = LogAnEvent(cint(TextBox1.Text))
    End Sub

    Private Sub CheckBox17_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox17.CheckedChanged
        If CheckBox17.Checked = False Then
            Timer1.Stop()
        Else
            Timer1.Start()
        End If
    End Sub

End Class
