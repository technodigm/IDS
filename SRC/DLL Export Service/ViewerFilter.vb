'Imports System.Data.OleDb

Public Class ViewerFilter
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
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents CheckBox1 As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox2 As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox3 As System.Windows.Forms.CheckBox
    Friend WithEvents ComboBox1 As System.Windows.Forms.ComboBox
    Friend WithEvents ComboBox2 As System.Windows.Forms.ComboBox
    Friend WithEvents ComboBox3 As System.Windows.Forms.ComboBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents CheckBox6 As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox7 As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox8 As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox9 As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox12 As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox13 As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox14 As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox15 As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox16 As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox17 As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox20 As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox21 As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox23 As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox22 As System.Windows.Forms.CheckBox
    Friend WithEvents DateTimePicker1 As System.Windows.Forms.DateTimePicker
    Friend WithEvents DateTimePicker2 As System.Windows.Forms.DateTimePicker
    Friend WithEvents RadioButton1 As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton2 As System.Windows.Forms.RadioButton
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents CheckBox19 As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox11 As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox18 As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox10 As System.Windows.Forms.CheckBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.CheckBox23 = New System.Windows.Forms.CheckBox
        Me.CheckBox3 = New System.Windows.Forms.CheckBox
        Me.CheckBox2 = New System.Windows.Forms.CheckBox
        Me.CheckBox1 = New System.Windows.Forms.CheckBox
        Me.ComboBox1 = New System.Windows.Forms.ComboBox
        Me.ComboBox2 = New System.Windows.Forms.ComboBox
        Me.ComboBox3 = New System.Windows.Forms.ComboBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label6 = New System.Windows.Forms.Label
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.RadioButton2 = New System.Windows.Forms.RadioButton
        Me.RadioButton1 = New System.Windows.Forms.RadioButton
        Me.DateTimePicker2 = New System.Windows.Forms.DateTimePicker
        Me.DateTimePicker1 = New System.Windows.Forms.DateTimePicker
        Me.Button1 = New System.Windows.Forms.Button
        Me.Button2 = New System.Windows.Forms.Button
        Me.GroupBox3 = New System.Windows.Forms.GroupBox
        Me.CheckBox22 = New System.Windows.Forms.CheckBox
        Me.CheckBox21 = New System.Windows.Forms.CheckBox
        Me.CheckBox20 = New System.Windows.Forms.CheckBox
        Me.CheckBox19 = New System.Windows.Forms.CheckBox
        Me.CheckBox18 = New System.Windows.Forms.CheckBox
        Me.CheckBox17 = New System.Windows.Forms.CheckBox
        Me.CheckBox16 = New System.Windows.Forms.CheckBox
        Me.CheckBox15 = New System.Windows.Forms.CheckBox
        Me.CheckBox14 = New System.Windows.Forms.CheckBox
        Me.CheckBox13 = New System.Windows.Forms.CheckBox
        Me.CheckBox12 = New System.Windows.Forms.CheckBox
        Me.CheckBox11 = New System.Windows.Forms.CheckBox
        Me.CheckBox10 = New System.Windows.Forms.CheckBox
        Me.CheckBox9 = New System.Windows.Forms.CheckBox
        Me.CheckBox8 = New System.Windows.Forms.CheckBox
        Me.CheckBox7 = New System.Windows.Forms.CheckBox
        Me.CheckBox6 = New System.Windows.Forms.CheckBox
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.CheckBox23)
        Me.GroupBox1.Controls.Add(Me.CheckBox3)
        Me.GroupBox1.Controls.Add(Me.CheckBox2)
        Me.GroupBox1.Controls.Add(Me.CheckBox1)
        Me.GroupBox1.Location = New System.Drawing.Point(16, 15)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(259, 89)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Type"
        '
        'CheckBox23
        '
        Me.CheckBox23.Checked = True
        Me.CheckBox23.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CheckBox23.Location = New System.Drawing.Point(136, 56)
        Me.CheckBox23.Name = "CheckBox23"
        Me.CheckBox23.Size = New System.Drawing.Size(96, 23)
        Me.CheckBox23.TabIndex = 3
        Me.CheckBox23.Text = "Administration"
        '
        'CheckBox3
        '
        Me.CheckBox3.Checked = True
        Me.CheckBox3.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CheckBox3.Location = New System.Drawing.Point(40, 56)
        Me.CheckBox3.Name = "CheckBox3"
        Me.CheckBox3.Size = New System.Drawing.Size(66, 23)
        Me.CheckBox3.TabIndex = 2
        Me.CheckBox3.Text = "Internal"
        '
        'CheckBox2
        '
        Me.CheckBox2.Checked = True
        Me.CheckBox2.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CheckBox2.Location = New System.Drawing.Point(136, 22)
        Me.CheckBox2.Name = "CheckBox2"
        Me.CheckBox2.Size = New System.Drawing.Size(87, 23)
        Me.CheckBox2.TabIndex = 1
        Me.CheckBox2.Text = "Information"
        '
        'CheckBox1
        '
        Me.CheckBox1.Checked = True
        Me.CheckBox1.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CheckBox1.Location = New System.Drawing.Point(40, 22)
        Me.CheckBox1.Name = "CheckBox1"
        Me.CheckBox1.Size = New System.Drawing.Size(87, 23)
        Me.CheckBox1.TabIndex = 0
        Me.CheckBox1.Text = "Alarm"
        '
        'ComboBox1
        '
        Me.ComboBox1.Location = New System.Drawing.Point(152, 120)
        Me.ComboBox1.Name = "ComboBox1"
        Me.ComboBox1.Size = New System.Drawing.Size(101, 21)
        Me.ComboBox1.TabIndex = 1
        Me.ComboBox1.Text = "(ALL)"
        '
        'ComboBox2
        '
        Me.ComboBox2.Location = New System.Drawing.Point(152, 152)
        Me.ComboBox2.Name = "ComboBox2"
        Me.ComboBox2.Size = New System.Drawing.Size(101, 21)
        Me.ComboBox2.TabIndex = 2
        Me.ComboBox2.Text = "(ALL)"
        '
        'ComboBox3
        '
        Me.ComboBox3.Location = New System.Drawing.Point(152, 184)
        Me.ComboBox3.Name = "ComboBox3"
        Me.ComboBox3.Size = New System.Drawing.Size(101, 21)
        Me.ComboBox3.TabIndex = 3
        Me.ComboBox3.Text = "(ALL)"
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(32, 120)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(67, 21)
        Me.Label1.TabIndex = 4
        Me.Label1.Text = "User ID"
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(32, 152)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(96, 22)
        Me.Label2.TabIndex = 5
        Me.Label2.Text = "Pattern File Name"
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(32, 184)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(67, 21)
        Me.Label3.TabIndex = 6
        Me.Label3.Text = "SPC Name"
        '
        'Label6
        '
        Me.Label6.Location = New System.Drawing.Point(122, 64)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(24, 22)
        Me.Label6.TabIndex = 12
        Me.Label6.Text = "To"
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.Label5)
        Me.GroupBox2.Controls.Add(Me.Label4)
        Me.GroupBox2.Controls.Add(Me.RadioButton2)
        Me.GroupBox2.Controls.Add(Me.RadioButton1)
        Me.GroupBox2.Controls.Add(Me.DateTimePicker2)
        Me.GroupBox2.Controls.Add(Me.Label6)
        Me.GroupBox2.Controls.Add(Me.DateTimePicker1)
        Me.GroupBox2.Location = New System.Drawing.Point(16, 216)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(259, 112)
        Me.GroupBox2.TabIndex = 16
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Date"
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(144, 84)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(48, 16)
        Me.Label5.TabIndex = 25
        Me.Label5.Text = "23:59:59"
        Me.Label5.Visible = False
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(32, 84)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(48, 16)
        Me.Label4.TabIndex = 24
        Me.Label4.Text = "00:00:00"
        Me.Label4.Visible = False
        '
        'RadioButton2
        '
        Me.RadioButton2.Location = New System.Drawing.Point(136, 24)
        Me.RadioButton2.Name = "RadioButton2"
        Me.RadioButton2.Size = New System.Drawing.Size(80, 24)
        Me.RadioButton2.TabIndex = 23
        Me.RadioButton2.Text = "Set Period"
        '
        'RadioButton1
        '
        Me.RadioButton1.Checked = True
        Me.RadioButton1.Location = New System.Drawing.Point(56, 24)
        Me.RadioButton1.Name = "RadioButton1"
        Me.RadioButton1.Size = New System.Drawing.Size(72, 24)
        Me.RadioButton1.TabIndex = 22
        Me.RadioButton1.TabStop = True
        Me.RadioButton1.Text = "All Dates"
        '
        'DateTimePicker2
        '
        Me.DateTimePicker2.CustomFormat = "dd/MM/yyyy"
        Me.DateTimePicker2.Enabled = False
        Me.DateTimePicker2.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DateTimePicker2.Location = New System.Drawing.Point(144, 60)
        Me.DateTimePicker2.MaxDate = New Date(2100, 12, 31, 23, 59, 59, 999)
        Me.DateTimePicker2.MinDate = New Date(2006, 1, 1, 0, 0, 0, 0)
        Me.DateTimePicker2.Name = "DateTimePicker2"
        Me.DateTimePicker2.Size = New System.Drawing.Size(88, 20)
        Me.DateTimePicker2.TabIndex = 21
        Me.DateTimePicker2.Value = New Date(2006, 6, 9, 0, 0, 0, 0)
        '
        'DateTimePicker1
        '
        Me.DateTimePicker1.CustomFormat = "dd/MM/yyyy"
        Me.DateTimePicker1.Enabled = False
        Me.DateTimePicker1.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DateTimePicker1.Location = New System.Drawing.Point(32, 60)
        Me.DateTimePicker1.MaxDate = New Date(2100, 12, 31, 23, 59, 59, 999)
        Me.DateTimePicker1.MinDate = New Date(2006, 1, 1, 0, 0, 0, 0)
        Me.DateTimePicker1.Name = "DateTimePicker1"
        Me.DateTimePicker1.Size = New System.Drawing.Size(88, 20)
        Me.DateTimePicker1.TabIndex = 20
        Me.DateTimePicker1.Value = New Date(2006, 6, 9, 0, 0, 0, 0)
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(400, 344)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(62, 22)
        Me.Button1.TabIndex = 17
        Me.Button1.Text = "OK"
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(480, 344)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(62, 22)
        Me.Button2.TabIndex = 18
        Me.Button2.Text = "Cancel"
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.CheckBox22)
        Me.GroupBox3.Controls.Add(Me.CheckBox21)
        Me.GroupBox3.Controls.Add(Me.CheckBox20)
        Me.GroupBox3.Controls.Add(Me.CheckBox19)
        Me.GroupBox3.Controls.Add(Me.CheckBox18)
        Me.GroupBox3.Controls.Add(Me.CheckBox17)
        Me.GroupBox3.Controls.Add(Me.CheckBox16)
        Me.GroupBox3.Controls.Add(Me.CheckBox15)
        Me.GroupBox3.Controls.Add(Me.CheckBox14)
        Me.GroupBox3.Controls.Add(Me.CheckBox13)
        Me.GroupBox3.Controls.Add(Me.CheckBox12)
        Me.GroupBox3.Controls.Add(Me.CheckBox11)
        Me.GroupBox3.Controls.Add(Me.CheckBox10)
        Me.GroupBox3.Controls.Add(Me.CheckBox9)
        Me.GroupBox3.Controls.Add(Me.CheckBox8)
        Me.GroupBox3.Controls.Add(Me.CheckBox7)
        Me.GroupBox3.Controls.Add(Me.CheckBox6)
        Me.GroupBox3.Location = New System.Drawing.Point(287, 15)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(260, 313)
        Me.GroupBox3.TabIndex = 19
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Column"
        '
        'CheckBox22
        '
        Me.CheckBox22.Checked = True
        Me.CheckBox22.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CheckBox22.Location = New System.Drawing.Point(13, 16)
        Me.CheckBox22.Name = "CheckBox22"
        Me.CheckBox22.Size = New System.Drawing.Size(43, 23)
        Me.CheckBox22.TabIndex = 16
        Me.CheckBox22.Text = "All"
        Me.CheckBox22.Visible = False
        '
        'CheckBox21
        '
        Me.CheckBox21.Location = New System.Drawing.Point(147, 272)
        Me.CheckBox21.Name = "CheckBox21"
        Me.CheckBox21.Size = New System.Drawing.Size(93, 22)
        Me.CheckBox21.TabIndex = 15
        Me.CheckBox21.Text = "SPC Name"
        '
        'CheckBox20
        '
        Me.CheckBox20.Location = New System.Drawing.Point(147, 224)
        Me.CheckBox20.Name = "CheckBox20"
        Me.CheckBox20.Size = New System.Drawing.Size(93, 23)
        Me.CheckBox20.TabIndex = 14
        Me.CheckBox20.Text = "Vol Cal Flag"
        '
        'CheckBox19
        '
        Me.CheckBox19.Location = New System.Drawing.Point(120, 282)
        Me.CheckBox19.Name = "CheckBox19"
        Me.CheckBox19.Size = New System.Drawing.Size(92, 23)
        Me.CheckBox19.TabIndex = 13
        Me.CheckBox19.Text = "Down Time"
        Me.CheckBox19.Visible = False
        '
        'CheckBox18
        '
        Me.CheckBox18.Location = New System.Drawing.Point(152, 8)
        Me.CheckBox18.Name = "CheckBox18"
        Me.CheckBox18.Size = New System.Drawing.Size(93, 22)
        Me.CheckBox18.TabIndex = 12
        Me.CheckBox18.Text = "Boards/h"
        Me.CheckBox18.Visible = False
        '
        'CheckBox17
        '
        Me.CheckBox17.Location = New System.Drawing.Point(147, 176)
        Me.CheckBox17.Name = "CheckBox17"
        Me.CheckBox17.Size = New System.Drawing.Size(93, 22)
        Me.CheckBox17.TabIndex = 11
        Me.CheckBox17.Text = "Failed Flag t"
        '
        'CheckBox16
        '
        Me.CheckBox16.Location = New System.Drawing.Point(147, 128)
        Me.CheckBox16.Name = "CheckBox16"
        Me.CheckBox16.Size = New System.Drawing.Size(93, 23)
        Me.CheckBox16.TabIndex = 10
        Me.CheckBox16.Text = "Pass Flag"
        '
        'CheckBox15
        '
        Me.CheckBox15.Location = New System.Drawing.Point(147, 80)
        Me.CheckBox15.Name = "CheckBox15"
        Me.CheckBox15.Size = New System.Drawing.Size(93, 22)
        Me.CheckBox15.TabIndex = 9
        Me.CheckBox15.Text = "Income Flag"
        '
        'CheckBox14
        '
        Me.CheckBox14.Location = New System.Drawing.Point(147, 32)
        Me.CheckBox14.Name = "CheckBox14"
        Me.CheckBox14.Size = New System.Drawing.Size(101, 22)
        Me.CheckBox14.TabIndex = 8
        Me.CheckBox14.Text = "Material Batch"
        '
        'CheckBox13
        '
        Me.CheckBox13.Location = New System.Drawing.Point(33, 272)
        Me.CheckBox13.Name = "CheckBox13"
        Me.CheckBox13.Size = New System.Drawing.Size(111, 22)
        Me.CheckBox13.TabIndex = 7
        Me.CheckBox13.Text = "Needle Cal Flag"
        '
        'CheckBox12
        '
        Me.CheckBox12.Location = New System.Drawing.Point(33, 224)
        Me.CheckBox12.Name = "CheckBox12"
        Me.CheckBox12.Size = New System.Drawing.Size(87, 23)
        Me.CheckBox12.TabIndex = 6
        Me.CheckBox12.Text = "Pause Flag"
        '
        'CheckBox11
        '
        Me.CheckBox11.Location = New System.Drawing.Point(7, 288)
        Me.CheckBox11.Name = "CheckBox11"
        Me.CheckBox11.Size = New System.Drawing.Size(87, 22)
        Me.CheckBox11.TabIndex = 5
        Me.CheckBox11.Text = "Run Time"
        Me.CheckBox11.Visible = False
        '
        'CheckBox10
        '
        Me.CheckBox10.Location = New System.Drawing.Point(64, 8)
        Me.CheckBox10.Name = "CheckBox10"
        Me.CheckBox10.Size = New System.Drawing.Size(87, 22)
        Me.CheckBox10.TabIndex = 4
        Me.CheckBox10.Text = "Time/B"
        Me.CheckBox10.Visible = False
        '
        'CheckBox9
        '
        Me.CheckBox9.Location = New System.Drawing.Point(33, 176)
        Me.CheckBox9.Name = "CheckBox9"
        Me.CheckBox9.Size = New System.Drawing.Size(95, 22)
        Me.CheckBox9.TabIndex = 3
        Me.CheckBox9.Text = "Failed Flag p"
        '
        'CheckBox8
        '
        Me.CheckBox8.Location = New System.Drawing.Point(33, 128)
        Me.CheckBox8.Name = "CheckBox8"
        Me.CheckBox8.Size = New System.Drawing.Size(87, 23)
        Me.CheckBox8.TabIndex = 2
        Me.CheckBox8.Text = "Outgo Flag"
        '
        'CheckBox7
        '
        Me.CheckBox7.Location = New System.Drawing.Point(33, 80)
        Me.CheckBox7.Name = "CheckBox7"
        Me.CheckBox7.Size = New System.Drawing.Size(87, 22)
        Me.CheckBox7.TabIndex = 1
        Me.CheckBox7.Text = "Type"
        '
        'CheckBox6
        '
        Me.CheckBox6.Location = New System.Drawing.Point(33, 32)
        Me.CheckBox6.Name = "CheckBox6"
        Me.CheckBox6.Size = New System.Drawing.Size(95, 22)
        Me.CheckBox6.TabIndex = 0
        Me.CheckBox6.Text = "Pattern Name"
        '
        'ViewerFilter
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(567, 381)
        Me.Controls.Add(Me.GroupBox3)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.ComboBox3)
        Me.Controls.Add(Me.ComboBox2)
        Me.Controls.Add(Me.ComboBox1)
        Me.Controls.Add(Me.GroupBox1)
        Me.Name = "ViewerFilter"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Event Viewer Filter"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox3.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region
#Region "shared val define"
    Shared type_alarm As Boolean = True
    Shared type_infor As Boolean = True
    Shared type_inter As Boolean = False
    Shared type_admin As Boolean = False

    Shared select_period As Boolean = False
    Shared old_select As Boolean = False

    Shared old_date1 As Date = DateTime.Today
    Shared old_date2 As Date = DateTime.Today

    'Shared all As Boolean = True
    Shared PatternName As Boolean = True
    Shared MaterialBatch As Boolean = False
    Shared Type As Boolean = True
    Shared IncomeFl As Boolean = False
    Shared OutgoFl As Boolean = False
    Shared PassFl As Boolean = False
    Shared FailedFlp As Boolean = False
    Shared FailedFlt As Boolean = False
    Shared TimeB As Boolean = False
    Shared Boardsh As Boolean = False
    Shared RunTime As Boolean = False
    Shared DownTime As Boolean = False
    Shared PauseFl As Boolean = False
    Shared VolcalFl As Boolean = False
    Shared NeedleCalFl As Boolean = False
    Shared SPCName As Boolean = True

    Shared UserID_sle As String = "*"
    Shared PatternName_sle As String = "*"
    Shared SPCNAME_sle As String = "*"
    Shared combo1_item As Integer = 0
    Shared combo2_item As Integer = 0
    Shared combo3_item As Integer = 0

    Public Shared combo1_max As Integer = 0
    Public Shared combo2_max As Integer = 0
    Public Shared combo3_max As Integer = 0

#End Region
    Public Shared SQLCmdStr As String = "SELECT ID, EvtID, EvtName, [Time], Source, [User/Group], PatternName, Type, SPCName, UserAction FROM Log " & _
                                        "WHERE (Type <> 'Internal') AND (Type <> 'Administration') ORDER BY ID"

    'Public Declare Sub Sleep Lib "kernel32" Alias "Sleep" (ByVal dwMilliseconds As Long)  'delay function
    Private Sub Button1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        'Sleep(1000)
        If CheckBox1.Checked Then : type_alarm = True : Else : type_alarm = False : End If
        If CheckBox2.Checked Then : type_infor = True : Else : type_infor = False : End If
        If CheckBox3.Checked Then : type_inter = True : Else : type_inter = False : End If
        If CheckBox23.Checked Then : type_admin = True : Else : type_admin = False : End If
        'If RadioButton1.Checked Then : select_period = False : Else : select_period = True : End If
        'If RadioButton2.Checked Then
        If CheckBox6.Checked Then : PatternName = True : Else : PatternName = False : End If
        If CheckBox14.Checked Then : MaterialBatch = True : Else : MaterialBatch = False : End If
        If CheckBox7.Checked Then : Type = True : Else : Type = False : End If
        If CheckBox15.Checked Then : IncomeFl = True : Else : IncomeFl = False : End If
        If CheckBox8.Checked Then : OutgoFl = True : Else : OutgoFl = False : End If
        If CheckBox16.Checked Then : PassFl = True : Else : PassFl = False : End If
        If CheckBox9.Checked Then : FailedFlp = True : Else : FailedFlp = False : End If
        If CheckBox17.Checked Then : FailedFlt = True : Else : FailedFlt = False : End If
        If CheckBox10.Checked Then : TimeB = True : Else : TimeB = False : End If
        If CheckBox18.Checked Then : Boardsh = True : Else : Boardsh = False : End If
        If CheckBox11.Checked Then : RunTime = True : Else : RunTime = False : End If
        If CheckBox19.Checked Then : DownTime = True : Else : DownTime = False : End If
        If CheckBox12.Checked Then : PauseFl = True : Else : PauseFl = False : End If
        If CheckBox20.Checked Then : VolcalFl = True : Else : VolcalFl = False : End If
        If CheckBox13.Checked Then : NeedleCalFl = True : Else : NeedleCalFl = False : End If
        If CheckBox21.Checked Then : SPCName = True : Else : SPCName = False : End If

        UserID_sle = cstr(ComboBox1.SelectedValue)
        PatternName_sle = cstr(ComboBox2.SelectedValue)
        SPCNAME_sle = cstr(ComboBox3.SelectedValue)

        combo1_item = ComboBox1.SelectedIndex
        combo2_item = ComboBox2.SelectedIndex
        combo3_item = ComboBox3.SelectedIndex

        If combo1_item = 0 Then
            UserID_sle = "*"
        End If
        If combo2_item = 0 Then
            PatternName_sle = "*"
        End If
        If combo3_item = 0 Then
            SPCNAME_sle = "*"
        End If
        'SELECT ID, EvtID, EvtName, [Time], Source, [User/Group], PatternName, MaterialBatch, 
        'Type, [in#], [out#], [pass#], [Fp#], [Ft#], [time/B], [B/h], RunTime, DownTime, Pause, 
        'VolCal(, NeedleCal, SPCName, UserAction)
        'FROM(Log)
        'ORDER BY ID
        'if
        GetNewSQLCmd()
        old_date1 = DateTimePicker1.Value
        old_date2 = DateTimePicker2.Value
        Close()
    End Sub

    Private Sub GetNewSQLCmd()
        SQLCmdStr = "SELECT ID, EvtID, EvtName, [Time], Source, [User/Group], "
        If PatternName Then : SQLCmdStr &= "PatternName, " : End If
        If MaterialBatch Then : SQLCmdStr &= "MaterialBatch, " : End If
        If Type Then : SQLCmdStr &= "Type, " : End If
        If IncomeFl Then : SQLCmdStr &= "[in#], " : End If
        If OutgoFl Then : SQLCmdStr &= "[out#], " : End If
        If PassFl Then : SQLCmdStr &= "[pass#], " : End If
        If FailedFlp Then : SQLCmdStr &= "[Fp#], " : End If
        If FailedFlt Then : SQLCmdStr &= "[Ft#], " : End If
        If TimeB Then : SQLCmdStr &= "[time/B], " : End If
        If Boardsh Then : SQLCmdStr &= "[B/h], " : End If
        If RunTime Then : SQLCmdStr &= "RunTime, " : End If
        If DownTime Then : SQLCmdStr &= "DownTime, " : End If
        If PauseFl Then : SQLCmdStr &= "Pause, " : End If
        If VolcalFl Then : SQLCmdStr &= "VolCal, " : End If
        If NeedleCalFl Then : SQLCmdStr &= "NeedleCal, " : End If
        If SPCName Then : SQLCmdStr &= "SPCName, " : End If
        SQLCmdStr &= "UserAction FROM Log"

        If UserID_sle <> "*" Then : SQLCmdStr &= " WHERE([User/Group] = '" & UserID_sle & "')" : End If
        If UserID_sle = "*" And PatternName_sle <> "*" Then : SQLCmdStr &= " WHERE(PatternName = '" & PatternName_sle & "')"
        ElseIf UserID_sle <> "*" And PatternName_sle <> "*" Then : SQLCmdStr &= "AND(PatternName = '" & PatternName_sle & "')"
        End If
        If (UserID_sle <> "*" Or PatternName_sle <> "*") And SPCNAME_sle <> "*" Then
            SQLCmdStr &= "AND(SPCName = '" & SPCNAME_sle & "')"
        ElseIf SPCNAME_sle <> "*" Then
            SQLCmdStr &= " WHERE(SPCName = '" & SPCNAME_sle & "')"
        End If

        Dim where_flag As Boolean = False

        If UserID_sle <> "*" Or SPCNAME_sle <> "*" Or PatternName_sle <> "*" Then
            where_flag = True
            If type_alarm = False Then : SQLCmdStr &= " AND (Type <> 'Alarm')" : End If
            If type_inter = False Then : SQLCmdStr &= " AND (Type <> 'Internal')" : End If
            If type_infor = False Then : SQLCmdStr &= " AND (Type <> 'Information')" : End If
            If type_admin = False Then : SQLCmdStr &= " AND (Type <> 'Administration')" : End If
        Else
            If type_alarm = False Then : where_flag = True : SQLCmdStr &= " WHERE (Type <> 'Alarm')" : End If
            If type_alarm = False And type_inter = False Then : where_flag = True : SQLCmdStr &= "AND (Type <> 'Internal')"
            ElseIf type_inter = False Then : where_flag = True : SQLCmdStr &= " WHERE (Type <> 'Internal')" : End If
            If (type_alarm = False Or type_inter = False) And type_infor = False Then : where_flag = True : SQLCmdStr &= "AND (Type <> 'Information')"
            ElseIf type_infor = False Then : where_flag = True : SQLCmdStr &= " WHERE (Type <> 'Information')" : End If
            If (type_alarm = False Or type_inter = False Or type_infor = False) And type_admin = False Then
                where_flag = True : SQLCmdStr &= "AND (Type <> 'Administration')"
            ElseIf type_admin = False Then : where_flag = True : SQLCmdStr &= " WHERE (Type <> 'Administration')" : End If
        End If

        If select_period Then
            Dim time_start_str As String = DateTimePicker1.Value.Month & "/" & DateTimePicker1.Value.Day & "/" & DateTimePicker1.Value.Year & " 00:00:00"
            Dim time_end_str As String = DateTimePicker2.Value.Month & "/" & DateTimePicker2.Value.Day & "/" & DateTimePicker2.Value.Year & " 23:59:59"
            'upper line is to: Format(time_start_str, "MM/dd/yyyy HH:mm:ss")
            'upper line is to: Format(time_end_str, "MM/dd/yyyy HH:mm:ss")
            If where_flag Then
                SQLCmdStr &= " AND ([Time] >= #" & time_start_str & "#) AND ([Time] <= #" & time_end_str & "#)"
            Else
                SQLCmdStr &= " WHERE ([Time] >= #" & time_start_str & "#) AND ([Time] <= #" & time_end_str & "#)"
            End If

        End If

        SQLCmdStr &= " ORDER BY ID"
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        select_period = old_select 'for radiobutton set back value
        Close()  'without changings
    End Sub

    Private Sub ViewerFilter_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        If type_alarm Then : CheckBox1.Checked = True : Else : CheckBox1.Checked = False : End If
        If type_infor Then : CheckBox2.Checked = True : Else : CheckBox2.Checked = False : End If
        If type_inter Then : CheckBox3.Checked = True : Else : CheckBox3.Checked = False : End If
        If type_admin Then : CheckBox23.Checked = True : Else : CheckBox23.Checked = False : End If

        old_select = select_period
        If select_period Then
            DateTimePicker1.Enabled = True
            DateTimePicker2.Enabled = True
            RadioButton1.Checked = False
            RadioButton2.Checked = True
            Label4.Visible = True
            Label5.Visible = True
        Else
            DateTimePicker1.Enabled = False
            DateTimePicker2.Enabled = False
            RadioButton2.Checked = False
            RadioButton1.Checked = True
            Label4.Visible = False
            Label5.Visible = False
        End If

        If PatternName Then : CheckBox6.Checked = True : Else : CheckBox6.Checked = False : End If
        If MaterialBatch Then : CheckBox14.Checked = True : Else : CheckBox14.Checked = False : End If
        If Type Then : CheckBox7.Checked = True : Else : CheckBox7.Checked = False : End If
        If IncomeFl Then : CheckBox15.Checked = True : Else : CheckBox15.Checked = False : End If
        If OutgoFl Then : CheckBox8.Checked = True : Else : CheckBox8.Checked = False : End If
        If PassFl Then : CheckBox16.Checked = True : Else : CheckBox16.Checked = False : End If
        If FailedFlp Then : CheckBox9.Checked = True : Else : CheckBox9.Checked = False : End If
        If FailedFlt Then : CheckBox17.Checked = True : Else : CheckBox17.Checked = False : End If
        If TimeB Then : CheckBox10.Checked = True : Else : CheckBox10.Checked = False : End If
        If Boardsh Then : CheckBox18.Checked = True : Else : CheckBox18.Checked = False : End If
        If RunTime Then : CheckBox11.Checked = True : Else : CheckBox11.Checked = False : End If
        If DownTime Then : CheckBox19.Checked = True : Else : CheckBox19.Checked = False : End If
        If PauseFl Then : CheckBox12.Checked = True : Else : CheckBox12.Checked = False : End If
        If VolcalFl Then : CheckBox20.Checked = True : Else : CheckBox20.Checked = False : End If
        If NeedleCalFl Then : CheckBox13.Checked = True : Else : CheckBox13.Checked = False : End If
        If SPCName Then : CheckBox21.Checked = True : Else : CheckBox21.Checked = False : End If

        If combo1_item <= combo1_max Then
            ComboBox1.SelectedIndex = combo1_item
        Else
            ComboBox1.SelectedIndex = combo1_max
        End If

        If combo3_item <= combo3_max Then
            ComboBox3.SelectedIndex = combo3_item
        Else
            ComboBox3.SelectedIndex = combo3_max
        End If

        If combo2_item <= combo2_max Then
            ComboBox2.SelectedIndex = combo2_item
        Else
            ComboBox2.SelectedIndex = combo2_max
        End If
        'Try : ComboBox2.SelectedIndex = combo2_item : Catch e1 As ArgumentOutOfRangeException : ComboBox2.SelectedIndex = 0 : End Try
        'Try : ComboBox3.SelectedIndex = combo3_item : Catch e1 As ArgumentOutOfRangeException : ComboBox3.SelectedIndex = 0 : End Try
        DateTimePicker1.Value = old_date1
        DateTimePicker2.Value = old_date2
    End Sub

    Private Sub RadioButton2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton2.CheckedChanged
        If RadioButton2.Checked Then
            select_period = True
            DateTimePicker1.Enabled = True
            DateTimePicker2.Enabled = True
            Label4.Visible = True
            Label5.Visible = True
        Else
            select_period = False
            DateTimePicker1.Enabled = False
            DateTimePicker2.Enabled = False
            Label4.Visible = False
            Label5.Visible = False
        End If
    End Sub

End Class
