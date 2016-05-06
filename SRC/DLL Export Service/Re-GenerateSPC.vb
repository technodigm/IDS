Imports System.Data.OleDb
Imports System.IO
Public Class Re_GenerateSPC
    Inherits System.Windows.Forms.Form

#Region " Windows 窗体设计器生成的代码 "

    Public Sub New()
        MyBase.New()

        '该调用是 Windows 窗体设计器所必需的。
        InitializeComponent()

        '在 InitializeComponent() 调用之后添加任何初始化

    End Sub

    '窗体重写 dispose 以清理组件列表。
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Windows 窗体设计器所必需的
    Private components As System.ComponentModel.IContainer

    '注意: 以下过程是 Windows 窗体设计器所必需的
    '可以使用 Windows 窗体设计器修改此过程。
    '不要使用代码编辑器修改它。
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents ListBox1 As System.Windows.Forms.ListBox
    Friend WithEvents RadioButton2 As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton1 As System.Windows.Forms.RadioButton
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents OleDbDataAdapter1 As System.Data.OleDb.OleDbDataAdapter
    Friend WithEvents OleDbSelectCommand1 As System.Data.OleDb.OleDbCommand
    Friend WithEvents OleDbInsertCommand1 As System.Data.OleDb.OleDbCommand
    Friend WithEvents OleDbUpdateCommand1 As System.Data.OleDb.OleDbCommand
    Friend WithEvents OleDbDeleteCommand1 As System.Data.OleDb.OleDbCommand
    Friend WithEvents OleDbConnection1 As System.Data.OleDb.OleDbConnection
    Friend WithEvents SaveFileDialog1 As System.Windows.Forms.SaveFileDialog
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.Button2 = New System.Windows.Forms.Button
        Me.Button1 = New System.Windows.Forms.Button
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.ListBox1 = New System.Windows.Forms.ListBox
        Me.RadioButton2 = New System.Windows.Forms.RadioButton
        Me.RadioButton1 = New System.Windows.Forms.RadioButton
        Me.Label3 = New System.Windows.Forms.Label
        Me.TextBox1 = New System.Windows.Forms.TextBox
        Me.Button3 = New System.Windows.Forms.Button
        Me.OleDbDataAdapter1 = New System.Data.OleDb.OleDbDataAdapter
        Me.OleDbDeleteCommand1 = New System.Data.OleDb.OleDbCommand
        Me.OleDbConnection1 = New System.Data.OleDb.OleDbConnection
        Me.OleDbInsertCommand1 = New System.Data.OleDb.OleDbCommand
        Me.OleDbSelectCommand1 = New System.Data.OleDb.OleDbCommand
        Me.OleDbUpdateCommand1 = New System.Data.OleDb.OleDbCommand
        Me.SaveFileDialog1 = New System.Windows.Forms.SaveFileDialog
        Me.SuspendLayout()
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(232, 392)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(62, 24)
        Me.Button2.TabIndex = 14
        Me.Button2.Text = "Cancel"
        '
        'Button1
        '
        Me.Button1.Enabled = False
        Me.Button1.Location = New System.Drawing.Point(152, 392)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(62, 24)
        Me.Button1.TabIndex = 13
        Me.Button1.Text = "Generate"
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(13, 17)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(267, 18)
        Me.Label2.TabIndex = 12
        Me.Label2.Text = "Re-generate SPC report from the events belong to"
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(16, 80)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(254, 15)
        Me.Label1.TabIndex = 11
        Me.Label1.Text = "Candidate items, for which to generate report:"
        '
        'ListBox1
        '
        Me.ListBox1.Location = New System.Drawing.Point(12, 104)
        Me.ListBox1.Name = "ListBox1"
        Me.ListBox1.Size = New System.Drawing.Size(280, 225)
        Me.ListBox1.TabIndex = 10
        '
        'RadioButton2
        '
        Me.RadioButton2.Location = New System.Drawing.Point(120, 39)
        Me.RadioButton2.Name = "RadioButton2"
        Me.RadioButton2.Size = New System.Drawing.Size(72, 20)
        Me.RadioButton2.TabIndex = 9
        Me.RadioButton2.Text = "A Pattern"
        '
        'RadioButton1
        '
        Me.RadioButton1.Location = New System.Drawing.Point(20, 39)
        Me.RadioButton1.Name = "RadioButton1"
        Me.RadioButton1.Size = New System.Drawing.Size(100, 20)
        Me.RadioButton1.TabIndex = 8
        Me.RadioButton1.Text = "An SPC report"
        '
        'Label3
        '
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(8, 344)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(35, 21)
        Me.Label3.TabIndex = 15
        Me.Label3.Text = "Path:"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'TextBox1
        '
        Me.TextBox1.BackColor = System.Drawing.Color.White
        Me.TextBox1.Location = New System.Drawing.Point(48, 344)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.ReadOnly = True
        Me.TextBox1.Size = New System.Drawing.Size(186, 20)
        Me.TextBox1.TabIndex = 16
        Me.TextBox1.Text = "Please Input"
        '
        'Button3
        '
        Me.Button3.Location = New System.Drawing.Point(248, 344)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(46, 21)
        Me.Button3.TabIndex = 17
        Me.Button3.Text = "Browse"
        '
        'OleDbDataAdapter1
        '
        Me.OleDbDataAdapter1.DeleteCommand = Me.OleDbDeleteCommand1
        Me.OleDbDataAdapter1.InsertCommand = Me.OleDbInsertCommand1
        Me.OleDbDataAdapter1.SelectCommand = Me.OleDbSelectCommand1
        Me.OleDbDataAdapter1.TableMappings.AddRange(New System.Data.Common.DataTableMapping() {New System.Data.Common.DataTableMapping("Table", "Log", New System.Data.Common.DataColumnMapping() {New System.Data.Common.DataColumnMapping("B/h", "B/h"), New System.Data.Common.DataColumnMapping("DownTime", "DownTime"), New System.Data.Common.DataColumnMapping("EvtID", "EvtID"), New System.Data.Common.DataColumnMapping("EvtName", "EvtName"), New System.Data.Common.DataColumnMapping("Fp#", "Fp#"), New System.Data.Common.DataColumnMapping("Ft#", "Ft#"), New System.Data.Common.DataColumnMapping("ID", "ID"), New System.Data.Common.DataColumnMapping("in#", "in#"), New System.Data.Common.DataColumnMapping("MaterialBatch", "MaterialBatch"), New System.Data.Common.DataColumnMapping("NeedleCal", "NeedleCal"), New System.Data.Common.DataColumnMapping("out#", "out#"), New System.Data.Common.DataColumnMapping("pass#", "pass#"), New System.Data.Common.DataColumnMapping("PatternName", "PatternName"), New System.Data.Common.DataColumnMapping("Pause", "Pause"), New System.Data.Common.DataColumnMapping("RunTime", "RunTime"), New System.Data.Common.DataColumnMapping("Source", "Source"), New System.Data.Common.DataColumnMapping("SPCName", "SPCName"), New System.Data.Common.DataColumnMapping("Time", "Time"), New System.Data.Common.DataColumnMapping("time/B", "time/B"), New System.Data.Common.DataColumnMapping("Type", "Type"), New System.Data.Common.DataColumnMapping("User/Group", "User/Group"), New System.Data.Common.DataColumnMapping("UserAction", "UserAction"), New System.Data.Common.DataColumnMapping("VolCal", "VolCal")})})
        Me.OleDbDataAdapter1.UpdateCommand = Me.OleDbUpdateCommand1
        '
        'OleDbDeleteCommand1
        '
        Me.OleDbDeleteCommand1.CommandText = "DELETE FROM Log WHERE (ID = ?) AND ([B/h] = ? OR ? IS NULL AND [B/h] IS NULL) AND" & _
        " (DownTime = ? OR ? IS NULL AND DownTime IS NULL) AND (EvtID = ? OR ? IS NULL AN" & _
        "D EvtID IS NULL) AND (EvtName = ? OR ? IS NULL AND EvtName IS NULL) AND ([Fp#] =" & _
        " ? OR ? IS NULL AND [Fp#] IS NULL) AND ([Ft#] = ? OR ? IS NULL AND [Ft#] IS NULL" & _
        ") AND (MaterialBatch = ? OR ? IS NULL AND MaterialBatch IS NULL) AND (NeedleCal " & _
        "= ? OR ? IS NULL AND NeedleCal IS NULL) AND (PatternName = ? OR ? IS NULL AND Pa" & _
        "tternName IS NULL) AND (Pause = ? OR ? IS NULL AND Pause IS NULL) AND (RunTime =" & _
        " ? OR ? IS NULL AND RunTime IS NULL) AND (SPCName = ? OR ? IS NULL AND SPCName I" & _
        "S NULL) AND (Source = ? OR ? IS NULL AND Source IS NULL) AND ([Time] = ? OR ? IS" & _
        " NULL AND [Time] IS NULL) AND (Type = ? OR ? IS NULL AND Type IS NULL) AND ([Use" & _
        "r/Group] = ? OR ? IS NULL AND [User/Group] IS NULL) AND (UserAction = ? OR ? IS " & _
        "NULL AND UserAction IS NULL) AND (VolCal = ? OR ? IS NULL AND VolCal IS NULL) AN" & _
        "D ([in#] = ? OR ? IS NULL AND [in#] IS NULL) AND ([out#] = ? OR ? IS NULL AND [o" & _
        "ut#] IS NULL) AND ([pass#] = ? OR ? IS NULL AND [pass#] IS NULL) AND ([time/B] =" & _
        " ? OR ? IS NULL AND [time/B] IS NULL)"
        Me.OleDbDeleteCommand1.Connection = Me.OleDbConnection1
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("OriginalIDS.__ID", System.Data.OleDb.OleDbType.Integer, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "ID", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_B_h", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "B/h", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_B_h1", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "B/h", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_DownTime", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "DownTime", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_DownTime1", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "DownTime", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_EvtID", System.Data.OleDb.OleDbType.Integer, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "EvtID", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_EvtID1", System.Data.OleDb.OleDbType.Integer, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "EvtID", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_EvtName", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "EvtName", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_EvtName1", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "EvtName", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Fp_", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Fp#", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Fp_1", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Fp#", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Ft_", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Ft#", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Ft_1", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Ft#", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_MaterialBatch", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "MaterialBatch", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_MaterialBatch1", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "MaterialBatch", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_NeedleCal", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "NeedleCal", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_NeedleCal1", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "NeedleCal", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_PatternName", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "PatternName", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_PatternName1", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "PatternName", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Pause", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Pause", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Pause1", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Pause", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_RunTime", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "RunTime", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_RunTime1", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "RunTime", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_SPCName", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "SPCName", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_SPCName1", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "SPCName", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Source", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Source", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Source1", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Source", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Time", System.Data.OleDb.OleDbType.DBDate, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Time", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Time1", System.Data.OleDb.OleDbType.DBDate, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Time", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Type", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Type", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Type1", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Type", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_User_Group", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "User/Group", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_User_Group1", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "User/Group", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_UserAction", System.Data.OleDb.OleDbType.VarWChar, 255, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "UserAction", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_UserAction1", System.Data.OleDb.OleDbType.VarWChar, 255, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "UserAction", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_VolCal", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "VolCal", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_VolCal1", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "VolCal", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_in_", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "in#", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_in_1", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "in#", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_out_", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "out#", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_out_1", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "out#", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_pass_", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "pass#", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_pass_1", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "pass#", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_time_B", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "time/B", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_time_B1", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "time/B", System.Data.DataRowVersion.Original, Nothing))
        '
        'OleDbConnection1
        '
        Me.OleDbConnection1.ConnectionString = "Jet OLEDB:Global Partial Bulk Ops=2;Jet OLEDB:Registry Path=;Jet OLEDB:Database L" & _
        "ocking Mode=1;Data Source=""C:\IDS\bin\evt_db_setting.mdb"";Jet OLEDB:Engine Type=" & _
        "5;Provider=""Microsoft.Jet.OLEDB.4.0"";Jet OLEDB:System database=;Jet OLEDB:SFP=Fa" & _
        "lse;persist security info=False;Extended Properties=;Mode=Share Deny None;Jet OL" & _
        "EDB:Encrypt Database=False;Jet OLEDB:Create System Database=False;Jet OLEDB:Don'" & _
        "t Copy Locale on Compact=False;Jet OLEDB:Compact Without Replica Repair=False;Us" & _
        "er ID=Admin;Jet OLEDB:Global Bulk Transactions=1"
        '
        'OleDbInsertCommand1
        '
        Me.OleDbInsertCommand1.CommandText = "INSERT INTO Log([B/h], DownTime, EvtID, EvtName, [Fp#], [Ft#], [in#], MaterialBat" & _
        "ch, NeedleCal, [out#], [pass#], PatternName, Pause, RunTime, Source, SPCName, [T" & _
        "ime], [time/B], Type, [User/Group], UserAction, VolCal) VALUES (?, ?, ?, ?, ?, ?" & _
        ", ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"
        Me.OleDbInsertCommand1.Connection = Me.OleDbConnection1
        Me.OleDbInsertCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("B_h", System.Data.OleDb.OleDbType.VarWChar, 50, "B/h"))
        Me.OleDbInsertCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("DownTime", System.Data.OleDb.OleDbType.VarWChar, 50, "DownTime"))
        Me.OleDbInsertCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("EvtID", System.Data.OleDb.OleDbType.Integer, 0, "EvtID"))
        Me.OleDbInsertCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("EvtName", System.Data.OleDb.OleDbType.VarWChar, 50, "EvtName"))
        Me.OleDbInsertCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Fp_", System.Data.OleDb.OleDbType.VarWChar, 50, "Fp#"))
        Me.OleDbInsertCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Ft_", System.Data.OleDb.OleDbType.VarWChar, 50, "Ft#"))
        Me.OleDbInsertCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("in_", System.Data.OleDb.OleDbType.VarWChar, 50, "in#"))
        Me.OleDbInsertCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("MaterialBatch", System.Data.OleDb.OleDbType.VarWChar, 50, "MaterialBatch"))
        Me.OleDbInsertCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("NeedleCal", System.Data.OleDb.OleDbType.VarWChar, 50, "NeedleCal"))
        Me.OleDbInsertCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("out_", System.Data.OleDb.OleDbType.VarWChar, 50, "out#"))
        Me.OleDbInsertCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("pass_", System.Data.OleDb.OleDbType.VarWChar, 50, "pass#"))
        Me.OleDbInsertCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("PatternName", System.Data.OleDb.OleDbType.VarWChar, 50, "PatternName"))
        Me.OleDbInsertCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Pause", System.Data.OleDb.OleDbType.VarWChar, 50, "Pause"))
        Me.OleDbInsertCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("RunTime", System.Data.OleDb.OleDbType.VarWChar, 50, "RunTime"))
        Me.OleDbInsertCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Source", System.Data.OleDb.OleDbType.VarWChar, 50, "Source"))
        Me.OleDbInsertCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("SPCName", System.Data.OleDb.OleDbType.VarWChar, 50, "SPCName"))
        Me.OleDbInsertCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Time", System.Data.OleDb.OleDbType.DBDate, 0, "Time"))
        Me.OleDbInsertCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("time_B", System.Data.OleDb.OleDbType.VarWChar, 50, "time/B"))
        Me.OleDbInsertCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Type", System.Data.OleDb.OleDbType.VarWChar, 50, "Type"))
        Me.OleDbInsertCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("User_Group", System.Data.OleDb.OleDbType.VarWChar, 50, "User/Group"))
        Me.OleDbInsertCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("UserAction", System.Data.OleDb.OleDbType.VarWChar, 255, "UserAction"))
        Me.OleDbInsertCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("VolCal", System.Data.OleDb.OleDbType.VarWChar, 50, "VolCal"))
        '
        'OleDbSelectCommand1
        '
        Me.OleDbSelectCommand1.CommandText = "SELECT [B/h], DownTime, EvtID, EvtName, [Fp#], [Ft#], ID, [in#], MaterialBatch, N" & _
        "eedleCal, [out#], [pass#], PatternName, Pause, RunTime, Source, SPCName, [Time]," & _
        " [time/B], Type, [User/Group], UserAction, VolCal FROM Log"
        Me.OleDbSelectCommand1.Connection = Me.OleDbConnection1
        '
        'OleDbUpdateCommand1
        '
        Me.OleDbUpdateCommand1.CommandText = "UPDATE Log SET [B/h] = ?, DownTime = ?, EvtID = ?, EvtName = ?, [Fp#] = ?, [Ft#] " & _
        "= ?, [in#] = ?, MaterialBatch = ?, NeedleCal = ?, [out#] = ?, [pass#] = ?, Patte" & _
        "rnName = ?, Pause = ?, RunTime = ?, Source = ?, SPCName = ?, [Time] = ?, [time/B" & _
        "] = ?, Type = ?, [User/Group] = ?, UserAction = ?, VolCal = ? WHERE (ID = ?) AND" & _
        " ([B/h] = ? OR ? IS NULL AND [B/h] IS NULL) AND (DownTime = ? OR ? IS NULL AND D" & _
        "ownTime IS NULL) AND (EvtID = ? OR ? IS NULL AND EvtID IS NULL) AND (EvtName = ?" & _
        " OR ? IS NULL AND EvtName IS NULL) AND ([Fp#] = ? OR ? IS NULL AND [Fp#] IS NULL" & _
        ") AND ([Ft#] = ? OR ? IS NULL AND [Ft#] IS NULL) AND (MaterialBatch = ? OR ? IS " & _
        "NULL AND MaterialBatch IS NULL) AND (NeedleCal = ? OR ? IS NULL AND NeedleCal IS" & _
        " NULL) AND (PatternName = ? OR ? IS NULL AND PatternName IS NULL) AND (Pause = ?" & _
        " OR ? IS NULL AND Pause IS NULL) AND (RunTime = ? OR ? IS NULL AND RunTime IS NU" & _
        "LL) AND (SPCName = ? OR ? IS NULL AND SPCName IS NULL) AND (Source = ? OR ? IS N" & _
        "ULL AND Source IS NULL) AND ([Time] = ? OR ? IS NULL AND [Time] IS NULL) AND (Ty" & _
        "pe = ? OR ? IS NULL AND Type IS NULL) AND ([User/Group] = ? OR ? IS NULL AND [Us" & _
        "er/Group] IS NULL) AND (UserAction = ? OR ? IS NULL AND UserAction IS NULL) AND " & _
        "(VolCal = ? OR ? IS NULL AND VolCal IS NULL) AND ([in#] = ? OR ? IS NULL AND [in" & _
        "#] IS NULL) AND ([out#] = ? OR ? IS NULL AND [out#] IS NULL) AND ([pass#] = ? OR" & _
        " ? IS NULL AND [pass#] IS NULL) AND ([time/B] = ? OR ? IS NULL AND [time/B] IS N" & _
        "ULL)"
        Me.OleDbUpdateCommand1.Connection = Me.OleDbConnection1
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("B_h", System.Data.OleDb.OleDbType.VarWChar, 50, "B/h"))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("DownTime", System.Data.OleDb.OleDbType.VarWChar, 50, "DownTime"))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("EvtID", System.Data.OleDb.OleDbType.Integer, 0, "EvtID"))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("EvtName", System.Data.OleDb.OleDbType.VarWChar, 50, "EvtName"))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Fp_", System.Data.OleDb.OleDbType.VarWChar, 50, "Fp#"))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Ft_", System.Data.OleDb.OleDbType.VarWChar, 50, "Ft#"))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("in_", System.Data.OleDb.OleDbType.VarWChar, 50, "in#"))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("MaterialBatch", System.Data.OleDb.OleDbType.VarWChar, 50, "MaterialBatch"))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("NeedleCal", System.Data.OleDb.OleDbType.VarWChar, 50, "NeedleCal"))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("out_", System.Data.OleDb.OleDbType.VarWChar, 50, "out#"))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("pass_", System.Data.OleDb.OleDbType.VarWChar, 50, "pass#"))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("PatternName", System.Data.OleDb.OleDbType.VarWChar, 50, "PatternName"))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Pause", System.Data.OleDb.OleDbType.VarWChar, 50, "Pause"))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("RunTime", System.Data.OleDb.OleDbType.VarWChar, 50, "RunTime"))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Source", System.Data.OleDb.OleDbType.VarWChar, 50, "Source"))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("SPCName", System.Data.OleDb.OleDbType.VarWChar, 50, "SPCName"))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Time", System.Data.OleDb.OleDbType.DBDate, 0, "Time"))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("time_B", System.Data.OleDb.OleDbType.VarWChar, 50, "time/B"))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Type", System.Data.OleDb.OleDbType.VarWChar, 50, "Type"))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("User_Group", System.Data.OleDb.OleDbType.VarWChar, 50, "User/Group"))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("UserAction", System.Data.OleDb.OleDbType.VarWChar, 255, "UserAction"))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("VolCal", System.Data.OleDb.OleDbType.VarWChar, 50, "VolCal"))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("OriginalIDS.__ID", System.Data.OleDb.OleDbType.Integer, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "ID", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_B_h", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "B/h", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_B_h1", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "B/h", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_DownTime", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "DownTime", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_DownTime1", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "DownTime", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_EvtID", System.Data.OleDb.OleDbType.Integer, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "EvtID", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_EvtID1", System.Data.OleDb.OleDbType.Integer, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "EvtID", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_EvtName", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "EvtName", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_EvtName1", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "EvtName", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Fp_", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Fp#", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Fp_1", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Fp#", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Ft_", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Ft#", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Ft_1", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Ft#", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_MaterialBatch", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "MaterialBatch", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_MaterialBatch1", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "MaterialBatch", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_NeedleCal", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "NeedleCal", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_NeedleCal1", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "NeedleCal", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_PatternName", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "PatternName", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_PatternName1", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "PatternName", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Pause", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Pause", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Pause1", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Pause", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_RunTime", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "RunTime", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_RunTime1", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "RunTime", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_SPCName", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "SPCName", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_SPCName1", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "SPCName", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Source", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Source", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Source1", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Source", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Time", System.Data.OleDb.OleDbType.DBDate, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Time", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Time1", System.Data.OleDb.OleDbType.DBDate, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Time", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Type", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Type", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Type1", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Type", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_User_Group", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "User/Group", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_User_Group1", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "User/Group", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_UserAction", System.Data.OleDb.OleDbType.VarWChar, 255, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "UserAction", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_UserAction1", System.Data.OleDb.OleDbType.VarWChar, 255, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "UserAction", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_VolCal", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "VolCal", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_VolCal1", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "VolCal", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_in_", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "in#", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_in_1", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "in#", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_out_", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "out#", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_out_1", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "out#", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_pass_", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "pass#", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_pass_1", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "pass#", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_time_B", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "time/B", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_time_B1", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "time/B", System.Data.DataRowVersion.Original, Nothing))
        '
        'Re_GenerateSPC
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(304, 426)
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.TextBox1)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.ListBox1)
        Me.Controls.Add(Me.RadioButton2)
        Me.Controls.Add(Me.RadioButton1)
        Me.Name = "Re_GenerateSPC"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Re-Generate SPC Report"
        Me.ResumeLayout(False)

    End Sub

#End Region
    'Public Generate_button As Integer = 0
    Dim for_pattern As Boolean
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Close()  'without actions
    End Sub

    Private Sub RadioButton1_click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton1.Click
        If RadioButton1.Checked Then
            for_pattern = False
            make_listbox_for1()
        End If
    End Sub

    Private Sub make_listbox_for1()
        Dim dataset11 As New DataSet
        Dim command11 As OleDbCommand = New OleDbCommand("SELECT DISTINCT SPCName FROM Log")
        command11.CommandType = CommandType.Text
        dataset11.Tables.Clear()
        OleDbConnection1.Open()
        command11.Connection = OleDbConnection1
        OleDbDataAdapter1.SelectCommand = command11
        OleDbDataAdapter1.Fill(dataset11)

        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '   Check the Table is empty. If there is no data, system also will clear the listbox    '
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If (dataset11.Tables("Log").Rows.Count = 0) Then
            MessageBox.Show("Data is empty!")
            OleDbConnection1.Close()
            ListBox1.DataSource = Nothing
            ListBox1.Items.Clear()
            Exit Sub
        End If

        ListBox1.DataSource = dataset11
        ListBox1.DisplayMember = "Log.SPCName"
        ListBox1.ValueMember = "Log.SPCName"
        OleDbConnection1.Close()
    End Sub

    Private Sub RadioButton2_click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton2.Click
        If RadioButton2.Checked Then
            for_pattern = True
            make_listbox_for2()
        End If
    End Sub

    Private Sub make_listbox_for2()
        Dim dataset11 As New DataSet
        Dim command11 As OleDbCommand = New OleDbCommand("SELECT DISTINCT PatternName FROM Log")
        command11.CommandType = CommandType.Text
        dataset11.Tables.Clear()
        OleDbConnection1.Open()
        command11.Connection = OleDbConnection1
        OleDbDataAdapter1.SelectCommand = command11
        OleDbDataAdapter1.Fill(dataset11)

        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '   Check the Table is empty. If there is no data, system also will clear the listbox    '
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If (dataset11.Tables("Log").Rows.Count = 0) Then
            MessageBox.Show("Data is empty!")
            OleDbConnection1.Close()
            ListBox1.DataSource = Nothing
            ListBox1.Items.Clear()
            Exit Sub
        End If

        ListBox1.DataSource = dataset11
        ListBox1.DisplayMember = "Log.PatternName"
        ListBox1.ValueMember = "Log.PatternName"
        OleDbConnection1.Close()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        'Dim savedia As New SaveFileDialog
        SaveFileDialog1.InitialDirectory = "c:\"
        SaveFileDialog1.Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*"
        SaveFileDialog1.FilterIndex = 0
        'SaveFileDialog1.CheckFileExists = True
        'SaveFileDialog1.RestoreDirectory = False
        'savedia.
        If SaveFileDialog1.ShowDialog = DialogResult.OK Then
            TextBox1.Text = SaveFileDialog1.FileName
        End If
        Button1.Enabled = True
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If (TextBox1.Text = "Please Input") Or (TextBox1.Text = Nothing) Then
            MsgBox("Please create file name and its path.")
            Exit Sub
        End If

        If (ListBox1.SelectedValue = Nothing) Then
            MsgBox("There is no data to report!")
            Exit Sub
        End If

        'Generate_button = 1
        Dim create_file As New FileStream(TextBox1.Text, FileMode.Create)
        create_file.Close()
        'If ReGener.Generate_button = 1 Then
        Do_Regenerate(TextBox1.Text, ListBox1.SelectedValue)
        'End If
        Close()
    End Sub

    Private Function Do_Regenerate(ByVal path As String, ByVal target As String)
        Dim SPC_content(20) As String
        'Dim content_flag(15) As Boolean
        Dim SPC_name As String
        Dim Pat_name As String
        Dim ToBeLogged As String
        Dim ToBeReturned As Integer = 0
        Dim select_words As String
        Dim temp_dataset As New DataSet
        Dim temp_command As OleDbCommand = New OleDbCommand(select_words)
        temp_command.CommandType = CommandType.Text
        Dim temp_id_start As String
        Dim temp_id_end As String
        Dim last_id_start As String = "0"
        Dim last_id_end As String = "0"
        Dim found As Boolean = True
        Dim pause_flag As Boolean = False

        If for_pattern Then
            While found
                SPC_content(0) = "0" : SPC_content(1) = "0" : SPC_content(2) = "0" : SPC_content(3) = "0" : SPC_content(4) = "0" : SPC_content(5) = "0"
                SPC_content(6) = "0" : SPC_content(7) = "0" : SPC_content(8) = "0" : SPC_content(9) = "0" : SPC_content(10) = "0" : SPC_content(11) = "0"
                SPC_content(12) = "0" : SPC_content(13) = "0" : SPC_content(14) = "0" : SPC_content(15) = "0" : SPC_content(16) = "0" : SPC_content(17) = "0"

                select_words = "SELECT MIN([ID]) FROM Log WHERE(EvtID = 201 AND PatternName = '" & target & "' AND [ID] > " & last_id_start & " )" 'find the last production end point
                temp_command = New OleDbCommand(select_words)
                temp_dataset.Tables.Clear()
                temp_command.Connection = OleDbConnection1
                OleDbDataAdapter1.SelectCommand = temp_command
                OleDbDataAdapter1.Fill(temp_dataset)
                OleDbConnection1.Close()
                If cstr(temp_dataset.Tables(0).Rows(0).Item(0)) <> "" Then
                    found = True

                    temp_id_start = cstr(temp_dataset.Tables(0).Rows(0).Item(0))
                    last_id_start = temp_id_start

                    select_words = "SELECT MIN([ID]) FROM Log WHERE(EvtID = 207 AND PatternName = '" & target & "' AND [ID] > " & last_id_end & " )" 'find the last production end point
                    temp_command = New OleDbCommand(select_words)
                    temp_dataset.Tables.Clear()
                    temp_command.Connection = OleDbConnection1
                    OleDbDataAdapter1.SelectCommand = temp_command
                    OleDbDataAdapter1.Fill(temp_dataset)
                    temp_id_end = cstr(temp_dataset.Tables(0).Rows(0).Item(0))
                    last_id_end = temp_id_end

                    select_words = "SELECT DISTINCT MaterialBatch FROM Log WHERE([ID] >= " & temp_id_start & " AND [ID] <= " & temp_id_end & ")" 'find the material batch
                    temp_command = New OleDbCommand(select_words)
                    temp_dataset.Tables.Clear()
                    temp_command.Connection = OleDbConnection1
                    OleDbDataAdapter1.SelectCommand = temp_command
                    OleDbDataAdapter1.Fill(temp_dataset)
                    Dim count As Integer = 1
                    For count = 0 To temp_dataset.Tables(0).Rows.Count() - 1
                        If count = 0 Then
                            SPC_content(5) = cstr(temp_dataset.Tables(0).Rows(count).Item(0))
                        Else
                            SPC_content(5) &= Chr(13) & Chr(10) & "					" & cstr(temp_dataset.Tables(0).Rows(count).Item(0))
                        End If
                    Next

                    select_words = "SELECT [in#], [out#], [pass#], [Fp#], [Ft#], Pause, [Time], VolCal, NeedleCal, [User/Group], PatternName FROM Log WHERE(ID >= " & temp_id_start & "AND ID <= " & temp_id_end & ")  ORDER BY [Time]" 'count incoming number
                    temp_command = New OleDbCommand(select_words)
                    temp_dataset.Tables.Clear()
                    temp_command.Connection = OleDbConnection1
                    OleDbDataAdapter1.SelectCommand = temp_command
                    OleDbDataAdapter1.Fill(temp_dataset)

                    SPC_content(0) = cstr(temp_dataset.Tables(0).Rows(0).Item(6))  'start time
                    SPC_content(1) = cstr(temp_dataset.Tables(0).Rows(temp_dataset.Tables(0).Rows.Count() - 1).Item(6))  'end time
                    SPC_content(2) = Format(DateDiff("s", Convert.ToDateTime(SPC_content(0)), Convert.ToDateTime(SPC_content(1))) / 3600, "0.000")  'duration
                    SPC_content(3) = cstr(temp_dataset.Tables(0).Rows(0).Item(9))  'user/group-id
                    SPC_content(4) = cstr(temp_dataset.Tables(0).Rows(0).Item(10))  'pattern name

                    For count = 0 To temp_dataset.Tables(0).Rows.Count() - 1
                        If temp_dataset.Tables(0).Rows(count).Item(0) = "1" Then
                            SPC_content(6) = cstr(cint(SPC_content(6)) + 1)   'number of incoming
                        End If
                        If temp_dataset.Tables(0).Rows(count).Item(1) = "1" Then
                            SPC_content(7) = cstr(cint(SPC_content(7)) + 1)   'number of outgoing
                        End If
                        'If temp_dataset.Tables(0).Rows(count).Item(2) = "1" Then
                        'SPC_content(8) = cstr(cint(SPC_content(8)) + 1)   'number of pass
                        'End If
                        If temp_dataset.Tables(0).Rows(count).Item(3) = "1" Then
                            SPC_content(9) = cstr(cint(SPC_content(9)) + 1)   'number of Fp
                        End If
                        If temp_dataset.Tables(0).Rows(count).Item(4) = "1" Then
                            SPC_content(10) = cstr(cint(SPC_content(10)) + 1)   'number of Ft
                        End If
                        If (temp_dataset.Tables(0).Rows(count).Item(5) = "1" Or pause_flag = True) And count < (temp_dataset.Tables(0).Rows.Count() - 1) Then
                            SPC_content(13) = cint(SPC_content(13)) + DateDiff("s", Convert.ToDateTime(temp_dataset.Tables(0).Rows(count).Item(6)), Convert.ToDateTime(temp_dataset.Tables(0).Rows(count + 1).Item(6)))  'down time
                            pause_flag = True
                        End If
                        If temp_dataset.Tables(0).Rows(count).Item(5) = "2" Then
                            pause_flag = False
                            SPC_content(15) = cstr(cint(SPC_content(15)) + 1)   'number of pause
                            'SPC_content(2) = Format(DateDiff("s", Convert.ToDateTime(SPC_content(0)), Convert.ToDateTime(SPC_content(1))) / 3600, "0.000")
                        End If
                        If temp_dataset.Tables(0).Rows(count).Item(7) = "1" Then
                            SPC_content(16) = cstr(cint(SPC_content(16)) + 1)   'number of VolCal
                        End If
                        If temp_dataset.Tables(0).Rows(count).Item(8) = "1" Then
                            SPC_content(17) = cstr(cint(SPC_content(17)) + 1)  'number of NeedleCal
                        End If
                    Next
                    SPC_content(8) = cstr(cint(SPC_content(7)) - cint(SPC_content(10)) - cint(SPC_content(9)))
                    SPC_content(11) = Format(DateDiff("s", Convert.ToDateTime(SPC_content(0)), Convert.ToDateTime(SPC_content(1))) / cint(SPC_content(7)), "0.000")
                    SPC_content(12) = Format(cint(SPC_content(7)) / DateDiff("s", Convert.ToDateTime(SPC_content(0)), Convert.ToDateTime(SPC_content(1))) * 3600, "0.000")
                    SPC_content(14) = DateDiff("s", Convert.ToDateTime(SPC_content(0)), Convert.ToDateTime(SPC_content(1))) - SPC_content(13)   'run time

                    Dim Writer As New StreamWriter(TextBox1.Text, True, System.Text.Encoding.Default)   'append=true
                    Writer.WriteLine("************************************************************************")
                    Writer.Write("SPC Report. Manually Re-generated at:   ") : Writer.WriteLine((Now))
                    Writer.WriteLine("************************************************************************")
                    Writer.Write("Production Start:			") : Writer.WriteLine(SPC_content(0))
                    Writer.Write("Production End:				") : Writer.WriteLine(SPC_content(1))
                    Writer.Write("Duration:				") : Writer.WriteLine(SPC_content(2) & " Hour(s)")
                    Writer.Write("Operator ID:				") : Writer.WriteLine(SPC_content(3))
                    Writer.Write("Pattern File:				") : Writer.WriteLine(SPC_content(4))
                    Writer.Write("Material Batch(es):			") : Writer.WriteLine(SPC_content(5))
                    Writer.Write("Number of Incoming boards:		") : Writer.WriteLine(SPC_content(6))
                    Writer.Write("Number of Outgoing boards:		") : Writer.WriteLine(SPC_content(7))
                    Writer.Write("Number of Passed boards:		") : Writer.WriteLine(SPC_content(8))
                    Writer.Write("Number of Failed boards (partially):	") : Writer.WriteLine(SPC_content(9))
                    Writer.Write("Number of Failed boards (totally):	") : Writer.WriteLine(SPC_content(10))
                    Writer.Write("Average time per boards:		") : Writer.WriteLine(SPC_content(11) & " Second(s)")
                    Writer.Write("Number of boards per hour:		") : Writer.WriteLine(SPC_content(12))
                    Writer.Write("Total down time of this production:	") : Writer.WriteLine(SPC_content(13) & " Second(s)")
                    Writer.Write("Total run time of this production:	") : Writer.WriteLine(SPC_content(14) & " Second(s)")
                    Writer.Write("Number of pauses:			") : Writer.WriteLine(SPC_content(15))
                    Writer.Write("Number of volume calibration:		") : Writer.WriteLine(SPC_content(16))
                    Writer.Write("Number of needle calibration:		") : Writer.WriteLine(SPC_content(17))
                    Writer.WriteLine("Report End #############################################################" & Chr(13) & Chr(10))
                    Writer.Close()
                Else
                    found = False
                End If
            End While
        Else
            While found
                SPC_content(0) = "0" : SPC_content(1) = "0" : SPC_content(2) = "0" : SPC_content(3) = "0" : SPC_content(4) = "0" : SPC_content(5) = "0"
                SPC_content(6) = "0" : SPC_content(7) = "0" : SPC_content(8) = "0" : SPC_content(9) = "0" : SPC_content(10) = "0" : SPC_content(11) = "0"
                SPC_content(12) = "0" : SPC_content(13) = "0" : SPC_content(14) = "0" : SPC_content(15) = "0" : SPC_content(16) = "0" : SPC_content(17) = "0"

                select_words = "SELECT MIN([ID]) FROM Log WHERE(EvtID = 201 AND SPCName = '" & target & "' AND [ID] > " & last_id_start & " )" 'find the last production end point
                temp_command = New OleDbCommand(select_words)
                temp_dataset.Tables.Clear()
                temp_command.Connection = OleDbConnection1
                OleDbDataAdapter1.SelectCommand = temp_command
                OleDbDataAdapter1.Fill(temp_dataset)
                OleDbConnection1.Close()
                If cstr(temp_dataset.Tables(0).Rows(0).Item(0)) <> "" Then
                    found = True

                    temp_id_start = cstr(temp_dataset.Tables(0).Rows(0).Item(0))
                    last_id_start = temp_id_start

                    select_words = "SELECT MIN([ID]) FROM Log WHERE(EvtID = 207 AND SPCName = '" & target & "' AND [ID] > " & last_id_end & " )" 'find the last production end point
                    temp_command = New OleDbCommand(select_words)
                    temp_dataset.Tables.Clear()
                    temp_command.Connection = OleDbConnection1
                    OleDbDataAdapter1.SelectCommand = temp_command
                    OleDbDataAdapter1.Fill(temp_dataset)
                    temp_id_end = cstr(temp_dataset.Tables(0).Rows(0).Item(0))
                    last_id_end = temp_id_end

                    select_words = "SELECT DISTINCT MaterialBatch FROM Log WHERE([ID] >= " & temp_id_start & " AND [ID] <= " & temp_id_end & ")" 'find the material batch
                    temp_command = New OleDbCommand(select_words)
                    temp_dataset.Tables.Clear()
                    temp_command.Connection = OleDbConnection1
                    OleDbDataAdapter1.SelectCommand = temp_command
                    OleDbDataAdapter1.Fill(temp_dataset)
                    Dim count As Integer = 1
                    For count = 0 To temp_dataset.Tables(0).Rows.Count() - 1
                        If count = 0 Then
                            SPC_content(5) = cstr(temp_dataset.Tables(0).Rows(count).Item(0))
                        Else
                            SPC_content(5) &= Chr(13) & Chr(10) & "					" & cstr(temp_dataset.Tables(0).Rows(count).Item(0))
                        End If
                    Next

                    select_words = "SELECT [in#], [out#], [pass#], [Fp#], [Ft#], Pause, [Time], VolCal, NeedleCal, [User/Group], PatternName FROM Log WHERE(ID >= " & temp_id_start & "AND ID <= " & temp_id_end & ")  ORDER BY [Time]" 'count incoming number
                    temp_command = New OleDbCommand(select_words)
                    temp_dataset.Tables.Clear()
                    temp_command.Connection = OleDbConnection1
                    OleDbDataAdapter1.SelectCommand = temp_command
                    OleDbDataAdapter1.Fill(temp_dataset)

                    SPC_content(0) = cstr(temp_dataset.Tables(0).Rows(0).Item(6))  'start time
                    SPC_content(1) = cstr(temp_dataset.Tables(0).Rows(temp_dataset.Tables(0).Rows.Count() - 1).Item(6))  'end time
                    SPC_content(2) = Format(DateDiff("s", Convert.ToDateTime(SPC_content(0)), Convert.ToDateTime(SPC_content(1))) / 3600, "0.000")  'duration
                    SPC_content(3) = cstr(temp_dataset.Tables(0).Rows(0).Item(9))  'user/group-id
                    SPC_content(4) = cstr(temp_dataset.Tables(0).Rows(0).Item(10))  'pattern name

                    For count = 0 To temp_dataset.Tables(0).Rows.Count() - 1
                        If temp_dataset.Tables(0).Rows(count).Item(0) = "1" Then
                            SPC_content(6) = cstr(cint(SPC_content(6)) + 1)   'number of incoming
                        End If
                        If temp_dataset.Tables(0).Rows(count).Item(1) = "1" Then
                            SPC_content(7) = cstr(cint(SPC_content(7)) + 1)   'number of outgoing
                        End If
                        'If temp_dataset.Tables(0).Rows(count).Item(2) = "1" Then
                        'SPC_content(8) = cstr(cint(SPC_content(8)) + 1)   'number of pass
                        'End If
                        If temp_dataset.Tables(0).Rows(count).Item(3) = "1" Then
                            SPC_content(9) = cstr(cint(SPC_content(9)) + 1)   'number of Fp
                        End If
                        If temp_dataset.Tables(0).Rows(count).Item(4) = "1" Then
                            SPC_content(10) = cstr(cint(SPC_content(10)) + 1)   'number of Ft
                        End If
                        If (temp_dataset.Tables(0).Rows(count).Item(5) = "1" Or pause_flag = True) And count < (temp_dataset.Tables(0).Rows.Count() - 1) Then
                            SPC_content(13) = cint(SPC_content(13)) + DateDiff("s", Convert.ToDateTime(temp_dataset.Tables(0).Rows(count).Item(6)), Convert.ToDateTime(temp_dataset.Tables(0).Rows(count + 1).Item(6)))  'down time
                            pause_flag = True
                        End If
                        If temp_dataset.Tables(0).Rows(count).Item(5) = "2" Then
                            pause_flag = False
                            SPC_content(15) = cstr(cint(SPC_content(15)) + 1)   'number of pause
                            'SPC_content(2) = Format(DateDiff("s", Convert.ToDateTime(SPC_content(0)), Convert.ToDateTime(SPC_content(1))) / 3600, "0.000")
                        End If
                        If temp_dataset.Tables(0).Rows(count).Item(7) = "1" Then
                            SPC_content(16) = cstr(cint(SPC_content(16)) + 1)   'number of VolCal
                        End If
                        If temp_dataset.Tables(0).Rows(count).Item(8) = "1" Then
                            SPC_content(17) = cstr(cint(SPC_content(17)) + 1)  'number of NeedleCal
                        End If
                    Next
                    SPC_content(8) = cstr(cint(SPC_content(7)) - cint(SPC_content(10)) - cint(SPC_content(9)))
                    SPC_content(11) = Format(DateDiff("s", Convert.ToDateTime(SPC_content(0)), Convert.ToDateTime(SPC_content(1))) / cint(SPC_content(7)), "0.000")
                    SPC_content(12) = Format(cint(SPC_content(7)) / DateDiff("s", Convert.ToDateTime(SPC_content(0)), Convert.ToDateTime(SPC_content(1))) * 3600, "0.000")
                    SPC_content(14) = DateDiff("s", Convert.ToDateTime(SPC_content(0)), Convert.ToDateTime(SPC_content(1))) - SPC_content(13)   'run time

                    Dim Writer As New StreamWriter(TextBox1.Text, True, System.Text.Encoding.Default)   'append=true
                    Writer.WriteLine("************************************************************************")
                    Writer.Write("SPC Report. Manually Re-generated at:   ") : Writer.WriteLine((Now))
                    Writer.WriteLine("************************************************************************")
                    Writer.Write("Production Start:			") : Writer.WriteLine(SPC_content(0))
                    Writer.Write("Production End:				") : Writer.WriteLine(SPC_content(1))
                    Writer.Write("Duration:				") : Writer.WriteLine(SPC_content(2) & " Hour(s)")
                    Writer.Write("Operator ID:				") : Writer.WriteLine(SPC_content(3))
                    Writer.Write("Pattern File:				") : Writer.WriteLine(SPC_content(4))
                    Writer.Write("Material Batch(es):			") : Writer.WriteLine(SPC_content(5))
                    Writer.Write("Number of Incoming boards:		") : Writer.WriteLine(SPC_content(6))
                    Writer.Write("Number of Outgoing boards:		") : Writer.WriteLine(SPC_content(7))
                    Writer.Write("Number of Passed boards:		") : Writer.WriteLine(SPC_content(8))
                    Writer.Write("Number of Failed boards (partially):	") : Writer.WriteLine(SPC_content(9))
                    Writer.Write("Number of Failed boards (totally):	") : Writer.WriteLine(SPC_content(10))
                    Writer.Write("Average time per boards:		") : Writer.WriteLine(SPC_content(11) & " Second(s)")
                    Writer.Write("Number of boards per hour:		") : Writer.WriteLine(SPC_content(12))
                    Writer.Write("Total down time of this production:	") : Writer.WriteLine(SPC_content(13) & " Second(s)")
                    Writer.Write("Total run time of this production:	") : Writer.WriteLine(SPC_content(14) & " Second(s)")
                    Writer.Write("Number of pauses:			") : Writer.WriteLine(SPC_content(15))
                    Writer.Write("Number of volume calibration:		") : Writer.WriteLine(SPC_content(16))
                    Writer.Write("Number of needle calibration:		") : Writer.WriteLine(SPC_content(17))
                    Writer.WriteLine("Report End #############################################################" & Chr(13) & Chr(10))
                    Writer.Close()
                Else
                    found = False
                End If
            End While
        End If
    End Function
End Class
