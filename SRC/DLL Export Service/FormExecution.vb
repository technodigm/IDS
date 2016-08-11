Imports System.IO
Public Class FormExecution
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
    Friend WithEvents RichTextBox1 As System.Windows.Forms.RichTextBox
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents OleDbDataAdapter1 As System.Data.OleDb.OleDbDataAdapter
    Friend WithEvents OleDbSelectCommand1 As System.Data.OleDb.OleDbCommand
    Friend WithEvents OleDbInsertCommand1 As System.Data.OleDb.OleDbCommand
    Friend WithEvents OleDbUpdateCommand1 As System.Data.OleDb.OleDbCommand
    Friend WithEvents OleDbDeleteCommand1 As System.Data.OleDb.OleDbCommand
    Friend WithEvents OleDbConnection1 As System.Data.OleDb.OleDbConnection
    Friend WithEvents DataSet11 As DLL_Export_Service.DataSet1
    Friend WithEvents Button4 As System.Windows.Forms.Button
    Friend WithEvents Button5 As System.Windows.Forms.Button
    Friend WithEvents Button6 As System.Windows.Forms.Button
    Friend WithEvents Button7 As System.Windows.Forms.Button
    Friend WithEvents Button8 As System.Windows.Forms.Button
    Friend WithEvents Button9 As System.Windows.Forms.Button
    Friend WithEvents Button11 As System.Windows.Forms.Button
    Friend WithEvents Button12 As System.Windows.Forms.Button
    Friend WithEvents Button13 As System.Windows.Forms.Button
    Public WithEvents Button10_ToBeAutoClicked As System.Windows.Forms.Button
    Friend WithEvents Button14 As System.Windows.Forms.Button
    Friend WithEvents Button15 As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.RichTextBox1 = New System.Windows.Forms.RichTextBox
        Me.Button1 = New System.Windows.Forms.Button
        Me.Button2 = New System.Windows.Forms.Button
        Me.Button3 = New System.Windows.Forms.Button
        Me.OleDbDataAdapter1 = New System.Data.OleDb.OleDbDataAdapter
        Me.OleDbDeleteCommand1 = New System.Data.OleDb.OleDbCommand
        Me.OleDbConnection1 = New System.Data.OleDb.OleDbConnection
        Me.OleDbInsertCommand1 = New System.Data.OleDb.OleDbCommand
        Me.OleDbSelectCommand1 = New System.Data.OleDb.OleDbCommand
        Me.OleDbUpdateCommand1 = New System.Data.OleDb.OleDbCommand
        Me.DataSet11 = New DLL_Export_Service.DataSet1
        Me.Button4 = New System.Windows.Forms.Button
        Me.Button5 = New System.Windows.Forms.Button
        Me.Button6 = New System.Windows.Forms.Button
        Me.Button7 = New System.Windows.Forms.Button
        Me.Button8 = New System.Windows.Forms.Button
        Me.Button9 = New System.Windows.Forms.Button
        Me.Button10_ToBeAutoClicked = New System.Windows.Forms.Button
        Me.Button11 = New System.Windows.Forms.Button
        Me.Button12 = New System.Windows.Forms.Button
        Me.Button13 = New System.Windows.Forms.Button
        Me.Button14 = New System.Windows.Forms.Button
        Me.Button15 = New System.Windows.Forms.Button
        CType(Me.DataSet11, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'RichTextBox1
        '
        Me.RichTextBox1.Cursor = System.Windows.Forms.Cursors.Default
        Me.RichTextBox1.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RichTextBox1.Location = New System.Drawing.Point(16, 16)
        Me.RichTextBox1.MaxLength = 4096
        Me.RichTextBox1.Name = "RichTextBox1"
        Me.RichTextBox1.ReadOnly = True
        Me.RichTextBox1.Size = New System.Drawing.Size(592, 312)
        Me.RichTextBox1.TabIndex = 0
        Me.RichTextBox1.Text = "RichTextBox1"
        '
        'Button1
        '
        Me.Button1.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button1.Location = New System.Drawing.Point(16, 344)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(112, 64)
        Me.Button1.TabIndex = 1
        Me.Button1.Text = "OK"
        Me.Button1.Visible = False
        '
        'Button2
        '
        Me.Button2.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button2.Location = New System.Drawing.Point(144, 344)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(112, 64)
        Me.Button2.TabIndex = 2
        Me.Button2.Text = "SKIP"
        Me.Button2.Visible = False
        '
        'Button3
        '
        Me.Button3.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button3.Location = New System.Drawing.Point(496, 344)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(112, 64)
        Me.Button3.TabIndex = 3
        Me.Button3.Text = "STOP"
        Me.Button3.Visible = False
        '
        'OleDbDataAdapter1
        '
        Me.OleDbDataAdapter1.DeleteCommand = Me.OleDbDeleteCommand1
        Me.OleDbDataAdapter1.InsertCommand = Me.OleDbInsertCommand1
        Me.OleDbDataAdapter1.SelectCommand = Me.OleDbSelectCommand1
        Me.OleDbDataAdapter1.TableMappings.AddRange(New System.Data.Common.DataTableMapping() {New System.Data.Common.DataTableMapping("Table", "Log", New System.Data.Common.DataColumnMapping() {New System.Data.Common.DataColumnMapping("ID", "ID"), New System.Data.Common.DataColumnMapping("EvtID", "EvtID"), New System.Data.Common.DataColumnMapping("EvtName", "EvtName"), New System.Data.Common.DataColumnMapping("Time", "Time"), New System.Data.Common.DataColumnMapping("Source", "Source"), New System.Data.Common.DataColumnMapping("User/Group", "User/Group"), New System.Data.Common.DataColumnMapping("PatternName", "PatternName"), New System.Data.Common.DataColumnMapping("MaterialBatch", "MaterialBatch"), New System.Data.Common.DataColumnMapping("Type", "Type"), New System.Data.Common.DataColumnMapping("in#", "in#"), New System.Data.Common.DataColumnMapping("out#", "out#"), New System.Data.Common.DataColumnMapping("pass#", "pass#"), New System.Data.Common.DataColumnMapping("Fp#", "Fp#"), New System.Data.Common.DataColumnMapping("Ft#", "Ft#"), New System.Data.Common.DataColumnMapping("time/B", "time/B"), New System.Data.Common.DataColumnMapping("B/h", "B/h"), New System.Data.Common.DataColumnMapping("RunTime", "RunTime"), New System.Data.Common.DataColumnMapping("DownTime", "DownTime"), New System.Data.Common.DataColumnMapping("Pause", "Pause"), New System.Data.Common.DataColumnMapping("VolCal", "VolCal"), New System.Data.Common.DataColumnMapping("NeedleCal", "NeedleCal"), New System.Data.Common.DataColumnMapping("SPCName", "SPCName"), New System.Data.Common.DataColumnMapping("UserAction", "UserAction")})})
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
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_EvtID", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "EvtID", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_EvtID1", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "EvtID", System.Data.DataRowVersion.Original, Nothing))
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
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Time", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Time", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Time1", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Time", System.Data.DataRowVersion.Original, Nothing))
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
        Me.OleDbInsertCommand1.CommandText = "INSERT INTO Log(EvtID, EvtName, [Time], Source, [User/Group], PatternName, Materi" & _
        "alBatch, Type, [in#], [out#], [pass#], [Fp#], [Ft#], [time/B], [B/h], RunTime, D" & _
        "ownTime, Pause, VolCal, NeedleCal, SPCName, UserAction) VALUES (?, ?, ?, ?, ?, ?" & _
        ", ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"
        Me.OleDbInsertCommand1.Connection = Me.OleDbConnection1
        Me.OleDbInsertCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("EvtID", System.Data.OleDb.OleDbType.VarWChar, 50, "EvtID"))
        Me.OleDbInsertCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("EvtName", System.Data.OleDb.OleDbType.VarWChar, 50, "EvtName"))
        Me.OleDbInsertCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Time", System.Data.OleDb.OleDbType.VarWChar, 50, "Time"))
        Me.OleDbInsertCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Source", System.Data.OleDb.OleDbType.VarWChar, 50, "Source"))
        Me.OleDbInsertCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("User_Group", System.Data.OleDb.OleDbType.VarWChar, 50, "User/Group"))
        Me.OleDbInsertCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("PatternName", System.Data.OleDb.OleDbType.VarWChar, 50, "PatternName"))
        Me.OleDbInsertCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("MaterialBatch", System.Data.OleDb.OleDbType.VarWChar, 50, "MaterialBatch"))
        Me.OleDbInsertCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Type", System.Data.OleDb.OleDbType.VarWChar, 50, "Type"))
        Me.OleDbInsertCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("in_", System.Data.OleDb.OleDbType.VarWChar, 50, "in#"))
        Me.OleDbInsertCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("out_", System.Data.OleDb.OleDbType.VarWChar, 50, "out#"))
        Me.OleDbInsertCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("pass_", System.Data.OleDb.OleDbType.VarWChar, 50, "pass#"))
        Me.OleDbInsertCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Fp_", System.Data.OleDb.OleDbType.VarWChar, 50, "Fp#"))
        Me.OleDbInsertCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Ft_", System.Data.OleDb.OleDbType.VarWChar, 50, "Ft#"))
        Me.OleDbInsertCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("time_B", System.Data.OleDb.OleDbType.VarWChar, 50, "time/B"))
        Me.OleDbInsertCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("B_h", System.Data.OleDb.OleDbType.VarWChar, 50, "B/h"))
        Me.OleDbInsertCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("RunTime", System.Data.OleDb.OleDbType.VarWChar, 50, "RunTime"))
        Me.OleDbInsertCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("DownTime", System.Data.OleDb.OleDbType.VarWChar, 50, "DownTime"))
        Me.OleDbInsertCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Pause", System.Data.OleDb.OleDbType.VarWChar, 50, "Pause"))
        Me.OleDbInsertCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("VolCal", System.Data.OleDb.OleDbType.VarWChar, 50, "VolCal"))
        Me.OleDbInsertCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("NeedleCal", System.Data.OleDb.OleDbType.VarWChar, 50, "NeedleCal"))
        Me.OleDbInsertCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("SPCName", System.Data.OleDb.OleDbType.VarWChar, 50, "SPCName"))
        Me.OleDbInsertCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("UserAction", System.Data.OleDb.OleDbType.VarWChar, 255, "UserAction"))
        '
        'OleDbSelectCommand1
        '
        Me.OleDbSelectCommand1.CommandText = "SELECT ID, EvtID, EvtName, [Time], Source, [User/Group], PatternName, MaterialBat" & _
        "ch, Type, [in#], [out#], [pass#], [Fp#], [Ft#], [time/B], [B/h], RunTime, DownTi" & _
        "me, Pause, VolCal, NeedleCal, SPCName, UserAction FROM Log"
        Me.OleDbSelectCommand1.Connection = Me.OleDbConnection1
        '
        'OleDbUpdateCommand1
        '
        Me.OleDbUpdateCommand1.CommandText = "UPDATE Log SET EvtID = ?, EvtName = ?, [Time] = ?, Source = ?, [User/Group] = ?, " & _
        "PatternName = ?, MaterialBatch = ?, Type = ?, [in#] = ?, [out#] = ?, [pass#] = ?" & _
        ", [Fp#] = ?, [Ft#] = ?, [time/B] = ?, [B/h] = ?, RunTime = ?, DownTime = ?, Paus" & _
        "e = ?, VolCal = ?, NeedleCal = ?, SPCName = ?, UserAction = ? WHERE (ID = ?) AND" & _
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
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("EvtID", System.Data.OleDb.OleDbType.VarWChar, 50, "EvtID"))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("EvtName", System.Data.OleDb.OleDbType.VarWChar, 50, "EvtName"))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Time", System.Data.OleDb.OleDbType.VarWChar, 50, "Time"))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Source", System.Data.OleDb.OleDbType.VarWChar, 50, "Source"))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("User_Group", System.Data.OleDb.OleDbType.VarWChar, 50, "User/Group"))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("PatternName", System.Data.OleDb.OleDbType.VarWChar, 50, "PatternName"))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("MaterialBatch", System.Data.OleDb.OleDbType.VarWChar, 50, "MaterialBatch"))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Type", System.Data.OleDb.OleDbType.VarWChar, 50, "Type"))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("in_", System.Data.OleDb.OleDbType.VarWChar, 50, "in#"))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("out_", System.Data.OleDb.OleDbType.VarWChar, 50, "out#"))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("pass_", System.Data.OleDb.OleDbType.VarWChar, 50, "pass#"))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Fp_", System.Data.OleDb.OleDbType.VarWChar, 50, "Fp#"))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Ft_", System.Data.OleDb.OleDbType.VarWChar, 50, "Ft#"))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("time_B", System.Data.OleDb.OleDbType.VarWChar, 50, "time/B"))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("B_h", System.Data.OleDb.OleDbType.VarWChar, 50, "B/h"))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("RunTime", System.Data.OleDb.OleDbType.VarWChar, 50, "RunTime"))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("DownTime", System.Data.OleDb.OleDbType.VarWChar, 50, "DownTime"))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Pause", System.Data.OleDb.OleDbType.VarWChar, 50, "Pause"))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("VolCal", System.Data.OleDb.OleDbType.VarWChar, 50, "VolCal"))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("NeedleCal", System.Data.OleDb.OleDbType.VarWChar, 50, "NeedleCal"))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("SPCName", System.Data.OleDb.OleDbType.VarWChar, 50, "SPCName"))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("UserAction", System.Data.OleDb.OleDbType.VarWChar, 255, "UserAction"))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("OriginalIDS.__ID", System.Data.OleDb.OleDbType.Integer, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "ID", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_B_h", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "B/h", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_B_h1", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "B/h", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_DownTime", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "DownTime", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_DownTime1", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "DownTime", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_EvtID", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "EvtID", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_EvtID1", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "EvtID", System.Data.DataRowVersion.Original, Nothing))
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
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Time", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Time", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Time1", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Time", System.Data.DataRowVersion.Original, Nothing))
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
        'DataSet11
        '
        Me.DataSet11.DataSetName = "DataSet1"
        Me.DataSet11.Locale = New System.Globalization.CultureInfo("zh-CN")
        '
        'Button4
        '
        Me.Button4.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button4.Location = New System.Drawing.Point(144, 344)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(112, 64)
        Me.Button4.TabIndex = 4
        Me.Button4.Text = "SKIP PATTERN"
        Me.Button4.Visible = False
        '
        'Button5
        '
        Me.Button5.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button5.Location = New System.Drawing.Point(496, 344)
        Me.Button5.Name = "Button5"
        Me.Button5.Size = New System.Drawing.Size(112, 64)
        Me.Button5.TabIndex = 5
        Me.Button5.Text = "SKIP CHIP"
        Me.Button5.Visible = False
        '
        'Button6
        '
        Me.Button6.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button6.Location = New System.Drawing.Point(16, 344)
        Me.Button6.Name = "Button6"
        Me.Button6.Size = New System.Drawing.Size(112, 64)
        Me.Button6.TabIndex = 6
        Me.Button6.Text = "IGNORE"
        Me.Button6.Visible = False
        '
        'Button7
        '
        Me.Button7.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button7.Location = New System.Drawing.Point(16, 344)
        Me.Button7.Name = "Button7"
        Me.Button7.Size = New System.Drawing.Size(112, 64)
        Me.Button7.TabIndex = 7
        Me.Button7.Text = "CONTINUE"
        Me.Button7.Visible = False
        '
        'Button8
        '
        Me.Button8.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button8.Location = New System.Drawing.Point(496, 344)
        Me.Button8.Name = "Button8"
        Me.Button8.Size = New System.Drawing.Size(112, 64)
        Me.Button8.TabIndex = 8
        Me.Button8.Text = "CHG. SYRINGE"
        Me.Button8.Visible = False
        '
        'Button9
        '
        Me.Button9.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button9.Location = New System.Drawing.Point(496, 344)
        Me.Button9.Name = "Button9"
        Me.Button9.Size = New System.Drawing.Size(112, 64)
        Me.Button9.TabIndex = 9
        Me.Button9.Text = "OFF HEATER"
        Me.Button9.Visible = False
        '
        'Button10_ToBeAutoClicked
        '
        Me.Button10_ToBeAutoClicked.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button10_ToBeAutoClicked.Location = New System.Drawing.Point(144, 344)
        Me.Button10_ToBeAutoClicked.Name = "Button10_ToBeAutoClicked"
        Me.Button10_ToBeAutoClicked.Size = New System.Drawing.Size(112, 64)
        Me.Button10_ToBeAutoClicked.TabIndex = 10
        Me.Button10_ToBeAutoClicked.Text = "STOP WAITING"
        Me.Button10_ToBeAutoClicked.Visible = False
        '
        'Button11
        '
        Me.Button11.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button11.Location = New System.Drawing.Point(528, 24)
        Me.Button11.Name = "Button11"
        Me.Button11.Size = New System.Drawing.Size(72, 48)
        Me.Button11.TabIndex = 11
        Me.Button11.Text = "STOP SIREN"
        '
        'Button12
        '
        Me.Button12.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button12.Location = New System.Drawing.Point(16, 344)
        Me.Button12.Name = "Button12"
        Me.Button12.Size = New System.Drawing.Size(112, 64)
        Me.Button12.TabIndex = 12
        Me.Button12.Text = "RESET"
        Me.Button12.Visible = False
        '
        'Button13
        '
        Me.Button13.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button13.Location = New System.Drawing.Point(144, 344)
        Me.Button13.Name = "Button13"
        Me.Button13.Size = New System.Drawing.Size(112, 64)
        Me.Button13.TabIndex = 13
        Me.Button13.Text = "RELEASE"
        Me.Button13.Visible = False
        '
        'Button14
        '
        Me.Button14.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button14.Location = New System.Drawing.Point(496, 344)
        Me.Button14.Name = "Button14"
        Me.Button14.Size = New System.Drawing.Size(112, 64)
        Me.Button14.TabIndex = 14
        Me.Button14.Text = "WAIT"
        Me.Button14.Visible = False
        '
        'Button15
        '
        Me.Button15.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button15.Location = New System.Drawing.Point(144, 344)
        Me.Button15.Name = "Button15"
        Me.Button15.Size = New System.Drawing.Size(112, 64)
        Me.Button15.TabIndex = 2
        Me.Button15.Text = "REDO"
        Me.Button15.Visible = False
        '
        'FormExecution
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(624, 421)
        Me.ControlBox = False
        Me.Controls.Add(Me.Button15)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.Button7)
        Me.Controls.Add(Me.Button6)
        Me.Controls.Add(Me.Button4)
        Me.Controls.Add(Me.Button10_ToBeAutoClicked)
        Me.Controls.Add(Me.Button13)
        Me.Controls.Add(Me.Button9)
        Me.Controls.Add(Me.Button14)
        Me.Controls.Add(Me.Button8)
        Me.Controls.Add(Me.Button5)
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.Button11)
        Me.Controls.Add(Me.Button12)
        Me.Controls.Add(Me.RichTextBox1)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "FormExecution"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Error"
        Me.TopMost = True
        CType(Me.DataSet11, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

    Dim clickflag As Boolean = False
    Friend ButtonSelection As Integer = 0
    Dim delimiter As String = "	"
    Dim ToBeLogged As String

    Private Shared m_instance As FormExecution
    Public Shared ReadOnly Property Instance() As FormExecution
        Get
            If m_instance Is Nothing Then
                m_instance = New FormExecution
            Else
            End If
            Return m_instance
        End Get
    End Property

    Private Delegate Sub zzz()
    Public Sub CenterPopup()
        m_instance.Show()
        m_instance.CenterToScreen()
    End Sub

    Private Sub HideAndDispose()
        m_instance.Hide()
    End Sub

    Public Sub LogAnEvent(ByVal ID As String)

        Dim RderEvtDef As New StreamReader("c:/IDS/BIN/EventDef.txt")
        RderEvtDef.BaseStream.Seek(0, SeekOrigin.Begin)
        Dim delimiter As String = "	"
        Dim strfieldsindex As Integer = 0
        Dim settingfieldsindex As Integer = 0
        Dim IDfound As Boolean = False
        Try
            While (RderEvtDef.Peek() > -1)
                IDfound = False

                Dim strfields As String

                For Each strfields In RderEvtDef.ReadLine().Split(delimiter)

                    If strfields = "\\" & CStr(ID) Then
                        IDfound = True
                    End If


                    If IDfound And strfields.Length > 0 Then
                        strfieldsindex += 1
                        settingfieldsindex += 1
                        Select Case settingfieldsindex
                            Case 1 'EVENT ID
                                ToBeLogged = CStr(ID)
                            Case 2 'EVENT NAME, TIME
                                If ID = "1007412" Or ID = "1006204" Then
                                    ToBeLogged &= "	" & strfields + Vision.GetQCSuccess
                                Else
                                    ToBeLogged &= "	" & strfields
                                End If
                                ToBeLogged &= "	" & CStr(Now())
                            Case 3 'SOURCE, USER/GROUP, PATTERNNAME, MATERIALBATCH
                                ToBeLogged &= "	" & strfields
                                ToBeLogged &= "	" & IDS.Data.Admin.User.ID 'user/groupID
                                ToBeLogged &= "	" & IDS.Data.Admin.Folder.PatternPath 'pattern_file
                                Dim m As String
                                m = "L:" & IDS.Data.Hardware.Dispenser.Left.MaterialInfo _
                                & " R:" & IDS.Data.Hardware.Dispenser.Right.MaterialInfo
                                ToBeLogged &= "	" & m
                            Case 4 'TYPE
                                ToBeLogged &= "	" & strfields
                            Case 5 'IN#
                                ToBeLogged &= "	" & strfields 'to_get_income_number"
                            Case 6 'OUT#
                                ToBeLogged &= "	" & strfields 'to_get_outgo_number"
                            Case 7 'PASS#
                                ToBeLogged &= "	" & strfields 'to_get_pass_number"
                            Case 8 'FP#
                                ToBeLogged &= "	" & strfields 'to_get_failed(P)_number"
                            Case 9 'FT#, T/B, B/H, runT, downT
                                ToBeLogged &= "	" & strfields 'to_get_failed(T)_number"
                                'ToBeLogged &= "	T/B"
                                'ToBeLogged &= "	B/H"
                                'ToBeLogged &= "	runT"
                                'ToBeLogged &= "	downT"
                                ToBeLogged &= "	NA"
                                ToBeLogged &= "	NA"
                                ToBeLogged &= "	NA"
                                ToBeLogged &= "	NA"
                            Case 10 'pause
                                ToBeLogged &= "	" & strfields
                            Case 11 'volume cal
                                ToBeLogged &= "	NA"
                            Case 12 'needle cal
                                ToBeLogged &= "	NA"
                        End Select
                    End If
                Next
            End While

            RderEvtDef.Close()
            Dim logFileName As String
            logFileName = IDS.Data.Hardware.SPC.ReportFileName

            'SPC name
            ToBeLogged &= "	" & logFileName
            'user action
            ToBeLogged &= "	NA"
            WriteTheEvent(ToBeLogged)

        Catch ex As Exception
            ExceptionDisplay(ex)
        End Try

    End Sub

    Public Sub DisplayErrorPopup(ByVal ID As String)

        Dim RderEvtDef As New StreamReader("c:/IDS/BIN/EventDef.txt")
        RderEvtDef.BaseStream.Seek(0, SeekOrigin.Begin)
        Dim delimiter As String = "	"
        Dim additional_information As String = ""
        Dim strfieldsindex As Integer = 0
        Dim settingfieldsindex As Integer = 0
        Dim IDfound As Boolean = False
        Dim FormattingString As String = ControlChars.CrLf + "		"

        Try
            While (RderEvtDef.Peek() > -1)

                If ID = "1006204" Then
                    additional_information = FormattingString + Vision.ReportQCFailure()
                ElseIf ID = "1003201" Then
                    additional_information = FormattingString + Conveyor.ReportWidthAdjustFailure()
                End If

                IDfound = False
                Dim strfields As String
                Dim display_text(20) As String

                For Each strfields In RderEvtDef.ReadLine().Split(delimiter)

                    If strfields = "\\" & CStr(ID) Then
                        IDfound = True
                    End If

                    If IDfound And strfields.Length > 0 Then
                        strfieldsindex += 1
                        Select Case strfieldsindex
                            Case 1
                                display_text(0) = "Event ID: "
                                display_text(1) = CStr(ID)
                            Case 2
                                display_text(2) = strfields
                            Case 13
                                display_text(3) = "Details:"
                                display_text(4) = "		" & strfields
                            Case 14
                                Dim corrections As String
                                Dim corrections_index As Integer = 0
                                For Each corrections In strfields.Split(":")
                                    corrections_index += 1
                                    If corrections_index = 1 Then
                                        display_text(19) = "Please do:"
                                        display_text(5) = corrections 'a.
                                    ElseIf corrections_index = 2 Then
                                        display_text(6) = corrections
                                    ElseIf corrections_index = 3 Then
                                        display_text(7) = corrections 'b.
                                    ElseIf corrections_index = 4 Then
                                        display_text(8) = corrections
                                    ElseIf corrections_index = 5 Then
                                        display_text(9) = corrections 'c.
                                    ElseIf corrections_index = 6 Then
                                        display_text(10) = corrections
                                    End If
                                Next
                                m_instance.RichTextBox1.Text = display_text(0) & display_text(1) & "	" & display_text(2) & ControlChars.CrLf & ControlChars.CrLf & display_text(3) & ControlChars.CrLf & display_text(4) & additional_information & ControlChars.CrLf & ControlChars.CrLf & "		" & display_text(19) & ControlChars.CrLf & ControlChars.CrLf & "		" & display_text(5) & display_text(6) & ControlChars.CrLf & "		" & display_text(7) & display_text(8) & ControlChars.CrLf & "		" & display_text(9) & display_text(10) & ControlChars.CrLf

                                m_instance.RichTextBox1.Find(display_text(1))
                                m_instance.RichTextBox1.SelectionFont = New Font(RichTextBox1.Font, FontStyle.Bold)
                                m_instance.RichTextBox1.Find(display_text(2))
                                m_instance.RichTextBox1.SelectionFont = New Font(RichTextBox1.Font, FontStyle.Bold)
                                m_instance.RichTextBox1.Find(display_text(4))
                                m_instance.RichTextBox1.SelectionFont = New Font(RichTextBox1.Font, FontStyle.Bold)
                                m_instance.RichTextBox1.SelectionColor = Color.Red
                                m_instance.RichTextBox1.Find(display_text(19))
                                m_instance.RichTextBox1.SelectionFont = New Font(RichTextBox1.Font, FontStyle.Underline)
                                m_instance.RichTextBox1.Select(0, 0)
                        End Select

                        If strfieldsindex = 15 Then
                            m_instance.Button1.Hide()
                            m_instance.Button2.Hide()
                            m_instance.Button3.Hide()
                            m_instance.Button4.Hide()
                            m_instance.Button5.Hide()
                            m_instance.Button6.Hide()
                            m_instance.Button7.Hide()
                            m_instance.Button8.Hide()
                            m_instance.Button8.Hide()
                            m_instance.Button10_ToBeAutoClicked.Hide()
                            m_instance.Button12.Hide()
                            m_instance.Button13.Hide()
                            m_instance.Button14.Hide()
                            m_instance.Button15.Hide()
                            Dim strbutton As String
                            For Each strbutton In strfields.Split(",")
                                If strbutton = "1" Then
                                    m_instance.Button1.Show()
                                ElseIf strbutton = "2" Then
                                    m_instance.Button2.Show()
                                ElseIf strbutton = "3" Then
                                    m_instance.Button3.Show()
                                ElseIf strbutton = "4" Then
                                    m_instance.Button4.Show()
                                ElseIf strbutton = "5" Then
                                    m_instance.Button5.Show()
                                ElseIf strbutton = "6" Then
                                    m_instance.Button6.Show()
                                ElseIf strbutton = "7" Then
                                    m_instance.Button7.Show()
                                ElseIf strbutton = "8" Then
                                    m_instance.Button8.Show()
                                ElseIf strbutton = "9" Then
                                    m_instance.Button9.Show()
                                ElseIf strbutton = "10" Then
                                    m_instance.Button10_ToBeAutoClicked.Show()
                                ElseIf strbutton = "12" Then
                                    m_instance.Button12.Show()
                                ElseIf strbutton = "13" Then
                                    m_instance.Button13.Show()
                                ElseIf strbutton = "14" Then
                                    m_instance.Button14.Show()
                                ElseIf strbutton = "15" Then
                                    m_instance.Button15.Show()
                                End If
                            Next
                        End If
                    End If
                Next
            End While

            'IDS.Devices.DIO.DIO.TurnOnTowerSiren()
            IDS.Devices.DIO.DIO.TurnOnTowerLightRed()
            Form_Service.SetEventCode(ID)
            'm_instance.Invoke(New zzz(AddressOf CenterPopup))
            CenterPopup()
            'GC.Collect()

            'Catch ex As InvalidOperationException
        Catch ex As Exception
            ExceptionDisplay(ex)
        End Try

    End Sub

    Private Function WriteTheEvent(ByVal logcontent As String)
        Dim drow As System.data.DataRow = DataSet11.Log.NewRow()
        Dim logstring As String
        Dim logindex As Integer = 0
        For Each logstring In logcontent.Split(delimiter)
            logindex += 1
            drow(logindex) = logstring
        Next
        DataSet11.Log.Rows.Add(drow)
        OleDbDataAdapter1.Update(DataSet11)
    End Function

#Region "buttons"
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        ButtonSelection = 1
        ToBeLogged &= "	" & CStr(ButtonSelection)
        IDS.__ID += 10000
        HideAndDispose()

        IDS.Devices.DIO.DIO.TurnOffTowerLightRed()
        IDS.Devices.DIO.DIO.TurnOffTowerSiren()
    End Sub
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        ButtonSelection = 2
        ToBeLogged &= "	" & CStr(ButtonSelection)
        IDS.__ID += 20000
        HideAndDispose()

        IDS.Devices.DIO.DIO.TurnOffTowerLightRed()
        IDS.Devices.DIO.DIO.TurnOffTowerSiren()
    End Sub
    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        ButtonSelection = 3
        ToBeLogged &= "	" & CStr(ButtonSelection)
        IDS.__ID += 30000
        HideAndDispose()

        IDS.Devices.DIO.DIO.TurnOffTowerLightRed()
        IDS.Devices.DIO.DIO.TurnOffTowerSiren()
    End Sub
    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        ButtonSelection = 4
        ToBeLogged &= "	" & CStr(ButtonSelection)
        IDS.__ID += 40000
        HideAndDispose()

        IDS.Devices.DIO.DIO.TurnOffTowerLightRed()
        IDS.Devices.DIO.DIO.TurnOffTowerSiren()
    End Sub
    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        ButtonSelection = 5
        ToBeLogged &= "	" & CStr(ButtonSelection)
        IDS.__ID += 50000
        HideAndDispose()

        IDS.Devices.DIO.DIO.TurnOffTowerLightRed()
        IDS.Devices.DIO.DIO.TurnOffTowerSiren()
    End Sub
    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        ButtonSelection = 6
        ToBeLogged &= "	" & CStr(ButtonSelection)
        IDS.__ID += 60000
        HideAndDispose()

        IDS.Devices.DIO.DIO.TurnOffTowerLightRed()
        IDS.Devices.DIO.DIO.TurnOffTowerSiren()
    End Sub
    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        ButtonSelection = 7
        ToBeLogged &= "	" & CStr(ButtonSelection)
        IDS.__ID += 70000
        HideAndDispose()
        IDS.Devices.DIO.DIO.TurnOffTowerLightRed()
        IDS.Devices.DIO.DIO.TurnOffTowerSiren()
    End Sub
    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        ButtonSelection = 8
        ToBeLogged &= "	" & CStr(ButtonSelection)
        IDS.__ID += 80000
        HideAndDispose()

        IDS.Devices.DIO.DIO.TurnOffTowerLightRed()
        IDS.Devices.DIO.DIO.TurnOffTowerSiren()
    End Sub
    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        ButtonSelection = 9
        ToBeLogged &= "	" & CStr(ButtonSelection)
        IDS.__ID += 90000
        HideAndDispose()

        IDS.Devices.DIO.DIO.TurnOffTowerLightRed()
        IDS.Devices.DIO.DIO.TurnOffTowerSiren()
    End Sub
    Private Sub Button10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button10_ToBeAutoClicked.Click
        ButtonSelection = 10
        ToBeLogged &= "	" & CStr(ButtonSelection)
        IDS.__ID += 100000
        HideAndDispose()

        IDS.Devices.DIO.DIO.TurnOffTowerLightRed()
        IDS.Devices.DIO.DIO.TurnOffTowerSiren()
    End Sub
    Private Sub Button12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button12.Click
        ButtonSelection = 12
        ToBeLogged &= "	" & CStr(ButtonSelection)
        IDS.__ID += 120000
        HideAndDispose()

        IDS.Devices.DIO.DIO.TurnOffTowerLightRed()
        IDS.Devices.DIO.DIO.TurnOffTowerSiren()
    End Sub
    Private Sub Button13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button13.Click
        ButtonSelection = 13
        ToBeLogged &= "	" & CStr(ButtonSelection)
        IDS.__ID += 130000
        HideAndDispose()

        IDS.Devices.DIO.DIO.TurnOffTowerLightRed()
        IDS.Devices.DIO.DIO.TurnOffTowerSiren()
    End Sub
    Private Sub Button14_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button14.Click
        ButtonSelection = 14
        ToBeLogged &= "	" & CStr(ButtonSelection)
        IDS.__ID += 140000
        HideAndDispose()

        IDS.Devices.DIO.DIO.TurnOffTowerLightRed()
        IDS.Devices.DIO.DIO.TurnOffTowerSiren()
    End Sub
    Private Sub Button11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button11.Click
        IDS.Devices.DIO.DIO.TurnOffTowerSiren()
    End Sub

    Private Sub Button15_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button15.Click
        ButtonSelection = 15
        ToBeLogged &= "	" & CStr(ButtonSelection)
        IDS.__ID += 150000
        HideAndDispose()

        IDS.Devices.DIO.DIO.TurnOffTowerLightRed()
        IDS.Devices.DIO.DIO.TurnOffTowerSiren()
    End Sub

#End Region

End Class








