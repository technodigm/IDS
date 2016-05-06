Imports System.IO
Imports System.Text.RegularExpressions
Imports System.Data.OleDb

Public Class SPCReport
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
    Friend WithEvents OleDbDataAdapter1 As System.Data.OleDb.OleDbDataAdapter
    Friend WithEvents OleDbSelectCommand1 As System.Data.OleDb.OleDbCommand
    Friend WithEvents OleDbInsertCommand1 As System.Data.OleDb.OleDbCommand
    Friend WithEvents OleDbUpdateCommand1 As System.Data.OleDb.OleDbCommand
    Friend WithEvents OleDbDeleteCommand1 As System.Data.OleDb.OleDbCommand
    Friend WithEvents OleDbConnection1 As System.Data.OleDb.OleDbConnection
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.OleDbDataAdapter1 = New System.Data.OleDb.OleDbDataAdapter
        Me.OleDbDeleteCommand1 = New System.Data.OleDb.OleDbCommand
        Me.OleDbConnection1 = New System.Data.OleDb.OleDbConnection
        Me.OleDbInsertCommand1 = New System.Data.OleDb.OleDbCommand
        Me.OleDbSelectCommand1 = New System.Data.OleDb.OleDbCommand
        Me.OleDbUpdateCommand1 = New System.Data.OleDb.OleDbCommand
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
        'SPCReport
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(6, 14)
        Me.ClientSize = New System.Drawing.Size(348, 284)
        Me.Name = "SPCReport"
        Me.Text = "SPCReport"

    End Sub

#End Region

    Private Sub RetrieveData(ByRef temp_dataset As DataSet, ByVal temp_command As OleDbCommand, ByVal query As String)
        temp_command = New OleDbCommand(query)
        temp_dataset.Tables.Clear()
        temp_command.Connection = OleDbConnection1
        OleDbDataAdapter1.SelectCommand = temp_command
        OleDbDataAdapter1.Fill(temp_dataset)
    End Sub

    Public Function generate(ByVal mode As Integer)
        Dim SPC_name As String
        Dim pause_flag As Boolean = False

        If mode = 0 Then 'auto generate SPC for the last production
            Dim start_time As String = "0"
            Dim end_time As String = "0"
            Dim duration As String = "0"
            Dim group_id As String = ""
            Dim material_batch As String = ""
            Dim pattern_name As String = ""
            Dim incoming_num As String = "0"
            Dim outgoing_num As String = "0"
            Dim pass_num As String = "0"
            Dim boards_partial_fail As String = "0" 'fp
            Dim boards_total_fail As String = "0" 'ft
            Dim time_per_board As String = "0"
            Dim board_per_hour As String = "0"
            Dim down_time As String = "0"
            Dim run_time As String = "0"
            Dim pause_num As String = "0"
            Dim volcal_num As String = "0"
            Dim needlecal_num As String = "0"
            Dim dot_size_average As String = "0"

            Dim report(20) As Boolean
            Dim query As String
            Dim temp_dataset As New DataSet
            Dim temp_command As OleDbCommand = New OleDbCommand(query)
            Dim temp_id_start As String
            Dim temp_id_end As String

            SPC_name = IDS.Data.Hardware.SPC.ReportFileName
            temp_command.CommandType = CommandType.Text

            '0    Writer.Write("Material Batch(es):			") : Writer.WriteLine(material_batch)
            '1    Writer.Write("Number of incoming boards:		") : Writer.WriteLine(incoming_num)
            '2    Writer.Write("Number of outgoing boards:		") : Writer.WriteLine(outgoing_num)
            '3    Writer.Write("Number of passed boards:		") : Writer.WriteLine(pass_num)
            '4    Writer.Write("Number of failed boards (partially):	") : Writer.WriteLine(boards_partial_fail)
            '5    Writer.Write("Number of failed boards (totally):	") : Writer.WriteLine(boards_total_fail)
            '6    Writer.Write("Average time per board:		") : Writer.WriteLine(time_per_board & " Second(s)")
            '7    Writer.Write("Number of boards per hour:		") : Writer.WriteLine(board_per_hour)
            '8    Writer.Write("Total down time of this production:	") : Writer.WriteLine(down_time & " Second(s)")
            '9    Writer.Write("Total run time of this production:	") : Writer.WriteLine(run_time & " Second(s)")
            '10    Writer.Write("Number of pauses:			") : Writer.WriteLine(pause_num)
            For i As Integer = 0 To IDS.Data.Hardware.SPC.ItemsToBeReported.Length - 1
                If Convert.ToByte(IDS.Data.Hardware.SPC.ItemsToBeReported.Chars(i)) - 48 Then
                    report(i) = True
                End If
            Next

            query = "SELECT MAX([ID]) FROM Log WHERE(EvtID = 1007401)" 'find the last production start point
            RetrieveData(temp_dataset, temp_command, query)
            temp_id_start = CStr(temp_dataset.Tables(0).Rows(0).Item(0))

            query = "SELECT MAX([ID]) FROM Log WHERE(EvtID = 1007410)" 'find the last production end point
            RetrieveData(temp_dataset, temp_command, query)
            temp_id_end = CStr(temp_dataset.Tables(0).Rows(0).Item(0))

            query = "SELECT DISTINCT MaterialBatch FROM Log WHERE([ID] >= " & temp_id_start & " AND [ID] <= " & temp_id_end & ")" 'find the material batch
            RetrieveData(temp_dataset, temp_command, query)
            Dim count As Integer = 1
            For count = 0 To temp_dataset.Tables(0).Rows.Count() - 1
                If count = 0 Then
                    material_batch = CStr(temp_dataset.Tables(0).Rows(count).Item(0))
                Else
                    material_batch &= Chr(13) & Chr(10) & "					" & CStr(temp_dataset.Tables(0).Rows(count).Item(0))
                End If
            Next

            'kr look for qc dot checks
            Dim dot_size_values As Double
            query = "SELECT EvtName FROM Log WHERE([ID] >= " & temp_id_start & " AND [ID] <= " & temp_id_end & " AND [EvtName] LIKE '%Dot size%')"
            RetrieveData(temp_dataset, temp_command, query)
            count = 1
            Dim num_qc_events As Integer = temp_dataset.Tables(0).Rows.Count() - 1
            If num_qc_events > 0 Then
                For count = 0 To num_qc_events
                    'replace all but numeric values and period
                    dot_size_values += CDbl(Regex.Replace(CStr(temp_dataset.Tables(0).Rows(count).Item(0)), "[^0-9.]", ""))
                Next
                dot_size_values = dot_size_values / (num_qc_events + 1)   'average it out
                dot_size_values = CInt(dot_size_values * 1000.0) / 1000.0 'to 3 dec places
            End If

            query = "SELECT [in#], [out#], [pass#], [Fp#], [Ft#], Pause, [Time], VolCal, NeedleCal, [User/Group], PatternName FROM Log WHERE(ID >= " & temp_id_start & "AND ID <= " & temp_id_end & ")  ORDER BY [Time]" 'count incoming number
            RetrieveData(temp_dataset, temp_command, query)

            '0 = in
            '1 = out
            '2 = pass
            '3 = failure partial
            '4 = failure total
            '5 = pause
            '6 = time
            '7 = volcal
            '8 = ndlcal
            '9 = usr/group
            '10 = pattern name
            start_time = CStr(temp_dataset.Tables(0).Rows(0).Item(6))  'start time
            end_time = CStr(temp_dataset.Tables(0).Rows(temp_dataset.Tables(0).Rows.Count() - 1).Item(6))  'end time
            duration = Format(DateDiff("s", Convert.ToDateTime(start_time), Convert.ToDateTime(end_time)) / 3600, "0.000")  'duration
            group_id = CStr(temp_dataset.Tables(0).Rows(0).Item(9))  'user/group-id
            pattern_name = CStr(temp_dataset.Tables(0).Rows(0).Item(10))  'pattern name

            For count = 0 To temp_dataset.Tables(0).Rows.Count() - 1
                If temp_dataset.Tables(0).Rows(count).Item(0) = "1" Then
                    incoming_num = CStr(CInt(incoming_num) + 1)   'number of incoming
                End If
                If temp_dataset.Tables(0).Rows(count).Item(1) = "1" Then
                    outgoing_num = CStr(CInt(outgoing_num) + 1)   'number of outgoing
                End If
                If temp_dataset.Tables(0).Rows(count).Item(3) = "1" Then
                    boards_partial_fail = CStr(CInt(boards_partial_fail) + 1)   'number of Fp
                End If
                If temp_dataset.Tables(0).Rows(count).Item(4) = "1" Then
                    boards_total_fail = CStr(CInt(boards_total_fail) + 1)   'number of Ft
                End If
                If temp_dataset.Tables(0).Rows(count).Item(5) = "1" Or pause_flag = True And count < (temp_dataset.Tables(0).Rows.Count() - 1) Then
                    down_time = CInt(down_time) + DateDiff("s", Convert.ToDateTime(temp_dataset.Tables(0).Rows(count).Item(6)), Convert.ToDateTime(temp_dataset.Tables(0).Rows(count + 1).Item(6)))  'down time
                    pause_flag = True
                End If
                If temp_dataset.Tables(0).Rows(count).Item(5) = "2" Then
                    pause_flag = False
                    pause_num = CStr(CInt(pause_num) + 1)   'number of pause
                End If
            Next

            pass_num = CStr(CInt(outgoing_num) - CInt(boards_total_fail) - CInt(boards_partial_fail))
            time_per_board = Format(DateDiff("s", Convert.ToDateTime(start_time), Convert.ToDateTime(end_time)) / CInt(outgoing_num), "0.000")
            board_per_hour = Format(CInt(outgoing_num) / DateDiff("s", Convert.ToDateTime(start_time), Convert.ToDateTime(end_time)) * 3600, "0.000")
            run_time = DateDiff("s", Convert.ToDateTime(start_time), Convert.ToDateTime(end_time)) - down_time   'run time
            dot_size_average = CStr(dot_size_values) + " mm"

            Dim Writer As New StreamWriter(SPC_name, True, System.Text.Encoding.Default)   'append=true
            Writer.WriteLine("************************************************************************")
            Writer.Write("SPC Report. Automatically Generated at: ")
            Writer.WriteLine((Now))
            Writer.WriteLine("************************************************************************")
            Writer.Write("Production Start:			") : Writer.WriteLine(start_time)
            Writer.Write("Production End:				") : Writer.WriteLine(end_time)
            Writer.Write("Duration:				") : Writer.WriteLine(duration & " Hour(s)")
            Writer.Write("Operator ID:				") : Writer.WriteLine(group_id)
            Writer.Write("Pattern File:				") : Writer.WriteLine(pattern_name)
            If num_qc_events > 0 Then
                Writer.Write("Average dot size:			") : Writer.WriteLine(dot_size_average)
            End If
            If report(0) Then
                Writer.Write("Material Batch(es):			") : Writer.WriteLine(material_batch)
            End If
            If report(1) Then
                Writer.Write("Number of incoming boards:		") : Writer.WriteLine(incoming_num)
            End If
            If report(2) Then
                Writer.Write("Number of outgoing boards:		") : Writer.WriteLine(outgoing_num)
            End If
            If report(3) Then
                Writer.Write("Number of passed boards:		") : Writer.WriteLine(pass_num)
            End If
            If report(4) Then
                Writer.Write("Number of failed boards (partially):	") : Writer.WriteLine(boards_partial_fail)
            End If
            If report(5) Then
                Writer.Write("Number of failed boards (totally):	") : Writer.WriteLine(boards_total_fail)
            End If
            If report(6) Then
                Writer.Write("Average time per board:                 ") : Writer.WriteLine(time_per_board & " Second(s)")
            End If
            If report(7) Then
                Writer.Write("Number of boards per hour:		") : Writer.WriteLine(board_per_hour)
            End If
            If report(8) Then
                Writer.Write("Total down time of this production:	") : Writer.WriteLine(down_time & " Second(s)")
            End If
            If report(9) Then
                Writer.Write("Total run time of this production:	") : Writer.WriteLine(run_time & " Second(s)")
            End If
            If report(10) Then
                Writer.Write("Number of pauses:			") : Writer.WriteLine(pause_num)
            End If
            Writer.WriteLine("Report End @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@" & Chr(13) & Chr(10))
            Writer.Close()
        End If
    End Function

End Class
