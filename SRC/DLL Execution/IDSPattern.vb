Option Explicit On 
Imports Microsoft.VisualBasic
Imports DLL_DataManager
Imports DLL_Export_Service
Imports Microsoft.Office.Interop


'position coordinate structure                                    
'                                                                 

Public Structure xyzCoordinate
    Dim X As Double
    Dim Y As Double
    Dim Z As Double
End Structure


'sub sheet structure                                              
'                                                                 

Public Structure SubSheetNameSub
    Dim SheetName As String
    Dim SheetInSelfUse As Integer
    Dim SheetInTotalUse As Integer
    Dim SheetRootRelated As Integer
End Structure


'Sub sheet name structure                                         
'                                                                 

Public Structure SubSheetNameManage
    Dim SheetName() As SubSheetNameSub
    Dim TotalNumberOfSheets As Integer
End Structure


'sub sheet data structure                                         
'                                                                 

Public Structure SubSheetForErrorCheck
    Dim SheetName As String
    Dim SubFirstPt As xyzCoordinate  'The first valid Pt
End Structure


'Sub sheet external reference structure                           
'                                                                 

Public Structure SubSheetExtRef
    Dim SheetIndex As Integer
    Dim ExternalRefPt As xyzCoordinate
    Dim RotationAngleRad As Double
    Dim SheetLevel As Integer

    Dim XMaxPtLineNo As Integer
    Dim YMaxPtLineNo As Integer
    Dim ZMaxPtLineNo As Integer
    Dim XMinPtLineNo As Integer
    Dim YMinPtLineNo As Integer
    Dim ZMinPtLineNo As Integer
End Structure


'Sub sheet error structure                                        
'                                                                 

Public Structure SubSheetErrorStruct
    Dim IntAbstract() As SubSheetForErrorCheck
    Dim ExtAbstract() As SubSheetExtRef
    Dim IntNumberOfSheets As Integer
    Dim ExtNumberOfSheets As Integer
End Structure


'Class CIDSPattern                                                                               
'  Despription:                                                                                   
'           Pattern file data handling class                                                       
'  Created by:                                                                                   
'           Shen Jian/Jiang     

Public Class CIDSPattern
    Private Const m_MaxColumnOfPatternWithVision As Integer = 78
    ' Private m_PatternCommand As String
    ' Dim m_Reference As New CIDSData.CIDSParameterID
    Public m_CurrentDPara As New CIDSData.CIDSPatternExecution.CIDSTemplate
    Public m_ErrorChk As New CIDSErrorCheck
    Public CopyPastePatternRec As PatternRecord

    Public SubCallSheetName As SubSheetNameManage

    Dim SubFilesClaster(200) As String



    Public Sub SubCallSheetInitialization(ByVal iMaxItem)
        SubCallSheetName.TotalNumberOfSheets = 0
        ReDim SubCallSheetName.SheetName(iMaxItem)

        Dim i As Integer
        For i = 0 To iMaxItem - 1
            SubCallSheetName.SheetName(i).SheetName = ""
            SubCallSheetName.SheetName(i).SheetInSelfUse = 0
            SubCallSheetName.SheetName(i).SheetInTotalUse = 0
            SubCallSheetName.SheetName(i).SheetRootRelated = 0
        Next
    End Sub


    'Array coordinate structure                                       
    '                                                                 

    Public Structure ArrayCoordinate
        Dim ArrayType As String

        Dim Pos1 As xyzCoordinate
        Dim Pos2 As xyzCoordinate
        Dim Pos3 As xyzCoordinate
        Dim Pos4 As xyzCoordinate
        Dim Pos5 As xyzCoordinate
        Dim Pos6 As xyzCoordinate
        Dim Pos7 As xyzCoordinate
        Dim Pos8 As xyzCoordinate
        Dim Pos9 As xyzCoordinate

        Dim XhDelta As Double
        Dim XvDelta As Double
        Dim YhDelta As Double
        Dim YvDelta As Double

        Dim xNumber As Integer
        Dim yNumber As Integer
    End Structure



    'Dispensing elements data structure                               
    '                                                                 

    Public Structure SpreadVisiableDataRecord
        Dim Name As String
        Dim Needle As String
        Dim DispenseFlag As String

        Dim Pos1 As xyzCoordinate
        Dim Pos2 As xyzCoordinate
        Dim Pos3 As xyzCoordinate

        Dim RowAndColumn As String 'added by kr for DotArray

        Dim TravelSpeed As Double
        Dim NeedleGap As Double
        Dim DispenseDuration As Double
        Dim TravelDelay As Double
        Dim RetractDelay As Double
        Dim ApproachDispHeight As Double
        Dim RetractSpeed As Double
        Dim RetractHeight As Double

        Dim ClearanceHeight As Double
        Dim DetailingDistance As Double
        Dim ArcRadius As Double
        Dim Pitch As Double
        Dim FilledHeight As Double

        Dim NumberOfRun As Integer
        Dim SpiralFlag As String
        Dim RotatingAngle As Double

        Dim EdgeClear As Double
        Dim IONumber As Integer
    End Structure



    'Vision data structure                                            
    '                                                                 

    Public Structure SpreadNoneVisiableDataRecord
        Dim Brightness As Integer
        Dim CheckBox As Boolean 'not sure
        Dim Binarized As Integer
        Dim CwCCw As Boolean
        Dim BlackDot As Boolean
        Dim DispModel As Integer
        Dim Open As Integer
        Dim InOut As Boolean
        Dim MainEdge As Integer
        Dim Close As Integer
        Dim PointX1 As Double
        Dim PointX2 As Double
        Dim PointX3 As Double
        Dim PointX4 As Double
        Dim PointX5 As Double
        Dim PointY1 As Double
        Dim PointY2 As Double
        Dim PointY3 As Double
        Dim PointY4 As Double
        Dim PointY5 As Double
        Dim position As Double
        Dim PosX As Double
        Dim PosY As Double
        Dim ROI As Double
        Dim Rot As Double
        Dim Size As Double
        Dim SizeX As Double
        Dim SizeY As Double
        Dim Threshold As Double
        Dim Vertical As Boolean
        Dim EdgeClearance As Double
        'QC
        Dim Compactness As Double
        Dim MaxArea As Double
        Dim MinArea As Double
        Dim MQC_OffX As Double
        Dim MQC_OffY As Double
        Dim MRegionX As Decimal
        Dim MRegionY As Decimal
        Dim MROIx As Decimal
        Dim MROIy As Decimal
        Dim Roughness As Double
        Dim Type As Integer
        'RM
        Dim AcceptRatio As Double
        Dim BlackWithoutRM As Boolean
        Dim BlackWithRM As Boolean
        Dim WhiteWithoutRM As Boolean
        Dim WhiteWithRM As Boolean
        Dim WoRM As Boolean
        Dim WRM As Boolean
    End Structure



    'Pattern data structure                                           
    '                                                                 

    Public Structure PatternRecord
        Dim pPara As SpreadVisiableDataRecord
        Dim VisionParameters As SpreadNoneVisiableDataRecord
    End Structure


    Public Sub New()

    End Sub


    'Load default dispensing parameters                               
    '   GroupID:  login group id                                      
    '   RecordID: record id                                           
    '                                                                 

    Public Function LoadDefaultPara(ByVal GroupID As String, ByVal RecordID As String, ByVal item As Integer) As Integer
        IDS.Data.ParameterID.GroupID = GroupID ' _IDS.Admin.User.Group.ID
        IDS.Data.ParameterID.RecordID = RecordID ''"FactoryDefault"
        InitDispara()
        Return 0
    End Function


    'Set initial data for dispensing element parameters               
    '                                                                 

    Public Sub InitDispara()
        m_CurrentDPara.Name = "Untitle"

        m_CurrentDPara.Pos1.X = 0.0
        m_CurrentDPara.Pos1.Y = 0.0
        m_CurrentDPara.Pos1.Z = 0.0

        m_CurrentDPara.Pos2.X = 0.0
        m_CurrentDPara.Pos2.Y = 0.0
        m_CurrentDPara.Pos2.Z = 0.0

        m_CurrentDPara.pos3.X = 0.0
        m_CurrentDPara.pos3.X = 0.0
        m_CurrentDPara.pos3.Z = 0.0

        m_CurrentDPara.DispenseFlag = "On"
        m_CurrentDPara.NeedleGap = 0.5
        m_CurrentDPara.DispenseDuration = 300
        m_CurrentDPara.ApproachDispHeight = 5.0
        m_CurrentDPara.TravelDelay = 100
        m_CurrentDPara.RetractDelay = 100
        m_CurrentDPara.TravelSpeed = 100
        m_CurrentDPara.RetractHeight = 5.0
        m_CurrentDPara.RetractSpeed = 50.0
        m_CurrentDPara.DetailingDistance = 0.0
        m_CurrentDPara.ClearanceHeight = 30.0
        m_CurrentDPara.ArcRadius = 3.0
        m_CurrentDPara.Pitch = 2.0
        m_CurrentDPara.NumofRun = 2
        m_CurrentDPara.FilledHeight = 10.0
        m_CurrentDPara.RotatingAngle = 0.0
        m_CurrentDPara.SubFileName = ""
        m_CurrentDPara.EdgeClearance = 0.0
        m_CurrentDPara.SpiralFlag = "IN"
        m_CurrentDPara.Needle = "Left"
    End Sub



    'Update display element data, (not in use currently)
    '   para:  Dsipaly element data                                
    '   inputStr:   input data string                                 
    '   currentcolumn:   the data string located column               
    '                                                                 

    Public Sub UpdateDispara(ByRef para As CIDSData.CIDSPatternExecution.CIDSTemplate, _
        ByVal inputStr As String, ByVal CurrentColumn As Integer)
        Select Case CurrentColumn
            Case gNeedleColumn
                para.Needle = inputStr
            Case gDispensColumn
                para.DispenseFlag = inputStr
            Case gPos1XColumn
                para.Pos1.X = CDbl(inputStr)
            Case gPos1YColumn
                para.Pos1.Y = CDbl(inputStr)
            Case gPos1ZColumn
                para.Pos1.Z = CDbl(inputStr)
            Case gPos2XColumn
                para.Pos2.X = CDbl(inputStr)
            Case gPos2YColumn
                para.Pos2.Y = CDbl(inputStr)
            Case gPos2ZColumn
                para.Pos2.Z = CDbl(inputStr)
            Case gPos3XColumn
                para.pos3.X = CDbl(inputStr)
            Case gPos3YColumn
                para.pos3.Y = CDbl(inputStr)
            Case gPos3ZColumn
                para.pos3.Z = CDbl(inputStr)

            Case gTravelSpeedColumn
                para.TravelSpeed = CDbl(inputStr)

            Case gNeedleGapColumn
                para.NeedleGap = CDbl(inputStr)

            Case gDurationColumn
                para.DispenseDuration = CDbl(inputStr)

            Case gTravelDelayColumn
                para.TravelDelay = CDbl(inputStr)

            Case gRetractDelayColumn
                para.RetractDelay = CDbl(inputStr)

            Case gApproachHtColumn
                para.ApproachDispHeight = CDbl(inputStr)

            Case gRetractSpeedColumn
                para.RetractSpeed = CDbl(inputStr)

            Case gRetractHtColumn
                para.RetractHeight = CDbl(inputStr)

            Case gClearanceHtColumn
                para.ClearanceHeight = CDbl(inputStr)

            Case gDTaiDistColumn
                para.DetailingDistance = CDbl(inputStr)

            Case gArcRadiusColumn
                para.ArcRadius = CDbl(inputStr)

            Case gPitchColumn
                para.Pitch = CDbl(inputStr)

            Case gFillHeightColumn
                para.FilledHeight = CDbl(inputStr)

            Case gNoOfRunColumn
                para.NumofRun = CInt(inputStr)

            Case gSprialColumn
                para.SpiralFlag = inputStr

            Case gRtAngleColumn
                para.RotatingAngle = CInt(inputStr)

            Case gEdgeClearColumn
                para.EdgeClearance = CDbl(inputStr)

            Case gIONumberColumn
                para.IONumber = CInt(inputStr)

            Case Else

        End Select
        TraceGCCollect
    End Sub


    'Set spreadsheet cells' color                                     
    '   cell1:  starting cell                                         
    '   cell2:  ending cell                                           
    '   flag:   color control flag = true:grey, = false: white        
    Public Sub Spreadsheet_CellGrey(ByVal cell1 As Object, ByVal cell2 As Object, _
    ByVal flag As Boolean, ByRef AxSpreadsheet As AxOWC10.AxSpreadsheet)
        AxSpreadsheet.ActiveSheet.Range(cell1, cell2).Select()
        With AxSpreadsheet.Selection
            If True = flag Then
                .Interior.Color = RGB(190, 190, 190)
            Else
                .Interior.Color = RGB(255, 255, 255)
            End If
        End With
        TraceGCCollect
    End Sub

    'Scan Spreadsheet to find unoptimizable ele                       
    '   Axspreadsheet:  ActiveX data table                            
    '   return:         False: Cannot be optimized,                   
    '                   True: can be optimized
    Public Function Spreadsheet_CheckOptimizable(ByRef AxSpreadsheet As AxOWC10.AxSpreadsheet) As Boolean

        Dim Rtn As Boolean = True
        Dim i, j As Integer
        Dim NumberOfSheet As Integer = AxSpreadsheet.Worksheets.Count()
        Dim strPageName As String
        Dim cmdStr As String
        Dim UsedRowNumber As Integer

        For i = 1 To NumberOfSheet
            strPageName = AxSpreadsheet.ActiveWorkbook.Worksheets(i).name()
            j = 1
            With AxSpreadsheet.Worksheets(strPageName)
                UsedRowNumber = .UsedRange.Rows.Count
                Do Until j > UsedRowNumber
                    cmdStr = CStr(.Cells(j, gCommandNameColumn).Value())

                    'This is verified by EET on 20112007
                    If "Line" = cmdStr Or "Arc" = cmdStr Or "Circle" = cmdStr Or "FillCircle" = cmdStr Or _
                        "Rectangle" = cmdStr Or "FillRectangle" = cmdStr Or "Clean" = cmdStr Or "Purge" = cmdStr Or _
                        "Wait" = cmdStr Or "Move" = cmdStr Or "ChipEdge" = cmdStr Or "QC" = cmdStr Then

                        Rtn = False
                        Exit For
                    End If
                    j += 1
                Loop
            End With
        Next i

        Return Rtn
        TraceGCCollect
    End Function


    'Find a given sheet's parent sheet                                
    '   InputSheetName:  a given sheet name                           
    '   return:   the parent sheet name                               
    '                                                                 

    Public Function Spreadsheet_FindParrentSheet(ByVal InputSheetName As String) As String
        Dim LineStr(5) As String
        Dim ParrentSheetName As String

        LineStr = InputSheetName.Split(".")

        If 0 = LineStr.Length Then

        ElseIf 1 = LineStr.Length Or 2 = LineStr.Length Then
            ParrentSheetName = "Main"
        ElseIf 3 = LineStr.Length Then
            ParrentSheetName = LineStr(0) + "." + LineStr(1)
        Else
            ParrentSheetName = "Main"
        End If

        Return ParrentSheetName
    End Function



    'Get the local reference point for the current inserting row      
    '   Axspreadsheet:  instance of activeX spreadsheet               
    '   CurrentRow:     current row number                            
    '   RefPos:         reference point coordinate                    
    '                                                                 

    Public Sub Spreadsheet_GetRowLocalReference(ByRef AxSpreadsheet As AxOWC10.AxSpreadsheet, _
    ByVal CurrentRow As Integer, ByRef RefPos() As Double)
        Dim i As Integer
        'Default ref point will be used it is not found
        RefPos(0) = 0
        RefPos(1) = 0
        RefPos(2) = 0

        With AxSpreadsheet.ActiveSheet
            '"Reference" is absolut value.  Thus its ref is zero
            If "Reference" = .Cells(CurrentRow, gCommandNameColumn).value Then
                Return
            End If

            'Inherent previous ref within the active page
            For i = 1 To CurrentRow
                If "Reference" = .Cells(i, gCommandNameColumn).value Then
                    RefPos(0) = .Cells(i, gPos1XColumn).value
                    RefPos(1) = .Cells(i, gPos1YColumn).value
                    RefPos(2) = .Cells(i, gPos1ZColumn).value
                End If
            Next i
        End With
    End Sub



    'Get the last point of the active sheet                           
    '   Axspreadsheet:  instance of activeX spreadsheet               
    '   StartRow:       starting row number                           
    '   EndRow:         ending row number                             
    '   position:            last point coordinate                         
    '                                                                 

    Public Function Spreadsheet_FindLastValidPoint(ByRef AxSpreadsheet As AxOWC10.AxSpreadsheet, _
        ByVal StartRow As Integer, ByVal EndRow As Integer, ByRef position() As Double) As Boolean

        Dim Rtn As Boolean = False
        Dim i As Integer
        Dim CmdName As String

        With AxSpreadsheet.ActiveSheet
            For i = StartRow To EndRow
                CmdName = .Cells(i, gCommandNameColumn).value
                If "" = CmdName Then

                ElseIf m_ErrorChk.CheckValidPtXY(CmdName) Then
                    position(0) = .Cells(i, gPos1XColumn).value
                    position(1) = .Cells(i, gPos1YColumn).value
                    position(2) = .Cells(i, gPos1ZColumn).value
                    Rtn = True
                End If
            Next i
        End With

        Return Rtn
    End Function





    'Check sheet whether exists                                       
    '   Axspreadsheet:  instance of activeX spreadsheet               
    '   SheetName:      sheet name to be checked                     
    '   return:         0: not exist, 1: exist                                                                                       '

    Public Function Spreadsheet_CheckSubsheetExist(ByRef AxSpreadsheet As AxOWC10.AxSpreadsheet, _
    ByRef SheetName As String) As Integer

        Dim Rtn As Integer = 0
        Dim i As Integer
        Dim NumberOfSheet As Integer = AxSpreadsheet.Worksheets.Count()
        Dim LocalSpreadSheetName As String

        For i = 1 To NumberOfSheet
            LocalSpreadSheetName = AxSpreadsheet.ActiveWorkbook.Worksheets(i).name()
            If LocalSpreadSheetName = SheetName Then
                Rtn = 1
                Exit For
            End If
        Next i

        Return Rtn
    End Function



    'Insert a work sheet into the activeX spreadsheet                 
    '   Axspreadsheet:  instance of activeX spreadsheet               
    '   sheet:          work sheet                                    
    '   sheet_name:     work sheet name                               
    '                                                                 

    Public Sub InsertSheet(ByRef AxSpreadsheet As AxOWC10.AxSpreadsheet, ByVal sheet As OWC10.Worksheet, ByVal sheet_name As String)
        Try
            Dim sub_sheet As OWC10.Worksheet = AxSpreadsheet.Worksheets.Add(, sheet)
            sub_sheet.Name = sheet_name
            sub_sheet.Range("B1:B1").Select()
            AxSpreadsheet.ActiveWindow.FreezePanes = True
            'sub_sheet.
        Catch ex As Exception

        End Try
        TraceGCCollect
    End Sub


    'Check a sheet whether it is an array sheet                       
    '   Axspreadsheet:  instance of activeX spreadsheet               
    '   SheetName:      sheet name to be checked                      
    '   return:         true: array sheet, false: non array sheet     
    '                                                                 

    Public Function Spreadsheet_IsAnArraySheet(ByRef AxSpreadsheet As AxOWC10.AxSpreadsheet, _
            ByVal SheetName As String) As Boolean

        Dim Rtn As Boolean = False
        Dim ArrString As String()

        ArrString = SheetName.Split("."c)
        Dim len As Integer = ArrString.Length

        If 1 = len Then
            If "Main" <> SheetName Then
                Rtn = True
            End If
        Else
            If 3 = len Then
                Rtn = True
            End If
        End If

        TraceGCCollect
        Return Rtn
    End Function



    'Build root connection for Array and Sub sheet, Used by copy/paste/cut only                     
    '   Axspreadsheet:  instance of activeX spreadsheet               
    '                                                                 

    Public Sub Spreadsheet_BuildRootConnectedArraySub(ByRef AxSpreadsheet As AxOWC10.AxSpreadsheet)
        Dim Rtn As Integer = 0
        Dim i, j, k, emptyLine As Integer
        Dim NumberOfSheet As Integer = AxSpreadsheet.Worksheets.Count()
        Dim SpreadSheetName As String
        Dim CurrentSheetName As String
        Dim strTmp As String
        Dim SheetName As String

        Dim sel As Microsoft.Office.Interop.OWC.Range = AxSpreadsheet.Selection
        Dim m_rowCount As Integer = sel.Rows.Count()
        Dim m_columnCount As Integer = sel.Columns.Count()
        Dim m_StartRow As Integer = sel.Row
        Dim m_columnStart As Integer = sel.Column


        'Build the connectivity of the first sub sheet.  It's also in self use.
        For i = 1 To NumberOfSheet
            SubCallSheetName.SheetName(i - 1).SheetName = AxSpreadsheet.ActiveWorkbook.Worksheets(i).name()
        Next


        'Build the connectivity from the current sheet in the selected range
        CurrentSheetName = AxSpreadsheet.ActiveWorkbook.ActiveSheet.Name
        j = m_StartRow - 1
        emptyLine = 0
        With AxSpreadsheet.Worksheets(CurrentSheetName)
            Do
                j = j + 1
                strTmp = .Cells(j, gCommandNameColumn).Value

                If "Array" = strTmp Or "SubPattern" = strTmp Then
                    emptyLine = 0

                    SheetName = .Cells(j, gSubnameColumn).Value

                    'Translate the sheet name in physical to Spreadsheet tab
                    If "SubPattern" = strTmp Then
                        SheetName = m_Execution.m_File.NameOnlyFromFullPath(SheetName)
                    Else
                        If CurrentSheetName = "Main" Then

                        Else
                            SheetName = CurrentSheetName + "." + SheetName
                        End If
                    End If

                    'Scan all sheet, mark "1" if it has the same name, it is connected
                    For k = 0 To NumberOfSheet - 1
                        If SheetName = SubCallSheetName.SheetName(k).SheetName Then
                            SubCallSheetName.SheetName(k).SheetRootRelated = 1
                            Exit For
                        End If
                    Next k
                End If
            Loop Until j > m_StartRow + m_rowCount - 1
        End With


        'Build the connectivity of the remaining sub sheets
        Dim iLayer As Integer = 0
        Dim iConnected As Integer = 0

        'Loop for all the layers
        Do
            iLayer += 1
            iConnected = 0
            'Loop for all the sheets
            For i = 0 To NumberOfSheet - 1

                If iLayer = SubCallSheetName.SheetName(i).SheetRootRelated Then
                    CurrentSheetName = SubCallSheetName.SheetName(i).SheetName

                    j = 0
                    emptyLine = 0
                    With AxSpreadsheet.Worksheets(CurrentSheetName)
                        'Loop within the sheet
                        Do
                            j = j + 1
                            strTmp = .Cells(j, gCommandNameColumn).Value

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
                                    If SheetName = SubCallSheetName.SheetName(k).SheetName And _
                                    0 = SubCallSheetName.SheetName(k).SheetRootRelated Then
                                        SubCallSheetName.SheetName(k).SheetRootRelated = iLayer + 1
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
        TraceGCCollect
    End Sub





    'Get sheet's last used row                                        
    '   Axspreadsheet:  instance of activeX spreadsheet               
    '   SheetName:      sheet name                                    
    '   return:         last used row number                          
    '                                                                 

    Public Function Spreadsheet_CountRowUsed(ByRef AxSpreadsheet As AxOWC10.AxSpreadsheet, ByRef SheetName As String) As Integer
        'This function can be replaced by an existing one build in the OWC10.  But OWC10 one may be unstable
        'sometimes.  We haven't found the root cause.

        Dim j, emptyLine As Integer
        j = 0
        emptyLine = 0
        With AxSpreadsheet.Worksheets(SheetName)
            Do
                j = j + 1
                If "" = .Cells(j, gCommandNameColumn).Value Then
                    emptyLine = emptyLine + 1
                Else
                    emptyLine = 0
                End If
            Loop Until 20 < emptyLine 'Successive 20 empty excel lines will terminate the saving record.
        End With
        Return j - 21
    End Function

    'Not used, not tested
    Public Function Spreadsheet_isOnlyFirstRowEffective( _
        ByRef AxSpreadsheet As AxOWC10.AxSpreadsheet, _
        ByRef SheetName As String, _
        ByVal firstRowNo As Integer, ByVal lastRowNo As Integer) As Boolean

        Dim onlyFirstRowEffective As Boolean = False
        Dim j, emptyLine As Integer
        j = firstRowNo
        emptyLine = 0
        With AxSpreadsheet.Worksheets(SheetName)
            Do
                j = j + 1
                If "" = .Cells(j, gCommandNameColumn).Value Then
                    emptyLine = emptyLine + 1
                Else
                    emptyLine = 0
                End If
            Loop Until (j >= lastRowNo)
        End With
        If (emptyLine >= lastRowNo - firstRowNo) Then onlyFirstRowEffective = True

        Return onlyFirstRowEffective
    End Function



    'Get the last point of the active sheet                           
    '   file:           instance of ids file handling class           
    '   Encrypt:        true: encrypt, false: non-encrypt             
    '                                                                 

    Function GotoNextPageTxtPattern(ByRef file As CIDSFileHandler, ByVal Encrypt As Boolean) As String
        Dim eof As String = ""
        Dim LineStr(70) As String

        Do Until eof.Equals("EOF")
            eof = file.Read(Encrypt)

            If "EOF" = eof Then
                Return "EOF"
            End If

            If "" = eof Then
                Do Until eof <> ""
                    eof = file.Read(Encrypt)
                Loop
            End If

            If (eof <> Nothing Or eof <> "") And (eof.Chars(0)) <> "#"c Then
                LineStr = eof.Split("=")
            End If

            If 2 = LineStr.Length Then
                LineStr(0) = LineStr(0).Trim(" ")
                LineStr(1) = LineStr(1).Trim(" ")

                If "[" = eof.ToString.Chars(0) Then
                    LineStr(0) = LineStr(0).Trim("[")
                    LineStr(1) = LineStr(1).Trim("]")

                    If "Page" = LineStr(0) Then

                        Return LineStr(1)
                    End If
                End If
            End If
        Loop
        TraceGCCollect

    End Function

    'Read data has three mode distinguished by variable AttachFlag
    'AttachFlag acts as, 
    '   0=new start
    '   1=sub sheet
    '   2=import attachment


    'Read pattern file into activeX spreadsheet only for copy/Cut--Paste use '
    '   Axspreadsheet:  instance of activeX spreadsheet               '
    '   Filename:       pattern file name                             '
    '   Subname:        sub pattern file name                         '
    '   AttachFlag:     three mod for attatchment                     '
    '                 0= new start, 1=sub sheet, 2= import attatchment'
    '   StartRow:       insertion starting row number                 '
    '   Encrypt:        file encrypt flag                             ' 
    '                                                                 '

    Function LoadTxtPatternParaAllSheets(ByRef AxSpreadsheet As AxOWC10.AxSpreadsheet, ByVal Filename As String, _
            ByVal Subname As String, ByVal AttachFlag As Integer, ByVal StartRow As Integer, _
            ByVal Encrypt As Boolean) As Integer
        Dim file As New CIDSFileHandler
        Dim ret As Integer = file.OpenR(Filename)
        Dim rowOfLoad As Integer = 0
        Dim strInfo As String

        If (ret <> 0) Then
            'Error Message
            file.Close()
            Return 1
        End If

        Dim eof As String = ""
        Dim LineStr(80) As String
        Dim SubPageExist As Boolean = False

        Do Until eof.Equals("EOF")
            eof = file.Read(Encrypt)

            If "EOF" = eof Then
                file.Close()
                If rowOfLoad > 1 Then
                    Return 0
                Else
                    Return 1
                End If
            End If

            If "" = eof Then
                Do Until eof <> ""
                    eof = file.Read(Encrypt)
                Loop
            End If

            If (eof <> Nothing Or eof <> "") And (eof.Chars(0)) <> "#"c Then
                LineStr = eof.Split("=")
                Dim value As Double = 0.0
                If 2 = LineStr.Length Then
                    LineStr(0) = LineStr(0).Trim(" ")
                    LineStr(1) = LineStr(1).Trim(" ")

                    If "[" = eof.ToString.Chars(0) Then
                        LineStr(0) = LineStr(0).Trim("[")
                        LineStr(1) = LineStr(1).Trim("]")

                        If "Page" = LineStr(0) Then
                            If "Main" <> LineStr(1) And "" <> LineStr(1) Then   'Attached as new
                                'Check the page name existance.  Skip if yes

                                Do
                                    SubPageExist = Spreadsheet_CheckSubsheetExist(AxSpreadsheet, LineStr(1))
                                    If True = SubPageExist Then
                                        LineStr(1) = GotoNextPageTxtPattern(file, Encrypt)
                                    End If
                                    If "EOF" = LineStr(1) Then
                                        Exit Do
                                    End If
                                Loop Until (False = SubPageExist)

                                If True = SubPageExist Then
                                    Exit Do
                                Else

                                    AxSpreadsheet.Sheets.Add.Name = LineStr(1)
                                    AxSpreadsheet.Worksheets(LineStr(1)).Activate()

                                    AxSpreadsheet.ActiveWindow.FreezePanes = False
                                    AxSpreadsheet.Worksheets(LineStr(1)).Range("B1:B1").Select()
                                    AxSpreadsheet.ActiveWindow.FreezePanes = True
                                    rowOfLoad = 0
                                End If

                            ElseIf 1 = AttachFlag And "Main" = LineStr(1) Then  'Attached as sub
                                LineStr(1) = Subname
                                'Check the page name existance.  Skip if yes

                                Do
                                    SubPageExist = Spreadsheet_CheckSubsheetExist(AxSpreadsheet, LineStr(1))
                                    If True = SubPageExist Then
                                        LineStr(1) = GotoNextPageTxtPattern(file, Encrypt)
                                    End If
                                    If "EOF" = LineStr(1) Then
                                        Exit Do
                                    End If
                                Loop Until (False = SubPageExist)

                                If True = SubPageExist Then
                                    Exit Do
                                Else
                                    AxSpreadsheet.Sheets.Add.Name = LineStr(1)
                                    AxSpreadsheet.Worksheets(LineStr(1)).Activate()

                                    AxSpreadsheet.ActiveWindow.FreezePanes = False
                                    AxSpreadsheet.Worksheets(LineStr(1)).Range("B1:B1").Select()
                                    AxSpreadsheet.ActiveWindow.FreezePanes = True
                                    rowOfLoad = 0
                                End If

                            ElseIf 2 = AttachFlag And "Main" = LineStr(1) Then  'Attached as import
                                rowOfLoad = StartRow - 1
                                If rowOfLoad < 0 Then
                                    rowOfLoad = 0
                                End If
                            Else
                                AxSpreadsheet.Worksheets(LineStr(1)).Activate()
                                rowOfLoad = 0
                            End If

                        End If
                    Else
                        LineStr(0) = LineStr(0).ToUpper()
                        LineStr(1) = LineStr(1).ToUpper()
                        Select Case LineStr(0)
                            Case "PATTERNFILENAME"

                            Case "PATTERNID"

                            Case Else

                        End Select
                    End If


                ElseIf 1 = LineStr.Length Then

                    'Insert an empty line to avoid data overwrite
                    rowOfLoad += 1
                    AxSpreadsheet.ActiveWindow.ActiveSheet.Cells(rowOfLoad, 1).Select()
                    AxSpreadsheet.Selection.EntireRow.Insert()
                    AxSpreadsheet.ActiveSheet.Rows(rowOfLoad).clear()

                    LineStr = eof.Split(",")

                    'Error checking for loading data record
                    'm_ErrorChk.ConvertToDataStruct(LineStr)

                    If LineStr.Length < gMaxColumns Or LineStr.Length > m_MaxColumnOfPatternWithVision Then
                        strInfo = "Number of data is incorrect in record:"
                        strInfo = strInfo + CStr(rowOfLoad - 1)
                        MessageBox.Show(strInfo, "Warning message")
                    Else
                        'Insert an empty line

                        'kr change to increase the number of columns loaded from linestr into spreadsheet
                        'this is because of invisible cells (30 + onwards) that are not loaded when 
                        'copying and pasting vision elements like QC
                        For i As Integer = 1 To m_MaxColumnOfPatternWithVision

                            LineStr(i - 1) = LineStr(i - 1).Trim(" ")
                            If LineStr(i - 1) = "" Then
                                Do Until LineStr(i - 1) <> ""
                                    i = i + 1
                                    'kr change to increase the number of columns loaded from linestr into spreadsheet
                                    'this is because of invisible cells (30 + onwards) that are not loaded when 
                                    'copying and pasting vision elements like QC
                                    If i = m_MaxColumnOfPatternWithVision + 1 Then
                                        Exit For
                                    End If
                                    LineStr(i - 1) = LineStr(i - 1).Trim(" ")
                                Loop
                            End If

                            If "" <> LineStr(i - 1) Then
                                AxSpreadsheet.ActiveSheet.Cells(rowOfLoad, i) = LineStr(i - 1)
                            End If
                        Next i
                    End If
                End If
            End If
        Loop

        file.Close()
        Return 0
        TraceGCCollect
    End Function



    'Read text pattern file into activeX spreadsheet                  '
    '   Axspreadsheet:  instance of activeX spreadsheet               '
    '   Filename:       pattern file name                             '
    '   AttachFlag:     three mod for attatchment                     '
    '                 0= new start, 1=sub sheet, 2= import attatchment'
    '   StartRow:       insertion starting row number                 '
    '   Encrypt:        file encrypt flag                             ' 
    '                                                                 '

    Function LoadTxtPatternPara(ByRef AxSpreadsheet As AxOWC10.AxSpreadsheet, ByVal Filename As String, _
            ByVal AttachFlag As Integer, ByVal StartRow As Integer, _
            ByVal Encrypt As Boolean) As Integer
        Dim Rtn As Integer = 0

        Dim NumOfSubFilesToRead As Integer = 0
        Dim i As Integer

        'Open "Main" and Arrays attached to it.  Rtn=0 if success
        Rtn = LoadTxtPatternSingleFile(AxSpreadsheet, Filename, AttachFlag, StartRow, Encrypt, SubFilesClaster, NumOfSubFilesToRead)

        'For all sub loadings if any.  It also need previous loading to be success
        Do Until (NumOfSubFilesToRead = 0 Or Rtn <> 0)
            Filename = SubFilesClaster(0)
            NumOfSubFilesToRead -= 1

            For i = 0 To NumOfSubFilesToRead
                SubFilesClaster(i) = SubFilesClaster(i + 1)
            Next
            SubFilesClaster(NumOfSubFilesToRead) = ""
            Rtn = LoadTxtPatternSingleFile(AxSpreadsheet, Filename, 1, 0, Encrypt, SubFilesClaster, NumOfSubFilesToRead)
        Loop

        Return Rtn  '0=no error
        TraceGCCollect
    End Function



    'Read pattern file into activeX spreadsheet                       '
    '   Axspreadsheet:  instance of activeX spreadsheet               '
    '   Filename:       pattern file name                             '
    '   Subname:        sub pattern file name                         '
    '   AttachFlag:     three mod for attatchment                     '
    '                 0= new start, 1=sub sheet, 2= import attatchment'
    '   StartRow:       insertion starting row number                 '
    '   Encrypt:        file encrypt flag                             ' 
    '                                                                 '

    'KR FILE NAME BUG

    Function LoadTxtPatternSingleFile(ByRef AxSpreadsheet As AxOWC10.AxSpreadsheet, ByVal Filename As String, _
        ByVal AttachFlag As Integer, ByVal StartRow As Integer, _
        ByVal Encrypt As Boolean, ByRef SubFilesClaster() As String, _
        ByRef NumOfSubFilesToRead As Integer) As Integer
        Dim file As New CIDSFileHandler
        Dim ret As Integer = file.OpenR(Filename)
        Dim rowOfLoad As Integer = 0
        Dim strInfo As String
        Dim Subname As String

        Try
            If (ret <> 0) Then
                'Error Message
                file.Close()
                Return 1
            End If

            If AttachFlag = 1 Or AttachFlag = 2 Then      'Attached as a sub
                Subname = file.NameOnlyFromFullPath(Filename)
            End If


            Dim eof As String = ""
            Dim LineStr(80) As String
            Dim SubPageExist As Boolean = False

            Do Until eof.Equals("EOF")
                eof = file.Read(Encrypt)

                If "EOF" = eof Then
                    file.Close()

                    ''''''''''
                    '   Xue Wen                                                                 '
                    '   Set a focus cell for programming after user load old pattern file.      '
                    '   Note: a.) This is only for "Main page".                                 '
                    '         b.) Check for "Sub/Array page".                                   '
                    ''''''''''
                    If (rowOfLoad >= 1) Then
                        AxSpreadsheet.ActiveWindow.ActiveSheet.Cells(rowOfLoad + 1, 1).Select()
                    End If

                    Return 0 'Empty data file with is allowed
                End If

                If "" = eof Then
                    Do Until eof <> ""
                        eof = file.Read(Encrypt)
                    Loop
                End If

                If (eof <> Nothing Or eof <> "") And (eof.Chars(0)) <> "#"c Then
                    LineStr = eof.Split("=")
                    Dim value As Double = 0.0
                    If 2 = LineStr.Length Then
                        LineStr(0) = LineStr(0).Trim(" ")
                        LineStr(1) = LineStr(1).Trim(" ")

                        If "[" = eof.ToString.Chars(0) Then
                            LineStr(0) = LineStr(0).Trim("[")
                            LineStr(1) = LineStr(1).Trim("]")

                            If "Page" = LineStr(0) Then
                                If "Main" <> LineStr(1) And "" <> LineStr(1) Then   'Attached as new
                                    'Check the page name existance.  Skip if yes
                                    If 1 = AttachFlag Or 2 = AttachFlag Then
                                        LineStr(1) = Subname + "." + LineStr(1)
                                    End If

                                    Do
                                        SubPageExist = Spreadsheet_CheckSubsheetExist(AxSpreadsheet, LineStr(1))
                                        If True = SubPageExist Then
                                            LineStr(1) = GotoNextPageTxtPattern(file, Encrypt)
                                        End If
                                        If "EOF" = LineStr(1) Then
                                            Exit Do
                                        End If
                                    Loop Until (False = SubPageExist)

                                    If True = SubPageExist Then
                                        Exit Do
                                    Else
                                        'If LineStr(1).Length > 31 Then
                                        '    LineStr(1) = LineStr(1).Remove(31, LineStr(1).Length - 31)
                                        'End If
                                        AxSpreadsheet.Sheets.Add.Name = LineStr(1)
                                        AxSpreadsheet.Worksheets(LineStr(1)).Activate()
                                        AxSpreadsheet.ActiveWindow.FreezePanes = False
                                        AxSpreadsheet.Worksheets(LineStr(1)).Range("B1:B1").Select()
                                        AxSpreadsheet.ActiveWindow.FreezePanes = True
                                        rowOfLoad = 0
                                        End If

                                ElseIf "Main" = LineStr(1) And 1 = AttachFlag Then   'Attached as sub
                                        LineStr(1) = Subname
                                        'Check the page name existance.  Skip if yes

                                        Do
                                            SubPageExist = Spreadsheet_CheckSubsheetExist(AxSpreadsheet, LineStr(1))
                                            If True = SubPageExist Then
                                                LineStr(1) = GotoNextPageTxtPattern(file, Encrypt)
                                            End If
                                            If "EOF" = LineStr(1) Then
                                                Exit Do
                                            End If
                                        Loop Until (False = SubPageExist)

                                        If True = SubPageExist Then
                                            Exit Do
                                        Else
                                            AxSpreadsheet.Sheets.Add.Name = LineStr(1)
                                            AxSpreadsheet.Worksheets(LineStr(1)).Activate()
                                            AxSpreadsheet.ActiveWindow.FreezePanes = False
                                            AxSpreadsheet.Worksheets(LineStr(1)).Range("B1:B1").Select()
                                            AxSpreadsheet.ActiveWindow.FreezePanes = True
                                            rowOfLoad = 0
                                        End If

                                ElseIf 2 = AttachFlag And "Main" = LineStr(1) Then  'Attached as import
                                        rowOfLoad = StartRow - 1
                                        If rowOfLoad < 0 Then
                                            rowOfLoad = 0
                                        End If
                                Else
                                        AxSpreadsheet.Worksheets(LineStr(1)).Activate()
                                        rowOfLoad = 0
                                End If

                            End If
                        Else
                            LineStr(0) = LineStr(0).ToUpper()
                            LineStr(1) = LineStr(1).ToUpper()
                            Select Case LineStr(0)
                                Case "UNIT"

                                Case "PATTERNID"

                                Case Else

                            End Select
                        End If

                    ElseIf 1 = LineStr.Length Then

                        'Insert an empty line to avoid data overwrite
                        rowOfLoad += 1
                        AxSpreadsheet.ActiveWindow.ActiveSheet.Cells(rowOfLoad, 1).Select()
                        AxSpreadsheet.Selection.EntireRow.Insert()
                        AxSpreadsheet.ActiveSheet.Rows(rowOfLoad).clear()

                        LineStr = eof.Split(",")

                        'Error checking for loading data record
                        'm_ErrorChk.ConvertToDataStruct(LineStr)

                        If LineStr.Length < gMaxColumns Or _
                            LineStr.Length > m_MaxColumnOfPatternWithVision Then
                            strInfo = "Number of data is incorrect in record:"
                            strInfo = strInfo + CStr(rowOfLoad - 1)
                            MessageBox.Show(strInfo, "Warning message")
                        Else

                            'kr put this in here. feel free to remove if bug suspected.
                            Dim ReadOnlyOnce As Boolean = False 'prevent multiple entries of NumOfSubFilesToRead

                            For i As Integer = 1 To gWRMCoulumn  'gMaxColumns

                                LineStr(i - 1) = LineStr(i - 1).Trim(" ")
                                If LineStr(i - 1) = "" Then
                                    Do Until LineStr(i - 1) <> ""
                                        i = i + 1
                                        If i > gWRMCoulumn Then
                                            Exit For
                                        End If
                                        LineStr(i - 1) = LineStr(i - 1).Trim(" ")
                                    Loop
                                End If

                                If "" <> LineStr(i - 1) Then
                                    Select Case i
                                        Case gCommandNameColumn To gMaxColumns
                                            If ReadOnlyOnce = False And "SubPattern" = CStr(AxSpreadsheet.ActiveSheet.Cells(rowOfLoad, gCommandNameColumn).value) Then
                                                SubFilesClaster(NumOfSubFilesToRead) = LineStr(i - 1)
                                                NumOfSubFilesToRead += 1
                                                ReadOnlyOnce = True
                                            End If
                                            AxSpreadsheet.ActiveSheet.Cells(rowOfLoad, i) = LineStr(i - 1)
                                            'Lim's code here =====
                                        Case gMaxColumns + 1 To gWRMCoulumn
                                            AxSpreadsheet.ActiveSheet.Cells(rowOfLoad, i) = LineStr(i - 1)
                                            '======
                                        Case Else
                                    End Select
                                End If
                            Next i
                        End If
                    End If
                End If
            Loop
        Catch ex As Exception
            ExceptionDisplay(ex)
        End Try

        file.Close()
        Return 0
        TraceGCCollect
    End Function



    'Write activeX spreadsheet selected data to pattern file, used by Copy/Cut/Paste
    '   Axspreadsheet:  instance of activeX spreadsheet               
    '   Filename:       pattern file name   
    '   SelectFlag:     0=All data will be saved, 
    '                   1=Only selected data will be saved
    '   Encrypt:        file encrypt flag                             
    '                                                                 

    Function SavePatternParaAllSheets(ByRef AxSpreadsheet As AxOWC10.AxSpreadsheet, ByVal Filename As String, _
            ByVal SelectFlag As Integer, ByVal Encrypt As Boolean) As Integer
        Dim file As New CIDSFileHandler
        Dim ret As Integer = file.OpenW(Filename)

        If (ret <> 0) Then
            'Error Message
            file.Close()
            Return 1
        End If

        WriteHeaderDispara(file, Filename, "ABCDE_GHIJ_KLMN", Encrypt)
        WritePatternParaAll(file, AxSpreadsheet, SelectFlag, Encrypt)

        file.Close()
        Return 0
        TraceGCCollect
    End Function




    'Write data array pattern file for optimization                           '
    '   DataArray:      array data table               '
    '   Filename:       pattern file name                             '
    '   Encrypt:        file encrypt flag                             ' 
    '                                                                 '

    Function SavePatternParaOpt(ByRef DataArray As ArrayList, ByVal Filename As String, _
        ByVal Encrypt As Boolean) As Integer
        Dim file As New CIDSFileHandler
        Dim Rtn As Integer
        Dim i, j, k As Integer
        Dim strPageName, tmpStrCmd As String

        Dim NumberOfSubpage As Integer = 0
        Dim UsedRowNumber, arrayDim As Integer


        'Save the Main with its Array
        Rtn = file.OpenW(Filename)
        If (Rtn <> 0) Then  'Error Message
            file.Close()
            Return Rtn
        End If
        WriteHeaderDispara(file, Filename, "ABCDE_GHIJ_KLMN", Encrypt)
        WritePatternParaOpt(file, DataArray, Encrypt, "Main")
        file.Close()

        Return Rtn
        TraceGCCollect
    End Function





    'Write activeX spreadsheet pattern file                           '
    '   Axspreadsheet:  instance of activeX spreadsheet               '
    '   Filename:       pattern file name                             '
    '   Encrypt:        file encrypt flag                             ' 
    '                                                                 '

    Function SavePatternPara(ByRef AxSpreadsheet As AxOWC10.AxSpreadsheet, ByVal Filename As String, _
        ByVal Encrypt As Boolean) As Integer
        Dim file As New CIDSFileHandler
        Dim Rtn As Integer
        Dim i, j, k As Integer
        Dim NumberOfSheets As Integer = AxSpreadsheet.Worksheets.Count()
        Dim strPageName, tmpStrCmd As String

        Dim NumberOfSubpage As Integer = 0
        Dim UsedRowNumber, arrayDim As Integer


        'Find all qualified subPage partialFileName.  This part is not quite necessory
        For i = 1 To NumberOfSheets
            strPageName = AxSpreadsheet.ActiveWorkbook.Worksheets(i).name()
            arrayDim = 1
            For j = 0 To Len(strPageName) - 1
                If "." = strPageName.Chars(j) Then
                    arrayDim += 1
                End If
            Next

            If 2 = arrayDim Then
                SubFilesClaster(NumberOfSubpage) = strPageName
                NumberOfSubpage += 1
            End If
        Next


        'Build all qualified subPage fullFileName
        Dim countUp As Integer
        Dim startCount As Boolean

        For i = 1 To NumberOfSheets
            strPageName = AxSpreadsheet.ActiveWorkbook.Worksheets(i).name()

            countUp = 1
            startCount = False
            j = 1

            With AxSpreadsheet.Worksheets(strPageName)
                UsedRowNumber = .UsedRange.Rows.Count

                Do Until (countUp > UsedRowNumber)
                    tmpStrCmd = CStr(.Cells(j, gCommandNameColumn).Value())

                    If (tmpStrCmd <> Nothing) Then
                        If "SubPattern" = tmpStrCmd Then
                            tmpStrCmd = CStr(.Cells(j, gSubnameColumn).Value())
                            For k = 0 To NumberOfSubpage - 1
                                If SubFilesClaster(k) = file.NameOnlyFromFullPath(tmpStrCmd) Then
                                    SubFilesClaster(k) = tmpStrCmd
                                End If
                            Next
                        End If

                        countUp = countUp + 1
                        startCount = True
                    Else
                        If (startCount = True) Then
                            countUp = countUp + 1
                        End If
                    End If

                    j += 1
                Loop
            End With
        Next


        'Save the Main with its Array
        Rtn = file.OpenW(Filename)
        If (Rtn <> 0) Then  'Error Message
            file.Close()
            Return Rtn
        End If
        WriteHeaderDispara(file, Filename, "ABCDE_GHIJ_KLMN", Encrypt)
        WritePatternPara(file, AxSpreadsheet, Encrypt, "Main")
        file.Close()

        'Only export the "Main" and the attached Array
        If "ptp" = file.ExtOnlyFromFullPath(Filename) Then
            TraceGCCollect
            Return Rtn
        End If

        'Save the sub with its Array
        i = 0
        Do Until (NumberOfSubpage = i Or Rtn <> 0)
            Filename = SubFilesClaster(i)
            strPageName = file.NameOnlyFromFullPath(Filename)

            Rtn = file.OpenW(Filename)
            If (Rtn <> 0) Then
                file.Close()
                Return 1
            End If
            WriteHeaderDispara(file, Filename, "ABCDE_GHIJ_KLMN", Encrypt)
            WritePatternPara(file, AxSpreadsheet, Encrypt, strPageName)

            file.Close()
            i += 1
        Loop
        Return Rtn
        TraceGCCollect
    End Function


    'Write pattern file header                                        '
    '   File:       instance of file handle class                     '
    '   filename:   pattern file name                                 '
    '   strPatternID:  pattern id number                              '
    '   Encrypt:        file encrypt flag                             ' 
    '                                                                 '

    Private Sub WriteHeaderDispara(ByRef file As CIDSFileHandler, ByRef Filename As String, _
        ByRef strPatternID As String, ByVal Encrypt As Boolean)
        Dim strLine As String = "#File name: "
        strLine = strLine + Filename
        file.Write(strLine, Encrypt)
        strLine = "#"
        file.Write(strLine, Encrypt)
        strLine = "#Objective: Data format defination for Pattern data. ptp file is used to store pattern defination in txt files."
        file.Write(strLine, Encrypt)
        strLine = "#"
        file.Write(strLine, Encrypt)
        strLine = "#"
        file.Write(strLine, Encrypt)
        strLine = "#Note: 1, Empty line is allowed"
        file.Write(strLine, Encrypt)
        strLine = "#      2, Items are seperated by commas"
        file.Write(strLine, Encrypt)
        strLine = ""
        file.Write(strLine, Encrypt)
        strLine = "#Pattern starts"
        file.Write(strLine, Encrypt)
        strLine = "PatternID =" + strPatternID
        file.Write(strLine, Encrypt)
        strLine = ""
        file.Write(strLine, Encrypt)

        strLine = "#  0  1       2         3   4   5   6   7   8   9   10  11  12    13        14       15       16       17   18    19    20     21        22     23    24    25      26     27      28   29"
        file.Write(strLine, Encrypt)
        strLine = "#Cmd, Needle, Dispense, x1, y1, z1, x2, y2, z2, x3, y3, z3, Speed,NeedleGap,Duration,TrvDelay,RetDelay,AppH,RetV, RetH, ClearH,DTailDist,ArcRad,Pitch,FillH,RunTime,Sprial,RtAngle,Edge,IO"
        file.Write(strLine, Encrypt)
        strLine = "#"
        file.Write(strLine, Encrypt)
        strLine = ""
        file.Write(strLine, Encrypt)
        TraceGCCollect
    End Sub



    'Write activeX spreadsheet to pattern file                        
    '   File:       instance of file handler class                    
    '   Axspreadsheet:  instance of activeX spreadsheet               
    '   Encrypt:        file encrypt flag                             
    '   strPageName:    page name                                                                                                    '

    Private Sub WritePatternParaOpt(ByRef file As CIDSFileHandler, ByRef DataArray As ArrayList, _
            ByVal Encrypt As Boolean, ByVal strPageName As String)
        Dim strArray(5) As String
        Dim strLine As String
        Dim strTmp As String
        Dim ii As Integer = 0

        Dim i, j As Integer

        Dim SheetName As String

        Dim kk, UsedRowNumber As Integer
        Dim foundArrayName As Boolean

        'Save subpage
        kk = 0
        ii = 0

        SheetName = strPageName

        strLine = ""
        file.Write(strLine, Encrypt)
        strLine = "[Page=Main]"
        file.Write(strLine, Encrypt)

        kk = 0

        Dim row As Integer = DataArray.Count
        Dim dotrec As CIDSDot
        Dim heightrec As CIDSHeight
        Dim fidrec As CIDSFiducial
        For i = 1 To row

            strLine = ""
            Dim rec As Object = DataArray(i - 1)
            Dim Type As String = rec.CmdType

            If Type.ToUpper = "DOT" Then
                strLine = "Dot"
                dotrec = rec
                strLine = strLine + ","
                strLine = strLine + dotrec.Param.Needle + ","
                If dotrec.Param.DispenseOn Then
                    strLine = strLine + "On,"
                Else
                    strLine = strLine + "Off,"
                End If
                strLine = strLine + CStr(dotrec.PosX) + "," + CStr(dotrec.PosY) + "," + CStr(dotrec.PosZ) + ","
                strLine = strLine + "," + "," + ","
                strLine = strLine + "," + "," + ","
                strLine = strLine + ","
                strLine = strLine + CStr(dotrec.Param.NeedleGap) + ","
                strLine = strLine + CStr(dotrec.Param.Duration) + ","
                strLine = strLine + ","
                strLine = strLine + CStr(dotrec.Param.RetractDelay) + ","
                strLine = strLine + CStr(dotrec.Param.ApproachHeight) + ","
                strLine = strLine + CStr(dotrec.Param.RetractSpeed) + ","
                strLine = strLine + CStr(dotrec.Param.RetractHeight) + ","
                strLine = strLine + CStr(dotrec.Param.ClearanceHeight) + ","
                strLine = strLine + ","
                strLine = strLine + CStr(dotrec.Param.ArcRadius) + ","
                strLine = strLine + "," + ","
                strLine = strLine + ",,,,,,,,,,,,,,,,,,,,"
                strLine = strLine + ",,,,,,,,,,,,,,,,,,,,"
                strLine = strLine + ",,,,,,,,,,,,"
                file.Write(strLine, Encrypt)
            ElseIf Type.ToUpper = "FIDUCIAL" Then
                strLine = "Fiducial"
                fidrec = rec
                strLine = strLine + ","
                strLine = strLine + fidrec.FirstFid + ","
                strLine = strLine + fidrec.SecondFid + ","
                strLine = strLine + CStr(fidrec.PosX1) + "," + CStr(fidrec.PosY1) + "," + CStr(fidrec.PosZ1) + ","
                strLine = strLine + CStr(fidrec.PosX2) + "," + CStr(fidrec.PosY2) + "," + CStr(fidrec.PosZ2) + ","
                strLine = strLine + ",,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,"
                file.Write(strLine, Encrypt)
            ElseIf Type.ToUpper = "HEIGHT" Then
                strLine = "Height"
                heightrec = rec
                strLine = strLine + ",,,"
                strLine = strLine + CStr(heightrec.PosX) + "," + CStr(heightrec.PosY) + "," + CStr(heightrec.PosZ) + ","
                strLine = strLine + ",,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,"
                file.Write(strLine, Encrypt)
            End If


        Next i
        'Do
        '    kk += 1
        '    ii = ii + 1
        '    strLine = ""
        '    strTmp = .Cells(ii, gCommandNameColumn).Value

        '    If "" <> strTmp Then
        '        strLine = strLine + strTmp
        '   For j = gNeedleColumn To gWRMCoulumn
        '            strLine = strLine + "," + CStr(.Cells(ii, j).Value)
        '        Next
        '        file.Write(strLine, Encrypt)
        '    End If
        'Loop Until kk > UsedRowNumber   'Array lines will terminate the saving record.

        TraceGCCollect
    End Sub



    'Write activeX spreadsheet to pattern file                        
    '   File:       instance of file handler class                    
    '   Axspreadsheet:  instance of activeX spreadsheet               
    '   Encrypt:        file encrypt flag                             
    '   strPageName:    page name                                                                                                    '

    Private Sub WritePatternPara(ByRef file As CIDSFileHandler, ByRef AxSpreadsheet As AxOWC10.AxSpreadsheet, _
            ByVal Encrypt As Boolean, ByVal strPageName As String)
        Dim strArray(5) As String
        Dim strLine As String
        Dim strTmp As String
        Dim ii As Integer = 0

        Dim TotalSheets As Integer = AxSpreadsheet.ActiveWorkbook.Worksheets.Count()
        Dim i, j As Integer

        Dim SheetName As String

        Dim kk, UsedRowNumber As Integer
        Dim foundArrayName As Boolean

        'Save subpage
        kk = 0
        ii = 0

        SheetName = strPageName
        UsedRowNumber = AxSpreadsheet.Worksheets(SheetName).UsedRange.Rows.Count

        'UsedRowNumber = Spreadsheet_CountRowUsed(AxSpreadsheet, SheetName)

        strLine = ""
        file.Write(strLine, Encrypt)
        strLine = "[Page=Main]"
        file.Write(strLine, Encrypt)

        '''''''''''''''''
        '   Origin      '
        '''''''''''''''''
        'kk = 0
        'With AxSpreadsheet.Worksheets(SheetName)
        '    Do
        '        kk += 1
        '        ii = ii + 1
        '        strLine = ""
        '        strTmp = .Cells(ii, gCommandNameColumn).Value

        '        If "" <> strTmp Then
        '            strLine = strLine + strTmp
        '            For j = gNeedleColumn To gWRMCoulumn
        '                strLine = strLine + "," + CStr(.Cells(ii, j).Value)
        '            Next
        '            file.Write(strLine, Encrypt)
        '        End If
        '    Loop Until kk > UsedRowNumber   'excel lines will terminate the saving record.

        'End With

        ''''''''''''''''''''''''''''''
        '   Xue Wen                                                                                     '
        '   Modify some parts.                                                                          '
        '   Note: 1.) if user doesn't teach the program from "row 1", system can't read out all data.   '
        '         2.) Test more.                                                                        '
        ''''''''''''''''''''''''''''''
        Dim startCount As Boolean = False

        kk = 0
        ii = 1
        With AxSpreadsheet.Worksheets(SheetName)
            Do
                kk += 1
                strLine = ""
                strTmp = .Cells(kk, gCommandNameColumn).Value

                If "" <> strTmp Then
                    strLine = strLine + strTmp
                    For j = gNeedleColumn To gWRMCoulumn
                        strLine = strLine + "," + CStr(.Cells(kk, j).Value)
                    Next
                    file.Write(strLine, Encrypt)
                    ii = ii + 1
                    startCount = True
                Else
                    If (startCount = True) Then
                        ii = ii + 1
                    End If
                End If
            Loop Until (ii > UsedRowNumber)   'excel lines will terminate the saving record.

        End With


        'Save arrays attached to it if any
        Dim arrayDim As Integer
        For i = 1 To TotalSheets
            foundArrayName = False
            kk = 0
            ii = 0
            startCount = False

            SheetName = AxSpreadsheet.ActiveWorkbook.Worksheets(i).name()

            arrayDim = 1
            For j = 0 To Len(SheetName) - 1
                If "." = SheetName.Chars(j) Then
                    arrayDim += 1
                End If
            Next

            If 3 = arrayDim Then
                strTmp = file.ExtOnlyFromFullPath(SheetName)
                If strPageName + "." + strTmp = SheetName Then
                    foundArrayName = True
                End If
            ElseIf 1 = arrayDim Then
                strTmp = SheetName
                If "Main" = strPageName And "Main" <> SheetName Then
                    foundArrayName = True
                End If
            End If


            If foundArrayName = True Then

                strLine = ""
                file.Write(strLine, Encrypt)
                strLine = "[Page=" + strTmp + "]"
                file.Write(strLine, Encrypt)

                ii = 1
                With AxSpreadsheet.Worksheets(SheetName)
                    UsedRowNumber = .UsedRange.Rows.Count
                    Do
                        kk += 1
                        strLine = ""
                        strTmp = .Cells(kk, gCommandNameColumn).Value

                        If "" <> strTmp Then
                            strLine = strLine + strTmp
                            For j = gNeedleColumn To gWRMCoulumn
                                strLine = strLine + "," + CStr(.Cells(kk, j).Value)
                            Next
                            file.Write(strLine, Encrypt)

                            ii = ii + 1
                            startCount = True
                        Else
                            If (startCount = True) Then
                                ii = ii + 1
                            End If
                        End If
                    Loop Until (ii > UsedRowNumber)   'excel lines will terminate the saving record.

                End With
            End If
        Next i
        TraceGCCollect
    End Sub



    'Write activeX spreadsheet to pattern file                        
    '   File:       instance of file handler class                    
    '   Axspreadsheet:  instance of activeX spreadsheet               
    '   SelectFlag:     0=All data will be saved, 
    '                   1=Only selected data will be saved
    '   Encrypt:        file encrypt flag                             
    '                                                                

    Private Sub WritePatternParaAll(ByRef file As CIDSFileHandler, ByRef AxSpreadsheet As AxOWC10.AxSpreadsheet, _
            ByVal SelectFlag As Integer, ByVal Encrypt As Boolean)
        Dim strLine As String
        Dim strTmp As String
        Dim m_rowLocal As Integer = 0

        Dim sel As Microsoft.Office.Interop.OWC.Range = AxSpreadsheet.Selection

        Dim m_rowCount As Integer = sel.Rows.Count()
        Dim m_columnCount As Integer = sel.Columns.Count()
        Dim m_StartRow As Integer = sel.Row
        Dim m_columnStart As Integer = sel.Column


        If 0 = SelectFlag Then  'For Undo instead of Copy/Paste/Cut
            m_StartRow = 1
            m_rowCount = AxSpreadsheet.ActiveSheet.UsedRange.Rows.Count
        End If


        'Step1: Save the selected as Main
        strLine = ""
        file.Write(strLine, Encrypt)
        strLine = "[Page=" + "Main" + "]"
        file.Write(strLine, Encrypt)

        Dim kk As Integer = 0
        Dim i, j As Integer
        m_rowLocal = m_StartRow - 1
        With AxSpreadsheet.ActiveSheet
            Do
                kk += 1
                m_rowLocal = m_rowLocal + 1
                strLine = ""
                strTmp = .Cells(m_rowLocal, gCommandNameColumn).Value

                If "" <> strTmp Then
                    strLine = strLine + strTmp
                    For j = gNeedleColumn To gWRMCoulumn
                        strLine = strLine + "," + CStr(.Cells(m_rowLocal, j).Value)
                    Next
                    file.Write(strLine, Encrypt)
                End If
            Loop Until kk >= m_rowCount   'exceed lines will terminate the saving record.
        End With

        'Step2: Find necessary sub/array to be saved using connection checking.
        'It will update SubCallSheetName data structure
        If 1 = SelectFlag Then          'For Copy/Paste/Cut
            Spreadsheet_BuildRootConnectedArraySub(AxSpreadsheet)
        End If

        'Step3: Save the Sub and Array selected
        Dim TotalSheets As Integer = AxSpreadsheet.ActiveWorkbook.Worksheets.Count()

        Dim SheetName As String
        Dim sub_sheet As OWC10.Worksheet
        Dim UsedRowNumber As Integer
        Dim k As Integer
        Dim bSubOrArrayFound As Boolean

        Dim SelectedInSheetName As String = AxSpreadsheet.ActiveSheet.Name

        For i = TotalSheets To 1 Step -1

            kk = 0
            m_rowLocal = 0
            bSubOrArrayFound = False

            SheetName = AxSpreadsheet.ActiveWorkbook.Worksheets(i).name()

            'Find the sheet should be saved or not
            If 1 = SelectFlag Then              'For Copy/Paste/Cut
                For k = 0 To TotalSheets - 1
                    If SheetName = SubCallSheetName.SheetName(k).SheetName And _
                        SubCallSheetName.SheetName(k).SheetRootRelated > 0 Then
                        bSubOrArrayFound = True
                        Exit For
                    End If
                Next k
            Else                                'For Undo
                If SelectedInSheetName = SheetName Then
                    bSubOrArrayFound = False    'Exclude active sheet
                Else
                    bSubOrArrayFound = True
                End If
            End If


            If bSubOrArrayFound = True Then
                UsedRowNumber = Spreadsheet_CountRowUsed(AxSpreadsheet, SheetName)

                strLine = ""
                file.Write(strLine, Encrypt)
                strLine = "[Page=" + SheetName + "]"
                file.Write(strLine, Encrypt)

                With AxSpreadsheet.Worksheets(SheetName)
                    Do
                        kk += 1
                        m_rowLocal = m_rowLocal + 1
                        strLine = ""
                        strTmp = .Cells(m_rowLocal, gCommandNameColumn).Value

                        If "" <> strTmp Then
                            strLine = strLine + strTmp
                            For j = gNeedleColumn To gWRMCoulumn
                                strLine = strLine + "," + CStr(.Cells(m_rowLocal, j).Value)
                            Next

                            file.Write(strLine, Encrypt)
                        End If
                    Loop Until kk > UsedRowNumber   'Rows at which it will terminate the saving record.

                End With
            End If
        Next i
        TraceGCCollect
    End Sub



    'Load excel pattern file to activeX spreadsheet                   '
    '   Axspreadsheet:  instance of activeX spreadsheet               '
    '   File:       file name                                         '
    '   StartRow:   inserting row number                              '
    '   ForUndo:    undo flag                                         ' 
    '                                                                 '

    Function LoadExcelPatternFile(ByRef AxSpreadsheet As AxOWC10.AxSpreadsheet, ByVal file As String, _
            ByVal StartRow As Integer, ByVal ForUndo As Boolean) As Integer

        Dim objExcel As New Excel.ApplicationClass
        Dim objSheet As Excel.Worksheet
        Try
            'Create and open the correct excel workbook
            objExcel = CreateObject("Excel.Application")
            objExcel.Workbooks.Open(file, , True)
            objExcel.Visible = False
            objExcel.Parent.Windows(1).Visible = True
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

        'Select all the cells in the relevant worksheet
        'and copy the contents
        'objSheet = objExcel.ActiveSheet
        'objSheet.Cells.Select()
        'objSheet.Cells.Copy()

        Dim Rtn As Integer = 0
        Dim Filename As String
        Dim NumOfSubFilesToRead As Integer = 0

        Dim TotalSheets As Integer = objExcel.ActiveWorkbook.Worksheets.Count()
        Dim emptyLine As Integer = 0
        Dim UnitLine As Integer = 0
        Dim i As Integer
        Dim j As Integer
        Dim SheetName As String
        Dim sub_sheet As OWC10.Worksheet
        Dim m_rowLocal As Integer
        Dim iSourceRow As Integer
        Dim SubPageExist As Boolean

        Dim kk, UsedRowNumber As Integer
        Dim UnitInMain As String
        Dim UnitInSub As String


        'Load "Main" and its attached "Array"
        For i = 1 To TotalSheets
            kk = 0
            objSheet = objExcel.ActiveWorkbook.Worksheets(i)
            emptyLine = 0
            UnitLine = 0
            iSourceRow = 0
            SubPageExist = False

            SheetName = objExcel.ActiveWorkbook.Worksheets(i).name()

            'UsedRowNumber = Spreadsheet_CountRowUsed(AxSpreadsheet, SheetName)

            If "Main" <> SheetName Then
                'Check sub Page existance
                SubPageExist = Spreadsheet_CheckSubsheetExist(AxSpreadsheet, SheetName)

                If False = SubPageExist Then
                    AxSpreadsheet.Sheets.Add.Name = SheetName

                    AxSpreadsheet.ActiveWindow.FreezePanes = False
                    AxSpreadsheet.Worksheets(SheetName).Range("B1:B1").Select()
                    AxSpreadsheet.ActiveWindow.FreezePanes = True
                    m_rowLocal = 0
                End If
            Else
                m_rowLocal = StartRow - 1
                If m_rowLocal < 0 Then
                    m_rowLocal = 0
                End If
            End If

            If False = SubPageExist Then
                AxSpreadsheet.Worksheets(SheetName).Activate()

                With AxSpreadsheet.ActiveSheet
                    Do
                        kk += 1
                        m_rowLocal = m_rowLocal + 1
                        iSourceRow = iSourceRow + 1

                        If "" = CStr(objSheet.Cells(iSourceRow, gCommandNameColumn).Value) Then
                            emptyLine = emptyLine + 1
                        ElseIf "Unit" = CStr(objSheet.Cells(iSourceRow, gCommandNameColumn).Value) Then
                            UnitInMain = CStr(objSheet.Cells(iSourceRow, gDispensColumn).Value)
                            UnitLine = 1
                        Else
                            'Insert an empty line to avoid data overwrite
                            AxSpreadsheet.ActiveSheet.Cells(m_rowLocal, 1).Select()
                            AxSpreadsheet.Selection.EntireRow.Insert()
                            AxSpreadsheet.ActiveSheet.Rows(m_rowLocal).clear()

                            If "SubPattern" = CStr(objSheet.Cells(iSourceRow, gCommandNameColumn).Value) Then
                                SubFilesClaster(NumOfSubFilesToRead) = CStr(objSheet.Cells(iSourceRow, gSubnameColumn).Value)
                                NumOfSubFilesToRead += 1
                            End If

                            For j = 1 To gMaxColumns
                                .Cells(m_rowLocal - UnitLine, j) = objSheet.Cells(iSourceRow, j)
                            Next
                            'lim
                            For j = gMaxColumns + 1 To gWRMCoulumn
                                .Cells(m_rowLocal - UnitLine, j) = objSheet.Cells(iSourceRow, j)
                            Next
                            emptyLine = 0
                        End If
                    Loop Until emptyLine > 20 'excel lines will terminate the saving record.
                End With
            End If
        Next i

        objExcel.Quit()
        objExcel = Nothing
        TraceGCCollect

        'For undo operation, the whole sheets are loaded once.  It not need to check the subs
        'comes with.
        If ForUndo Then
            Return Rtn
        End If

        Dim strTmp As String
        'For all sub loadings if any
        Do Until (NumOfSubFilesToRead = 0 Or Rtn <> 0)
            Filename = SubFilesClaster(0)
            NumOfSubFilesToRead -= 1

            For i = 0 To NumOfSubFilesToRead
                SubFilesClaster(i) = SubFilesClaster(i + 1)
            Next
            SubFilesClaster(NumOfSubFilesToRead) = ""
            Rtn = LoadTxtPatternSingleFile(AxSpreadsheet, Filename, 1, 0, True, SubFilesClaster, NumOfSubFilesToRead)
        Loop


        Return Rtn  '0=no error
        TraceGCCollect
    End Function



    'Generate element record for displaying in spreadsheet            '
    '   Type:       element type                                      '
    '   CamPos:     camera position in work coord or reference coord  '
    '   return:     element record                                    ' 
    '                                                                 '

    Function AddCommand(ByVal Type As String, ByVal CamPos() As Double) As CIDSData.CIDSPatternExecution.CIDSTemplate
        Dim CmdStr As String = Type.ToUpper()
        Select Case CmdStr
            Case "WAIT"
                Return AddWait(CamPos)
            Case "MOVE"
                Return AddMove(CamPos)
            Case "DOT"
                Return AddDot(CamPos)
            Case "LINE"
                Return AddLine(CamPos)
            Case "ARC"
                Return AddArc(CamPos)
            Case "CIRCLE"
                Return AddCircle(CamPos)
            Case "FILLCIRCLE"
                Return AddFillCircle(CamPos)
            Case "RECTANGLE"
                Return AddRectangle(CamPos)
            Case "FILLRECTANGLE"
                Return AddFillRectangle(CamPos)
            Case "CHIPEDGE"
                Return AddChipEdge(CamPos)
            Case "PURGE"
                Return AddPurge()
            Case "VOLUMECALIBRATION"
                Return AddVolCal()
            Case "CLEAN"
                Return AddClean()
            Case "QC"
                Return AddQCCheck(CamPos)
            Case "LINKSTART", "LINKEND"
                Return AddLink()
            Case "SUBPATTERN"
                Return AddSubPattern(CamPos)
            Case "FIDUCIAL"
                Return AddFiducial(CamPos)
            Case "REFERENCE"
                Return AddReference(CamPos)
            Case "REJECT"
                Return AddReject(CamPos)
            Case "HEIGHT"
                Return AddHeight(CamPos)
            Case "ARRAY"
                Return AddArray(CamPos)
            Case "SETIO"
                Return SetIO()
            Case "GETIO"
                Return GetIO()
            Case "DOTARRAY"
                Return AddDotArray(CamPos)
        End Select
    End Function



    'Generate wait element record                                     '
    '   CamPos:     camera position in work coord or reference coord  '
    '   return:     element record                                    ' 
    '                                                                 '

    Public Function AddWait(ByVal CamPos() As Double) As CIDSData.CIDSPatternExecution.CIDSTemplate
        m_CurrentDPara.Pos1.X = CamPos(0)
        m_CurrentDPara.Pos1.Y = CamPos(1)
        m_CurrentDPara.Pos1.Z = CamPos(2)
        m_CurrentDPara.DispenseFlag = ""
        m_CurrentDPara.Needle = ""
        Return m_CurrentDPara
    End Function



    'Generate move element record                                     '
    '   CamPos:     camera position in work coord or reference coord  '
    '   return:     element record                                    ' 
    '                                                                 '

    Public Function AddMove(ByVal CamPos() As Double) As CIDSData.CIDSPatternExecution.CIDSTemplate
        m_CurrentDPara.Pos1.X = CamPos(0)
        m_CurrentDPara.Pos1.Y = CamPos(1)
        m_CurrentDPara.Pos1.Z = CamPos(2)
        m_CurrentDPara.DispenseFlag = "True"
        m_CurrentDPara.Needle = "Left"
        Return m_CurrentDPara
    End Function


    Public Function AddDotArray(ByVal CamPos() As Double) As CIDSData.CIDSPatternExecution.CIDSTemplate
        m_CurrentDPara.Pos1.X = CamPos(0)
        m_CurrentDPara.Pos1.Y = CamPos(1)
        m_CurrentDPara.Pos1.Z = CamPos(2)
        m_CurrentDPara.DispenseFlag = "On"
        m_CurrentDPara.Needle = "Left"
        Return m_CurrentDPara
    End Function


    'Generate dot element record                                      '
    '   CamPos:     camera position in work coord or reference coord  '
    '   return:     element record                                    ' 
    '                                                                 '

    Public Function AddDot(ByVal CamPos() As Double) As CIDSData.CIDSPatternExecution.CIDSTemplate
        m_CurrentDPara.Pos1.X = CamPos(0)
        m_CurrentDPara.Pos1.Y = CamPos(1)
        m_CurrentDPara.Pos1.Z = CamPos(2)
        m_CurrentDPara.DispenseFlag = "On"
        m_CurrentDPara.Needle = "Left"
        Return m_CurrentDPara
    End Function


    'Generate link element record                                     '
    '   return:     element record                                    ' 
    '                                                                 '

    Public Function AddLink() As CIDSData.CIDSPatternExecution.CIDSTemplate
        Return m_CurrentDPara
    End Function



    'Generate line element record                                     '
    '   CamPos:     camera position in work coord or reference coord  '
    '   return:     element record                                    ' 
    '                                                                 '

    Public Function AddLine(ByVal CamPos() As Double) As CIDSData.CIDSPatternExecution.CIDSTemplate
        m_CurrentDPara.Pos1.X = CamPos(0)
        m_CurrentDPara.Pos1.Y = CamPos(1)
        m_CurrentDPara.Pos1.Z = CamPos(2)
        m_CurrentDPara.DispenseFlag = "On"
        m_CurrentDPara.Needle = "Left"
        Return m_CurrentDPara
    End Function


    'Generate arc element record                                      '
    '   CamPos:     camera position in work coord or reference coord  '
    '   return:     element record                                    ' 
    '                                                                 '

    Public Function AddArc(ByVal CamPos() As Double) As CIDSData.CIDSPatternExecution.CIDSTemplate
        m_CurrentDPara.Pos1.X = CamPos(0)
        m_CurrentDPara.Pos1.Y = CamPos(1)
        m_CurrentDPara.Pos1.Z = CamPos(2)
        m_CurrentDPara.DispenseFlag = "True"
        m_CurrentDPara.Needle = "Left"
        Return m_CurrentDPara
    End Function



    'Generate circle element record                                   '
    '   CamPos:     camera position in work coord or reference coord  '
    '   return:     element record                                    ' 
    '                                                                 '

    Public Function AddCircle(ByVal CamPos() As Double) As CIDSData.CIDSPatternExecution.CIDSTemplate
        m_CurrentDPara.Pos1.X = CamPos(0)
        m_CurrentDPara.Pos1.Y = CamPos(1)
        m_CurrentDPara.Pos1.Z = CamPos(2)
        m_CurrentDPara.DispenseFlag = "True"
        m_CurrentDPara.Needle = "Left"
        Return m_CurrentDPara
    End Function



    'Generate fill circle element record                              '
    '   CamPos:     camera position in work coord or reference coord  '
    '   return:     element record                                    ' 
    '                                                                 '

    Public Function AddFillCircle(ByVal CamPos() As Double) As CIDSData.CIDSPatternExecution.CIDSTemplate
        m_CurrentDPara.Pos1.X = CamPos(0)
        m_CurrentDPara.Pos1.Y = CamPos(1)
        m_CurrentDPara.Pos1.Z = CamPos(2)
        m_CurrentDPara.DispenseFlag = "True"
        m_CurrentDPara.Needle = "Left"
        Return m_CurrentDPara
    End Function



    'Generate rectangle element record                                '
    '   CamPos:     camera position in work coord or reference coord  '
    '   return:     element record                                    ' 
    '                                                                 '

    Public Function AddRectangle(ByVal CamPos() As Double) As CIDSData.CIDSPatternExecution.CIDSTemplate
        m_CurrentDPara.Pos1.X = CamPos(0)
        m_CurrentDPara.Pos1.Y = CamPos(1)
        m_CurrentDPara.Pos1.Z = CamPos(2)
        m_CurrentDPara.DispenseFlag = "True"
        m_CurrentDPara.Needle = "Left"
        Return m_CurrentDPara
    End Function



    'Generate fill rectangle element record                           '
    '   CamPos:     camera position in work coord or reference coord  '
    '   return:     element record                                    ' 
    '                                                                 '

    Public Function AddFillRectangle(ByVal CamPos() As Double) As CIDSData.CIDSPatternExecution.CIDSTemplate
        m_CurrentDPara.Pos1.X = CamPos(0)
        m_CurrentDPara.Pos1.Y = CamPos(1)
        m_CurrentDPara.Pos1.Z = CamPos(2)
        m_CurrentDPara.DispenseFlag = "True"
        m_CurrentDPara.Needle = "Left"
        Return m_CurrentDPara
    End Function



    'Generate chip edge element record                                '
    '   CamPos:     camera position in work coord or reference coord  '
    '   return:     element record                                    ' 
    '                                                                 '

    Public Function AddChipEdge(ByVal CamPos() As Double) As CIDSData.CIDSPatternExecution.CIDSTemplate
        Dim CmdStr As String = ""
        m_CurrentDPara.Pos1.X = CamPos(0)
        m_CurrentDPara.Pos1.Y = CamPos(1)
        m_CurrentDPara.Pos1.Z = CamPos(2)
        m_CurrentDPara.DispenseFlag = "True"
        m_CurrentDPara.Needle = "Left"
        Return m_CurrentDPara
    End Function



    'Generate QC element record                                       '
    '   CamPos:     camera position in work coord or reference coord  '
    '   return:     element record                                    ' 
    '                                                                 '

    Public Function AddQCCheck(ByVal CamPos() As Double) As CIDSData.CIDSPatternExecution.CIDSTemplate
        m_CurrentDPara.Pos1.X = CamPos(0)
        m_CurrentDPara.Pos1.Y = CamPos(1)
        m_CurrentDPara.Pos1.Z = CamPos(2)
        m_CurrentDPara.DispenseFlag = "On"
        m_CurrentDPara.Needle = "Left"
        Return m_CurrentDPara
    End Function



    'Generate purge element record                                    '
    '   return:     element record                                    ' 
    '                                                                 '

    Public Function AddPurge() As CIDSData.CIDSPatternExecution.CIDSTemplate
        m_CurrentDPara.Pos1.X = Nothing
        m_CurrentDPara.Pos1.Y = Nothing
        m_CurrentDPara.Pos1.Z = Nothing
        'm_CurrentDPara.DispenseDuration = 0
        m_CurrentDPara.Needle = "Left"
        Return m_CurrentDPara
    End Function


    'Generate clean element record                                    '
    '   CamPos:     camera position in work coord or reference coord  '
    '   return:     element record                                    ' 
    '                                                                 '

    Public Function AddVolCal() As CIDSData.CIDSPatternExecution.CIDSTemplate
        Dim CmdStr As String = ""
        m_CurrentDPara.Pos1.X = Nothing
        m_CurrentDPara.Pos1.Y = Nothing
        m_CurrentDPara.Pos1.Z = Nothing
        m_CurrentDPara.DispenseFlag = ""
        m_CurrentDPara.Needle = ""
        Return m_CurrentDPara
    End Function

    'Generate clean element record                                    '
    '   CamPos:     camera position in work coord or reference coord  '
    '   return:     element record                                    ' 
    '                                                                 '

    Public Function AddClean() As CIDSData.CIDSPatternExecution.CIDSTemplate
        Dim CmdStr As String = ""
        m_CurrentDPara.Pos1.X = Nothing
        m_CurrentDPara.Pos1.Y = Nothing
        m_CurrentDPara.Pos1.Z = Nothing
        m_CurrentDPara.DispenseFlag = ""
        m_CurrentDPara.Needle = ""
        Return m_CurrentDPara
    End Function



    'Generate call sub element record                                 '
    '   CamPos:     camera position in work coord or reference coord  '
    '   return:     element record                                    ' 
    '                                                                 '

    Public Function AddSubPattern(ByVal CamPos() As Double) As CIDSData.CIDSPatternExecution.CIDSTemplate
        m_CurrentDPara.Pos1.X = CamPos(0)
        m_CurrentDPara.Pos1.Y = CamPos(1)
        m_CurrentDPara.Pos1.Z = CamPos(2)
        m_CurrentDPara.Needle = ""
        m_CurrentDPara.RotatingAngle = 0.0
        m_CurrentDPara.SubFileName = "Untitle"
        Return m_CurrentDPara
    End Function



    'Generate fiducial element record                                 '
    '   CamPos:     camera position in work coord or reference coord  '
    '   return:     element record                                    ' 
    '                                                                 '

    Public Function AddFiducial(ByVal CamPos() As Double) As CIDSData.CIDSPatternExecution.CIDSTemplate
        m_CurrentDPara.Pos1.X = CamPos(0)
        m_CurrentDPara.Pos1.Y = CamPos(1)
        m_CurrentDPara.Pos1.Z = CamPos(2)
        m_CurrentDPara.Needle = ""
        m_CurrentDPara.DispenseFlag = ""
        m_CurrentDPara.SubFileName = ""
        Return m_CurrentDPara
    End Function



    'Generate reference element record                                '
    '   CamPos:     camera position in work coord or reference coord  '
    '   return:     element record                                    ' 
    '                                                                 '

    Public Function AddReference(ByVal CamPos() As Double) As CIDSData.CIDSPatternExecution.CIDSTemplate
        m_CurrentDPara.Pos1.X = CamPos(0)
        m_CurrentDPara.Pos1.Y = CamPos(1)
        m_CurrentDPara.Pos1.Z = CamPos(2)
        m_CurrentDPara.Needle = ""
        m_CurrentDPara.DispenseFlag = ""
        Return m_CurrentDPara
    End Function



    'Generate reject element record                                   '
    '   CamPos:     camera position in work coord or reference coord  '
    '   return:     element record                                    ' 
    '                                                                 '

    Public Function AddReject(ByVal CamPos() As Double) As CIDSData.CIDSPatternExecution.CIDSTemplate
        m_CurrentDPara.Pos1.X = CamPos(0)
        m_CurrentDPara.Pos1.Y = CamPos(1)
        m_CurrentDPara.Pos1.Z = CamPos(2)
        m_CurrentDPara.Needle = ""
        m_CurrentDPara.DispenseFlag = ""
        Return m_CurrentDPara
    End Function



    'Generate height element record                                   '
    '   CamPos:     camera position in work coord or reference coord  '
    '   return:     element record                                    ' 
    '                                                                 '

    Public Function AddHeight(ByVal CamPos() As Double) As CIDSData.CIDSPatternExecution.CIDSTemplate
        m_CurrentDPara.Pos1.X = CamPos(0)
        m_CurrentDPara.Pos1.Y = CamPos(1)
        m_CurrentDPara.Pos1.Z = CamPos(2)
        m_CurrentDPara.Needle = ""
        m_CurrentDPara.DispenseFlag = ""
        Return m_CurrentDPara
    End Function



    'Generate array element record                                    '
    '   return:     element record                                    ' 
    '                                                                 '

    Public Function AddArray(ByVal CamPos() As Double) As CIDSData.CIDSPatternExecution.CIDSTemplate
        m_CurrentDPara.Pos1.X = CamPos(0)
        m_CurrentDPara.Pos1.Y = CamPos(1)
        m_CurrentDPara.Pos1.Z = CamPos(2)
        m_CurrentDPara.DispenseDuration = 0
        m_CurrentDPara.Needle = "Left"
        Return m_CurrentDPara
    End Function



    'Generate setio element record                                    '
    '   return:     element record                                    ' 
    '                                                                 '

    Public Function SetIO() As CIDSData.CIDSPatternExecution.CIDSTemplate
        Dim CmdStr As String = ""
        m_CurrentDPara.Pos1.X = Nothing
        m_CurrentDPara.Pos1.Y = Nothing
        m_CurrentDPara.Pos1.Z = Nothing
        m_CurrentDPara.Needle = "Left"
        Return m_CurrentDPara
    End Function



    'Generate getio element record                                    '
    '   CamPos:     camera position in work coord or reference coord  '
    '   return:     element record                                    ' 
    '                                                                 '

    Public Function GetIO() As CIDSData.CIDSPatternExecution.CIDSTemplate
        m_CurrentDPara.Pos1.X = Nothing
        m_CurrentDPara.Pos1.Y = Nothing
        m_CurrentDPara.Pos1.Z = Nothing
        Return m_CurrentDPara
    End Function

End Class


'''''''''''''''''''''''''''''''
'                                                                                                '
'Class CIDSPattern                                                                               '
'  Despription:                                                                                  '  
'           Pattern file handling class                                                          '    
'  Created by:                                                                                   ' 
'           Shen Jian                                                                            '
'  Modified by:                                                                                  ' 
'           Jinag Ting Ying                                                                      '
'                                                                                                '
'''''''''''''''''''''''''''''''
Public Class CIDSFileHandler
    Protected m_FreeFileNo As Integer


    'open file for reading                                            '
    '   Filename:     filename to be opened                           '
    '   ext:          file extension name                             ' 
    '                                                                 '

    Public Function Open(ByVal Filemane As String, ByVal ext As String) As Integer
        If ext = "ptp" Then
            Dim ret As Integer = 0
            Try
                Dim file As String
                file = Filemane & "." & ext
                m_FreeFileNo = FreeFile()
                FileOpen(m_FreeFileNo, file, OpenMode.Input, OpenAccess.Read)

            Catch ex As Exception
                Return 1
            End Try
        End If
    End Function



    'open file for reading                                            '
    '   Filename:     filename with extension to be opened            '
    '                                                                 '

    Public Function OpenR(ByVal Filemane As String) As Integer
        Dim ret As Integer = 0
        Try
            m_FreeFileNo = FreeFile()
            FileOpen(m_FreeFileNo, Filemane, OpenMode.Input, OpenAccess.Read)
        Catch ex As Exception
            Return 1
        End Try
    End Function



    'open file for writing                                            '
    '   Filename:     filename with extention to be opened            '
    '                                                                 '

    Public Function OpenW(ByVal Filemane As String) As Integer
        Dim ret As Integer = 0
        Try
            m_FreeFileNo = FreeFile()
            FileOpen(m_FreeFileNo, Filemane, OpenMode.Output, OpenAccess.Write)
        Catch ex As Exception
            Return 1
        End Try
    End Function



    'Read file one line data                                          '
    '   Encrypt:          true:encrypt, false:non encrypt             ' 
    '                                                                 '

    Public Function Read(ByVal Encrypt As Boolean) As String
        If Not EOF(m_FreeFileNo) Then
            If Encrypt Then
                Return TextCrypt(LineInput(m_FreeFileNo))
            Else
                Return LineInput(m_FreeFileNo)
            End If

        Else
            Return ("EOF")
        End If
    End Function



    'Write one line data                                              '
    '   Encrypt:          true:encrypt, false:non encrypt             ' 
    '                                                                 '

    Public Function Write(ByVal line As String, ByVal Encrypt As Boolean) As Integer 'Jiang
        If Encrypt Then
            line = TextCrypt(line)
        End If
        PrintLine(m_FreeFileNo, line)
    End Function


    'Close file                                                       '
    '                                                                 '

    Public Function Close() As Integer
        FileClose(m_FreeFileNo)
    End Function



    'extract folder and file name                                     '
    '   FileFullPath:     file name with folder name                  ' 
    '   return:           folder name with file name                  '
    '                                                                 '

    Public Function FolderWithNameFromFileName(ByVal FileFullPath As String) As String
        'EXAMPLE: input ="C:\winnt\system32\kernel.dll, 
        'output = C:\winnt\system32\kernel
        Try
            Dim intPos As Integer
            intPos = FileFullPath.LastIndexOfAny(".")
            Return FileFullPath.Substring(0, intPos)
        Catch ex As Exception
            ExceptionDisplay(ex)
        End Try

    End Function



    'extract folder name                                              '
    '   FileFullPath:     file name with folder name                  ' 
    '   return:           folder name                                 '
    '                                                                 '

    Public Function FolderFromFileName(ByVal FileFullPath As String) As String
        'EXAMPLE: input ="C:\winnt\system32\kernel.dll, 
        'output = C:\winnt\system32\

        Dim intPos As Integer
        intPos = FileFullPath.LastIndexOfAny("\")
        intPos += 1

        Return FileFullPath.Substring(0, intPos)
    End Function



    'extract file name                                                '
    '   FileFullPath:     file name with folder name                  ' 
    '   return:           file name                                   '
    '                                                                 '

    Public Function NameOnlyFromFullPath(ByVal FileFullPath As String) As String

        'EXAMPLE: input ="C:\winnt\system32\kernel.dll, 
        'output = kernel.dll

        Dim intPos As Integer

        intPos = FileFullPath.LastIndexOfAny("\")
        intPos += 1

        Return FileFullPath.Substring(intPos, (Len(FileFullPath) - intPos))
    End Function



    'extract extention name                                           '
    '   FileFullPath:     file name with folder name                  ' 
    '   return:           extention name                              '
    '                                                                 '

    Public Function ExtOnlyFromFullPath(ByVal FileFullPath As String) As String

        'EXAMPLE: input ="C:\winnt\system32\kernel.dll, 
        'output = dll

        Dim intPos As Integer

        intPos = FileFullPath.LastIndexOfAny(".")
        intPos += 1

        Return FileFullPath.Substring(intPos, (Len(FileFullPath) - intPos))

    End Function



    'Text encrypt                                                     '
    '   Text:     text to be encrypted                                ' 
    '   return:   encrypted text                                      '
    '                                                                 '

    Public Function TextCrypt( _
        ByVal Text As String) As String
        ' Encrypts/decrypts the passed string using
        ' a simple ASCII value-swapping algorithm
        Dim strTempChar As String, i As Integer
        For i = 1 To Len(Text)
            If Asc(Mid$(Text, i, 1)) < 128 Then
                strTempChar = CType(Asc(Mid$(Text, i, 1)) + 128, String)
            ElseIf Asc(Mid$(Text, i, 1)) > 128 Then
                strTempChar = CType(Asc(Mid$(Text, i, 1)) - 128, String)
            End If
            Mid$(Text, i, 1) = Chr(CType(strTempChar, Integer))
        Next i
        Return Text
    End Function

    'How to use
    'Dim MyText As String
    '    ' Encrypt
    'MyText = "Karl Moore"
    'MyText = Crypt(MyText)
    'MessageBox.Show(MyText)
    '    ' Decrypt
    'MyText = Crypt(MyText)
    'MessageBox.Show(MyText)

End Class


'''''''''''''''''''''''''''''''
'                                                                                                '
'Class CIDSItemsLUT                                                                              '
'  Despription:                                                                                  '  
'           Fast indexing for Spreadsheet cell filling: indentify the part of cells no           '
'           filling data                                                                         '
'  Created by:                                                                                   ' 
'           Jiang Ting Ying                                                                      '
'                                                                                                '
'''''''''''''''''''''''''''''''
Public Class CIDSItemsLUT
    Const MaxNumberOfCmd As Integer = 24
    Const MaxNumberOfPar As Integer = 30

    'It will have a cooresponding command
    Dim CurrentCmdIndex As Integer


    Dim theCmd() As String = {"FIDUCIAL", "REFERENCE", "HEIGHT", "REJECT", _
        "DOT", "DOTARRAY", "LINE", "ARC", "LINK", "RECTANGLE", "CIRCLE", "FILLRECTANGLE", _
        "FILLCIRCLE", "CHIPEDGE", "QC", "ARRAY", "SUBPATTERN", "MOVE", "WAIT", "PURGE", _
        "CLEAN", "VOLUMECALIBRATION", "SETIO", "GETIO", "ITEMNO"}

    'The last column is ITEMNO, will not be used

    '                               0,           1
    '   --------->Cmd (for example: "FIDUCIAL", "REFERENCE",...)
    '   |
    '   |
    '   |
    '   V
    ' Data (for example: gCommandNameColumn, 
    '                    gNeedleColumn, 
    '                    gSubnameColumn, 
    '                    X1, 
    '                    Y1,
    '                    ...)
    '

    '                               |------------------- FILLRECTANGLE
    '                               | |------------------FILLCIRCLE
    '                               | |
    '                               | |           |----- SUBPATTERN
    '                               | |           |
    '0, 1, 2, 3, 4, 5, 6, 7, 8, 9,10,11,12,13,14,15,16,17,18,19,20,21,22,23
    Dim theFlag(,) As Byte = { _
    {1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 0}, _
    {0, 0, 0, 0, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 0, 0, 0, 1, 1, 1, 1, 1, 0, 0, 1}, _
    {0, 0, 0, 0, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 0, 0, 1, 0, 0, 0, 0, 0, 1, 0, 2}, _
    {1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 0, 1, 1, 1, 0, 0, 0, 0, 0, 3}, _
    {1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 0, 1, 1, 1, 0, 0, 0, 0, 0, 4}, _
    {1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 0, 1, 1, 1, 0, 0, 0, 0, 0, 5}, _
    {1, 0, 0, 0, 0, 1, 1, 1, 1, 1, 1, 1, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 6}, _
    {1, 0, 0, 0, 0, 1, 1, 1, 1, 1, 1, 1, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 7}, _
    {1, 0, 0, 0, 0, 1, 1, 1, 1, 1, 1, 1, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 8}, _
    {0, 0, 0, 0, 0, 1, 0, 1, 1, 1, 1, 1, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 9}, _
    {0, 0, 0, 0, 0, 1, 0, 1, 1, 1, 1, 1, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 10}, _
    {0, 0, 0, 0, 0, 1, 0, 1, 1, 1, 1, 1, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 11}, _
    {0, 0, 0, 0, 0, 1, 1, 1, 1, 1, 1, 1, 1, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 12}, _
    {0, 0, 0, 0, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 13}, _
    {0, 0, 0, 0, 1, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 0, 0, 0, 14}, _
    {0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 1, 1, 1, 1, 0, 0, 0, 0, 0, 1, 1, 0, 0, 0, 15}, _
    {0, 0, 0, 0, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 16}, _
    {0, 0, 0, 0, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 17}, _
    {0, 0, 0, 0, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 18}, _
    {0, 0, 0, 0, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 19}, _
    {0, 0, 0, 0, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 20}, _
    {0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 1, 1, 1, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 21}, _
    {0, 0, 0, 0, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 0, 0, 0, 1, 0, 0, 0, 0, 0, 0, 22}, _
    {0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 23}, _
    {0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 24}, _
    {0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 25}, _
    {0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 26}, _
    {0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1, 0, 0, 0, 0, 0, 0, 0, 27}, _
    {0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 28}, _
    {0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 29}}



    'Convert a command to index to speedup running speed                          
    '   Cmd:  the input command     
    '   return:   None                                                                                    

    Public Function Cmd2Index(ByVal Cmd As String) As Byte
        Dim i As Integer
        Dim CmdStr As String = Cmd.ToUpper()

        For i = 0 To MaxNumberOfCmd - 1
            If theCmd(i) = CmdStr Then
                Exit For
            End If
        Next

        CurrentCmdIndex = i
    End Function


    'Check the flag setting                          
    '   Parameter:  the element in the Spreadsheet     
    '   return:   True=The cell should have data                                                                                    

    Public Function GetFlag(ByVal Parameter As Integer) As Byte
        Dim rtn As Byte = 0

        'Indexes need to be inside the LUT searching range
        If (CurrentCmdIndex < MaxNumberOfCmd) And (Parameter - 1 < MaxNumberOfPar) Then
            rtn = theFlag(Parameter - 1, CurrentCmdIndex)
        End If

        Return rtn
    End Function

End Class

Public Class ArrayCoordinate

    Dim ArrayType As String

    Dim Pos1 As xyzCoordinate
    Dim Pos2 As xyzCoordinate
    Dim Pos3 As xyzCoordinate
    Dim Pos4 As xyzCoordinate
    Dim Pos5 As xyzCoordinate
    Dim Pos6 As xyzCoordinate
    Dim Pos7 As xyzCoordinate
    Dim Pos8 As xyzCoordinate
    Dim Pos9 As xyzCoordinate

    Dim XhDelta As Double
    Dim XvDelta As Double
    Dim YhDelta As Double
    Dim YvDelta As Double

    Dim xNumber As Integer
    Dim yNumber As Integer

    Public Sub CalculateStride()

    End Sub


End Class




'Pattern error checking class in file: IDSPattern.vb
Structure WorkAreaDef
    Dim X As Double
    Dim Y As Double
    Dim XMin As Double
    Dim XMax As Double
    Dim YMin As Double
    Dim YMax As Double
    Dim ZMin As Double
    Dim ZMax As Double
End Structure

Structure SystemOriginDef
    Dim X As Double
    Dim Y As Double
    Dim Z As Double
End Structure


'''''''''''''''''''''''''''''''
'                                                                                                
'Class CIDSErrorCheck                                                                            
'  Despription:                                                                                   
'           Pattern data error checking handling class, unit change is also included               
'                                                                                               
'''''''''''''''''''''''''''''''
Public Class CIDSErrorCheck
    Dim SystemOrigin As SystemOriginDef
    Dim WorkArea As WorkAreaDef
    Dim MinSpeedLimit As Double = 0
    Dim MaxSpeedLimit As Double

    Dim ReferencePtX As Double = 0
    Dim ReferencePtY As Double = 0
    Dim ReferencePtZ As Double = 0

    Dim Record As CIDSPattern.PatternRecord

    'Get position coordination limits for error checking 

    Public Sub GetErrorCheckParameter()

        WorkArea.XMin = gWorkLimitXmin  'SystemOrigin.X
        WorkArea.XMax = gWorkLimitXmax  'IDSData.Hardware.Gantry.WorkArea.X + SystemOrigin.X
        WorkArea.YMin = gWorkLimitYmin  'SystemOrigin.Y
        WorkArea.YMax = gWorkLimitYmax  'IDSData.Hardware.Gantry.WorkArea.Y + SystemOrigin.Y
        WorkArea.ZMin = (gWorkLimitZmin - gLeftNeedleOffs(2))  'SystemOrigin.Z + gRightNeedleOffs(2)
        WorkArea.ZMax = (gWorkLimitZmax - gLeftNeedleOffs(2))  'IDSData.Hardware.Gantry.WorkArea.Z.Max + SystemOrigin.Z + gRightNeedleOffs(2)

        MinSpeedLimit = 0
        MaxSpeedLimit = IDS.Data.Hardware.Gantry.MaxSpeedLimit  ' IDSData.Hardware.Gantry.MaxSpeedLimit


        'Set test value for debug before system paras are ready
        'WorkArea.XMin = -150
        'WorkArea.XMax = 150
        'WorkArea.YMin = -150
        'WorkArea.YMax = 150
        'WorkArea.ZMin = -10
        'WorkArea.ZMax = 60

        'MaxSpeedLimit = 1000
        TraceGCCollect
    End Sub


    'Convert external Excel pattern file to internal data record
    Public Function ConvertToDataStruct(ByRef Rec As CIDSPattern.PatternRecord, ByRef sheet As AxOWC10.AxSpreadsheet, ByVal SheetName As String, ByVal CurrentRow As Integer) As Boolean
        Dim ValidData As New CIDSItemsLUT

        Dim Rtn As Boolean = False

        Rec.pPara.Name = sheet.ActiveSheet.Cells(CurrentRow, gCommandNameColumn).Value
        If "" = Rec.pPara.Name Then
            Return Rtn
        Else
            ValidData.Cmd2Index(Rec.pPara.Name)
        End If


        With sheet.ActiveSheet
            If ValidData.GetFlag(gNeedleColumn) Then
                Rec.pPara.Needle = .Cells(CurrentRow, gNeedleColumn).Value
            End If
            If ValidData.GetFlag(gDispensColumn) Then
                Rec.pPara.DispenseFlag = .Cells(CurrentRow, gDispensColumn).Value
            End If
            If ValidData.GetFlag(gPos1XColumn) Then
                Rec.pPara.Pos1.X = .Cells(CurrentRow, gPos1XColumn).Value
                Rec.pPara.Pos1.Y = .Cells(CurrentRow, gPos1YColumn).Value
                Rec.pPara.Pos1.Z = .Cells(CurrentRow, gPos1ZColumn).Value
            End If
            If ValidData.GetFlag(gPos2XColumn) Then
                Rec.pPara.Pos2.X = .Cells(CurrentRow, gPos2XColumn).Value
                Rec.pPara.Pos2.Y = .Cells(CurrentRow, gPos2YColumn).Value
                Rec.pPara.Pos2.Z = .Cells(CurrentRow, gPos2ZColumn).Value
            End If
            If ValidData.GetFlag(gPos3XColumn) Then
                Rec.pPara.Pos3.X = .Cells(CurrentRow, gPos3XColumn).Value
                Rec.pPara.Pos3.Y = .Cells(CurrentRow, gPos3YColumn).Value
                Rec.pPara.Pos3.Z = .Cells(CurrentRow, gPos3ZColumn).Value
            End If
            If ValidData.GetFlag(gTravelSpeedColumn) Then
                If Rec.pPara.Name = "DotArray" Then
                    Rec.pPara.RowAndColumn = .Cells(CurrentRow, gDotArrayRowsColumn).Value
                Else
                    Rec.pPara.TravelSpeed = .Cells(CurrentRow, gTravelSpeedColumn).Value
                End If
            End If
            If ValidData.GetFlag(gNeedleGapColumn) Then
                Rec.pPara.NeedleGap = .Cells(CurrentRow, gNeedleGapColumn).Value
            End If
            If ValidData.GetFlag(gDurationColumn) Then
                Rec.pPara.DispenseDuration = .Cells(CurrentRow, gDurationColumn).Value
            End If
            If ValidData.GetFlag(gTravelDelayColumn) Then
                Rec.pPara.TravelDelay = .Cells(CurrentRow, gTravelDelayColumn).Value
            End If
            If m_Execution.m_ItemLUT.GetFlag(gRetractDelayColumn) Then
                Rec.pPara.RetractDelay = .Cells(CurrentRow, gRetractDelayColumn).Value
            End If
            If ValidData.GetFlag(gApproachHtColumn) Then
                Rec.pPara.ApproachDispHeight = .Cells(CurrentRow, gApproachHtColumn).Value
            End If
            If ValidData.GetFlag(gRetractSpeedColumn) Then
                Rec.pPara.RetractSpeed = .Cells(CurrentRow, gRetractSpeedColumn).Value
            End If
            If ValidData.GetFlag(gRetractHtColumn) Then
                Rec.pPara.RetractHeight = .Cells(CurrentRow, gRetractHtColumn).Value
            End If
            If ValidData.GetFlag(gClearanceHtColumn) Then
                Rec.pPara.ClearanceHeight = .Cells(CurrentRow, gClearanceHtColumn).Value
            End If
            If ValidData.GetFlag(gDTaiDistColumn) Then
                Rec.pPara.DetailingDistance = .Cells(CurrentRow, gDTaiDistColumn).Value
            End If
            If ValidData.GetFlag(gArcRadiusColumn) Then
                Rec.pPara.ArcRadius = .Cells(CurrentRow, gArcRadiusColumn).Value
            End If
            If ValidData.GetFlag(gPitchColumn) Then
                Rec.pPara.Pitch = .Cells(CurrentRow, gPitchColumn).Value
            End If
            If ValidData.GetFlag(gFillHeightColumn) Then
                Rec.pPara.FilledHeight = .Cells(CurrentRow, gFillHeightColumn).Value
            End If
            If ValidData.GetFlag(gNoOfRunColumn) Then
                Rec.pPara.NumberOfRun = .Cells(CurrentRow, gNoOfRunColumn).Value
            End If
            If ValidData.GetFlag(gSprialColumn) Then
                Rec.pPara.SpiralFlag = .Cells(CurrentRow, gSprialColumn).Value
            End If
            If ValidData.GetFlag(gRtAngleColumn) Then
                Rec.pPara.RotatingAngle = .Cells(CurrentRow, gRtAngleColumn).Value
            End If
            If ValidData.GetFlag(gEdgeClearColumn) Then
                Rec.pPara.EdgeClear = .Cells(CurrentRow, gEdgeClearColumn).Value
            End If
            If ValidData.GetFlag(gIONumberColumn) Then
                Rec.pPara.IONumber = .Cells(CurrentRow, gIONumberColumn).Value
            End If
        End With
        Rtn = True

        Return Rtn
        TraceGCCollect
    End Function


    Public Function ConvertFromDataStruct(ByRef Rec As CIDSPattern.PatternRecord, ByRef sheet As AxOWC10.AxSpreadsheet, ByVal CurrentRow As Integer) As Boolean
        Dim ValidData As New CIDSItemsLUT

        Dim Rtn As Boolean = False

        If "" = Rec.pPara.Name Then
            Return Rtn
        Else
            ValidData.Cmd2Index(Rec.pPara.Name)
        End If

        With sheet.ActiveSheet
            sheet.ActiveSheet.Cells(CurrentRow, gCommandNameColumn).Value = Rec.pPara.Name

            If ValidData.GetFlag(gNeedleColumn) Then
                .Cells(CurrentRow, gNeedleColumn).Value = Rec.pPara.Needle
            End If
            If ValidData.GetFlag(gDispensColumn) Then
                .Cells(CurrentRow, gDispensColumn).Value = Rec.pPara.DispenseFlag
            End If
            If ValidData.GetFlag(gPos1XColumn) Then
                .Cells(CurrentRow, gPos1XColumn).Value = Rec.pPara.Pos1.X
                .Cells(CurrentRow, gPos1YColumn).Value = Rec.pPara.Pos1.Y
                .Cells(CurrentRow, gPos1ZColumn).Value = Rec.pPara.Pos1.Z
            End If
            If ValidData.GetFlag(gPos2XColumn) Then
                .Cells(CurrentRow, gPos2XColumn).Value = Rec.pPara.Pos2.X
                .Cells(CurrentRow, gPos2YColumn).Value = Rec.pPara.Pos2.Y
                .Cells(CurrentRow, gPos2ZColumn).Value = Rec.pPara.Pos2.Z
            End If
            If ValidData.GetFlag(gPos3XColumn) Then
                .Cells(CurrentRow, gPos3XColumn).Value = Rec.pPara.Pos3.X
                .Cells(CurrentRow, gPos3YColumn).Value = Rec.pPara.Pos3.Y
                .Cells(CurrentRow, gPos3ZColumn).Value = Rec.pPara.Pos3.Z
            End If
            If ValidData.GetFlag(gTravelSpeedColumn) Then
                .Cells(CurrentRow, gTravelSpeedColumn).Value = Rec.pPara.TravelSpeed
            End If
            If ValidData.GetFlag(gNeedleGapColumn) Then
                .Cells(CurrentRow, gNeedleGapColumn).Value = Rec.pPara.NeedleGap
            End If
            If ValidData.GetFlag(gDurationColumn) Then
                .Cells(CurrentRow, gDurationColumn).Value = Rec.pPara.DispenseDuration
            End If
            If ValidData.GetFlag(gTravelDelayColumn) Then
                .Cells(CurrentRow, gTravelDelayColumn).Value = Rec.pPara.TravelDelay
            End If
            If ValidData.GetFlag(gApproachHtColumn) Then
                .Cells(CurrentRow, gApproachHtColumn).Value = Rec.pPara.ApproachDispHeight
            End If
            If ValidData.GetFlag(gRetractSpeedColumn) Then
                .Cells(CurrentRow, gRetractSpeedColumn).Value = Rec.pPara.RetractSpeed
            End If
            If ValidData.GetFlag(gRetractHtColumn) Then
                .Cells(CurrentRow, gRetractHtColumn).Value = Rec.pPara.RetractHeight
            End If
            If ValidData.GetFlag(gClearanceHtColumn) Then
                .Cells(CurrentRow, gClearanceHtColumn).Value = Rec.pPara.ClearanceHeight
            End If
            If ValidData.GetFlag(gDTaiDistColumn) Then
                .Cells(CurrentRow, gDTaiDistColumn).Value = Rec.pPara.DetailingDistance
            End If
            If ValidData.GetFlag(gArcRadiusColumn) Then
                .Cells(CurrentRow, gArcRadiusColumn).Value = Rec.pPara.ArcRadius
            End If
            If ValidData.GetFlag(gPitchColumn) Then
                .Cells(CurrentRow, gPitchColumn).Value = Rec.pPara.Pitch
            End If
            If ValidData.GetFlag(gFillHeightColumn) Then
                .Cells(CurrentRow, gFillHeightColumn).Value = Rec.pPara.FilledHeight
            End If
            If ValidData.GetFlag(gNoOfRunColumn) Then
                .Cells(CurrentRow, gNoOfRunColumn).Value = Rec.pPara.NumberOfRun
            End If
            If ValidData.GetFlag(gSprialColumn) Then
                .Cells(CurrentRow, gSprialColumn).Value = Rec.pPara.SpiralFlag
            End If
            If ValidData.GetFlag(gRtAngleColumn) Then
                .Cells(CurrentRow, gRtAngleColumn).Value = Rec.pPara.RotatingAngle
            End If
            If ValidData.GetFlag(gEdgeClearColumn) Then
                .Cells(CurrentRow, gEdgeClearColumn).Value = Rec.pPara.EdgeClear
            End If
            If ValidData.GetFlag(gIONumberColumn) Then
                .Cells(CurrentRow, gIONumberColumn).Value = Rec.pPara.IONumber
            End If
        End With
        Rtn = True

        Return Rtn
        TraceGCCollect
    End Function

    'Convert external txt pattern file to internal data
    Public Sub ConvertToDataStruct(ByRef rec() As String)
        Try

            Record.pPara.Name = rec(gCommandNameColumn)
            Record.pPara.Needle = rec(gNeedleColumn)
            Record.pPara.DispenseFlag = rec(gDispensColumn)

            Record.pPara.Pos1.X = rec(gPos1XColumn)
            Record.pPara.Pos1.Y = rec(gPos1YColumn)
            Record.pPara.Pos1.Z = rec(gPos1ZColumn)
            Record.pPara.Pos2.X = rec(gPos2XColumn)
            Record.pPara.Pos2.Y = rec(gPos2YColumn)
            Record.pPara.Pos2.Z = rec(gPos2ZColumn)
            Record.pPara.Pos3.X = rec(gPos3XColumn)
            Record.pPara.Pos3.Y = rec(gPos3YColumn)
            Record.pPara.Pos3.Z = rec(gPos3ZColumn)

            Record.pPara.NeedleGap = rec(gNeedleGapColumn)
            Record.pPara.DispenseDuration = rec(gDurationColumn)
            Record.pPara.ApproachDispHeight = rec(gApproachHtColumn)
            Record.pPara.TravelDelay = rec(gTravelDelayColumn)
            Record.pPara.TravelSpeed = rec(gTravelSpeedColumn)
            Record.pPara.DetailingDistance = rec(gDTaiDistColumn)
            Record.pPara.RetractDelay = rec(gRetractDelayColumn)
            Record.pPara.RetractSpeed = rec(gRetractSpeedColumn)

            Record.pPara.RetractHeight = rec(gRetractHtColumn)
            Record.pPara.ClearanceHeight = rec(gClearanceHtColumn)
            Record.pPara.ArcRadius = rec(gArcRadiusColumn)
            Record.pPara.Pitch = rec(gPitchColumn)
            Record.pPara.FilledHeight = rec(gFillHeightColumn)
            Record.pPara.RotatingAngle = rec(gRtAngleColumn)

            Record.pPara.ClearanceHeight = rec(gClearanceHtColumn)
            Record.pPara.SpiralFlag = rec(gSprialColumn)

            If "Reference" = Record.pPara.Name Then
                ReferencePtX = Record.pPara.Pos1.X
                ReferencePtY = Record.pPara.Pos1.Y
                ReferencePtZ = Record.pPara.Pos1.Z
            End If

        Catch ex As Exception

        End Try
        TraceGCCollect
    End Sub


    Public Sub getXYMinMax(ByRef Xmin As Double, ByRef Xmax As Double, ByRef Ymin As Double, ByRef Ymax As Double)

        If "Dot" = Record.pPara.Name Or "Wait" = Record.pPara.Name Or "Move" = Record.pPara.Name Or "Line" = Record.pPara.Name Or _
            "Arc" = Record.pPara.Name Or "Rectangle" = Record.pPara.Name Or _
            "SubPattern" = Record.pPara.Name Or "FillRectangle" = Record.pPara.Name Or _
            "ChipEdge" = Record.pPara.Name Or "QC" = Record.pPara.Name Or "Offset" = Record.pPara.Name Or "Measure" = Record.pPara.Name Then
            Xmin = Record.pPara.Pos1.X
            Ymin = Record.pPara.Pos1.Y
            Xmax = Record.pPara.Pos1.X
            Ymax = Record.pPara.Pos1.Y
        End If

        If "Line" = Record.pPara.Name Or "Arc" = Record.pPara.Name Or "Rectangle" = Record.pPara.Name Or _
            "FillRectangle" = Record.pPara.Name Then
            If Xmin > Record.pPara.Pos2.X Then
                Xmin = Record.pPara.Pos2.X
            End If
            If Ymin > Record.pPara.Pos2.Y Then
                Ymin = Record.pPara.Pos2.Y
            End If
            If Xmax < Record.pPara.Pos2.X Then
                Xmax = Record.pPara.Pos2.X
            End If
            If Ymax < Record.pPara.Pos2.Y Then
                Ymax = Record.pPara.Pos2.Y
            End If
        End If

        If "Arc" = Record.pPara.Name Or "Rectangle" = Record.pPara.Name Or "FillRectangle" = Record.pPara.Name Then
            If Xmin > Record.pPara.Pos3.X Then
                Xmin = Record.pPara.Pos3.X
            End If
            If Ymin > Record.pPara.Pos3.Y Then
                Ymin = Record.pPara.Pos3.Y
            End If
            If Xmax < Record.pPara.Pos3.X Then
                Xmax = Record.pPara.Pos3.X
            End If
            If Ymax < Record.pPara.Pos3.Y Then
                Ymax = Record.pPara.Pos3.Y
            End If
        End If

        Dim pt1(3), pt2(3), pt3(3), cen(2), radius As Double
        If "Circle" = Record.pPara.Name Or "FillCircle" = Record.pPara.Name Then
            pt1(0) = Record.pPara.Pos1.X : pt1(1) = Record.pPara.Pos1.Y : pt1(2) = Record.pPara.Pos1.Z
            pt2(0) = Record.pPara.Pos2.X : pt2(1) = Record.pPara.Pos2.Y : pt2(2) = Record.pPara.Pos2.Z
            pt3(0) = Record.pPara.Pos3.X : pt3(1) = Record.pPara.Pos3.Y : pt3(2) = Record.pPara.Pos3.Z
            MathFunction.Get3dCircleCentre(pt1, pt2, pt3, cen, radius)

            Xmin = cen(0) - radius
            Xmax = cen(0) + radius
            Ymin = cen(1) - radius
            Ymax = cen(1) + radius
        End If

        If "SubPattern" = Record.pPara.Name Then



        End If
        TraceGCCollect
    End Sub


    'Error checking for one row        
    '                   
    '   return: 0=no error found

    Public Function CheckAllErrorForOneRow() As Integer
        Dim errorCode As Integer = 0
        Dim Xmin, Xmax, Ymin, Ymax As Double

        'Check the x and y range
        getXYMinMax(Xmin, Xmax, Ymin, Ymax)

        If Xmin + ReferencePtX >= SystemOrigin.X And Xmax + ReferencePtX <= WorkArea.X And _
            Ymin + ReferencePtY >= SystemOrigin.Y And Ymax + ReferencePtY <= WorkArea.Y Then
        Else
            errorCode += 1
            Return (errorCode)
        End If


        'Check logic order of z direction constraints
        If Record.pPara.ClearanceHeight >= Record.pPara.RetractHeight And Record.pPara.RetractHeight >= Record.pPara.NeedleGap And _
            WorkArea.ZMax - WorkArea.ZMin > Record.pPara.ClearanceHeight Then

        ElseIf Record.pPara.ClearanceHeight >= Record.pPara.NeedleGap And Record.pPara.NeedleGap > Record.pPara.RetractHeight And _
            Record.pPara.RetractHeight < 0.001 And WorkArea.ZMax - WorkArea.ZMin > Record.pPara.ClearanceHeight Then

        Else
            errorCode += 2
            Return (errorCode)
        End If

        'Check the z range
        If Record.pPara.Pos1.Z + ReferencePtZ >= WorkArea.ZMin And Record.pPara.Pos1.Z + +ReferencePtZ <= WorkArea.ZMax Then

        Else
            errorCode += 4
            Return (errorCode)
        End If
        TraceGCCollect
    End Function


    'Speed error checking of the data structure        
    '   AxSpreadsheet:    spreadsheet activeX component              
    '   SubSheetStruct:   data structure
    '                   
    '   return: 0=no error found

    Public Function CheckSpeedError(ByRef AxSpreadsheet As AxOWC10.AxSpreadsheet, _
            ByRef SubSheetStruct As SubSheetErrorStruct) As Integer
        Dim Rtn As Integer = 0
        Dim i, j, emptyLine As Integer
        Dim NumberOfSheet As Integer = AxSpreadsheet.Worksheets.Count()
        Dim SpreadSheetName As String
        Dim strTmp As String
        Dim ErrorMsg As String


        For i = 1 To NumberOfSheet
            SpreadSheetName = AxSpreadsheet.Worksheets(i).name()
            j = 0
            emptyLine = 0
            With AxSpreadsheet.Worksheets(SpreadSheetName)
                Do
                    j = j + 1
                    strTmp = .Cells(j, gCommandNameColumn).Value

                    If "" = strTmp Then
                        emptyLine = emptyLine + 1
                    ElseIf CheckRequiredTSpeed(strTmp) Then
                        emptyLine = 0
                        If MaxSpeedLimit < .Cells(j, gTravelSpeedColumn).Value Or _
                        MinSpeedLimit >= .Cells(j, gTravelSpeedColumn).Value Then
                            ErrorMsg = "Element speed error on Sheet: " + SpreadSheetName + ", "
                            ErrorMsg = ErrorMsg + "Line: " + CStr(j)
                            MyMsgBox(ErrorMsg, MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly, "Element speed error")
                            Rtn = 1
                            Exit For
                        End If
                        'Retract speed will be checked if the Cmd has valid Travel speed
                        If MaxSpeedLimit < .Cells(j, gRetractSpeedColumn).Value Or _
                            MinSpeedLimit >= .Cells(j, gRetractSpeedColumn).Value Then
                            ErrorMsg = "Retract speed error on Sheet: " + SpreadSheetName + ", "
                            ErrorMsg = ErrorMsg + "Line: " + CStr(j)
                            MyMsgBox(ErrorMsg, MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly, "Retract speed error")
                            Rtn = 1
                            Exit For
                        End If
                    ElseIf CheckRequiredRSpeed(strTmp) Then
                        'Retract speed will be checked for the Cmd without valid Travel speed
                        If MaxSpeedLimit < .Cells(j, gRetractSpeedColumn).Value Then
                            ErrorMsg = "Retract speed error on Sheet: " + SpreadSheetName + ", "
                            ErrorMsg = ErrorMsg + "Line: " + CStr(j)
                            MyMsgBox(ErrorMsg, MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly, "Retract speed error")
                            Rtn = 1
                            Exit For
                        End If
                    End If
                Loop Until 20 < emptyLine 'Successive 20 empty excel lines will terminate the saving record.
            End With
        Next i
        Return Rtn
        TraceGCCollect
    End Function


    'Height error checking of the data structure        
    '   AxSpreadsheet:    spreadsheet activeX component              
    '   SubSheetStruct:   data structure
    '                   
    '   return: 0=no error found

    Public Function CheckHeightError(ByRef AxSpreadsheet As AxOWC10.AxSpreadsheet, _
            ByRef SubSheetStruct As SubSheetErrorStruct) As Integer
        Dim Rtn As Integer = 0
        Dim i, j, emptyLine As Integer
        Dim NumberOfSheet As Integer = AxSpreadsheet.Worksheets.Count()
        Dim SpreadSheetName, cmdStr, ErrorMsg As String
        Dim ExtractedPt, fClearanceHeight, fRetractHeight, fNeedleGap As Double

        'Rtn = CheckPtXYError(AbsolutePt)

        For i = 1 To NumberOfSheet
            SpreadSheetName = AxSpreadsheet.ActiveWorkbook.Worksheets(i).name()
            j = 0
            emptyLine = 0
            With AxSpreadsheet.Worksheets(SpreadSheetName)
                Do
                    j = j + 1
                    cmdStr = .Cells(j, gCommandNameColumn).Value

                    If "" = cmdStr Then
                        emptyLine = emptyLine + 1
                    Else
                        emptyLine = 0

                        'Check the z range.  For 3 or 2 pts, the z are need to be checked.
                        If CheckRequiredPt3XY(cmdStr) Then
                            ExtractedPt = .Cells(j, gPos1ZColumn).Value()
                            If ExtractedPt < WorkArea.ZMin Or ExtractedPt > WorkArea.ZMax Then
                                Rtn = 1
                            End If

                            ExtractedPt = .Cells(j, gPos2ZColumn).Value()
                            If ExtractedPt < WorkArea.ZMin Or ExtractedPt > WorkArea.ZMax Then
                                Rtn = 1
                            End If
                            ExtractedPt = .Cells(j, gPos3ZColumn).Value()
                            If ExtractedPt < WorkArea.ZMin Or ExtractedPt > WorkArea.ZMax Then
                                Rtn = 1
                            End If

                        ElseIf CheckRequiredPt2XY(cmdStr) Then
                            ExtractedPt = .Cells(j, gPos1ZColumn).Value()
                            If ExtractedPt < WorkArea.ZMin Or ExtractedPt > WorkArea.ZMax Then
                                Rtn = 1
                            End If

                            ExtractedPt = .Cells(j, gPos2ZColumn).Value()
                            If ExtractedPt < WorkArea.ZMin Or ExtractedPt > WorkArea.ZMax Then
                                Rtn = 1
                            End If

                        ElseIf CheckRequiredPtHeight(cmdStr) Then   'For 1 pt, the height may not exists.

                            ExtractedPt = .Cells(j, gPos1ZColumn).Value()
                            If ExtractedPt < WorkArea.ZMin Or ExtractedPt > WorkArea.ZMax Then
                                Rtn = 1
                            End If

                        End If

                        If Rtn = 1 Then
                            ErrorMsg = "Height error on Sheet: " + SpreadSheetName + ", "
                            ErrorMsg = ErrorMsg + "Line: " + CStr(j)
                            MyMsgBox(ErrorMsg, MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly, "Height coordinate error")
                            Rtn = 1
                            Exit For
                        End If

                        'Check logic order of z direction constraints
                        If CheckValidRCHeight(cmdStr) Then

                            fClearanceHeight = .Cells(j, gClearanceHtColumn).Value
                            fRetractHeight = .Cells(j, gRetractHtColumn).Value
                            fNeedleGap = .Cells(j, gNeedleGapColumn).Value

                            If fClearanceHeight > WorkArea.ZMax Or fRetractHeight > WorkArea.ZMax Or fNeedleGap > WorkArea.ZMax Then
                                Rtn = 1
                            ElseIf fClearanceHeight < 0 Or fRetractHeight < 0 Or fNeedleGap < 0 Then
                                Rtn = 1
                            ElseIf fClearanceHeight < fRetractHeight Or fClearanceHeight < fNeedleGap Then
                                Rtn = 1
                            ElseIf WorkArea.ZMax - WorkArea.ZMin < fClearanceHeight Then
                                Rtn = 1
                            End If

                            If Rtn = 1 Then
                                ErrorMsg = "Height logical error on Sheet: " + SpreadSheetName + ", "
                                ErrorMsg = ErrorMsg + "Line: " + CStr(j)
                                MyMsgBox(ErrorMsg, MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly, "Height logical error")
                                Rtn = 1
                                Exit For
                            End If

                        End If
                    End If

                Loop Until 20 < emptyLine 'Successive 20 empty excel lines will terminate the saving record.
            End With
        Next i
        Return Rtn
        TraceGCCollect
    End Function



    'Update external reference for error checking
    '   iIndex:  index number
    '   iLevel:  level for external data structure depth
    '   RefPt:   Reference of this sub
    '   RotRad:  Rotation angle of this sub
    '
    '   SheetsStruct: data structure suitable for error checking            

    Public Sub UpdateExtReference(ByVal iIndex As Integer, ByVal iLevel As Integer, _
            ByVal RefPt As xyzCoordinate, ByVal RotRad As Double, ByRef SheetsStruct As SubSheetErrorStruct)

        'SheetsStruct.TotalNumberOfSheets is ini to zero before it is used

        SheetsStruct.ExtAbstract(SheetsStruct.ExtNumberOfSheets).SheetIndex = iIndex
        SheetsStruct.ExtAbstract(SheetsStruct.ExtNumberOfSheets).SheetLevel = iLevel

        SheetsStruct.ExtAbstract(SheetsStruct.ExtNumberOfSheets).ExternalRefPt.X = RefPt.X
        SheetsStruct.ExtAbstract(SheetsStruct.ExtNumberOfSheets).ExternalRefPt.Y = RefPt.Y
        SheetsStruct.ExtAbstract(SheetsStruct.ExtNumberOfSheets).ExternalRefPt.Z = RefPt.Z
        SheetsStruct.ExtAbstract(SheetsStruct.ExtNumberOfSheets).RotationAngleRad = RotRad

        SheetsStruct.ExtNumberOfSheets += 1
        TraceGCCollect
    End Sub



    'One point 2D coordinate transformation for ratation and translation
    '   PtInput:  input point coordinate
    '   PtReference: reference point for the translation
    '   RotAngleRad: rotation angle
    '
    '   PtOutput: output point coordinate          

    Public Function Point2DRotateTranslate(ByRef PtOutput As xyzCoordinate, ByVal PtInput As xyzCoordinate, _
        ByVal PtReference As xyzCoordinate, ByVal RotAngleRad As Double)
        Dim fCos As Double = Math.Cos(RotAngleRad)
        Dim fSin As Double = Math.Sin(RotAngleRad)
        Dim Pt As xyzCoordinate

        Pt.X = PtInput.X - PtReference.X
        Pt.Y = PtInput.Y - PtReference.Y

        PtOutput.X = Pt.X * fCos - Pt.Y * fSin
        PtOutput.Y = Pt.X * fSin + Pt.Y * fCos
        TraceGCCollect
    End Function



    'Build external reference relations for error checking
    '   AxSpreadsheet:  activeX spreadsheet handler
    '
    '   SubSheetStruct: data structure suitable for error checking            
    '   return: 0=no error found

    Public Function BuildExtReference(ByRef AxSpreadsheet As AxOWC10.AxSpreadsheet, _
            ByRef SubSheetStruct As SubSheetErrorStruct) As Integer
        Dim Rtn As Integer = 0
        Dim i, j, k, emptyLine As Integer
        Dim NumberOfSheet As Integer = AxSpreadsheet.Worksheets.Count()
        Dim SpreadSheetName As String
        Dim strTmp As String
        Dim SheetName As String
        Dim SheetName2 As String
        Dim ReferencePt As xyzCoordinate
        Dim TotalOffsetPt As xyzCoordinate
        Dim IntRotateAngleRad As Double
        Dim ExtRotateAngleRad As Double


        Dim file As New CIDSFileHandler


        SubSheetStruct.ExtNumberOfSheets = 0  'Include overlapping of use

        ReferencePt.X = 0
        ReferencePt.Y = 0
        ReferencePt.Z = 0
        TotalOffsetPt.X = 0
        TotalOffsetPt.Y = 0
        TotalOffsetPt.Z = 0
        IntRotateAngleRad = 0
        ExtRotateAngleRad = 0

        Dim iLevel As Integer = 0
        'Add "main" to the data structure.  For "Main" the iIndex=0
        UpdateExtReference(0, iLevel, TotalOffsetPt, IntRotateAngleRad, SubSheetStruct)


        'Scan each spreadsheet (Main first) to allocate Reference point.  The Ref Pt will be reflected 
        'in the corresponds sub only.

        j = 0
        iLevel = 1
        With AxSpreadsheet.Worksheets("Main")
            Do
                j = j + 1
                strTmp = .Cells(j, gCommandNameColumn).Value

                If "" = strTmp Then
                    emptyLine = emptyLine + 1
                ElseIf "Reference" = strTmp Then
                    emptyLine = 0
                    ReferencePt.X = .Cells(j, gPos1XColumn).Value
                    ReferencePt.Y = .Cells(j, gPos1YColumn).Value
                    ReferencePt.Z = .Cells(j, gPos1ZColumn).Value
                ElseIf "Array" = strTmp Or "SubPattern" = strTmp Then
                    emptyLine = 0

                    If "Array" = strTmp Then
                        SheetName = .Cells(j, gSubnameColumn).Value
                    ElseIf "SubPattern" = strTmp Then
                        SheetName = file.NameOnlyFromFullPath(.Cells(j, gSubnameColumn).Value)
                    End If

                    ExtRotateAngleRad = 0

                    'Get sub-sheet index number in the data structure
                    For k = 0 To NumberOfSheet - 1
                        If SheetName = SubSheetStruct.IntAbstract(k).SheetName Then
                            Exit For
                        End If
                    Next k

                    If "SubPattern" = strTmp Then
                        TotalOffsetPt.X = .Cells(j, gPos1XColumn).Value + ReferencePt.X
                        TotalOffsetPt.Y = .Cells(j, gPos1YColumn).Value + ReferencePt.Y
                        TotalOffsetPt.Z = .Cells(j, gPos1ZColumn).Value + ReferencePt.Z
                        IntRotateAngleRad = (.Cells(j, gRtAngleColumn).Value) * 0.017453277
                    ElseIf "Array" = strTmp Then
                        TotalOffsetPt.X = ReferencePt.X
                        TotalOffsetPt.Y = ReferencePt.Y
                        TotalOffsetPt.Z = ReferencePt.Z
                    End If
                    UpdateExtReference(k, iLevel, TotalOffsetPt, IntRotateAngleRad, SubSheetStruct)
                Else
                    emptyLine = 0
                End If
            Loop Until 20 < emptyLine 'Successive 20 empty excel lines will terminate the saving record.
        End With

        Dim SubFirstPt As xyzCoordinate
        Dim OffsetRefPt As xyzCoordinate
        Dim TemRefPt As xyzCoordinate
        Dim ExtRefPt As xyzCoordinate
        Dim SheetIndex As Integer
        Dim bRefExistedFlag As Boolean

        Dim iLoopin As Integer
        iLevel = 1
        i = 1   '"Main" will be skipped

        Do      'Loop for entire structure
            iLoopin = 0
            iLevel = iLevel + 1

            Do  'Loop for each level.  It will exit if current level completed
                'The next level will be built (if it is available) by current level

                '"Main" in excluded
                If iLevel - 1 = SubSheetStruct.ExtAbstract(i).SheetLevel Then
                    'Ini of local ref
                    ReferencePt.X = 0
                    ReferencePt.Y = 0
                    ReferencePt.Z = 0

                    'Get the external ref
                    ExtRefPt.X = SubSheetStruct.ExtAbstract(i).ExternalRefPt.X
                    ExtRefPt.Y = SubSheetStruct.ExtAbstract(i).ExternalRefPt.Y
                    ExtRefPt.Z = SubSheetStruct.ExtAbstract(i).ExternalRefPt.Z
                    ExtRotateAngleRad = SubSheetStruct.ExtAbstract(i).RotationAngleRad

                    'Get the real sheet name
                    SheetIndex = SubSheetStruct.ExtAbstract(i).SheetIndex

                    SheetName = SubSheetStruct.IntAbstract(SheetIndex).SheetName

                    If SheetName = "" Then
                        MyMsgBox("File name not found")
                        Return 1
                    End If

                    'Fast extract subSheet first valid pt
                    SubFirstPt.X = SubSheetStruct.IntAbstract(SheetIndex).SubFirstPt.X
                    SubFirstPt.Y = SubSheetStruct.IntAbstract(SheetIndex).SubFirstPt.Y
                    SubFirstPt.Z = SubSheetStruct.IntAbstract(SheetIndex).SubFirstPt.Z

                    j = 0

                    'Scan inside the sub to find any "Array" or "SubSub"
                    With AxSpreadsheet.Worksheets(SheetName)
                        Do
                            j = j + 1
                            strTmp = .Cells(j, gCommandNameColumn).Value

                            If "" = strTmp Then
                                emptyLine = emptyLine + 1
                            ElseIf "Reference" = strTmp Then
                                emptyLine = 0

                                ReferencePt.X = .Cells(j, gPos1XColumn).Value
                                ReferencePt.Y = .Cells(j, gPos1YColumn).Value
                                ReferencePt.Z = .Cells(j, gPos1ZColumn).Value

                                bRefExistedFlag = True

                            ElseIf "Array" = strTmp Or "SubPattern" = strTmp Then
                                emptyLine = 0

                                If "Array" = strTmp Then
                                    SheetName2 = SheetName + "." + .Cells(j, gSubnameColumn).Value

                                    'if dotFound=0 then 
                                    '   no change for sheetname
                                    'elseif dotFound=3 then 
                                    '   
                                    'else
                                    '   error msg
                                    'end if

                                ElseIf "SubPattern" = strTmp Then
                                    SheetName2 = file.NameOnlyFromFullPath(.Cells(j, gSubnameColumn).Value)
                                End If


                                'Get sub-sheet index number in the data structure
                                For k = 0 To NumberOfSheet - 1
                                    If SheetName2 = SubSheetStruct.IntAbstract(k).SheetName Then
                                        Exit For
                                    End If
                                Next k

                                If "SubPattern" = strTmp Then
                                    'The reference pt is rotated accordinadingly
                                    OffsetRefPt.X = ReferencePt.X + .Cells(j, gPos1XColumn).Value
                                    OffsetRefPt.Y = ReferencePt.Y + .Cells(j, gPos1YColumn).Value
                                    OffsetRefPt.Z = ReferencePt.Z + .Cells(j, gPos1ZColumn).Value

                                    'Rotate offseted Ref point
                                    Point2DRotateTranslate(TemRefPt, OffsetRefPt, SubFirstPt, ExtRotateAngleRad)

                                    'Combine the offseted Ref pt with External Ref Pt
                                    TotalOffsetPt.X = ExtRefPt.X + TemRefPt.X
                                    TotalOffsetPt.Y = ExtRefPt.Y + TemRefPt.Y
                                    TotalOffsetPt.Z = ExtRefPt.Z + TemRefPt.Z

                                    'Accumulate rotation angle
                                    IntRotateAngleRad = ExtRotateAngleRad + (.Cells(j, gRtAngleColumn).Value) * 0.017453277

                                    'Normalize rotation angle range[0, 360]
                                    If IntRotateAngleRad > 6.28318 Then    '360*3.14159/180=6.28318
                                        IntRotateAngleRad = IntRotateAngleRad - 6.28318
                                    ElseIf IntRotateAngleRad < 0 Then
                                        IntRotateAngleRad = IntRotateAngleRad + 6.28318
                                    End If
                                ElseIf "Array" = strTmp Then
                                    'The reference pt is rotated accordinadingly
                                    Point2DRotateTranslate(TemRefPt, ReferencePt, SubFirstPt, ExtRotateAngleRad)

                                    TotalOffsetPt.X = ExtRefPt.X + TemRefPt.X
                                    TotalOffsetPt.Y = ExtRefPt.Y + TemRefPt.Y
                                    TotalOffsetPt.Z = ExtRefPt.Z + TemRefPt.Z
                                    IntRotateAngleRad = ExtRotateAngleRad
                                End If

                                UpdateExtReference(k, iLevel, TotalOffsetPt, IntRotateAngleRad, SubSheetStruct)

                            Else
                                emptyLine = 0
                            End If
                        Loop Until 20 < emptyLine 'Successive 20 empty excel lines will terminate the saving record.
                    End With
                    iLoopin = 1
                Else
                    Exit Do
                End If
                i = i + 1
            Loop
        Loop Until 0 = iLoopin
        TraceGCCollect
    End Function



    'Check the Pt has valid x,y or not
    '   cmdStr:  command string
    '          
    '   return: TRUE=the Pt has valid x,y 

    Public Function CheckValidPtXY(ByRef cmdStr As String) As Boolean
        'Check the Pt has valid x,y or not
        Dim ValidPtFound As Boolean = False
        If "Fiducial" = cmdStr Or "Reference" = cmdStr Or "Height" = cmdStr Or "Reject" = cmdStr Or _
                "Dot" = cmdStr Or "Line" = cmdStr Or "Arc" = cmdStr Or "Circle" = cmdStr Or "FillCircle" = cmdStr Or _
                "Rectangle" = cmdStr Or "FillRectangle" = cmdStr Or "SubPattern" = cmdStr Or _
                "Wait" = cmdStr Or "Move" = cmdStr Or "ChipEdge" = cmdStr Or "QC" = cmdStr Or _
                "Subpattern" = cmdStr Or "DotArray" = cmdStr Then
            ValidPtFound = True
        End If
        Return ValidPtFound
        TraceGCCollect
    End Function



    'Find the 1st Pt x,y need to be checked or not
    '   cmdStr:  command string
    '          
    '   return: TRUE=the 1st Pt x,y need to be checked

    Public Function CheckRequiredPt1XY(ByRef cmdStr As String) As Boolean
        Dim ValidPtFound As Boolean = False
        If "Fiducial" = cmdStr Or "Height" = cmdStr Or "Reject" = cmdStr Or "Reference" = cmdStr Or _
                "Dot" = cmdStr Or "Line" = cmdStr Or "Arc" = cmdStr Or "Circle" = cmdStr Or "FillCircle" = cmdStr Or _
                "Rectangle" = cmdStr Or "FillRectangle" = cmdStr Or "SubPattern" = cmdStr Or "ChipEdge" = cmdStr Or _
                "Wait" = cmdStr Or "Move" = cmdStr Or "ChipEdge" = cmdStr Or "QC" = cmdStr Then
            ValidPtFound = True
        End If
        Return ValidPtFound
        TraceGCCollect
    End Function



    'Find the 1st and 2nd Pts x,y need to be checked or not
    '   cmdStr:  command string
    '          
    '   return: TRUE=the 1st and 2nd Pts x,y need to be checked

    Public Function CheckRequiredPt2XY(ByRef cmdStr As String) As Boolean
        Dim ValidPtFound As Boolean = False
        'If "Fiducial" = cmdStr Or "Line" = cmdStr Or "Arc" = cmdStr Or "Circle" = cmdStr Or _
        '       "FillCircle" = cmdStr Or "Rectangle" = cmdStr Or "FillRectangle" = cmdStr Or "ChipEdge" = cmdStr Then
        'ValidPtFound = True
        'End If
        'lim modify chipedge as 1 xy point
        If "Fiducial" = cmdStr Or "Line" = cmdStr Or "Arc" = cmdStr Or "Circle" = cmdStr Or _
                "FillCircle" = cmdStr Or "Rectangle" = cmdStr Or "FillRectangle" = cmdStr Then
            ValidPtFound = True
        End If
        Return ValidPtFound
        TraceGCCollect
    End Function

    'Find the 1st, 2nd and 3rd Pts x,y need to be checked or not
    '   cmdStr:  command string
    '          
    '   return: TRUE=the 1st, 2nd and 3rd Pts x,y need to be checked

    Public Function CheckRequiredPt3XY(ByRef cmdStr As String) As Boolean
        Dim ValidPtFound As Boolean = False
        If "Arc" = cmdStr Or "Circle" = cmdStr Or "FillCircle" = cmdStr Or _
                "Rectangle" = cmdStr Or "FillRectangle" = cmdStr Or "DotArray" = cmdStr Then
            ValidPtFound = True
        End If
        Return ValidPtFound
        TraceGCCollect
    End Function

    'Find the Msg z needs to be checked or not
    '   cmdStr:  command string
    '          
    '   return: TRUE=z needs to be checked

    Public Function CheckRequiredPtHeight(ByRef cmdStr As String) As Boolean
        Dim ValidPtFound As Boolean = False
        If "Dot" = cmdStr Or "Line" = cmdStr Or "Arc" = cmdStr Or "Circle" = cmdStr Or "FillCircle" = cmdStr Or _
                "Rectangle" = cmdStr Or "FillRectangle" = cmdStr Or _
                "Wait" = cmdStr Or "Move" = cmdStr Or "ChipEdge" = cmdStr Then
            ValidPtFound = True
        End If
        Return ValidPtFound
        TraceGCCollect
    End Function


    'Find the Msg height information needs to be checked or not
    '   cmdStr:  command string
    '          
    '   return: TRUE=height information needs to be checked

    Public Function CheckValidRCHeight(ByRef cmdStr As String) As Boolean
        Dim ValidPtFound As Boolean = False
        If "Dot" = cmdStr Or "Line" = cmdStr Or "Arc" = cmdStr Or "Circle" = cmdStr Or _
                "FillCircle" = cmdStr Or "Rectangle" = cmdStr Or "FillRectangle" = cmdStr Or "ChipEdge" = cmdStr Then
            ValidPtFound = True
        End If
        Return ValidPtFound
        TraceGCCollect
    End Function


    'Find the Msg's Travel Speed needs to be checked or not
    '   cmdStr:  command string
    '          
    '   return: TRUE=Travel Speed needs to be checked

    Public Function CheckRequiredTSpeed(ByRef cmdStr As String) As Boolean
        Dim ValidPtFound As Boolean = False
        If "Line" = cmdStr Or "Arc" = cmdStr Or "Rectangle" = cmdStr Or "FillRectangle" = cmdStr Or _
            "Circle" = cmdStr Or "FillCircle" = cmdStr Or "ChipEdge" = cmdStr Then
            ValidPtFound = True
        End If
        Return ValidPtFound
        TraceGCCollect
    End Function


    'Find the Msg's Retract Speed needs to be checked or not
    '   cmdStr:  command string
    '          
    '   return: TRUE=Retract Speed needs to be checked

    Public Function CheckRequiredRSpeed(ByRef cmdStr As String) As Boolean
        Dim ValidPtFound As Boolean = False
        If "Dot" = cmdStr Or "Line" = cmdStr Or "Arc" = cmdStr Or "Circle" = cmdStr Or "FillCircle" = cmdStr Or _
                "Rectangle" = cmdStr Or "FillRectangle" = cmdStr Or "ChipEdge" = cmdStr Then
            ValidPtFound = True
        End If
        Return ValidPtFound
        TraceGCCollect
    End Function


    'Get the first valid x,y (mainly) in a Sub or Array
    '   Line number which has valid first pt x and y
    '
    '   AxSpreadsheet:  activeX spreadsheet handler
    '   SheetName: input sheet name
    '          
    '   return: the row number which contains the first valid point

    Public Function GetFirstValidPtXY(ByRef AxSpreadsheet As AxOWC10.AxSpreadsheet, ByRef SheetName As String) As Integer
        Dim i, j, kk As Integer
        Dim cmdStr As String

        Dim UsedRowNumber As Integer
        UsedRowNumber = Spreadsheet_CountRowUsed(AxSpreadsheet, SheetName)

        With AxSpreadsheet.Worksheets(SheetName)
            kk = 0
            j = 0
            Do
                kk = kk + 1
                j = j + 1
                cmdStr = .Cells(j, gCommandNameColumn).Value
                If "" <> cmdStr Then
                    If CheckValidPtXY(cmdStr) Then          'Find the row with valid point
                        Exit Do
                    End If
                End If
            Loop Until kk > UsedRowNumber

            If kk > UsedRowNumber Then
                Return -1
            Else
                Return j
            End If
        End With
        TraceGCCollect
    End Function


    'Build internal reference relations for error checking
    '   AxSpreadsheet:  activeX spreadsheet handler
    '
    '   SubSheetStruct: data structure suitable for error checking            
    '   return: 0=no error found

    Public Function BuildIntReference(ByRef AxSpreadsheet As AxOWC10.AxSpreadsheet, _
            ByRef SubSheetStruct As SubSheetErrorStruct) As Integer

        Dim Rtn As Integer = 0
        Dim i, j, k As Integer
        Dim NumberOfSheet As Integer = AxSpreadsheet.Worksheets.Count()
        Dim SpreadSheetName As String
        Dim pt1(3), pt2(3), pt3(3), cen1(2), radius As Double
        Dim sheetIsAnArray As Boolean
        Dim DotDim As Integer

        SubSheetStruct.IntNumberOfSheets = NumberOfSheet

        'Build ini error checking data structure.  Initialize data structure for error checking.
        'Same sheetName with same reference point will only appear once
        For i = 0 To NumberOfSheet - 1
            SpreadSheetName = AxSpreadsheet.ActiveWorkbook.Worksheets(NumberOfSheet - i).name()
            SubSheetStruct.IntAbstract(i).SheetName = SpreadSheetName

            If "Main" = SpreadSheetName Then
                SubSheetStruct.IntAbstract(i).SubFirstPt.X = 0
                SubSheetStruct.IntAbstract(i).SubFirstPt.Y = 0
            Else

                'Get line No of the first valid point (x,y)
                j = GetFirstValidPtXY(AxSpreadsheet, SpreadSheetName)
                If -1 = j Then
                    MyMsgBox("First Pt with valid x, y cannot be found!", MsgBoxStyle.OKOnly, "Error found")
                    Rtn = 1
                Else

                    'Determine the sheet is a Array or a sub
                    DotDim = 1
                    For k = 0 To Len(SpreadSheetName) - 1
                        If "." = SpreadSheetName.Chars(k) Then
                            DotDim += 1
                        End If
                    Next

                    sheetIsAnArray = False
                    If 1 = DotDim Or 3 = DotDim Then
                        sheetIsAnArray = True
                    End If


                    'Check whether the first Cmd with a valid Pt is a circle (include FillCircle) or not.  
                    'If it is a circle in a Array, the Center will be used as the first Pt
                    With AxSpreadsheet.Worksheets(SpreadSheetName)
                        If sheetIsAnArray And _
                            ("Circle" = .Cells(j, gCommandNameColumn).Value Or "FillCircle" = .Cells(j, gCommandNameColumn).Value) Then
                            'Find center of a circle.  It will be used as "Array" first pt

                            pt1(0) = .Cells(j, gPos1XColumn).Value : pt1(1) = .Cells(j, gPos1YColumn).Value : pt1(2) = .Cells(j, gPos1ZColumn).Value
                            pt2(0) = .Cells(j, gPos2XColumn).Value : pt2(1) = .Cells(j, gPos2YColumn).Value : pt2(2) = .Cells(j, gPos2ZColumn).Value
                            pt3(0) = .Cells(j, gPos3XColumn).Value : pt3(1) = .Cells(j, gPos3YColumn).Value : pt3(2) = .Cells(j, gPos3ZColumn).Value
                            MathFunction.Get3dCircleCentre(pt1, pt2, pt3, cen1, radius)

                            SubSheetStruct.IntAbstract(i).SubFirstPt.X = cen1(0)
                            SubSheetStruct.IntAbstract(i).SubFirstPt.Y = cen1(1)
                        Else
                            SubSheetStruct.IntAbstract(i).SubFirstPt.X = _
                                AxSpreadsheet.Worksheets(SpreadSheetName).Cells(j, gPos1XColumn).Value
                            SubSheetStruct.IntAbstract(i).SubFirstPt.Y = _
                                AxSpreadsheet.Worksheets(SpreadSheetName).Cells(j, gPos1YColumn).Value

                        End If
                    End With
                End If
            End If

        Next i

        Return (Rtn)
        TraceGCCollect
    End Function



    'Check whether the XY in the range
    '   AbsolutePt:  absolute coordinate system               
    '                   
    '   return: 0=no error found

    Public Function CheckPtXYError(ByRef AbsolutePt As xyzCoordinate) As Integer
        Dim Rtn As Integer = 0
        If AbsolutePt.X < WorkArea.XMin Or AbsolutePt.X > WorkArea.XMax Or _
            AbsolutePt.Y < WorkArea.YMin Or AbsolutePt.Y > WorkArea.YMax Then
            Rtn = 1
        End If
        Return (Rtn)
        TraceGCCollect
    End Function


    'Check XY error in points
    '   AxSpreadsheet:  activeX component handler               
    '   SubSheetStruct:  data structure for the subsheets
    '                   
    '   return: 0=no error found

    Public Function CheckPointsXYError(ByRef AxSpreadsheet As AxOWC10.AxSpreadsheet, _
                ByRef SubSheetStruct As SubSheetErrorStruct) As Integer
        Dim Rtn As Integer = 0
        Dim i, j, k, emptyLine As Integer
        Dim NumberOfSheet As Integer = AxSpreadsheet.Worksheets.Count()
        Dim SheetName As String
        Dim cmdStr As String
        Dim ErrorMsg As String
        Dim ReferencePt As xyzCoordinate
        Dim ExtRefPt As xyzCoordinate
        Dim SubFirstPt As xyzCoordinate
        Dim ExtractedPt As xyzCoordinate
        Dim AbsolutePt As xyzCoordinate
        Dim SheetIndex As Integer
        Dim ExtRotateAngleRad As Double

        Dim pt1(3), pt2(3), pt3(3), cen1(2), radius As Double

        ReferencePt.X = 0
        ReferencePt.Y = 0
        ReferencePt.Z = 0


        For i = 0 To SubSheetStruct.ExtNumberOfSheets

            ReferencePt.X = 0
            ReferencePt.Y = 0
            ReferencePt.Z = 0

            'Get the external ref
            ExtRefPt.X = SubSheetStruct.ExtAbstract(i).ExternalRefPt.X
            ExtRefPt.Y = SubSheetStruct.ExtAbstract(i).ExternalRefPt.Y
            ExtRefPt.Z = SubSheetStruct.ExtAbstract(i).ExternalRefPt.Z
            ExtRotateAngleRad = SubSheetStruct.ExtAbstract(i).RotationAngleRad

            'Get the real sheet name
            SheetIndex = SubSheetStruct.ExtAbstract(i).SheetIndex

            SheetName = SubSheetStruct.IntAbstract(SheetIndex).SheetName

            'Fast extract subSheet first valid pt
            SubFirstPt.X = SubSheetStruct.IntAbstract(SheetIndex).SubFirstPt.X
            SubFirstPt.Y = SubSheetStruct.IntAbstract(SheetIndex).SubFirstPt.Y
            SubFirstPt.Z = SubSheetStruct.IntAbstract(SheetIndex).SubFirstPt.Z



            j = 0
            emptyLine = 0
            With AxSpreadsheet.Worksheets(SheetName)
                Do
                    j = j + 1
                    cmdStr = .Cells(j, gCommandNameColumn).Value

                    If "" = cmdStr Then
                        emptyLine = emptyLine + 1
                    ElseIf "Reference" = cmdStr Then
                        emptyLine = 0
                        'Don't need to check Ref Pts error.  If a RefPt is outside the effective working area, the
                        'offset result data may still be inside the area.
                        ReferencePt.X = .Cells(j, gPos1XColumn).Value
                        ReferencePt.Y = .Cells(j, gPos1YColumn).Value
                        ReferencePt.Z = .Cells(j, gPos1ZColumn).Value
                    ElseIf "Array" = cmdStr Then
                        emptyLine = 0
                    Else
                        If CheckRequiredPt3XY(cmdStr) Then
                            If "Circle" = cmdStr Or "FillCircle" = cmdStr Then
                                'Convert 3Pts to a circle with Center and radius

                                pt1(0) = .Cells(j, gPos1XColumn).Value : pt1(1) = .Cells(j, gPos1YColumn).Value : pt1(2) = .Cells(j, gPos1ZColumn).Value
                                pt2(0) = .Cells(j, gPos2XColumn).Value : pt2(1) = .Cells(j, gPos2YColumn).Value : pt2(2) = .Cells(j, gPos2ZColumn).Value
                                pt3(0) = .Cells(j, gPos3XColumn).Value : pt3(1) = .Cells(j, gPos3YColumn).Value : pt3(2) = .Cells(j, gPos3ZColumn).Value
                                MathFunction.Get3dCircleCentre(pt1, pt2, pt3, cen1, radius)
                                ExtractedPt.X = cen1(0) + ReferencePt.X
                                ExtractedPt.Y = cen1(1) + ReferencePt.Y
                                Point2DRotateTranslate(AbsolutePt, ExtractedPt, SubFirstPt, ExtRotateAngleRad)

                                'Check 4 Pts for a circle
                                ExtractedPt.X = AbsolutePt.X
                                ExtractedPt.Y = AbsolutePt.Y

                                AbsolutePt.X = ExtractedPt.X + radius
                                Rtn = CheckPtXYError(AbsolutePt)
                                AbsolutePt.X = ExtractedPt.X - radius
                                Rtn = Rtn + CheckPtXYError(AbsolutePt)

                                AbsolutePt.X = ExtractedPt.X
                                AbsolutePt.Y = ExtractedPt.Y + radius
                                Rtn = Rtn + CheckPtXYError(AbsolutePt)
                                AbsolutePt.Y = ExtractedPt.Y - radius
                                Rtn = Rtn + CheckPtXYError(AbsolutePt)

                            Else
                                'Check the 1st Pt
                                ExtractedPt.X = .Cells(j, gPos1XColumn).Value() + ReferencePt.X
                                ExtractedPt.Y = .Cells(j, gPos1YColumn).Value() + ReferencePt.Y

                                Point2DRotateTranslate(AbsolutePt, ExtractedPt, SubFirstPt, ExtRotateAngleRad)
                                AbsolutePt.X = ExtRefPt.X + AbsolutePt.X
                                AbsolutePt.Y = ExtRefPt.Y + AbsolutePt.Y
                                Rtn = CheckPtXYError(AbsolutePt)

                                'Check the 2nd Pt
                                ExtractedPt.X = .Cells(j, gPos2XColumn).Value() + ReferencePt.X
                                ExtractedPt.Y = .Cells(j, gPos2YColumn).Value() + ReferencePt.Y

                                Point2DRotateTranslate(AbsolutePt, ExtractedPt, SubFirstPt, ExtRotateAngleRad)
                                AbsolutePt.X = ExtRefPt.X + AbsolutePt.X
                                AbsolutePt.Y = ExtRefPt.Y + AbsolutePt.Y
                                Rtn = Rtn + CheckPtXYError(AbsolutePt)

                                'Check the 3rd Pt
                                ExtractedPt.X = .Cells(j, gPos3XColumn).Value() + ReferencePt.X
                                ExtractedPt.Y = .Cells(j, gPos3YColumn).Value() + ReferencePt.Y

                                Point2DRotateTranslate(AbsolutePt, ExtractedPt, SubFirstPt, ExtRotateAngleRad)
                                AbsolutePt.X = ExtRefPt.X + AbsolutePt.X
                                AbsolutePt.Y = ExtRefPt.Y + AbsolutePt.Y
                                Rtn = Rtn + CheckPtXYError(AbsolutePt)
                            End If

                        ElseIf CheckRequiredPt2XY(cmdStr) Then
                            ExtractedPt.X = .Cells(j, gPos1XColumn).Value() + ReferencePt.X
                            ExtractedPt.Y = .Cells(j, gPos1YColumn).Value() + ReferencePt.Y

                            Point2DRotateTranslate(AbsolutePt, ExtractedPt, SubFirstPt, ExtRotateAngleRad)
                            AbsolutePt.X = ExtRefPt.X + AbsolutePt.X
                            AbsolutePt.Y = ExtRefPt.Y + AbsolutePt.Y
                            Rtn = CheckPtXYError(AbsolutePt)


                            ExtractedPt.X = .Cells(j, gPos2XColumn).Value() + ReferencePt.X
                            ExtractedPt.Y = .Cells(j, gPos2YColumn).Value() + ReferencePt.Y

                            Point2DRotateTranslate(AbsolutePt, ExtractedPt, SubFirstPt, ExtRotateAngleRad)
                            AbsolutePt.X = ExtRefPt.X + AbsolutePt.X
                            AbsolutePt.Y = ExtRefPt.Y + AbsolutePt.Y
                            Rtn = Rtn + CheckPtXYError(AbsolutePt)

                        ElseIf CheckRequiredPt1XY(cmdStr) Then
                            ExtractedPt.X = .Cells(j, gPos1XColumn).Value() + ReferencePt.X
                            ExtractedPt.Y = .Cells(j, gPos1YColumn).Value() + ReferencePt.Y

                            Point2DRotateTranslate(AbsolutePt, ExtractedPt, SubFirstPt, ExtRotateAngleRad)
                            AbsolutePt.X = ExtRefPt.X + AbsolutePt.X
                            AbsolutePt.Y = ExtRefPt.Y + AbsolutePt.Y
                            Rtn = CheckPtXYError(AbsolutePt)
                        End If

                        emptyLine = 0

                        'If 0 <> Rtn Then
                        '    ErrorMsg = "Element XY coordinate error on Sheet: " + SheetName + ", "
                        '    ErrorMsg = ErrorMsg + "Row: " + CStr(j)
                        '    MyMsgBox(ErrorMsg, MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly, "Element coordinate error")
                        '    Return (Rtn)
                        'End If
                    End If

                Loop Until 20 < emptyLine 'Successive 20 empty excel lines will terminate the saving record.
            End With
        Next i
        TraceGCCollect
    End Function



    'Check XY error in one sheet
    '   CurrentSheetName: sheet name
    '   AxSpreadsheet:  activeX component handler               
    '   SubSheetStruct:  data structure for the subsheets
    '                   
    '   return: 0=no error found

    Public Function CheckXYErrorInOneSheet(ByVal CurrentSheetName As String, _
            ByRef AxSpreadsheet As AxOWC10.AxSpreadsheet, _
            ByRef SubSheetStruct As SubSheetErrorStruct) As Integer
        Dim Rtn As Integer = 0
        Dim i, j, k, emptyLine As Integer
        Dim NumberOfSheet As Integer = AxSpreadsheet.Worksheets.Count()
        Dim SheetName As String
        Dim cmdStr As String
        Dim ErrorMsg As String
        Dim ReferencePt As xyzCoordinate
        Dim ExtRefPt As xyzCoordinate
        Dim SubFirstPt As xyzCoordinate
        Dim ExtractedPt As xyzCoordinate
        Dim AbsolutePt As xyzCoordinate
        Dim SheetIndex As Integer
        Dim ExtRotateAngleRad As Double

        Dim pt1(3), pt2(3), pt3(3), cen1(2), radius As Double

        ReferencePt.X = 0
        ReferencePt.Y = 0
        ReferencePt.Z = 0


        For i = 0 To SubSheetStruct.ExtNumberOfSheets

            ReferencePt.X = 0
            ReferencePt.Y = 0
            ReferencePt.Z = 0

            'Get the external ref
            ExtRefPt.X = SubSheetStruct.ExtAbstract(i).ExternalRefPt.X
            ExtRefPt.Y = SubSheetStruct.ExtAbstract(i).ExternalRefPt.Y
            ExtRefPt.Z = SubSheetStruct.ExtAbstract(i).ExternalRefPt.Z
            ExtRotateAngleRad = SubSheetStruct.ExtAbstract(i).RotationAngleRad

            'Get the real sheet name
            SheetIndex = SubSheetStruct.ExtAbstract(i).SheetIndex

            SheetName = SubSheetStruct.IntAbstract(SheetIndex).SheetName

            'Fast extract subSheet first valid pt
            SubFirstPt.X = SubSheetStruct.IntAbstract(SheetIndex).SubFirstPt.X
            SubFirstPt.Y = SubSheetStruct.IntAbstract(SheetIndex).SubFirstPt.Y
            SubFirstPt.Z = SubSheetStruct.IntAbstract(SheetIndex).SubFirstPt.Z



            j = 0
            emptyLine = 0
            With AxSpreadsheet.Worksheets(SheetName)
                Do
                    'Only check the error inside the specify sheet
                    If SheetName <> CurrentSheetName Then
                        Exit Do
                    End If

                    j = j + 1
                    cmdStr = .Cells(j, gCommandNameColumn).Value

                    If "" = cmdStr Then
                        emptyLine = emptyLine + 1
                    ElseIf "Reference" = cmdStr Then
                        emptyLine = 0
                        'Don't need to check Ref Pts error.  If a RefPt is outside the effective working area, the
                        'offset result data may still be inside the area.
                        ReferencePt.X = .Cells(j, gPos1XColumn).Value
                        ReferencePt.Y = .Cells(j, gPos1YColumn).Value
                        ReferencePt.Z = .Cells(j, gPos1ZColumn).Value
                    ElseIf "Array" = cmdStr Then
                        emptyLine = 0
                    Else
                        If CheckRequiredPt3XY(cmdStr) Then
                            If "Circle" = cmdStr Or "FillCircle" = cmdStr Then
                                'Convert 3Pts to a circle with Center and radius

                                pt1(0) = .Cells(j, gPos1XColumn).Value : pt1(1) = .Cells(j, gPos1YColumn).Value : pt1(2) = .Cells(j, gPos1ZColumn).Value
                                pt2(0) = .Cells(j, gPos2XColumn).Value : pt2(1) = .Cells(j, gPos2YColumn).Value : pt2(2) = .Cells(j, gPos2ZColumn).Value
                                pt3(0) = .Cells(j, gPos3XColumn).Value : pt3(1) = .Cells(j, gPos3YColumn).Value : pt3(2) = .Cells(j, gPos3ZColumn).Value
                                MathFunction.Get3dCircleCentre(pt1, pt2, pt3, cen1, radius)
                                ExtractedPt.X = cen1(0) + ReferencePt.X
                                ExtractedPt.Y = cen1(1) + ReferencePt.Y
                                Point2DRotateTranslate(AbsolutePt, ExtractedPt, SubFirstPt, ExtRotateAngleRad)

                                'Check 4 Pts for a circle
                                ExtractedPt.X = AbsolutePt.X
                                ExtractedPt.Y = AbsolutePt.Y

                                AbsolutePt.X = ExtractedPt.X + radius
                                Rtn = CheckPtXYError(AbsolutePt)
                                AbsolutePt.X = ExtractedPt.X - radius
                                Rtn = Rtn + CheckPtXYError(AbsolutePt)

                                AbsolutePt.X = ExtractedPt.X
                                AbsolutePt.Y = ExtractedPt.Y + radius
                                Rtn = Rtn + CheckPtXYError(AbsolutePt)
                                AbsolutePt.Y = ExtractedPt.Y - radius
                                Rtn = Rtn + CheckPtXYError(AbsolutePt)

                            Else
                                'Check the 1st Pt
                                ExtractedPt.X = .Cells(j, gPos1XColumn).Value() + ReferencePt.X
                                ExtractedPt.Y = .Cells(j, gPos1YColumn).Value() + ReferencePt.Y

                                Point2DRotateTranslate(AbsolutePt, ExtractedPt, SubFirstPt, ExtRotateAngleRad)
                                AbsolutePt.X = ExtRefPt.X + AbsolutePt.X
                                AbsolutePt.Y = ExtRefPt.Y + AbsolutePt.Y
                                Rtn = CheckPtXYError(AbsolutePt)

                                'Check the 2nd Pt
                                ExtractedPt.X = .Cells(j, gPos2XColumn).Value() + ReferencePt.X
                                ExtractedPt.Y = .Cells(j, gPos2YColumn).Value() + ReferencePt.Y

                                Point2DRotateTranslate(AbsolutePt, ExtractedPt, SubFirstPt, ExtRotateAngleRad)
                                AbsolutePt.X = ExtRefPt.X + AbsolutePt.X
                                AbsolutePt.Y = ExtRefPt.Y + AbsolutePt.Y
                                Rtn = Rtn + CheckPtXYError(AbsolutePt)

                                'Check the 3rd Pt
                                ExtractedPt.X = .Cells(j, gPos3XColumn).Value() + ReferencePt.X
                                ExtractedPt.Y = .Cells(j, gPos3YColumn).Value() + ReferencePt.Y

                                Point2DRotateTranslate(AbsolutePt, ExtractedPt, SubFirstPt, ExtRotateAngleRad)
                                AbsolutePt.X = ExtRefPt.X + AbsolutePt.X
                                AbsolutePt.Y = ExtRefPt.Y + AbsolutePt.Y
                                Rtn = Rtn + CheckPtXYError(AbsolutePt)
                            End If

                        ElseIf CheckRequiredPt2XY(cmdStr) Then
                            ExtractedPt.X = .Cells(j, gPos1XColumn).Value() + ReferencePt.X
                            ExtractedPt.Y = .Cells(j, gPos1YColumn).Value() + ReferencePt.Y

                            Point2DRotateTranslate(AbsolutePt, ExtractedPt, SubFirstPt, ExtRotateAngleRad)
                            AbsolutePt.X = ExtRefPt.X + AbsolutePt.X
                            AbsolutePt.Y = ExtRefPt.Y + AbsolutePt.Y
                            Rtn = CheckPtXYError(AbsolutePt)


                            ExtractedPt.X = .Cells(j, gPos2XColumn).Value() + ReferencePt.X
                            ExtractedPt.Y = .Cells(j, gPos2YColumn).Value() + ReferencePt.Y

                            Point2DRotateTranslate(AbsolutePt, ExtractedPt, SubFirstPt, ExtRotateAngleRad)
                            AbsolutePt.X = ExtRefPt.X + AbsolutePt.X
                            AbsolutePt.Y = ExtRefPt.Y + AbsolutePt.Y
                            Rtn = Rtn + CheckPtXYError(AbsolutePt)

                        ElseIf CheckRequiredPt1XY(cmdStr) Then
                            ExtractedPt.X = .Cells(j, gPos1XColumn).Value() + ReferencePt.X
                            ExtractedPt.Y = .Cells(j, gPos1YColumn).Value() + ReferencePt.Y

                            Point2DRotateTranslate(AbsolutePt, ExtractedPt, SubFirstPt, ExtRotateAngleRad)
                            AbsolutePt.X = ExtRefPt.X + AbsolutePt.X
                            AbsolutePt.Y = ExtRefPt.Y + AbsolutePt.Y
                            Rtn = CheckPtXYError(AbsolutePt)
                        End If

                        emptyLine = 0

                        'If 0 <> Rtn Then
                        '    ErrorMsg = "Element XY coordinate error on Sheet: " + SheetName + ", "
                        '    ErrorMsg = ErrorMsg + "Row: " + CStr(j)
                        '    MyMsgBox(ErrorMsg, MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly, "Element coordinate error")
                        '    Return (Rtn)
                        'End If
                    End If

                Loop Until 20 < emptyLine 'Successive 20 empty excel lines will terminate the saving record.
            End With
        Next i
        TraceGCCollect
    End Function



    'Count the rows used by current sheet
    '   AxSpreadsheet: the activeX component handler
    '   SheetName: the input sheet name
    '
    '   Return: Number of used rows
    '
    'Note: It may not be used.  It may be superceeded by an existing method 
    '      provided by AxSpreasheed

    Public Function Spreadsheet_CountRowUsed(ByRef AxSpreadsheet As AxOWC10.AxSpreadsheet, ByRef SheetName As String) As Integer

        Dim j, emptyLine As Integer

        j = 0
        emptyLine = 0
        With AxSpreadsheet.Worksheets(SheetName)
            Do
                j = j + 1
                If "" = .Cells(j, gCommandNameColumn).Value Then
                    emptyLine = emptyLine + 1
                Else
                    emptyLine = 0
                End If
            Loop Until 20 < emptyLine 'Successive 20 empty excel lines will terminate the saving record.
        End With
        Return j - 21
        TraceGCCollect
    End Function


    'Overall error checking            
    '   sheet:  activeX component handler               
    '   ErrorSheetData:  data structure for the error checking
    '                   
    '   return: 0=no error found

    Public Function CheckAllError(ByRef sheet As AxOWC10.AxSpreadsheet, ByRef ErrorSheetData As SubSheetErrorStruct) As Integer
        Dim Rtn As Integer = 0
        If 0 = CheckSpeedError(sheet, ErrorSheetData) And 0 = CheckHeightError(sheet, ErrorSheetData) Then
            'No error checked for Speed.  Then we check Points
            If 0 = BuildIntReference(sheet, ErrorSheetData) Then
                If BuildExtReference(sheet, ErrorSheetData) = 1 Then
                    Rtn = 1
                End If
            Else
                Rtn = 1
            End If
            If 0 <> CheckPointsXYError(sheet, ErrorSheetData) Then
                Rtn = 1
            End If
        Else
            Rtn = 1
        End If
        Return Rtn
        TraceGCCollect
    End Function


    'Speed error checking            
    '   CurrentSpeed:          speed related variables               
    '                   
    '   return: 0=no error found

    Public Function CheckCellSpeedError(ByVal CurrentSpeed) As Integer
        Dim Rtn As Integer = 0
        If CurrentSpeed > MaxSpeedLimit Or CurrentSpeed < MinSpeedLimit Then
            Rtn = 1
        End If
        Return Rtn
        TraceGCCollect
    End Function


    'Height error checking of specific data input        
    '   CurrentHeight:          height related variables               
    '   fClearanceHeight:  
    '   fRetractHeight:
    '   fNeedleGap:
    '                   
    '   return: 0=no error found

    Public Function CheckCellHeightError(ByVal CurrentHeight, _
        ByVal fClearanceHeight, ByVal fRetractHeight, ByVal fNeedleGap) As Integer
        Dim Rtn As Integer = 0

        If CurrentHeight < WorkArea.ZMin Or CurrentHeight > WorkArea.ZMax Then
            Rtn = 1
        End If

        If fClearanceHeight > WorkArea.ZMax Or fRetractHeight > WorkArea.ZMax Or fNeedleGap > WorkArea.ZMax Then
            Rtn = 1
            'ElseIf fClearanceHeight < 0 Or fRetractHeight < 0 Or fNeedleGap < 0 Then
        ElseIf fClearanceHeight < 0 Or fRetractHeight < 0 Then
            Rtn = 1
        ElseIf fClearanceHeight < fRetractHeight Or fClearanceHeight < fNeedleGap Then
            Rtn = 1
        ElseIf WorkArea.ZMax - WorkArea.ZMin < fClearanceHeight Then
            Rtn = 1
        End If

        Return Rtn
        TraceGCCollect
    End Function



    'Check errors for XY cell only
    '   CurrentSheetName: Current active sheet name
    '   InputStr: Data value of the cell
    '   row: the row number of the cell
    '   colum: the column number of the cell
    '   sheet: the ActiveX spreadsheet
    '   ErrorSheetData: error checking data structure
    '
    '   Return: 0=no error, 1=error found
    '

    Public Function CheckCellXYError(ByVal CurrentSheetName As String, ByVal InputStr As String, _
        ByVal row As Integer, ByVal colum As Integer, ByRef sheet As AxOWC10.AxSpreadsheet, _
        ByRef ErrorSheetData As SubSheetErrorStruct) As Integer

        Dim Rtn As Integer = 0  'No error


        'Step 1: Fillin the activeX data cell
        Dim strBackup As String = sheet.ActiveSheet.Cells(row, colum).value
        sheet.ActiveSheet.Cells(row, colum).value() = InputStr


        'Step 2: Check error
        If 0 = BuildIntReference(sheet, ErrorSheetData) Then
            If BuildExtReference(sheet, ErrorSheetData) = 1 Then
                Rtn = 1
            End If
            If 0 <> CheckXYErrorInOneSheet(CurrentSheetName, sheet, ErrorSheetData) Then
                Rtn = 1
            End If
        Else
            Rtn = 1
        End If


        'Step 3: Restore activeX data cell if error found
        If 1 = Rtn Then
            sheet.ActiveSheet.Cells(row, colum).value() = strBackup
        End If


        Return Rtn
        TraceGCCollect
    End Function





    'Check errors for cells exclude XY and height related
    '   CmdStr: command string
    '   InputStr: Data value of the cell
    '   row: the row number of the cell
    '   colum: the column number of the cell
    '   sheet: the ActiveX spreadsheet
    '   ErrorSheetData: error checking data structure
    '
    '   Return: 0=no error, 1=error found
    '

    Public Function CheckOtherCellCharError(ByVal CmdStr As String, ByVal InputStr As String, ByVal Column As Integer) As Integer
        Dim Rtn As Integer = 0
        Dim m_ItemLUT As New CIDSItemsLUT

        m_ItemLUT.Cmd2Index(CmdStr.ToUpper)

        'Check if the cell should be empty
        If m_ItemLUT.GetFlag(Column) Then
            If InputStr = "" Then
                Rtn = 1
            End If
        Else
            If InputStr <> "" Then
                Rtn = 1
            End If
        End If

        'Check if the cell data is Appropriate
        If Rtn = 0 And InputStr <> "" Then

            Select Case (Column)
                Case gCommandNameColumn To gSubnameColumn
                    If True = IsNumeric(InputStr) Then
                        Rtn = 1
                    End If
                Case gTravelSpeedColumn To gNoOfRunColumn, gRtAngleColumn To gIONumberColumn
                    If CmdStr.ToUpper = "DOTARRAY" And Column = gTravelSpeedColumn Then
                        'check whether in proper format of integer x integer
                        Dim RowsAndColumns(1) As String
                        RowsAndColumns = InputStr.Split("x")
                        If (Int32.Parse(RowsAndColumns(0)) > 0 And Int32.Parse(RowsAndColumns(1)) > 0) Then
                            Rtn = 0
                        Else
                            Rtn = 1
                        End If
                        Exit Select
                    End If
                    If False = IsNumeric(InputStr) Then
                        Rtn = 1
                        Exit Select
                    Else
                        If Column = gTravelSpeedColumn Or Column = gRetractSpeedColumn Then
                            If (CInt(InputStr) <= 0) Or (CInt(InputStr) > 800) Then
                                Rtn = 1
                                Exit Select
                            End If
                        ElseIf (Column = gDurationColumn) Or (Column = gTravelDelayColumn) Or (Column = gRetractDelayColumn) Then
                            If (CInt(InputStr) < 0) Then
                                Rtn = 1
                                Exit Select
                            End If
                        End If
                    End If
                    If Column = gNoOfRunColumn Or Column = gIONumberColumn Then
                        'Data need to be integer type
                        If Math.Abs(CSng(InputStr) - CInt(InputStr)) > 0.000001 Then
                            Rtn = 1
                        End If
                        Exit Select
                    End If
                Case gSprialColumn
                    If "IN" <> InputStr.ToUpper() And "OUT" <> InputStr.ToUpper() Then
                        Rtn = 1
                    End If
            End Select

        End If

        Return Rtn
        TraceGCCollect
    End Function

End Class





'''''''''''''''''''''''''''''''
'                                                                                               
'Class CIDSUndo                                                                                  
'  Despription:                                                                                  
'           Pattern data undo handling class                                                      
'  Created by:                                                                                   
'                                                                                     
'                                                                                               
'''''''''''''''''''''''''''''''
Public Class CIDSUndo
    Public hasBackupData As Boolean = False 'Empty data sheet cannot be saved
    Public UndoFilename As String = ""      'Filename use as current

    Public UndoLevel As Integer
    'To optimize the speed, undo is classified in 2 levels.  Level 0 is slow
    '0: cross pages or multi-row include row delete; ; 
    '1: row level (one cell or one row)

    'Filename/Cell type "A" and "B" is used as undo and redo alternatively
    Public CurrentPageName_A As String
    Public SelectedCell1_A, SelectedCell2A As Object
    Public CurrentPageName_B As String
    Public SelectedCell1_B, SelectedCell2B As Object

    'Row level undo backup data
    Public UndoRow As Integer
    Public UndoPatternRec_A As CIDSPattern.PatternRecord
    Public UndoPatternRec_B As CIDSPattern.PatternRecord

    'Read activeX spreadsheet pattern file in excel or text format         
    '   Axspreadsheet:  instance of activeX spreadsheet               
    '   Filename:       Either type "A" or "B"                        
    '   Encrypt:        No                                             

    Public Sub DataSaveFor_Undo(ByVal AxSpreadsheet As AxOWC10.AxSpreadsheet)
        Try
            'Save data in text format --> faster
            m_Execution.m_Pattern.SavePatternParaAllSheets(AxSpreadsheet, _
                UndoFilename, 0, False)

            'Save data in excel format  --> Slower
            'AxSpreadsheet.Export(UndoFilename, _
            '   OWC.SheetExportActionEnum.ssExportActionNone, OWC.SheetExportFormat.ssExportAsAppropriate)
        Catch ex As SystemException
            ExceptionDisplay(ex)
        End Try
        hasBackupData = True
        TraceGCCollect
    End Sub

    'Write activeX spreadsheet pattern file in excel or text format            
    '   Axspreadsheet:  instance of activeX spreadsheet               
    '   Filename:       Either type "A" or "B"                        
    '   Encrypt:        No                                             

    Public Sub DataLoadFor_Undo(ByVal AxSpreadsheet As AxOWC10.AxSpreadsheet)
        Dim Ptn As New CIDSPattern

        Try
            'Load data in text format --> faster
            m_Execution.m_Pattern.LoadTxtPatternParaAllSheets(AxSpreadsheet, _
                UndoFilename, "", 0, 1, False)

            'Load data in excel format  --> Slower
            'Ptn.LoadExcelPatternFile(AxSpreadsheet, UndoFilename, 1, True, strUnit)
        Catch ex As SystemException
            ExceptionDisplay(ex)
        End Try
        TraceGCCollect
    End Sub
End Class