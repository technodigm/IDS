Option Explicit On 
Imports System
Imports System.Threading
Imports Microsoft.VisualBasic
Imports DLL_Export_Service
Imports Microsoft.DirectX.DirectInput
Imports System.Text
Imports Microsoft.Win32
Imports Microsoft.Office.Interop
'
'Class CIDSCommand                                                                               '
'  Despription:                                                                                  '  
'           Dispensing element data handling base class                                          '    
'  Created by:                                                                                   ' 
'           Shen Jian   

Public Class CIDSCommand

    Protected m_PatternList As New ArrayList  'pattern command list
    Protected m_MainSubList As New ArrayList  'main/sub sheet name list
    Protected m_DispenseList As New ArrayList 'dispensing command list
    Protected m_GlobalQCList As New ArrayList
    Protected m_CompilerStatus As Integer = 0 'compiling status flag
    'Protected m_Tri As CIDSTrioController = IDS.Devices.Motor
    Public globalQCParam As DLL_Export_Device_Vision.QC.QCParam
    Protected m_Optim As Integer = 0

    Public Sub New()

    End Sub



    'Get current pattern file ompilling status
    '   return:  1-finished and sucess;     0-not yet complied                                                                           
    Public ReadOnly Property CompileStatus() As Integer
        Get
            Return m_CompilerStatus
        End Get
    End Property


    'Get gattern file commands hold list
    '                                                                                

    Public ReadOnly Property PattenList() As ArrayList
        Get
            Return m_PatternList
        End Get
    End Property


    'Get dispensing command list after compiling
    '                                                                                

    Public ReadOnly Property DispenseList() As ArrayList
        Get
            Return m_DispenseList
        End Get
    End Property

    Public Function SetOptimizFlag(ByVal flag As Integer) As Integer
        m_Optim = flag
        Return 0
    End Function


    'Read pattern commands from spread sheet into pattern list
    '   file:  filename name for the export                                                                                

    Public Function ReadPattern(ByVal sheet As AxOWC10.AxSpreadsheet) As Integer
        Dim loadSheet As New CIDSPatternLoader(sheet)

        If m_MainSubList.Count > 0 Then
            m_MainSubList.Clear() 'RemoveRange(0, m_MainSubList.Count)
        End If

        Dim count As Integer = sheet.Worksheets.Count 'total sheets number
        Dim i, rtn As Integer
        Dim name As String

        'Set sheets name list
        For i = 1 To count
            name = sheet.Worksheets.Item(i).name
            m_MainSubList.Add(name)
        Next i

        'Generate pattern list from spread sheet
        loadSheet.SetOptimizFlag(m_Optim)
        If (m_Optim) Then
            Dim tmpList As New ArrayList
            tmpList = m_Execution.m_Command.DispenseList
            rtn = loadSheet.SetPatternList(m_MainSubList, tmpList, Me.globalQCParam) 'originally (m_MainSubList, tmpList)
            If rtn = 0 Or rtn = 100 Then DotOptimize(tmpList, m_PatternList)
        Else
            rtn = loadSheet.SetPatternList(m_MainSubList, m_PatternList, Me.globalQCParam)
        End If

        If rtn <> 0 Then
            Return rtn
        Else
            Return 0
        End If

    End Function

    Public Function DotOptimize(ByVal dotpatternList As ArrayList, ByRef optiPatternList As ArrayList)
        If dotpatternList.Count < 1 Then
            Return -1
        End If
        If optiPatternList.Count <> 0 Then
            optiPatternList.Clear()
        End If

        Dim elementRec, elementRec1 As Object
        Dim i, j, start_no, row, col As Integer
        Dim Type, Type1 As String
        Dim dot As CIDSDot


        Dim Index(100) As Integer
        Dim xelements(100) As Double
        Dim yelements(100) As Double
        Dim elementstmp(100) As Double

        Dim indextmp As Integer
        ' Dim xytmp() As Double
        Dim ytmp As Double
        row = dotpatternList.Count

        elementRec = dotpatternList(0)
        Type = elementRec.CmdType
        Type = Type.ToUpper
        If row > 1 Then
            elementRec1 = dotpatternList(1)
            Type1 = elementRec1.CmdType
            Type1 = Type1.ToUpper
        End If

        Dim indexRow As Integer
        If Type = "FIDUCIAL" And Type1 = "HEIGHT" Then
            optiPatternList.Add(elementRec)
            optiPatternList.Add(elementRec1)
            start_no = 3
            indexRow = row - 2
        ElseIf Type = "FIDUCIAL" Or Type = "HEIGHT" Then
            optiPatternList.Add(elementRec)
            start_no = 2
            indexRow = row - 1
        Else
            start_no = 1
            indexRow = row
        End If


        ReDim Index(indexRow - 1)
        ReDim xelements(indexRow - 1)
        ReDim yelements(indexRow - 1)
        ReDim elementstmp(indexRow - 1)
        For i = start_no To row
            elementRec = dotpatternList(i - 1)
            Type = elementRec.CmdType
            Type = Type.ToUpper
            If Type = "DOT" Then
                Type = "Dot"
                dot = elementRec
                xelements(i - start_no) = dot.PosX
                yelements(i - start_no) = dot.PosY
                'If start_no = 1 Then
                '    Index(i - 1) = i
                'ElseIf start_no = 2 Then
                '    Index(i - 2) = i - 1
                'Else
                '    Index(i - 3) = i - 2
                'End If
                Index(i - start_no) = i - start_no + 1
                'For j = 0 To col
                'dotarray(i - 1, j) = elementarray(j)
                'Next j
                'dot = elementRec
                'aa(1, 0) = dot.PosX
            End If
        Next i

        Array.Sort(xelements, Index)

        For i = 0 To indexRow - 1 Step 1
            indextmp = Index(i) + start_no - 1
            'elementstmp = dotpatternList(i - 1)
            optiPatternList.Add(dotpatternList(indextmp - 1))

        Next i

        Return 0
    End Function



    'Compiling pattern command list into dispensing command list
    '   runmode:  0-vision 1-dry 2-dry left 3-dry right 4-wet 5-dry left 6-dry right                                                                               

    Public Function Compile(ByVal runmode As Integer, Optional ByVal GlobalQCEnabled As Boolean = False) As Integer
        m_CompilerStatus = 0
        If m_PatternList.Count < 1 Then
            m_CompilerStatus = -1
            Return -1
        End If

        Dim comP As New IDSPattnCompiler(m_PatternList)
        comP.SetRunMode(runmode) 'set run mode for compiling
        'kr dec4 turn off when running

        If runmode <> 0 Then
            SwitchToRealTimeCamera()
        End If
        'generate dispensing command list
        If GlobalQCEnabled Then
            m_GlobalQCList.Clear()
            If comP.Compile(m_DispenseList, m_GlobalQCList) < 0 Then
                m_CompilerStatus = -1
                Return -1
            End If
            Dim lst As ArrayList = New ArrayList
            Dim vParam As DLL_Export_Device_Vision.QC.QCParam
            lst.Add(m_Execution.m_Command.globalQCParam._Binarized)
            If m_Execution.m_Command.globalQCParam._BlackDot Then
                lst.Add(1)
            Else
                lst.Add(0)
            End If
            lst.Add(m_Execution.m_Command.globalQCParam._Brightness)
            lst.Add(m_Execution.m_Command.globalQCParam._Close)
            lst.Add(m_Execution.m_Command.globalQCParam._Compactness)
            lst.Add(m_Execution.m_Command.globalQCParam._MaxArea)
            lst.Add(m_Execution.m_Command.globalQCParam._MinArea)
            lst.Add(m_Execution.m_Command.globalQCParam._MRegionX)
            lst.Add(m_Execution.m_Command.globalQCParam._MRegionY)
            lst.Add(m_Execution.m_Command.globalQCParam._MROIx)
            lst.Add(m_Execution.m_Command.globalQCParam._MROIy)
            lst.Add(m_Execution.m_Command.globalQCParam._Open)
            lst.Add(m_Execution.m_Command.globalQCParam._Roughness)
            lst.Add(m_Execution.m_Command.globalQCParam._Diameter)
            lst.Add(m_Execution.m_Command.globalQCParam._Tolerance)
            lst.Add(1) 'Turn of is globalQC
            'Insert into element 8th which is after first QC command
            'For global QC check, it only need to set once for the QC check settings
            'm_GlobalQCList.InsertRange(11, lst)
            m_GlobalQCList.InsertRange(13, lst)
            m_DispenseList.InsertRange(m_DispenseList.Count - 1, m_GlobalQCList)
            'm_DispenseList.AddRange(m_GlobalQCList)
        Else
            If comP.Compile(m_DispenseList) < 0 Then
                m_CompilerStatus = -1
                Return -1
            End If
        End If

        m_CompilerStatus = 1 'compiling finished flag
        Return 0
    End Function
End Class


'
'                                                                                              
'Class 3                                                                             
'  Despription:                                                                                  
'           Dispensing element data handling base class                                         
'  Created by:                                                                                  
'           Shen Jian                                                                          
'                          
Public Class CIDSPatternLoader
    Private m_MaxRows As Integer 'sheet occupied rows
    Private m_MaxColums As Integer 'sheet occupied colums
    Private m_Sheet As AxOWC10.AxSpreadsheet
    Protected m_Optim As Integer = 0
    Private Testing As Boolean = False

    Public Sub DebugAddList(ByVal lst As ArrayList, ByVal val As Object)
        lst.Add(val)
    End Sub

    'Set spreadsheet to be handled
    '   spreadsheet: an instance of spreadsheet                                                                              

    Public Sub New(ByVal spreadsheet As AxOWC10.AxSpreadsheet)
        m_Sheet = spreadsheet
    End Sub

    Public Function SetOptimizFlag(ByVal flag As Integer) As Integer
        m_Optim = flag
        Return 0
    End Function


    'Get worksheet used rows and columns
    '                                                                               

    Public Function GetSheetUsedRowsColums(ByVal actSheet As OWC10.Worksheet, ByRef Rows As Integer, ByRef Colums As Integer) As Integer
        If actSheet Is Nothing Then
            Return -1
        End If
        m_MaxRows = actSheet.UsedRange.Rows.Count
        Rows = m_MaxRows
        m_MaxColums = actSheet.UsedRange.Columns.Count
        Colums = m_MaxColums
        If m_MaxRows < 1 Or m_MaxColums = 0 Then
            Return -1
        End If
        Return (0)
    End Function


    'Fiducial and sub transform operation
    '   pt():       point to be transformed to hard home coornidates
    '   referPt():  reference point
    '   fidComp:    fiducial ompensation data
    '   newPt():    transformed point

    Public Function FidandSubTransform(ByVal pt() As Double, ByVal referPt() As Double, ByVal fidComp As FidCompData, ByRef newPt() As Double)
        Dim tmp(3), comp(3) As Double
        Dim z As Double
        'transform to work coordinate
        ReferToSys(pt, tmp, referPt)
        z = tmp(2)

        'sub-sub insert transform
        If fidComp.sub2Flag Then
            Dim suboff() As Double = {-fidComp.sub2FirstPtX, -fidComp.sub2FirstPtY}
            Translate2D(tmp, suboff, pt)
            Rotate2D(pt, fidComp.sub2InsAngle, tmp)
            Dim insPt() As Double = {fidComp.sub2InsPtX, fidComp.sub2InsPtY}
            Translate2D(tmp, insPt, tmp)
        End If

        'sub insert transform
        If fidComp.sub1Flag Then
            Dim suboff() As Double = {-fidComp.sub1FirstPtX, -fidComp.sub1FirstPtY}
            Translate2D(tmp, suboff, pt)
            Rotate2D(pt, fidComp.sub1InsAngle, tmp)
            Dim insPt() As Double = {fidComp.sub1InsPtX, fidComp.sub1InsPtY}
            Translate2D(tmp, insPt, tmp)
        End If

        'Fiducial Compensation
        tmp(2) = z
        If m_Optim = 0 Then
            'main fiducial
            If fidComp.gparentFidFlag Then
                Dim fdPt() As Double = {fidComp.gparentFidPtX, fidComp.gparentFidPtY, fidComp.gparentFidPtZ}
                Dim offset() As Double = {fidComp.gparentOffsetX, fidComp.gparentOffsetY, fidComp.gparentOffsetZ}
                FiducialComp(tmp, fdPt, offset, fidComp.gparentCompAngle, comp) 'wrf work co.
                comp.CopyTo(tmp, 0)
            End If
            'sub fiducial
            If fidComp.parentFidFlag Then
                Dim fdPt() As Double = {fidComp.parentFidPtX, fidComp.parentFidPtY, fidComp.parentFidPtZ}
                Dim offset() As Double = {fidComp.parentOffsetX, fidComp.parentOffsetY, fidComp.parentOffsetZ}
                FiducialComp(tmp, fdPt, offset, fidComp.parentCompAngle, comp) 'wrf work co.
                comp.CopyTo(tmp, 0)
            End If
            'sub-sub fiducial
            If fidComp.FidFlag Then
                Dim fdPt() As Double = {fidComp.FidPtX, fidComp.FidPtY, fidComp.FidPtZ}
                Dim offset() As Double = {fidComp.OffsetX, fidComp.OffsetY, fidComp.OffsetZ}
                FiducialComp(tmp, fdPt, offset, fidComp.CompAngle, comp) 'wrf work co.
                comp.CopyTo(tmp, 0)
            End If
            SysToHard(tmp, newPt)
        Else
            tmp.CopyTo(newPt, 0)
        End If
    End Function


    'Get current reference point's coordinates
    '  

    Public Function GetReferencePtData(ByVal sheet As OWC10.Worksheet, ByVal row As Integer, ByRef referPt() As Double) As Integer
        Dim Ratio As Double = 1.0
        Dim array As Object(,) ' array start at (1,1)
        array = sheet.Range("A" & row, "AD" & row).Value
        referPt(0) = array(1, gPos1XColumn)
        referPt(1) = array(1, gPos1YColumn)
        referPt(2) = 0.0 'sheet.Cells(row, gPos1ZColum).value

        Return 0
    End Function

    'Move robot to position x,y to detect height 

    Public Function MoveToCheckHeight(ByVal x As Double, ByVal y As Double, ByRef height As Double) As Integer

        Dim pos() As Double = {x, y}
        m_Tri.Set_XY_Speed(IDS.Data.Hardware.Gantry.ElementXYSpeed)

        If Not m_Tri.Move_XY(pos) Then Return -1
        If CheckButtonState() = -1 Then
            Return 100
        End If
        Dim rtn As Boolean
        Dim height1, height2 As Double
        OnLaser()
        rtn = Laser.WaitForReadingToStabilize
        OffLaser()
        'Dim start_time As DateTime = Now
        'Dim stop_time As DateTime = Now
        'Dim elapsed_time As TimeSpan
        'Do
        '    Application.DoEvents()
        '    stop_time = Now
        '    elapsed_time = stop_time.Subtract(start_time)
        'Loop Until elapsed_time.TotalSeconds > 3
        If rtn Then
            height = Laser.LASER_Reading
        Else
            If CheckButtonState() = -1 Then
                Return 100
            Else
                Return -1
            End If
        End If
        Return 0

    End Function



    'Set height record
    '      sheet: pattern sheet
    '      row:   pattern command row number

    Public Function SetHeightRecData(ByVal sheet As OWC10.Worksheet, ByVal row As Integer, ByRef data As CIDSHeight) As Integer
        Dim array As Object(,) ' array start at (1,1)
        Dim j = 1
        array = sheet.Range("A" & row, "AD" & row).Value

        data.PosX = array(j, gPos1XColumn)
        data.PosY = array(j, gPos1YColumn)
        data.PosZ = array(j, gPos1ZColumn)
        data.CmdLineNo = row
        data.CmdType = "HEIGHT"
        Return 0
    End Function




    'Move robot to position x,y to detect height 
    '      sheet: pattern sheet
    '      row:   pattern command row number
    '      referPt: reference point
    '      fidComp: sub/fiducial compensation data
    '      heightComps:  height compensation value  

    Public Function GetHeightCompensation(ByVal sheet As OWC10.Worksheet, ByVal row As Integer, _
                                          ByVal referPt() As Double, ByVal fidComp As FidCompData, ByRef heightComps As Double) As Integer
        Dim pt(3), comp(3), tmp(3) As Double
        Dim PosX, PosY, height As Double

        Dim offX1, offY1, offX2, offY2 As Double
        heightComps = 0.0

        Dim array As Object(,) ' array start at (1,1)
        array = sheet.Range("A" & row, "AD" & row).Value
        Dim j As Integer = 1
        'get detecting height position wrf work co.
        pt(0) = array(j, gPos1XColumn)
        pt(1) = array(j, gPos1YColumn)
        pt(2) = array(j, gPos1ZColumn)
        FidandSubTransform(pt, referPt, fidComp, comp) 'transform to hard home coordinate

        'offset laser 
        PosX = comp(0) - gLaserOffX
        PosY = comp(1) - gLaserOffY
        Dim sheetname As String = sheet.Name
        Dim temp() As Double = {PosX, PosY, comp(2)}
        'If Not WorkAreaErrorCheckXY(temp) Then 'check XY range 
        '    CompileErrorDisplay(sheetname, row, 0)
        '    Return -1
        'End If
        Dim rtn As Integer
        'If ProgrammingMode() Then
        '    'Programming.tbLsHeight.Text = ""
        'End If
        rtn = MoveToCheckHeight(PosX, PosY, height) 'Move to get height
        If rtn < 0 Then
            CompileErrorDisplay(sheetname, row, 6)
            Return -1
        ElseIf rtn > 0 Then
            Return rtn
        End If
        heightComps = height - IDS.Data.Hardware.HeightSensor.Laser.HeightReference
        Return 0
    End Function


    'Move robot to position x,y to check reject mark
    '  

    Public Function MoveToCheckRejectMark(ByVal x As Double, ByVal y As Double, ByVal vpara As DLL_Export_Device_Vision.RejectPoint.RMParam) As Integer

        'move robot to position (x,y)
        Dim pos() As Double = {x, y}
        m_Tri.Set_XY_Speed(IDS.Data.Hardware.Gantry.ElementXYSpeed)

        If Not m_Tri.Move_XY(pos) Then Return -1
        If CheckButtonState() = -1 Then
            Return 100
        End If


        Dim rtn As Boolean
        'Check reject mark
        rtn = Vision.IDSV_RM(vpara)
        ClearDisplay()
        If rtn = False Then
            If CheckButtonState() = -1 Then
                Return 100
            Else
                Return -1
            End If
        Else
            Return 0
        End If

    End Function


    'Check reject mark 
    '      sheet: pattern sheet
    '      row:   pattern command row number
    '      data:  chipedge record data
    '      referPt: reference point
    '      fidComp: sub/fiducial compensation data

    Public Function CheckRejectMark(ByVal sheet As OWC10.Worksheet, ByVal row As Integer, _
                                          ByVal referPt() As Double, ByVal fidComp As FidCompData) As Integer
        Dim pt(3), comp(3), tmp(3) As Double
        Dim PosX, PosY As Double

        Dim vpara As DLL_Export_Device_Vision.RejectPoint.RMParam
        Dim offX1, offY1, offX2, offY2 As Double

        Dim array As Object(,) ' array start at (1,1)
        array = sheet.Range("A" & row, "BZ" & row).Value
        Dim j As Integer = 1

        'Get reject mark position wrt work coord.
        pt(0) = array(j, gPos1XColumn)
        pt(1) = array(j, gPos1YColumn)
        pt(2) = array(j, gPos1ZColumn)
        FidandSubTransform(pt, referPt, fidComp, comp) 'Transform to hard home coord.

        If Not WorkAreaErrorCheckXY(comp) Then
            Dim sheetname As String = sheet.Name
            CompileErrorDisplay(sheetname, row, 0)
            Return -1
        End If

        PosX = comp(0)
        PosY = comp(1)

        'set vision parameters for reject mark detection
        vpara._AcceptRatio = array(j, gAcceptRatioCoulumn)
        vpara._Binarized = array(j, gBinarizedColumn)
        vpara._BlackWithoutRM = array(j, gBlackWithoutRMCoulumn)
        vpara._BlackWithRM = array(j, gBlackWithRMCoulumn)
        vpara._Brightness = array(j, gBrightnessColumn)
        vpara._MRegionX = array(j, gMRegionXColumn)
        vpara._MRegionY = array(j, gMRegionYColumn)
        vpara._MROIx = array(j, gMROIxColumn)
        vpara._MROIy = array(j, gMROIyColumn)
        vpara._WhiteWithoutRM = array(j, gWhiteWithoutRMCoulumn)
        vpara._WhiteWithRM = array(j, gWhiteWithRMCoulumn)
        vpara._WoRM = array(j, gWoRMCoulumn)
        vpara._WRM = array(j, gWRMCoulumn)

        ''
        '   Xue Wen                                                                                 '
        '   Set the brightness before doing the movement. This will affect vision(ActiveX) less.    '
        ''
        Vision.IDSV_SetBrightness(array(j, gBrightnessColumn))

        Dim rtn As Integer = 0
        rtn = MoveToCheckRejectMark(PosX, PosY, vpara)
        If rtn < 0 Then 'detected reject mark
            Return -1
        ElseIf rtn > 0 Then
            Return rtn
        End If
        Return 0
    End Function

    'Set Fiducial record dara
    '      sheet: pattern sheet
    '      row:   pattern command row number
    '      data:  Fiducial record data

    Public Function SetFiducialRecData(ByVal sheet As OWC10.Worksheet, ByVal row As Integer, ByRef data As CIDSFiducial) As Integer
        Dim pt(3), comp(3) As Double
        Dim para As DispensePara

        data.CmdLineNo = row
        data.CmdType = "FIDUCIAL"
        data.SheetName = sheet.Name

        Dim array As Object(,) ' array start at (1,1)
        Dim j = 1
        array = sheet.Range("A" & row, "AD" & row).Value

        data.PosX1 = array(j, gPos1XColumn)
        data.PosY1 = array(j, gPos1YColumn)
        data.PosZ1 = array(j, gPos1ZColumn)
        data.PosX2 = array(j, gPos2XColumn)
        data.PosY2 = array(j, gPos2YColumn)
        data.PosZ2 = array(j, gPos2ZColumn)
        data.FirstFid = array(j, gFid1Column)
        data.SecondFid = array(j, gFid2Column)
        If data.SecondFid = Nothing Then
            data.NumOfFid = 1
        Else
            data.NumOfFid = 2
        End If

        Return 0
    End Function



    'Check fiducial points
    '      sheet: pattern sheet
    '      row:   pattern command row number
    '      referPt: reference point
    '      fidComp: sub/fiducial compensation data
    '      fdPt:  fiducial point
    '      offset: offset between command and real fiducial location
    '      angle: orintation difference of teaching with real angle 

    Public Function CheckFiducialPt(ByVal sheet As OWC10.Worksheet, ByVal row As Integer, _
                                          ByVal referPt() As Double, ByVal fidComp As FidCompData, ByRef fdPt() As Double, ByRef offset() As Double, ByRef angle As Double) As Integer
        Dim pt(3), comp(3), tmp(3) As Double
        Dim PosX1, PosX2, PosY1, PosY2 As Double
        ' Dim NumofFid As Integer
        Dim FidNames As String
        Dim offX1, offY1, offX2, offY2 As Double
        'Get fiducial point 1 position

        Dim array As Object(,) ' array start at (1,1)
        array = sheet.Range("A" & row, "BZ" & row).Value
        Dim j As Integer = 1

        pt(0) = array(j, gPos1XColumn)
        pt(1) = array(j, gPos1YColumn)
        pt(2) = array(j, gPos1ZColumn)
        FidandSubTransform(pt, referPt, fidComp, comp)

        If Not WorkAreaErrorCheckXY(comp) Then Return -1

        PosX1 = comp(0)
        PosY1 = comp(1)
        'Get fiducial point 2 position
        pt(0) = array(j, gPos2XColumn)
        pt(1) = array(j, gPos2YColumn)
        pt(2) = array(j, gPos2ZColumn)
        FidandSubTransform(pt, referPt, fidComp, comp)
        Dim sheetname As String = sheet.Name
        If Not WorkAreaErrorCheckXY(comp) Then
            CompileErrorDisplay(sheetname, row, 0)
            Return -1
        End If

        PosX2 = comp(0)
        PosY2 = comp(1)

        Dim rtn As Integer = 0
        Dim result As Boolean
        'only one fiducial point
        If Math.Abs(PosX1 - PosX2) < 1.0 And Math.Abs(PosY1 - PosY2) < 1.0 Then
            FidNames = array(j, gFid1Column)
            FidNames = gFidFileName + FidNames
            Dim Brightness As Integer
            Brightness = array(j, gBrightnessColumn)
            Vision.IDSV_SetBrightness(Brightness)
            position(0) = PosX1
            position(1) = PosY1
            m_Tri.Set_XY_Speed(IDS.Data.Hardware.Gantry.ElementXYSpeed)
            If CheckButtonState() = -1 Then
                Return 100
            End If
            'Vision.FrmVision.CameraIdle()
            If Not m_Tri.Move_XY(position) Then Return -1
            If CheckButtonState() = -1 Then
                Return 100
            End If
            'OneGrab()
            'MySleep(50)
            DisplayCrossHair()
            result = Vision.IDSV_FI(FidNames, offX1, offY1)
            'ClearDisplay()
            If result = False Then
                If CheckButtonState() = -1 Then
                    Return 100
                Else
                    Return -1
                End If
            Else
                rtn = 0
            End If
            If rtn < 0 Then  'no autoskip so check!
                CompileErrorDisplay(sheetname, row, 5)
                Return -1
            ElseIf rtn > 0 Then
                Return rtn
            End If
            fdPt(0) = PosX1 - IDS.Data.Hardware.Gantry.SystemOriginPos.X
            fdPt(1) = PosY1 - IDS.Data.Hardware.Gantry.SystemOriginPos.Y
            offset(0) = offX1
            offset(1) = -offY1
            angle = 0.0
            Return 0
        End If
        'two fiducial points checking
        'fiducial 1
        FidNames = array(j, gFid1Column)
        FidNames = gFidFileName + FidNames
        Dim Brightness1 As Integer
        Brightness1 = array(j, gBrightnessColumn)
        Vision.IDSV_SetBrightness(Brightness1)
        position(0) = PosX1
        position(1) = PosY1
        m_Tri.Set_XY_Speed(IDS.Data.Hardware.Gantry.ElementXYSpeed)
        If CheckButtonState() = -1 Then
            Return 100
        End If
        If Not m_Tri.Move_XY(position) Then Return -1
        If CheckButtonState() = -1 Then
            Return 100
        End If
        'OneGrab()
        'MySleep(50)
        DisplayCrossHair()
        result = Vision.IDSV_FI(FidNames, offX1, offY1)
        'ClearDisplay()
        If result = False Then
            If CheckButtonState() = -1 Then
                Return 100
            Else
                Return -1
            End If
        Else
            rtn = 0
        End If
        If rtn < 0 Then
            CompileErrorDisplay(sheetname, row, 5)
            Return -1
        ElseIf rtn > 0 Then
            Return rtn
        End If

        'fiducial 2
        FidNames = array(j, gFid2Column)
        FidNames = gFidFileName + FidNames
        Dim Brightness2 As Integer
        Brightness2 = array(j, gThresholdColumn)
        Vision.IDSV_SetBrightness(Brightness2)
        position(0) = PosX2
        position(1) = PosY2
        m_Tri.Set_XY_Speed(IDS.Data.Hardware.Gantry.ElementXYSpeed)
        If Not m_Tri.Move_XY(position) Then Return -1
        If CheckButtonState() = -1 Then
            Return 100
        End If
        'OneGrab()
        'MySleep(50)
        DisplayCrossHair()
        result = Vision.IDSV_FI(FidNames, offX2, offY2)
        'ClearDisplay()
        If result = False Then
            If CheckButtonState() = -1 Then
                Return 100
            Else
                Return -1
            End If
        Else
            rtn = 0
        End If

        If rtn < 0 Then
            CompileErrorDisplay(sheetname, row, 5)
            Return -1
        ElseIf rtn > 0 Then
            Return rtn
        End If

        'Set fiducial compensation parameter rotating point, angle and offset value
        fdPt(0) = (PosX1 + PosX2) / 2.0 - gSystemOrigin(0)
        fdPt(1) = (PosY1 + PosY2) / 2.0 - gSystemOrigin(1)
        Dim L1pt1() As Double = {PosX1, PosY1}
        Dim L1pt2() As Double = {PosX2, PosY2}
        Dim L2pt1(2), L2pt2(2) As Double
        L2pt1(0) = PosX1 + offX1
        L2pt1(1) = PosY1 - offY1
        L2pt2(0) = PosX2 + offX2
        L2pt2(1) = PosY2 - offY2
        If Get2LineAngleOffset(L1pt1, L1pt2, L2pt1, L2pt2, offset, angle) < 0 Then Return -1
        Return 0
    End Function
    'only for debug test
    Public Function MoveToCheckQC(ByVal x As Double, ByVal y As Double, ByVal vParam As DLL_Export_Device_Vision.QC.QCParam) As Integer

        Dim pos() As Double = {x, y}
        m_Tri.Set_XY_Speed(IDS.Data.Hardware.Gantry.ElementXYSpeed)

        If Not m_Tri.Move_XY(pos) Then Return -1
        If CheckButtonState() = -1 Then
            Return 100
        End If

        Dim id As Object = "0"

        Dim rtn As Boolean = Vision.IDSV_QC(vParam, True)
        ClearDisplay()
        If rtn = False Then
            If CheckButtonState() = -1 Then
                Return 100
            Else
                Return -1
            End If
        End If
        Return 0

    End Function
    Public Function SetGlobalQCData(ByVal sheet As OWC10.Worksheet, ByVal row As Integer, ByRef vPara As DLL_Export_Device_Vision.QC.QCParam)
        Dim array As Object(,) ' array start at (1,1)
        Dim j = 1
        array = sheet.Range("A" & row, "BZ" & row).Value
        'set vision data
        vPara._Brightness = array(j, gBrightnessColumn)
        vPara._Binarized = array(j, gBinarizedColumn)
        vPara._BlackDot = array(j, gBlackDotColumn)
        vPara._Open = array(j, gOpenColumn)
        vPara._Close = array(j, gCloseColumn)
        vPara._Compactness = array(j, gCompactnessColumn)
        vPara._MaxArea = array(j, gMaxAreaColumn)
        vPara._MinArea = array(j, gMinAreaColumn)
        vPara._MRegionX = array(j, gMRegionXColumn)
        vPara._MRegionY = array(j, gMRegionYColumn)
        vPara._MROIx = array(j, gMROIxColumn)
        vPara._MROIy = array(j, gMROIyColumn)
        vPara._Roughness = array(j, gRoughnessColumn)
        vPara._Diameter = array(j, gDiameterColumn)
        vPara._Tolerance = array(j, gToleranceColumn)
    End Function

    'Set QC record dara
    '      sheet: pattern sheet
    '      row:   pattern command row number
    '      data:  QC record data
    '      referPt: reference point
    '      fidComp: sub/fiducial compensation data
    '      height:  height compensation value

    Public Function SetQCRecData(ByVal sheet As OWC10.Worksheet, ByVal row As Integer, ByRef data As CIDSQC, _
                                             ByVal referPt() As Double, ByVal fidComp As FidCompData) As Integer
        Dim pt(3), comp(3) As Double
        Dim para As DispensePara
        Dim vPara As DLL_Export_Device_Vision.QC.QCParam

        data.CmdLineNo = row
        data.CmdType = "QC"
        data.SheetName = sheet.Name

        Dim array As Object(,) ' array start at (1,1)
        Dim j = 1
        array = sheet.Range("A" & row, "BZ" & row).Value

        pt(0) = array(j, gPos1XColumn)
        pt(1) = array(j, gPos1YColumn)
        pt(2) = array(j, gPos1ZColumn)
        FidandSubTransform(pt, referPt, fidComp, comp) 'Transform QC position to hard home coord.

        Dim sheetname As String = sheet.Name
        If Not WorkAreaErrorCheckXY(comp) Then 'check xy range
            CompileErrorDisplay(sheetname, row, 0)
            Return -1
        End If

        data.PosX = comp(0)
        data.PosY = comp(1)
        data.PosZ = comp(2)

        'set vision data
        vPara._Brightness = array(j, gBrightnessColumn)
        vPara._Binarized = array(j, gBinarizedColumn)
        vPara._BlackDot = array(j, gBlackDotColumn)
        vPara._Open = array(j, gOpenColumn)
        vPara._Close = array(j, gCloseColumn)
        vPara._Compactness = array(j, gCompactnessColumn)
        vPara._MaxArea = array(j, gMaxAreaColumn)
        vPara._MinArea = array(j, gMinAreaColumn)
        vPara._MRegionX = array(j, gMRegionXColumn)
        vPara._MRegionY = array(j, gMRegionYColumn)
        vPara._MROIx = array(j, gMROIxColumn)
        vPara._MROIy = array(j, gMROIyColumn)
        vPara._Roughness = array(j, gRoughnessColumn)
        vPara._Diameter = array(j, gDiameterColumn)
        vPara._Tolerance = array(j, gToleranceColumn)

        data.vParam = vPara

        Return 0
    End Function


    'Check chipedge and set its record data
    '      sheet: pattern sheet
    '      row:   pattern command row number
    '      data:  chipedge record data
    '      referPt: reference point
    '      fidComp: sub/fiducial compensation data
    '      height:  height compensation value

    Public Function ChecknSetChipedgeRecData(ByVal sheet As OWC10.Worksheet, ByVal row As Integer, ByRef data As CIDSChipEdge, _
                                             ByVal referPt() As Double, ByVal fidComp As FidCompData, ByVal height As Double) As Integer
        Dim pt(3), comp(3) As Double
        Dim PosX, PosY As Double
        Dim para As DispensePara
        Dim vPara As DLL_Export_Device_Vision.ChipEdgePoints.ChipEdgeParam

        data.CmdLineNo = row
        data.CmdType = "CHIPEDGE"
        data.SheetName = sheet.Name

        Dim array As Object(,) ' array start at (1,1)
        array = sheet.Range("A" & row, "BZ" & row).Value
        Dim j As Integer = 1

        'Set chip edge detecting position
        pt(0) = array(j, gPos1XColumn)
        pt(1) = array(j, gPos1YColumn)
        pt(2) = array(j, gPos1ZColumn)
        FidandSubTransform(pt, referPt, fidComp, comp)

        Dim sheetname As String = sheet.Name
        If Not WorkAreaErrorCheckXY(comp) Then
            CompileErrorDisplay(sheetname, row, 0)
            Return -1
        End If


        data.PosX = comp(0)
        data.PosY = comp(1)
        data.PosZ = comp(2)

        PosX = data.PosX
        PosY = data.PosY

        'set vision paramerers
        vPara._Brightness = array(j, gBrightnessColumn)
        vPara._EdgeClearance = array(j, gEdgeClearColumn)
        vPara._CheckBox_ChipRec_Enable = array(j, gCheckBoxColumn)
        vPara._Cw_CCw = array(j, gCwCCwColumn)
        vPara._DispenseModel = array(j, gDispModelColumn)
        vPara._Inside_out = array(j, gInOutColumn)
        vPara._MainEdge = array(j, gMainEdgeColumn)
        vPara._PointX1 = array(j, gPointX1Column)
        vPara._PointX2 = array(j, gPointX2Column)
        vPara._PointX3 = array(j, gPointX3Column)
        vPara._PointX4 = array(j, gPointX4Column)
        vPara._PointX5 = array(j, gPointX5Column)
        vPara._PointY1 = array(j, gPointY1Column)
        vPara._PointY2 = array(j, gPointY2Column)
        vPara._PointY3 = array(j, gPointY3Column)
        vPara._PointY4 = array(j, gPointY4Column)
        vPara._PointY5 = array(j, gPointY5Column)
        vPara._Pos = array(j, gPosColumn)
        vPara._PosX = array(j, gPosXColumn)
        vPara._PosY = array(j, gPosYColumn)
        vPara._ROI = array(j, gROIColumn)
        vPara._Rot = array(j, gRotColumn)
        vPara._Size = array(j, gSizeColumn)
        vPara._SizeX = array(j, gSizeXColumn)
        vPara._SizeY = array(j, gSizeYColumn)
        vPara._Threshold = array(j, gThresholdColumn)
        vPara._Vertical = array(j, gVerticalColumn)
        vPara._Polarity = array(j, gPolarityColumn)
        vPara._EdgeStrength = array(j, gEdgeStrengthColumn)
        vPara._DotDispensingDuration = array(j, gDurationColumn)
        vPara._Contrast = array(j, gCompactnessColumn)
        ''
        '   Xue Wen                                                                                 '
        '   Set the brightness before doing the movement. This will affect vision(ActiveX) less.    '
        ''
        Vision.IDSV_SetBrightness(array(j, gBrightnessColumn))

        Dim rtn As Integer = 0
        rtn = MoveToCheckChipEdge(PosX, PosY, vPara) 'move to check chip edge
        If rtn < 0 Then
            'If ProgrammingMode() Then
            '    CompileErrorDisplay(sheetname, row, 8)
            'End If
            Return -1
        ElseIf rtn > 0 Then
            Return rtn
        End If
        If Testing Then
            Dim sTime As Long = DateTime.Now.Ticks
            While ((DateTime.Now.Ticks - sTime) / 10000) < 2000
                Application.DoEvents()
            End While
        End If
       

        data.vParam = vPara
        data.HeightCompS = height
        Dim dispStr As String = array(j, gDispensColumn)
        dispStr = dispStr.ToUpper()
        If dispStr = "ON" Then
            para.DispenseOn = True
        Else
            para.DispenseOn = False
        End If

        para.NeedleGap = array(j, gNeedleGapColumn)
        para.TravelDelay = array(j, gTravelDelayColumn)
        para.TravelSpeed = array(j, gTravelSpeedColumn)
        para.DeTailDist = array(j, gDTaiDistColumn)
        para.ApproachHeight = array(j, gApproachHtColumn)
        para.RetractDelay = array(j, gRetractDelayColumn)
        para.RetractSpeed = array(j, gRetractSpeedColumn)
        para.RetractHeight = array(j, gRetractHtColumn)
        para.ClearanceHeight = array(j, gClearanceHtColumn)
        para.ArcRadius = array(j, gArcRadiusColumn)
        para.Needle = array(j, gNeedleColumn)
        para.EdgeClearance = array(j, gEdgeClearColumn)
        data.Param = para
        Return 0
    End Function
    Private Sub OneGrab()
        Dim sTime As Long = DateTime.Now.Ticks
        'Motion settling time
        While (DateTime.Now.Ticks - sTime) / 10000 < 100
            Application.DoEvents()
        End While
        Vision.FrmVision.CameraResume()
        sTime = DateTime.Now.Ticks
        While Not (Vision.FrmVision.ImageRefreshed) And (DateTime.Now.Ticks - sTime) / 10000 < 300
            Application.DoEvents()
        End While
        Vision.FrmVision.CameraIdle()
    End Sub

    'Move to check chip edge
    '      x,y:   checking position
    '      para:  detected chip edge data
    '
    Public lastError As String = ""
    Public Function MoveToCheckChipEdge(ByVal x As Double, ByVal y As Double, ByRef para As DLL_Export_Device_Vision.ChipEdgePoints.ChipEdgeParam) As Integer
        Dim pos() As Double = {x, y}
        m_Tri.Set_XY_Speed(IDS.Data.Hardware.Gantry.ElementXYSpeed)
        If CheckButtonState() = -1 Then
            Return 100
        End If
        'Vision.FrmVision.CameraIdle()
        If Not m_Tri.Move_XY(pos) Then Return -1
        'Dim sTime As Long = DateTime.Now.Ticks
        ''Motion settling time
        'While (DateTime.Now.Ticks - sTime) / 10000 < 100
        '    Application.DoEvents()
        'End While
        'Vision.FrmVision.CameraResume()
        'sTime = DateTime.Now.Ticks
        'While Not (Vision.FrmVision.ImageRefreshed) And (DateTime.Now.Ticks - sTime) / 10000 < 100
        '    Application.DoEvents()
        'End While
        Vision.FrmVision.CameraIdle()
        OneGrab()
        If CheckButtonState() = -1 Then
            Vision.FrmVision.CameraResume()
            Return 100
        End If
        Vision.FrmVision.DisplayIndicator()
        Dim x1, y1, x2, y2, x3, y3, x4, y4 As Double
        Dim retry As Integer = 0
        While (retry < 5)
            retry += 1
            OneGrab()
            If Not Vision.IDSV_CE(para) Then 'detecting chip edge
                LabelMessage("Chip edge detect retry #" & retry)
                lastError = Vision.lastError
                If CheckButtonState() = -1 Then
                    Vision.FrmVision.CameraResume()
                    Return 100
                Else
                    If retry >= 4 Then
                        Vision.FrmVision.CameraResume()
                        Return -1
                    End If
                End If
            Else : Exit While
            End If
        End While
        'ClearDisplay()
        Vision.IDSV_CEOutput(x1, y1, x2, y2, x3, y3, x4, y4) 'get four points wrt pos(x,y)
        Vision.FrmVision.CameraResume()
        'convert four points to hard home coord.
        para._PointX1 = x + x1
        para._PointY1 = y - y1
        para._PointX2 = x + x2
        para._PointY2 = y - y2
        para._PointX3 = x + x3
        para._PointY3 = y - y3
        para._PointX4 = x + x4
        para._PointY4 = y - y4
        Return 0
    End Function

    'Set wait command record data
    '      sheet: pattern sheet
    '      row:   pattern command row number
    '      data:  wait record data
    '      referPt: reference point
    '      fidComp: sub/fiducial compensation data
    '      height:  height compensation value

    Public Function SetWaitRecordData(ByVal sheet As OWC10.Worksheet, ByVal row As Integer, ByRef data As CIDSWait, ByVal referPt() As Double, ByVal fidComp As FidCompData, ByVal height As Double) As Integer
        Dim para As DispensePara
        Dim pt(3), comp(3), tmp(3) As Double
        Dim array As Object(,) ' array start at (1,1)
        Dim j = 1
        array = sheet.Range("A" & row, "AD" & row).Value

        pt(0) = array(j, gPos1XColumn)
        pt(1) = array(j, gPos1YColumn)
        pt(2) = array(j, gPos1ZColumn)
        para.Needle = array(j, gNeedleColumn)
        FidandSubTransform(pt, referPt, fidComp, comp)
        If Not WorkAreaErrorCheckXY(comp) Then
            Dim sheetname As String = sheet.Name
            CompileErrorDisplay(sheetname, row, 0)
            Return -1
        End If
        data.CmdLineNo = row
        data.CmdType = "WAIT"
        data.SheetName = sheet.Name
        para.Duration = array(j, gDurationColumn)
        data.PosX = comp(0)
        data.PosY = comp(1)
        data.PosZ = comp(2)
        data.HeightCompS = height
        data.Param = para
        Return 0
    End Function

    'Set setio command record data
    '      sheet: pattern sheet
    '      row:   pattern command row number
    '      data:  setio record data

    Public Function SetSetIORecordData(ByVal sheet As OWC10.Worksheet, ByVal row As Integer, ByRef data As CIDSSetIO) As Integer
        Dim OnOff As String
        data.CmdLineNo = row
        data.CmdType = "SETIO"
        data.SheetName = sheet.Name
        data.SetIONo = sheet.Cells(row, gIONumberColumn).value
        If data.SetIONo < 32 Or data.SetIONo > 47 Then 'check io range
            Dim sheetname As String = sheet.Name
            CompileErrorDisplay(sheetname, row, 14)
            Return -1
        End If
        OnOff = sheet.Cells(row, gSetIOOnOffColumn).value
        If OnOff.ToUpper = "ON" Then
            data.SetIOOnOffFlag = 1
        ElseIf OnOff.ToUpper = "OFF" Then
            data.SetIOOnOffFlag = 0
        Else
            Return -1
        End If
        Return 0
    End Function


    'Set getio command record data
    '      sheet: pattern sheet
    '      row:   pattern command row number
    '      data:  getio record data
    '

    Public Function SetGetIORecordData(ByVal sheet As OWC10.Worksheet, ByVal row As Integer, ByRef data As CIDSGetIO) As Integer
        data.CmdLineNo = row
        data.CmdType = "GETIO"
        data.SheetName = sheet.Name
        data.GetIONo = sheet.Cells(row, gIONumberColumn).value
        If data.GetIONo < 32 Or data.GetIONo > 47 Then
            Dim sheetname As String = sheet.Name
            CompileErrorDisplay(sheetname, row, 14)
            Return -1
        End If
        Return 0
    End Function


    'Set clean command record data
    '      sheet: pattern sheet
    '      row:   pattern command row number
    '      data:  clean record data
    '

    Public Function SetVolCalRecordData(ByVal sheet As OWC10.Worksheet, ByVal row As Integer, ByRef data As CIDSVolumeCalibration) As Integer
        Dim para As DispensePara
        data.CmdLineNo = row
        data.CmdType = "VOLUMECALIBRATION"
        data.SheetName = sheet.Name
        Dim array As Object(,) ' array start at (1,1)
        Dim j = 1
        array = sheet.Range("A" & row, "AD" & row).Value

        para.Needle = array(j, gNeedleColumn)
        para.Duration = array(j, gDurationColumn)
        data.Param = para
        Return 0
    End Function

    'Set clean command record data
    '      sheet: pattern sheet
    '      row:   pattern command row number
    '      data:  clean record data
    '

    Public Function SetCleanRecordData(ByVal sheet As OWC10.Worksheet, ByVal row As Integer, ByRef data As CIDSClean) As Integer
        Dim para As DispensePara
        data.CmdLineNo = row
        data.CmdType = "CLEAN"
        data.SheetName = sheet.Name
        Dim array As Object(,) ' array start at (1,1)
        Dim j = 1
        array = sheet.Range("A" & row, "AD" & row).Value
        para.Needle = array(j, gNeedleColumn)
        para.Duration = array(j, gDurationColumn)
        data.Param = para
        Return 0
    End Function


    'Set purge command record data
    '      sheet: pattern sheet
    '      row:   pattern command row number
    '      data:  purge record data
    ' 

    Public Function SetPurgeRecordData(ByVal sheet As OWC10.Worksheet, ByVal row As Integer, ByRef data As CIDSPurge) As Integer
        Dim para As DispensePara
        data.CmdLineNo = row
        data.CmdType = "PURGE"
        data.SheetName = sheet.Name
        Dim array As Object(,) ' array start at (1,1)
        Dim j = 1
        array = sheet.Range("A" & row, "AD" & row).Value

        para.Needle = array(j, gNeedleColumn)
        para.Duration = array(j, gDurationColumn)
        data.Param = para
        Return 0
    End Function


    'Set move command record data
    '      sheet: pattern sheet
    '      row:   pattern command row number
    '      data:  move record data
    '      referPt: reference point
    '      fidComp: sub/fiducial compensation data
    '      height:  height compensation value

    Public Function SetMoveRecordData(ByVal sheet As OWC10.Worksheet, ByVal row As Integer, ByRef data As CIDSMove, _
                                      ByVal referPt() As Double, ByVal fidComp As FidCompData, ByVal height As Double) As Integer
        Dim para As DispensePara
        Dim pt(3), comp(3), tmp(3) As Double

        Dim array As Object(,) ' array start at (1,1)
        Dim j = 1
        array = sheet.Range("A" & row, "AD" & row).Value

        pt(0) = array(j, gPos1XColumn)
        pt(1) = array(j, gPos1YColumn)
        pt(2) = array(j, gPos1ZColumn)
        para.ArcRadius = array(j, gArcRadiusColumn)
        para.Needle = array(j, gNeedleColumn)
        FidandSubTransform(pt, referPt, fidComp, comp)

        If Not WorkAreaErrorCheckXY(comp) Then
            Dim sheetname As String = sheet.Name
            CompileErrorDisplay(sheetname, row, 0)
            Return -1
        End If


        data.CmdLineNo = row
        data.CmdType = "MOVE"
        data.SheetName = sheet.Name
        data.PosX = comp(0)
        data.PosY = comp(1)
        data.PosZ = comp(2)
        data.HeightCompS = height
        data.Param = para
        Return 0
    End Function


    'Set dot command record data
    '      sheet: pattern sheet
    '      row:   pattern command row number
    '      data:  dot record data
    '      referPt: reference point
    '      fidComp: sub/fiducial compensation data
    '      height:  height compensation value

    Public Function SetDotRecordData(ByVal sheet As OWC10.Worksheet, ByVal row As Integer, ByRef data As CIDSDot, _
                ByVal referPt() As Double, ByVal fidComp As FidCompData, ByVal height As Double) As Integer
        Dim enterTime As Long = DateTime.Now.Ticks
        Dim para As DispensePara
        Dim pt(3), comp(3), tmp(3) As Double
        Dim dispStr As String = "ON"
        Dim Ratio As Double = 1.0

        'Dim oneCell As OWC10.Range
        Dim i As Integer = 1
        Dim array As Object(,) ' array start at (1,1)
        Dim j = 1
        array = sheet.Range("A" & row, "AD" & row).Value
        'Console.WriteLine(array(1, 1))
        'For Each oneCell In sheetRange
        '    'modify the array if the value looks like a date, else skip it
        '    Console.WriteLine(oneCell.Value2)
        '    i += 1
        'Next oneCell
        pt(0) = array(j, gPos1XColumn) '(sheet.Cells(row, gPos1XColumn).value)
        pt(1) = array(j, gPos1YColumn) '(sheet.Cells(row, gPos1YColumn).value)
        pt(2) = array(j, gPos1ZColumn) '(sheet.Cells(row, gPos1ZColumn).value)
        'Console.WriteLine("#1 :" & ((DateTime.Now.Ticks - enterTime) / 10000).ToString())
        Dim testTime As Long = DateTime.Now.Ticks
        dispStr = array(j, gDispensColumn) 'sheet.Cells(row, gDispensColumn).value
        para.NeedleGap = array(j, gNeedleGapColumn) * Ratio '(sheet.Cells(row, gNeedleGapColumn).value) * Ratio
        para.Duration = array(j, gDurationColumn) 'sheet.Cells(row, gDurationColumn).value
        para.ApproachHeight = array(j, gApproachHtColumn) * Ratio ' (sheet.Cells(row, gApproachHtColumn).value) * Ratio
        para.RetractDelay = array(j, gRetractDelayColumn) 'sheet.Cells(row, gRetractDelayColumn).value
        para.RetractSpeed = array(j, gRetractSpeedColumn) * Ratio '(sheet.Cells(row, gRetractSpeedColumn).value) * Ratio
        para.RetractHeight = array(j, gRetractHtColumn) * Ratio ' (sheet.Cells(row, gRetractHtColumn).value) * Ratio
        para.ClearanceHeight = array(j, gClearanceHtColumn) * Ratio '(sheet.Cells(row, gClearanceHtColumn).value) * Ratio
        para.ArcRadius = array(j, gPos1ZColumn) * Ratio '(sheet.Cells(row, gArcRadiusColumn).value) * Ratio
        para.Needle = array(j, gNeedleColumn) 'sheet.Cells(row, gNeedleColumn).value
        'Console.WriteLine("#2 :" & ((DateTime.Now.Ticks - testTime) / 10000).ToString())
        testTime = DateTime.Now.Ticks
        FidandSubTransform(pt, referPt, fidComp, comp)
        'Console.WriteLine("#3 :" & ((DateTime.Now.Ticks - testTime) / 10000).ToString())
        testTime = DateTime.Now.Ticks
        If m_Optim = 0 Then
            If Not WorkAreaErrorCheckXY(comp) Then
                Dim sheetname As String = sheet.Name
                CompileErrorDisplay(sheetname, row, 0)
                Return -1
            End If
        End If
        'Console.WriteLine("#4 :" & ((DateTime.Now.Ticks - testTime) / 10000).ToString())
        testTime = DateTime.Now.Ticks
        data.CmdLineNo = row
        data.CmdType = "DOT"
        data.SheetName = sheet.Name
        data.PosX = comp(0)
        data.PosY = comp(1)
        data.PosZ = comp(2)
        data.HeightCompS = height
        'Console.WriteLine("#5 :" & ((DateTime.Now.Ticks - testTime) / 10000).ToString())
        dispStr = dispStr.ToUpper()
        If dispStr = "ON" Then
            para.DispenseOn = True
        Else
            para.DispenseOn = False
        End If

        data.Param = para
        Dim duration As Long = (DateTime.Now.Ticks - enterTime) / 10000 'miliseconds
        'Console.WriteLine("Set Dot Record Data = " & duration.ToString())
        Return 0
    End Function

    Public Function TSetDotRecordData(ByRef array As Object(,), ByVal sheet As OWC10.Worksheet, ByVal row As Integer, ByRef data As CIDSDot, _
                                     ByVal referPt() As Double, ByVal fidComp As FidCompData, ByVal height As Double) As Integer
        Dim enterTime As Long = DateTime.Now.Ticks
        Dim para As DispensePara
        Dim pt(3), comp(3), tmp(3) As Double
        Dim dispStr As String = "ON"
        Dim Ratio As Double = 1.0
        'Dim sheetRange As OWC10.Range
        ' sheetRange = sheet.Range("A" & row, "AD" & row)

        'Dim oneCell As OWC10.Range
        'Dim i As Integer = 1
        'Dim array As Object(,) ' array start at (1,1)
        'array = sheet.Range("A" & row, "AD" & row).Value
        'Console.WriteLine(array(1, 1))
        'For Each oneCell In sheetRange
        '    'modify the array if the value looks like a date, else skip it
        '    Console.WriteLine(oneCell.Value2)
        '    i += 1
        'Next oneCell
        pt(0) = array(row, gPos1XColumn) '(sheet.Cells(row, gPos1XColumn).value)
        pt(1) = array(row, gPos1YColumn) '(sheet.Cells(row, gPos1YColumn).value)
        pt(2) = array(row, gPos1ZColumn) '(sheet.Cells(row, gPos1ZColumn).value)
        'Console.WriteLine("#1 :" & ((DateTime.Now.Ticks - enterTime) / 10000).ToString())
        Dim testTime As Long = DateTime.Now.Ticks
        dispStr = array(row, gDispensColumn) 'sheet.Cells(row, gDispensColumn).value
        para.NeedleGap = array(row, gNeedleGapColumn) * Ratio '(sheet.Cells(row, gNeedleGapColumn).value) * Ratio
        para.Duration = array(row, gDurationColumn) 'sheet.Cells(row, gDurationColumn).value
        para.ApproachHeight = array(row, gApproachHtColumn) * Ratio ' (sheet.Cells(row, gApproachHtColumn).value) * Ratio
        para.RetractDelay = array(row, gRetractDelayColumn) 'sheet.Cells(row, gRetractDelayColumn).value
        para.RetractSpeed = array(row, gRetractSpeedColumn) * Ratio '(sheet.Cells(row, gRetractSpeedColumn).value) * Ratio
        para.RetractHeight = array(row, gRetractHtColumn) * Ratio ' (sheet.Cells(row, gRetractHtColumn).value) * Ratio
        para.ClearanceHeight = array(row, gClearanceHtColumn) * Ratio '(sheet.Cells(row, gClearanceHtColumn).value) * Ratio
        para.ArcRadius = array(row, gPos1ZColumn) * Ratio '(sheet.Cells(row, gArcRadiusColumn).value) * Ratio
        para.Needle = array(row, gNeedleColumn) 'sheet.Cells(row, gNeedleColumn).value
        'Console.WriteLine("#2 :" & ((DateTime.Now.Ticks - testTime) / 10000).ToString())
        testTime = DateTime.Now.Ticks
        FidandSubTransform(pt, referPt, fidComp, comp)
        'Console.WriteLine("#3 :" & ((DateTime.Now.Ticks - testTime) / 10000).ToString())
        testTime = DateTime.Now.Ticks
        If m_Optim = 0 Then
            If Not WorkAreaErrorCheckXY(comp) Then
                Dim sheetname As String = sheet.Name
                CompileErrorDisplay(sheetname, row, 0)
                Return -1
            End If
        End If
        'Console.WriteLine("#4 :" & ((DateTime.Now.Ticks - testTime) / 10000).ToString())
        testTime = DateTime.Now.Ticks
        data.CmdLineNo = row
        data.CmdType = "DOT"
        data.SheetName = sheet.Name
        data.PosX = comp(0)
        data.PosY = comp(1)
        data.PosZ = comp(2)
        data.HeightCompS = height
        'Console.WriteLine("#5 :" & ((DateTime.Now.Ticks - testTime) / 10000).ToString())
        dispStr = dispStr.ToUpper()
        If dispStr = "ON" Then
            para.DispenseOn = True
        Else
            para.DispenseOn = False
        End If

        data.Param = para
        Dim duration As Long = (DateTime.Now.Ticks - enterTime) / 10000 'miliseconds
        'Console.WriteLine("Set Dot Record Data = " & duration.ToString())
        Return 0
    End Function


    'Set link start command record data
    '      sheet: pattern sheet
    '      row:   pattern command row number
    '      data:  link start record data
    '

    Public Function SetLinkSRecordData(ByVal sheet As OWC10.Worksheet, ByVal row As Integer, ByRef data As CIDSLinkS) As Integer
        data.CmdLineNo = row
        data.CmdType = "LINKSTART"
        data.SheetName = sheet.Name
        Return 0
    End Function


    'Set link end command record data
    '      sheet: pattern sheet
    '      row:   pattern command row number
    '      data:  link end record data
    '

    Public Function SetLinkERecordData(ByVal sheet As OWC10.Worksheet, ByVal row As Integer, ByRef data As CIDSLinkE) As Integer
        data.CmdLineNo = row
        data.CmdType = "LINKEND"
        data.SheetName = sheet.Name
        Return 0
    End Function


    'Set line command record data
    '      sheet: pattern sheet
    '      row:   pattern command row number
    '      data:  line record data
    '      referPt: reference point
    '      fidComp: sub/fiducial compensation data
    '      height:  height compensation value

    Public Function SetLineRecordData(ByVal sheet As OWC10.Worksheet, ByVal row As Integer, ByRef data As CIDSLine, _
                                     ByVal referPt() As Double, ByVal fidComp As FidCompData, ByVal height As Double) As Integer
        Dim para As DispensePara
        Dim pt1(3), comp1(3) As Double
        Dim pt2(3), comp2(3) As Double
        data.CmdLineNo = row
        data.CmdType = "LINE"
        data.SheetName = sheet.Name

        Dim array As Object(,) ' array start at (1,1)
        Dim j = 1
        array = sheet.Range("A" & row, "AD" & row).Value

        pt1(0) = array(j, gPos1XColumn)
        pt1(1) = array(j, gPos1YColumn)
        pt1(2) = array(j, gPos1ZColumn)
        FidandSubTransform(pt1, referPt, fidComp, comp1)
        Dim sheetname As String = sheet.Name
        If Not WorkAreaErrorCheckXY(comp1) Then
            CompileErrorDisplay(sheetname, row, 0)
            Return -1
        End If

        pt2(0) = array(j, gPos2XColumn)
        pt2(1) = array(j, gPos2YColumn)
        pt2(2) = array(j, gPos2ZColumn)
        FidandSubTransform(pt2, referPt, fidComp, comp2)
        If Not WorkAreaErrorCheckXY(comp2) Then
            CompileErrorDisplay(sheetname, row, 0)
            Return -1
        End If

        data.PosX1 = comp1(0)
        data.PosY1 = comp1(1)
        data.PosZ1 = comp1(2)
        data.PosX2 = comp2(0)
        data.PosY2 = comp2(1)
        data.PosZ2 = comp2(2)
        data.HeightCompS = height
        Dim dispStr As String = array(j, gDispensColumn)
        dispStr = dispStr.ToUpper()
        If dispStr = "ON" Then
            para.DispenseOn = True
        Else
            para.DispenseOn = False
        End If

        para.NeedleGap = array(j, gNeedleGapColumn)
        para.TravelDelay = array(j, gTravelDelayColumn)
        para.TravelSpeed = array(j, gTravelSpeedColumn)
        para.DeTailDist = array(j, gDTaiDistColumn)
        para.ApproachHeight = array(j, gApproachHtColumn)
        para.RetractDelay = array(j, gRetractDelayColumn)
        para.RetractSpeed = array(j, gRetractSpeedColumn)
        para.RetractHeight = array(j, gRetractHtColumn)
        para.ClearanceHeight = array(j, gClearanceHtColumn)
        para.ArcRadius = array(j, gArcRadiusColumn)
        para.Needle = array(j, gNeedleColumn)
        data.Param = para
        Return 0

    End Function

    Public Function SetDotArrayRecordData(ByVal sheet As OWC10.Worksheet, ByVal row As Integer, ByRef data As CIDSDotArray, ByVal referPt() As Double, ByVal fidComp As FidCompData, ByVal height As Double) As Integer
        Dim para As DispensePara
        Dim pt1(3), comp1(3), tmp1(3) As Double
        Dim pt2(3), comp2(3), tmp2(3) As Double
        Dim pt3(3), comp3(3), tmp3(3) As Double

        Dim array As Object(,) ' array start at (1,1)
        Dim j = 1
        array = sheet.Range("A" & row, "AD" & row).Value

        pt1(0) = array(j, gPos1XColumn)
        pt1(1) = array(j, gPos1YColumn)
        pt1(2) = array(j, gPos1ZColumn)
        FidandSubTransform(pt1, referPt, fidComp, comp1)
        Dim sheetname As String = sheet.Name
        If Not WorkAreaErrorCheckXY(comp1) Then
            CompileErrorDisplay(sheetname, row, 0)
            Return -1
        End If

        pt2(0) = array(j, gPos2XColumn)
        pt2(1) = array(j, gPos2YColumn)
        pt2(2) = array(j, gPos2ZColumn)
        FidandSubTransform(pt2, referPt, fidComp, comp2)

        If Not WorkAreaErrorCheckXY(comp2) Then
            CompileErrorDisplay(sheetname, row, 0)
            Return -1
        End If

        pt3(0) = array(j, gPos3XColumn)
        pt3(1) = array(j, gPos3YColumn)
        pt3(2) = array(j, gPos3ZColumn)
        FidandSubTransform(pt3, referPt, fidComp, comp3)

        If Not WorkAreaErrorCheckXY(comp3) Then
            CompileErrorDisplay(sheetname, row, 0)
            Return -1
        End If

        data.PosX1 = comp1(0)
        data.PosY1 = comp1(1)
        data.PosZ1 = comp1(2)
        data.PosX2 = comp2(0)
        data.PosY2 = comp2(1)
        data.PosZ2 = comp2(2)
        data.PosX3 = comp3(0)
        data.PosY3 = comp3(1)
        data.PosZ3 = comp3(2)
        data.CmdLineNo = row
        data.CmdType = "DOTARRAY"
        data.SheetName = sheet.Name

        data.HeightCompS = height
        Dim dispStr As String = array(j, gDispensColumn)
        dispStr = dispStr.ToUpper()
        If dispStr = "ON" Then
            para.DispenseOn = True
        Else
            para.DispenseOn = False
        End If

        para.NeedleGap = array(j, gNeedleGapColumn)
        para.Duration = array(j, gDurationColumn)
        para.ApproachHeight = array(j, gApproachHtColumn)
        para.RetractDelay = array(j, gRetractDelayColumn)
        para.RetractSpeed = array(j, gRetractSpeedColumn)
        para.RetractHeight = array(j, gRetractHtColumn)
        para.ClearanceHeight = array(j, gClearanceHtColumn)
        para.ArcRadius = array(j, gArcRadiusColumn)
        para.Needle = array(j, gNeedleColumn)

        Dim RowColumnInformationString As String = array(j, gTravelSpeedColumn)
        Dim RowsAndColumns(1) As String
        RowsAndColumns = RowColumnInformationString.Split("x")
        para.Rows = CInt(RowsAndColumns(0))
        para.Columns = CInt(RowsAndColumns(1))
        data.Param = para

        Return 0
    End Function


    'Set arc command record data
    '      sheet: pattern sheet
    '      row:   pattern command row number
    '      data:  arc record data
    '      referPt: reference point
    '      fidComp: sub/fiducial compensation data
    '      height:  height compensation value

    Public Function SetArcRecordData(ByVal sheet As OWC10.Worksheet, ByVal row As Integer, ByRef data As CIDSArc, _
                                     ByVal referPt() As Double, ByVal fidComp As FidCompData, ByVal height As Double) As Integer
        Dim para As DispensePara
        Dim pt1(3), comp1(3), tmp1(3) As Double
        Dim pt2(3), comp2(3), tmp2(3) As Double
        Dim pt3(3), comp3(3), tmp3(3) As Double
        data.CmdLineNo = row
        data.CmdType = "ARC"
        data.SheetName = sheet.Name

        Dim array As Object(,) ' array start at (1,1)
        Dim j = 1
        array = sheet.Range("A" & row, "AD" & row).Value

        pt1(0) = array(j, gPos1XColumn)
        pt1(1) = array(j, gPos1YColumn)
        pt1(2) = array(j, gPos1ZColumn)
        FidandSubTransform(pt1, referPt, fidComp, comp1)
        Dim sheetname As String = sheet.Name
        If Not WorkAreaErrorCheckXY(comp1) Then
            CompileErrorDisplay(sheetname, row, 0)
            Return -1
        End If

        pt2(0) = array(j, gPos2XColumn)
        pt2(1) = array(j, gPos2YColumn)
        pt2(2) = array(j, gPos2ZColumn)
        FidandSubTransform(pt2, referPt, fidComp, comp2)

        If Not WorkAreaErrorCheckXY(comp2) Then
            CompileErrorDisplay(sheetname, row, 0)
            Return -1
        End If

        pt3(0) = array(j, gPos3XColumn)
        pt3(1) = array(j, gPos3YColumn)
        pt3(2) = array(j, gPos3ZColumn)
        FidandSubTransform(pt3, referPt, fidComp, comp3)

        If Not WorkAreaErrorCheckXY(comp3) Then
            CompileErrorDisplay(sheetname, row, 0)
            Return -1
        End If

        data.PosX1 = comp1(0)
        data.PosY1 = comp1(1)
        data.PosZ1 = comp1(2)
        data.PosX2 = comp2(0)
        data.PosY2 = comp2(1)
        data.PosZ2 = comp2(2)
        data.PosX3 = comp3(0)
        data.PosY3 = comp3(1)
        data.PosZ3 = comp3(2)

        data.HeightCompS = height
        Dim dispStr As String = array(j, gDispensColumn)
        dispStr = dispStr.ToUpper()
        If dispStr = "ON" Then
            para.DispenseOn = True
        Else
            para.DispenseOn = False
        End If

        para.NeedleGap = array(j, gNeedleGapColumn)
        para.TravelDelay = array(j, gTravelDelayColumn)
        para.TravelSpeed = array(j, gTravelSpeedColumn)
        para.DeTailDist = array(j, gDTaiDistColumn)
        para.ApproachHeight = array(j, gApproachHtColumn)
        para.RetractDelay = array(j, gRetractDelayColumn)
        para.RetractSpeed = array(j, gRetractSpeedColumn)
        para.RetractHeight = array(j, gRetractHtColumn)
        para.ClearanceHeight = array(j, gClearanceHtColumn)
        para.ArcRadius = array(j, gArcRadiusColumn)
        para.Needle = array(j, gNeedleColumn)
        data.Param = para
        Return 0
    End Function


    'Set circle command record data
    '      sheet: pattern sheet
    '      row:   pattern command row number
    '      data:  circle record data
    '      referPt: reference point
    '      fidComp: sub/fiducial compensation data
    '      height:  height compensation value

    Public Function SetCircleRecordData(ByVal sheet As OWC10.Worksheet, ByVal row As Integer, ByRef data As CIDSCircle, _
                                     ByVal referPt() As Double, ByVal fidComp As FidCompData, ByVal height As Double) As Integer
        Dim para As DispensePara
        Dim array_read, cell_read As Double
        Dim pt1(3), comp1(3), tmp1(3) As Double
        Dim pt2(3), comp2(3), tmp2(3) As Double
        Dim pt3(3), comp3(3), tmp3(3) As Double

        data.CmdLineNo = row
        data.CmdType = "CIRCLE"
        data.SheetName = sheet.Name

        Dim array As Object(,) ' array start at (1,1)
        Dim j = 1
        array = sheet.Range("A" & row, "AD" & row).Value

        Dim dispStr As String = array(j, gDispensColumn)
        pt1(0) = array(j, gPos1XColumn)
        pt1(1) = array(j, gPos1YColumn)
        pt1(2) = array(j, gPos1ZColumn)
        pt2(0) = array(j, gPos2XColumn)
        pt2(1) = array(j, gPos2YColumn)
        pt2(2) = array(j, gPos2ZColumn)
        pt3(0) = array(j, gPos3XColumn)
        pt3(1) = array(j, gPos3YColumn)
        pt3(2) = array(j, gPos3ZColumn)
        para.NeedleGap = array(j, gNeedleGapColumn)
        para.TravelDelay = array(j, gTravelDelayColumn)
        para.TravelSpeed = array(j, gTravelSpeedColumn)
        para.DeTailDist = array(j, gDTaiDistColumn)
        para.ApproachHeight = array(j, gApproachHtColumn)
        para.RetractDelay = array(j, gRetractDelayColumn)
        para.RetractSpeed = array(j, gRetractSpeedColumn)
        para.RetractHeight = array(j, gRetractHtColumn)
        para.ClearanceHeight = array(j, gClearanceHtColumn)
        para.ArcRadius = array(j, gArcRadiusColumn)
        para.Needle = array(j, gNeedleColumn)


        FidandSubTransform(pt1, referPt, fidComp, comp1)
        Dim sheetname As String = sheet.Name
        If Not WorkAreaErrorCheckXY(comp1) Then
            CompileErrorDisplay(sheetname, row, 0)
            Return -1
        End If

        FidandSubTransform(pt2, referPt, fidComp, comp2)

        If Not WorkAreaErrorCheckXY(comp2) Then
            CompileErrorDisplay(sheetname, row, 0)
            Return -1
        End If

        FidandSubTransform(pt3, referPt, fidComp, comp3)
        If Not WorkAreaErrorCheckXY(comp3) Then
            CompileErrorDisplay(sheetname, row, 0)
            Return -1
        End If

        data.PosX1 = comp1(0)
        data.PosY1 = comp1(1)
        data.PosZ1 = comp1(2)
        data.PosX2 = comp2(0)
        data.PosY2 = comp2(1)
        data.PosZ2 = comp2(2)
        data.PosX3 = comp3(0)
        data.PosY3 = comp3(1)
        data.PosZ3 = comp3(2)

        data.HeightCompS = height
        dispStr = dispStr.ToUpper()
        If dispStr = "ON" Then
            para.DispenseOn = True
        Else
            para.DispenseOn = False
        End If

        data.Param = para
        Return 0
    End Function



    'Set rectangle command record data
    '      sheet: pattern sheet
    '      row:   pattern command row number
    '      data:  rectangle record data
    '      referPt: reference point
    '      fidComp: sub/fiducial compensation data
    '      height:  height compensation value

    Public Function SetRectRecordData(ByVal sheet As OWC10.Worksheet, ByVal row As Integer, ByRef data As CIDSRectangle, _
                                     ByVal referPt() As Double, ByVal fidComp As FidCompData, ByVal height As Double) As Integer
        Dim para As DispensePara
        Dim pt1(3), comp1(3), tmp1(3) As Double
        Dim pt2(3), comp2(3), tmp2(3) As Double
        Dim pt3(3), comp3(3), tmp3(3) As Double
        data.CmdLineNo = row
        data.CmdType = "RECTANGLE"
        data.SheetName = sheet.Name

        Dim array As Object(,) ' array start at (1,1)
        Dim j = 1
        array = sheet.Range("A" & row, "AD" & row).Value

        pt1(0) = array(j, gPos1XColumn)
        pt1(1) = array(j, gPos1YColumn)
        pt1(2) = array(j, gPos1ZColumn)
        FidandSubTransform(pt1, referPt, fidComp, comp1)
        Dim sheetname As String = sheet.Name
        If Not WorkAreaErrorCheckXY(comp1) Then
            CompileErrorDisplay(sheetname, row, 0)
            Return -1
        End If

        pt2(0) = array(j, gPos2XColumn)
        pt2(1) = array(j, gPos2YColumn)
        pt2(2) = array(j, gPos2ZColumn)
        FidandSubTransform(pt2, referPt, fidComp, comp2)

        If Not WorkAreaErrorCheckXY(comp2) Then
            CompileErrorDisplay(sheetname, row, 0)
            Return -1
        End If

        pt3(0) = array(j, gPos3XColumn)
        pt3(1) = array(j, gPos3YColumn)
        pt3(2) = array(j, gPos3ZColumn)
        FidandSubTransform(pt3, referPt, fidComp, comp3)

        If Not WorkAreaErrorCheckXY(comp3) Then
            CompileErrorDisplay(sheetname, row, 0)
            Return -1
        End If

        data.PosX1 = comp1(0)
        data.PosY1 = comp1(1)
        data.PosZ1 = comp1(2)
        data.PosX2 = comp2(0)
        data.PosY2 = comp2(1)
        data.PosZ2 = comp2(2)
        data.PosX3 = comp3(0)
        data.PosY3 = comp3(1)
        data.PosZ3 = comp3(2)
        data.HeightCompS = height
        Dim dispStr As String = array(j, gDispensColumn)
        dispStr = dispStr.ToUpper()
        If dispStr = "ON" Then
            para.DispenseOn = True
        Else
            para.DispenseOn = False
        End If

        para.NeedleGap = array(j, gNeedleGapColumn)
        para.TravelDelay = array(j, gTravelDelayColumn)
        para.TravelSpeed = array(j, gTravelSpeedColumn)
        para.DeTailDist = array(j, gDTaiDistColumn)
        para.ApproachHeight = array(j, gApproachHtColumn)
        para.RetractDelay = array(j, gRetractDelayColumn)
        para.RetractSpeed = array(j, gRetractSpeedColumn)
        para.RetractHeight = array(j, gRetractHtColumn)
        para.ClearanceHeight = array(j, gClearanceHtColumn)
        para.ArcRadius = array(j, gArcRadiusColumn)
        para.Needle = array(j, gNeedleColumn)
        data.Param = para
        Return 0
    End Function


    'Set fill circle command record data
    '      sheet: pattern sheet
    '      row:   pattern command row number
    '      data:  fill circle record data
    '      referPt: reference point
    '      fidComp: sub/fiducial compensation data
    '      height:  height compensation value

    Public Function SetFillCircleRecordData(ByVal sheet As OWC10.Worksheet, ByVal row As Integer, ByRef data As CIDSFillCircle, _
                                      ByVal referPt() As Double, ByVal fidComp As FidCompData, ByVal height As Double) As Integer
        Dim para As DispensePara
        Dim pt1(3), comp1(3), tmp1(3) As Double
        Dim pt2(3), comp2(3), tmp2(3) As Double
        Dim pt3(3), comp3(3), tmp3(3) As Double
        data.CmdLineNo = row
        data.CmdType = "FILLCIRCLE"
        data.SheetName = sheet.Name

        Dim array As Object(,) ' array start at (1,1)
        Dim j = 1
        array = sheet.Range("A" & row, "AD" & row).Value

        pt1(0) = array(j, gPos1XColumn)
        pt1(1) = array(j, gPos1YColumn)
        pt1(2) = array(j, gPos1ZColumn)
        FidandSubTransform(pt1, referPt, fidComp, comp1)
        Dim sheetname As String = sheet.Name
        If Not WorkAreaErrorCheckXY(comp1) Then
            CompileErrorDisplay(sheetname, row, 0)
            Return -1
        End If

        pt2(0) = array(j, gPos2XColumn)
        pt2(1) = array(j, gPos2YColumn)
        pt2(2) = array(j, gPos2ZColumn)
        FidandSubTransform(pt2, referPt, fidComp, comp2)

        If Not WorkAreaErrorCheckXY(comp2) Then
            CompileErrorDisplay(sheetname, row, 0)
            Return -1
        End If


        pt3(0) = array(j, gPos3XColumn)
        pt3(1) = array(j, gPos3YColumn)
        pt3(2) = array(j, gPos3ZColumn)
        FidandSubTransform(pt3, referPt, fidComp, comp3)

        If Not WorkAreaErrorCheckXY(comp3) Then
            CompileErrorDisplay(sheetname, row, 0)
            Return -1
        End If

        data.PosX1 = comp1(0)
        data.PosY1 = comp1(1)
        data.PosZ1 = comp1(2)
        data.PosX2 = comp2(0)
        data.PosY2 = comp2(1)
        data.PosZ2 = comp2(2)
        data.PosX3 = comp3(0)
        data.PosY3 = comp3(1)
        data.PosZ3 = comp3(2)
        data.HeightCompS = height
        Dim dispStr As String = array(j, gDispensColumn)
        dispStr = dispStr.ToUpper()
        If dispStr = "ON" Then
            para.DispenseOn = True
        Else
            para.DispenseOn = False
        End If

        para.NeedleGap = array(j, gNeedleGapColumn)
        para.TravelDelay = array(j, gTravelDelayColumn)
        para.TravelSpeed = array(j, gTravelSpeedColumn)
        para.DeTailDist = array(j, gDTaiDistColumn)
        para.ApproachHeight = array(j, gApproachHtColumn)
        para.RetractDelay = array(j, gRetractDelayColumn)
        para.RetractSpeed = array(j, gRetractSpeedColumn)
        para.RetractHeight = array(j, gRetractHtColumn)
        para.ClearanceHeight = array(j, gClearanceHtColumn)
        para.ArcRadius = array(j, gArcRadiusColumn)
        para.Needle = array(j, gNeedleColumn)
        para.Pitch = array(j, gPitchColumn)
        para.NoOfRun = array(j, gNoOfRunColumn)
        para.SprialIn = array(j, gSprialColumn)
        para.FillHeight = array(j, gFillHeightColumn)
        data.Param = para
        Return 0
    End Function


    'Set fill rectangle command record data
    '      sheet: pattern sheet
    '      row:   pattern command row number
    '      data:  fill rectangle record data
    '      referPt: reference point
    '      fidComp: sub/fiducial compensation data
    '      height:  height compensation value

    Public Function SetFillRectRecordData(ByVal sheet As OWC10.Worksheet, ByVal row As Integer, ByRef data As CIDSFillRectangle, _
                                     ByVal referPt() As Double, ByVal fidComp As FidCompData, ByVal height As Double) As Integer
        Dim para As DispensePara
        Dim pt1(3), comp1(3), tmp1(3) As Double
        Dim pt2(3), comp2(3), tmp2(3) As Double
        Dim pt3(3), comp3(3), tmp3(3) As Double
        data.CmdLineNo = row
        data.CmdType = "FILLRECTANGLE"
        data.SheetName = sheet.Name

        Dim array As Object(,) ' array start at (1,1)
        Dim j = 1
        array = sheet.Range("A" & row, "AD" & row).Value

        pt1(0) = array(j, gPos1XColumn)
        pt1(1) = array(j, gPos1YColumn)
        pt1(2) = array(j, gPos1ZColumn)
        FidandSubTransform(pt1, referPt, fidComp, comp1)
        Dim sheetname As String = sheet.Name
        If Not WorkAreaErrorCheckXY(comp1) Then
            CompileErrorDisplay(sheetname, row, 0)
            Return -1
        End If

        pt2(0) = array(j, gPos2XColumn)
        pt2(1) = array(j, gPos2YColumn)
        pt2(2) = array(j, gPos2ZColumn)
        FidandSubTransform(pt2, referPt, fidComp, comp2)

        If Not WorkAreaErrorCheckXY(comp2) Then
            CompileErrorDisplay(sheetname, row, 0)
            Return -1
        End If

        pt3(0) = array(j, gPos3XColumn)
        pt3(1) = array(j, gPos3YColumn)
        pt3(2) = array(j, gPos3ZColumn)
        FidandSubTransform(pt3, referPt, fidComp, comp3)

        If Not WorkAreaErrorCheckXY(comp3) Then
            CompileErrorDisplay(sheetname, row, 0)
            Return -1
        End If

        data.PosX1 = comp1(0)
        data.PosY1 = comp1(1)
        data.PosZ1 = comp1(2)
        data.PosX2 = comp2(0)
        data.PosY2 = comp2(1)
        data.PosZ2 = comp2(2)
        data.PosX3 = comp3(0)
        data.PosY3 = comp3(1)
        data.PosZ3 = comp3(2)
        data.HeightCompS = height
        Dim dispStr As String = array(j, gDispensColumn)
        dispStr = dispStr.ToUpper()
        If dispStr = "ON" Then
            para.DispenseOn = True
        Else
            para.DispenseOn = False
        End If

        para.NeedleGap = array(j, gNeedleGapColumn)
        para.TravelDelay = array(j, gTravelDelayColumn)
        para.TravelSpeed = array(j, gTravelSpeedColumn)
        para.DeTailDist = array(j, gDTaiDistColumn)
        para.ApproachHeight = array(j, gApproachHtColumn)
        para.RetractDelay = array(j, gRetractDelayColumn)
        para.RetractSpeed = array(j, gRetractSpeedColumn)
        para.RetractHeight = array(j, gRetractHtColumn)
        para.ClearanceHeight = array(j, gClearanceHtColumn)
        para.ArcRadius = array(j, gArcRadiusColumn)
        para.Needle = array(j, gNeedleColumn)
        para.Pitch = array(j, gPitchColumn)
        para.NoOfRun = array(j, gNoOfRunColumn)
        para.SprialIn = array(j, gSprialColumn)
        para.FillHeight = array(j, gFillHeightColumn)
        data.Param = para
        Return 0
    End Function

    'Get array first element's position
    '      sheetlist: sheet name list
    '      arraysheetname:   array sheet name
    '      referencePt: reference point
    '      firstPos:  first element position

    Public Function GetArrayFirstPos(ByVal sheetlist As ArrayList, ByVal arraysheetname As String, ByVal referencePt() As Double, ByRef firstPos() As Double) As Integer
        Dim arraysheet As OWC10.Worksheet
        Dim numofsheet As Integer = sheetlist.Count
        Dim isExist As Boolean = False
        firstPos(0) = 0.0
        firstPos(1) = 0.0
        firstPos(2) = 0.0
        Dim i As Integer
        'KR FILE NAME BUG
        'Dim shortened_arraysheetname = arraysheetname.Remove(31, arraysheetname.Length - 31)
        For i = 0 To numofsheet - 1 'check array sheet name whether exists
            If arraysheetname = sheetlist(i) Then
                arraysheet = m_Sheet.Workbooks.Item(1).Worksheets(arraysheetname)
                isExist = True
                Exit For
            End If
        Next i

        If Not isExist Then
            Return -1
        End If

        If arraysheet Is Nothing Then
            Return -1
        End If

        Dim Rows, Colums As Integer
        Dim ret As Integer = GetSheetUsedRowsColums(arraysheet, Rows, Colums)
        If Rows < 1 Or ret < 0 Then
            Return -1  'empty sheet or error
        End If

        Dim x, y, z As Double
        Dim Type As String
        'set first element position
        Dim array As Object(,) ' array start at (1,1)
        Dim j = 1
        array = arraysheet.Range("A1", "AD" & Rows).Value

        For i = 1 To Rows
            'Type = arraysheet.Cells(i, gCommandNameColumn).Value
            Type = array(i, gCommandNameColumn)
            If Type <> Nothing Then
                Type = Type.Trim(" ")
                Type = Type.ToUpper
                Select Case Type
                    Case "DOT", "LINE", "RECTANGLE", "FILLRECTANGLE"
                        x = array(i, gPos1XColumn)
                        y = array(i, gPos1YColumn)
                        z = array(i, gPos1ZColumn)
                        firstPos(0) = x + referencePt(0)
                        firstPos(1) = y + referencePt(1)
                        firstPos(2) = z + referencePt(2)
                        Return 0
                    Case "ARC", "CIRCLE", "FILLCIRCLE"
                        x = array(i, gPos1XColumn)
                        y = array(i, gPos1YColumn)
                        z = array(i, gPos1ZColumn)
                        firstPos(0) = x + referencePt(0)
                        firstPos(1) = y + referencePt(1)
                        firstPos(2) = z + referencePt(2)
                        Return 0
                    Case "SUBPATTERN", "CHIPEDGE"
                        x = array(i, gPos1XColumn)
                        y = array(i, gPos1YColumn)
                        z = array(i, gPos1ZColumn)
                        firstPos(0) = x + referencePt(0)
                        firstPos(1) = y + referencePt(1)
                        firstPos(2) = z + referencePt(2)
                        Return 0
                End Select
            End If
            'SJ add for GUI freezing
            'TraceDoEvents()
        Next i
        Return 0
    End Function



    'Get pattern sheet first element's position
    '      subname: sub pattern name
    '      sheetlist:  whole pattern sheet name list
    '      sheet:  sub pattern sheet
    '      rows:   sub pattern sheet total rows
    '      referencePt: reference point
    '      firstPos: first element's position in sub pattern
    '    

    Public Function GetFirstElementPos(ByVal subname As String, ByVal sheetlist As ArrayList, ByVal sheet As OWC10.Worksheet, ByVal rows As Integer, ByVal referencePt() As Double, ByRef firstPos() As Double) As Integer
        Dim type As String
        Dim i As Integer
        Dim x, y, z As Double

        firstPos(0) = 0.0
        firstPos(1) = 0.0
        firstPos(2) = 0.0

        Dim array As Object(,) ' array start at (1,1)
        Dim j = 1
        array = sheet.Range("A1", "AD" & rows).Value

        For i = 1 To rows
            type = array(i, gCommandNameColumn)
            If type <> Nothing Then
                type = type.Trim(" ")
                type = type.ToUpper
                Select Case type
                    Case "REFERENCE"
                        referencePt(0) = array(i, gPos1XColumn)
                        referencePt(1) = array(i, gPos1YColumn)
                        referencePt(2) = array(i, gPos1ZColumn)
                    Case "DOT", "LINE", "RECTANGLE", "FILLRECTANGLE", "DOTARRAY"
                        x = array(i, gPos1XColumn)
                        y = array(i, gPos1YColumn)
                        z = array(i, gPos1ZColumn)
                        firstPos(0) = x + referencePt(0)
                        firstPos(1) = y + referencePt(1)
                        firstPos(2) = z + referencePt(2)
                        Return 0
                    Case "ARC", "CIRCLE", "FILLCIRCLE"
                        x = array(i, gPos1XColumn)
                        y = array(i, gPos1YColumn)
                        z = array(i, gPos1ZColumn)
                        firstPos(0) = x + referencePt(0)
                        firstPos(1) = y + referencePt(1)
                        firstPos(2) = z + referencePt(2)
                        Return 0
                    Case "SUBPATTERN", "CHIPEDGE"
                        x = array(i, gPos1XColumn)
                        y = array(i, gPos1YColumn)
                        z = array(i, gPos1ZColumn)
                        firstPos(0) = x + referencePt(0)
                        firstPos(1) = y + referencePt(1)
                        firstPos(2) = z + referencePt(2)
                        Return 0
                    Case "ARRAY"
                        Dim arrayname As String = array(i, gArraynameColumn)
                        If arrayname = Nothing Then Return -1
                        Dim arraysheetname As String = subname + "." + arrayname
                        Return GetArrayFirstPos(sheetlist, arraysheetname, referencePt, firstPos)
                End Select
            End If
            'SJ add for GUI freezing
            'TraceDoEvents()
        Next i
        Return -1
    End Function


    'Remove element objects from pattern command list
    '

    Public Function RemoveElements(ByRef list As ArrayList, ByVal startingElem As Object) As Integer
        Dim count As Integer = list.Count
        If count < 1 Then Return -1
        Dim index As Integer = list.LastIndexOf(startingElem)
        If index < 0 Then Return -1
        list.RemoveRange(index, count - index)
        Return 0
    End Function



    'Set second level sub pattern command record data
    '      sheetlist: whole pattern sheet name list
    '      sheet:     sub sheet
    '      row:   sub pattern command row number
    '      list:  pattern command list
    '      referencePt: reference point
    '      heightcomps: height compensation
    '      compData: sub/fiducial compensation data
    '      

    Public Function SetSubSubPatternRecordData(ByVal sheetlist As ArrayList, ByVal sheet As OWC10.Worksheet, ByVal row As Integer, ByRef list As ArrayList, ByVal referencePt() As Double, ByVal heightcomps As Double, _
                                               ByRef compData As FidCompData) As Integer
        Dim i, Rows, Colums As Integer
        Dim fiducialpt(3) As Double
        Dim offset(3) As Double
        Dim compangle As Double = 0.0
        Dim heightcomp As Double = heightcomps
        Dim subsheet As OWC10.Worksheet

        Dim array As Object(,) ' array start at (1,1)
        Dim j = 1
        array = sheet.Range("A" & row, "AD" & row).Value

        Dim subfilename As String = array(j, gSubnameColumn)
        Dim sheetname As String
        sheetname = sheet.Name
        If subfilename = Nothing Then
            CompileErrorDisplay(sheetname, row, 1)
            Return -1
        End If

        Dim numofsheet As Integer = sheetlist.Count
        If numofsheet < 2 Then
            CompileErrorDisplay(sheetname, row, 2)
            Return -1
        End If

        Dim intPos As Integer
        'kr added to change gFidFileName for subpatterns
        intPos = subfilename.LastIndexOfAny(".")
        gFidFileName = subfilename.Remove(intPos, 4)

        intPos = subfilename.LastIndexOfAny("\")
        intPos += 1
        subfilename = subfilename.Substring(intPos, (Len(subfilename) - intPos)) 'get sub pattern file name without directory tree


        Dim isExist As Boolean = False
        For i = 0 To numofsheet - 1  'check whether sub pattern sheet exists
            If subfilename = sheetlist(i) Then
                subsheet = m_Sheet.Workbooks.Item(1).Worksheets(subfilename)
                isExist = True
                Exit For
            End If
        Next i

        If Not isExist Then
            CompileErrorDisplay(sheetname, row, 2)
            Return -1
        End If
        If subsheet Is Nothing Then
            Return -1
        End If

        Dim ret As Integer = GetSheetUsedRowsColums(subsheet, Rows, Colums) 'get sub pattern sheet's used rows and columns
        sheetname = subsheet.Name
        If Rows < 1 Or ret < 0 Then
            CompileErrorDisplay(sheetname, 0, 3)
            Return -1  'empty sheet or error
        End If

        Dim subInsPt(3), subInsAngle As Double  'get sub insert point and angle
        subInsPt(0) = array(j, gPos1XColumn)
        subInsPt(1) = array(j, gPos1YColumn)
        subInsPt(2) = array(j, gPos1ZColumn)
        subInsAngle = array(j, gRtAngleColumn) * Math.PI / 180.0

        Dim firstPos(3) As Double ' get sub's first point location
        If GetFirstElementPos(subfilename, sheetlist, subsheet, Rows, referencePt, firstPos) < 0 Then
            CompileErrorDisplay(sheetname, 0, 4)
            Return -1
        End If

        compData.SetSub2(subInsPt, firstPos, subInsAngle) 'set sub transform parameters

        'set breakpoint for skip pattern when chipedge checking fails
        Dim subRecS As New CIDSSub
        subRecS.CmdType = "SUBS"
        subRecS.subName = subfilename
        subRecS.level = 2
        DebugAddList(list, subRecS)  'add breakpoint to pattern list
        '

        array = subsheet.Range("A1" & row, "AD" & Rows).Value

        Dim rtn As Integer = 0
        For i = 1 To Rows  'handling all pattern commands 
            Dim type As String = array(i, gCommandNameColumn)
            If type <> Nothing Then
                type = type.Trim(" ")
                type = type.ToUpper
                Select Case type
                    Case "FIDUCIAL"     'Check fiducial point and get fiducial compensation 
                        If m_Optim = 0 Then
                            rtn = CheckFiducialPt(subsheet, i, referencePt, compData, fiducialpt, offset, compangle)
                            If rtn < 0 Then
                                If IDS.Data.Hardware.SPC.FiducialFailedAction = False Then    'auto skip
                                    compData.ClearFid()
                                    compData.ClearSub2()
                                    If ProductionMode() And Production.ContinuousMode.Checked Then
                                        Form_Service.LogEventInSPCReport("Fiducial Failed")
                                        Form_Service.LogEventInSPCReport("Board Partial Failure")
                                        If ShouldLog() Then Form_Service.LogEventInSPCReport("Resume")
                                    End If
                                    Return 0
                                Else     'manu / user decision

                                    If ProductionMode() And Production.ContinuousMode.Checked Then
                                        Form_Service.LogEventInSPCReport("Fiducial Failed")
                                        Form_Service.DisplayErrorMessage("Fiducial Failed")
                                    Else
                                        Form_Service.DisplayErrorMessage("Fiducial Failed")
                                    End If

                                    WaitUntilErrorMessagesCleared()

                                    If Form_Service.FiducialSkipped() Then
                                        compData.ClearFid()
                                        compData.ClearSub2()
                                        If ShouldLog() Then Form_Service.LogEventInSPCReport("Board Partial Failure")
                                        If ShouldLog() Then Form_Service.LogEventInSPCReport("Resume")
                                        Return 0
                                    Else
                                        If ShouldLog() Then Form_Service.LogEventInSPCReport("Board Total Failure")
                                        Return 99
                                    End If
                                End If
                            ElseIf rtn > 0 Then
                                Return rtn
                            Else
                                compData.SetFid(fiducialpt, offset, compangle)
                            End If
                        End If
                    Case "HEIGHT"  'height detection
                        If m_Optim = 0 Then
                            Dim heightc As Double
                            If Laser Is Nothing Then
                                MessageBox.Show("This software not support height compensation with laser! Process abort!")
                                Return -1
                            End If
                            rtn = GetHeightCompensation(subsheet, i, referencePt, compData, heightc)
                            If rtn < 0 Then   'check fail
                                If IDS.Data.Hardware.SPC.HeightFailedAction = False Then    'auto skip
                                    compData.ClearFid()
                                    compData.ClearSub2()
                                    If ProductionMode() Then
                                        Form_Service.LogEventInSPCReport("Height Point Failed")
                                        Form_Service.LogEventInSPCReport("Board Partial Failure")
                                        If ShouldLog() Then Form_Service.LogEventInSPCReport("Resume")
                                    End If
                                    'delete elements from list
                                    Return RemoveElements(list, subRecS)
                                Else     'manu /user decision



                                    If ProductionMode() Then
                                        Form_Service.LogEventInSPCReport("Height Point Failed")
                                        Form_Service.DisplayErrorMessage("Height Point Failed")
                                    Else
                                        Form_Service.DisplayErrorMessage("Height Point Failed")
                                    End If

                                    WaitUntilErrorMessagesCleared()

                                    If Form_Service.HeightSkipped Then
                                        'skip pattern
                                        compData.ClearFid()
                                        compData.ClearSub2()
                                        If ShouldLog() Then Form_Service.LogEventInSPCReport("Board Partial Failure")
                                        If ShouldLog() Then Form_Service.LogEventInSPCReport("Resume")
                                        'delete elements from list
                                        Return RemoveElements(list, subRecS)
                                    Else
                                        'stop
                                        If ShouldLog() Then Form_Service.LogEventInSPCReport("Board Total Failure")
                                        Return 99
                                    End If
                                End If
                            ElseIf rtn > 0 Then  'user stop execution
                                Return rtn
                            Else
                                heightcomp = heightc
                            End If
                        End If
                    Case "REJECT" 'detect reject mark
                        If m_Optim = 0 Then
                            rtn = CheckRejectMark(subsheet, i, referencePt, compData)
                            If rtn < 0 Then   'matched
                                'auto skip
                                compData.ClearFid() 'skip pattern
                                compData.ClearSub2()
                                If ShouldLog() Then Form_Service.LogEventInSPCReport("Board Partial Failure")
                                If ShouldLog() Then Form_Service.LogEventInSPCReport("Resume")
                                'delete elements from list
                                Return RemoveElements(list, subRecS)
                            ElseIf rtn > 0 Then  'user stop execution
                                Return rtn
                            End If
                        End If
                    Case "REFERENCE"  'get reference point
                        GetReferencePtData(subsheet, i, referencePt)
                    Case "WAIT" 'set wait element data
                        If m_Optim = 0 Then
                            Dim waitData As New CIDSWait
                            If SetWaitRecordData(subsheet, i, waitData, referencePt, compData, heightcomp) < 0 Then
                                Return -1
                            End If
                            DebugAddList(list, waitData)
                        End If
                    Case "SETIO"  'set setio element data
                        If m_Optim = 0 Then
                            Dim setIOData As New CIDSSetIO
                            SetSetIORecordData(subsheet, i, setIOData)
                            DebugAddList(list, setIOData)
                        End If
                    Case "GETIO"  'set getio element data
                        If m_Optim = 0 Then
                            Dim getIOData As New CIDSGetIO
                            SetGetIORecordData(subsheet, i, getIOData)
                            DebugAddList(list, getIOData)
                        End If
                    Case "VOLUMECALIBRATION"
                        If m_Optim = 0 Then
                            Dim volcalData As New CIDSVolumeCalibration
                            SetVolCalRecordData(sheet, i, volcalData)
                            DebugAddList(list, volcalData)
                        End If
                    Case "CLEAN" 'set clean element data
                        If m_Optim = 0 Then
                            Dim cleanData As New CIDSClean
                            SetCleanRecordData(sheet, i, cleanData)
                            DebugAddList(list, cleanData)
                        End If
                    Case "DOTARRAY"
                        Dim dotArrayData As New CIDSDotArray
                        If SetDotArrayRecordData(subsheet, i, dotArrayData, referencePt, compData, heightcomp) < 0 Then
                            Return -1
                        End If
                        DebugAddList(list, dotArrayData)
                    Case "PURGE"  'set purge element data
                        If m_Optim = 0 Then
                            Dim purgeData As New CIDSPurge
                            SetPurgeRecordData(sheet, i, purgeData)
                            DebugAddList(list, purgeData)
                        End If
                    Case "SUBPATTERN"
                        Return -1     ' error: only 2 levels sub call
                    Case "ARRAY"  'set array elements data
                        Dim skipFlag As Integer = 0
                        rtn = SetArrayRecordData(3, subfilename, skipFlag, sheetlist, subsheet, i, list, referencePt, compData, heightcomp)
                        If skipFlag = 1 Then  'skip pattern
                            compData.ClearFid()
                            compData.ClearSub2()
                            If ShouldLog() Then Form_Service.LogEventInSPCReport("Board Partial Failure")
                            If ShouldLog() Then Form_Service.LogEventInSPCReport("Resume")
                            'delete elements from list
                            Return RemoveElements(list, subRecS)
                        ElseIf rtn < 0 Then
                            Return -1
                        ElseIf rtn > 0 Then
                            Return rtn
                        End If
                    Case "CHIPEDGE" 'set chipedge element data
                        If m_Optim = 0 Then
                            Dim chipData As New CIDSChipEdge
                            rtn = ChecknSetChipedgeRecData(subsheet, i, chipData, referencePt, compData, heightcomp)
                            If rtn < 0 Then   'check fail
                                If IDS.Data.Hardware.SPC.ChipFailedAction = False Then    'auto skip
                                    If ProductionMode() Then
                                        Form_Service.LogEventInSPCReport("Chip Finding Failed")
                                        Form_Service.LogEventInSPCReport("Board Partial Failure")
                                        If ShouldLog() Then Form_Service.LogEventInSPCReport("Resume")
                                    End If
                                Else
                                    If ProductionMode() Then
                                        Form_Service.LogEventInSPCReport("Chip Finding Failed")
                                        Form_Service.DisplayErrorMessage("Chip Finding Failed")
                                    Else
                                        Form_Service.DisplayErrorMessage("Chip Finding Failed")
                                    End If

                                    WaitUntilErrorMessagesCleared()

                                    If ShouldLog() Then Form_Service.LogEventInSPCReport("Board Partial Failure")
                                    If Form_Service.ChipEdgeSkipped Then
                                        If ShouldLog() Then Form_Service.LogEventInSPCReport("Resume")
                                        'skip chip
                                    Else
                                        'skip pattern 
                                        compData.ClearFid()
                                        compData.ClearSub2()
                                        'delete elements from list
                                        Return RemoveElements(list, subRecS)
                                    End If
                                End If
                            ElseIf rtn > 0 Then
                                Return rtn
                            Else
                                DebugAddList(list, chipData)
                            End If
                        End If
                    Case "QC"  'set QC element data
                        If m_Optim = 0 Then
                            Dim QCData As New CIDSQC
                            If SetQCRecData(subsheet, i, QCData, referencePt, compData) < 0 Then
                                Return -1
                            End If
                            DebugAddList(list, QCData)
                        End If
                    Case "MOVE"  'set move element data
                        If m_Optim = 0 Then
                            Dim moveData As New CIDSMove
                            If SetMoveRecordData(subsheet, i, moveData, referencePt, compData, heightcomp) < 0 Then
                                Return -1
                            End If
                            DebugAddList(list, moveData)
                        End If
                    Case "DOT"  'set dot element data
                        Dim dotData As New CIDSDot
                        If SetDotRecordData(subsheet, i, dotData, referencePt, compData, heightcomp) < 0 Then
                            Return -1
                        End If
                        DebugAddList(list, dotData)
                    Case "LINE" 'set line element data

                        If m_Optim = 0 Then
                            Dim lineData As New CIDSLine
                            If SetLineRecordData(subsheet, i, lineData, referencePt, compData, heightcomp) < 0 Then
                                Return -1
                            End If
                            DebugAddList(list, lineData)
                        End If
                    Case "ARC"  'set arc element data
                        If m_Optim = 0 Then
                            Dim arcData As New CIDSArc
                            If SetArcRecordData(subsheet, i, arcData, referencePt, compData, heightcomp) < 0 Then
                                Return -1
                            End If
                            DebugAddList(list, arcData)
                        End If
                    Case "CIRCLE"  'set circle element data
                        If m_Optim = 0 Then
                            Dim circleData As New CIDSCircle
                            If SetCircleRecordData(subsheet, i, circleData, referencePt, compData, heightcomp) < 0 Then
                                Return -1
                            End If
                            DebugAddList(list, circleData)
                        End If
                    Case "LINKSTART"  'set link start element data

                        If m_Optim = 0 Then
                            Dim linkSData As New CIDSLinkS
                            SetLinkSRecordData(subsheet, i, linkSData)
                            DebugAddList(list, linkSData)
                        End If
                    Case "LINKEND"   'set link end element data
                        If m_Optim = 0 Then
                            Dim linkEData As New CIDSLinkE
                            SetLinkERecordData(subsheet, i, linkEData)
                            DebugAddList(list, linkEData)
                        End If
                    Case "FILLCIRCLE"  'set fill circle element data
                        If m_Optim = 0 Then
                            Dim FillCircleData As New CIDSFillCircle
                            If SetFillCircleRecordData(subsheet, i, FillCircleData, referencePt, compData, heightcomp) < 0 Then
                                Return -1
                            End If
                            DebugAddList(list, FillCircleData)
                        End If
                    Case "RECTANGLE"  'set rectangle element data
                        If m_Optim = 0 Then
                            Dim rectData As New CIDSRectangle
                            If SetRectRecordData(subsheet, i, rectData, referencePt, compData, heightcomp) < 0 Then
                                Return -1
                            End If
                            DebugAddList(list, rectData)
                        End If
                    Case "FILLRECTANGLE"  'set fill rectangle element data
                        If m_Optim = 0 Then
                            Dim rectData As New CIDSFillRectangle
                            If SetFillRectRecordData(subsheet, i, rectData, referencePt, compData, heightcomp) < 0 Then
                                Return -1
                            End If
                            DebugAddList(list, rectData)
                        End If
                End Select
            End If
            'SJ add for GUI freezing
            'TraceDoEvents()
        Next i
        compData.ClearFid()  'initialize fiducial compensation data
        compData.ClearSub2() 'initailize sub transform data
        'set sub pattern end record for qc check fail skipping pattern
        Dim subRecE As New CIDSSub
        subRecE.CmdType = "SUBE"
        subRecE.subName = subfilename
        subRecS.level = 2
        DebugAddList(list, subRecE)
        'list.Remove(subRecS)
        Return 0
    End Function


    'Set first level sub pattern command record data
    '      sheetlist: whole pattern sheet name list
    '      sheet:     sub sheet
    '      row:   sub pattern command row number
    '      list:  pattern command list
    '      referencePt: reference point
    '      heightcomps: height compensation
    '      compData: sub/fiducial compensation data

    Public Function SetSubPatternRecordData(ByVal sheetlist As ArrayList, ByVal sheet As OWC10.Worksheet, ByVal row As Integer, ByRef list As ArrayList, ByVal referencePt() As Double, ByVal heightcomps As Double, ByRef compData As FidCompData) As Integer
        Dim i, Rows, Colums As Integer
        Dim fiducialpt(3) As Double
        Dim offset(3) As Double
        Dim compangle As Double = 0.0
        Dim heightcomp As Double = heightcomps
        Dim subsheet As OWC10.Worksheet
        Dim subfilename As String = sheet.Cells(row, gSubnameColumn).value
        Dim sheetname As String
        sheetname = sheet.Name
        Dim one = sheet.Parent
        Dim two = sheet.Names
        Dim three = sheet.Rows
        Dim four = sheet.Columns
        Dim five = sheet.Cells
        If subfilename = Nothing Then
            CompileErrorDisplay(sheetname, row, 1)
            Return -1
        End If
        Dim numofsheet As Integer = sheetlist.Count
        If numofsheet < 2 Then
            CompileErrorDisplay(sheetname, row, 2)
            Return -1
        End If
        Dim intPos As Integer

        'kr added to change gFidFileName for subpatterns
        intPos = subfilename.LastIndexOfAny(".")
        gFidFileName = subfilename.Remove(intPos, 4)

        intPos = subfilename.LastIndexOfAny("\")
        intPos += 1
        subfilename = subfilename.Substring(intPos, subfilename.Length - intPos) 'get sub pattern file name without directory tree

        Dim isExist As Boolean = False
        For i = 0 To numofsheet - 1  'check whether sub pattern sheet exists
            If subfilename = sheetlist(i) Then
                subsheet = m_Sheet.Workbooks.Item(1).Worksheets(subfilename)
                isExist = True
                Exit For
            End If
        Next
        If Not isExist Then
            CompileErrorDisplay(sheetname, row, 2)
            Return -1
        End If

        If subsheet Is Nothing Then
            Return -1
        End If
        Dim ret As Integer = GetSheetUsedRowsColums(subsheet, Rows, Colums)
        sheetname = subsheet.Name
        If Rows < 1 Or ret < 0 Then
            CompileErrorDisplay(sheetname, 0, 3)
            Return -1  'empty sheet or error
        End If

        Dim array As Object(,) ' array start at (1,1)
        Dim j = 1
        array = sheet.Range("A" & row, "AD" & row).Value

        Dim subInsPt(3), subInsAngle As Double
        subInsPt(0) = array(j, gPos1XColumn)
        subInsPt(1) = array(j, gPos1YColumn)
        subInsPt(2) = array(j, gPos1ZColumn)
        subInsAngle = array(j, gRtAngleColumn) * Math.PI / 180.0
        Dim firstPos(3) As Double
        If GetFirstElementPos(subfilename, sheetlist, subsheet, Rows, referencePt, firstPos) < 0 Then
            CompileErrorDisplay(sheetname, 0, 4)
            Return -1
        End If
        compData.SetSub1(subInsPt, firstPos, subInsAngle)
        'set temp breakpoint for skip pattern when QC chipedge checking fails
        Dim subRecS As New CIDSSub
        subRecS.CmdType = "SUBS"
        subRecS.subName = subfilename
        subRecS.level = 1
        DebugAddList(list, subRecS)
        '
        Dim type As String
        Dim rtn As Integer = 0
        array = subsheet.Range("A1", "AD" & Rows).Value
        For i = 1 To Rows
            'type = subsheet.Cells(i, gCommandNameColumn).Value
            type = array(i, gCommandNameColumn)
            If type <> Nothing Then
                type = type.Trim(" ")
                type = type.ToUpper
                Select Case type
                    Case "FIDUCIAL"   'Check fiducial point and get fiducial compensation 
                        If m_Optim = 0 Then
                            rtn = CheckFiducialPt(subsheet, i, referencePt, compData, fiducialpt, offset, compangle)
                            If rtn < 0 Then
                                If IDS.Data.Hardware.SPC.FiducialFailedAction = False Then    'auto skip
                                    compData.ClearParentFid()
                                    compData.ClearSub1()
                                    If ProductionMode() Then
                                        Form_Service.LogEventInSPCReport("Fiducial Failed")
                                        Form_Service.LogEventInSPCReport("Board Partial Failure")
                                        If ShouldLog() Then Form_Service.LogEventInSPCReport("Resume")
                                    End If
                                    Return 0
                                Else     'manu


                                    If ProductionMode() Then
                                        Form_Service.LogEventInSPCReport("Fiducial Failed")
                                        Form_Service.DisplayErrorMessage("Fiducial Failed")
                                    Else
                                        Form_Service.DisplayErrorMessage("Fiducial Failed")
                                    End If

                                    WaitUntilErrorMessagesCleared()

                                    If Form_Service.FiducialSkipped Then
                                        'skip
                                        compData.ClearParentFid()
                                        compData.ClearSub1()
                                        If ShouldLog() Then Form_Service.LogEventInSPCReport("Board Partial Failure")
                                        If ShouldLog() Then Form_Service.LogEventInSPCReport("Resume")
                                        Return 0
                                    Else
                                        'stop
                                        If ShouldLog() Then Form_Service.LogEventInSPCReport("Board Total Failure")
                                        Return 99
                                    End If
                                End If
                            ElseIf rtn > 0 Then
                                Return rtn
                            Else
                                compData.SetParentFid(fiducialpt, offset, compangle) 'set compensation data
                            End If
                        Else
                            'DebugAddList(list, )
                        End If
                    Case "HEIGHT"   'height detection
                        If m_Optim = 0 Then
                            Dim heightc As Double
                            If Laser Is Nothing Then
                                MessageBox.Show("This software not support height compensation with laser! Process abort!")
                                Return -1
                            End If
                            rtn = GetHeightCompensation(subsheet, i, referencePt, compData, heightc)
                            If rtn < 0 Then   'check fail
                                If IDS.Data.Hardware.SPC.HeightFailedAction = False Then    'auto skip pattern
                                    compData.ClearParentFid()
                                    compData.ClearSub1()
                                    If ProductionMode() Then
                                        Form_Service.LogEventInSPCReport("Height Point Failed")
                                        Form_Service.LogEventInSPCReport("Board Partial Failure")
                                        If ShouldLog() Then Form_Service.LogEventInSPCReport("Resume")
                                    End If
                                    'delete elements from list
                                    Return RemoveElements(list, subRecS)
                                Else     'manu



                                    If ProductionMode() Then
                                        Form_Service.LogEventInSPCReport("Height Point Failed")
                                        Form_Service.DisplayErrorMessage("Height Point Failed")
                                    Else
                                        Form_Service.DisplayErrorMessage("Height Point Failed")
                                    End If

                                    WaitUntilErrorMessagesCleared()

                                    If Form_Service.HeightSkipped Then
                                        'skip pattern
                                        compData.ClearParentFid()
                                        compData.ClearSub1()
                                        If ShouldLog() Then Form_Service.LogEventInSPCReport("Board Partial Failure")
                                        If ShouldLog() Then Form_Service.LogEventInSPCReport("Resume")
                                        'delete elements from list
                                        Return RemoveElements(list, subRecS)
                                    Else
                                        'stop
                                        If ShouldLog() Then Form_Service.LogEventInSPCReport("Board Total Failure")
                                        Return 99
                                    End If

                                End If
                            ElseIf rtn > 0 Then
                                Return rtn
                            Else
                                heightcomp = heightc
                            End If

                        End If
                    Case "REJECT"  'detect reject mark
                        If m_Optim = 0 Then
                            rtn = CheckRejectMark(subsheet, i, referencePt, compData)
                            If rtn < 0 Then   'matched
                                'auto skip pattern
                                compData.ClearParentFid()
                                compData.ClearSub1()
                                If ShouldLog() Then Form_Service.LogEventInSPCReport("Board Partial Failure")
                                If ShouldLog() Then Form_Service.LogEventInSPCReport("Resume")
                                'delete elements from list
                                Return RemoveElements(list, subRecS)
                            ElseIf rtn > 0 Then
                                Return rtn
                            End If
                        End If
                    Case "REFERENCE"    'get reference point
                        GetReferencePtData(subsheet, i, referencePt)
                    Case "WAIT"         'set wait element data
                        If m_Optim = 0 Then
                            Dim waitData As New CIDSWait
                            If SetWaitRecordData(subsheet, i, waitData, referencePt, compData, heightcomp) < 0 Then
                                Return -1
                            End If
                            DebugAddList(list, waitData)
                        End If
                    Case "DOTARRAY"
                        Dim dotArrayData As New CIDSDotArray
                        If SetDotArrayRecordData(subsheet, i, dotArrayData, referencePt, compData, heightcomp) < 0 Then
                            Return -1
                        End If
                        DebugAddList(list, dotArrayData)
                    Case "SETIO"         'set setio element data
                        If m_Optim = 0 Then
                            Dim setIOData As New CIDSSetIO
                            SetSetIORecordData(subsheet, i, setIOData)
                            DebugAddList(list, setIOData)
                        End If
                    Case "GETIO"         'set getio element data
                        If m_Optim = 0 Then
                            Dim getIOData As New CIDSGetIO
                            SetGetIORecordData(subsheet, i, getIOData)
                            DebugAddList(list, getIOData)
                        End If
                    Case "VOLUMECALIBRATION"          'set clean element data
                        If m_Optim = 0 Then
                            Dim volcalData As New CIDSVolumeCalibration
                            SetVolCalRecordData(subsheet, i, volcalData)
                            DebugAddList(list, volcalData)
                        End If
                    Case "CLEAN"          'set clean element data
                        If m_Optim = 0 Then
                            Dim cleanData As New CIDSClean
                            SetCleanRecordData(subsheet, i, cleanData)
                            DebugAddList(list, cleanData)
                        End If
                    Case "PURGE"          'set clean element data 
                        If m_Optim = 0 Then
                            Dim purgeData As New CIDSPurge
                            SetPurgeRecordData(subsheet, i, purgeData)
                            DebugAddList(list, purgeData)
                        End If
                    Case "SUBPATTERN"     'set sub pattern elements data
                        rtn = SetSubSubPatternRecordData(sheetlist, subsheet, i, list, referencePt, heightcomp, compData)
                        If rtn < 0 Then
                            Return -1
                        ElseIf rtn > 0 Then
                            Return rtn
                        End If
                    Case "ARRAY"           'set array element data
                        Dim skipFlag As Integer = 0
                        rtn = SetArrayRecordData(2, subfilename, skipFlag, sheetlist, subsheet, i, list, referencePt, compData, heightcomp)
                        If skipFlag = 1 Then   'skipping 
                            compData.ClearParentFid()
                            compData.ClearSub1()
                            If ShouldLog() Then Form_Service.LogEventInSPCReport("Board Partial Failure")
                            If ShouldLog() Then Form_Service.LogEventInSPCReport("Resume")
                            'delete elements from list
                            Return RemoveElements(list, subRecS)
                        ElseIf rtn < 0 Then
                            Return -1
                        ElseIf rtn > 0 Then
                            Return rtn
                        End If
                    Case "CHIPEDGE"          'set chipedge element data
                        If m_Optim = 0 Then
                            Dim chipData As New CIDSChipEdge
                            rtn = ChecknSetChipedgeRecData(subsheet, i, chipData, referencePt, compData, heightcomp)
                            'MessageBox.Show("Step ChipEdge")
                            If rtn < 0 Then   'check fail
                                If IDS.Data.Hardware.SPC.ChipFailedAction = False Then    'auto skip
                                    If ProductionMode() Then
                                        Form_Service.LogEventInSPCReport("Chip Finding Failed")
                                        Form_Service.LogEventInSPCReport("Board Partial Failure")
                                        If ShouldLog() Then Form_Service.LogEventInSPCReport("Resume")
                                    End If
                                Else     'manu

                                    If ProductionMode() Then
                                        Form_Service.LogEventInSPCReport("Chip Finding Failed")
                                        Form_Service.DisplayErrorMessage("Chip Finding Failed")
                                    Else
                                        Form_Service.DisplayErrorMessage("Chip Finding Failed")
                                    End If

                                    WaitUntilErrorMessagesCleared()

                                    If ShouldLog() Then Form_Service.LogEventInSPCReport("Board Partial Failure")
                                    If Form_Service.ChipEdgeSkipped Then
                                        'skip chip
                                        If ShouldLog() Then Form_Service.LogEventInSPCReport("Resume")
                                    Else
                                        'skip pattern 
                                        compData.ClearParentFid()
                                        compData.ClearSub1()
                                        'delete elements from list
                                        'Return RemoveElements(list, subRecS)
                                        RemoveElements(list, subRecS)
                                        Return rtn
                                    End If
                                End If
                            ElseIf rtn > 0 Then  'user stop execution
                                Return rtn
                            Else
                                DebugAddList(list, chipData)
                            End If
                        End If
                    Case "QC"            'set QC element data
                        If m_Optim = 0 Then
                            Dim QCData As New CIDSQC
                            If SetQCRecData(subsheet, i, QCData, referencePt, compData) < 0 Then
                                Return -1
                            End If
                            DebugAddList(list, QCData)
                        End If
                    Case "MOVE"           'set move element data
                        If m_Optim = 0 Then
                            Dim moveData As New CIDSMove
                            If SetMoveRecordData(subsheet, i, moveData, referencePt, compData, heightcomp) < 0 Then
                                Return -1
                            End If
                            DebugAddList(list, moveData)
                        End If
                    Case "DOT"            'set dot element data
                        Dim dotData As New CIDSDot
                        If SetDotRecordData(subsheet, i, dotData, referencePt, compData, heightcomp) < 0 Then
                            Return -1
                        End If
                        DebugAddList(list, dotData)
                    Case "LINE"           'set line element data
                        If m_Optim = 0 Then
                            Dim lineData As New CIDSLine
                            If SetLineRecordData(subsheet, i, lineData, referencePt, compData, heightcomp) < 0 Then
                                Return -1
                            End If
                            DebugAddList(list, lineData)
                        End If
                    Case "ARC"            'set arc element data
                        If m_Optim = 0 Then
                            Dim arcData As New CIDSArc
                            If SetArcRecordData(subsheet, i, arcData, referencePt, compData, heightcomp) < 0 Then
                                Return -1
                            End If
                            DebugAddList(list, arcData)
                        End If
                    Case "CIRCLE"         'set circle element data
                        If m_Optim = 0 Then
                            Dim circleData As New CIDSCircle
                            If SetCircleRecordData(subsheet, i, circleData, referencePt, compData, heightcomp) < 0 Then
                                Return -1
                            End If
                            DebugAddList(list, circleData)
                        End If
                    Case "LINKSTART"       'set linkstart element data
                        If m_Optim = 0 Then
                            Dim linkSData As New CIDSLinkS
                            SetLinkSRecordData(subsheet, i, linkSData)
                            DebugAddList(list, linkSData)
                        End If
                    Case "LINKEND"          'set linkend element data
                        If m_Optim = 0 Then
                            Dim linkEData As New CIDSLinkE
                            SetLinkERecordData(subsheet, i, linkEData)
                            DebugAddList(list, linkEData)
                        End If
                    Case "FILLCIRCLE"       'set fill circle element data
                        If m_Optim = 0 Then
                            Dim FillCircleData As New CIDSFillCircle
                            If SetFillCircleRecordData(subsheet, i, FillCircleData, referencePt, compData, heightcomp) < 0 Then
                                Return -1
                            End If
                            DebugAddList(list, FillCircleData)
                        End If
                    Case "RECTANGLE"        'set rectangle element data
                        If m_Optim = 0 Then
                            Dim rectData As New CIDSRectangle
                            SetRectRecordData(subsheet, i, rectData, referencePt, compData, heightcomp)
                            DebugAddList(list, rectData)
                        End If
                    Case "FILLRECTANGLE"    'set fill rectangle element data
                        If m_Optim = 0 Then
                            Dim rectData As New CIDSFillRectangle
                            If SetFillRectRecordData(subsheet, i, rectData, referencePt, compData, heightcomp) < 0 Then
                                Return -1
                            End If
                            DebugAddList(list, rectData)
                        End If
                End Select
            End If
            'SJ add for GUI freezing
            'TraceDoEvents()
        Next i
        compData.ClearParentFid()
        compData.ClearSub1()
        Dim subRecE As New CIDSSub
        subRecE.CmdType = "SUBE"
        subRecE.subName = subfilename
        subRecE.level = 1
        DebugAddList(list, subRecE)
        'list.Remove(subRecS)
        Return 0
    End Function


    'Set array command record data
    '      level: 1-main, 2,3-sub
    '      parentName:  parent sheet name
    '      skipFlag:  1-skip, 0-none skip
    '      sheetlist: whole pattern sheet name list
    '      sheet:     sub sheet
    '      row:   sub pattern command row number
    '      list:  pattern command list
    '      referencePt: reference point
    '      compData: sub/fiducial compensation data
    '      heightcomp: height compensation
    '      

    Public Function SetArrayRecordData(ByVal level As Integer, ByVal parentName As String, ByRef skipFlag As Integer, ByVal sheetlist As ArrayList, ByVal sheet As OWC10.Worksheet, ByVal row As Integer, ByRef list As ArrayList, _
                                       ByVal referencePt() As Double, ByRef compData As FidCompData, ByVal heightComp As Double) As Integer
        Dim i, j, Rows, Colums As Integer
        Dim arraysheet As OWC10.Worksheet
        skipFlag = 0
        Dim arrayfilename As String = sheet.Cells(row, gArraynameColumn).value
        Dim sheetname As String
        sheetname = sheet.Name
        If arrayfilename = Nothing Then
            CompileErrorDisplay(sheetname, row, 1)
            Return -1
        End If
        Dim numofsheet As Integer = sheetlist.Count
        If numofsheet < 2 Then
            CompileErrorDisplay(sheetname, row, 2)
            Return -1
        End If

        If level = 1 Then 'in Main

        Else 'in sub
            arrayfilename = parentName + "." + arrayfilename
        End If

        Dim isExist As Boolean = False
        For i = 0 To numofsheet - 1
            If arrayfilename = sheetlist(i) Then
                arraysheet = m_Sheet.Workbooks.Item(1).Worksheets(arrayfilename)
                isExist = True
                Exit For
            End If
        Next
        If Not isExist Then
            CompileErrorDisplay(sheetname, row, 2)
            Return -1
        End If
        If arraysheet Is Nothing Then
            Return -1
        End If
        Dim ret As Integer = GetSheetUsedRowsColums(arraysheet, Rows, Colums)
        sheetname = arraysheet.Name
        If Rows < 1 Or ret < 0 Then
            CompileErrorDisplay(sheetname, 0, 3)
            Return -1  'empty sheet or error
        End If
        Dim enterTime As Long = DateTime.Now.Ticks
        Dim type As String
        Dim rtn As Integer = 0
        'array should be of one type only
        Dim array As Object(,) ' array start at (1,1)
        array = arraysheet.Range("A1", "AD" & Rows).Value
        For i = 1 To Rows
            type = array(i, gCommandNameColumn)
            If type <> Nothing Then
                type = type.Trim(" ")
                type = type.ToUpper
                Select Case type
                    Case "DOT" 'set dot record
                        Dim dotData As New CIDSDot
                        'If SetDotRecordData(arraysheet, i, dotData, referencePt, compData, heightComp) < 0 Then
                        '    Return -1
                        'End If
                        If TSetDotRecordData(array, arraysheet, i, dotData, referencePt, compData, heightComp) < 0 Then
                            Return -1
                        End If
                        DebugAddList(list, dotData)
                    Case "SUBPATTERN"  'call sub
                        If (level = 1) Then
                            rtn = SetSubPatternRecordData(sheetlist, arraysheet, i, list, referencePt, heightComp, compData)
                        ElseIf level = 2 Then
                            rtn = SetSubSubPatternRecordData(sheetlist, arraysheet, i, list, referencePt, heightComp, compData)
                        Else
                            CompileErrorDisplay(sheetname, 0, 9)
                            rtn = -1 '3 levels only
                        End If

                        If rtn < 0 Then
                            Return -1
                        ElseIf rtn > 0 Then
                            Return rtn
                        End If
                    Case "LINE"
                        If m_Optim = 0 Then
                            Dim lineData As New CIDSLine
                            If SetLineRecordData(arraysheet, i, lineData, referencePt, compData, heightComp) < 0 Then
                                Return -1
                            End If
                            DebugAddList(list, lineData)
                        End If
                    Case "ARC"
                        If m_Optim = 0 Then
                            Dim arcData As New CIDSArc
                            If SetArcRecordData(arraysheet, i, arcData, referencePt, compData, heightComp) < 0 Then
                                Return -1
                            End If
                            DebugAddList(list, arcData)
                        End If
                    Case "CIRCLE"
                        If m_Optim = 0 Then
                            Dim circleData As New CIDSCircle
                            If SetCircleRecordData(arraysheet, i, circleData, referencePt, compData, heightComp) < 0 Then
                                Return -1
                            End If
                            DebugAddList(list, circleData)
                        End If
                    Case "FILLCIRCLE"
                        If m_Optim = 0 Then
                            Dim FillCircleData As New CIDSFillCircle
                            If SetFillCircleRecordData(arraysheet, i, FillCircleData, referencePt, compData, heightComp) < 0 Then
                                Return -1
                            End If
                            DebugAddList(list, FillCircleData)
                        End If
                    Case "RECTANGLE"
                        If m_Optim = 0 Then
                            Dim rectData As New CIDSRectangle
                            If SetRectRecordData(arraysheet, i, rectData, referencePt, compData, heightComp) < 0 Then
                                Return -1
                            End If
                            DebugAddList(list, rectData)
                        End If
                    Case "FILLRECTANGLE"
                        If m_Optim = 0 Then
                            Dim fillrectData As New CIDSFillRectangle
                            If SetFillRectRecordData(arraysheet, i, fillrectData, referencePt, compData, heightComp) < 0 Then
                                Return -1
                            End If
                            DebugAddList(list, fillrectData)
                        End If
                    Case "CHIPEDGE"
                        If m_Optim = 0 Then
                            Dim chipData As New CIDSChipEdge
                            rtn = ChecknSetChipedgeRecData(arraysheet, i, chipData, referencePt, compData, heightComp)
                            If rtn < 0 Then   'check fail
                                If IDS.Data.Hardware.SPC.ChipFailedAction = False Then    'auto skip
                                    If ProductionMode() Then
                                        Form_Service.LogEventInSPCReport("Chip Finding Failed")
                                        Form_Service.LogEventInSPCReport("Board Partial Failure")
                                        If ShouldLog() Then Form_Service.LogEventInSPCReport("Resume")
                                    End If
                                Else     'manu


                                    If ProductionMode() Then
                                        Form_Service.LogEventInSPCReport("Chip Finding Failed")
                                        Form_Service.DisplayErrorMessage("Chip Finding Failed")
                                    Else
                                        Form_Service.DisplayErrorMessage("Chip Finding Failed")
                                    End If

                                    WaitUntilErrorMessagesCleared()

                                    If ShouldLog() Then Form_Service.LogEventInSPCReport("Board Partial Failure")
                                    If Form_Service.ChipEdgeSkipped Then
                                        If ShouldLog() Then Form_Service.LogEventInSPCReport("Resume")
                                        'skip chip
                                    Else
                                        'skip pattern 
                                        skipFlag = 1
                                        Return 0
                                    End If
                                End If
                            ElseIf rtn > 0 Then
                                Return rtn
                            Else
                                DebugAddList(list, chipData)
                            End If
                        End If
                End Select
            End If
            'SJ add for GUI freezing
            'TraceDoEvents()
        Next i
        Dim Duration As Long = (DateTime.Now.Ticks - enterTime) / 10000
        'Console.WriteLine("Total time(ms) set array record data: " & Duration.ToString())

        Return 0
    End Function

    'Set main pattern command record data
    '      sheetlist: whole pattern sheet name list
    '      list:  pattern command list
    Public Function SetPatternList(ByVal sheetlist As ArrayList, ByRef list As ArrayList, ByRef GQCParam As DLL_Export_Device_Vision.QC.QCParam) As Integer
        Dim I, J As Integer
        Dim compData As FidCompData
        Dim referencePt(3) As Double
        Dim heightcomp As Double = 0.0
        Dim fiducialpt(3) As Double
        Dim offset(3) As Double
        Dim compangle As Double = 0.0
        compData.ClearAll()

        referencePt(0) = 0.0
        referencePt(1) = 0.0
        referencePt(2) = 0.0
        fiducialpt(0) = 0.0
        fiducialpt(1) = 0.0
        fiducialpt(2) = 0.0
        offset(0) = 0.0
        offset(1) = 0.0
        offset(2) = 0.0

        Dim Rows, Colums As Integer
        If list.Count > 0 Then
            list.Clear()
        End If

        If m_Sheet Is Nothing Then
            Return -1
        End If

        Dim sheetname As String = "Main"
        Dim sheet As OWC10.Worksheet = m_Sheet.Workbooks.Item(1).Worksheets(sheetname)
        If sheet Is Nothing Then
            Return -1
        End If
        Dim ret As Integer = GetSheetUsedRowsColums(sheet, Rows, Colums)
        If Rows < 1 Or ret < 0 Then
            Return -1  'empty sheet or error
        End If

        Dim rtn As Integer = 0


        '   Xue Wen                                                                                     '
        '   Change reading method.                                                                      '
        '   Note: 1.) if user doesn't teach the program from "row 1", system can't read out all data.   '
        '         2.) Test more.                                                                        '

        Dim countUp As Integer = 1
        Dim startCount As Boolean = False
        Dim array As Object(,) ' array start at (1,1)
        array = sheet.Range("A1", "AD" & Rows).Value
        I = 1
        Do Until (countUp > Rows)
            If (I > Rows) Then
                If (countUp = 1) Then
                    Return -1
                End If
            Else
                If Not m_Optim = 1 Then
                    If CheckButtonState() = -1 Then
                        Return 100
                    End If
                End If
            End If

            gFidFileName = Programming.gPatternFileName 'for fiducial. added by kr

            Dim type As String = array(I, gCommandNameColumn)
            Dim skipRetry As Boolean = True
            If type <> Nothing Then
                type = type.Trim(" ")
                type = type.ToUpper
                Select Case type
                    Case "FIDUCIAL"  'Check fiducial point and get fiducial compensation 
                        If m_Optim = 0 Then
                            LabelMessage("Checking fiducial")
                            Do
                                rtn = CheckFiducialPt(sheet, I, referencePt, compData, fiducialpt, offset, compangle)
                                If rtn < 0 Then
                                    Dim fm As InfoForm = New InfoForm(True)
                                    fm.SetTitle("Fiducial")
                                    fm.SetCancelButtonText("Abort")
                                    fm.Size = New Size(584, 180)
                                    fm.SetOKButtonText("Retry")
                                    fm.AddNewLine("Fiducial Failed!")
                                    fm.AddNewLine("Would you like to retry?")
                                    If fm.ShowDialog() = DialogResult.OK Then
                                        skipRetry = False
                                    Else
                                        skipRetry = True
                                    End If
                                Else
                                    skipRetry = True
                                End If
                            Loop Until skipRetry

                            If rtn < 0 Then   'check fail
                                LabelMessage("Unable to find fiducial")
                                If ProductionMode() Then
                                    If ShouldLog() Then Form_Service.LogEventInSPCReport("Fiducial Failed")
                                    If ShouldLog() Then Form_Service.LogEventInSPCReport("Board Total Failure")
                                End If
                                Return 99
                            ElseIf rtn > 0 Then
                                LabelMessage("Unable perform fiducial finding operation")
                                Return rtn
                            Else
                                compData.SetGparentFid(fiducialpt, offset, compangle)
                                LabelMessage("Found fiducial")
                            End If
                        Else
                            Dim FidData As New CIDSFiducial
                            SetFiducialRecData(sheet, I, FidData)
                            DebugAddList(list, FidData)
                        End If

                    Case "HEIGHT"  'height detection
                        If m_Optim = 0 Then
                            Dim heightc As Double
                            LabelMessage("Height compensation in progress")
                            'OnLaser()
                            If Laser Is Nothing Then
                                MessageBox.Show("This software not support height compensation with laser! Process abort!")
                                Return -1
                            End If
                            rtn = GetHeightCompensation(sheet, I, referencePt, compData, heightc)
                            'OffLaser()
                            If rtn < 0 Then   'check fail
                                LabelMessage("Height compensation failed")
                                If ShouldLog() Then Form_Service.LogEventInSPCReport("Board Total Failure")
                                Return 99
                            ElseIf rtn > 0 Then
                                LabelMessage("Unable to perform height compensation")
                                Return rtn
                            Else
                                If heightcomp > 10 Then
                                    ''MyMsgBox("ahhh!")
                                    LabelMessage("Height compensation failed, value > 10")
                                    Return 101
                                End If
                                Dim h As Double = (CInt(heightc * 1000)) / 1000
                                LabelMessage("Height compensation done, different = " & h)
                                heightcomp = heightc
                            End If
                        Else
                            Dim htdata As New CIDSHeight
                            SetHeightRecData(sheet, I, htdata)
                            DebugAddList(list, htdata)
                        End If

                    Case "REJECT"  'detect reject mark
                        If m_Optim = 0 Then
                            rtn = CheckRejectMark(sheet, I, referencePt, compData)
                            If rtn < 0 Then   'matched
                                'auto skip same as stopping execution 
                                If ShouldLog() Then Form_Service.LogEventInSPCReport("Board Total Failure")
                                Return 98
                            ElseIf rtn > 0 Then  'user stops execution
                                Return rtn
                            End If
                        End If
                    Case "REFERENCE"
                        GetReferencePtData(sheet, I, referencePt)
                    Case "WAIT"
                        If m_Optim = 0 Then
                            Dim waitData As New CIDSWait
                            If SetWaitRecordData(sheet, I, waitData, referencePt, compData, heightcomp) < 0 Then
                                Return -1
                            End If
                            DebugAddList(list, waitData)
                        End If
                    Case "SETIO"
                        If m_Optim = 0 Then
                            Dim setIOData As New CIDSSetIO
                            '   If IO number is not between 32 and 47, system will not run anything.  
                            If (SetSetIORecordData(sheet, I, setIOData) < 0) Then
                                Return -1
                            End If
                            DebugAddList(list, setIOData)
                        End If
                    Case "GETIO"
                        If m_Optim = 0 Then
                            Dim getIOData As New CIDSGetIO
                            '   If IO number is not between 32 and 47, system will not run anything.  
                            If (SetGetIORecordData(sheet, I, getIOData) < 0) Then
                                Return -1
                            End If
                            DebugAddList(list, getIOData)
                        End If
                    Case "GLOBALQC"
                        Programming.GlobalQCEnabled = True
                        Me.SetGlobalQCData(sheet, I, GQCParam)
                    Case "VOLUMECALIBRATION"
                        If m_Optim = 0 Then
                            Dim volcalData As New CIDSVolumeCalibration
                            SetVolCalRecordData(sheet, I, volcalData)
                            DebugAddList(list, volcalData)
                        End If
                    Case "CLEAN"
                        If m_Optim = 0 Then
                            Dim cleanData As New CIDSClean
                            SetCleanRecordData(sheet, I, cleanData)
                            DebugAddList(list, cleanData)
                        End If
                    Case "PURGE"
                        If m_Optim = 0 Then
                            Dim purgeData As New CIDSPurge
                            SetPurgeRecordData(sheet, I, purgeData)
                            DebugAddList(list, purgeData)
                        End If
                    Case "SUBPATTERN"
                        rtn = SetSubPatternRecordData(sheetlist, sheet, I, list, referencePt, heightcomp, compData)
                        If rtn < 0 Then
                            Return -1
                        ElseIf rtn > 0 Then
                            Return rtn
                        End If
                    Case "ARRAY"
                        Dim skipFlag As Integer = 0
                        rtn = SetArrayRecordData(1, "", skipFlag, sheetlist, sheet, I, list, referencePt, compData, heightcomp)
                        If skipFlag = 1 Then
                            If ShouldLog() Then Form_Service.LogEventInSPCReport("Board Total Failure")
                            Return 99
                        ElseIf rtn < 0 Then
                            Return -1
                        ElseIf rtn > 0 Then
                            Return rtn
                        End If
                    Case "CHIPEDGE"
                        If m_Optim = 0 Then
                            Dim chipData As New CIDSChipEdge
                            LabelMessage("Performing Chip Edges detection")
                            Do
                                rtn = ChecknSetChipedgeRecData(sheet, I, chipData, referencePt, compData, heightcomp)
                                If rtn < 0 Then   'check fail
                                    LabelMessage("Unable to detect Chip Edges")
                                    If IDS.Data.Hardware.SPC.ChipFailedAction = False Then    'auto skip
                                        If ProductionMode() Then
                                            Form_Service.LogEventInSPCReport("Chip Finding Failed")
                                            Form_Service.LogEventInSPCReport("Board Partial Failure")
                                            If ShouldLog() Then Form_Service.LogEventInSPCReport("Resume")
                                        End If
                                    Else     'manu

                                        'If ProductionMode() Then

                                        'Form_Service.LogEventInSPCReport("Chip Finding Failed")
                                        'End If
                                        'Form_Service.DisplayErrorMessage("Chip Finding Failed")
                                        Dim fm As InfoForm = New InfoForm(True)
                                        fm.SetTitle("Chip Edge")
                                        fm.ShowIgnoreButton()
                                        fm.SetCancelButtonText("Abort")
                                        fm.SetOKButtonText("Retry")
                                        fm.AddNewLine("Chip Finding Failed")
                                        fm.AddNewLine("Would you like to retry?")
                                        Dim result As DialogResult = fm.ShowDialog()
                                        If result = DialogResult.OK Then
                                            skipRetry = False
                                        ElseIf result = DialogResult.Cancel Then
                                            skipRetry = True
                                            If ProductionMode() Then
                                                Form_Service.LogEventInSPCReport("Chip Finding Failed")
                                            End If
                                            Return rtn
                                        ElseIf result = DialogResult.Ignore Then
                                            skipRetry = True
                                            If ProductionMode() Then
                                                Form_Service.LogEventInSPCReport("Chip Finding Failed")
                                            End If
                                        End If
                                        'Else
                                        'Form_Service.DisplayErrorMessage("Chip Finding Failed")
                                        'End If
                                    End If
                                ElseIf rtn > 0 Then
                                    Return rtn
                                Else
                                    skipRetry = True
                                    DebugAddList(list, chipData)
                                End If
                            Loop Until skipRetry
                        End If

                    Case "QC"
                        If m_Optim = 0 Then
                            Dim QCData As New CIDSQC
                            If SetQCRecData(sheet, I, QCData, referencePt, compData) < 0 Then
                                Return -1
                            End If
                            DebugAddList(list, QCData)
                        End If
                    Case "MOVE"
                        If m_Optim = 0 Then
                            Dim moveData As New CIDSMove
                            If SetMoveRecordData(sheet, I, moveData, referencePt, compData, heightcomp) < 0 Then
                                Return -1
                            End If
                            DebugAddList(list, moveData)
                        End If
                    Case "DOTARRAY"
                        Dim dotArrayData As New CIDSDotArray
                        If SetDotArrayRecordData(sheet, I, dotArrayData, referencePt, compData, heightcomp) < 0 Then
                            Return -1
                        End If
                        DebugAddList(list, dotArrayData)
                    Case "DOT"
                        Dim dotData As New CIDSDot
                        If SetDotRecordData(sheet, I, dotData, referencePt, compData, heightcomp) < 0 Then
                            Return -1
                        End If
                        DebugAddList(list, dotData)
                    Case "LINE"
                        If m_Optim = 0 Then
                            Dim lineData As New CIDSLine
                            If SetLineRecordData(sheet, I, lineData, referencePt, compData, heightcomp) < 0 Then
                                Return -1
                            End If
                            DebugAddList(list, lineData)
                        End If
                    Case "ARC"
                        If m_Optim = 0 Then
                            Dim arcData As New CIDSArc
                            If SetArcRecordData(sheet, I, arcData, referencePt, compData, heightcomp) < 0 Then
                                Return -1
                            End If
                            DebugAddList(list, arcData)
                        End If
                    Case "CIRCLE"
                        If m_Optim = 0 Then
                            Dim circleData As New CIDSCircle
                            If SetCircleRecordData(sheet, I, circleData, referencePt, compData, heightcomp) < 0 Then
                                Return -1
                            End If
                            DebugAddList(list, circleData)
                        End If
                    Case "LINKSTART"
                        If m_Optim = 0 Then
                            Dim linkSData As New CIDSLinkS
                            SetLinkSRecordData(sheet, I, linkSData)
                            DebugAddList(list, linkSData)
                        End If
                    Case "LINKEND"
                        If m_Optim = 0 Then
                            Dim linkEData As New CIDSLinkE
                            SetLinkERecordData(sheet, I, linkEData)
                            DebugAddList(list, linkEData)
                        End If
                    Case "FILLCIRCLE"
                        If m_Optim = 0 Then
                            Dim FillCircleData As New CIDSFillCircle
                            If SetFillCircleRecordData(sheet, I, FillCircleData, referencePt, compData, heightcomp) < 0 Then
                                Return -1
                            End If
                            DebugAddList(list, FillCircleData)
                        End If
                    Case "RECTANGLE"
                        If m_Optim = 0 Then
                            Dim rectData As New CIDSRectangle
                            If SetRectRecordData(sheet, I, rectData, referencePt, compData, heightcomp) < 0 Then
                                Return -1
                            End If
                            DebugAddList(list, rectData)
                        End If
                    Case "FILLRECTANGLE"
                        If m_Optim = 0 Then
                            Dim rectData As New CIDSFillRectangle
                            If SetFillRectRecordData(sheet, I, rectData, referencePt, compData, heightcomp) < 0 Then
                                Return -1
                            End If
                            DebugAddList(list, rectData)
                        End If
                End Select

                countUp = countUp + 1
                startCount = True
            Else
                If (startCount = True) Then
                    countUp = countUp + 1
                End If
            End If

            I = I + 1 ' to trigger the I>Rows condition, making this function return -1
            'SJ add for GUI freezing
            'TraceDoEvents()
        Loop

        If list.Count = 0 Then Return -1

        Return 0


    End Function


    'Insert a sheet after active sheet

    Public Sub InsertSheetAfter(ByVal sheet As OWC10.Worksheet, ByVal sheet_name As String)
        Try
            Dim sub_sheet As OWC10.Worksheet = m_Sheet.Worksheets.Add(, sheet)
            sub_sheet.Name = sheet_name
            sub_sheet.Range("B1:B1").Select()
            m_Sheet.ActiveWindow.FreezePanes = True
            'sub_sheet.
        Catch ex As Exception

        End Try

    End Sub

End Class



'  Pattern command compiling class
'  Description:                                                                                  
'           Compiling pattern element data to dispensing command list                                         
'  Created by:                                                                                  
'           Shen Jian                

Public Class IDSPattnCompiler
    Private pos(3) As Double
    Private m_PatternList As ArrayList  'pattern command elements list
    Private m_NoMoveList As ArrayList   'no motion element 
    Private m_NeedleIO As Integer
    Private m_PrevNeedleIO As Integer
    Private m_Plane As Integer          'motion plane: 0-XY; 1-YZ; 2-XZ
    Private m_PrevPlane As Integer
    Private m_Speed As Double
    Private m_PrevSpeed As Double
    Private m_PrevNeedlePt(3) As Double
    Private m_PrevRetractPt(3) As Double
    Private m_PrevClearPt(3) As Double
    Private m_ArcRadius As Double
    Private m_PrevArcRadius As Double
    Private m_PrevRetractSpeed As Double
    Private m_heightComp As Double        ' = board height - reference block height
    'this is a flag for whether the generated item is the first element of the dispensing program
    Private m_IsFirstElement As Boolean
    'this is a flag for whether the generated item is the last element of the dispensing program
    Private m_IsEndElement As Boolean
    'this is a flag for whether the generated item is inside a link command set
    Private m_IsLinkElement As Boolean
    'this is a flag for whether the generated item is the first element in a link command set
    Private m_IsFirstElementLink As Boolean
    'this is a flag for whether the generated item is the last element in a link command set
    Private m_IsEndElementLink As Boolean
    Private m_LinkPara As DispensePara
    Private m_LinkHeightCompS As Double
    '0: vision 1:Dry Needle 2:Dry left 3:Dry right 4:Wet 5:Wet left 6:Wet right
    Private m_RunMode As Integer

    Public Sub New(ByVal list As ArrayList)
        m_PatternList = list
        m_NoMoveList = New ArrayList
        m_NeedleIO = -1
        m_PrevNeedleIO = 25
        m_heightComp = 0.0
        m_Speed = 0.0
        m_PrevSpeed = 0.0
        m_IsFirstElement = True
        m_IsEndElement = False
        m_IsLinkElement = False
        m_IsFirstElementLink = True
        m_IsEndElementLink = False
        m_LinkHeightCompS = 0.0
        m_RunMode = 0
        m_Plane = -1
        m_PrevClearPt(2) = 0
    End Sub

    'Needle change action
    Public Sub NeedleChange(ByVal needleIOno As Integer, ByRef disCmdList As ArrayList)

        'Dim CmdStr As String
        'If (Not m_IsFirstElement) Then 'change needle during dispensing
        '    Dim speed As Double = m_PrevRetractSpeed
        '    Dim RetractHeightPt As Double = m_PrevRetractPt(2)
        '    Dim ClearanceZPt As Double = m_PrevClearPt(2)
        '    GenerateEndBlock(1, m_PrevRetractPt, speed, RetractHeightPt, ClearanceZPt, disCmdList)
        '    m_IsFirstElement = True
        'Else                    'start dispensing
        '    If needleIOno = gIORightNeedle Then
        '        AddOffIO(disCmdList, gIOLNeedleSlide)

        '    ElseIf needleIOno = gIOLeftNeedle Then
        '        AddOffIO(disCmdList, gIORNeedleSlide)

        '    End If
        'End If
        'If NotXYPlane(m_Plane) Then AddXYZBase(disCmdList) 'set motion plane
        'SetToXYPlane(m_Plane)

        'AddSpeed(disCmdList, IDS.Data.Hardware.Gantry.ElementXYSpeed)

        'Xue(Wen)                                                                    '
        '   No need to go needle changed position. It will increase the cycle time.     '
        '   Note: Test more.     

        'AddMoveXYZ(disCmdList, m_PrevRetractPt(0), m_PrevRetractPt(1), gSoftHome(2))




        'needle silde io on
        'If needleIOno = gIOLeftNeedle Then
        '    AddOnIO(disCmdList, gIOLNeedleSlide)


        'ElseIf needleIOno = gIORightNeedle Then
        '    AddOnIO(disCmdList, gIORNeedleSlide)

        'End If

    End Sub

    '  ensure approach height greater than needle gap
    Public Sub CheckApproachHeight(ByVal needdleGap As Double, ByRef approachHeight As Double)
        If approachHeight <= needdleGap Then
            approachHeight = needdleGap + 0.001
        End If
    End Sub

    '  ensure retract height greater than needle gap and clearence height,  greater than retract height
    Public Sub CheckRetractClearHeight(ByVal needdleGap As Double, ByRef retractHeight As Double, ByRef clearHeight As Double)
        If retractHeight <= needdleGap Then
            retractHeight = needdleGap + 0.001
        End If
        If clearHeight <= retractHeight Then
            clearHeight = retractHeight + 0.001
        End If

    End Sub

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Generate transit motion command between two elements
    '  comp:  dispensing posiotion
    '  ApproachZ: approach height
    '  disCmdList: dispensing command list
    '
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Public Function GenerateStartBlock(ByVal comp() As Double, ByVal ApproachZ As Double, ByRef disCmdList As ArrayList) As Integer
        Dim direction1, direction2, plane1, plane2 As Integer
        Dim Tolerance As Double = 0.0015
        Dim rtn1, rtn2 As Integer
        Dim xc1, yc1, xc2, yc2, pa, aa As Double
        Dim p1(3), p21(3), p31(3), p22(3), p32(3), p4(3) As Double
        Dim dx, dy, dz, dxc, dyc, dzc As Double
        Dim CmdStr As String
        If (m_IsFirstElement = False And NotVisionMode(m_RunMode)) Then 'not firt first element or vision mode
            m_PrevRetractPt.CopyTo(p1, 0)
            m_PrevClearPt.CopyTo(p21, 0)
            m_PrevClearPt.CopyTo(p31, 0)
            p4(0) = comp(0)
            p4(1) = comp(1)
            p4(2) = m_PrevClearPt(2)
            'insert radius element A between previous element motion and travel motion
            rtn1 = Linelinefillet3d(p1, p21, p31, p4, m_PrevArcRadius, m_ArcRadius, xc1, yc1, pa, aa, direction1, plane1)
            m_PrevClearPt.CopyTo(p1, 0)
            p4.CopyTo(p22, 0)
            p4.CopyTo(p32, 0)
            p4(0) = comp(0)
            p4(1) = comp(1)
            p4(2) = ApproachZ
            'insert second radius element B between travel motion and approach motion
            rtn2 = Linelinefillet3d(p1, p22, p32, p4, m_ArcRadius, m_PrevArcRadius, xc2, yc2, pa, aa, direction2, plane2)
            ''''''''''''''''''
            'Temperay disable helical SJ
            'rtn1 = -1
            'rtn2 = -1
        End If

        RoundTo3DecimalPoints(comp)
        RoundTo3DecimalPoints(m_PrevClearPt)
        RoundTo3DecimalPoints(m_PrevRetractPt)
        RoundTo3DecimalPoints(p21)
        RoundTo3DecimalPoints(p31)
        RoundTo3DecimalPoints(p22)
        RoundTo3DecimalPoints(p32)

        If m_IsFirstElement Or IsVisionMode(m_RunMode) Then 'vision mode or first motion
            If NotXYPlane(m_Plane) Then AddXYZBase(disCmdList) 'set plane
            SetToXYPlane(m_Plane)

            m_Speed = IDS.Data.Hardware.Gantry.ElementXYSpeed       'set motion speed
            AddSpeed(disCmdList, m_Speed)

            If (m_NoMoveList.Count > 0) Then   'add no motion elements command ex. setio,getio etc.
                If GenerateNoMoveCommandSet(disCmdList) < 0 Then Return -1
            End If

            AddMoveXY(disCmdList, comp)
            AddWaitIdle(disCmdList)   'wait for finished

        ElseIf m_IsFirstElementLink Then    'first link element
            If (rtn1 <> 0) Then      'can't insert radius
                If NotXYPlane(m_Plane) Then AddXYZBase(disCmdList)
                SetToXYPlane(m_Plane)
                'If m_Speed <> dot.Param.RetractSpeed Then

                m_Speed = m_PrevRetractSpeed
                AddSpeed(disCmdList, m_Speed)
                ' End If
                'move to retract position
                If (Math.Abs(m_PrevRetractPt(2) - m_PrevNeedlePt(2)) > Tolerance) Then
                    AddMoveXYZ(disCmdList, m_PrevRetractPt)
                End If

                m_Speed = IDS.Data.Hardware.Gantry.ElementZSpeed
                AddSpeed(disCmdList, m_Speed)
                'move to clearence position
                If (Math.Abs(m_PrevRetractPt(2) - m_PrevClearPt(2)) > Tolerance) Then
                    AddMoveXYZ(disCmdList, m_PrevClearPt)
                End If

            Else                     'radius inserted
                If (plane1 = 1) Then 'YZ
                    If NotYZPlane(m_Plane) Then AddYZXBase(disCmdList)
                    SetToYZPlane(m_Plane)
                    m_Speed = m_PrevRetractSpeed
                    AddSpeed(disCmdList, m_Speed)
                    'move to retract position
                    pos(0) = m_PrevRetractPt(1)
                    pos(1) = m_PrevRetractPt(2)
                    pos(2) = m_PrevRetractPt(0)
                    AddMoveXYZ(disCmdList, pos)
                    m_Speed = IDS.Data.Hardware.Gantry.ElementZSpeed
                    AddSpeed(disCmdList, m_Speed)
                    'move to arc's first endpoint
                    pos(0) = p21(1)
                    pos(1) = p21(2)
                    pos(2) = p21(0)
                    AddMoveXYZ(disCmdList, pos)
                ElseIf (plane1 = 2) Then 'XZ
                    If NotXZPlane(m_Plane) Then AddXZYBase(disCmdList)
                    SetToXZPlane(m_Plane)
                    m_Speed = m_PrevRetractSpeed
                    AddSpeed(disCmdList, m_Speed)
                    'move to retract position
                    pos(0) = m_PrevRetractPt(0)
                    pos(1) = m_PrevRetractPt(2)
                    pos(2) = m_PrevRetractPt(1)
                    AddMoveXYZ(disCmdList, pos)
                    m_Speed = IDS.Data.Hardware.Gantry.ElementZSpeed
                    AddSpeed(disCmdList, m_Speed)
                    'move to arc's first endpoint
                    pos(0) = p21(0)
                    pos(1) = p21(2)
                    pos(2) = p21(1)
                    AddMoveXYZ(disCmdList, pos)
                End If
            End If

            AddWaitLoaded(disCmdList)
            'set travel speed
            AddSpeed(disCmdList, IDS.Data.Hardware.Gantry.ElementXYSpeed)

            If (rtn1 <> 0) Then  'can't insert radius
                AddWaitIdle(disCmdList)
            Else
                'move helical to arc's second endpoint
                dx = p31(0) - p21(0)
                RoundTo3DecimalPoints(dx)
                dy = p31(1) - p21(1)
                RoundTo3DecimalPoints(dy)
                dz = p31(2) - p21(2)
                RoundTo3DecimalPoints(dz)
                If (plane1 = 2) Then 'XZ
                    dxc = xc1 - p21(0)
                    RoundTo3DecimalPoints(dxc)
                    dzc = yc1 - p21(2)
                    RoundTo3DecimalPoints(dzc)
                    If NotXZPlane(m_Plane) Then AddXZYBase(disCmdList)
                    AddMoveHelical(disCmdList, dx, dz, dxc, dzc, direction1, dy)
                    SetToXZPlane(m_Plane)
                ElseIf (plane1 = 1) Then  'YZ
                    dyc = xc1 - p21(1)
                    RoundTo3DecimalPoints(dyc)
                    dzc = yc1 - p21(2)
                    RoundTo3DecimalPoints(dzc)
                    If NotYZPlane(m_Plane) Then AddYZXBase(disCmdList)
                    AddMoveHelical(disCmdList, dy, dz, dyc, dzc, direction1, dx)
                    SetToYZPlane(m_Plane)
                End If

            End If

            'add none motion element command
            If (m_NoMoveList.Count > 0) Then
                If GenerateNoMoveCommandSet(disCmdList) < 0 Then Return -1
            End If

            If (rtn2 <> 0) Then     ' can't insert radius element B
                'travel to dispensing position
                If IsXYPlane(m_Plane) Then
                    pos(0) = comp(0)
                    pos(1) = comp(1)
                    pos(2) = m_PrevClearPt(2)
                    AddMoveXYZ(disCmdList, pos)
                ElseIf IsYZPlane(m_Plane) Then
                    pos(0) = comp(1)
                    pos(1) = m_PrevClearPt(2)
                    pos(2) = comp(0)
                    AddMoveXYZ(disCmdList, pos)
                Else
                    pos(0) = comp(0)
                    pos(1) = m_PrevClearPt(2)
                    pos(2) = comp(1)
                    AddMoveXYZ(disCmdList, pos)
                End If
                AddWaitIdle(disCmdList)
            Else
                'travel to radius element B's first endpoint
                If IsXYPlane(m_Plane) Then
                    pos(0) = p22(0)
                    pos(1) = p22(1)
                    pos(2) = p22(2)
                    AddMoveXYZ(disCmdList, pos)
                ElseIf IsYZPlane(m_Plane) Then
                    pos(0) = p22(1)
                    pos(1) = p22(2)
                    pos(2) = p22(0)
                    AddMoveXYZ(disCmdList, pos)
                Else
                    pos(0) = p22(0)
                    pos(1) = p22(2)
                    pos(2) = p22(1)
                    AddMoveXYZ(disCmdList, pos)
                End If

                'travel helically to radius element B's second endpoint
                dx = p32(0) - p22(0)
                RoundTo3DecimalPoints(dx)
                dy = p32(1) - p22(1)
                RoundTo3DecimalPoints(dy)
                dz = p32(2) - p22(2)
                RoundTo3DecimalPoints(dz)
                If (plane2 = 2) Then 'XZ
                    dxc = xc2 - p22(0)
                    RoundTo3DecimalPoints(dxc)
                    dzc = yc2 - p22(2)
                    RoundTo3DecimalPoints(dzc)
                    If NotXZPlane(m_Plane) Then AddXZYBase(disCmdList)
                    AddMoveHelical(disCmdList, dx, dz, dxc, dzc, direction2, dy)
                    SetToXZPlane(m_Plane)
                ElseIf (plane2 = 1) Then 'YZ
                    dyc = xc2 - p22(1)
                    RoundTo3DecimalPoints(dyc)
                    dzc = yc2 - p22(2)
                    RoundTo3DecimalPoints(dzc)
                    If NotYZPlane(m_Plane) Then AddYZXBase(disCmdList)
                    AddMoveHelical(disCmdList, dy, dz, dyc, dzc, direction2, dx)
                    SetToYZPlane(m_Plane)
                End If
            End If
        End If
        Return 0
    End Function
    'Generate transit motion command between two elements
    '  slideFlag: needle slide control flag, 0-no action; 1- retract slide
    '  comp:  dispensing posiotion
    '  speed: motion speed
    '  ApproachZ: approach height
    '  disCmdList: dispensing command list
    '
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Public Sub GenerateEndBlock(ByVal slideFlag As Integer, ByVal comp() As Double, ByVal speed As Double, ByVal RetractZ As Double, ByVal ClearanceZ As Double, ByRef disCmdList As ArrayList)
        Dim CmdStr As String
        Dim Tolerance As Double = 0.0015
        If IsVisionMode(m_RunMode) Then Exit Sub 'vision mode exit
        'set retract speed
        m_Speed = speed
        AddSpeed(disCmdList, m_Speed)

        'move to retract and clearence position
        RoundTo3DecimalPoints(RetractZ)
        RoundTo3DecimalPoints(ClearanceZ)
        If IsXYPlane(m_Plane) Then
            If (Math.Abs(RetractZ - ClearanceZ) > Tolerance) Then ' SJ 25/09/2012
                pos(0) = comp(0)
                pos(1) = comp(1)
                pos(2) = RetractZ
                AddMoveXYZ(disCmdList, pos)
                pos(0) = comp(0)
                pos(1) = comp(1)
                pos(2) = ClearanceZ
                m_Speed = IDS.Data.Hardware.Gantry.ElementZSpeed
                AddSpeed(disCmdList, m_Speed)
                AddMoveXYZ(disCmdList, pos)
            End If

        ElseIf IsYZPlane(m_Plane) Then
            pos(0) = comp(1)
            pos(1) = RetractZ
            pos(2) = comp(0)
            AddMoveXYZ(disCmdList, pos)
            If (Math.Abs(RetractZ - ClearanceZ) > Tolerance) Then ' SJ 25/09/2012
                pos(0) = comp(1)
                pos(1) = ClearanceZ
                pos(2) = comp(0)
                m_Speed = IDS.Data.Hardware.Gantry.ElementZSpeed
                AddSpeed(disCmdList, m_Speed)
                AddMoveXYZ(disCmdList, pos)
            End If

        Else
            pos(0) = comp(0)
            pos(1) = RetractZ
            pos(2) = comp(1)
            AddMoveXYZ(disCmdList, pos)
            If (Math.Abs(RetractZ - ClearanceZ) > Tolerance) Then ' SJ 25/09/2012
                pos(0) = comp(0)
                pos(1) = ClearanceZ
                pos(2) = comp(1)
                m_Speed = IDS.Data.Hardware.Gantry.ElementZSpeed
                AddSpeed(disCmdList, m_Speed)
                AddMoveXYZ(disCmdList, pos)
            End If
        End If

        AddWaitLoaded(disCmdList)
        AddSpeed(disCmdList, IDS.Data.Hardware.Gantry.ElementXYSpeed) 'set travel speed
        'retract left/rignt slides
        If slideFlag = 1 Then
            AddOffIO(disCmdList, gIOLNeedleSlide)
            AddOffIO(disCmdList, gIORNeedleSlide)
        End If
    End Sub

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Retrieve needle setup data
    '  mode:    0: vision 1:Dry Needle 2:Dry left 3:Dry right 4:Wet 5:Wet left 6:Wet right
    '  needle:  left or right
    '  off:     offset wrt camera 
    '  ioNo:    needle io no.
    '
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Public Sub GetNeedleData(ByVal mode As Integer, ByVal needle As String, ByRef off() As Double, ByRef ioNo As Integer)
        off(0) = 0.0
        off(1) = 0.0
        off(2) = 0.0
        ioNo = 0
        If (mode = 1 Or mode = 4) And needle.ToUpper = "LEFT" Or mode = 2 Or mode = 5 Then  'Dry/wet left
            off(0) = gLeftNeedleOffs(0)
            off(1) = gLeftNeedleOffs(1)
            off(2) = gLeftNeedleOffs(2)
            ioNo = gIOLeftNeedle
        End If
    End Sub

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Generate setio action
    '  elemData:      pattern element data
    '  disCmdList:    dispensing command list
    '
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Public Function GenerateSetIOCommandSet(ByVal elemData As Object, ByRef disCmdList As ArrayList) As Integer
        Dim setIO As CIDSSetIO = elemData
        'check element type
        If setIO.CmdType <> "SETIO" Then
            Return -1
        End If

        'dio on/off
        Dim CmdStr As String
        AddWaitIdle(disCmdList)
        If setIO.SetIOOnOffFlag = 1 Then
            AddOnIO(disCmdList, setIO.SetIONo)
        ElseIf setIO.SetIOOnOffFlag = 0 Then
            AddOffIO(disCmdList, setIO.SetIONo)
        Else
            MsgBox("contact engineer! program error")
        End If

        Return 0
    End Function

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Generate getio action
    '  elemData:      pattern element data
    '  disCmdList:    dispensing command list
    '
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Public Function GenerateGetIOCommandSet(ByVal elemData As Object, ByRef disCmdList As ArrayList) As Integer
        Dim getIO As CIDSGetIO = elemData
        'check element type
        If getIO.CmdType <> "GETIO" Then
            Return -1
        End If

        'wait untill dio on
        Dim CmdStr As String
        AddWaitIdle(disCmdList)
        AddWaitUntil(disCmdList, getIO.GetIONo)
        Return 0
    End Function

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Generate none motion command 
    '  dispenselist:    dispensing command list
    '
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Public Function GenerateNoMoveCommandSet(ByRef dispenselist As ArrayList) As Integer
        Dim row As Integer
        Dim elemData As Object
        Dim type As String

        Dim count As Integer = m_NoMoveList.Count
        If count < 1 Then Return 0 'no none motion elements

        For row = 1 To count
            elemData = m_NoMoveList(row - 1)
            If elemData Is Nothing Then
                Return -1
            End If

            type = elemData.CmdType

            Select Case type

                Case "SETIO"
                    If GenerateSetIOCommandSet(elemData, dispenselist) < 0 Then
                        Return -1
                    End If
                Case "GETIO"
                    If GenerateGetIOCommandSet(elemData, dispenselist) < 0 Then
                        Return -1
                    End If
                Case Else
                    Return -1
            End Select
        Next row
        m_NoMoveList.Clear()
        Return 0
    End Function

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Generate chipedge motion command 
    '  disCmdList:    dispensing command list
    '
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Public Function GenerateChipEdgeCommandSet(ByVal elemData As Object, ByRef disCmdList As ArrayList) As Integer
        '''''''''''''''''''''''''''''''''''
        Dim chipEdge As CIDSChipEdge = elemData
        Dim sheetname As String = chipEdge.SheetName
        Dim CmdStr As String

        Dim p1(3), p2(3), p3(3), comp1(3), comp2(3), comp3(3), comp4(3), comp5(3), detatchPt(3), dEndPt(3), midPt(3), off(3) As Double
        Dim ApproachZ, zDisp1, zDisp2, zDisp3, RetractZ, ClearanceZ As Double
        Dim DispenseModel As Integer = chipEdge.vParam._DispenseModel
        Dim needlename As String = chipEdge.Param.Needle

        GetNeedleData(m_RunMode, needlename, off, m_NeedleIO)
        'Change needle 
        If NotVisionMode(m_RunMode) And m_NeedleIO <> m_PrevNeedleIO Then
            NeedleChange(m_NeedleIO, disCmdList)
            m_PrevNeedleIO = m_NeedleIO
        End If
        'calculate approach height and dispensing height
        Dim fixHeight As Double = off(2) - chipEdge.HeightCompS
        ApproachZ = chipEdge.PosZ + fixHeight + chipEdge.Param.ApproachHeight
        zDisp1 = chipEdge.PosZ + fixHeight + chipEdge.Param.NeedleGap
        CheckApproachHeight(zDisp1, ApproachZ)

        'check work range
        If NotVisionMode(m_RunMode) Then
            If Not WorkAreaErrorCheckZ(zDisp1) Then
                CompileErrorDisplay(sheetname, chipEdge.CmdLineNo, 0)
                Return -1
            End If
            If Not WorkAreaErrorCheckZ(ApproachZ) Then
                CompileErrorDisplay(sheetname, chipEdge.CmdLineNo, 0)
                Return -1
            End If
        End If

        Dim bak_z As Double
        If chipEdge.vParam._Cw_CCw Then  'CW
            Select Case chipEdge.vParam._MainEdge
                Case 1   'the first chip edge
                    If (NotVisionMode(m_RunMode)) Then 'except vision mode
                        bak_z = zDisp1
                        comp1(0) = chipEdge.vParam._PointX1 - off(0)  ' needle Pos
                        comp1(1) = chipEdge.vParam._PointY1 - off(1)
                        comp1(2) = zDisp1
                        comp2(0) = chipEdge.vParam._PointX2 - off(0)  ' needle Pos
                        comp2(1) = chipEdge.vParam._PointY2 - off(1)
                        comp2(2) = zDisp1
                        comp3(0) = chipEdge.vParam._PointX3 - off(0)  ' needle Pos
                        comp3(1) = chipEdge.vParam._PointY3 - off(1)
                        comp3(2) = zDisp1
                    Else
                        bak_z = chipEdge.PosZ
                        comp1(0) = chipEdge.vParam._PointX1
                        comp1(1) = chipEdge.vParam._PointY1
                        comp1(2) = chipEdge.PosZ
                        comp2(0) = chipEdge.vParam._PointX2
                        comp2(1) = chipEdge.vParam._PointY2
                        comp2(2) = chipEdge.PosZ
                        comp3(0) = chipEdge.vParam._PointX3
                        comp3(1) = chipEdge.vParam._PointY3
                        comp3(2) = chipEdge.PosZ
                    End If
                Case 2   'the second chip edge
                    If (NotVisionMode(m_RunMode)) Then 'except vision mode
                        bak_z = zDisp1
                        comp1(0) = chipEdge.vParam._PointX2 - off(0) ' needle Pos
                        comp1(1) = chipEdge.vParam._PointY2 - off(1)
                        comp1(2) = zDisp1
                        comp2(0) = chipEdge.vParam._PointX3 - off(0) ' needle Pos
                        comp2(1) = chipEdge.vParam._PointY3 - off(1)
                        comp2(2) = zDisp1
                        comp3(0) = chipEdge.vParam._PointX4 - off(0) ' needle Pos
                        comp3(1) = chipEdge.vParam._PointY4 - off(1)
                        comp3(2) = zDisp1
                    Else
                        bak_z = chipEdge.PosZ
                        comp1(0) = chipEdge.vParam._PointX2
                        comp1(1) = chipEdge.vParam._PointY2
                        comp1(2) = chipEdge.PosZ
                        comp2(0) = chipEdge.vParam._PointX3
                        comp2(1) = chipEdge.vParam._PointY3
                        comp2(2) = chipEdge.PosZ
                        comp3(0) = chipEdge.vParam._PointX4
                        comp3(1) = chipEdge.vParam._PointY4
                        comp3(2) = chipEdge.PosZ
                    End If
                Case 3  'the third edge
                    If (NotVisionMode(m_RunMode)) Then 'except vision mode
                        bak_z = zDisp1
                        comp1(0) = chipEdge.vParam._PointX3 - off(0)  ' needle Pos
                        comp1(1) = chipEdge.vParam._PointY3 - off(1)
                        comp1(2) = zDisp1
                        comp2(0) = chipEdge.vParam._PointX4 - off(0)  ' needle Pos
                        comp2(1) = chipEdge.vParam._PointY4 - off(1)
                        comp2(2) = zDisp1
                        comp3(0) = chipEdge.vParam._PointX1 - off(0)  ' needle Pos
                        comp3(1) = chipEdge.vParam._PointY1 - off(1)
                        comp3(2) = zDisp1
                    Else
                        bak_z = chipEdge.PosZ
                        comp1(0) = chipEdge.vParam._PointX3
                        comp1(1) = chipEdge.vParam._PointY3
                        comp1(2) = chipEdge.PosZ
                        comp2(0) = chipEdge.vParam._PointX4
                        comp2(1) = chipEdge.vParam._PointY4
                        comp2(2) = chipEdge.PosZ
                        comp3(0) = chipEdge.vParam._PointX1
                        comp3(1) = chipEdge.vParam._PointY1
                        comp3(2) = chipEdge.PosZ
                    End If
                Case 4  ' the fourth edge
                    If (NotVisionMode(m_RunMode)) Then 'except vision mode
                        bak_z = zDisp1
                        comp1(0) = chipEdge.vParam._PointX4 - off(0)  ' needle Pos
                        comp1(1) = chipEdge.vParam._PointY4 - off(1)
                        comp1(2) = zDisp1
                        comp2(0) = chipEdge.vParam._PointX1 - off(0)  ' needle Pos
                        comp2(1) = chipEdge.vParam._PointY1 - off(1)
                        comp2(2) = zDisp1
                        comp3(0) = chipEdge.vParam._PointX2 - off(0)  ' needle Pos
                        comp3(1) = chipEdge.vParam._PointY2 - off(1)
                        comp3(2) = zDisp1
                    Else
                        bak_z = chipEdge.PosZ
                        comp1(0) = chipEdge.vParam._PointX4
                        comp1(1) = chipEdge.vParam._PointY4
                        comp1(2) = chipEdge.PosZ
                        comp2(0) = chipEdge.vParam._PointX1
                        comp2(1) = chipEdge.vParam._PointY1
                        comp2(2) = chipEdge.PosZ
                        comp3(0) = chipEdge.vParam._PointX2
                        comp3(1) = chipEdge.vParam._PointY2
                        comp3(2) = chipEdge.PosZ
                    End If
            End Select
        Else  'CCW
            Select Case chipEdge.vParam._MainEdge
                Case 1
                    If (NotVisionMode(m_RunMode)) Then 'except vision mode
                        bak_z = zDisp1
                        comp1(0) = chipEdge.vParam._PointX2 - off(0)  ' needle Pos
                        comp1(1) = chipEdge.vParam._PointY2 - off(1)
                        comp1(2) = zDisp1
                        comp2(0) = chipEdge.vParam._PointX1 - off(0)  ' needle Pos
                        comp2(1) = chipEdge.vParam._PointY1 - off(1)
                        comp2(2) = zDisp1
                        comp3(0) = chipEdge.vParam._PointX4 - off(0)  ' needle Pos
                        comp3(1) = chipEdge.vParam._PointY4 - off(1)
                        comp3(2) = zDisp1
                    Else
                        bak_z = chipEdge.PosZ
                        comp1(0) = chipEdge.vParam._PointX2
                        comp1(1) = chipEdge.vParam._PointY2
                        comp1(2) = bak_z
                        comp2(0) = chipEdge.vParam._PointX1
                        comp2(1) = chipEdge.vParam._PointY1
                        comp2(2) = bak_z
                        comp3(0) = chipEdge.vParam._PointX4
                        comp3(1) = chipEdge.vParam._PointY4
                        comp3(2) = bak_z
                    End If
                Case 2
                    If (NotVisionMode(m_RunMode)) Then 'except vision mode
                        bak_z = zDisp1
                        comp1(0) = chipEdge.vParam._PointX3 - off(0) ' needle Pos
                        comp1(1) = chipEdge.vParam._PointY3 - off(1)
                        comp1(2) = zDisp1
                        comp2(0) = chipEdge.vParam._PointX2 - off(0) ' needle Pos
                        comp2(1) = chipEdge.vParam._PointY2 - off(1)
                        comp2(2) = zDisp1
                        comp3(0) = chipEdge.vParam._PointX1 - off(0) ' needle Pos
                        comp3(1) = chipEdge.vParam._PointY1 - off(1)
                        comp3(2) = zDisp1
                    Else
                        bak_z = chipEdge.PosZ
                        comp1(0) = chipEdge.vParam._PointX3
                        comp1(1) = chipEdge.vParam._PointY3
                        comp1(2) = bak_z
                        comp2(0) = chipEdge.vParam._PointX2
                        comp2(1) = chipEdge.vParam._PointY2
                        comp2(2) = bak_z
                        comp3(0) = chipEdge.vParam._PointX1
                        comp3(1) = chipEdge.vParam._PointY1
                        comp3(2) = bak_z
                    End If
                Case 3
                    If (NotVisionMode(m_RunMode)) Then 'except vision mode
                        bak_z = zDisp1
                        comp1(0) = chipEdge.vParam._PointX4 - off(0)  ' needle Pos
                        comp1(1) = chipEdge.vParam._PointY4 - off(1)
                        comp1(2) = zDisp1
                        comp2(0) = chipEdge.vParam._PointX3 - off(0) ' needle Pos
                        comp2(1) = chipEdge.vParam._PointY3 - off(1)
                        comp2(2) = zDisp1
                        comp3(0) = chipEdge.vParam._PointX2 - off(0)  ' needle Pos
                        comp3(1) = chipEdge.vParam._PointY2 - off(1)
                        comp3(2) = zDisp1
                    Else
                        bak_z = chipEdge.PosZ
                        comp1(0) = chipEdge.vParam._PointX4
                        comp1(1) = chipEdge.vParam._PointY4
                        comp1(2) = bak_z
                        comp2(0) = chipEdge.vParam._PointX3
                        comp2(1) = chipEdge.vParam._PointY3
                        comp2(2) = bak_z
                        comp3(0) = chipEdge.vParam._PointX2
                        comp3(1) = chipEdge.vParam._PointY2
                        comp3(2) = bak_z
                    End If
                Case 4
                    If (NotVisionMode(m_RunMode)) Then 'except vision mode
                        bak_z = zDisp1
                        comp1(0) = chipEdge.vParam._PointX1 - off(0)  ' needle Pos
                        comp1(1) = chipEdge.vParam._PointY1 - off(1)
                        comp1(2) = zDisp1
                        comp2(0) = chipEdge.vParam._PointX4 - off(0)  ' needle Pos
                        comp2(1) = chipEdge.vParam._PointY4 - off(1)
                        comp2(2) = zDisp1
                        comp3(0) = chipEdge.vParam._PointX3 - off(0)  ' needle Pos
                        comp3(1) = chipEdge.vParam._PointY3 - off(1)
                        comp3(2) = zDisp1
                    Else
                        bak_z = chipEdge.PosZ
                        comp1(0) = chipEdge.vParam._PointX1
                        comp1(1) = chipEdge.vParam._PointY1
                        comp1(2) = bak_z
                        comp2(0) = chipEdge.vParam._PointX4
                        comp2(1) = chipEdge.vParam._PointY4
                        comp2(2) = bak_z
                        comp3(0) = chipEdge.vParam._PointX3
                        comp3(1) = chipEdge.vParam._PointY3
                        comp3(2) = bak_z
                    End If
            End Select
        End If

        Dim Ptlist As New ArrayList
        Dim clearoff As Double = chipEdge.Param.EdgeClearance
        Dim centre(3) As Double
        comp2.CopyTo(centre, 0)

        If clearoff < 0.001 Then clearoff = 0.001
        'offset chip edge by clear offset
        If GetDiaOffsetVertexlist(comp1, comp2, comp3, clearoff, Ptlist) < 0 Then
            CompileErrorDisplay(sheetname, chipEdge.CmdLineNo, 10)
            Return -1
        End If
        Ptlist(0).CopyTo(comp1, 0)
        Ptlist(1).CopyTo(comp2, 0)
        Ptlist(2).CopyTo(comp3, 0)
        Ptlist(3).CopyTo(comp4, 0)
        Ptlist(4).CopyTo(comp5, 0)
        comp1(2) = bak_z
        comp2(2) = bak_z
        comp3(2) = bak_z
        comp4(2) = bak_z
        comp5(2) = bak_z
        'calculate dispensing positions
        If DispenseModel = 0 Then   'Dot 
            comp1(0) = (comp1(0) + comp2(0)) / 2.0
            comp1(1) = (comp1(1) + comp2(1)) / 2.0
            comp1.CopyTo(dEndPt, 0)
            comp1.CopyTo(midPt, 0)
        ElseIf DispenseModel = 1 Then  'One line
            If PointOnLine3d(comp1, comp2, chipEdge.Param.DeTailDist, detatchPt) < 0 Then
                CompileErrorDisplay(sheetname, chipEdge.CmdLineNo, 11)
                Return -1
            End If
            RoundTo3DecimalPoints(detatchPt)
            If chipEdge.Param.DeTailDist > 0.001 Then
                detatchPt.CopyTo(dEndPt, 0)
                comp2.CopyTo(midPt, 0)
            ElseIf chipEdge.Param.DeTailDist < -0.001 Then
                comp2.CopyTo(dEndPt, 0)
                detatchPt.CopyTo(midPt, 0)
            Else
                comp2.CopyTo(midPt, 0)
                comp2.CopyTo(dEndPt, 0)
            End If
        ElseIf DispenseModel = 2 Then  'Two lines
            If PointOnLine3d(comp4, comp5, chipEdge.Param.DeTailDist, detatchPt) < 0 Then
                CompileErrorDisplay(sheetname, chipEdge.CmdLineNo, 10)
                Return -1
            End If
            RoundTo3DecimalPoints(detatchPt)
            If chipEdge.Param.DeTailDist > 0.001 Then
                detatchPt.CopyTo(dEndPt, 0)
                comp5.CopyTo(midPt, 0)
            ElseIf chipEdge.Param.DeTailDist < -0.001 Then
                comp5.CopyTo(dEndPt, 0)
                detatchPt.CopyTo(midPt, 0)
            Else
                comp5.CopyTo(midPt, 0)
                comp5.CopyTo(dEndPt, 0)
            End If
        End If
        Dim dc(3), de(3) As Double
        Dim direction As Integer
        If DispenseModel = 2 Then
            dc(0) = centre(0) - comp2(0)
            dc(1) = centre(1) - comp2(1)
            dc(2) = 0.0
            de(0) = comp4(0) - comp2(0)
            de(1) = comp4(1) - comp2(1)
            de(2) = 0.0
            RoundTo3DecimalPoints(de)
            RoundTo3DecimalPoints(dc)
            If chipEdge.vParam._Cw_CCw Then
                direction = 1
            Else
                direction = 0
            End If
        End If

        RoundTo3DecimalPoints(comp1)
        If Not WorkAreaErrorCheckXYZ(m_RunMode, comp1) Then
            CompileErrorDisplay(sheetname, chipEdge.CmdLineNo, 0)
            Return -1
        End If

        RoundTo3DecimalPoints(comp2)
        If Not WorkAreaErrorCheckXYZ(m_RunMode, comp2) Then
            CompileErrorDisplay(sheetname, chipEdge.CmdLineNo, 0)
            Return -1
        End If

        RoundTo3DecimalPoints(comp3)
        If Not WorkAreaErrorCheckXYZ(m_RunMode, comp3) Then
            CompileErrorDisplay(sheetname, chipEdge.CmdLineNo, 0)
            Return -1
        End If

        RoundTo3DecimalPoints(midPt)
        If Not WorkAreaErrorCheckXYZ(m_RunMode, midPt) Then
            CompileErrorDisplay(sheetname, chipEdge.CmdLineNo, 0)
            Return -1
        End If

        RoundTo3DecimalPoints(dEndPt)
        If Not WorkAreaErrorCheckXYZ(m_RunMode, dEndPt) Then
            CompileErrorDisplay(sheetname, chipEdge.CmdLineNo, 0)
            Return -1
        End If

        'calculate reatract and cleraence height
        Dim ndgapZ As Double = dEndPt(2)
        Dim surfZ As Double = ndgapZ - chipEdge.Param.NeedleGap
        RetractZ = surfZ + chipEdge.Param.RetractHeight
        ClearanceZ = surfZ + chipEdge.Param.ClearanceHeight
        CheckRetractClearHeight(ndgapZ, RetractZ, ClearanceZ)
        If NotVisionMode(m_RunMode) Then
            If Not WorkAreaErrorCheckZ(RetractZ) Then
                CompileErrorDisplay(sheetname, chipEdge.CmdLineNo, 0)
                Return -1
            End If
            If Not WorkAreaErrorCheckZ(ClearanceZ) Then
                CompileErrorDisplay(sheetname, chipEdge.CmdLineNo, 0)
                Return -1
            End If
        End If

        RoundTo3DecimalPoints(RetractZ)
        RoundTo3DecimalPoints(ClearanceZ)
        '
        If (chipEdge.Param.ArcRadius < 0.001) Then
            m_ArcRadius = gMinArcRadius
        Else
            m_ArcRadius = chipEdge.Param.ArcRadius
        End If

        'add transit motion between two elements
        If GenerateStartBlock(comp1, ApproachZ, disCmdList) < 0 Then Return -1

        If NotVisionMode(m_RunMode) Then 'except for vision mode
            RoundTo3DecimalPoints(ApproachZ)
            RoundTo3DecimalPoints(zDisp1)
            'move to approach and dispensing height

            m_Speed = IDS.Data.Hardware.Gantry.ElementZSpeed
            AddSpeed(disCmdList, m_Speed) 'set motion speed

            If IsXYPlane(m_Plane) Then
                pos(0) = comp1(0)
                pos(1) = comp1(1)
                pos(2) = ApproachZ
                AddMoveXYZ(disCmdList, pos)
                pos(0) = comp1(0)
                pos(1) = comp1(1)
                pos(2) = zDisp1
                AddMoveXYZ(disCmdList, pos)
            ElseIf IsYZPlane(m_Plane) Then
                pos(0) = comp1(1)
                pos(1) = ApproachZ
                pos(2) = comp1(0)
                AddMoveXYZ(disCmdList, pos)
                pos(0) = comp1(1)
                pos(1) = zDisp1
                pos(2) = comp1(0)
                AddMoveXYZ(disCmdList, pos)
            Else
                pos(0) = comp1(0)
                pos(1) = ApproachZ
                pos(2) = comp1(1)
                AddMoveXYZ(disCmdList, pos)
                pos(0) = comp1(0)
                pos(1) = zDisp1
                pos(2) = comp1(1)
                AddMoveXYZ(disCmdList, pos)
            End If

            AddWaitLoaded(disCmdList)
            If chipEdge.Param.DispenseOn And m_RunMode > 3 Then ' wet and dispensing on
                AddOnIO(disCmdList, m_NeedleIO)   'dispensing io on
            End If

        End If

        If Not (DispenseModel = 0) And chipEdge.Param.TravelDelay > 0 Then    'travel delay
            AddWaitDuration(disCmdList, chipEdge.Param.TravelDelay)
        End If

        If NotXYPlane(m_Plane) Then AddXYZBase(disCmdList) 'set motion base plane
        SetToXYPlane(m_Plane)

        m_Speed = chipEdge.Param.TravelSpeed  'set dispensing speed
        AddSpeed(disCmdList, m_Speed)

        If DispenseModel = 0 Then    'dot
            AddWaitDuration(disCmdList, chipEdge.vParam._DotDispensingDuration)
            If chipEdge.Param.DispenseOn And m_RunMode > 3 Then  'wet and dispensing on
                AddOffIO(disCmdList, m_NeedleIO)     'dispensing io off
            End If
        Else     'line(s)
            If System.Math.Abs(chipEdge.Param.DeTailDist) <= 0.001 Then  'no detailing 
                AddMergeOn(disCmdList)
                If NotVisionMode(m_RunMode) Then    'except for vision
                    If DispenseModel = 2 Then 'two lines
                        'dispensing to arc's first endpoint
                        pos(0) = comp2(0)
                        pos(1) = comp2(1)
                        pos(2) = comp2(2)
                        AddMoveXYZ(disCmdList, pos)
                        'AddSpeed(disCmdList, m_Speed * 3)
                        AddSpeed(disCmdList, m_Speed)
                        'AddMoveArc(disCmdList, de(0), de(1), dc(2), dc(1), direction)
                        AddMoveArc(disCmdList, de(0), de(1), dc(0), dc(1), direction)
                        AddSpeed(disCmdList, m_Speed)
                    End If
                    pos(0) = dEndPt(0)
                    pos(1) = dEndPt(1)
                    pos(2) = dEndPt(2)
                    AddMoveXYZ(disCmdList, pos)
                Else
                    If DispenseModel = 2 Then
                        pos(0) = comp2(0)
                        pos(1) = comp2(1)
                        AddMoveXY(disCmdList, pos)
                        'AddSpeed(disCmdList, m_Speed * 3)
                        AddSpeed(disCmdList, m_Speed)
                        AddMoveArc(disCmdList, de(0), de(1), dc(0), dc(1), direction)
                        AddSpeed(disCmdList, m_Speed)
                    End If
                    pos(0) = dEndPt(0)
                    pos(1) = dEndPt(1)
                    AddMoveXY(disCmdList, pos)
                End If

                AddWaitIdle(disCmdList)
                AddMergeOff(disCmdList)
                If chipEdge.Param.DispenseOn And m_RunMode > 3 Then
                    AddOffIO(disCmdList, m_NeedleIO)
                End If
            Else                            'detailing
                If (NotVisionMode(m_RunMode)) Then    'except for vision mode
                    If DispenseModel = 2 Then  'two lines
                        pos(0) = comp2(0)
                        pos(1) = comp2(1)
                        pos(2) = comp2(2)
                        AddMoveXYZ(disCmdList, pos)
                        'AddSpeed(disCmdList, m_Speed * 3)
                        AddSpeed(disCmdList, m_Speed)
                        AddMoveArc(disCmdList, de(0), de(1), dc(0), dc(1), direction)
                        AddSpeed(disCmdList, m_Speed)
                    End If
                    pos(0) = midPt(0)
                    pos(1) = midPt(1)
                    pos(2) = midPt(2)
                    AddMoveXYZ(disCmdList, pos)
                    pos(0) = dEndPt(0)
                    pos(1) = dEndPt(1)
                    pos(2) = dEndPt(2)
                    AddMoveXYZ(disCmdList, pos)
                Else
                    If DispenseModel = 2 Then
                        pos(0) = comp2(0)
                        pos(1) = comp2(1)
                        AddMoveXY(disCmdList, pos)
                        'AddSpeed(disCmdList, m_Speed * 3)
                        AddSpeed(disCmdList, m_Speed)
                        AddMoveArc(disCmdList, de(0), de(1), dc(0), dc(1), direction)
                        AddSpeed(disCmdList, m_Speed)
                    End If
                    pos(0) = dEndPt(0)
                    pos(1) = dEndPt(1)
                    AddMoveXY(disCmdList, pos)
                End If

                AddWaitIdle(disCmdList)
                If chipEdge.Param.DispenseOn And m_RunMode > 3 Then  'wet and dis
                    AddOffIO(disCmdList, m_NeedleIO)
                End If
            End If

        End If

        If chipEdge.Param.RetractDelay > 0 Then     'retract delay
            AddWaitDuration(disCmdList, chipEdge.Param.RetractDelay)
        End If

        If (m_IsEndElement) Then  'last pattern element
            Dim speed As Double = chipEdge.Param.RetractSpeed
            RoundTo3DecimalPoints(dEndPt)
            GenerateEndBlock(1, dEndPt, speed, RetractZ, ClearanceZ, disCmdList)
        End If

        'backup current element's positions

        m_PrevRetractPt(0) = dEndPt(0)
        m_PrevRetractPt(1) = dEndPt(1)
        m_PrevRetractPt(2) = RetractZ
        m_PrevClearPt(0) = dEndPt(0)
        m_PrevClearPt(1) = dEndPt(1)
        m_PrevClearPt(2) = ClearanceZ
        m_PrevArcRadius = m_ArcRadius
        m_PrevRetractSpeed = chipEdge.Param.RetractSpeed
        Return 0
    End Function

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Generate clean motion command 
    '  disCmdList:    dispensing command list
    '
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Public Function GenerateCleanCommandSet(ByVal elemData As Object, ByRef disCmdList As ArrayList) As Integer
        Dim clean As CIDSClean = elemData
        'check command type
        If clean.CmdType <> "CLEAN" Then
            Return -1
        End If

        Dim CmdStr As String
        Dim p(3), comp(3), off(3) As Double
        Dim ApproachZ, NeedleGapZ, RetractZ, ClearanceZ As Double
        p(0) = IDS.Data.Hardware.Gantry.CleanPosition.X   ' Camera pos wrt hardhome set by calibration module
        p(1) = IDS.Data.Hardware.Gantry.CleanPosition.Y

        Dim needlename As String = clean.Param.Needle

        GetNeedleData(m_RunMode, needlename, off, m_NeedleIO) 'get needle data

        If NotVisionMode(m_RunMode) And m_NeedleIO <> m_PrevNeedleIO Then  'needle change
            NeedleChange(m_NeedleIO, disCmdList)
            m_PrevNeedleIO = m_NeedleIO
        End If

        ApproachZ = IDS.Data.Hardware.Gantry.CleanPosition.Z + off(2)
        NeedleGapZ = ApproachZ
        RetractZ = m_PrevRetractPt(2)
        ClearanceZ = m_PrevClearPt(2)

        If NotVisionMode(m_RunMode) Then 'except vision mode
            comp(0) = p(0) - off(0)  ' needle Pos
            comp(1) = p(1) - off(1)
            comp(2) = ClearanceZ 'ApproachZ
        Else
            comp(0) = p(0)
            comp(1) = p(1)
        End If

        m_ArcRadius = 0.001

        Dim p1(3), p2(3), p3(3), p4(3), xc, yc, pa, aa As Double
        Dim dx, dy, dz, dxc, dyc, dzc As Double
        Dim rtn, direction, plane As Integer

        If (m_IsFirstElement = False And NotVisionMode(m_RunMode)) Then 'not first element and other than vision mode
            m_PrevRetractPt.CopyTo(p1, 0)
            m_PrevClearPt.CopyTo(p2, 0)
            m_PrevClearPt.CopyTo(p3, 0)
            comp.CopyTo(p4, 0)
            'insert fillet arc
            rtn = Linelinefillet3d(p1, p2, p3, p4, m_PrevArcRadius, m_ArcRadius, xc, yc, pa, aa, direction, plane)
        End If

        RoundTo3DecimalPoints(comp)
        RoundTo3DecimalPoints(m_PrevClearPt)
        RoundTo3DecimalPoints(m_PrevRetractPt)
        RoundTo3DecimalPoints(p2)
        RoundTo3DecimalPoints(p3)
        RoundTo3DecimalPoints(ApproachZ)
        Dim posZ(0) As Double
        posZ(0) = 0

        If (m_IsFirstElement Or m_RunMode = 0) Then 'vision mode and not first element
            If NotXYPlane(m_Plane) Then AddXYZBase(disCmdList)
            SetToXYPlane(m_Plane)

            m_Speed = IDS.Data.Hardware.Gantry.ServiceXYSpeed
            AddSpeed(disCmdList, m_Speed) 'travel speed

            If (m_NoMoveList.Count > 0) Then
                If GenerateNoMoveCommandSet(disCmdList) < 0 Then Return -1
            End If

            pos(0) = comp(0)
            pos(1) = comp(1)
            AddMoveXY(disCmdList, pos)

            AddWaitIdle(disCmdList)

        Else
            If (rtn <> 0) Then   'no arc inserted
                'If NotXYPlane(m_Plane) Then AddXYZBase(disCmdList)
                'SetToXYPlane(m_Plane)
                'm_Speed = m_PrevRetractSpeed
                'AddSpeed(disCmdList, m_Speed)
                'pos(0) = m_PrevRetractPt(0)
                'pos(1) = m_PrevRetractPt(1)
                'pos(2) = m_PrevRetractPt(2)
                'AddMoveXYZ(disCmdList, pos)
                'pos(0) = m_PrevClearPt(0)
                'pos(1) = m_PrevClearPt(1)
                'pos(2) = m_PrevClearPt(2)
                'AddMoveXYZ(disCmdList, pos)
                AddZXYBase(disCmdList)
                'SetToZXPlane(m_Plane)
                m_Speed = IDS.Data.Hardware.Gantry.ServiceZSpeed
                AddSpeed(disCmdList, m_Speed)
                AddMoveZ(disCmdList, posZ)
                AddWaitIdle(disCmdList)
            Else  'arc inserted
                AddZXYBase(disCmdList)
                'SetToZXPlane(m_Plane)
                m_Speed = IDS.Data.Hardware.Gantry.ServiceZSpeed
                AddSpeed(disCmdList, m_Speed)
                AddMoveZ(disCmdList, posZ)
                AddWaitIdle(disCmdList)
                'If IsYZPlane(plane) Then 'YZ
                '    If NotYZPlane(m_Plane) Then AddYZXBase(disCmdList)
                '    SetToYZPlane(m_Plane)
                '    m_Speed = m_PrevRetractSpeed
                '    AddSpeed(disCmdList, m_Speed)
                '    pos(0) = m_PrevRetractPt(1)
                '    pos(1) = m_PrevRetractPt(2)
                '    pos(2) = m_PrevRetractPt(0)
                '    AddMoveXYZ(disCmdList, pos)
                '    pos(0) = p2(1)
                '    pos(1) = p2(2)
                '    pos(2) = p2(0)
                '    AddMoveXYZ(disCmdList, pos)
                'ElseIf IsXZPlane(plane) Then 'XZ
                '    If NotXZPlane(m_Plane) Then AddXZYBase(disCmdList)
                '    SetToXZPlane(m_Plane)
                '    AddSpeed(disCmdList, m_Speed)
                '    pos(0) = m_PrevRetractPt(0)
                '    pos(1) = m_PrevRetractPt(2)
                '    pos(2) = m_PrevRetractPt(1)
                '    AddMoveXYZ(disCmdList, pos)
                '    pos(0) = p2(0)
                '    pos(1) = p2(2)
                '    pos(2) = p2(1)
                '    AddMoveXYZ(disCmdList, pos)
                'End If
            End If

            If IsXZPlane(plane) Then 'XZ
                AddXZYBase(disCmdList)
            ElseIf IsYZPlane(plane) Then
                AddYZXBase(disCmdList)
            ElseIf IsXYPlane(plane) Then
                AddXYZBase(disCmdList)
            End If

            AddWaitLoaded(disCmdList)
            m_Speed = IDS.Data.Hardware.Gantry.ServiceXYSpeed
            AddSpeed(disCmdList, m_Speed)

            If (m_NoMoveList.Count > 0) Then 'none motion elements exist
                If GenerateNoMoveCommandSet(disCmdList) < 0 Then Return -1
            End If

            If (rtn <> 0) Then  'no arc inserted
                AddWaitIdle(disCmdList)
            Else
                dx = p3(0) - p2(0)
                RoundTo3DecimalPoints(dx)
                dy = p3(1) - p2(1)
                RoundTo3DecimalPoints(dy)
                dz = p3(2) - p2(2)
                RoundTo3DecimalPoints(dz)
                If IsXZPlane(plane) Then 'XZ
                    dxc = xc - p2(0)
                    RoundTo3DecimalPoints(dxc)
                    dzc = yc - p2(2)
                    RoundTo3DecimalPoints(dzc)
                    If NotXZPlane(m_Plane) Then AddXZYBase(disCmdList)
                    AddMoveHelical(disCmdList, dx, dz, dxc, dzc, direction, dy)
                    SetToXZPlane(m_Plane)
                ElseIf IsYZPlane(plane) Then 'YZ
                    dyc = xc - p2(1)
                    RoundTo3DecimalPoints(dyc)
                    dzc = yc - p2(2)
                    RoundTo3DecimalPoints(dzc)
                    If NotYZPlane(m_Plane) Then AddYZXBase(disCmdList)
                    AddMoveHelical(disCmdList, dy, dz, dyc, dzc, direction, dx)
                    SetToYZPlane(m_Plane)
                End If

            End If
        End If

        If NotVisionMode(m_RunMode) Then  'except for vision mode
            'move to approach position
            m_Speed = IDS.Data.Hardware.Gantry.ServiceXYSpeed       'set motion speed
            If IsXYPlane(m_Plane) Then
                pos(0) = comp(0)
                pos(1) = comp(1)
                pos(2) = 0 'comp(2)
                AddMoveXYZ(disCmdList, pos)
                AddWaitIdle(disCmdList) 'yy
                pos(0) = comp(0)
                pos(1) = comp(1)
                pos(2) = ApproachZ
                m_Speed = IDS.Data.Hardware.Gantry.ServiceZSpeed       'set motion speed
                AddSpeed(disCmdList, m_Speed)
                AddMoveXYZ(disCmdList, pos)

            ElseIf IsYZPlane(m_Plane) Then
                pos(0) = comp(1)
                pos(1) = 0 'comp(2)
                pos(2) = comp(0)
                AddMoveXYZ(disCmdList, pos)
                AddWaitIdle(disCmdList) 'yy
                pos(0) = comp(1)
                pos(1) = ApproachZ
                pos(2) = comp(0)
                m_Speed = IDS.Data.Hardware.Gantry.ServiceZSpeed       'set motion speed
                AddSpeed(disCmdList, m_Speed)
                AddMoveXYZ(disCmdList, pos)
            Else
                pos(0) = comp(0)
                pos(1) = 0 'comp(2)
                pos(2) = comp(1)
                AddMoveXYZ(disCmdList, pos)
                AddWaitIdle(disCmdList) 'yy
                pos(0) = comp(0)
                pos(1) = ApproachZ
                pos(2) = comp(1)
                m_Speed = IDS.Data.Hardware.Gantry.ServiceZSpeed       'set motion speed
                AddSpeed(disCmdList, m_Speed)
                AddMoveXYZ(disCmdList, pos)

            End If
        Else
            pos(0) = comp(0)
            pos(1) = comp(1)
            AddMoveXY(disCmdList, pos)
        End If
        AddWaitIdle(disCmdList)

        Dim wTime As Integer = CInt(clean.Param.Duration)
        Dim rz As Double = 0
        RoundTo3DecimalPoints(rz)
        If NotVisionMode(m_RunMode) Then  'except for vision mode
            AddOnIO(disCmdList, gIOClean)
            AddWaitDuration(disCmdList, wTime)
            AddOffIO(disCmdList, gIOClean)
            'retract to tool change position
            If IsXYPlane(m_Plane) Then
                pos(0) = comp(0)
                pos(1) = comp(1)
                pos(2) = rz
                AddMoveXYZ(disCmdList, pos)
            ElseIf IsYZPlane(m_Plane) Then
                pos(0) = comp(1)
                pos(1) = rz
                pos(2) = comp(0)
                AddMoveXYZ(disCmdList, pos)
            Else
                pos(0) = comp(0)
                pos(1) = rz
                pos(2) = comp(1)
                AddMoveXYZ(disCmdList, pos)
            End If
            AddWaitIdle(disCmdList)
        Else
            AddWaitDuration(disCmdList, wTime)
        End If

        If (m_IsEndElement) Then 'last pattern element
            AddOffIO(disCmdList, gIOLNeedleSlide)
            AddOffIO(disCmdList, gIORNeedleSlide)
        End If

        m_IsFirstElement = True
        m_PrevRetractPt(0) = comp(0)
        m_PrevRetractPt(1) = comp(1)
        m_PrevRetractPt(2) = rz
        m_PrevClearPt(0) = comp(0)
        m_PrevClearPt(1) = comp(1)
        m_PrevClearPt(2) = rz
        m_PrevArcRadius = m_ArcRadius
        'm_PrevRetractSpeed = move.Param.RetractSpeed
        Return 0
    End Function

    'Generate volume calibration motion command 
    '  disCmdList:    dispensing command list
    Public Function GenerateVolumeCalibrationCommandSet(ByVal elemData As Object, ByRef disCmdList As ArrayList) As Integer
        'similar as clean element
        Dim volcal As CIDSVolumeCalibration = elemData
        If volcal.CmdType <> "VOLUMECALIBRATION" Then
            Return -1
        End If
        Dim CmdStr As String

        Dim p(3), comp(3), off(3) As Double
        Dim ApproachZ, NeedleGapZ, RetractZ, ClearanceZ As Double
        p(0) = IDS.Data.Hardware.Gantry.WeighingScalePosition.X   ' Camera pos wrt hardhome
        p(1) = IDS.Data.Hardware.Gantry.WeighingScalePosition.Y

        Dim needlename As String = volcal.Param.Needle

        GetNeedleData(m_RunMode, needlename, off, m_NeedleIO)

        If NotVisionMode(m_RunMode) And m_NeedleIO <> m_PrevNeedleIO Then
            NeedleChange(m_NeedleIO, disCmdList)
            m_PrevNeedleIO = m_NeedleIO
        End If

        ApproachZ = IDS.Data.Hardware.Gantry.WeighingScalePosition.Z + off(2)
        NeedleGapZ = ApproachZ
        RetractZ = m_PrevRetractPt(2)
        ClearanceZ = m_PrevClearPt(2)

        If (NotVisionMode(m_RunMode)) Then 'except vision mode
            comp(0) = p(0) - off(0)  ' needle Pos
            comp(1) = p(1) - off(1)
            comp(2) = ClearanceZ 'ApproachZ
        Else
            comp(0) = p(0)
            comp(1) = p(1)
        End If

        m_ArcRadius = 0.001

        Dim p1(3), p2(3), p3(3), p4(3), xc, yc, pa, aa As Double
        Dim dx, dy, dz, dxc, dyc, dzc As Double
        Dim rtn, direction, plane As Integer

        If (m_IsFirstElement = False And NotVisionMode(m_RunMode)) Then
            m_PrevRetractPt.CopyTo(p1, 0)
            m_PrevClearPt.CopyTo(p2, 0)
            m_PrevClearPt.CopyTo(p3, 0)
            comp.CopyTo(p4, 0)
            rtn = Linelinefillet3d(p1, p2, p3, p4, m_PrevArcRadius, m_ArcRadius, xc, yc, pa, aa, direction, plane)
        End If

        RoundTo3DecimalPoints(comp)
        RoundTo3DecimalPoints(m_PrevClearPt)
        RoundTo3DecimalPoints(m_PrevRetractPt)
        RoundTo3DecimalPoints(p2)
        RoundTo3DecimalPoints(p3)
        RoundTo3DecimalPoints(ApproachZ)


        If (m_IsFirstElement Or m_RunMode = 0) Then 'vision mode
            If NotXYPlane(m_Plane) Then AddXYZBase(disCmdList)
            SetToXYPlane(m_Plane)

            m_Speed = IDS.Data.Hardware.Gantry.ServiceXYSpeed
            AddSpeed(disCmdList, m_Speed)

            If (m_NoMoveList.Count > 0) Then
                If GenerateNoMoveCommandSet(disCmdList) < 0 Then Return -1
            End If

            pos(0) = comp(0)
            pos(1) = comp(1)
            AddMoveXY(disCmdList, pos)
            AddWaitIdle(disCmdList)

        Else
            If (rtn <> 0) Then   'no arc inserted
                If NotXYPlane(m_Plane) Then AddXYZBase(disCmdList)
                SetToXYPlane(m_Plane)
                m_Speed = m_PrevRetractSpeed
                AddSpeed(disCmdList, m_Speed)
                pos(0) = m_PrevRetractPt(0)
                pos(1) = m_PrevRetractPt(1)
                pos(2) = m_PrevRetractPt(2)
                AddMoveXY(disCmdList, pos)
                AddMoveXYZ(disCmdList, pos)
                pos(0) = m_PrevClearPt(0)
                pos(1) = m_PrevClearPt(1)
                pos(2) = m_PrevClearPt(2)
                AddMoveXYZ(disCmdList, pos)
            Else
                If IsYZPlane(plane) Then 'YZ
                    If NotYZPlane(m_Plane) Then AddYZXBase(disCmdList)
                    SetToYZPlane(m_Plane)
                    m_Speed = m_PrevRetractSpeed
                    AddSpeed(disCmdList, m_Speed)
                    pos(0) = m_PrevRetractPt(1)
                    pos(1) = m_PrevRetractPt(2)
                    pos(2) = m_PrevRetractPt(0)
                    AddMoveXY(disCmdList, pos)
                    AddMoveXYZ(disCmdList, pos)
                    pos(0) = p2(1)
                    pos(1) = p2(2)
                    pos(2) = p2(0)
                    AddMoveXYZ(disCmdList, pos)
                ElseIf IsXZPlane(plane) Then 'XZ
                    If NotXZPlane(m_Plane) Then AddXZYBase(disCmdList)
                    SetToXZPlane(m_Plane)
                    m_Speed = m_PrevRetractSpeed
                    AddSpeed(disCmdList, m_Speed)
                    pos(0) = m_PrevRetractPt(0)
                    pos(1) = m_PrevRetractPt(2)
                    pos(2) = m_PrevRetractPt(1)
                    AddMoveXY(disCmdList, pos)
                    AddMoveXYZ(disCmdList, pos)
                    pos(0) = p2(0)
                    pos(1) = p2(2)
                    pos(2) = p2(1)
                    AddMoveXYZ(disCmdList, pos)
                End If
            End If

            AddWaitLoaded(disCmdList)
            m_Speed = IDS.Data.Hardware.Gantry.ServiceXYSpeed
            AddSpeed(disCmdList, m_Speed)

            If (m_NoMoveList.Count > 0) Then
                If GenerateNoMoveCommandSet(disCmdList) < 0 Then Return -1
            End If

            If (rtn <> 0) Then
                AddWaitIdle(disCmdList)
            Else
                dx = p3(0) - p2(0)
                RoundTo3DecimalPoints(dx)
                dy = p3(1) - p2(1)
                RoundTo3DecimalPoints(dy)
                dz = p3(2) - p2(2)
                RoundTo3DecimalPoints(dz)
                If IsXZPlane(plane) Then 'XZ
                    dxc = xc - p2(0)
                    RoundTo3DecimalPoints(dxc)
                    dzc = yc - p2(2)
                    RoundTo3DecimalPoints(dzc)
                    If NotXZPlane(m_Plane) Then AddXZYBase(disCmdList)
                    AddMoveHelical(disCmdList, dx, dz, dxc, dzc, direction, dy)
                    SetToXZPlane(m_Plane)
                ElseIf IsYZPlane(plane) Then 'YZ
                    dyc = xc - p2(1)
                    RoundTo3DecimalPoints(dyc)
                    dzc = yc - p2(2)
                    RoundTo3DecimalPoints(dzc)
                    If NotYZPlane(m_Plane) Then AddYZXBase(disCmdList)
                    AddMoveHelical(disCmdList, dy, dz, dyc, dzc, direction, dx)
                    SetToYZPlane(m_Plane)
                End If

            End If
        End If

        If NotVisionMode(m_RunMode) Then
            m_Speed = IDS.Data.Hardware.Gantry.ServiceXYSpeed       'set motion speed
            AddSpeed(disCmdList, m_Speed)
            If IsXYPlane(m_Plane) Then
                pos(0) = comp(0)
                pos(1) = comp(1)
                pos(2) = comp(2)
                AddMoveXYZ(disCmdList, pos)
            ElseIf IsYZPlane(m_Plane) Then
                pos(0) = comp(1)
                pos(1) = comp(2)
                pos(2) = comp(0)
                AddMoveXYZ(disCmdList, pos)
            Else
                pos(0) = comp(0)
                pos(1) = comp(2)
                pos(2) = comp(1)
                AddMoveXYZ(disCmdList, pos)
            End If
        Else
            pos(0) = comp(0)
            pos(1) = comp(1)
            AddMoveXY(disCmdList, pos)
        End If
        AddWaitIdle(disCmdList)

        Dim wTime As Integer = CInt(volcal.Param.Duration)
        RoundTo3DecimalPoints(m_PrevClearPt(2))
        If NotVisionMode(m_RunMode) Then  'except forvision mode
            If m_RunMode > 3 Then AddOnIO(disCmdList, m_NeedleIO)
            AddWaitDuration(disCmdList, wTime)
            If m_RunMode > 3 Then AddOffIO(disCmdList, m_NeedleIO)
            If IsXYPlane(m_Plane) Then
                pos(0) = comp(0)
                pos(1) = comp(1)
                pos(2) = m_PrevClearPt(2)
                AddMoveXYZ(disCmdList, pos)
            ElseIf IsYZPlane(m_Plane) Then
                pos(0) = comp(1)
                pos(1) = m_PrevClearPt(2)
                pos(2) = comp(0)
                AddMoveXYZ(disCmdList, pos)
            Else
                pos(0) = comp(0)
                pos(1) = m_PrevClearPt(2)
                pos(2) = comp(1)
                AddMoveXYZ(disCmdList, pos)
            End If
            AddWaitIdle(disCmdList)
        Else
            AddWaitDuration(disCmdList, wTime)
        End If

        If m_RunMode > 3 Then
            AddVolumeCalibration(disCmdList)
        End If

        If (m_IsEndElement) Then
            AddOffIO(disCmdList, gIOLNeedleSlide)
            AddOffIO(disCmdList, gIORNeedleSlide)
        End If

        m_IsFirstElement = True
        m_PrevRetractPt(0) = comp(0)
        m_PrevRetractPt(1) = comp(1)
        m_PrevRetractPt(2) = m_PrevRetractPt(2)
        m_PrevClearPt(0) = comp(0)
        m_PrevClearPt(1) = comp(1)
        m_PrevClearPt(2) = m_PrevClearPt(2) 'ClearanceZ
        m_PrevArcRadius = m_ArcRadius
        Return 0
    End Function

    'Generate purge motion command 
    '  disCmdList:    dispensing command list
    Public Function GeneratePurgeCommandSet(ByVal elemData As Object, ByRef disCmdList As ArrayList) As Integer
        'similar as clean element
        Dim purge As CIDSPurge = elemData
        If purge.CmdType <> "PURGE" Then
            Return -1
        End If
        Dim CmdStr As String

        Dim p(3), comp(3), off(3) As Double
        Dim ApproachZ, NeedleGapZ, RetractZ, ClearanceZ As Double
        p(0) = IDS.Data.Hardware.Gantry.PurgePosition.X   ' Camera pos wrt hardhome
        p(1) = IDS.Data.Hardware.Gantry.PurgePosition.Y

        Dim needlename As String = purge.Param.Needle

        GetNeedleData(m_RunMode, needlename, off, m_NeedleIO)

        If NotVisionMode(m_RunMode) And m_NeedleIO <> m_PrevNeedleIO Then
            NeedleChange(m_NeedleIO, disCmdList)
            m_PrevNeedleIO = m_NeedleIO
        End If

        ApproachZ = IDS.Data.Hardware.Gantry.PurgePosition.Z + off(2)
        NeedleGapZ = ApproachZ
        RetractZ = m_PrevRetractPt(2)
        ClearanceZ = m_PrevClearPt(2)

        If (NotVisionMode(m_RunMode)) Then 'except vision mode
            comp(0) = p(0) - off(0)  ' needle Pos
            comp(1) = p(1) - off(1)
            comp(2) = ClearanceZ 'ApproachZ
        Else
            comp(0) = p(0)
            comp(1) = p(1)
        End If

        m_ArcRadius = 0.001

        Dim p1(3), p2(3), p3(3), p4(3), xc, yc, pa, aa As Double
        Dim dx, dy, dz, dxc, dyc, dzc As Double
        Dim rtn, direction, plane As Integer

        If (m_IsFirstElement = False And NotVisionMode(m_RunMode)) Then
            m_PrevRetractPt.CopyTo(p1, 0)
            m_PrevClearPt.CopyTo(p2, 0)
            m_PrevClearPt.CopyTo(p3, 0)
            comp.CopyTo(p4, 0)
            rtn = Linelinefillet3d(p1, p2, p3, p4, m_PrevArcRadius, m_ArcRadius, xc, yc, pa, aa, direction, plane)
        End If

        RoundTo3DecimalPoints(comp)
        RoundTo3DecimalPoints(m_PrevClearPt)
        RoundTo3DecimalPoints(m_PrevRetractPt)
        RoundTo3DecimalPoints(p2)
        RoundTo3DecimalPoints(p3)
        RoundTo3DecimalPoints(ApproachZ)
        Dim posZ(0) As Double
        posZ(0) = 0

        If (m_IsFirstElement Or m_RunMode = 0) Then 'vision mode
            If NotXYPlane(m_Plane) Then AddXYZBase(disCmdList)
            SetToXYPlane(m_Plane)

            m_Speed = IDS.Data.Hardware.Gantry.ServiceXYSpeed
            AddSpeed(disCmdList, m_Speed)

            If (m_NoMoveList.Count > 0) Then
                If GenerateNoMoveCommandSet(disCmdList) < 0 Then Return -1
            End If

            pos(0) = comp(0)
            pos(1) = comp(1)
            AddMoveXY(disCmdList, pos)
            AddWaitIdle(disCmdList)

        Else
            If (rtn <> 0) Then   'no arc inserted
                'If NotXYPlane(m_Plane) Then AddXYZBase(disCmdList)
                'SetToXYPlane(m_Plane)
                'm_Speed = m_PrevRetractSpeed
                'AddSpeed(disCmdList, m_Speed)
                'pos(0) = m_PrevRetractPt(0)
                'pos(1) = m_PrevRetractPt(1)
                'pos(2) = m_PrevRetractPt(2)
                'AddMoveXY(disCmdList, pos)
                'AddMoveXYZ(disCmdList, pos)
                'pos(0) = m_PrevClearPt(0)
                'pos(1) = m_PrevClearPt(1)
                'pos(2) = m_PrevClearPt(2)
                'AddMoveXYZ(disCmdList, pos)
                AddZXYBase(disCmdList)
                'SetToZXPlane(m_Plane)
                m_Speed = IDS.Data.Hardware.Gantry.ServiceZSpeed
                AddSpeed(disCmdList, m_Speed)
                AddMoveZ(disCmdList, posZ)
                AddWaitIdle(disCmdList)
            Else
                AddZXYBase(disCmdList)
                m_Speed = IDS.Data.Hardware.Gantry.ServiceZSpeed
                AddSpeed(disCmdList, m_Speed)
                'SetToZXPlane(m_Plane)
                AddMoveZ(disCmdList, posZ)
                AddWaitIdle(disCmdList)
                'If IsYZPlane(plane) Then 'YZ
                '    If NotYZPlane(m_Plane) Then AddYZXBase(disCmdList)
                '    SetToYZPlane(m_Plane)
                '    m_Speed = m_PrevRetractSpeed
                '    AddSpeed(disCmdList, m_Speed)
                '    pos(0) = m_PrevRetractPt(1)
                '    pos(1) = m_PrevRetractPt(2)
                '    pos(2) = m_PrevRetractPt(0)
                '    AddMoveXY(disCmdList, pos)
                '    'AddWaitIdle(disCmdList) 'yy
                '    AddMoveXYZ(disCmdList, pos)
                '    pos(0) = p2(1)
                '    pos(1) = p2(2)
                '    pos(2) = p2(0)
                '    AddMoveXYZ(disCmdList, pos)
                'ElseIf IsXZPlane(plane) Then 'XZ
                '    If NotXZPlane(m_Plane) Then AddXZYBase(disCmdList)
                '    SetToXZPlane(m_Plane)
                '    m_Speed = m_PrevRetractSpeed
                '    AddSpeed(disCmdList, m_Speed)
                '    pos(0) = m_PrevRetractPt(0)
                '    pos(1) = m_PrevRetractPt(2)
                '    pos(2) = m_PrevRetractPt(1)
                '    AddMoveXY(disCmdList, pos)
                '    'AddWaitIdle(disCmdList) 'yy
                '    AddMoveXYZ(disCmdList, pos)
                '    pos(0) = p2(0)
                '    pos(1) = p2(2)
                '    pos(2) = p2(1)
                '    AddMoveXYZ(disCmdList, pos)
                'End If
            End If
            If IsXZPlane(plane) Then 'XZ
                AddXZYBase(disCmdList)
            ElseIf IsYZPlane(plane) Then
                AddYZXBase(disCmdList)
            ElseIf IsXYPlane(plane) Then
                AddXYZBase(disCmdList)
            End If
            AddWaitLoaded(disCmdList)
            m_Speed = IDS.Data.Hardware.Gantry.ServiceXYSpeed
            AddSpeed(disCmdList, m_Speed)

            If (m_NoMoveList.Count > 0) Then
                If GenerateNoMoveCommandSet(disCmdList) < 0 Then Return -1
            End If

            If (rtn <> 0) Then
                AddWaitIdle(disCmdList)
            Else
                dx = p3(0) - p2(0)
                RoundTo3DecimalPoints(dx)
                dy = p3(1) - p2(1)
                RoundTo3DecimalPoints(dy)
                dz = p3(2) - p2(2)
                RoundTo3DecimalPoints(dz)
                If IsXZPlane(plane) Then 'XZ
                    dxc = xc - p2(0)
                    RoundTo3DecimalPoints(dxc)
                    dzc = yc - p2(2)
                    RoundTo3DecimalPoints(dzc)
                    If NotXZPlane(m_Plane) Then AddXZYBase(disCmdList)
                    AddMoveHelical(disCmdList, dx, dz, dxc, dzc, direction, dy)
                    SetToXZPlane(m_Plane)
                ElseIf IsYZPlane(plane) Then 'YZ
                    dyc = xc - p2(1)
                    RoundTo3DecimalPoints(dyc)
                    dzc = yc - p2(2)
                    RoundTo3DecimalPoints(dzc)
                    If NotYZPlane(m_Plane) Then AddYZXBase(disCmdList)
                    AddMoveHelical(disCmdList, dy, dz, dyc, dzc, direction, dx)
                    SetToYZPlane(m_Plane)
                End If

            End If
        End If

        If NotVisionMode(m_RunMode) Then
            m_Speed = IDS.Data.Hardware.Gantry.ServiceXYSpeed       'set motion speed
            AddSpeed(disCmdList, m_Speed)
            'AddZXYBase(disCmdList)
            'Dim zpos(0) As Double
            'zpos(0) = 0
            'AddMoveZ(disCmdList, zpos) 'yy
            'AddWaitIdle(disCmdList)
            If IsXYPlane(m_Plane) Then
                AddXYZBase(disCmdList)
                pos(0) = comp(0)
                pos(1) = comp(1)
                pos(2) = 0 'comp(2) 'yy
                AddMoveXYZ(disCmdList, pos)
                AddWaitIdle(disCmdList) 'yy
                pos(0) = comp(0)
                pos(1) = comp(1)
                pos(2) = ApproachZ
                m_Speed = IDS.Data.Hardware.Gantry.ServiceZSpeed       'set motion speed
                AddSpeed(disCmdList, m_Speed)
                AddMoveXYZ(disCmdList, pos)
            ElseIf IsYZPlane(m_Plane) Then
                AddYZXBase(disCmdList)
                pos(0) = comp(1)
                pos(1) = 0 'comp(2)
                pos(2) = comp(0)
                AddMoveXYZ(disCmdList, pos)
                AddWaitIdle(disCmdList) 'yy
                pos(0) = comp(1)
                pos(1) = ApproachZ
                pos(2) = comp(0)
                m_Speed = IDS.Data.Hardware.Gantry.ServiceZSpeed       'set motion speed
                AddSpeed(disCmdList, m_Speed)
                AddMoveXYZ(disCmdList, pos)
            Else
                AddXYZBase(disCmdList)
                pos(0) = comp(0)
                pos(1) = 0 'comp(2) 'yy
                pos(2) = comp(1)
                AddMoveXYZ(disCmdList, pos)
                AddWaitIdle(disCmdList) 'yy
                pos(0) = comp(0)
                pos(1) = ApproachZ
                pos(2) = comp(1)
                m_Speed = IDS.Data.Hardware.Gantry.ServiceZSpeed       'set motion speed
                AddSpeed(disCmdList, m_Speed)
                AddMoveXYZ(disCmdList, pos)
            End If
        Else
            pos(0) = comp(0)
            pos(1) = comp(1)
            AddMoveXY(disCmdList, pos)
        End If
        AddWaitIdle(disCmdList)

        Dim wTime As Integer = CInt(purge.Param.Duration)
        Dim CleanPositionZ As Double = IDS.Data.Hardware.Gantry.PurgePosition.Z
        RoundTo3DecimalPoints(CleanPositionZ)
        If NotVisionMode(m_RunMode) Then  'except forvision mode
            If m_RunMode > 3 Then AddOnIO(disCmdList, m_NeedleIO)
            AddWaitDuration(disCmdList, wTime)
            If m_RunMode > 3 Then AddOffIO(disCmdList, m_NeedleIO)
            If IsXYPlane(m_Plane) Then
                pos(0) = comp(0)
                pos(1) = comp(1)
                pos(2) = CleanPositionZ
                AddMoveXYZ(disCmdList, pos)
            ElseIf IsYZPlane(m_Plane) Then
                pos(0) = comp(1)
                pos(1) = CleanPositionZ
                pos(2) = comp(0)
                AddMoveXYZ(disCmdList, pos)
            Else
                pos(0) = comp(0)
                pos(1) = CleanPositionZ
                pos(2) = comp(1)
                AddMoveXYZ(disCmdList, pos)
            End If
            AddWaitIdle(disCmdList)
        Else
            AddWaitDuration(disCmdList, wTime)
        End If

        If (m_IsEndElement) Then
            AddOffIO(disCmdList, gIOLNeedleSlide)
            AddOffIO(disCmdList, gIORNeedleSlide)
        End If

        m_IsFirstElement = True
        m_PrevRetractPt(0) = comp(0)
        m_PrevRetractPt(1) = comp(1)
        m_PrevRetractPt(2) = CleanPositionZ
        m_PrevClearPt(0) = comp(0)
        m_PrevClearPt(1) = comp(1)
        m_PrevClearPt(2) = CleanPositionZ 'ClearanceZ
        m_PrevArcRadius = m_ArcRadius
        'm_PrevRetractSpeed = move.Param.RetractSpeed
        Return 0
    End Function

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Generate move motion command 
    '  disCmdList:    dispensing command list
    '
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Public Function GenerateMoveCommandSet(ByVal elemData As Object, ByRef disCmdList As ArrayList) As Integer
        Dim move As CIDSMove = elemData
        Dim sheetname As String = move.SheetName
        If move.CmdType <> "MOVE" Then
            Return -1
        End If
        Dim CmdStr As String

        Dim p(3), comp(3), off(3) As Double
        Dim ApproachZ, NeedleGapZ, RetractZ, ClearanceZ As Double
        p(0) = move.PosX    ' Camera pos wrt hardhome
        p(1) = move.PosY

        Dim needlename As String = move.Param.Needle

        GetNeedleData(m_RunMode, needlename, off, m_NeedleIO)

        If NotVisionMode(m_RunMode) And m_NeedleIO <> m_PrevNeedleIO Then
            NeedleChange(m_NeedleIO, disCmdList)
            m_PrevNeedleIO = m_NeedleIO
        End If

        ApproachZ = move.PosZ + off(2) - move.HeightCompS

        NeedleGapZ = ApproachZ
        RetractZ = ApproachZ
        ClearanceZ = ApproachZ

        If (NotVisionMode(m_RunMode)) Then 'except vision mode
            comp(0) = p(0) - off(0)  ' needle Pos
            comp(1) = p(1) - off(1)
            comp(2) = ApproachZ
        Else
            comp(0) = p(0)
            comp(1) = p(1)
            comp(2) = 0.0
        End If

        If (Not WorkAreaErrorCheckXYZ(m_RunMode, comp)) Then 'Or (move.PosZ <= 0.0)
            CompileErrorDisplay(sheetname, move.CmdLineNo, 0)
            Return -1
        End If

        m_ArcRadius = 0.001

        Dim p1(3), p2(3), p3(3), p4(3), xc, yc, pa, aa As Double
        Dim dx, dy, dz, dxc, dyc, dzc As Double
        Dim rtn, direction, plane As Integer

        If (m_IsFirstElement = False And NotVisionMode(m_RunMode)) Then 'not first pattern element and other than vision mode
            m_PrevRetractPt.CopyTo(p1, 0)
            m_PrevClearPt.CopyTo(p2, 0)
            m_PrevClearPt.CopyTo(p3, 0)
            comp.CopyTo(p4, 0)
            rtn = Linelinefillet3d(p1, p2, p3, p4, m_PrevArcRadius, m_ArcRadius, xc, yc, pa, aa, direction, plane)
        End If

        RoundTo3DecimalPoints(comp)
        RoundTo3DecimalPoints(m_PrevClearPt)
        RoundTo3DecimalPoints(m_PrevRetractPt)
        RoundTo3DecimalPoints(p2)
        RoundTo3DecimalPoints(p3)

        If (m_IsFirstElement Or m_RunMode = 0) Then 'vision mode or first pattern element
            If NotXYPlane(m_Plane) Then AddXYZBase(disCmdList) 'motion base plane
            SetToXYPlane(m_Plane)
            m_Speed = IDS.Data.Hardware.Gantry.ElementXYSpeed    'travel speed
            AddSpeed(disCmdList, m_Speed)
            If (m_NoMoveList.Count > 0) Then  'no motion elements exist
                If GenerateNoMoveCommandSet(disCmdList) < 0 Then Return -1
            End If
            pos(0) = comp(0)
            pos(1) = comp(1)
            AddMoveXY(disCmdList, pos)
            AddWaitIdle(disCmdList)
        Else
            If (rtn <> 0) Then   'no arc inserted
                If NotXYPlane(m_Plane) Then AddXYZBase(disCmdList)
                SetToXYPlane(m_Plane)
                m_Speed = m_PrevRetractSpeed
                AddSpeed(disCmdList, m_Speed)
                'move to retract position then clearance postion
                pos(0) = m_PrevRetractPt(0)
                pos(1) = m_PrevRetractPt(1)
                pos(2) = m_PrevRetractPt(2)
                AddMoveXYZ(disCmdList, pos)
                pos(0) = m_PrevClearPt(0)
                pos(1) = m_PrevClearPt(1)
                pos(2) = m_PrevClearPt(2)
                AddMoveXYZ(disCmdList, pos)
            Else      'arc inserted
                If IsYZPlane(plane) Then 'YZ
                    If NotYZPlane(m_Plane) Then AddYZXBase(disCmdList)
                    SetToYZPlane(m_Plane)
                    m_Speed = m_PrevRetractSpeed
                    AddSpeed(disCmdList, m_Speed)
                    pos(0) = m_PrevRetractPt(1)
                    pos(1) = m_PrevRetractPt(2)
                    pos(2) = m_PrevRetractPt(0)
                    AddMoveXYZ(disCmdList, pos)
                    pos(0) = p2(1)
                    pos(1) = p2(2)
                    pos(2) = p2(0)
                    AddMoveXYZ(disCmdList, pos)
                ElseIf IsXZPlane(plane) Then 'XZ
                    If NotXZPlane(m_Plane) Then AddXZYBase(disCmdList)
                    SetToXZPlane(m_Plane)
                    m_Speed = m_PrevRetractSpeed
                    AddSpeed(disCmdList, m_Speed)
                    pos(0) = m_PrevRetractPt(0)
                    pos(1) = m_PrevRetractPt(2)
                    pos(2) = m_PrevRetractPt(1)
                    AddMoveXYZ(disCmdList, pos)
                    pos(0) = p2(0)
                    pos(1) = p2(2)
                    pos(2) = p2(1)
                    AddMoveXYZ(disCmdList, pos)
                End If
            End If

            AddWaitLoaded(disCmdList)
            m_Speed = IDS.Data.Hardware.Gantry.ElementXYSpeed   'travel speed
            AddSpeed(disCmdList, m_Speed)

            If (m_NoMoveList.Count > 0) Then
                If GenerateNoMoveCommandSet(disCmdList) < 0 Then Return -1
            End If

            If (rtn <> 0) Then  'no arc inserted
                AddWaitIdle(disCmdList)
            Else
                dx = p3(0) - p2(0)
                RoundTo3DecimalPoints(dx)
                dy = p3(1) - p2(1)
                RoundTo3DecimalPoints(dy)
                dz = p3(2) - p2(2)
                RoundTo3DecimalPoints(dz)
                If IsXZPlane(plane) Then 'XZ
                    dxc = xc - p2(0)
                    RoundTo3DecimalPoints(dxc)
                    dzc = yc - p2(2)
                    RoundTo3DecimalPoints(dzc)
                    If NotXZPlane(m_Plane) Then AddXZYBase(disCmdList)
                    AddMoveHelical(disCmdList, dx, dz, dxc, dzc, direction, dy)
                    SetToXZPlane(m_Plane)
                ElseIf IsYZPlane(plane) Then 'YZ
                    dyc = xc - p2(1)
                    RoundTo3DecimalPoints(dyc)
                    dzc = yc - p2(2)
                    RoundTo3DecimalPoints(dzc)
                    If NotYZPlane(m_Plane) Then AddYZXBase(disCmdList)
                    AddMoveHelical(disCmdList, dy, dz, dyc, dzc, direction, dx)
                    SetToYZPlane(m_Plane)
                End If

            End If
        End If

        If NotVisionMode(m_RunMode) Then 'not vision mode
            m_Speed = IDS.Data.Hardware.Gantry.ElementZSpeed       'set motion speed
            AddSpeed(disCmdList, m_Speed)
            If IsXYPlane(m_Plane) Then
                pos(0) = comp(0)
                pos(1) = comp(1)
                pos(2) = comp(2)
                AddMoveXYZ(disCmdList, pos)
            ElseIf IsYZPlane(m_Plane) Then
                pos(0) = comp(1)
                pos(1) = comp(2)
                pos(2) = comp(0)
                AddMoveXYZ(disCmdList, pos)
            Else
                pos(0) = comp(0)
                pos(1) = comp(2)
                pos(2) = comp(1)
                AddMoveXYZ(disCmdList, pos)
            End If
        Else
            pos(0) = comp(0)
            pos(1) = comp(1)
            AddMoveXY(disCmdList, pos)
        End If
        AddWaitIdle(disCmdList)

        If (m_IsEndElement) Then
            AddOffIO(disCmdList, gIOLNeedleSlide)
            AddOffIO(disCmdList, gIORNeedleSlide)
        End If

        m_IsFirstElement = True
        m_PrevRetractPt(0) = comp(0)
        m_PrevRetractPt(1) = comp(1)
        m_PrevRetractPt(2) = RetractZ
        m_PrevClearPt(0) = comp(0)
        m_PrevClearPt(1) = comp(1)
        m_PrevClearPt(2) = ClearanceZ
        m_PrevArcRadius = m_ArcRadius
        'm_PrevRetractSpeed = move.Param.RetractSpeed
        Return 0
    End Function

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Generate wait motion command 
    '  disCmdList:    dispensing command list
    '
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Public Function GenerateWaitCommandSet(ByVal elemData As Object, ByRef disCmdList As ArrayList) As Integer
        'similar as move element
        Dim wait As CIDSWait = elemData
        Dim sheetname As String = wait.SheetName
        If wait.CmdType <> "WAIT" Then
            Return -1
        End If
        Dim CmdStr As String
        Dim posZ(0) As Double
        posZ(0) = 0
        Dim p(3), comp(3), off(3) As Double
        Dim ApproachZ, NeedleGapZ, RetractZ, ClearanceZ As Double
        p(0) = wait.PosX    ' Camera pos wrt hardhome
        p(1) = wait.PosY

        Dim needlename As String = wait.Param.Needle

        GetNeedleData(m_RunMode, needlename, off, m_NeedleIO)

        If NotVisionMode(m_RunMode) And m_NeedleIO <> m_PrevNeedleIO Then
            NeedleChange(m_NeedleIO, disCmdList)
            m_PrevNeedleIO = m_NeedleIO
        End If

        ApproachZ = wait.PosZ + off(2) - wait.HeightCompS

        NeedleGapZ = ApproachZ
        RetractZ = ApproachZ
        ClearanceZ = ApproachZ

        If (NotVisionMode(m_RunMode)) Then 'except vision mode
            comp(0) = p(0) - off(0)  ' needle Pos
            comp(1) = p(1) - off(1)
            comp(2) = ApproachZ
        Else
            comp(0) = p(0)
            comp(1) = p(1)
            comp(2) = 0.0
        End If
        If (Not WorkAreaErrorCheckXYZ(m_RunMode, comp)) Or (wait.PosZ < 0.0) Then
            CompileErrorDisplay(sheetname, wait.CmdLineNo, 0)
            Return -1
        End If

        m_ArcRadius = 0.001

        Dim p1(3), p2(3), p3(3), p4(3), xc, yc, pa, aa As Double
        Dim dx, dy, dz, dxc, dyc, dzc As Double
        Dim rtn, direction, plane As Integer

        If (m_IsFirstElement = False And NotVisionMode(m_RunMode)) Then
            m_PrevRetractPt.CopyTo(p1, 0)
            m_PrevClearPt.CopyTo(p2, 0)
            m_PrevClearPt.CopyTo(p3, 0)
            comp.CopyTo(p4, 0)
            rtn = Linelinefillet3d(p1, p2, p3, p4, m_PrevArcRadius, m_ArcRadius, xc, yc, pa, aa, direction, plane)
        End If

        RoundTo3DecimalPoints(comp)
        RoundTo3DecimalPoints(m_PrevClearPt)
        RoundTo3DecimalPoints(m_PrevRetractPt)
        RoundTo3DecimalPoints(p2)
        RoundTo3DecimalPoints(p3)

        If (m_IsFirstElement Or m_RunMode = 0) Then 'vision mode
            If NotXYPlane(m_Plane) Then AddXYZBase(disCmdList)
            SetToXYPlane(m_Plane)

            m_Speed = IDS.Data.Hardware.Gantry.ElementXYSpeed
            AddSpeed(disCmdList, m_Speed)

            If (m_NoMoveList.Count > 0) Then
                If GenerateNoMoveCommandSet(disCmdList) < 0 Then Return -1
            End If

            pos(0) = comp(0)
            pos(1) = comp(1)
            AddMoveXY(disCmdList, pos)
            AddWaitIdle(disCmdList)

        Else
            If (rtn <> 0) Then   'no arc inserted
                If NotXYPlane(m_Plane) Then AddXYZBase(disCmdList)
                SetToXYPlane(m_Plane)
                m_Speed = m_PrevRetractSpeed
                AddSpeed(disCmdList, m_Speed)
                'pos(0) = m_PrevRetractPt(0)
                'pos(1) = m_PrevRetractPt(1)
                'pos(2) = m_PrevRetractPt(2)
                'AddMoveXYZ(disCmdList, pos)
                pos(0) = m_PrevClearPt(0)
                pos(1) = m_PrevClearPt(1)
                pos(2) = m_PrevClearPt(2)
                AddMoveXYZ(disCmdList, pos)
            Else
                If IsYZPlane(plane) Then 'YZ
                    If NotYZPlane(m_Plane) Then AddYZXBase(disCmdList)
                    SetToYZPlane(m_Plane)
                    m_Speed = m_PrevRetractSpeed
                    AddSpeed(disCmdList, m_Speed)
                    pos(0) = m_PrevRetractPt(1)
                    pos(1) = m_PrevRetractPt(2)
                    pos(2) = m_PrevRetractPt(0)
                    AddMoveXYZ(disCmdList, pos)
                    pos(0) = p2(1)
                    pos(1) = p2(2)
                    pos(2) = p2(0)
                    AddMoveXYZ(disCmdList, pos)
                ElseIf IsXZPlane(plane) Then 'XZ
                    If NotXZPlane(m_Plane) Then AddXZYBase(disCmdList)
                    SetToXZPlane(m_Plane)
                    AddSpeed(disCmdList, m_Speed)
                    pos(0) = m_PrevRetractPt(0)
                    pos(1) = m_PrevRetractPt(2)
                    pos(2) = m_PrevRetractPt(1)
                    AddMoveXYZ(disCmdList, pos)
                    pos(0) = p2(0)
                    pos(1) = p2(2)
                    pos(2) = p2(1)
                    AddMoveXYZ(disCmdList, pos)
                End If
            End If

            AddWaitLoaded(disCmdList)
            m_Speed = IDS.Data.Hardware.Gantry.ElementXYSpeed
            AddSpeed(disCmdList, m_Speed)

            If (m_NoMoveList.Count > 0) Then
                If GenerateNoMoveCommandSet(disCmdList) < 0 Then Return -1
            End If

            If (rtn <> 0) Then
                AddWaitIdle(disCmdList)
            Else
                dx = p3(0) - p2(0)
                RoundTo3DecimalPoints(dx)
                dy = p3(1) - p2(1)
                RoundTo3DecimalPoints(dy)
                dz = p3(2) - p2(2)
                RoundTo3DecimalPoints(dz)
                If IsXZPlane(plane) Then 'XZ
                    dxc = xc - p2(0)
                    RoundTo3DecimalPoints(dxc)
                    dzc = yc - p2(2)
                    RoundTo3DecimalPoints(dzc)
                    If NotXZPlane(m_Plane) Then AddXZYBase(disCmdList)
                    AddMoveHelical(disCmdList, dx, dz, dxc, dzc, direction, dy)
                    SetToXZPlane(m_Plane)
                ElseIf IsYZPlane(plane) Then 'YZ
                    dyc = xc - p2(1)
                    RoundTo3DecimalPoints(dyc)
                    dzc = yc - p2(2)
                    RoundTo3DecimalPoints(dzc)
                    If NotYZPlane(m_Plane) Then AddYZXBase(disCmdList)
                    AddMoveHelical(disCmdList, dy, dz, dyc, dzc, direction, dx)
                    SetToYZPlane(m_Plane)
                End If

            End If
        End If

        If NotVisionMode(m_RunMode) Then

            m_Speed = IDS.Data.Hardware.Gantry.ElementZSpeed       'set motion speed
            AddSpeed(disCmdList, m_Speed)

            comp(2) = 0 'kr

            If IsXYPlane(m_Plane) Then
                pos(0) = comp(0)
                pos(1) = comp(1)
                pos(2) = comp(2)
                AddZXYBase(disCmdList)
                AddMoveZ(disCmdList, posZ)
                AddWaitIdle(disCmdList)
                AddXYZBase(disCmdList)
                AddMoveXYZ(disCmdList, pos)
            ElseIf IsYZPlane(m_Plane) Then
                pos(0) = comp(1)
                pos(1) = comp(2)
                pos(2) = comp(0)
                AddZXYBase(disCmdList)
                AddMoveZ(disCmdList, posZ)
                AddWaitIdle(disCmdList)
                AddYZXBase(disCmdList)
                AddMoveXYZ(disCmdList, pos)
            Else
                pos(0) = comp(0)
                pos(1) = comp(2)
                pos(2) = comp(1)
                AddZXYBase(disCmdList)
                AddMoveZ(disCmdList, posZ)
                AddWaitIdle(disCmdList)
                AddXYZBase(disCmdList)
                AddMoveXYZ(disCmdList, pos)
            End If
        Else
            pos(0) = comp(0)
            pos(1) = comp(1)
            AddMoveXY(disCmdList, pos)
        End If

        AddWaitDuration(disCmdList, CInt(wait.Param.Duration))

        If (m_IsEndElement) Then
            AddOffIO(disCmdList, gIOLNeedleSlide)
            AddOffIO(disCmdList, gIORNeedleSlide)
        End If

        m_IsFirstElement = True
        m_PrevRetractPt(0) = comp(0)
        m_PrevRetractPt(1) = comp(1)
        m_PrevRetractPt(2) = RetractZ
        m_PrevClearPt(0) = comp(0)
        m_PrevClearPt(1) = comp(1)
        m_PrevClearPt(2) = ClearanceZ
        m_PrevArcRadius = m_ArcRadius
        'm_PrevRetractSpeed = move.Param.RetractSpeed
        Return 0
    End Function

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Generate QC motion command 
    '  disCmdList:    dispensing command list
    '
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Public Function GenerateQCCommandSet(ByVal elemData As Object, ByRef disCmdList As ArrayList) As Integer
        Dim QC As CIDSQC = elemData
        Dim sheetname As String = QC.SheetName
        If QC.CmdType <> "QC" Then
            Return -1
        End If
        Dim CmdStr As String

        Dim p(3), comp(3), off(3) As Double
        Dim ApproachZ, NeedleGapZ, RetractZ, ClearanceZ As Double
        'QC check position
        p(0) = QC.PosX    ' Camera pos wrt hardhome
        p(1) = QC.PosY

        ApproachZ = m_PrevClearPt(2)
        NeedleGapZ = ApproachZ
        RetractZ = ApproachZ
        ClearanceZ = ApproachZ

        If (NotVisionMode(m_RunMode)) Then 'except vision mode
            comp(0) = p(0)
            comp(1) = p(1)
            comp(2) = ApproachZ
        Else
            comp(0) = p(0)
            comp(1) = p(1)
            comp(2) = 0.0
        End If

        If (Not WorkAreaErrorCheckXYZ(m_RunMode, comp)) Then
            CompileErrorDisplay(sheetname, QC.CmdLineNo, 0)
            Return -1
        End If

        m_ArcRadius = 0.001

        Dim p1(3), p2(3), p3(3), p4(3), xc, yc, pa, aa As Double
        Dim dx, dy, dz, dxc, dyc, dzc As Double
        Dim rtn, direction, plane As Integer

        'insert arc
        If (m_IsFirstElement = False And NotVisionMode(m_RunMode)) Then
            m_PrevRetractPt.CopyTo(p1, 0)
            m_PrevClearPt.CopyTo(p2, 0)
            m_PrevClearPt.CopyTo(p3, 0)
            comp.CopyTo(p4, 0)
            rtn = Linelinefillet3d(p1, p2, p3, p4, m_PrevArcRadius, m_ArcRadius, xc, yc, pa, aa, direction, plane)
        End If

        RoundTo3DecimalPoints(comp)
        RoundTo3DecimalPoints(m_PrevClearPt)
        RoundTo3DecimalPoints(m_PrevRetractPt)
        RoundTo3DecimalPoints(p2)
        RoundTo3DecimalPoints(p3)

        If (m_IsFirstElement Or m_RunMode = 0) Then 'vision mode
            If NotXYPlane(m_Plane) Then AddXYZBase(disCmdList)
            SetToXYPlane(m_Plane)

            m_Speed = IDS.Data.Hardware.Gantry.ElementXYSpeed
            AddSpeed(disCmdList, m_Speed)

            If (m_NoMoveList.Count > 0) Then
                If GenerateNoMoveCommandSet(disCmdList) < 0 Then Return -1
            End If

            pos(0) = comp(0)
            pos(1) = comp(1)
            AddMoveXY(disCmdList, pos)
            AddWaitIdle(disCmdList)

        Else
            m_Speed = m_PrevRetractSpeed
            If (rtn <> 0) Then   'no arc inserted
                If NotXYPlane(m_Plane) Then AddXYZBase(disCmdList)
                SetToXYPlane(m_Plane)
                AddSpeed(disCmdList, m_Speed)
                ' move to clearnce position
                pos(0) = m_PrevRetractPt(0)
                pos(1) = m_PrevRetractPt(1)
                pos(2) = m_PrevRetractPt(2)
                AddMoveXYZ(disCmdList, pos)
                pos(0) = m_PrevClearPt(0)
                pos(1) = m_PrevClearPt(1)
                pos(2) = m_PrevClearPt(2)
                AddMoveXYZ(disCmdList, pos)
            Else
                'move to inserted arc's first endpoint
                If IsYZPlane(plane) Then 'YZ
                    If NotYZPlane(m_Plane) Then AddYZXBase(disCmdList)
                    SetToYZPlane(m_Plane)
                    AddSpeed(disCmdList, m_Speed)
                    pos(0) = m_PrevRetractPt(1)
                    pos(1) = m_PrevRetractPt(2)
                    pos(2) = m_PrevRetractPt(0)
                    AddMoveXYZ(disCmdList, pos)
                    pos(0) = p2(1)
                    pos(1) = p2(2)
                    pos(2) = p2(0)
                    AddMoveXYZ(disCmdList, pos)
                ElseIf IsXZPlane(plane) Then 'XZ
                    If NotXZPlane(m_Plane) Then AddXZYBase(disCmdList)
                    SetToXZPlane(m_Plane)
                    AddSpeed(disCmdList, m_Speed)
                    pos(0) = m_PrevRetractPt(0)
                    pos(1) = m_PrevRetractPt(2)
                    pos(2) = m_PrevRetractPt(1)
                    AddMoveXYZ(disCmdList, pos)
                    pos(0) = p2(0)
                    pos(1) = p2(2)
                    pos(2) = p2(1)
                    AddMoveXYZ(disCmdList, pos)
                End If
            End If

            AddWaitLoaded(disCmdList)
            m_Speed = IDS.Data.Hardware.Gantry.ElementXYSpeed
            AddSpeed(disCmdList, m_Speed)

            If (m_NoMoveList.Count > 0) Then
                If GenerateNoMoveCommandSet(disCmdList) < 0 Then Return -1
            End If

            If (rtn <> 0) Then 'no arc inserted
                AddWaitIdle(disCmdList)
            Else
                'move helical to arc's second endpoint
                dx = p3(0) - p2(0)
                RoundTo3DecimalPoints(dx)
                dy = p3(1) - p2(1)
                RoundTo3DecimalPoints(dy)
                dz = p3(2) - p2(2)
                RoundTo3DecimalPoints(dz)
                If IsXZPlane(plane) Then 'XZ
                    dxc = xc - p2(0)
                    RoundTo3DecimalPoints(dxc)
                    dzc = yc - p2(2)
                    RoundTo3DecimalPoints(dzc)
                    If NotXZPlane(m_Plane) Then AddXZYBase(disCmdList)
                    AddMoveHelical(disCmdList, dx, dz, dxc, dzc, direction, dy)
                    SetToXZPlane(m_Plane)
                ElseIf IsYZPlane(plane) Then 'YZ
                    dyc = xc - p2(1)
                    RoundTo3DecimalPoints(dyc)
                    dzc = yc - p2(2)
                    RoundTo3DecimalPoints(dzc)
                    If NotYZPlane(m_Plane) Then AddYZXBase(disCmdList)
                    AddMoveHelical(disCmdList, dy, dz, dyc, dzc, direction, dx)
                    SetToYZPlane(m_Plane)
                End If

            End If
        End If

        'move to qc check position
        If NotVisionMode(m_RunMode) Then
            If IsXYPlane(m_Plane) Then
                pos(0) = comp(0)
                pos(1) = comp(1)
                pos(2) = comp(2)
                AddMoveXYZ(disCmdList, pos)
            ElseIf IsYZPlane(m_Plane) Then
                pos(0) = comp(1)
                pos(1) = comp(2)
                pos(2) = comp(0)
                AddMoveXYZ(disCmdList, pos)
            Else
                pos(0) = comp(0)
                pos(1) = comp(2)
                pos(2) = comp(1)
                AddMoveXYZ(disCmdList, pos)
            End If
        Else
            pos(0) = comp(0)
            pos(1) = comp(1)
            AddMoveXY(disCmdList, pos)
        End If
        AddWaitIdle(disCmdList)

        With QC.vParam
            'Set QC check vision parameters in format QC(a,b,c,d,...)
            disCmdList.Add(5)
            disCmdList.Add(._Binarized)
            If QC.vParam._BlackDot Then
                disCmdList.Add(1)
            Else
                disCmdList.Add(0)
            End If
            disCmdList.Add(._Brightness)
            disCmdList.Add(._Close)
            disCmdList.Add(._Compactness)
            disCmdList.Add(._MaxArea)
            disCmdList.Add(._MinArea)
            disCmdList.Add(._MRegionX)
            disCmdList.Add(._MRegionY)
            disCmdList.Add(._MROIx)
            disCmdList.Add(._MROIy)
            disCmdList.Add(._Open)
            disCmdList.Add(._Roughness)
            disCmdList.Add(._Diameter)
            disCmdList.Add(._Tolerance)
        End With

        If (m_IsEndElement) Then
            AddOffIO(disCmdList, gIOLNeedleSlide)
            AddOffIO(disCmdList, gIORNeedleSlide)
        End If

        m_IsFirstElement = True
        m_PrevRetractPt(0) = comp(0)
        m_PrevRetractPt(1) = comp(1)
        m_PrevRetractPt(2) = RetractZ
        m_PrevClearPt(0) = comp(0)
        m_PrevClearPt(1) = comp(1)
        m_PrevClearPt(2) = ClearanceZ
        m_PrevArcRadius = m_ArcRadius
        Return 0
    End Function

    Public Function GenerateDotArrayCommandSet(ByVal elemData As Object, ByRef disCmdList As ArrayList) As Integer
        Dim dotarray As CIDSDotArray = elemData
        Dim sheetname As String = dotarray.SheetName
        If dotarray.CmdType <> "DOTARRAY" Then
            Return -1
        End If
        Dim CmdStr As String

        Dim p1(3), p2(3), p3(3) As Double
        Dim comp1(3), comp2(2), comp3(2), off(3) As Double
        Dim ApproachZ, NeedleGapZ, RetractZ, ClearanceZ As Double

        p1(0) = dotarray.PosX1   ' Camera pos wrt hardhome
        p1(1) = dotarray.PosY1
        p1(2) = dotarray.PosZ1
        p2(0) = dotarray.PosX2
        p2(1) = dotarray.PosY2
        p3(0) = dotarray.PosX3
        p3(1) = dotarray.PosY3

        Dim needlename As String = dotarray.Param.Needle
        GetNeedleData(m_RunMode, needlename, off, m_NeedleIO)

        If NotVisionMode(m_RunMode) And m_NeedleIO <> m_PrevNeedleIO Then
            NeedleChange(m_NeedleIO, disCmdList)
            m_PrevNeedleIO = m_NeedleIO
        End If

        'approach and dispensing height
        Dim fixHeight As Double = off(2) - dotarray.HeightCompS
        ApproachZ = dotarray.PosZ1 + fixHeight + dotarray.Param.ApproachHeight
        NeedleGapZ = dotarray.PosZ1 + fixHeight + dotarray.Param.NeedleGap
        CheckApproachHeight(NeedleGapZ, ApproachZ)

        If NotVisionMode(m_RunMode) Then 'except vision mode
            comp1(0) = p1(0) - off(0)  ' needle Pos
            comp1(1) = p1(1) - off(1)
            comp1(2) = NeedleGapZ
            comp2(0) = p2(0) - off(0)  ' needle Pos
            comp2(1) = p2(1) - off(1)
            comp3(0) = p3(0) - off(0)  ' needle Pos
            comp3(1) = p3(1) - off(1)
        Else
            comp1(0) = p1(0)
            comp1(1) = p1(1)
            comp1(2) = 0.0
            comp2(0) = p2(0)
            comp2(1) = p2(1)
            comp3(0) = p3(0)
            comp3(1) = p3(1)
        End If

        'retract and clearence height wrt hard home
        RetractZ = dotarray.PosZ1 + fixHeight + dotarray.Param.RetractHeight
        ClearanceZ = dotarray.PosZ1 + fixHeight + dotarray.Param.ClearanceHeight
        CheckRetractClearHeight(NeedleGapZ, RetractZ, ClearanceZ)
        SetArcRadius(m_ArcRadius, dotarray.Param.ArcRadius)
        RoundTo3DecimalPoints(comp1)
        RoundTo3DecimalPoints(comp2)
        RoundTo3DecimalPoints(comp3)

        If GenerateStartBlock(comp1, ApproachZ, disCmdList) < 0 Then Return -1

        RoundTo3DecimalPoints(ApproachZ)
        RoundTo3DecimalPoints(NeedleGapZ)
        Dim dispense_on As Integer
        If dotarray.Param.DispenseOn = True Then
            dispense_on = 1
        Else
            dispense_on = 0
        End If
        AddDotArray(disCmdList, comp1(0), comp1(1), NeedleGapZ, comp2(0), comp2(1), comp3(0), comp3(1), dotarray.Param.Duration, RetractZ, dotarray.Param.RetractSpeed, ClearanceZ, ApproachZ, dotarray.Param.RetractDelay, IDS.Data.Hardware.Gantry.ElementXYSpeed, IDS.Data.Hardware.Gantry.ElementZSpeed, dotarray.Param.Rows, dotarray.Param.Columns, dispense_on)

        Dim LastPoint(2) As Double
        'if even then last point is not the 3rd point
        If dotarray.Param.Rows Mod 2 = 0 Then
            LastPoint(0) = comp3(0) + (comp1(0) - comp2(0))
            LastPoint(1) = comp3(1) + (comp1(1) - comp2(1))
            'odd rows then last point is the 3rd point
        Else
            LastPoint(0) = comp3(0)
            LastPoint(1) = comp3(1)
        End If

        If (m_IsEndElement) Then
            Dim speed As Double = dotarray.Param.RetractSpeed
            GenerateEndBlock(1, LastPoint, speed, RetractZ, ClearanceZ, disCmdList)
        End If

        m_PrevNeedlePt(0) = LastPoint(0)
        m_PrevNeedlePt(1) = LastPoint(1)
        m_PrevNeedlePt(2) = NeedleGapZ
        m_PrevRetractPt(0) = LastPoint(0)
        m_PrevRetractPt(1) = LastPoint(1)
        m_PrevRetractPt(2) = RetractZ
        m_PrevClearPt(0) = LastPoint(0)
        m_PrevClearPt(1) = LastPoint(1)
        m_PrevClearPt(2) = ClearanceZ
        m_PrevArcRadius = m_ArcRadius
        m_PrevRetractSpeed = dotarray.Param.RetractSpeed
        Return 0

    End Function

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Generate dot motion command 
    '  disCmdList:    dispensing command list
    '
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Public Function GenerateDotCommandSet(ByVal elemData As Object, ByRef disCmdList As ArrayList, Optional ByRef GlobalQCList As ArrayList = Nothing) As Integer
        Dim dot As CIDSDot = elemData
        Dim sheetname As String = dot.SheetName
        If dot.CmdType <> "DOT" Then
            Return -1
        End If
        Dim CmdStr As String

        Dim p(3), comp(3), off(3) As Double
        Dim ApproachZ, NeedleGapZ, RetractZ, ClearanceZ As Double

        p(0) = dot.PosX   ' Camera pos wrt hardhome
        p(1) = dot.PosY

        Dim needlename As String = dot.Param.Needle

        GetNeedleData(m_RunMode, needlename, off, m_NeedleIO)

        If NotVisionMode(m_RunMode) And m_NeedleIO <> m_PrevNeedleIO Then
            NeedleChange(m_NeedleIO, disCmdList)
            m_PrevNeedleIO = m_NeedleIO
        End If

        'approach and dispensing height
        'IDS.Data.Hardware.Needle.Left.NeedleCalibrationPosition.Z
        'current laser height reading - IDS.Data.Hardware.HeightSensor.Laser.HeightReference
        Dim fixHeight As Double = off(2) - dot.HeightCompS
        'Console.WriteLine("Laser Height: " + dot.HeightCompS.ToString())
        ApproachZ = dot.PosZ + fixHeight + dot.Param.ApproachHeight
        NeedleGapZ = dot.PosZ + fixHeight + dot.Param.NeedleGap
        CheckApproachHeight(NeedleGapZ, ApproachZ)

        If (NotVisionMode(m_RunMode)) Then 'except vision mode
            comp(0) = p(0) - off(0)  ' needle Pos
            comp(1) = p(1) - off(1)
            comp(2) = NeedleGapZ
            If Not WorkAreaErrorCheckZ(ApproachZ) Then
                CompileErrorDisplay(sheetname, dot.CmdLineNo, ApproachZError)
                Return -1
            End If
        Else
            comp(0) = p(0)
            comp(1) = p(1)
            comp(2) = 0.0

        End If
        Dim qcPost(2) As Double
        If Not (GlobalQCList Is Nothing) Then
            qcPost(0) = p(0)
            qcPost(1) = p(1)
            qcPost(2) = 0.0
            'Add global QC
            Dim qcZPost(0) As Double
            qcZPost(0) = 0.0
            'SetToXYPlane(m_Plane)
            AddZXYBase(GlobalQCList)
            AddMoveZ(GlobalQCList, qcZPost)
            AddWaitIdle(GlobalQCList)
            AddXYZBase(GlobalQCList)
            AddSpeed(GlobalQCList, IDS.Data.Hardware.Gantry.ElementXYSpeed) 'set travel speed
            AddMoveXYZ(GlobalQCList, qcPost)
            AddWaitIdle(GlobalQCList)
            GlobalQCList.Add(27) 'GlobalQC
        End If
        If Not WorkAreaErrorCheckXYZ(m_RunMode, comp) Then
            CompileErrorDisplay(sheetname, dot.CmdLineNo, 0)
            Return -1
        End If

        'retract and clearence height wrt hard home
        RetractZ = dot.PosZ + fixHeight + dot.Param.RetractHeight
        ClearanceZ = dot.PosZ + fixHeight + dot.Param.ClearanceHeight
        CheckRetractClearHeight(NeedleGapZ, RetractZ, ClearanceZ)

        If NotVisionMode(m_RunMode) Then
            If Not WorkAreaErrorCheckZ(RetractZ) Then
                CompileErrorDisplay(sheetname, dot.CmdLineNo, RetractZError)
                Return -1
            End If
            If Not WorkAreaErrorCheckZ(ClearanceZ) Then
                CompileErrorDisplay(sheetname, dot.CmdLineNo, 0)
                Return -1
            End If
        End If

        If (dot.Param.ArcRadius < 0.001) Then
            m_ArcRadius = gMinArcRadius
        Else
            m_ArcRadius = dot.Param.ArcRadius
        End If

        RoundTo3DecimalPoints(comp)
        If GenerateStartBlock(comp, ApproachZ, disCmdList) < 0 Then Return -1

        If NotVisionMode(m_RunMode) Then 'except for vision mode
            Dim Tolerance As Double = 0.0015
            RoundTo3DecimalPoints(ApproachZ)
            RoundTo3DecimalPoints(NeedleGapZ)
            m_Speed = IDS.Data.Hardware.Gantry.ElementZSpeed       'set motion speed
            AddSpeed(disCmdList, m_Speed)
            'move to approach then dispensing height
            If IsXYPlane(m_Plane) Then
                If (Math.Abs(m_PrevClearPt(2) - ApproachZ) > Tolerance) Then
                    pos(0) = comp(0)
                    pos(1) = comp(1)
                    pos(2) = ApproachZ
                    AddMoveXYZ(disCmdList, pos)
                End If

                If (Math.Abs(ApproachZ - NeedleGapZ) > Tolerance) Then  ' SJ 25/09/2012
                    pos(0) = comp(0)
                    pos(1) = comp(1)
                    pos(2) = NeedleGapZ
                    AddMoveXYZ(disCmdList, pos)
                End If
            ElseIf IsYZPlane(m_Plane) Then
                If (Math.Abs(m_PrevClearPt(2) - ApproachZ) > Tolerance) Then
                    pos(0) = comp(1)
                    pos(1) = ApproachZ
                    pos(2) = comp(0)
                    AddMoveXYZ(disCmdList, pos)
                End If

                If (Math.Abs(ApproachZ - NeedleGapZ) > Tolerance) Then ' SJ 25/09/2012
                    pos(0) = comp(1)
                    pos(1) = NeedleGapZ
                    pos(2) = comp(0)
                    AddMoveXYZ(disCmdList, pos)
                End If

            Else
                If (Math.Abs(m_PrevClearPt(2) - ApproachZ) > Tolerance) Then
                    pos(0) = comp(0)
                    pos(1) = ApproachZ
                    pos(2) = comp(1)
                    AddMoveXYZ(disCmdList, pos)
                End If
                If (Math.Abs(ApproachZ - NeedleGapZ) > Tolerance) Then ' SJ 25/09/2012
                    pos(0) = comp(0)
                    pos(1) = NeedleGapZ
                    pos(2) = comp(1)
                    AddMoveXYZ(disCmdList, pos)
                End If

            End If

            AddWaitLoaded(disCmdList)
            If dot.Param.DispenseOn And m_RunMode > 3 Then 'wet mode only
                AddOnIO(disCmdList, m_NeedleIO)   'dispensing on
            End If
        End If


        If dot.Param.Duration > 0 Then  'wait duration
            AddWaitDuration(disCmdList, dot.Param.Duration)
        End If

        If dot.Param.DispenseOn And m_RunMode > 3 Then   'dispensing off
            AddOffIO(disCmdList, m_NeedleIO)
        End If

        If dot.Param.RetractDelay > 0 Then
            AddWaitDuration(disCmdList, dot.Param.RetractDelay)
        End If

        If (m_IsEndElement) Then
            Dim speed As Double = dot.Param.RetractSpeed
            GenerateEndBlock(1, comp, speed, RetractZ, ClearanceZ, disCmdList)
        End If
        m_PrevNeedlePt(0) = comp(0)
        m_PrevNeedlePt(1) = comp(1)
        m_PrevNeedlePt(2) = comp(2)
        m_PrevRetractPt(0) = comp(0)
        m_PrevRetractPt(1) = comp(1)
        m_PrevRetractPt(2) = RetractZ
        m_PrevClearPt(0) = comp(0)
        m_PrevClearPt(1) = comp(1)
        m_PrevClearPt(2) = ClearanceZ
        m_PrevArcRadius = m_ArcRadius
        m_PrevRetractSpeed = dot.Param.RetractSpeed
        Return 0
    End Function

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Set link start flags
    ' 
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Public Function GenerateLinkSCommandSet(ByVal elemData As Object) As Integer

        Dim linkS As CIDSLinkS = elemData
        If linkS.CmdType <> "LINKSTART" Then
            Return -1
        End If
        m_IsLinkElement = True       'link elements start
        m_IsFirstElementLink = True  'is the first element within link elements
        m_IsEndElementLink = False   'not last link element
        Return 0
    End Function

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Set link end flags
    ' 
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Public Function GenerateLinkECommandSet(ByVal elemData As Object) As Integer

        Dim linkE As CIDSLinkE = elemData
        If linkE.CmdType <> "LINKEND" Then
            Return -1
        End If
        m_IsLinkElement = False      'link elements end
        m_IsFirstElementLink = True  'reset 
        m_IsEndElementLink = False   'reset
        Return 0
    End Function

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Generate line motion command 
    '  disCmdList:    dispensing command list
    '
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Public Function GenerateLineCommandSet(ByVal elemData As Object, ByRef disCmdList As ArrayList) As Integer
        '''''''''''''''''''''''''''''''''''
        Dim Line As CIDSLine = elemData
        If Line.CmdType <> "LINE" Then
            Return -1
        End If
        Dim CmdStr As String

        Dim p1(3), p2(3), comp1(3), comp2(3), detatchPt(3), dEndPt(3), midPt(3), off(3) As Double
        Dim ApproachZ, zDisp1, zDisp2, RetractZ, ClearanceZ As Double
        p1(0) = Line.PosX1    ' Camera pos
        p1(1) = Line.PosY1
        p2(0) = Line.PosX2    ' Camera pos
        p2(1) = Line.PosY2

        Dim needlename As String = Line.Param.Needle
        GetNeedleData(m_RunMode, needlename, off, m_NeedleIO)

        'needle change
        If NotVisionMode(m_RunMode) And m_NeedleIO <> m_PrevNeedleIO Then
            NeedleChange(m_NeedleIO, disCmdList)
            m_PrevNeedleIO = m_NeedleIO
        End If

        'approach height and dispensing height
        Dim fixHeight As Double = off(2) - Line.HeightCompS
        ApproachZ = Line.PosZ1 + Line.Param.ApproachHeight + fixHeight
        zDisp1 = Line.PosZ1 + Line.Param.NeedleGap + fixHeight
        zDisp2 = Line.PosZ2 + Line.Param.NeedleGap + fixHeight
        CheckApproachHeight(zDisp1, ApproachZ)

        If (NotVisionMode(m_RunMode)) Then 'except vision mode
            comp1(0) = p1(0) - off(0)  ' needle Pos
            comp1(1) = p1(1) - off(1)
            comp1(2) = zDisp1
            comp2(0) = p2(0) - off(0)  ' needle Pos
            comp2(1) = p2(1) - off(1)
            comp2(2) = zDisp2
            If Not WorkAreaErrorCheckZ(ApproachZ) Then
                CompileErrorDisplay(Line.SheetName, Line.CmdLineNo, ApproachZError)
                Return -1
            End If
        Else
            comp1(0) = p1(0)
            comp1(1) = p1(1)
            comp1(2) = Line.PosZ1
            comp2(0) = p2(0)
            comp2(1) = p2(1)
            comp2(2) = Line.PosZ2
        End If

        If Not WorkAreaErrorCheckXYZ(m_RunMode, comp1) Then
            CompileErrorDisplay(Line.SheetName, Line.CmdLineNo, 0)
            Return -1
        End If
        If Not WorkAreaErrorCheckXYZ(m_RunMode, comp2) Then
            CompileErrorDisplay(Line.SheetName, Line.CmdLineNo, 0)
            Return -1
        End If

        'detailing point
        If PointOnLine3d(comp1, comp2, Line.Param.DeTailDist, detatchPt) < 0 Then
            CompileErrorDisplay(Line.SheetName, Line.CmdLineNo, LineLengthError)
            Return -1
        End If
        If Not WorkAreaErrorCheckXYZ(m_RunMode, detatchPt) Then
            CompileErrorDisplay(Line.SheetName, Line.CmdLineNo, 0)
            Return -1
        End If

        RoundTo3DecimalPoints(comp1)
        RoundTo3DecimalPoints(comp2)
        RoundTo3DecimalPoints(detatchPt)
        If Line.Param.DeTailDist > 0.001 Then
            detatchPt.CopyTo(dEndPt, 0)
            comp2.CopyTo(midPt, 0)
        ElseIf Line.Param.DeTailDist < -0.001 Then
            comp2.CopyTo(dEndPt, 0)
            detatchPt.CopyTo(midPt, 0)
        Else
            comp2.CopyTo(dEndPt, 0)
        End If

        'retract and clearence height
        Dim ndgapZ As Double = dEndPt(2)
        Dim surfZ As Double = ndgapZ - Line.Param.NeedleGap
        RetractZ = surfZ + Line.Param.RetractHeight
        ClearanceZ = surfZ + Line.Param.ClearanceHeight
        CheckRetractClearHeight(ndgapZ, RetractZ, ClearanceZ)
        If NotVisionMode(m_RunMode) Then
            If Not WorkAreaErrorCheckZ(RetractZ) Then
                CompileErrorDisplay(Line.SheetName, Line.CmdLineNo, RetractZError)
                Return -1
            End If
            If Not WorkAreaErrorCheckZ(ClearanceZ) Then
                CompileErrorDisplay(Line.SheetName, Line.CmdLineNo, ClearanceZError)
                Return -1
            End If
        End If

        RoundTo3DecimalPoints(RetractZ)
        RoundTo3DecimalPoints(ClearanceZ)

        If (Line.Param.ArcRadius < 0.001) Then
            m_ArcRadius = gMinArcRadius
        Else
            m_ArcRadius = Line.Param.ArcRadius
        End If

        'transit motion
        If GenerateStartBlock(comp1, ApproachZ, disCmdList) < 0 Then Return -1
        m_Speed = IDS.Data.Hardware.Gantry.ElementZSpeed       'set motion speed
        AddSpeed(disCmdList, m_Speed)
        If (NotVisionMode(m_RunMode)) Then 'except for vision mode
            RoundTo3DecimalPoints(ApproachZ)
            RoundTo3DecimalPoints(zDisp1)
            'move to first dispensing endpoint
            If IsXYPlane(m_Plane) Then
                pos(0) = comp1(0)
                pos(1) = comp1(1)
                pos(2) = ApproachZ
                AddMoveXYZ(disCmdList, pos)
                pos(0) = comp1(0)
                pos(1) = comp1(1)
                pos(2) = zDisp1
                AddMoveXYZ(disCmdList, pos)
            ElseIf IsYZPlane(m_Plane) Then
                pos(0) = comp1(1)
                pos(1) = ApproachZ
                pos(2) = comp1(0)
                AddMoveXYZ(disCmdList, pos)
                pos(0) = comp1(1)
                pos(1) = zDisp1
                pos(2) = comp1(0)
                AddMoveXYZ(disCmdList, pos)
            Else
                pos(0) = comp1(0)
                pos(1) = ApproachZ
                pos(2) = comp1(1)
                AddMoveXYZ(disCmdList, pos)
                pos(0) = comp1(0)
                pos(1) = zDisp1
                pos(2) = comp1(1)
                AddMoveXYZ(disCmdList, pos)
            End If

            AddWaitLoaded(disCmdList)
            If Line.Param.DispenseOn And m_RunMode > 3 Then 'wet mode
                AddOnIO(disCmdList, m_NeedleIO)
            End If
        End If

        If Line.Param.TravelDelay > 0 Then
            AddWaitDuration(disCmdList, Line.Param.TravelDelay)
        End If

        If NotXYPlane(m_Plane) Then AddXYZBase(disCmdList)
        SetToXYPlane(m_Plane)

        m_Speed = Line.Param.TravelSpeed
        AddSpeed(disCmdList, m_Speed)

        If System.Math.Abs(Line.Param.DeTailDist) <= 0.001 Then 'no detailing
            If NotVisionMode(m_RunMode) Then
                pos(0) = dEndPt(0)
                pos(1) = dEndPt(1)
                pos(2) = dEndPt(2)
                AddMoveXYZ(disCmdList, pos)
            Else
                pos(0) = dEndPt(0)
                pos(1) = dEndPt(1)
                AddMoveXY(disCmdList, pos)
            End If
            AddWaitIdle(disCmdList)

            If Line.Param.DispenseOn And m_RunMode > 3 Then
                AddOffIO(disCmdList, m_NeedleIO)
            End If

        Else  'detailing 
            If (NotVisionMode(m_RunMode)) Then
                pos(0) = midPt(0)
                pos(1) = midPt(1)
                pos(2) = midPt(2)
                AddMoveXYZ(disCmdList, pos)
                pos(0) = dEndPt(0)
                pos(1) = dEndPt(1)
                pos(2) = dEndPt(2)
                AddMoveXYZ(disCmdList, pos)
            Else
                pos(0) = dEndPt(0)
                pos(1) = dEndPt(1)
                AddMoveXY(disCmdList, pos)
            End If

            If Line.Param.DispenseOn And m_RunMode > 3 Then
                AddWaitLoaded(disCmdList)
                AddOffIO(disCmdList, m_NeedleIO)
            End If
            AddWaitIdle(disCmdList)
        End If

        If Line.Param.RetractDelay > 0 Then
            AddWaitDuration(disCmdList, Line.Param.RetractDelay)
        End If

        If (m_IsEndElement) Then
            Dim speed As Double = Line.Param.RetractSpeed
            RoundTo3DecimalPoints(dEndPt)
            GenerateEndBlock(1, dEndPt, speed, RetractZ, ClearanceZ, disCmdList)
        End If

        m_PrevRetractPt(0) = dEndPt(0)
        m_PrevRetractPt(1) = dEndPt(1)
        m_PrevRetractPt(2) = RetractZ
        m_PrevClearPt(0) = dEndPt(0)
        m_PrevClearPt(1) = dEndPt(1)
        m_PrevClearPt(2) = ClearanceZ
        m_PrevArcRadius = m_ArcRadius
        m_PrevRetractSpeed = Line.Param.RetractSpeed

        Return 0
    End Function

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Generate link line motion command 
    '  disCmdList:    dispensing command list
    '
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Public Function GenerateLinkLineCommandSet(ByVal elemData As Object, ByRef disCmdList As ArrayList) As Integer
        '''''''''''''''''''''''''''''''''''
        Dim Line As CIDSLine = elemData
        If Line.CmdType <> "LINE" Then
            Return -1
        End If

        Dim CmdStr As String
        If (m_IsFirstElementLink) Then   'is the first link element
            'set link element dispensing parameters which will be referred by the rest link elements
            m_LinkPara = Line.Param
            m_LinkHeightCompS = Line.HeightCompS
        End If

        Dim p1(3), p2(3), comp1(3), comp2(3), detatchPt(3), dEndPt(3), midPt(3), off(3) As Double
        Dim ApproachZ, zDisp1, zDisp2, RetractZ, ClearanceZ As Double
        p1(0) = Line.PosX1   ' Camera pos
        p1(1) = Line.PosY1
        p2(0) = Line.PosX2   ' Camera pos
        p2(1) = Line.PosY2

        Dim needlename As String = m_LinkPara.Needle
        GetNeedleData(m_RunMode, needlename, off, m_NeedleIO)

        If NotVisionMode(m_RunMode) And m_NeedleIO <> m_PrevNeedleIO Then
            NeedleChange(m_NeedleIO, disCmdList)
            m_PrevNeedleIO = m_NeedleIO
        End If

        'approach and dispensing height
        Dim fixHeight As Double = off(2) - m_LinkHeightCompS
        ApproachZ = Line.PosZ1 + m_LinkPara.ApproachHeight + fixHeight
        zDisp1 = Line.PosZ1 + m_LinkPara.NeedleGap + fixHeight
        zDisp2 = Line.PosZ2 + m_LinkPara.NeedleGap + fixHeight
        CheckApproachHeight(zDisp1, ApproachZ)

        If (NotVisionMode(m_RunMode)) Then 'except vision mode
            comp1(0) = p1(0) - off(0)  ' needle Pos
            comp1(1) = p1(1) - off(1)
            comp1(2) = zDisp1
            comp2(0) = p2(0) - off(0)  ' needle Pos
            comp2(1) = p2(1) - off(1)
            comp2(2) = zDisp2
            If Not WorkAreaErrorCheckZ(ApproachZ) Then
                CompileErrorDisplay(Line.SheetName, Line.CmdLineNo, ApproachZError)
                Return -1
            End If
        Else
            comp1(0) = p1(0)
            comp1(1) = p1(1)
            comp1(2) = Line.PosZ1
            comp2(0) = p2(0)
            comp2(1) = p2(1)
            comp2(2) = Line.PosZ2
        End If

        If Not WorkAreaErrorCheckXYZ(m_RunMode, comp1) Then
            CompileErrorDisplay(Line.SheetName, Line.CmdLineNo, 0)
            Return -1
        End If
        If Not WorkAreaErrorCheckXYZ(m_RunMode, comp2) Then
            CompileErrorDisplay(Line.SheetName, Line.CmdLineNo, 0)
            Return -1
        End If

        If m_IsEndElementLink Then   'last link element
            If PointOnLine3d(comp1, comp2, m_LinkPara.DeTailDist, detatchPt) < 0 Then 'detailing point
                CompileErrorDisplay(Line.SheetName, Line.CmdLineNo, LineLengthError)
                Return -1
            End If

            If Not WorkAreaErrorCheckXYZ(m_RunMode, detatchPt) Then
                CompileErrorDisplay(Line.SheetName, Line.CmdLineNo, 0)
                Return -1
            End If

            RoundTo3DecimalPoints(comp1)
            RoundTo3DecimalPoints(comp2)
            RoundTo3DecimalPoints(detatchPt)
            If m_LinkPara.DeTailDist > 0.001 Then
                detatchPt.CopyTo(dEndPt, 0)
                comp2.CopyTo(midPt, 0)
            ElseIf m_LinkPara.DeTailDist < -0.001 Then
                comp2.CopyTo(dEndPt, 0)
                detatchPt.CopyTo(midPt, 0)
            Else
                comp2.CopyTo(dEndPt, 0)
            End If
            'retract and clearence height
            Dim ndgapZ As Double = dEndPt(2)
            Dim surfZ As Double = ndgapZ - m_LinkPara.NeedleGap
            RetractZ = surfZ + m_LinkPara.RetractHeight
            ClearanceZ = surfZ + m_LinkPara.ClearanceHeight
            CheckRetractClearHeight(ndgapZ, RetractZ, ClearanceZ)
            If NotVisionMode(m_RunMode) Then
                If Not WorkAreaErrorCheckZ(RetractZ) Then
                    CompileErrorDisplay(Line.SheetName, Line.CmdLineNo, RetractZError)
                    Return -1
                End If
                If Not WorkAreaErrorCheckZ(ClearanceZ) Then
                    CompileErrorDisplay(Line.SheetName, Line.CmdLineNo, ClearanceZError)
                    Return -1
                End If
            End If
            RoundTo3DecimalPoints(RetractZ)
            RoundTo3DecimalPoints(ClearanceZ)
        End If

        If (m_LinkPara.ArcRadius < 0.001) Then
            m_ArcRadius = gMinArcRadius
        Else
            m_ArcRadius = m_LinkPara.ArcRadius
        End If

        If (m_IsFirstElementLink) Then   'transit motion
            If GenerateStartBlock(comp1, ApproachZ, disCmdList) < 0 Then Return -1
        End If

        If (NotVisionMode(m_RunMode) And m_IsFirstElementLink) Then  'not vision mode and  first link element
            RoundTo3DecimalPoints(ApproachZ)
            RoundTo3DecimalPoints(zDisp1)
            m_Speed = IDS.Data.Hardware.Gantry.ElementZSpeed       'set motion speed
            AddSpeed(disCmdList, m_Speed)
            'move to the first endpoint of the first link element
            If IsXYPlane(m_Plane) Then
                pos(0) = comp1(0)
                pos(1) = comp1(1)
                pos(2) = ApproachZ
                AddMoveXYZ(disCmdList, pos)
                pos(0) = comp1(0)
                pos(1) = comp1(1)
                pos(2) = zDisp1
                AddMoveXYZ(disCmdList, pos)
            ElseIf IsYZPlane(m_Plane) Then
                pos(0) = comp1(1)
                pos(1) = ApproachZ
                pos(2) = comp1(0)
                AddMoveXYZ(disCmdList, pos)
                pos(0) = comp1(1)
                pos(1) = zDisp1
                pos(2) = comp1(0)
                AddMoveXYZ(disCmdList, pos)
            Else
                pos(0) = comp1(0)
                pos(1) = ApproachZ
                pos(2) = comp1(1)
                AddMoveXYZ(disCmdList, pos)
                pos(0) = comp1(0)
                pos(1) = zDisp1
                pos(2) = comp1(1)
                AddMoveXYZ(disCmdList, pos)
            End If

            AddWaitLoaded(disCmdList)
            If Line.Param.DispenseOn And m_RunMode > 3 Then 'wet mode
                AddOnIO(disCmdList, m_NeedleIO)
            End If
        End If

        If m_LinkPara.TravelDelay > 0 And m_IsFirstElementLink Then
            AddWaitDuration(disCmdList, m_LinkPara.TravelDelay)
        End If

        m_Speed = Line.Param.TravelSpeed 'dispensing speed
        AddSpeed(disCmdList, m_Speed)

        If m_IsFirstElementLink Then
            If NotXYPlane(m_Plane) Then AddXYZBase(disCmdList)
            SetToXYPlane(m_Plane)
        End If

        If (m_IsEndElementLink = False) Then
            RoundTo3DecimalPoints(comp2)

            If NotVisionMode(m_RunMode) Then
                AddSpeed(disCmdList, m_Speed)
                pos(0) = comp2(0)
                pos(1) = comp2(1)
                pos(2) = comp2(2)
                AddMoveXYZ(disCmdList, pos)
            Else
                pos(0) = comp2(0)
                pos(1) = comp2(1)
                AddMoveXY(disCmdList, pos)
            End If
            If m_IsFirstElementLink = False Then

                AddWaitLoaded(disCmdList)

                If Line.Param.DispenseOn And m_RunMode > 3 Then
                    AddWaitLoaded(disCmdList)
                    AddOnIO(disCmdList, m_NeedleIO)
                ElseIf Line.Param.DispenseOn = False And m_RunMode > 3 Then
                    AddWaitLoaded(disCmdList)
                    AddOffIO(disCmdList, m_NeedleIO)
                End If
            End If
            'test merge ON
            'AddWaitIdle(disCmdList)
        Else
            If System.Math.Abs(m_LinkPara.DeTailDist) <= 0.001 Then
                If NotVisionMode(m_RunMode) Then
                    AddSpeed(disCmdList, m_Speed)
                    pos(0) = dEndPt(0)
                    pos(1) = dEndPt(1)
                    pos(2) = dEndPt(2)
                    AddMoveXYZ(disCmdList, pos)
                Else
                    pos(0) = dEndPt(0)
                    pos(1) = dEndPt(1)
                    AddMoveXY(disCmdList, pos)
                End If

                AddWaitLoaded(disCmdList)

                If Line.Param.DispenseOn And m_RunMode > 3 Then
                    AddWaitLoaded(disCmdList)
                    AddOnIO(disCmdList, m_NeedleIO)
                ElseIf Line.Param.DispenseOn = False And m_RunMode > 3 Then
                    AddWaitLoaded(disCmdList)
                    AddOffIO(disCmdList, m_NeedleIO)
                End If

                AddWaitIdle(disCmdList)
                If Line.Param.DispenseOn And m_RunMode > 3 Then
                    AddOffIO(disCmdList, m_NeedleIO)
                End If
            Else
                If (NotVisionMode(m_RunMode)) Then
                    m_Speed = Line.Param.TravelSpeed
                    AddSpeed(disCmdList, m_Speed)

                    pos(0) = midPt(0)
                    pos(1) = midPt(1)
                    pos(2) = midPt(2)
                    AddMoveXYZ(disCmdList, pos)

                    AddWaitLoaded(disCmdList)

                    If Line.Param.DispenseOn And m_RunMode > 3 Then
                        AddWaitLoaded(disCmdList)
                        AddOnIO(disCmdList, m_NeedleIO)
                    ElseIf Line.Param.DispenseOn = False And m_RunMode > 3 Then
                        AddWaitLoaded(disCmdList)
                        AddOffIO(disCmdList, m_NeedleIO)
                    End If

                    pos(0) = dEndPt(0)
                    pos(1) = dEndPt(1)
                    pos(2) = dEndPt(2)
                    AddMoveXYZ(disCmdList, pos)
                Else
                    m_Speed = Line.Param.TravelSpeed
                    AddSpeed(disCmdList, m_Speed)
                    pos(0) = dEndPt(0)
                    pos(1) = dEndPt(1)
                    AddMoveXY(disCmdList, pos)
                    AddWaitLoaded(disCmdList)
                End If

                If Line.Param.DispenseOn And m_RunMode > 3 Then
                    AddWaitLoaded(disCmdList)
                    AddOffIO(disCmdList, m_NeedleIO)
                    AddOffIO(disCmdList, m_NeedleIO)
                End If
                AddWaitIdle(disCmdList)
            End If
        End If

        If m_LinkPara.RetractDelay > 0 And m_IsEndElementLink Then
            AddWaitDuration(disCmdList, m_LinkPara.RetractDelay)
        End If

        If (m_IsEndElement) Then
            Dim speed As Double = m_LinkPara.RetractSpeed
            RoundTo3DecimalPoints(dEndPt)
            GenerateEndBlock(1, dEndPt, speed, RetractZ, ClearanceZ, disCmdList)
        End If

        m_PrevRetractPt(0) = dEndPt(0)
        m_PrevRetractPt(1) = dEndPt(1)
        m_PrevRetractPt(2) = RetractZ
        m_PrevClearPt(0) = dEndPt(0)
        m_PrevClearPt(1) = dEndPt(1)
        m_PrevClearPt(2) = ClearanceZ
        m_PrevArcRadius = m_ArcRadius
        m_PrevRetractSpeed = Line.Param.RetractSpeed

        If (m_IsFirstElementLink) Then
            m_IsFirstElementLink = False
        End If

        Return 0
    End Function

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Generate arc motion command 
    '  disCmdList:    dispensing command list
    '
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Public Function GenerateArcCommandSet(ByVal elemData As Object, ByRef disCmdList As ArrayList) As Integer
        '''''''''''''''''''''''''''''''''''
        Dim Arc As CIDSArc = elemData
        If Arc.CmdType <> "ARC" Then
            Return -1
        End If
        Dim CmdStr As String
        Dim p1(3), p2(3), p3(3), normal(3) As Double
        p1(0) = Arc.PosX1
        p1(1) = Arc.PosY1
        p1(2) = Arc.PosZ1
        p2(0) = Arc.PosX2
        p2(1) = Arc.PosY2
        p2(2) = Arc.PosZ2
        p3(0) = Arc.PosX3
        p3(1) = Arc.PosY3
        p3(2) = Arc.PosZ3

        If Get3dNormal(p1, p2, p3, normal) < 0 Then
            CompileErrorDisplay(Arc.SheetName, Arc.CmdLineNo, 12)
            Return -1
        End If

        If (Math.Abs(Math.Abs(normal(2)) - 1.0) < 0.001) Then '2d arc
            If GenerateArc2dCommandSet(elemData, disCmdList) < 0 Then Return -1
        Else   '3d arc
            If GenerateArc3dCommandSet(elemData, disCmdList) < 0 Then Return -1
        End If

        Return 0
    End Function

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Generate 2d arc motion command 
    '  disCmdList:    dispensing command list
    '
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Public Function GenerateArc2dCommandSet(ByVal elemData As Object, ByRef disCmdList As ArrayList) As Integer
        '''''''''''''''''''''''''''''''''''
        Dim Arc As CIDSArc = elemData
        Dim CmdStr As String

        Dim p1(3), p2(3), p3(3), comp1(3), comp2(3), comp3(3), detatchPt(3), dEndPt(3), midPt(3), off(3) As Double
        Dim ApproachZ, zDisp1, zDisp2, zDisp3, RetractZ, ClearanceZ As Double
        Dim centre(3), radius As Double
        Dim direction As Integer
        p1(0) = Arc.PosX1   ' Camera pos
        p1(1) = Arc.PosY1
        p1(2) = 0.0
        p2(0) = Arc.PosX2    ' Camera pos
        p2(1) = Arc.PosY2
        p2(2) = 0.0
        p3(0) = Arc.PosX3    ' Camera pos
        p3(1) = Arc.PosY3
        p3(2) = 0.0

        Dim needlename As String = Arc.Param.Needle
        GetNeedleData(m_RunMode, needlename, off, m_NeedleIO)
        If NotVisionMode(m_RunMode) And m_NeedleIO <> m_PrevNeedleIO Then
            NeedleChange(m_NeedleIO, disCmdList)
            m_PrevNeedleIO = m_NeedleIO
        End If

        Dim fixHeight As Double = off(2) - Arc.HeightCompS
        ApproachZ = Arc.PosZ1 + Arc.Param.ApproachHeight + fixHeight
        zDisp1 = Arc.PosZ1 + Arc.Param.NeedleGap + fixHeight
        zDisp2 = Arc.PosZ2 + Arc.Param.NeedleGap + fixHeight
        zDisp3 = Arc.PosZ3 + Arc.Param.NeedleGap + fixHeight
        CheckApproachHeight(zDisp1, ApproachZ)


        If (NotVisionMode(m_RunMode)) Then 'except vision mode
            comp1(0) = p1(0) - off(0)  ' needle Pos
            comp1(1) = p1(1) - off(1)
            comp1(2) = zDisp1
            comp2(0) = p2(0) - off(0)  ' needle Pos
            comp2(1) = p2(1) - off(1)
            comp2(2) = zDisp2
            comp3(0) = p3(0) - off(0)  ' needle Pos
            comp3(1) = p3(1) - off(1)
            comp3(2) = zDisp3
            If Not WorkAreaErrorCheckZ(ApproachZ) Then
                CompileErrorDisplay(Arc.SheetName, Arc.CmdLineNo, ApproachZError)
                Return -1
            End If
        Else
            comp1(0) = p1(0)
            comp1(1) = p1(1)
            comp1(2) = Arc.PosZ1
            comp2(0) = p2(0)
            comp2(1) = p2(1)
            comp2(2) = Arc.PosZ2
            comp3(0) = p3(0)
            comp3(1) = p3(1)
            comp3(2) = Arc.PosZ3
        End If
        If Not WorkAreaErrorCheckXYZ(m_RunMode, comp1) Then
            CompileErrorDisplay(Arc.SheetName, Arc.CmdLineNo, 0)
            Return -1
        End If
        If Not WorkAreaErrorCheckXYZ(m_RunMode, comp2) Then
            CompileErrorDisplay(Arc.SheetName, Arc.CmdLineNo, 0)
            Return -1
        End If
        If Not WorkAreaErrorCheckXYZ(m_RunMode, comp3) Then
            CompileErrorDisplay(Arc.SheetName, Arc.CmdLineNo, 0)
            Return -1
        End If

        If PointOnArc3d(comp1, comp2, comp3, Arc.Param.DeTailDist, centre, radius, direction, detatchPt) < 0 Then
            CompileErrorDisplay(Arc.SheetName, Arc.CmdLineNo, 11)
            Return -1
        End If

        If Not WorkAreaErrorCheckXYZ(m_RunMode, detatchPt) Then
            CompileErrorDisplay(Arc.SheetName, Arc.CmdLineNo, 0)
            Return -1
        End If

        If Arc.Param.DeTailDist > 0.001 Then
            detatchPt.CopyTo(dEndPt, 0)
            comp3.CopyTo(midPt, 0)
        ElseIf Arc.Param.DeTailDist < -0.001 Then
            comp3.CopyTo(dEndPt, 0)
            detatchPt.CopyTo(midPt, 0)
        Else
            comp3.CopyTo(midPt, 0)
            comp3.CopyTo(dEndPt, 0)
        End If

        Dim dc1(3), dc2(3), de1(3), de2(3) As Double
        dc1(0) = centre(0) - comp1(0)
        dc1(1) = centre(1) - comp1(1)
        dc1(2) = 0.0
        dc2(0) = centre(0) - midPt(0)
        dc2(1) = centre(1) - midPt(1)
        dc2(2) = 0.0

        de1(0) = midPt(0) - comp1(0)
        de1(1) = midPt(1) - comp1(1)
        de1(2) = 0.0
        de2(0) = dEndPt(0) - midPt(0)
        de2(1) = dEndPt(1) - midPt(1)
        de2(2) = 0.0
        RoundTo3DecimalPoints(dc1)
        RoundTo3DecimalPoints(dc2)
        RoundTo3DecimalPoints(de1)
        RoundTo3DecimalPoints(de2)
        RoundTo3DecimalPoints(comp1)

        Dim ndgapZ As Double = dEndPt(2)
        Dim surfZ As Double = ndgapZ - Arc.Param.NeedleGap
        RetractZ = surfZ + Arc.Param.RetractHeight
        ClearanceZ = surfZ + Arc.Param.ClearanceHeight
        CheckRetractClearHeight(ndgapZ, RetractZ, ClearanceZ)
        If NotVisionMode(m_RunMode) Then
            If Not WorkAreaErrorCheckZ(RetractZ) Then
                CompileErrorDisplay(Arc.SheetName, Arc.CmdLineNo, RetractZError)
                Return -1
            End If

            If Not WorkAreaErrorCheckZ(ClearanceZ) Then
                CompileErrorDisplay(Arc.SheetName, Arc.CmdLineNo, ClearanceZError)
                Return -1
            End If
        End If

        RoundTo3DecimalPoints(RetractZ)
        RoundTo3DecimalPoints(ClearanceZ)
        '
        If (Arc.Param.ArcRadius < 0.001) Then
            m_ArcRadius = gMinArcRadius
        Else
            m_ArcRadius = Arc.Param.ArcRadius
        End If

        If GenerateStartBlock(comp1, ApproachZ, disCmdList) < 0 Then Return -1
        m_Speed = IDS.Data.Hardware.Gantry.ElementZSpeed       'set motion speed
        AddSpeed(disCmdList, m_Speed)
        If NotVisionMode(m_RunMode) Then
            RoundTo3DecimalPoints(ApproachZ)
            RoundTo3DecimalPoints(zDisp1)
            If IsXYPlane(m_Plane) Then
                pos(0) = comp1(0)
                pos(1) = comp1(1)
                pos(2) = ApproachZ
                AddMoveXYZ(disCmdList, pos)
                pos(0) = comp1(0)
                pos(1) = comp1(1)
                pos(2) = zDisp1
                AddMoveXYZ(disCmdList, pos)
            ElseIf IsYZPlane(m_Plane) Then
                pos(0) = comp1(1)
                pos(1) = ApproachZ
                pos(2) = comp1(0)
                AddMoveXYZ(disCmdList, pos)
                pos(0) = comp1(1)
                pos(1) = zDisp1
                pos(2) = comp1(0)
                AddMoveXYZ(disCmdList, pos)
            Else
                pos(0) = comp1(0)
                pos(1) = ApproachZ
                pos(2) = comp1(1)
                AddMoveXYZ(disCmdList, pos)
                pos(0) = comp1(0)
                pos(1) = zDisp1
                pos(2) = comp1(1)
                AddMoveXYZ(disCmdList, pos)
            End If

            AddWaitLoaded(disCmdList)
            If Arc.Param.DispenseOn And m_RunMode > 3 Then
                AddOnIO(disCmdList, m_NeedleIO)
            End If

        End If

        If Arc.Param.TravelDelay > 0 Then
            AddWaitDuration(disCmdList, Arc.Param.TravelDelay)
        End If

        If NotXYPlane(m_Plane) Then AddXYZBase(disCmdList)
        SetToXYPlane(m_Plane)

        m_Speed = Arc.Param.TravelSpeed
        AddSpeed(disCmdList, m_Speed)

        If System.Math.Abs(Arc.Param.DeTailDist) <= 0.001 Then
            AddMoveArc(disCmdList, de1(0), de1(1), dc1(0), dc1(1), direction)

            AddWaitIdle(disCmdList)
            If Arc.Param.DispenseOn And m_RunMode > 3 Then
                AddOffIO(disCmdList, m_NeedleIO)
            End If
        Else
            AddMoveArc(disCmdList, de1(0), de1(1), dc1(0), dc1(1), direction)
            AddMoveArc(disCmdList, de2(0), de2(1), dc2(0), dc2(1), direction)
            If Arc.Param.DispenseOn And m_RunMode > 3 Then
                AddWaitLoaded(disCmdList)
                AddOffIO(disCmdList, m_NeedleIO)
            End If
            AddWaitIdle(disCmdList)
        End If

        If Arc.Param.RetractDelay > 0 Then
            AddWaitDuration(disCmdList, Arc.Param.RetractDelay)
        End If

        If (m_IsEndElement) Then
            Dim speed As Double = Arc.Param.RetractSpeed
            RoundTo3DecimalPoints(dEndPt)
            GenerateEndBlock(1, dEndPt, speed, RetractZ, ClearanceZ, disCmdList)
        End If

        m_PrevRetractPt(0) = dEndPt(0)
        m_PrevRetractPt(1) = dEndPt(1)
        m_PrevRetractPt(2) = RetractZ
        m_PrevClearPt(0) = dEndPt(0)
        m_PrevClearPt(1) = dEndPt(1)
        m_PrevClearPt(2) = ClearanceZ
        m_PrevArcRadius = m_ArcRadius
        m_PrevRetractSpeed = Arc.Param.RetractSpeed
        Return 0
    End Function


    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Generate 3d arc motion command 
    '  disCmdList:    dispensing command list
    '
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Public Function GenerateArc3dCommandSet(ByVal elemData As Object, ByRef disCmdList As ArrayList) As Integer
        '''''''''''''''''''''''''''''''''''
        Dim Arc As CIDSArc = elemData
        Dim CmdStr As String

        Dim p1(3), p2(3), p3(3), comp1(3), comp2(3), comp3(3), detatchPt(3), dEndPt(3), midPt(3), off(3) As Double
        Dim ApproachZ, zDisp1, zDisp2, zDisp3, RetractZ, ClearanceZ As Double
        Dim centre(3), radius As Double
        Dim direction As Integer
        p1(0) = Arc.PosX1 '+ m_referencePt(0)   ' Camera pos
        p1(1) = Arc.PosY1 '+ m_referencePt(1)
        p1(2) = 0.0
        p2(0) = Arc.PosX2 '+ m_referencePt(0)   ' Camera pos
        p2(1) = Arc.PosY2 '+ m_referencePt(1)
        p2(2) = 0.0
        p3(0) = Arc.PosX3 '+ m_referencePt(0)   ' Camera pos
        p3(1) = Arc.PosY3 '+ m_referencePt(1)
        p3(2) = 0.0

        Dim needlename As String = Arc.Param.Needle
        GetNeedleData(m_RunMode, needlename, off, m_NeedleIO)
        If NotVisionMode(m_RunMode) And m_NeedleIO <> m_PrevNeedleIO Then
            NeedleChange(m_NeedleIO, disCmdList)
            m_PrevNeedleIO = m_NeedleIO
        End If

        Dim fixHeight As Double = off(2) - Arc.HeightCompS
        ApproachZ = Arc.PosZ1 + Arc.Param.ApproachHeight + fixHeight
        zDisp1 = Arc.PosZ1 + Arc.Param.NeedleGap + fixHeight
        zDisp2 = Arc.PosZ2 + Arc.Param.NeedleGap + fixHeight
        zDisp3 = Arc.PosZ3 + Arc.Param.NeedleGap + fixHeight
        CheckApproachHeight(zDisp1, ApproachZ)

        If (NotVisionMode(m_RunMode)) Then 'except vision mode
            comp1(0) = p1(0) - off(0)  ' needle Pos
            comp1(1) = p1(1) - off(1)
            comp1(2) = zDisp1
            comp2(0) = p2(0) - off(0)  ' needle Pos
            comp2(1) = p2(1) - off(1)
            comp2(2) = zDisp2
            comp3(0) = p3(0) - off(0)  ' needle Pos
            comp3(1) = p3(1) - off(1)
            comp3(2) = zDisp3
            If Not WorkAreaErrorCheckZ(ApproachZ) Then
                CompileErrorDisplay(Arc.SheetName, Arc.CmdLineNo, ApproachZError)
                Return -1
            End If
        Else
            comp1(0) = p1(0)
            comp1(1) = p1(1)
            comp1(2) = Arc.PosZ1
            comp2(0) = p2(0)
            comp2(1) = p2(1)
            comp2(2) = Arc.PosZ2
            comp3(0) = p3(0)
            comp3(1) = p3(1)
            comp3(2) = Arc.PosZ3
        End If
        If Not WorkAreaErrorCheckXYZ(m_RunMode, comp1) Then
            CompileErrorDisplay(Arc.SheetName, Arc.CmdLineNo, 0)
            Return -1
        End If
        If Not WorkAreaErrorCheckXYZ(m_RunMode, comp2) Then
            CompileErrorDisplay(Arc.SheetName, Arc.CmdLineNo, 0)
            Return -1
        End If
        If Not WorkAreaErrorCheckXYZ(m_RunMode, comp3) Then
            CompileErrorDisplay(Arc.SheetName, Arc.CmdLineNo, 0)
            Return -1
        End If

        Dim detaildist As Double = Arc.Param.DeTailDist
        Dim pointlist As New ArrayList

        If GetArc3dPointList(comp1, comp2, comp3, detaildist, pointlist) < 0 Then  'line segments
            CompileErrorDisplay(Arc.SheetName, Arc.CmdLineNo, 11)
            Return -1
        End If

        Dim pointno As Integer = pointlist.Count
        Dim pointrec As Point3d = pointlist(pointno - 1)
        dEndPt(0) = pointrec.x
        dEndPt(1) = pointrec.y
        dEndPt(2) = pointrec.z
        If Not WorkAreaErrorCheckXYZ(m_RunMode, dEndPt) Then
            CompileErrorDisplay(Arc.SheetName, Arc.CmdLineNo, 0)
            Return -1
        End If
        Dim ndgapZ As Double = dEndPt(2)
        Dim surfZ As Double = ndgapZ - Arc.Param.NeedleGap
        RetractZ = surfZ + Arc.Param.RetractHeight
        ClearanceZ = surfZ + Arc.Param.ClearanceHeight
        CheckRetractClearHeight(ndgapZ, RetractZ, ClearanceZ)
        If NotVisionMode(m_RunMode) Then
            If Not WorkAreaErrorCheckZ(RetractZ) Then
                CompileErrorDisplay(Arc.SheetName, Arc.CmdLineNo, RetractZError)
                Return -1
            End If
            If Not WorkAreaErrorCheckZ(ClearanceZ) Then
                CompileErrorDisplay(Arc.SheetName, Arc.CmdLineNo, ClearanceZError)
                Return -1
            End If
        End If

        RoundTo3DecimalPoints(RetractZ)
        RoundTo3DecimalPoints(ClearanceZ)
        '
        If (Arc.Param.ArcRadius < 0.001) Then
            m_ArcRadius = gMinArcRadius
        Else
            m_ArcRadius = Arc.Param.ArcRadius
        End If
        RoundTo3DecimalPoints(comp1)
        If GenerateStartBlock(comp1, ApproachZ, disCmdList) < 0 Then Return -1
        m_Speed = IDS.Data.Hardware.Gantry.ElementZSpeed       'set motion speed
        AddSpeed(disCmdList, m_Speed)
        If (NotVisionMode(m_RunMode)) Then
            RoundTo3DecimalPoints(ApproachZ)
            RoundTo3DecimalPoints(zDisp1)
            If IsXYPlane(m_Plane) Then
                pos(0) = comp1(0)
                pos(1) = comp1(1)
                pos(2) = ApproachZ
                AddMoveXYZ(disCmdList, pos)
                pos(0) = comp1(0)
                pos(1) = comp1(1)
                pos(2) = zDisp1
                AddMoveXYZ(disCmdList, pos)
            ElseIf IsYZPlane(m_Plane) Then
                pos(0) = comp1(1)
                pos(1) = ApproachZ
                pos(2) = comp1(0)
                AddMoveXYZ(disCmdList, pos)
                pos(0) = comp1(1)
                pos(1) = zDisp1
                pos(2) = comp1(0)
                AddMoveXYZ(disCmdList, pos)
            Else
                pos(0) = comp1(0)
                pos(1) = ApproachZ
                pos(2) = comp1(1)
                AddMoveXYZ(disCmdList, pos)
                pos(0) = comp1(0)
                pos(1) = zDisp1
                pos(2) = comp1(1)
                AddMoveXYZ(disCmdList, pos)
            End If

            AddWaitLoaded(disCmdList)
            If Arc.Param.DispenseOn And m_RunMode > 3 Then
                AddOnIO(disCmdList, m_NeedleIO)
            End If
        End If

        If Arc.Param.TravelDelay > 0 Then
            AddWaitDuration(disCmdList, Arc.Param.TravelDelay)
        End If

        If NotXYPlane(m_Plane) Then AddXYZBase(disCmdList)
        SetToXYPlane(m_Plane)

        m_Speed = Arc.Param.TravelSpeed
        AddSpeed(disCmdList, m_Speed)


        Dim count As Integer
        Dim tmp(3) As Double
        Dim isFirst As Boolean = True
        For count = 0 To pointno - 1 'move line segments 
            pointrec = pointlist(count)
            tmp(0) = pointrec.x
            tmp(1) = pointrec.y
            tmp(2) = pointrec.z
            RoundTo3DecimalPoints(tmp)
            If (NotVisionMode(m_RunMode)) Then
                pos(0) = tmp(0)
                pos(1) = tmp(1)
                pos(2) = tmp(2)
                AddMoveXYZ(disCmdList, pos)
            Else
                pos(0) = tmp(0)
                pos(1) = tmp(1)
                AddMoveXY(disCmdList, pos)
            End If

            If pointrec.DispenseOn = False And m_RunMode > 3 And isFirst Then
                isFirst = False
                If count = pointno - 1 Then
                    AddWaitIdle(disCmdList)
                    AddOffIO(disCmdList, m_NeedleIO)
                Else
                    AddWaitLoaded(disCmdList)
                    AddOffIO(disCmdList, m_NeedleIO)
                End If
            End If
        Next
        AddWaitIdle(disCmdList)

        If Arc.Param.RetractDelay > 0 Then
            AddWaitDuration(disCmdList, Arc.Param.RetractDelay)
        End If

        If (m_IsEndElement) Then
            Dim speed As Double = Arc.Param.RetractSpeed
            RoundTo3DecimalPoints(dEndPt)
            GenerateEndBlock(1, dEndPt, speed, RetractZ, ClearanceZ, disCmdList)
        End If

        m_PrevRetractPt(0) = dEndPt(0)
        m_PrevRetractPt(1) = dEndPt(1)
        m_PrevRetractPt(2) = RetractZ
        m_PrevClearPt(0) = dEndPt(0)
        m_PrevClearPt(1) = dEndPt(1)
        m_PrevClearPt(2) = ClearanceZ
        m_PrevArcRadius = m_ArcRadius
        m_PrevRetractSpeed = Arc.Param.RetractSpeed
        Return 0
    End Function

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Generate link arc motion command 
    '  disCmdList:    dispensing command list
    '
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Public Function GenerateLinkArcCommandSet(ByVal elemData As Object, ByRef disCmdList As ArrayList) As Integer
        '''''''''''''''''''''''''''''''''''
        Dim Arc As CIDSArc = elemData
        If Arc.CmdType <> "ARC" Then
            Return -1
        End If
        Dim CmdStr As String
        Dim p1(3), p2(3), p3(3), normal(3) As Double
        p1(0) = Arc.PosX1
        p1(1) = Arc.PosY1
        p1(2) = Arc.PosZ1
        p2(0) = Arc.PosX2
        p2(1) = Arc.PosY2
        p2(2) = Arc.PosZ2
        p3(0) = Arc.PosX3
        p3(1) = Arc.PosY3
        p3(2) = Arc.PosZ3
        If Get3dNormal(p1, p2, p3, normal) < 0 Then
            CompileErrorDisplay(Arc.SheetName, Arc.CmdLineNo, 12)
            Return -1
        End If
        If (Math.Abs(Math.Abs(normal(2)) - 1.0) < 0.001) Then
            If GenerateLinkArc2dCommandSet(elemData, disCmdList) < 0 Then Return -1
        Else
            If GenerateLinkArc3dCommandSet(elemData, disCmdList) < 0 Then Return -1
        End If
        If (m_IsFirstElementLink) Then
            m_IsFirstElementLink = False
        End If
        Return 0
    End Function

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Generate 2d link arc motion command 
    '  disCmdList:    dispensing command list
    '
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Public Function GenerateLinkArc2dCommandSet(ByVal elemData As Object, ByRef disCmdList As ArrayList) As Integer
        '''''''''''''''''''''''''''''''''''
        Dim Arc As CIDSArc = elemData
        Dim CmdStr As String
        If (m_IsFirstElementLink) Then
            m_LinkPara = Arc.Param
            m_LinkHeightCompS = Arc.HeightCompS
        End If
        Dim p1(3), p2(3), p3(3), comp1(3), comp2(3), comp3(3), detatchPt(3), dEndPt(3), midPt(3), off(3) As Double
        Dim ApproachZ, zDisp1, zDisp2, zDisp3, RetractZ, ClearanceZ As Double
        Dim centre(3), radius As Double
        Dim direction As Integer
        p1(0) = Arc.PosX1 '+ m_referencePt(0)   ' Camera pos
        p1(1) = Arc.PosY1 '+ m_referencePt(1)
        p1(2) = 0.0
        p2(0) = Arc.PosX2 '+ m_referencePt(0)   ' Camera pos
        p2(1) = Arc.PosY2 '+ m_referencePt(1)
        p2(2) = 0.0
        p3(0) = Arc.PosX3 '+ m_referencePt(0)   ' Camera pos
        p3(1) = Arc.PosY3 '+ m_referencePt(1)
        p3(2) = 0.0

        Dim needlename As String = m_LinkPara.Needle
        GetNeedleData(m_RunMode, needlename, off, m_NeedleIO)
        If NotVisionMode(m_RunMode) And m_NeedleIO <> m_PrevNeedleIO Then
            NeedleChange(m_NeedleIO, disCmdList)
            m_PrevNeedleIO = m_NeedleIO
        End If

        Dim fixHeight As Double = off(2) - m_LinkHeightCompS
        ApproachZ = Arc.PosZ1 + m_LinkPara.ApproachHeight + fixHeight
        zDisp1 = Arc.PosZ1 + m_LinkPara.NeedleGap + fixHeight
        zDisp2 = Arc.PosZ2 + m_LinkPara.NeedleGap + fixHeight
        zDisp3 = Arc.PosZ3 + m_LinkPara.NeedleGap + fixHeight
        CheckApproachHeight(zDisp1, ApproachZ)

        If (NotVisionMode(m_RunMode)) Then 'except vision mode
            comp1(0) = p1(0) - off(0)  ' needle Pos
            comp1(1) = p1(1) - off(1)
            comp1(2) = zDisp1
            comp2(0) = p2(0) - off(0)  ' needle Pos
            comp2(1) = p2(1) - off(1)
            comp2(2) = zDisp2
            comp3(0) = p3(0) - off(0)  ' needle Pos
            comp3(1) = p3(1) - off(1)
            comp3(2) = zDisp3
            If Not WorkAreaErrorCheckZ(ApproachZ) Then
                CompileErrorDisplay(Arc.SheetName, Arc.CmdLineNo, ApproachZError)
                Return -1
            End If
        Else
            comp1(0) = p1(0)
            comp1(1) = p1(1)
            comp1(2) = Arc.PosZ1
            comp2(0) = p2(0)
            comp2(1) = p2(1)
            comp2(2) = Arc.PosZ2
            comp3(0) = p3(0)
            comp3(1) = p3(1)
            comp3(2) = Arc.PosZ3
        End If
        If Not WorkAreaErrorCheckXYZ(m_RunMode, comp1) Then
            CompileErrorDisplay(Arc.SheetName, Arc.CmdLineNo, 0)
            Return -1
        End If
        If Not WorkAreaErrorCheckXYZ(m_RunMode, comp2) Then
            CompileErrorDisplay(Arc.SheetName, Arc.CmdLineNo, 0)
            Return -1
        End If
        If Not WorkAreaErrorCheckXYZ(m_RunMode, comp3) Then
            CompileErrorDisplay(Arc.SheetName, Arc.CmdLineNo, 0)
            Return -1
        End If

        Dim dc1(3), dc2(3), de1(3), de2(3), ddist As Double
        If (m_IsEndElementLink) Then
            ddist = m_LinkPara.DeTailDist
        Else
            ddist = 0.0
        End If

        If PointOnArc3d(comp1, comp2, comp3, ddist, centre, radius, direction, detatchPt) < 0 Then
            CompileErrorDisplay(Arc.SheetName, Arc.CmdLineNo, 11)
            Return -1
        End If
        If Not WorkAreaErrorCheckXYZ(m_RunMode, detatchPt) Then
            CompileErrorDisplay(Arc.SheetName, Arc.CmdLineNo, 0)
            Return -1
        End If

        If ddist > 0.001 Then
            detatchPt.CopyTo(dEndPt, 0)
            comp3.CopyTo(midPt, 0)
        ElseIf ddist < -0.001 Then
            comp3.CopyTo(dEndPt, 0)
            detatchPt.CopyTo(midPt, 0)
        Else
            comp3.CopyTo(midPt, 0)
            comp3.CopyTo(dEndPt, 0)
        End If

        dc1(0) = centre(0) - comp1(0)
        dc1(1) = centre(1) - comp1(1)
        dc1(2) = 0.0
        dc2(0) = centre(0) - midPt(0)
        dc2(1) = centre(1) - midPt(1)
        dc2(2) = 0.0

        de1(0) = midPt(0) - comp1(0)
        de1(1) = midPt(1) - comp1(1)
        de1(2) = 0.0
        de2(0) = dEndPt(0) - midPt(0)
        de2(1) = dEndPt(1) - midPt(1)
        de2(2) = 0.0
        RoundTo3DecimalPoints(dc1)
        RoundTo3DecimalPoints(dc2)
        RoundTo3DecimalPoints(de1)
        RoundTo3DecimalPoints(de2)
        RoundTo3DecimalPoints(comp1)

        Dim ndgapZ As Double = dEndPt(2)
        Dim surfZ As Double = ndgapZ - m_LinkPara.NeedleGap
        RetractZ = surfZ + m_LinkPara.RetractHeight
        ClearanceZ = surfZ + m_LinkPara.ClearanceHeight
        CheckRetractClearHeight(ndgapZ, RetractZ, ClearanceZ)

        If NotVisionMode(m_RunMode) Then
            If Not WorkAreaErrorCheckZ(RetractZ) Then
                CompileErrorDisplay(Arc.SheetName, Arc.CmdLineNo, RetractZError)
                Return -1
            End If
            If Not WorkAreaErrorCheckZ(ClearanceZ) Then
                CompileErrorDisplay(Arc.SheetName, Arc.CmdLineNo, ClearanceZError)
                Return -1
            End If
        End If

        RoundTo3DecimalPoints(RetractZ)
        RoundTo3DecimalPoints(ClearanceZ)

        '
        If (m_LinkPara.ArcRadius < 0.001) Then
            m_ArcRadius = gMinArcRadius
        Else
            m_ArcRadius = m_LinkPara.ArcRadius
        End If

        If (m_IsFirstElementLink) Then
            If GenerateStartBlock(comp1, ApproachZ, disCmdList) < 0 Then Return -1
        End If

        If NotVisionMode(m_RunMode) And m_IsFirstElementLink Then
            m_Speed = IDS.Data.Hardware.Gantry.ElementZSpeed       'set motion speed
            AddSpeed(disCmdList, m_Speed)
            RoundTo3DecimalPoints(ApproachZ)
            RoundTo3DecimalPoints(zDisp1)
            If IsXYPlane(m_Plane) Then
                pos(0) = comp1(0)
                pos(1) = comp1(1)
                pos(2) = ApproachZ
                AddMoveXYZ(disCmdList, pos)
                pos(0) = comp1(0)
                pos(1) = comp1(1)
                pos(2) = zDisp1
                AddMoveXYZ(disCmdList, pos)
            ElseIf IsYZPlane(m_Plane) Then
                pos(0) = comp1(1)
                pos(1) = ApproachZ
                pos(2) = comp1(0)
                AddMoveXYZ(disCmdList, pos)
                pos(0) = comp1(1)
                pos(1) = zDisp1
                pos(2) = comp1(0)
                AddMoveXYZ(disCmdList, pos)
            Else
                pos(0) = comp1(0)
                pos(1) = ApproachZ
                pos(2) = comp1(1)
                AddMoveXYZ(disCmdList, pos)
                pos(0) = comp1(0)
                pos(1) = zDisp1
                pos(2) = comp1(1)
                AddMoveXYZ(disCmdList, pos)
            End If

            AddWaitLoaded(disCmdList)
            If Arc.Param.DispenseOn And m_RunMode > 3 Then
                AddOnIO(disCmdList, m_NeedleIO)
            End If

        End If

        If m_LinkPara.TravelDelay > 0 And m_IsFirstElementLink Then
            AddWaitDuration(disCmdList, Arc.Param.TravelDelay)
        End If

        If m_IsFirstElementLink Then
            If NotXYPlane(m_Plane) Then AddXYZBase(disCmdList)
            SetToXYPlane(m_Plane)
            m_Speed = m_LinkPara.TravelSpeed
            AddSpeed(disCmdList, m_Speed)

        End If

        If (m_IsEndElementLink = False) Then
            m_Speed = Arc.Param.TravelSpeed
            AddSpeed(disCmdList, m_Speed)
            AddMoveArc(disCmdList, de1(0), de1(1), dc1(0), dc1(1), direction)
            If m_IsFirstElementLink = False Then
                AddWaitLoaded(disCmdList)

                If Arc.Param.DispenseOn And m_RunMode > 3 Then
                    AddWaitLoaded(disCmdList)
                    AddOnIO(disCmdList, m_NeedleIO)
                ElseIf Arc.Param.DispenseOn = False And m_RunMode > 3 Then
                    AddWaitLoaded(disCmdList)
                    AddOffIO(disCmdList, m_NeedleIO)
                End If
            End If

        Else
            If System.Math.Abs(ddist) <= 0.001 Then

                m_Speed = Arc.Param.TravelSpeed
                AddSpeed(disCmdList, m_Speed)
                AddMoveArc(disCmdList, de1(0), de1(1), dc1(0), dc1(1), direction)
                AddWaitLoaded(disCmdList)

                If Arc.Param.DispenseOn And m_RunMode > 3 Then
                    AddWaitLoaded(disCmdList)
                    AddOnIO(disCmdList, m_NeedleIO)
                ElseIf Arc.Param.DispenseOn = False And m_RunMode > 3 Then
                    AddWaitLoaded(disCmdList)
                    AddOffIO(disCmdList, m_NeedleIO)
                End If

                AddWaitIdle(disCmdList)
                If Arc.Param.DispenseOn And m_RunMode > 3 Then
                    AddOffIO(disCmdList, m_NeedleIO)
                End If
            Else
                m_Speed = Arc.Param.TravelSpeed
                AddSpeed(disCmdList, m_Speed)
                AddMoveArc(disCmdList, de1(0), de1(1), dc1(0), dc1(1), direction)
                AddWaitLoaded(disCmdList)

                If Arc.Param.DispenseOn And m_RunMode > 3 Then
                    AddWaitLoaded(disCmdList)
                    AddOnIO(disCmdList, m_NeedleIO)
                ElseIf Arc.Param.DispenseOn = False And m_RunMode > 3 Then
                    AddWaitLoaded(disCmdList)
                    AddOffIO(disCmdList, m_NeedleIO)
                End If

                AddMoveArc(disCmdList, de2(0), de2(1), dc2(0), dc2(1), direction)

                If Arc.Param.DispenseOn And m_RunMode > 3 Then
                    AddWaitLoaded(disCmdList)
                    AddOffIO(disCmdList, m_NeedleIO)
                End If
                AddWaitIdle(disCmdList)
            End If
        End If

        If m_LinkPara.RetractDelay > 0 And m_IsEndElementLink Then
            AddWaitDuration(disCmdList, Arc.Param.RetractDelay)
        End If

        If (m_IsEndElement) Then
            Dim speed As Double = Arc.Param.RetractSpeed
            RoundTo3DecimalPoints(dEndPt)
            GenerateEndBlock(1, dEndPt, speed, RetractZ, ClearanceZ, disCmdList)
        End If

        m_PrevRetractPt(0) = dEndPt(0)
        m_PrevRetractPt(1) = dEndPt(1)
        m_PrevRetractPt(2) = RetractZ
        m_PrevClearPt(0) = dEndPt(0)
        m_PrevClearPt(1) = dEndPt(1)
        m_PrevClearPt(2) = ClearanceZ
        m_PrevArcRadius = m_ArcRadius
        m_PrevRetractSpeed = Arc.Param.RetractSpeed
        Return 0
    End Function


    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Generate 3d link arc motion command 
    '  disCmdList:    dispensing command list
    '
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Public Function GenerateLinkArc3dCommandSet(ByVal elemData As Object, ByRef disCmdList As ArrayList) As Integer
        '''''''''''''''''''''''''''''''''''
        Dim Arc As CIDSArc = elemData
        Dim CmdStr As String
        If (m_IsFirstElementLink) Then
            m_LinkPara = Arc.Param
            m_LinkHeightCompS = Arc.HeightCompS
        End If
        Dim p1(3), p2(3), p3(3), comp1(3), comp2(3), comp3(3), detatchPt(3), dEndPt(3), midPt(3), off(3) As Double
        Dim ApproachZ, zDisp1, zDisp2, zDisp3, RetractZ, ClearanceZ As Double
        Dim centre(3), radius As Double
        Dim direction As Integer
        p1(0) = Arc.PosX1 '+ m_referencePt(0)   ' Camera pos
        p1(1) = Arc.PosY1 '+ m_referencePt(1)
        p1(2) = 0.0
        p2(0) = Arc.PosX2 '+ m_referencePt(0)   ' Camera pos
        p2(1) = Arc.PosY2 '+ m_referencePt(1)
        p2(2) = 0.0
        p3(0) = Arc.PosX3 '+ m_referencePt(0)   ' Camera pos
        p3(1) = Arc.PosY3 '+ m_referencePt(1)
        p3(2) = 0.0

        Dim needlename As String = m_LinkPara.Needle
        GetNeedleData(m_RunMode, needlename, off, m_NeedleIO)
        If NotVisionMode(m_RunMode) And m_NeedleIO <> m_PrevNeedleIO Then
            NeedleChange(m_NeedleIO, disCmdList)
            m_PrevNeedleIO = m_NeedleIO
        End If

        Dim fixHeight As Double = off(2) - m_LinkHeightCompS
        ApproachZ = Arc.PosZ1 + m_LinkPara.ApproachHeight + fixHeight
        zDisp1 = Arc.PosZ1 + m_LinkPara.NeedleGap + fixHeight
        zDisp2 = Arc.PosZ2 + m_LinkPara.NeedleGap + fixHeight
        zDisp3 = Arc.PosZ3 + m_LinkPara.NeedleGap + fixHeight
        CheckApproachHeight(zDisp1, ApproachZ)

        If (NotVisionMode(m_RunMode)) Then 'except vision mode
            comp1(0) = p1(0) - off(0)  ' needle Pos
            comp1(1) = p1(1) - off(1)
            comp1(2) = zDisp1
            comp2(0) = p2(0) - off(0)  ' needle Pos
            comp2(1) = p2(1) - off(1)
            comp2(2) = zDisp2
            comp3(0) = p3(0) - off(0)  ' needle Pos
            comp3(1) = p3(1) - off(1)
            comp3(2) = zDisp3
            If Not WorkAreaErrorCheckZ(ApproachZ) Then
                CompileErrorDisplay(Arc.SheetName, Arc.CmdLineNo, ApproachZError)
                Return -1
            End If
        Else
            comp1(0) = p1(0)
            comp1(1) = p1(1)
            comp1(2) = Arc.PosZ1
            comp2(0) = p2(0)
            comp2(1) = p2(1)
            comp2(2) = Arc.PosZ2
            comp3(0) = p3(0)
            comp3(1) = p3(1)
            comp3(2) = Arc.PosZ3
        End If
        If Not WorkAreaErrorCheckXYZ(m_RunMode, comp1) Then
            CompileErrorDisplay(Arc.SheetName, Arc.CmdLineNo, 0)
            Return -1
        End If
        If Not WorkAreaErrorCheckXYZ(m_RunMode, comp2) Then
            CompileErrorDisplay(Arc.SheetName, Arc.CmdLineNo, 0)
            Return -1
        End If
        If Not WorkAreaErrorCheckXYZ(m_RunMode, comp3) Then
            CompileErrorDisplay(Arc.SheetName, Arc.CmdLineNo, 0)
            Return -1
        End If

        Dim detaildist As Double
        If (m_IsEndElementLink) Then
            detaildist = m_LinkPara.DeTailDist
        Else
            detaildist = 0.0
        End If

        Dim pointlist As New ArrayList
        If GetArc3dPointList(comp1, comp2, comp3, detaildist, pointlist) < 0 Then
            CompileErrorDisplay(Arc.SheetName, Arc.CmdLineNo, 11)
            Return -1
        End If

        Dim pointno As Integer = pointlist.Count
        Dim pointrec As Point3d = pointlist(pointno - 1)
        dEndPt(0) = pointrec.x
        dEndPt(1) = pointrec.y
        dEndPt(2) = pointrec.z
        If Not WorkAreaErrorCheckXYZ(m_RunMode, dEndPt) Then
            CompileErrorDisplay(Arc.SheetName, Arc.CmdLineNo, 0)
            Return -1
        End If

        Dim ndgapZ As Double = dEndPt(2)
        Dim surfZ As Double = ndgapZ - m_LinkPara.NeedleGap
        RetractZ = surfZ + m_LinkPara.RetractHeight
        ClearanceZ = surfZ + m_LinkPara.ClearanceHeight
        CheckRetractClearHeight(ndgapZ, RetractZ, ClearanceZ)
        If NotVisionMode(m_RunMode) Then
            If Not WorkAreaErrorCheckZ(RetractZ) Then
                CompileErrorDisplay(Arc.SheetName, Arc.CmdLineNo, RetractZError)
                Return -1
            End If
            If Not WorkAreaErrorCheckZ(ClearanceZ) Then
                CompileErrorDisplay(Arc.SheetName, Arc.CmdLineNo, ClearanceZError)
                Return -1
            End If
        End If
        RoundTo3DecimalPoints(RetractZ)
        RoundTo3DecimalPoints(ClearanceZ)
        '
        If (Arc.Param.ArcRadius < 0.001) Then
            m_ArcRadius = gMinArcRadius
        Else
            m_ArcRadius = m_LinkPara.ArcRadius
        End If
        RoundTo3DecimalPoints(comp1)

        If (m_IsFirstElementLink) Then
            If GenerateStartBlock(comp1, ApproachZ, disCmdList) < 0 Then Return -1
        End If

        If (NotVisionMode(m_RunMode) And m_IsFirstElementLink) Then
            RoundTo3DecimalPoints(ApproachZ)
            RoundTo3DecimalPoints(zDisp1)
            If IsXYPlane(m_Plane) Then
                pos(0) = comp1(0)
                pos(1) = comp1(1)
                pos(2) = ApproachZ
                AddMoveXYZ(disCmdList, pos)
                pos(0) = comp1(0)
                pos(1) = comp1(1)
                pos(2) = zDisp1
                AddMoveXYZ(disCmdList, pos)
            ElseIf IsYZPlane(m_Plane) Then
                pos(0) = comp1(1)
                pos(1) = ApproachZ
                pos(2) = comp1(0)
                AddMoveXYZ(disCmdList, pos)
                pos(0) = comp1(1)
                pos(1) = zDisp1
                pos(2) = comp1(0)
                AddMoveXYZ(disCmdList, pos)
            Else
                pos(0) = comp1(0)
                pos(1) = ApproachZ
                pos(2) = comp1(1)
                AddMoveXYZ(disCmdList, pos)
                pos(0) = comp1(0)
                pos(1) = zDisp1
                pos(2) = comp1(1)
                AddMoveXYZ(disCmdList, pos)
            End If

            AddWaitLoaded(disCmdList)
            If Arc.Param.DispenseOn And m_RunMode > 3 Then
                AddOnIO(disCmdList, m_NeedleIO)
            End If

        End If

        If m_LinkPara.TravelDelay > 0 And m_IsFirstElementLink Then
            AddWaitDuration(disCmdList, Arc.Param.TravelDelay)
        End If

        If m_IsFirstElementLink Then
            If NotXYPlane(m_Plane) Then AddXYZBase(disCmdList)
            SetToXYPlane(m_Plane)
            m_Speed = m_LinkPara.TravelSpeed
            AddSpeed(disCmdList, m_Speed)
        End If

        Dim count As Integer
        Dim tmp(3) As Double
        Dim isFirst As Boolean = True
        For count = 0 To pointno - 1
            pointrec = pointlist(count)
            tmp(0) = pointrec.x
            tmp(1) = pointrec.y
            tmp(2) = pointrec.z
            RoundTo3DecimalPoints(tmp)
            If m_IsFirstElementLink = False And count = 0 Then

                m_Speed = Arc.Param.TravelSpeed
                AddSpeed(disCmdList, m_Speed)
            End If

            If (NotVisionMode(m_RunMode)) Then
                pos(0) = tmp(0)
                pos(1) = tmp(1)
                pos(2) = tmp(2)
                AddMoveXYZ(disCmdList, pos)
            Else
                pos(0) = tmp(0)
                pos(1) = tmp(1)
                AddMoveXY(disCmdList, pos)
            End If

            If m_RunMode > 3 And count = 0 And m_IsFirstElementLink = False Then
                If Arc.Param.DispenseOn Then
                    AddWaitLoaded(disCmdList)
                    AddOnIO(disCmdList, m_NeedleIO)
                ElseIf Arc.Param.DispenseOn = False Then
                    AddWaitLoaded(disCmdList)
                    AddOffIO(disCmdList, m_NeedleIO)
                End If
            End If

            If (m_IsEndElementLink) Then
                If pointrec.DispenseOn = False And m_RunMode > 3 And isFirst Then
                    isFirst = False
                    If count = pointno - 1 Then
                        AddWaitIdle(disCmdList)
                        AddOffIO(disCmdList, m_NeedleIO)
                    Else
                        AddWaitLoaded(disCmdList)
                        AddOffIO(disCmdList, m_NeedleIO)
                    End If
                End If
            End If
        Next
        AddWaitIdle(disCmdList)

        If m_LinkPara.RetractDelay > 0 And m_IsEndElementLink Then
            AddWaitDuration(disCmdList, Arc.Param.RetractDelay)
        End If

        If (m_IsEndElement) Then
            Dim speed As Double = Arc.Param.RetractSpeed
            RoundTo3DecimalPoints(dEndPt)
            GenerateEndBlock(1, dEndPt, speed, RetractZ, ClearanceZ, disCmdList)
        End If

        m_PrevRetractPt(0) = dEndPt(0)
        m_PrevRetractPt(1) = dEndPt(1)
        m_PrevRetractPt(2) = RetractZ
        m_PrevClearPt(0) = dEndPt(0)
        m_PrevClearPt(1) = dEndPt(1)
        m_PrevClearPt(2) = ClearanceZ
        m_PrevArcRadius = m_ArcRadius
        m_PrevRetractSpeed = Arc.Param.RetractSpeed
        Return 0
    End Function

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Generate circle motion command 
    '  disCmdList:    dispensing command list
    '
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Public Function GenerateCircleCommandSet(ByVal elemData As Object, ByRef disCmdList As ArrayList) As Integer
        '''''''''''''''''''''''''''''''''''
        Dim Circle As CIDSCircle = elemData
        If Circle.CmdType <> "CIRCLE" Then
            Return -1
        End If
        Dim CmdStr As String
        Dim p1(3), p2(3), p3(3), normal(3) As Double
        p1(0) = Circle.PosX1
        p1(1) = Circle.PosY1
        p1(2) = Circle.PosZ1
        p2(0) = Circle.PosX2
        p2(1) = Circle.PosY2
        p2(2) = Circle.PosZ2
        p3(0) = Circle.PosX3
        p3(1) = Circle.PosY3
        p3(2) = Circle.PosZ3
        If Get3dNormal(p1, p2, p3, normal) < 0 Then
            CompileErrorDisplay(Circle.SheetName, Circle.CmdLineNo, 12)
            Return -1
        End If

        If (Math.Abs(Math.Abs(normal(2)) - 1.0) < 0.001) Then
            If GenerateCircle2dCommandSet(elemData, disCmdList) < 0 Then Return -1
        Else
            If GenerateCircle3dCommandSet(elemData, disCmdList) < 0 Then Return -1
        End If

        Return 0
    End Function


    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Generate 2d circle motion command 
    '  disCmdList:    dispensing command list
    '
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Public Function GenerateCircle2dCommandSet(ByVal elemData As Object, ByRef disCmdList As ArrayList) As Integer
        '''''''''''''''''''''''''''''''''''
        Dim Circle As CIDSCircle = elemData

        Dim CmdStr As String

        Dim p1(3), p2(3), p3(3), comp1(3), comp2(3), comp3(3), detatchPt(3), dEndPt(3), midPt(3), off(3) As Double
        Dim ApproachZ, zDisp1, zDisp2, zDisp3, RetractZ, ClearanceZ As Double
        Dim centre(3), radius As Double
        Dim direction As Integer
        p1(0) = Circle.PosX1 ' + m_referencePt(0)   ' Camera pos
        p1(1) = Circle.PosY1 ' + m_referencePt(1)
        p1(2) = 0.0
        p2(0) = Circle.PosX2 '+ m_referencePt(0)   ' Camera pos
        p2(1) = Circle.PosY2 '+ m_referencePt(1)
        p2(2) = 0.0
        p3(0) = Circle.PosX3 '+ m_referencePt(0)   ' Camera pos
        p3(1) = Circle.PosY3 '+ m_referencePt(1)
        p3(2) = 0.0

        Dim needlename As String = Circle.Param.Needle
        GetNeedleData(m_RunMode, needlename, off, m_NeedleIO)
        If NotVisionMode(m_RunMode) And m_NeedleIO <> m_PrevNeedleIO Then
            NeedleChange(m_NeedleIO, disCmdList)
            m_PrevNeedleIO = m_NeedleIO
        End If

        Dim fixHeight As Double = off(2) - Circle.HeightCompS
        ApproachZ = Circle.PosZ1 + Circle.Param.ApproachHeight + fixHeight
        zDisp1 = Circle.PosZ1 + Circle.Param.NeedleGap + fixHeight
        zDisp2 = Circle.PosZ2 + Circle.Param.NeedleGap + fixHeight
        zDisp3 = Circle.PosZ3 + Circle.Param.NeedleGap + fixHeight
        CheckApproachHeight(zDisp1, ApproachZ)

        If (NotVisionMode(m_RunMode)) Then 'except vision mode
            comp1(0) = p1(0) - off(0)  ' needle Pos
            comp1(1) = p1(1) - off(1)
            comp1(2) = zDisp1
            comp2(0) = p2(0) - off(0)  ' needle Pos
            comp2(1) = p2(1) - off(1)
            comp2(2) = zDisp2
            comp3(0) = p3(0) - off(0)  ' needle Pos
            comp3(1) = p3(1) - off(1)
            comp3(2) = zDisp3
            If Not WorkAreaErrorCheckZ(ApproachZ) Then
                CompileErrorDisplay(Circle.SheetName, Circle.CmdLineNo, ApproachZError)
                Return -1
            End If
        Else
            comp1(0) = p1(0)
            comp1(1) = p1(1)
            comp1(2) = Circle.PosZ1
            comp2(0) = p2(0)
            comp2(1) = p2(1)
            comp2(2) = Circle.PosZ2
            comp3(0) = p3(0)
            comp3(1) = p3(1)
            comp3(2) = Circle.PosZ3
        End If
        If Not WorkAreaErrorCheckXYZ(m_RunMode, comp1) Then
            CompileErrorDisplay(Circle.SheetName, Circle.CmdLineNo, 0)
            Return -1
        End If
        If Not WorkAreaErrorCheckXYZ(m_RunMode, comp2) Then
            CompileErrorDisplay(Circle.SheetName, Circle.CmdLineNo, 0)
            Return -1
        End If
        If Not WorkAreaErrorCheckXYZ(m_RunMode, comp3) Then
            CompileErrorDisplay(Circle.SheetName, Circle.CmdLineNo, 0)
            Return -1
        End If

        If PointOnCircle3d(comp1, comp2, comp3, Circle.Param.DeTailDist, centre, radius, direction, detatchPt) < 0 Then
            CompileErrorDisplay(Circle.SheetName, Circle.CmdLineNo, 11)
            Return -1
        End If

        If Circle.Param.DeTailDist > 0.001 Then
            detatchPt.CopyTo(dEndPt, 0)
            comp1.CopyTo(midPt, 0)
        ElseIf Circle.Param.DeTailDist < -0.001 Then
            comp1.CopyTo(dEndPt, 0)
            detatchPt.CopyTo(midPt, 0)
        Else
            comp1.CopyTo(midPt, 0)
            comp1.CopyTo(dEndPt, 0)
        End If

        If Not WorkAreaErrorCheckXYZ(m_RunMode, dEndPt) Then
            CompileErrorDisplay(Circle.SheetName, Circle.CmdLineNo, 0)
            Return -1
        End If

        Dim dc1(3), dc2(3), de1(3), de2(3) As Double
        dc1(0) = centre(0) - comp1(0)
        dc1(1) = centre(1) - comp1(1)
        dc1(2) = 0.0
        dc2(0) = centre(0) - midPt(0)
        dc2(1) = centre(1) - midPt(1)
        dc2(2) = 0.0

        de1(0) = midPt(0) - comp1(0)
        de1(1) = midPt(1) - comp1(1)
        de1(2) = 0.0
        de2(0) = dEndPt(0) - midPt(0)
        de2(1) = dEndPt(1) - midPt(1)
        de2(2) = 0.0
        RoundTo3DecimalPoints(dc1)
        RoundTo3DecimalPoints(dc2)
        RoundTo3DecimalPoints(de1)
        RoundTo3DecimalPoints(de2)

        Dim ndgapZ As Double = dEndPt(2)
        Dim surfZ As Double = ndgapZ - Circle.Param.NeedleGap
        RetractZ = surfZ + Circle.Param.RetractHeight
        ClearanceZ = surfZ + Circle.Param.ClearanceHeight
        CheckRetractClearHeight(ndgapZ, RetractZ, ClearanceZ)
        If NotVisionMode(m_RunMode) Then
            If Not WorkAreaErrorCheckZ(RetractZ) Then
                CompileErrorDisplay(Circle.SheetName, Circle.CmdLineNo, RetractZError)
                Return -1
            End If
            If Not WorkAreaErrorCheckZ(ClearanceZ) Then
                CompileErrorDisplay(Circle.SheetName, Circle.CmdLineNo, ClearanceZError)
                Return -1
            End If
        End If

        RoundTo3DecimalPoints(RetractZ)
        RoundTo3DecimalPoints(ClearanceZ)
        '
        If (Circle.Param.ArcRadius < 0.001) Then
            m_ArcRadius = gMinArcRadius
        Else
            m_ArcRadius = Circle.Param.ArcRadius
        End If

        RoundTo3DecimalPoints(comp1)
        If GenerateStartBlock(comp1, ApproachZ, disCmdList) < 0 Then Return -1

        AddSpeed(disCmdList, IDS.Data.Hardware.Gantry.ElementZSpeed)
        If (NotVisionMode(m_RunMode)) Then
            RoundTo3DecimalPoints(ApproachZ)
            RoundTo3DecimalPoints(zDisp1)
            If IsXYPlane(m_Plane) Then
                pos(0) = comp1(0)
                pos(1) = comp1(1)
                pos(2) = ApproachZ
                AddMoveXYZ(disCmdList, pos)
                pos(0) = comp1(0)
                pos(1) = comp1(1)
                pos(2) = zDisp1
                AddMoveXYZ(disCmdList, pos)
            ElseIf IsYZPlane(m_Plane) Then
                pos(0) = comp1(1)
                pos(1) = ApproachZ
                pos(2) = comp1(0)
                AddMoveXYZ(disCmdList, pos)
                pos(0) = comp1(1)
                pos(1) = zDisp1
                pos(2) = comp1(0)
                AddMoveXYZ(disCmdList, pos)
            Else
                pos(0) = comp1(0)
                pos(1) = ApproachZ
                pos(2) = comp1(1)
                AddMoveXYZ(disCmdList, pos)
                pos(0) = comp1(0)
                pos(1) = zDisp1
                pos(2) = comp1(1)
                AddMoveXYZ(disCmdList, pos)
            End If

            AddWaitLoaded(disCmdList)
            If Circle.Param.DispenseOn And m_RunMode > 3 Then
                AddOnIO(disCmdList, m_NeedleIO)
            End If
        End If

        If Circle.Param.TravelDelay > 0 Then
            AddWaitDuration(disCmdList, Circle.Param.TravelDelay)
        End If

        If NotXYPlane(m_Plane) Then AddXYZBase(disCmdList)
        SetToXYPlane(m_Plane)

        m_Speed = Circle.Param.TravelSpeed
        AddSpeed(disCmdList, m_Speed)

        If System.Math.Abs(Circle.Param.DeTailDist) <= 0.001 Then
            AddMoveArc(disCmdList, de1(0), de1(1), dc1(0), dc1(1), direction)
            AddWaitIdle(disCmdList)
            If Circle.Param.DispenseOn And m_RunMode > 3 Then
                AddOffIO(disCmdList, m_NeedleIO)
            End If
        Else

            AddMoveArc(disCmdList, de1(0), de1(1), dc1(0), dc1(1), direction)
            AddMoveArc(disCmdList, de2(0), de2(1), dc2(0), dc2(1), direction)

            If Circle.Param.DispenseOn And m_RunMode > 3 Then
                AddWaitLoaded(disCmdList)
                AddOffIO(disCmdList, m_NeedleIO)
            End If
            AddWaitIdle(disCmdList)
        End If

        If Circle.Param.RetractDelay > 0 Then
            AddWaitDuration(disCmdList, Circle.Param.RetractDelay)
        End If

        If (m_IsEndElement) Then
            Dim speed As Double = Circle.Param.RetractSpeed
            RoundTo3DecimalPoints(dEndPt)
            GenerateEndBlock(1, dEndPt, speed, RetractZ, ClearanceZ, disCmdList)
        End If

        m_PrevRetractPt(0) = dEndPt(0)
        m_PrevRetractPt(1) = dEndPt(1)
        m_PrevRetractPt(2) = RetractZ
        m_PrevClearPt(0) = dEndPt(0)
        m_PrevClearPt(1) = dEndPt(1)
        m_PrevClearPt(2) = ClearanceZ
        m_PrevArcRadius = m_ArcRadius
        m_PrevRetractSpeed = Circle.Param.RetractSpeed
        Return 0
    End Function

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Generate 3d circle motion command 
    '  disCmdList:    dispensing command list
    '
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Public Function GenerateCircle3dCommandSet(ByVal elemData As Object, ByRef disCmdList As ArrayList) As Integer
        '''''''''''''''''''''''''''''''''''
        Dim Circle As CIDSCircle = elemData
        Dim CmdStr As String

        Dim p1(3), p2(3), p3(3), comp1(3), comp2(3), comp3(3), detatchPt(3), dEndPt(3), midPt(3), off(3) As Double
        Dim ApproachZ, zDisp1, zDisp2, zDisp3, RetractZ, ClearanceZ As Double
        Dim centre(3), radius As Double
        Dim direction As Integer
        p1(0) = Circle.PosX1 '+ m_referencePt(0)   ' Camera pos
        p1(1) = Circle.PosY1 '+ m_referencePt(1)
        p1(2) = 0.0
        p2(0) = Circle.PosX2 '+ m_referencePt(0)   ' Camera pos
        p2(1) = Circle.PosY2 '+ m_referencePt(1)
        p2(2) = 0.0
        p3(0) = Circle.PosX3 '+ m_referencePt(0)   ' Camera pos
        p3(1) = Circle.PosY3 '+ m_referencePt(1)
        p3(2) = 0.0

        Dim needlename As String = Circle.Param.Needle
        GetNeedleData(m_RunMode, needlename, off, m_NeedleIO)
        If NotVisionMode(m_RunMode) And m_NeedleIO <> m_PrevNeedleIO Then
            NeedleChange(m_NeedleIO, disCmdList)
            m_PrevNeedleIO = m_NeedleIO
        End If

        Dim fixHeight As Double = off(2) - Circle.HeightCompS
        ApproachZ = Circle.PosZ1 + Circle.Param.ApproachHeight + fixHeight
        zDisp1 = Circle.PosZ1 + Circle.Param.NeedleGap + fixHeight
        zDisp2 = Circle.PosZ2 + Circle.Param.NeedleGap + fixHeight
        zDisp3 = Circle.PosZ3 + Circle.Param.NeedleGap + fixHeight
        CheckApproachHeight(zDisp1, ApproachZ)

        If (NotVisionMode(m_RunMode)) Then 'except vision mode
            comp1(0) = p1(0) - off(0)  ' needle Pos
            comp1(1) = p1(1) - off(1)
            comp1(2) = zDisp1
            comp2(0) = p2(0) - off(0)  ' needle Pos
            comp2(1) = p2(1) - off(1)
            comp2(2) = zDisp2
            comp3(0) = p3(0) - off(0)  ' needle Pos
            comp3(1) = p3(1) - off(1)
            comp3(2) = zDisp3
            If Not WorkAreaErrorCheckZ(ApproachZ) Then
                CompileErrorDisplay(Circle.SheetName, Circle.CmdLineNo, ApproachZError)
                Return -1
            End If
        Else
            comp1(0) = p1(0)
            comp1(1) = p1(1)
            comp1(2) = Circle.PosZ1
            comp2(0) = p2(0)
            comp2(1) = p2(1)
            comp2(2) = Circle.PosZ2
            comp3(0) = p3(0)
            comp3(1) = p3(1)
            comp3(2) = Circle.PosZ3
        End If
        If Not WorkAreaErrorCheckXYZ(m_RunMode, comp1) Then
            CompileErrorDisplay(Circle.SheetName, Circle.CmdLineNo, 0)
            Return -1
        End If
        If Not WorkAreaErrorCheckXYZ(m_RunMode, comp2) Then
            CompileErrorDisplay(Circle.SheetName, Circle.CmdLineNo, 0)
            Return -1
        End If
        If Not WorkAreaErrorCheckXYZ(m_RunMode, comp3) Then
            CompileErrorDisplay(Circle.SheetName, Circle.CmdLineNo, 0)
            Return -1
        End If

        Dim detaildist As Double = Circle.Param.DeTailDist
        Dim pointlist As New ArrayList

        If GetCircle3dPointList(comp1, comp2, comp3, detaildist, pointlist) < 0 Then
            CompileErrorDisplay(Circle.SheetName, Circle.CmdLineNo, 11)
            Return -1
        End If

        Dim pointno As Integer = pointlist.Count
        Dim pointrec As Point3d = pointlist(pointno - 1)
        dEndPt(0) = pointrec.x
        dEndPt(1) = pointrec.y
        dEndPt(2) = pointrec.z
        If Not WorkAreaErrorCheckXYZ(m_RunMode, dEndPt) Then
            CompileErrorDisplay(Circle.SheetName, Circle.CmdLineNo, 0)
            Return -1
        End If
        Dim ndgapZ As Double = dEndPt(2)
        Dim surfZ As Double = ndgapZ - Circle.Param.NeedleGap
        RetractZ = surfZ + Circle.Param.RetractHeight
        ClearanceZ = surfZ + Circle.Param.ClearanceHeight
        CheckRetractClearHeight(ndgapZ, RetractZ, ClearanceZ)
        If NotVisionMode(m_RunMode) Then
            If Not WorkAreaErrorCheckZ(RetractZ) Then
                CompileErrorDisplay(Circle.SheetName, Circle.CmdLineNo, RetractZError)
                Return -1
            End If
            If Not WorkAreaErrorCheckZ(ClearanceZ) Then
                CompileErrorDisplay(Circle.SheetName, Circle.CmdLineNo, ClearanceZError)
                Return -1
            End If
        End If
        RoundTo3DecimalPoints(RetractZ)
        RoundTo3DecimalPoints(ClearanceZ)
        '
        If (Circle.Param.ArcRadius < 0.001) Then
            m_ArcRadius = gMinArcRadius
        Else
            m_ArcRadius = Circle.Param.ArcRadius
        End If

        RoundTo3DecimalPoints(comp1)
        If GenerateStartBlock(comp1, ApproachZ, disCmdList) < 0 Then Return -1

        AddSpeed(disCmdList, IDS.Data.Hardware.Gantry.ElementZSpeed)

        If (NotVisionMode(m_RunMode)) Then
            RoundTo3DecimalPoints(ApproachZ)
            RoundTo3DecimalPoints(zDisp1)
            If IsXYPlane(m_Plane) Then
                pos(0) = comp1(0)
                pos(1) = comp1(1)
                pos(2) = ApproachZ
                AddMoveXYZ(disCmdList, pos)
                pos(0) = comp1(0)
                pos(1) = comp1(1)
                pos(2) = zDisp1
                AddMoveXYZ(disCmdList, pos)
            ElseIf IsYZPlane(m_Plane) Then
                pos(0) = comp1(1)
                pos(1) = ApproachZ
                pos(2) = comp1(0)
                AddMoveXYZ(disCmdList, pos)
                pos(0) = comp1(1)
                pos(1) = zDisp1
                pos(2) = comp1(0)
                AddMoveXYZ(disCmdList, pos)
            Else
                pos(0) = comp1(0)
                pos(1) = ApproachZ
                pos(2) = comp1(1)
                AddMoveXYZ(disCmdList, pos)
                pos(0) = comp1(0)
                pos(1) = zDisp1
                pos(2) = comp1(1)
                AddMoveXYZ(disCmdList, pos)
            End If

            AddWaitLoaded(disCmdList)
            If Circle.Param.DispenseOn And m_RunMode > 3 Then
                AddOnIO(disCmdList, m_NeedleIO)
            End If
        End If

        If Circle.Param.TravelDelay > 0 Then
            AddWaitDuration(disCmdList, Circle.Param.TravelDelay)
        End If

        If NotXYPlane(m_Plane) Then AddXYZBase(disCmdList)
        SetToXYPlane(m_Plane)

        m_Speed = Circle.Param.TravelSpeed
        AddSpeed(disCmdList, m_Speed)

        Dim count As Integer
        Dim tmp(3) As Double
        Dim isFirst As Boolean = True
        For count = 0 To pointno - 1
            pointrec = pointlist(count)
            tmp(0) = pointrec.x
            tmp(1) = pointrec.y
            tmp(2) = pointrec.z
            RoundTo3DecimalPoints(tmp)
            If (NotVisionMode(m_RunMode)) Then
                pos(0) = tmp(0)
                pos(1) = tmp(1)
                pos(2) = tmp(2)
                AddMoveXYZ(disCmdList, pos)
            Else
                pos(0) = tmp(0)
                pos(1) = tmp(1)
                AddMoveXY(disCmdList, pos)
            End If

            If pointrec.DispenseOn = False And m_RunMode > 3 And isFirst Then
                isFirst = False
                If count = pointno - 1 Then
                    AddWaitIdle(disCmdList)
                    AddOffIO(disCmdList, m_NeedleIO)
                Else
                    AddWaitLoaded(disCmdList)
                    AddOffIO(disCmdList, m_NeedleIO)
                End If
            End If
        Next
        AddWaitIdle(disCmdList)

        If Circle.Param.RetractDelay > 0 Then
            AddWaitDuration(disCmdList, Circle.Param.RetractDelay)
        End If

        If (m_IsEndElement) Then
            Dim speed As Double = Circle.Param.RetractSpeed
            RoundTo3DecimalPoints(dEndPt)
            GenerateEndBlock(1, dEndPt, speed, RetractZ, ClearanceZ, disCmdList)
        End If

        m_PrevRetractPt(0) = dEndPt(0)
        m_PrevRetractPt(1) = dEndPt(1)
        m_PrevRetractPt(2) = RetractZ
        m_PrevClearPt(0) = dEndPt(0)
        m_PrevClearPt(1) = dEndPt(1)
        m_PrevClearPt(2) = ClearanceZ
        m_PrevArcRadius = m_ArcRadius
        m_PrevRetractSpeed = Circle.Param.RetractSpeed
        Return 0
    End Function

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Generate rectangle motion command 
    '  disCmdList:    dispensing command list
    '
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Public Function GenerateRoundTo3DecimalPointsRectCommandSet(ByVal elemData As Object, ByRef disCmdList As ArrayList) As Integer
        '''''''''''''''''''''''''''''''''''
        Dim Rect As CIDSRectangle = elemData
        If Rect.CmdType <> "RECTANGLE" Then
            Return -1
        End If
        Dim CmdStr As String

        Dim p1(3), p2(3), p3(3), p4(3), comp1(3), comp2(3), comp3(3), comp4(3), detatchPt(3), dEndPt(3), midPt(3), off(3) As Double
        Dim ApproachZ, zDisp1, zDisp2, zDisp3, RetractZ, ClearanceZ As Double

        Dim radius, xc, yc, startAngle, arcAngle, a, b, c, d As Double
        Dim dc(3), de(3), dz As Double
        Dim i, direction As Integer
        p1(0) = Rect.PosX1 '+ m_referencePt(0)   ' Camera pos
        p1(1) = Rect.PosY1 '+ m_referencePt(1)
        p1(2) = 0.0
        p2(0) = Rect.PosX2 '+ m_referencePt(0)   ' Camera pos
        p2(1) = Rect.PosY2 '+ m_referencePt(1)
        p2(2) = 0.0
        p3(0) = Rect.PosX3 '+ m_referencePt(0)   ' Camera pos
        p3(1) = Rect.PosY3 '+ m_referencePt(1)
        p3(2) = 0.0

        Dim needlename As String = Rect.Param.Needle
        GetNeedleData(m_RunMode, needlename, off, m_NeedleIO)
        If NotVisionMode(m_RunMode) And m_NeedleIO <> m_PrevNeedleIO Then
            NeedleChange(m_NeedleIO, disCmdList)
            m_PrevNeedleIO = m_NeedleIO
        End If

        Dim fixHeight As Double = off(2) - Rect.HeightCompS
        ApproachZ = Rect.PosZ1 + Rect.Param.ApproachHeight + fixHeight
        zDisp1 = Rect.PosZ1 + Rect.Param.NeedleGap + fixHeight
        zDisp2 = Rect.PosZ2 + Rect.Param.NeedleGap + fixHeight
        zDisp3 = Rect.PosZ3 + Rect.Param.NeedleGap + fixHeight
        CheckApproachHeight(zDisp1, ApproachZ)

        If (NotVisionMode(m_RunMode)) Then 'except vision mode
            p1(0) = p1(0) - off(0)  ' needle Pos
            p1(1) = p1(1) - off(1)
            p1(2) = zDisp1
            p2(0) = p2(0) - off(0)  ' needle Pos
            p2(1) = p2(1) - off(1)
            p2(2) = zDisp2
            p3(0) = p3(0) - off(0)  ' needle Pos
            p3(1) = p3(1) - off(1)
            p3(2) = zDisp3
            If Not WorkAreaErrorCheckZ(ApproachZ) Then
                CompileErrorDisplay(Rect.SheetName, Rect.CmdLineNo, ApproachZError)
                Return -1
            End If
        Else
            p1(2) = Rect.PosZ1
            p2(2) = Rect.PosZ2
            p3(2) = Rect.PosZ3
        End If
        If Not WorkAreaErrorCheckXYZ(m_RunMode, p1) Then
            CompileErrorDisplay(Rect.SheetName, Rect.CmdLineNo, 0)
            Return -1
        End If
        If Not WorkAreaErrorCheckXYZ(m_RunMode, p2) Then
            CompileErrorDisplay(Rect.SheetName, Rect.CmdLineNo, 0)
            Return -1
        End If
        If Not WorkAreaErrorCheckXYZ(m_RunMode, p3) Then
            CompileErrorDisplay(Rect.SheetName, Rect.CmdLineNo, 0)
            Return -1
        End If

        Dim ptlist As New ArrayList
        GetDiaVertexlist(p1, p2, p3, ptlist)  'get rectangle vertex list
        If ptlist.Count <> 5 Then Return -1

        ptlist(3).CopyTo(p4, 0)

        If Not WorkAreaErrorCheckXYZ(m_RunMode, p4) Then
            CompileErrorDisplay(Rect.SheetName, Rect.CmdLineNo, 0)
            Return -1
        End If
        p1.CopyTo(comp1, 0)

        If (Rect.Param.ArcRadius < 0.001) Then
            m_ArcRadius = gMinArcRadius
        Else
            m_ArcRadius = Rect.Param.ArcRadius
        End If

        RoundTo3DecimalPoints(comp1)
        If GenerateStartBlock(comp1, ApproachZ, disCmdList) < 0 Then Return -1
        AddSpeed(disCmdList, IDS.Data.Hardware.Gantry.ElementZSpeed)
        If (NotVisionMode(m_RunMode)) Then 'except vision mode
            RoundTo3DecimalPoints(ApproachZ)
            RoundTo3DecimalPoints(zDisp1)
            If IsXYPlane(m_Plane) Then
                pos(0) = comp1(0)
                pos(1) = comp1(1)
                pos(2) = ApproachZ
                AddMoveXYZ(disCmdList, pos)
                pos(0) = comp1(0)
                pos(1) = comp1(1)
                pos(2) = zDisp1
                AddMoveXYZ(disCmdList, pos)
            ElseIf IsYZPlane(m_Plane) Then
                pos(0) = comp1(1)
                pos(1) = ApproachZ
                pos(2) = comp1(0)
                AddMoveXYZ(disCmdList, pos)
                pos(0) = comp1(1)
                pos(1) = zDisp1
                pos(2) = comp1(0)
                AddMoveXYZ(disCmdList, pos)
            Else
                pos(0) = comp1(0)
                pos(1) = ApproachZ
                pos(2) = comp1(1)
                AddMoveXYZ(disCmdList, pos)
                pos(0) = comp1(0)
                pos(1) = zDisp1
                pos(2) = comp1(1)
                AddMoveXYZ(disCmdList, pos)
            End If

            AddWaitLoaded(disCmdList)
            If Rect.Param.DispenseOn And m_RunMode > 3 Then
                AddOnIO(disCmdList, m_NeedleIO)
            End If
        End If

        If Rect.Param.TravelDelay > 0 Then
            AddWaitDuration(disCmdList, Rect.Param.TravelDelay)
        End If

        AddXYZBase(disCmdList)
        SetToXYPlane(m_Plane)

        m_Speed = Rect.Param.TravelSpeed
        AddSpeed(disCmdList, m_Speed)

        radius = GetFilletRadius(m_Speed)
        Planecoef(p1, p2, p3, a, b, c, d) 'rectangle's plane

        Dim rtn As Integer
        'insert fillet radius to smoothing motion
        For i = 0 To 2
            ptlist(i).CopyTo(comp1, 0)
            ptlist(i + 1).CopyTo(comp2, 0)
            ptlist(i + 1).CopyTo(comp3, 0)
            ptlist(i + 2).CopyTo(comp4, 0)
            rtn = Linelinefillet(comp1, comp2, comp3, comp4, radius, xc, yc, startAngle, arcAngle, direction)
            If rtn = 0 Then
                If c <> 0.0 Then
                    comp2(2) = -(a * comp2(0) + b * comp2(1) + d) / c
                    comp3(2) = -(a * comp3(0) + b * comp3(1) + d) / c
                    dz = comp3(2) - comp2(2)
                Else
                    comp2(2) = comp1(2)
                    comp3(2) = comp1(2)
                    dz = 0.0
                End If

                dc(0) = xc - comp2(0)
                dc(1) = yc - comp2(1)
                dc(2) = 0.0

                de(0) = comp3(0) - comp2(0)
                de(1) = comp3(1) - comp2(1)
                de(2) = 0.0
                RoundTo3DecimalPoints(dc)
                RoundTo3DecimalPoints(de)
                RoundTo3DecimalPoints(comp2)
                pos(0) = comp2(0)
                pos(1) = comp2(1)
                pos(2) = comp2(2)
                AddMoveXYZ(disCmdList, pos)

                If Math.Abs(dz) < 0.001 Then  'or m_RunMode = 0
                    RoundTo3DecimalPoints(dz)
                    'AddSpeed(disCmdList, m_Speed * 3)
                    AddSpeed(disCmdList, m_Speed)
                    AddMoveArc(disCmdList, de(0), de(1), dc(0), dc(1), direction)
                    AddSpeed(disCmdList, m_Speed)
                Else
                    RoundTo3DecimalPoints(dz)
                    'AddSpeed(disCmdList, m_Speed * 3)
                    AddSpeed(disCmdList, m_Speed)
                    AddMoveHelical(disCmdList, de(0), de(1), dc(0), dc(1), direction, dz)
                    AddSpeed(disCmdList, m_Speed)
                End If
            Else
                RoundTo3DecimalPoints(comp2)
                pos(0) = comp2(0)
                pos(1) = comp2(1)
                pos(2) = comp2(2)
                AddMoveXYZ(disCmdList, pos)
                AddWaitIdle(disCmdList)
            End If
        Next i
        Dim ndgapZ, surfZ As Double
        'detailing operation
        If System.Math.Abs(Rect.Param.DeTailDist) <= 0.001 Then
            ptlist(0).CopyTo(comp1, 0)
            ndgapZ = comp1(2)
            surfZ = ndgapZ - Rect.Param.NeedleGap
            RetractZ = surfZ + Rect.Param.RetractHeight
            ClearanceZ = surfZ + Rect.Param.ClearanceHeight
            CheckRetractClearHeight(ndgapZ, RetractZ, ClearanceZ)
            If NotVisionMode(m_RunMode) Then
                If Not WorkAreaErrorCheckZ(RetractZ) Then
                    CompileErrorDisplay(Rect.SheetName, Rect.CmdLineNo, RetractZError)
                    Return -1
                End If
                If Not WorkAreaErrorCheckZ(ClearanceZ) Then
                    CompileErrorDisplay(Rect.SheetName, Rect.CmdLineNo, ClearanceZError)
                    Return -1
                End If
            End If

            RoundTo3DecimalPoints(comp1)
            comp1.CopyTo(dEndPt, 0)
            RoundTo3DecimalPoints(RetractZ)
            RoundTo3DecimalPoints(ClearanceZ)

            pos(0) = comp1(0)
            pos(1) = comp1(1)
            pos(2) = comp1(2)
            AddMoveXYZ(disCmdList, pos)
            AddWaitIdle(disCmdList)

            If Rect.Param.DispenseOn And m_RunMode > 3 Then
                AddOffIO(disCmdList, m_NeedleIO)
            End If
            If Rect.Param.RetractDelay > 0 Then
                AddWaitDuration(disCmdList, Rect.Param.RetractDelay)
            End If

        ElseIf Rect.Param.DeTailDist < -0.001 Then
            ptlist(3).CopyTo(comp1, 0)
            ptlist(4).CopyTo(comp2, 0)
            If PointOnLine3d(comp1, comp2, Rect.Param.DeTailDist, detatchPt) < 0 Then
                CompileErrorDisplay(Rect.SheetName, Rect.CmdLineNo, 11)
                Return -1
            End If

            ndgapZ = comp2(2)
            surfZ = ndgapZ - Rect.Param.NeedleGap
            RetractZ = surfZ + Rect.Param.RetractHeight
            ClearanceZ = surfZ + Rect.Param.ClearanceHeight
            CheckRetractClearHeight(ndgapZ, RetractZ, ClearanceZ)
            If NotVisionMode(m_RunMode) Then
                If Not WorkAreaErrorCheckZ(RetractZ) Then
                    CompileErrorDisplay(Rect.SheetName, Rect.CmdLineNo, RetractZError)
                    Return -1
                End If
                If Not WorkAreaErrorCheckZ(ClearanceZ) Then
                    CompileErrorDisplay(Rect.SheetName, Rect.CmdLineNo, ClearanceZError)
                    Return -1
                End If
            End If
            RoundTo3DecimalPoints(comp2)
            RoundTo3DecimalPoints(detatchPt)
            comp2.CopyTo(dEndPt, 0)
            RoundTo3DecimalPoints(RetractZ)
            RoundTo3DecimalPoints(ClearanceZ)
            pos(0) = detatchPt(0)
            pos(1) = detatchPt(1)
            pos(2) = detatchPt(2)
            AddMoveXYZ(disCmdList, pos)
            pos(0) = comp2(0)
            pos(1) = comp2(1)
            pos(2) = comp2(2)
            AddMoveXYZ(disCmdList, pos)
            If Rect.Param.DispenseOn And m_RunMode > 3 Then
                AddWaitLoaded(disCmdList)
                AddOffIO(disCmdList, m_NeedleIO)
            End If
            AddWaitIdle(disCmdList)

            If Rect.Param.RetractDelay > 0 Then
                AddWaitDuration(disCmdList, Rect.Param.RetractDelay)
            End If

        Else
            ptlist(1).CopyTo(comp1, 0)
            ptlist(0).CopyTo(comp2, 0)
            If PointOnLine3d(comp1, comp2, -Rect.Param.DeTailDist, detatchPt) < 0 Then
                CompileErrorDisplay(Rect.SheetName, Rect.CmdLineNo, 11)
                Return -1
            End If

            ndgapZ = detatchPt(2)
            surfZ = ndgapZ - Rect.Param.NeedleGap
            RetractZ = surfZ + Rect.Param.RetractHeight
            ClearanceZ = surfZ + Rect.Param.ClearanceHeight
            CheckRetractClearHeight(ndgapZ, RetractZ, ClearanceZ)
            If NotVisionMode(m_RunMode) Then
                If Not WorkAreaErrorCheckZ(RetractZ) Then
                    CompileErrorDisplay(Rect.SheetName, Rect.CmdLineNo, RetractZError)
                    Return -1
                End If
                If Not WorkAreaErrorCheckZ(ClearanceZ) Then
                    CompileErrorDisplay(Rect.SheetName, Rect.CmdLineNo, ClearanceZError)
                    Return -1
                End If
            End If

            RoundTo3DecimalPoints(comp2)
            RoundTo3DecimalPoints(detatchPt)
            detatchPt.CopyTo(dEndPt, 0)
            RoundTo3DecimalPoints(RetractZ)
            RoundTo3DecimalPoints(ClearanceZ)
            pos(0) = comp2(0)
            pos(1) = comp2(1)
            pos(2) = comp2(2)
            AddMoveXYZ(disCmdList, pos)
            AddWaitIdle(disCmdList)
            pos(0) = detatchPt(0)
            pos(1) = detatchPt(1)
            pos(2) = detatchPt(2)
            AddMoveXYZ(disCmdList, pos)
            If Rect.Param.DispenseOn And m_RunMode > 3 Then
                AddWaitLoaded(disCmdList)
                AddOffIO(disCmdList, m_NeedleIO)
            End If
            AddWaitIdle(disCmdList)

            If Rect.Param.RetractDelay > 0 Then
                AddWaitDuration(disCmdList, Rect.Param.RetractDelay)
            End If

        End If

        If Not WorkAreaErrorCheckXYZ(m_RunMode, dEndPt) Then
            CompileErrorDisplay(Rect.SheetName, Rect.CmdLineNo, 0)
            Return -1
        End If

        If (m_IsEndElement) Then
            Dim speed As Double = Rect.Param.RetractSpeed
            RoundTo3DecimalPoints(dEndPt)
            GenerateEndBlock(1, dEndPt, speed, RetractZ, ClearanceZ, disCmdList)
        End If

        m_PrevRetractPt(0) = dEndPt(0)
        m_PrevRetractPt(1) = dEndPt(1)
        m_PrevRetractPt(2) = RetractZ
        m_PrevClearPt(0) = dEndPt(0)
        m_PrevClearPt(1) = dEndPt(1)
        m_PrevClearPt(2) = ClearanceZ
        m_PrevArcRadius = m_ArcRadius
        m_PrevRetractSpeed = Rect.Param.RetractSpeed

        Return 0
    End Function

    Public Function GenerateFillCircleCommandSet(ByVal elemData As Object, ByRef disCmdList As ArrayList) As Integer
        '''''''''''''''''''''''''''''''''''
        Dim FillCircle As CIDSFillCircle = elemData
        If FillCircle.CmdType <> "FILLCIRCLE" Then
            Return -1
        End If
        Dim CmdStr As String

        Dim p1(3), p2(3), p3(3), comp1(3), comp2(3), comp3(3), detatchPt(3), dEndPt(3), midPt(3), off(3) As Double
        Dim zApht, zDisp1, zDisp2, zDisp3, zRetr, zClear As Double
        Dim centre(3), radius As Double
        Dim direction As Integer
        p1(0) = FillCircle.PosX1 '+ m_referencePt(0)   ' Camera pos
        p1(1) = FillCircle.PosY1 '+ m_referencePt(1)
        p1(2) = 0.0
        p2(0) = FillCircle.PosX2 '+ m_referencePt(0)   ' Camera pos
        p2(1) = FillCircle.PosY2 '+ m_referencePt(1)
        p2(2) = 0.0
        p3(0) = FillCircle.PosX3 '+ m_referencePt(0)   ' Camera pos
        p3(1) = FillCircle.PosY3 '+ m_referencePt(1)
        p3(2) = 0.0

        Dim needlename As String = FillCircle.Param.Needle
        GetNeedleData(m_RunMode, needlename, off, m_NeedleIO)
        If NotVisionMode(m_RunMode) And m_NeedleIO <> m_PrevNeedleIO Then
            NeedleChange(m_NeedleIO, disCmdList)
            m_PrevNeedleIO = m_NeedleIO
        End If

        Dim fixHeight As Double = off(2) - FillCircle.HeightCompS
        zApht = FillCircle.PosZ1 + FillCircle.Param.ApproachHeight + fixHeight
        zDisp1 = FillCircle.PosZ1 + FillCircle.Param.NeedleGap + fixHeight
        zDisp2 = FillCircle.PosZ2 + FillCircle.Param.NeedleGap + fixHeight
        zDisp3 = FillCircle.PosZ3 + FillCircle.Param.NeedleGap + fixHeight
        CheckApproachHeight(zDisp1, zApht)

        If (NotVisionMode(m_RunMode)) Then 'except vision mode
            comp1(0) = p1(0) - off(0)  ' needle Pos
            comp1(1) = p1(1) - off(1)
            comp1(2) = zDisp1
            comp2(0) = p2(0) - off(0)  ' needle Pos
            comp2(1) = p2(1) - off(1)
            comp2(2) = zDisp2
            comp3(0) = p3(0) - off(0)  ' needle Pos
            comp3(1) = p3(1) - off(1)
            comp3(2) = zDisp3
            If Not WorkAreaErrorCheckZ(zApht) Then
                CompileErrorDisplay(FillCircle.SheetName, FillCircle.CmdLineNo, ApproachZError)
                Return -1
            End If
        Else
            comp1(0) = p1(0)
            comp1(1) = p1(1)
            comp1(2) = FillCircle.PosZ1
            comp2(0) = p2(0)
            comp2(1) = p2(1)
            comp2(2) = FillCircle.PosZ2
            comp3(0) = p3(0)
            comp3(1) = p3(1)
            comp3(2) = FillCircle.PosZ3
        End If

        If Not WorkAreaErrorCheckXYZ(m_RunMode, comp1) Then
            CompileErrorDisplay(FillCircle.SheetName, FillCircle.CmdLineNo, 0)
            Return -1
        End If
        If Not WorkAreaErrorCheckXYZ(m_RunMode, comp2) Then
            CompileErrorDisplay(FillCircle.SheetName, FillCircle.CmdLineNo, 0)
            Return -1
        End If
        If Not WorkAreaErrorCheckXYZ(m_RunMode, comp3) Then
            CompileErrorDisplay(FillCircle.SheetName, FillCircle.CmdLineNo, 0)
            Return -1
        End If

        Dim r_pitch As Double = FillCircle.Param.Pitch
        'Dim rr_pitch As Double = 0.0
        If (r_pitch = 0) Then
            MsgBox("Pitch cannot be set to zero for Filled Circle.")
            Return -1
        End If
        Dim height As Double = FillCircle.Param.FillHeight
        Dim NoOfRun As Integer = FillCircle.Param.NoOfRun
        Dim inout As String = FillCircle.Param.SprialIn
        Dim detaildist As Double = FillCircle.Param.DeTailDist
        Dim pointlist As New ArrayList
        'generate spiral point list to fit spiral
        If GetSpiralPointList(comp1, comp2, comp3, r_pitch, NoOfRun, height, detaildist, inout, pointlist) < 0 Then
            CompileErrorDisplay(FillCircle.SheetName, FillCircle.CmdLineNo, 13)
            Return -1
        End If

        Dim pointno As Integer = pointlist.Count
        If pointno <= 0 Then Return -1
        Dim endptrec As Point3d = pointlist(pointno - 1)
        Dim startptrec As Point3d = pointlist(0)
        Dim pointrec As Point3d
        If inout.ToUpper = "OUT" Then
            comp1(0) = startptrec.x
            comp1(1) = startptrec.y
            comp1(2) = startptrec.z
            zApht = startptrec.z - FillCircle.Param.NeedleGap + FillCircle.Param.ApproachHeight
            zDisp1 = startptrec.z
            CheckApproachHeight(zDisp1, zApht)
            If NotVisionMode(m_RunMode) Then
                If Not WorkAreaErrorCheckZ(zDisp1) Then
                    CompileErrorDisplay(FillCircle.SheetName, FillCircle.CmdLineNo, 0)
                    Return -1
                End If
                If Not WorkAreaErrorCheckZ(zApht) Then
                    CompileErrorDisplay(FillCircle.SheetName, FillCircle.CmdLineNo, ApproachZError)
                    Return -1
                End If
            End If
        End If

        dEndPt(0) = endptrec.x
        dEndPt(1) = endptrec.y
        dEndPt(2) = endptrec.z
        If Not WorkAreaErrorCheckXYZ(m_RunMode, dEndPt) Then
            CompileErrorDisplay(FillCircle.SheetName, FillCircle.CmdLineNo, 0)
            Return -1
        End If
        Dim ndgapZ As Double = dEndPt(2)
        Dim surfZ As Double = ndgapZ - FillCircle.Param.NeedleGap
        zRetr = surfZ + FillCircle.Param.RetractHeight
        zClear = surfZ + FillCircle.Param.ClearanceHeight
        CheckRetractClearHeight(ndgapZ, zRetr, zClear)

        If NotVisionMode(m_RunMode) Then
            If Not WorkAreaErrorCheckZ(zRetr) Then
                CompileErrorDisplay(FillCircle.SheetName, FillCircle.CmdLineNo, RetractZError)
                Return -1
            End If
            If Not WorkAreaErrorCheckZ(zClear) Then
                CompileErrorDisplay(FillCircle.SheetName, FillCircle.CmdLineNo, ClearanceZError)
                Return -1
            End If
        End If

        RoundTo3DecimalPoints(zRetr)
        RoundTo3DecimalPoints(zClear)
        '
        If (FillCircle.Param.ArcRadius < 0.001) Then
            m_ArcRadius = gMinArcRadius
        Else
            m_ArcRadius = FillCircle.Param.ArcRadius
        End If
        RoundTo3DecimalPoints(comp1)
        If GenerateStartBlock(comp1, zApht, disCmdList) < 0 Then Return -1

        If IsVisionMode(m_RunMode) Then AddWaitDuration(disCmdList, 150)
        AddSpeed(disCmdList, IDS.Data.Hardware.Gantry.ElementZSpeed)

        If (NotVisionMode(m_RunMode)) Then 'except vision mode
            RoundTo3DecimalPoints(zApht)
            RoundTo3DecimalPoints(zDisp1)
            If IsXYPlane(m_Plane) Then
                pos(0) = comp1(0)
                pos(1) = comp1(1)
                pos(2) = zApht
                AddMoveXYZ(disCmdList, pos)
                AddWaitIdle(disCmdList)
                pos(0) = comp1(0)
                pos(1) = comp1(1)
                pos(2) = zDisp1
                AddMoveXYZ(disCmdList, pos)
            ElseIf IsYZPlane(m_Plane) Then
                pos(0) = comp1(1)
                pos(1) = zApht
                pos(2) = comp1(0)
                AddMoveXYZ(disCmdList, pos)
                AddWaitIdle(disCmdList)
                pos(0) = comp1(1)
                pos(1) = zDisp1
                pos(2) = comp1(0)
                AddMoveXYZ(disCmdList, pos)
            Else
                pos(0) = comp1(0)
                pos(1) = zApht
                pos(2) = comp1(1)
                AddMoveXYZ(disCmdList, pos)
                AddWaitIdle(disCmdList)
                pos(0) = comp1(0)
                pos(1) = zDisp1
                pos(2) = comp1(1)
                AddMoveXYZ(disCmdList, pos)
            End If

            If FillCircle.Param.DispenseOn And m_RunMode > 3 Then
                AddWaitLoaded(disCmdList)
                AddOnIO(disCmdList, m_NeedleIO)
            End If
            AddWaitIdle(disCmdList)
        End If
        If FillCircle.Param.TravelDelay > 0 Then
            AddWaitDuration(disCmdList, FillCircle.Param.TravelDelay)
        End If

        If NotXYPlane(m_Plane) Then AddXYZBase(disCmdList)
        SetToXYPlane(m_Plane)

        m_Speed = FillCircle.Param.TravelSpeed
        AddSpeed(disCmdList, m_Speed)

        Dim count As Integer
        Dim tmp(3), dz As Double
        Dim isFirst As Boolean = True
        Dim dummy(3) As Double
        Dim dc1(3), dc2(3), de1(3), de2(3) As Double
        Dim k As Integer = 0

        p1(0) = comp1(0)
        p1(1) = comp1(1)
        p1(2) = zDisp1

        For count = 1 To pointno - 1
            k += 1
            pointrec = pointlist(count)

            If (k = 1) Then
                p2(0) = pointrec.x
                p2(1) = pointrec.y
                p2(2) = pointrec.z
            ElseIf (k = 2) Then
                p3(0) = pointrec.x
                p3(1) = pointrec.y
                p3(2) = pointrec.z
            End If

            If (k = 2) Then
                If PointOnArc3d(p1, p2, p3, 0.0, centre, radius, direction, dummy) < 0 Then
                    CompileErrorDisplay(FillCircle.SheetName, FillCircle.CmdLineNo, 11)
                    Return -1
                End If

                dc1(0) = centre(0) - p1(0)
                dc1(1) = centre(1) - p1(1)
                dc1(2) = 0.0

                de1(0) = p3(0) - p1(0)
                de1(1) = p3(1) - p1(1)
                de1(2) = 0.0
                dz = p3(2) - p1(2)
                RoundTo3DecimalPoints(dc1)
                RoundTo3DecimalPoints(de1)

                If (NotVisionMode(m_RunMode)) Then 'except vision mode
                    AddMoveHelical(disCmdList, de1(0), de1(1), dc1(0), dc1(1), direction, dz)
                Else
                    AddMoveArc(disCmdList, de1(0), de1(1), dc1(0), dc1(1), direction)
                End If

                k = 0
                p1(0) = p3(0)
                p1(1) = p3(1)
                p1(2) = p3(2)
            End If

            If pointrec.DispenseOn = False And m_RunMode > 3 And isFirst Then
                isFirst = False
                AddWaitLoaded(disCmdList)
                AddOffIO(disCmdList, m_NeedleIO)
            End If
        Next

        If (k = 1) Then
            RoundTo3DecimalPoints(p2)
            If (NotVisionMode(m_RunMode)) Then 'except vision mode
                pos(0) = p2(0)
                pos(1) = p2(1)
                pos(2) = p2(2)
                AddMoveXYZ(disCmdList, pos)
            Else
                pos(0) = p2(0)
                pos(1) = p2(1)
                AddMoveXY(disCmdList, pos)
            End If
        End If


        AddWaitIdle(disCmdList)
        If FillCircle.Param.RetractDelay > 0 Then
            AddWaitDuration(disCmdList, FillCircle.Param.RetractDelay)
        End If

        If m_IsEndElement Then
            Dim speed As Double = FillCircle.Param.RetractSpeed
            RoundTo3DecimalPoints(dEndPt)
            GenerateEndBlock(1, dEndPt, speed, zRetr, zClear, disCmdList)
        End If

        m_PrevRetractPt(0) = dEndPt(0)
        m_PrevRetractPt(1) = dEndPt(1)
        m_PrevRetractPt(2) = zRetr
        m_PrevClearPt(0) = dEndPt(0)
        m_PrevClearPt(1) = dEndPt(1)
        m_PrevClearPt(2) = zClear
        m_PrevArcRadius = m_ArcRadius
        m_PrevRetractSpeed = FillCircle.Param.RetractSpeed
        Return 0
    End Function

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Generate fill circle motion command 
    '  disCmdList:    dispensing command list
    '
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Public Function GenerateFillCircleCommandSet(ByVal elemData As Object, ByRef disCmdList As ArrayList) As Integer
    '    '''''''''''''''''''''''''''''''''''
    '    Dim FillCircle As CIDSFillCircle = elemData
    '    If FillCircle.CmdType <> "FILLCIRCLE" Then
    '        Return -1
    '    End If
    '    Dim CmdStr As String

    '    Dim p1(3), p2(3), p3(3), comp1(3), comp2(3), comp3(3), detatchPt(3), dEndPt(3), midPt(3), off(3) As Double
    '    Dim ApproachZ, zDisp1, zDisp2, zDisp3, RetractZ, ClearanceZ As Double
    '    Dim centre(3), radius As Double
    '    Dim direction As Integer
    '    p1(0) = FillCircle.PosX1 '+ m_referencePt(0)   ' Camera pos
    '    p1(1) = FillCircle.PosY1 '+ m_referencePt(1)
    '    p1(2) = 0.0
    '    p2(0) = FillCircle.PosX2 '+ m_referencePt(0)   ' Camera pos
    '    p2(1) = FillCircle.PosY2 '+ m_referencePt(1)
    '    p2(2) = 0.0
    '    p3(0) = FillCircle.PosX3 '+ m_referencePt(0)   ' Camera pos
    '    p3(1) = FillCircle.PosY3 '+ m_referencePt(1)
    '    p3(2) = 0.0

    '    Dim needlename As String = FillCircle.Param.Needle
    '    GetNeedleData(m_RunMode, needlename, off, m_NeedleIO)
    '    If NotVisionMode(m_RunMode) And m_NeedleIO <> m_PrevNeedleIO Then
    '        NeedleChange(m_NeedleIO, disCmdList)
    '        m_PrevNeedleIO = m_NeedleIO
    '    End If

    '    Dim fixHeight As Double = off(2) - FillCircle.HeightCompS
    '    ApproachZ = FillCircle.PosZ1 + FillCircle.Param.ApproachHeight + fixHeight
    '    zDisp1 = FillCircle.PosZ1 + FillCircle.Param.NeedleGap + fixHeight
    '    zDisp2 = FillCircle.PosZ2 + FillCircle.Param.NeedleGap + fixHeight
    '    zDisp3 = FillCircle.PosZ3 + FillCircle.Param.NeedleGap + fixHeight
    '    CheckApproachHeight(zDisp1, ApproachZ)

    '    If (NotVisionMode(m_RunMode)) Then 'except vision mode
    '        comp1(0) = p1(0) - off(0)  ' needle Pos
    '        comp1(1) = p1(1) - off(1)
    '        comp1(2) = zDisp1
    '        comp2(0) = p2(0) - off(0)  ' needle Pos
    '        comp2(1) = p2(1) - off(1)
    '        comp2(2) = zDisp2
    '        comp3(0) = p3(0) - off(0)  ' needle Pos
    '        comp3(1) = p3(1) - off(1)
    '        comp3(2) = zDisp3
    '        If Not WorkAreaErrorCheckZ(ApproachZ) Then
    '            CompileErrorDisplay(FillCircle.SheetName, FillCircle.CmdLineNo, 0)
    '            Return -1
    '        End If
    '    Else
    '        comp1(0) = p1(0)
    '        comp1(1) = p1(1)
    '        comp1(2) = FillCircle.PosZ1
    '        comp2(0) = p2(0)
    '        comp2(1) = p2(1)
    '        comp2(2) = FillCircle.PosZ2
    '        comp3(0) = p3(0)
    '        comp3(1) = p3(1)
    '        comp3(2) = FillCircle.PosZ3
    '    End If

    '    If Not WorkAreaErrorCheckXYZ(m_RunMode, comp1) Then
    '        CompileErrorDisplay(FillCircle.SheetName, FillCircle.CmdLineNo, 0)
    '        Return -1
    '    End If
    '    If Not WorkAreaErrorCheckXYZ(m_RunMode, comp2) Then
    '        CompileErrorDisplay(FillCircle.SheetName, FillCircle.CmdLineNo, 0)
    '        Return -1
    '    End If
    '    If Not WorkAreaErrorCheckXYZ(m_RunMode, comp3) Then
    '        CompileErrorDisplay(FillCircle.SheetName, FillCircle.CmdLineNo, 0)
    '        Return -1
    '    End If

    '    Dim r_pitch As Double = FillCircle.Param.Pitch
    '    'Dim rr_pitch As Double = 0.0
    '    If (r_pitch = 0) Then
    '        MsgBox("Pitch cannot be set to zero for Filled Circle.")
    '        Return -1
    '    End If
    '    Dim height As Double = FillCircle.Param.FillHeight
    '    Dim NoOfRun As Integer = FillCircle.Param.NoOfRun
    '    Dim inout As String = FillCircle.Param.SprialIn
    '    Dim detaildist As Double = FillCircle.Param.DeTailDist
    '    Dim pointlist As New ArrayList
    '    'generate spiral point list to fit spiral
    '    If GetSpiralPointList(comp1, comp2, comp3, r_pitch, NoOfRun, height, detaildist, inout, pointlist) < 0 Then
    '        CompileErrorDisplay(FillCircle.SheetName, FillCircle.CmdLineNo, 13)
    '        Return -1
    '    End If

    '    Dim pointno As Integer = pointlist.Count
    '    If pointno <= 0 Then Return -1
    '    Dim endptrec As Point3d = pointlist(pointno - 1)
    '    Dim startptrec As Point3d = pointlist(0)
    '    Dim pointrec As Point3d
    '    If inout.ToUpper = "OUT" Then
    '        comp1(0) = startptrec.x
    '        comp1(1) = startptrec.y
    '        comp1(2) = startptrec.z
    '        ApproachZ = startptrec.z - FillCircle.Param.NeedleGap + FillCircle.Param.ApproachHeight
    '        zDisp1 = startptrec.z
    '        CheckApproachHeight(zDisp1, ApproachZ)
    '        If NotVisionMode(m_RunMode) Then
    '            If Not WorkAreaErrorCheckZ(zDisp1) Then
    '                CompileErrorDisplay(FillCircle.SheetName, FillCircle.CmdLineNo, 0)
    '                Return -1
    '            End If
    '            If Not WorkAreaErrorCheckZ(ApproachZ) Then
    '                CompileErrorDisplay(FillCircle.SheetName, FillCircle.CmdLineNo, 0)
    '                Return -1
    '            End If
    '        End If
    '    End If

    '    dEndPt(0) = endptrec.x
    '    dEndPt(1) = endptrec.y
    '    dEndPt(2) = endptrec.z
    '    If Not WorkAreaErrorCheckXYZ(m_RunMode, dEndPt) Then
    '        CompileErrorDisplay(FillCircle.SheetName, FillCircle.CmdLineNo, 0)
    '        Return -1
    '    End If
    '    Dim ndgapZ As Double = dEndPt(2)
    '    Dim surfZ As Double = ndgapZ - FillCircle.Param.NeedleGap
    '    RetractZ = surfZ + FillCircle.Param.RetractHeight
    '    ClearanceZ = surfZ + FillCircle.Param.ClearanceHeight
    '    CheckRetractClearHeight(ndgapZ, RetractZ, ClearanceZ)

    '    If NotVisionMode(m_RunMode) Then
    '        If Not WorkAreaErrorCheckZ(RetractZ) Then
    '            CompileErrorDisplay(FillCircle.SheetName, FillCircle.CmdLineNo, 0)
    '            Return -1
    '        End If
    '        If Not WorkAreaErrorCheckZ(ClearanceZ) Then
    '            CompileErrorDisplay(FillCircle.SheetName, FillCircle.CmdLineNo, 0)
    '            Return -1
    '        End If
    '    End If

    '    RoundTo3DecimalPoints(RetractZ)
    '    RoundTo3DecimalPoints(ClearanceZ)
    '    '
    '    If (FillCircle.Param.ArcRadius < 0.001) Then
    '        m_ArcRadius = gMinArcRadius
    '    Else
    '        m_ArcRadius = FillCircle.Param.ArcRadius
    '    End If
    '    RoundTo3DecimalPoints(comp1)
    '    If GenerateStartBlock(comp1, ApproachZ, disCmdList) < 0 Then Return -1
    '    AddSpeed(disCmdList, IDS.Data.Hardware.Gantry.ElementZSpeed)

    '    If IsVisionMode(m_RunMode) Then
    '        AddWaitDuration(disCmdList, 150)
    '    End If

    '    If (NotVisionMode(m_RunMode)) Then 'except vision mode
    '        RoundTo3DecimalPoints(ApproachZ)
    '        RoundTo3DecimalPoints(zDisp1)
    '        If IsXYPlane(m_Plane) Then
    '            pos(0) = comp2(0)
    '            pos(1) = comp2(1)
    '            pos(2) = ApproachZ
    '            AddMoveXYZ(disCmdList, pos)
    '            AddWaitIdle(disCmdList)
    '            pos(0) = comp2(0)
    '            pos(1) = comp2(1)
    '            pos(2) = zDisp1
    '            AddMoveXYZ(disCmdList, pos)
    '        ElseIf IsYZPlane(m_Plane) Then
    '            pos(0) = comp2(1)
    '            pos(1) = ApproachZ
    '            pos(2) = comp2(0)
    '            AddMoveXYZ(disCmdList, pos)
    '            AddWaitIdle(disCmdList)
    '            pos(0) = comp2(1)
    '            pos(1) = zDisp1
    '            pos(2) = comp2(0)
    '            AddMoveXYZ(disCmdList, pos)
    '        Else
    '            pos(0) = comp2(0)
    '            pos(1) = ApproachZ
    '            pos(2) = comp2(1)
    '            AddMoveXYZ(disCmdList, pos)
    '            AddWaitIdle(disCmdList)
    '            pos(0) = comp2(0)
    '            pos(1) = zDisp1
    '            pos(2) = comp2(1)
    '            AddMoveXYZ(disCmdList, pos)
    '        End If

    '        If FillCircle.Param.DispenseOn And m_RunMode > 3 Then
    '            AddWaitLoaded(disCmdList)
    '            AddOnIO(disCmdList, m_NeedleIO)
    '        End If
    '        AddWaitIdle(disCmdList)
    '    End If
    '    If FillCircle.Param.TravelDelay > 0 Then
    '        AddWaitDuration(disCmdList, FillCircle.Param.TravelDelay)
    '    End If

    '    If NotXYPlane(m_Plane) Then AddXYZBase(disCmdList)
    '    SetToXYPlane(m_Plane)

    '    m_Speed = FillCircle.Param.TravelSpeed
    '    AddSpeed(disCmdList, m_Speed)

    '    Dim count As Integer
    '    Dim tmp(3), dz As Double
    '    Dim isFirst As Boolean = True
    '    Dim dummy(3) As Double
    '    Dim dc1(3), dc2(3), de1(3), de2(3) As Double
    '    Dim k As Integer = 0

    '    p1(0) = comp1(0)
    '    p1(1) = comp1(1)
    '    p1(2) = zDisp1

    '    For count = 1 To pointno - 1
    '        k += 1
    '        pointrec = pointlist(count)

    '        If (k = 1) Then
    '            p2(0) = pointrec.x
    '            p2(1) = pointrec.y
    '            p2(2) = pointrec.z
    '        ElseIf (k = 2) Then
    '            p3(0) = pointrec.x
    '            p3(1) = pointrec.y
    '            p3(2) = pointrec.z
    '        End If

    '        If (k = 2) Then
    '            If PointOnArc3d(p1, p2, p3, 0.0, centre, radius, direction, dummy) < 0 Then
    '                CompileErrorDisplay(FillCircle.SheetName, FillCircle.CmdLineNo, 11)
    '                Return -1
    '            End If

    '            dc1(0) = centre(0) - p1(0)
    '            dc1(1) = centre(1) - p1(1)
    '            dc1(2) = 0.0

    '            de1(0) = p3(0) - p1(0)
    '            de1(1) = p3(1) - p1(1)
    '            de1(2) = 0.0
    '            dz = p3(2) - p1(2)
    '            RoundTo3DecimalPoints(dc1)
    '            RoundTo3DecimalPoints(de1)
    '            'RoundTo3DecimalPoints(tmp)

    '            If (NotVisionMode(m_RunMode)) Then 'except vision mode
    '                AddMoveHelical(disCmdList, de1(0), de1(1), dc1(0), dc1(1), direction, dz)
    '            Else
    '                AddMoveArc(disCmdList, de1(0), de1(1), dc1(0), dc1(1), direction)
    '            End If
    '            k = 0
    '            p1(0) = p3(0)
    '            p1(1) = p3(1)
    '            p1(2) = p3(2)
    '        End If

    '        If pointrec.DispenseOn = False And m_RunMode > 3 And isFirst Then
    '            isFirst = False
    '            AddWaitLoaded(disCmdList)
    '            AddOffIO(disCmdList, m_NeedleIO)

    '        End If
    '    Next

    '    If (k = 1) Then
    '        RoundTo3DecimalPoints(p2)
    '        If (NotVisionMode(m_RunMode)) Then 'except vision mode
    '            pos(0) = p2(0)
    '            pos(1) = p2(1)
    '            pos(2) = p2(2)
    '            AddMoveXYZ(disCmdList, pos)
    '        Else
    '            pos(0) = p2(0)
    '            pos(1) = p2(1)
    '            AddMoveXY(disCmdList, pos)
    '        End If
    '    End If

    '    AddWaitIdle(disCmdList)
    '    If FillCircle.Param.RetractDelay > 0 Then
    '        AddWaitDuration(disCmdList, FillCircle.Param.RetractDelay)
    '    End If

    '    If (m_IsEndElement) Then
    '        Dim speed As Double = FillCircle.Param.RetractSpeed
    '        RoundTo3DecimalPoints(dEndPt)
    '        GenerateEndBlock(1, dEndPt, speed, RetractZ, ClearanceZ, disCmdList)
    '    End If

    '    m_PrevRetractPt(0) = dEndPt(0)
    '    m_PrevRetractPt(1) = dEndPt(1)
    '    m_PrevRetractPt(2) = RetractZ
    '    m_PrevClearPt(0) = dEndPt(0)
    '    m_PrevClearPt(1) = dEndPt(1)
    '    m_PrevClearPt(2) = ClearanceZ
    '    m_PrevArcRadius = m_ArcRadius
    '    m_PrevRetractSpeed = FillCircle.Param.RetractSpeed
    '    Return 0
    'End Function

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Generate fill reactangle motion command 
    '  disCmdList:    dispensing command list
    '
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Public Function GenerateFillRectCommandSet(ByVal elemData As Object, ByRef disCmdList As ArrayList) As Integer
        '''''''''''''''''''''''''''''''''''
        Dim FillRect As CIDSFillRectangle = elemData
        If FillRect.CmdType <> "FILLRECTANGLE" Then
            Return -1
        End If
        Dim CmdStr As String

        Dim p1(3), p2(3), p3(3), comp1(3), comp2(3), comp3(3), detatchPt(3), dEndPt(3), midPt(3), off(3) As Double
        Dim ApproachZ, zDisp1, zDisp2, zDisp3, RetractZ, ClearanceZ As Double
        Dim centre(3), radius As Double
        Dim direction As Integer
        p1(0) = FillRect.PosX1 '+ m_referencePt(0)   ' Camera pos
        p1(1) = FillRect.PosY1 '+ m_referencePt(1)
        p1(2) = 0.0
        p2(0) = FillRect.PosX2 '+ m_referencePt(0)   ' Camera pos
        p2(1) = FillRect.PosY2 '+ m_referencePt(1)
        p2(2) = 0.0
        p3(0) = FillRect.PosX3 '+ m_referencePt(0)   ' Camera pos
        p3(1) = FillRect.PosY3 '+ m_referencePt(1)
        p3(2) = 0.0

        Dim needlename As String = FillRect.Param.Needle
        GetNeedleData(m_RunMode, needlename, off, m_NeedleIO)
        If NotVisionMode(m_RunMode) And m_NeedleIO <> m_PrevNeedleIO Then
            NeedleChange(m_NeedleIO, disCmdList)
            m_PrevNeedleIO = m_NeedleIO
        End If

        Dim fixHeight As Double = off(2) - FillRect.HeightCompS
        ApproachZ = FillRect.PosZ1 + fixHeight + FillRect.Param.ApproachHeight
        zDisp1 = FillRect.PosZ1 + fixHeight + FillRect.Param.NeedleGap
        zDisp2 = FillRect.PosZ2 + fixHeight + FillRect.Param.NeedleGap
        zDisp3 = FillRect.PosZ3 + fixHeight + FillRect.Param.NeedleGap
        CheckApproachHeight(zDisp1, ApproachZ)

        If (NotVisionMode(m_RunMode)) Then 'except vision mode
            comp1(0) = p1(0) - off(0)  ' needle Pos
            comp1(1) = p1(1) - off(1)
            comp1(2) = zDisp1
            comp2(0) = p2(0) - off(0)  ' needle Pos
            comp2(1) = p2(1) - off(1)
            comp2(2) = zDisp2
            comp3(0) = p3(0) - off(0)  ' needle Pos
            comp3(1) = p3(1) - off(1)
            comp3(2) = zDisp3
            If Not WorkAreaErrorCheckZ(ApproachZ) Then
                CompileErrorDisplay(FillRect.SheetName, FillRect.CmdLineNo, ApproachZError)
                Return -1
            End If
        Else
            comp1(0) = p1(0)
            comp1(1) = p1(1)
            comp1(2) = FillRect.PosZ1
            comp2(0) = p2(0)
            comp2(1) = p2(1)
            comp2(2) = FillRect.PosZ2
            comp3(0) = p3(0)
            comp3(1) = p3(1)
            comp3(2) = FillRect.PosZ3
        End If

        If Not WorkAreaErrorCheckXYZ(m_RunMode, comp1) Then
            CompileErrorDisplay(FillRect.SheetName, FillRect.CmdLineNo, 0)
            Return -1
        End If
        If Not WorkAreaErrorCheckXYZ(m_RunMode, comp2) Then
            CompileErrorDisplay(FillRect.SheetName, FillRect.CmdLineNo, 0)
            Return -1
        End If
        If Not WorkAreaErrorCheckXYZ(m_RunMode, comp3) Then
            CompileErrorDisplay(FillRect.SheetName, FillRect.CmdLineNo, 0)
            Return -1
        End If

        Dim r_pitch As Double = FillRect.Param.Pitch
        Dim height As Double = FillRect.Param.FillHeight
        Dim NoOfRun As Integer = FillRect.Param.NoOfRun
        Dim inout As String = FillRect.Param.SprialIn
        Dim detaildist As Double = FillRect.Param.DeTailDist
        Dim pointlist As New ArrayList
        'Calculate sprial rectangle enpoints
        If GetSpiralRectPointList(comp1, comp2, comp3, r_pitch, NoOfRun, height, detaildist, inout, pointlist) < 0 Then
            CompileErrorDisplay(FillRect.SheetName, FillRect.CmdLineNo, 13)
            Return -1
        End If

        Dim pointno As Integer = pointlist.Count
        If pointno <= 0 Then Return -1
        Dim endptrec As Point3d = pointlist(pointno - 1)
        Dim startptrec As Point3d = pointlist(0)
        Dim pointrec As Point3d
        If inout.ToUpper = "OUT" Then
            comp1(0) = startptrec.x
            comp1(1) = startptrec.y
            comp1(2) = startptrec.z
            ApproachZ = startptrec.z - FillRect.Param.NeedleGap + FillRect.Param.ApproachHeight
            zDisp1 = startptrec.z
            CheckApproachHeight(zDisp1, ApproachZ)
            If NotVisionMode(m_RunMode) Then
                If Not WorkAreaErrorCheckZ(zDisp1) Then
                    CompileErrorDisplay(FillRect.SheetName, FillRect.CmdLineNo, 0)
                    Return -1
                End If
                If Not WorkAreaErrorCheckZ(ApproachZ) Then
                    CompileErrorDisplay(FillRect.SheetName, FillRect.CmdLineNo, ApproachZError)
                    Return -1
                End If
            End If
        End If

        dEndPt(0) = endptrec.x
        dEndPt(1) = endptrec.y
        dEndPt(2) = endptrec.z
        If Not WorkAreaErrorCheckXYZ(m_RunMode, dEndPt) Then
            CompileErrorDisplay(FillRect.SheetName, FillRect.CmdLineNo, 0)
            Return -1
        End If

        Dim ndgapZ As Double = dEndPt(2)
        Dim surfZ As Double = ndgapZ - FillRect.Param.NeedleGap
        RetractZ = surfZ + FillRect.Param.RetractHeight
        ClearanceZ = surfZ + FillRect.Param.ClearanceHeight
        CheckRetractClearHeight(ndgapZ, RetractZ, ClearanceZ)

        If NotVisionMode(m_RunMode) Then
            If Not WorkAreaErrorCheckZ(RetractZ) Then
                CompileErrorDisplay(FillRect.SheetName, FillRect.CmdLineNo, RetractZError)
                Return -1
            End If
            If Not WorkAreaErrorCheckZ(ClearanceZ) Then
                CompileErrorDisplay(FillRect.SheetName, FillRect.CmdLineNo, ClearanceZError)
                Return -1
            End If
        End If

        RoundTo3DecimalPoints(RetractZ)
        RoundTo3DecimalPoints(ClearanceZ)
        '
        If (FillRect.Param.ArcRadius < 0.001) Then
            m_ArcRadius = gMinArcRadius
        Else
            m_ArcRadius = FillRect.Param.ArcRadius
        End If
        RoundTo3DecimalPoints(comp1)
        If GenerateStartBlock(comp1, ApproachZ, disCmdList) < 0 Then Return -1
        AddSpeed(disCmdList, IDS.Data.Hardware.Gantry.ElementZSpeed)

        If (NotVisionMode(m_RunMode)) Then 'except vision mode
            RoundTo3DecimalPoints(ApproachZ)
            RoundTo3DecimalPoints(zDisp1)
            If IsXYPlane(m_Plane) Then
                pos(0) = comp1(0)
                pos(1) = comp1(1)
                pos(2) = ApproachZ
                AddMoveXYZ(disCmdList, pos)
                pos(0) = comp1(0)
                pos(1) = comp1(1)
                pos(2) = zDisp1
                AddMoveXYZ(disCmdList, pos)
            ElseIf IsYZPlane(m_Plane) Then
                pos(0) = comp1(1)
                pos(1) = ApproachZ
                pos(2) = comp1(0)
                AddMoveXYZ(disCmdList, pos)
                pos(0) = comp1(1)
                pos(1) = zDisp1
                pos(2) = comp1(0)
                AddMoveXYZ(disCmdList, pos)
            Else
                pos(0) = comp1(0)
                pos(1) = ApproachZ
                pos(2) = comp1(1)
                AddMoveXYZ(disCmdList, pos)
                pos(0) = comp1(0)
                pos(1) = zDisp1
                pos(2) = comp1(1)
                AddMoveXYZ(disCmdList, pos)
            End If

            AddWaitLoaded(disCmdList)
            If FillRect.Param.DispenseOn And m_RunMode > 3 Then
                AddOnIO(disCmdList, m_NeedleIO)
            End If
        End If
        If FillRect.Param.TravelDelay > 0 Then
            AddWaitDuration(disCmdList, FillRect.Param.TravelDelay)
        End If

        If NotXYPlane(m_Plane) Then AddXYZBase(disCmdList)
        SetToXYPlane(m_Plane)


        m_Speed = FillRect.Param.TravelSpeed
        AddSpeed(disCmdList, m_Speed)

        Dim count As Integer
        Dim tmp(3) As Double
        Dim isFirst As Boolean = True
        For count = 0 To pointno - 1
            pointrec = pointlist(count)
            tmp(0) = pointrec.x
            tmp(1) = pointrec.y
            tmp(2) = pointrec.z
            RoundTo3DecimalPoints(tmp)
            If (NotVisionMode(m_RunMode)) Then 'except vision mode
                pos(0) = tmp(0)
                pos(1) = tmp(1)
                pos(2) = tmp(2)
                AddMoveXYZ(disCmdList, pos)
            Else
                pos(0) = tmp(0)
                pos(1) = tmp(1)
                AddMoveXY(disCmdList, pos)
            End If
            If pointrec.DispenseOn = True Then
                AddWaitIdle(disCmdList)
            End If

            If pointrec.DispenseOn = False And m_RunMode > 3 And isFirst Then
                isFirst = False
                AddWaitLoaded(disCmdList)
                AddOffIO(disCmdList, m_NeedleIO)
            End If
        Next
        AddWaitIdle(disCmdList)
        If FillRect.Param.RetractDelay > 0 Then
            AddWaitDuration(disCmdList, FillRect.Param.RetractDelay)
        End If

        If (m_IsEndElement) Then
            Dim speed As Double = FillRect.Param.RetractSpeed
            RoundTo3DecimalPoints(dEndPt)
            GenerateEndBlock(1, dEndPt, speed, RetractZ, ClearanceZ, disCmdList)
        End If

        m_PrevRetractPt(0) = dEndPt(0)
        m_PrevRetractPt(1) = dEndPt(1)
        m_PrevRetractPt(2) = RetractZ
        m_PrevClearPt(0) = dEndPt(0)
        m_PrevClearPt(1) = dEndPt(1)
        m_PrevClearPt(2) = ClearanceZ
        m_PrevArcRadius = m_ArcRadius
        m_PrevRetractSpeed = FillRect.Param.RetractSpeed
        Return 0
    End Function

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Find last dispensing element' index in the pattern element list
    '  
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Public Function getLastDispenseElementIndex(ByRef index As Integer) As Integer
        Dim elemData As Object
        Dim count As Integer = m_PatternList.Count
        Dim row As Integer
        Dim Type As String
        index = 1
        For row = count To 1 Step -1
            elemData = m_PatternList(row - 1)
            If elemData Is Nothing Then
                Return -1
            End If

            Type = elemData.CmdType
            Type = Type.ToUpper

            Select Case Type
                Case "SETIO", "GETIO", "LINKEND"  '"WAIT", 'none dispensing element

                Case Else
                    index = row
                    Return 0
            End Select
        Next
        Return 0
    End Function

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Generate sub pattern start/end command for QC fail skipping
    '  disCmdList:    dispensing command list
    '
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Public Function GenerateSubSECommandSet(ByVal elemData As Object, ByRef disCmdList As ArrayList) As Integer
        Dim subSE As CIDSSub = elemData
        Dim type As String = subSE.CmdType
        If type <> "SUBS" And type <> "SUBE" Then
            Return -1
        End If
        Dim level As Integer = subSE.level 'sub level
        Dim CmdStr As String
        If type = "SUBS" Then
            AddSubStart(disCmdList, level)
        Else
            Dim speed, RetractZ, ClearanceZ As Double
            RetractZ = m_PrevRetractPt(2)
            ClearanceZ = m_PrevClearPt(2)
            GenerateEndBlock(0, m_PrevRetractPt, m_PrevRetractSpeed, RetractZ, ClearanceZ, disCmdList)
            AddSubEnd(disCmdList, level)
            m_IsFirstElement = True
        End If
        Return 0
    End Function

    Public Function Compile(ByRef dispenselist As ArrayList, Optional ByRef GlobalQCList As ArrayList = Nothing) As Integer
        Dim row As Integer
        Dim elemData, endelemdata, nextelemdata As Object
        Dim type As String
        If dispenselist.Count > 0 Then
            dispenselist.Clear()
        End If
        m_NoMoveList.Clear()

        m_IsFirstElement = True
        m_IsEndElement = False
        m_IsLinkElement = False
        m_IsFirstElementLink = True
        m_IsEndElementLink = False

        Dim count As Integer = m_PatternList.Count
        endelemdata = m_PatternList(count - 1)
        If endelemdata Is Nothing Then
            Return -1
        End If
        type = endelemdata.CmdType

        Dim index As Integer = 0
        If getLastDispenseElementIndex(index) < 0 Then Return -1 'last dispensing element index

        For row = 1 To count     'handling whole pattern command list
            elemData = m_PatternList(row - 1)
            If elemData Is Nothing Then
                Return -1
            End If
            Dim nextEleNo As Integer = row + 1
            If (nextEleNo <= count) Then
                nextelemdata = m_PatternList(nextEleNo - 1)
                If nextelemdata Is Nothing Then
                    Return -1
                End If
                type = nextelemdata.CmdType
                If (type.ToUpper = "LINKEND") Then m_IsEndElementLink = True
            End If

            type = elemData.CmdType

            If m_IsLinkElement Then
                If type.ToUpper = "LINE" Then
                    type = "LINKLINE"
                ElseIf type.ToUpper = "ARC" Then
                    type = "LINKARC"
                End If
            End If

            If row = index Then
                m_IsEndElement = True
            End If

            Select Case type
                Case "SETIO", "GETIO"  '"WAIT",
                    m_NoMoveList.Add(elemData)
                    If row = count Then
                        If GenerateNoMoveCommandSet(dispenselist) < 0 Then Return -1
                    End If
                Case "QC"
                    If GenerateQCCommandSet(elemData, dispenselist) < 0 Then
                        Return -1
                    End If
                Case "WAIT"
                    If GenerateWaitCommandSet(elemData, dispenselist) < 0 Then
                        Return -1
                    End If
                Case "PURGE"
                    If GeneratePurgeCommandSet(elemData, dispenselist) < 0 Then
                        Return -1
                    End If
                Case "VOLUMECALIBRATION"
                    If GenerateVolumeCalibrationCommandSet(elemData, dispenselist) < 0 Then
                        Return -1
                    End If
                Case "CLEAN"
                    If GenerateCleanCommandSet(elemData, dispenselist) < 0 Then
                        Return -1
                    End If
                Case "CHIPEDGE"
                    If GenerateChipEdgeCommandSet(elemData, dispenselist) < 0 Then
                        Return -1
                    End If
                    m_IsFirstElement = False
                Case "MOVE"
                    If GenerateMoveCommandSet(elemData, dispenselist) < 0 Then
                        Return -1
                    End If
                Case "DOT"
                    If GenerateDotCommandSet(elemData, dispenselist, GlobalQCList) < 0 Then
                        Return -1
                    End If
                    m_IsFirstElement = False
                Case "DOTARRAY"
                    If GenerateDotArrayCommandSet(elemData, dispenselist) < 0 Then
                        Return -1
                    End If
                    m_IsFirstElement = False
                Case "LINE"
                    If GenerateLineCommandSet(elemData, dispenselist) < 0 Then
                        Return -1
                    End If
                    m_IsFirstElement = False
                Case "ARC"
                    If GenerateArcCommandSet(elemData, dispenselist) < 0 Then
                        Return -1
                    End If
                    m_IsFirstElement = False
                Case "CIRCLE"
                    If GenerateCircleCommandSet(elemData, dispenselist) < 0 Then
                        Return -1
                    End If
                    m_IsFirstElement = False
                Case "FILLCIRCLE"
                    If GenerateFillCircleCommandSet(elemData, dispenselist) < 0 Then
                        Return -1
                    End If
                    m_IsFirstElement = False
                Case "LINKSTART"
                    If GenerateLinkSCommandSet(elemData) < 0 Then
                        Return -1
                    End If
                Case "LINKLINE"
                    If GenerateLinkLineCommandSet(elemData, dispenselist) < 0 Then
                        Return -1
                    End If
                    m_IsFirstElement = False
                Case "LINKARC"
                    If GenerateLinkArcCommandSet(elemData, dispenselist) < 0 Then
                        Return -1
                    End If
                    m_IsFirstElement = False
                Case "LINKEND"
                    If GenerateLinkECommandSet(elemData) < 0 Then
                        Return -1
                    End If

                Case "RECTANGLE"
                    If GenerateRoundTo3DecimalPointsRectCommandSet(elemData, dispenselist) < 0 Then
                        Return -1
                    End If
                    m_IsFirstElement = False
                Case "FILLRECTANGLE"
                    If GenerateFillRectCommandSet(elemData, dispenselist) < 0 Then
                        Return -1
                    End If
                    m_IsFirstElement = False
                Case "SUBS", "SUBE"
                    GenerateSubSECommandSet(elemData, dispenselist)
                Case Else
                    Return -1

            End Select
            'SJ add for freeze GUI
            'TraceDoEvents()
        Next row

        dispenselist.Add(18)
        Return 0
    End Function

    Dim CommandString As String

    'Set execution mode
    '   RunMode:  0: vision 1:Dry Needle 2:Dry left 3:Dry right 4:Wet 5:Wet left 6:Wet right
    Public Sub SetRunMode(ByVal RunMode As Integer)
        m_RunMode = RunMode
    End Sub

    Function NotVisionMode(ByVal status As Integer)
        Return status <> 0
    End Function
    Function IsVisionMode(ByVal status As Integer)
        Return status = 0
    End Function

    Sub SetArcRadius(ByVal ArcToSet As Double, ByVal ArcValue As Double)
        If ArcValue < 0.001 Then
            ArcToSet = gMinArcRadius
        Else
            ArcToSet = ArcValue
        End If
    End Sub

    'motion plane: 0-XY; 1-YZ; 2-XZ
    Function NotXYPlane(ByVal status As Integer)
        Return status <> 0
    End Function
    Function NotYZPlane(ByVal status As Integer)
        Return status <> 1
    End Function
    Function NotXZPlane(ByVal status As Integer)
        Return status <> 2
    End Function
    Function IsXYPlane(ByVal status As Integer)
        Return status = 0
    End Function
    Function IsYZPlane(ByVal status As Integer)
        Return status = 1
    End Function
    Function IsXZPlane(ByVal status As Integer)
        Return status = 2
    End Function
    Sub SetToXYPlane(ByRef status As Integer)
        status = 0
    End Sub
    Sub SetToYZPlane(ByRef status As Integer)
        status = 1
    End Sub
    Sub SetToXZPlane(ByRef status As Integer)
        status = 2
    End Sub
    Sub SetToZXPlane(ByRef status As Integer)
        status = 3
    End Sub

    Sub AddMoveHelical(ByVal CommandList As ArrayList, ByVal end1 As Double, ByVal end2 As Double, ByVal centre1 As Double, ByVal centre2 As Double, ByVal direction As Integer, ByVal distance3 As Double)
        CommandList.Add(3)
        CommandList.Add(end1)
        CommandList.Add(end2)
        CommandList.Add(centre1)
        CommandList.Add(centre2)
        CommandList.Add(direction)
        CommandList.Add(distance3)
    End Sub

    Sub AddNormalDotDispense(ByVal CommandList As ArrayList, ByVal IO_number As Integer, ByVal WaitDuration As Double) ' millisecond resolution
        AddOnIO(CommandList, IO_number)
        AddWaitDuration(CommandList, WaitDuration)
        AddOffIO(CommandList, IO_number)
    End Sub

    Sub AddJettingDotDispense(ByVal CommandList As ArrayList, ByVal IO_number As Integer, ByVal WaitDuration As Double) 'microsecond resolution
        AddHWTrigger(CommandList, WaitDuration)
    End Sub

    Sub AddOnIO(ByVal CommandList As ArrayList, ByVal IO_number As Integer)
        CommandList.Add(17)
        CommandList.Add(IO_number)
        CommandList.Add(1)
    End Sub

    Sub AddOffIO(ByVal CommandList As ArrayList, ByVal IO_number As Integer)
        CommandList.Add(17)
        CommandList.Add(IO_number)
        CommandList.Add(0)
    End Sub

    Sub AddHWTrigger(ByVal CommandList As ArrayList, ByVal Duration As Double)
        'Duration = Duration * 1000
        'CommandString = "HW_TIMER(" & Duration.ToString & ")"
        'If Duration > 0 Then
        '    CommandList.Add(CommandString)
        'End If
    End Sub

    Sub AddWaitDuration(ByVal CommandList As ArrayList, ByVal Duration As Double)
        If Duration > 0 Then
            CommandList.Add(16)
            CommandList.Add(Duration)
        End If
    End Sub

    Sub AddVolumeCalibration(ByVal CommandList As ArrayList)
        CommandList.Add(4)
    End Sub

    Sub AddWaitIdle(ByVal CommandList As ArrayList)
        CommandList.Add(14)
    End Sub

    Sub AddWaitLoaded(ByVal CommandList As ArrayList)
        CommandList.Add(15)
    End Sub

    Sub AddDotArray(ByVal CommandList As ArrayList, ByVal x1 As Single, ByVal y1 As Single, ByVal disp_height As Single, ByVal x2 As Single, ByVal y2 As Single, ByVal x3 As Single, ByVal y3 As Single, ByVal disp_duration As Single, ByVal ret_height As Single, ByVal ret_speed As Single, ByVal clear_height As Single, ByVal appr_height As Single, ByVal retract_delay As Single, ByVal xyspeed As Single, ByVal zspeed As Single, ByVal rows As Single, ByVal cols As Single, ByVal dispense_onoff As Single)
        CommandList.Add(22)
        CommandList.Add(x1)
        CommandList.Add(y1)
        CommandList.Add(disp_height)
        CommandList.Add(x2)
        CommandList.Add(y2)
        CommandList.Add(x3)
        CommandList.Add(y3)
        CommandList.Add(disp_duration)
        CommandList.Add(ret_height)
        CommandList.Add(ret_speed)
        CommandList.Add(clear_height)
        CommandList.Add(appr_height)
        CommandList.Add(retract_delay)
        CommandList.Add(xyspeed)
        CommandList.Add(zspeed)
        CommandList.Add(rows)
        CommandList.Add(cols)
        CommandList.Add(dispense_onoff)
    End Sub

    Sub AddSubStart(ByVal CommandList As ArrayList, ByVal level As Integer)
        CommandList.Add(20)
        CommandList.Add(level)
    End Sub

    Sub AddSubEnd(ByVal CommandList As ArrayList, ByVal level As Integer)
        CommandList.Add(21)
        CommandList.Add(level)
    End Sub

    Sub AddWaitUntil(ByVal CommandList As ArrayList, ByVal IO_number As Integer)
        CommandList.Add(19)
        CommandList.Add(IO_number)
    End Sub

    Sub AddSpeed(ByVal CommandList As ArrayList, ByVal Speed As Double)
        CommandList.Add(11)
        CommandList.Add(Speed)
    End Sub

    Sub AddXYZBase(ByVal CommandList As ArrayList)
        CommandList.Add(7)
    End Sub

    Sub AddYZXBase(ByVal CommandList As ArrayList)
        CommandList.Add(8)
    End Sub

    Sub AddXZYBase(ByVal CommandList As ArrayList)
        CommandList.Add(9)
    End Sub

    Sub AddZXYBase(ByVal CommandList As ArrayList)
        CommandList.Add(25)
    End Sub
    'Add Move Z only command
    Sub AddMoveZ(ByVal CommandList As ArrayList, ByVal pos As Double())
        CommandList.Add(26)
        CommandList.Add(pos(0))
    End Sub
    Sub AddMoveXY(ByVal CommandList As ArrayList, ByVal pos As Double())
        CommandList.Add(0)
        CommandList.Add(pos(0))
        CommandList.Add(pos(1))
    End Sub

    Sub AddMoveXYZ(ByVal CommandList As ArrayList, ByVal pos As Double())
        CommandList.Add(1)
        CommandList.Add(pos(0))
        CommandList.Add(pos(1))
        CommandList.Add(pos(2))
    End Sub

    Sub AddMoveArc(ByVal CommandList As ArrayList, ByVal end1 As Double, ByVal end2 As Double, ByVal centre1 As Double, ByVal centre2 As Double, ByVal direction As Integer)
        CommandList.Add(2)
        CommandList.Add(end1)
        CommandList.Add(end2)
        CommandList.Add(centre1)
        CommandList.Add(centre2)
        CommandList.Add(direction)
    End Sub

    Sub AddMergeOn(ByVal CommandList As ArrayList)
        CommandList.Add(23)
    End Sub
    Sub AddMergeOff(ByVal CommandList As ArrayList)
        CommandList.Add(24)
    End Sub

    Function NoDetailing(ByVal DetailDistance As Double)
        Return System.Math.Abs(DetailDistance) <= 0.001
    End Function

End Class

'Downloading motion data to triomotion controller class
Public Class CIDSPattnBurn
    'Protected m_Tri As CIDSTrioController = IDS.Devices.Motor

    'Public Sub New(ByVal tri As CIDSTrioController)
    '    m_Tri = tri
    'End Sub

    Public Function OnlineRun(ByVal CmdList As ArrayList) As Integer
        Return 0
    End Function

    Dim CmdSize As Integer = 22
    Dim CmdName() As String = {"", "MOVEABS", "MOVECIRC", "MHELICAL", "MVSPLINE", "QC", "", "", "", "", "BASE", "SPEED", "ACCEL", "DECEL", "WAIT IDLE", "WAIT LOADED", "WA", "OP", "STOP", "WAIT UNTIL", "SUBS", "SUBE", "DOTARRAY"}
    Dim CmdLength() As Integer = {0, 4, 5, 6, 3, 37, 0, 0, 0, 0, _
                                   3, 1, 1, 1, 0, 0, 1, 2, 0, 1, 1, 1, 16}

    Public Function GetCmdInfo(ByVal commandstr As String, ByRef cmdtype As Integer, ByRef cmddata() As Double) As Integer
        Dim i As Integer
        Dim strArray(40), cmdstr As String
        cmdstr = commandstr

        cmdtype = -1
        For i = 0 To CmdSize
            If CmdName(i) <> "" Then
                If cmdstr.IndexOf(CmdName(i)) = 0 Then
                    cmdtype = i
                    Exit For
                End If
            End If
        Next i

        Select Case (cmdtype)
            Case 1 'Moveabs
                cmdstr = cmdstr.Replace(CmdName(cmdtype), "")
                cmdstr = cmdstr.Replace("(", "")
                cmdstr = cmdstr.Replace(")", "")
                strArray = cmdstr.Split(","c)
                cmddata(0) = strArray.Length
                For i = 0 To strArray.Length - 1
                    cmddata(i + 1) = System.Double.Parse(strArray(i))
                Next i
            Case 2, 3, 4, 5, 10, 17 'arc, helical, spline, QC, base, DIO
                cmdstr = cmdstr.Replace(CmdName(cmdtype), "")
                cmdstr = cmdstr.Replace("(", "")
                cmdstr = cmdstr.Replace(")", "")
                strArray = cmdstr.Split(","c)
                For i = 0 To strArray.Length - 1
                    cmddata(i) = System.Double.Parse(strArray(i))
                Next i
            Case 11, 12, 13 'Speed, Accel, Decel
                cmdstr = cmdstr.Replace(CmdName(cmdtype), "")
                cmdstr = cmdstr.Replace("=", "")
                cmddata(0) = System.Double.Parse(cmdstr)
            Case 14
            Case 15
            Case 16, 19, 20, 21 'wait, wait until, subs, sube
                cmdstr = cmdstr.Replace(CmdName(cmdtype), "")
                cmdstr = cmdstr.Replace("(", "")
                cmdstr = cmdstr.Replace(")", "")
                cmddata(0) = System.Double.Parse(cmdstr)
            Case 18 'STOP
                cmddata(0) = -1
            Case 22
                cmdstr = cmdstr.Replace(CmdName(cmdtype), "")
                cmdstr = cmdstr.Replace("(", "")
                cmdstr = cmdstr.Replace(")", "")
                strArray = cmdstr.Split(","c)
                For i = 0 To strArray.Length - 1
                    cmddata(i) = System.Double.Parse(strArray(i))
                Next i
            Case Else
                Return -1
        End Select

        Return 0
    End Function

    'Download motion data to trio controller
    '   CmdList: dispensing element list

    Public Function BurnTable(ByVal dispenselist As ArrayList) As Boolean
        If dispenselist.Count > 999 Then
            Dim download_number As Integer = OnePageElements - 2
            Dim downloads As Integer = Math.Ceiling(dispenselist.Count / download_number)
            Dim page_number, table_pos, count, values_left As Integer
            Dim cmddata(40) As Double
            Dim Buffer(OnePageElements - 1) As Single
            Dim download_page_number = 1
            count = 0
            m_Tri.RunTrioMotionProgram("CLEARTABLE", 2)
            Try
                For page_number = 1 To downloads
                    If CheckButtonState() = -1 Then Return False
                    table_pos = StartPosition + OnePageElements * page_number

                    'i.e. if onePageElements is 1000, we download 998 values of m_dispenselist2 into Buffer. 
                    'if m_dispenselist2 has 2500 elements, on our last download, we set the entire array to 0 values, then download the last 500
                    values_left = dispenselist.Count Mod download_number
                    If page_number = downloads And values_left <> 0 Then
                        Array.Clear(Buffer, 0, OnePageElements)
                        For buffer_index As Integer = 1 To values_left
                            Buffer(buffer_index) = dispenselist(count)
                            count += 1
                        Next
                    Else
                        For buffer_index As Integer = 1 To download_number
                            Buffer(buffer_index) = dispenselist(count)
                            count += 1
                        Next
                    End If

                    If DownloadOnePageTable(download_page_number, table_pos, Buffer) = False Then
                        Return False
                    End If

                    If page_number = 1 Then
                        Dim rtn As Boolean = False
                        Dim counter As Integer = 1
                        Do
                            counter += 1
                            rtn = m_Tri.RunTrioMotionProgram("DISPENSE", 8)
                        Loop Until rtn = True Or counter = 5 Or m_Tri.EStopActivated
                        If rtn = False Or m_Tri.EStopActivated Then
                            Return False
                        End If
                        SetLampsToRunningMode()
                        LabelMessage("Start Dispensing")
                        'Vision.FrmVision.DisplayIndicator()
                    End If

                Next
                LabelMessage("Download complete")
                Return True

            Catch ex As ThreadAbortException
            Catch ex As Exception
                ExceptionDisplay(ex)
            End Try
        Else
            If DownloadToTable_Below1000(dispenselist) Then
                Return True
            End If
            Return False
        End If
    End Function

    'Download one page motion data to trio controller
    '   pageno: current page no.
    '   tablepos: current Vr index
    '   bufferpos: current index of one page buffer


    Public Function DownloadOnePageTable(ByRef download_page_number As Integer, ByRef tablepos As Integer, ByRef buffer() As Single) As Boolean

        Dim rtn As Boolean = False
        Dim firstpos, firstEleValue As Integer
        Dim runmode As Integer = 0
        Dim counter As Integer = 1
        Dim readBuff(OnePageElements) As Single

        firstpos = (download_page_number - 1) * OnePageElements + 1 + StartPosition
        readBuff(0) = 1.0

        Do  ' wait for writable
            If (m_Tri.m_TriCtrl.GetTable(firstpos, 1, readBuff) = False) Then
                firstEleValue = 1
            Else
                firstEleValue = CInt(readBuff(0))
                If (firstEleValue = -1) Then Return -1
            End If
            TraceDoEvents()
        Loop While firstEleValue <> 0

        buffer(0) = 1.0
        buffer(OnePageElements - 1) = 9999.0

        Do
            counter += 1
            rtn = m_Tri.m_TriCtrl.SetTable(firstpos, OnePageElements, buffer)
        Loop Until rtn = True Or counter = 5 Or m_Tri.EStopActivated

        If m_Tri.EStopActivated Then
            Return False
        End If

        download_page_number = download_page_number + 1
        If download_page_number > MaxPages Then
            download_page_number = 1
            tablepos = 2 + StartPosition
        Else
            tablepos = tablepos + 2
        End If
        Return True
    End Function

    Public Function DownloadOnePageTable(ByRef pageno As Integer, ByRef tablepos As Integer, ByRef bufferpos As Integer, ByRef buffer() As Single) As Integer
        Dim firstpos, firstEleValue As Integer
        Dim runmode As Integer = 0
        Dim readBuff(OnePageElements) As Single

        firstpos = (pageno - 1) * OnePageElements + 1 + StartPosition
        readBuff(0) = 1.0

        Do  ' wait for writable
            'If (m_TriMemHandle.GetVr(firstpos, firstpos, readBuff(0)) = False) Then
            If (m_Tri.m_TriCtrl.GetTable(firstpos, 1, readBuff) = False) Then
                firstEleValue = 1
            Else
                firstEleValue = CInt(readBuff(0))
                If (firstEleValue = -1) Then Return -1
            End If
            TraceDoEvents()
        Loop While firstEleValue <> 0

        buffer(0) = 1.0
        buffer(OnePageElements - 1) = 9999.0
        m_Tri.m_TriCtrl.SetTable(firstpos, OnePageElements, buffer)
        pageno = pageno + 1
        If pageno > MaxPages Then
            pageno = 1
            tablepos = 2 + StartPosition
            bufferpos = 1
        Else
            tablepos = tablepos + 2
            bufferpos = 1
        End If
        Return 0
    End Function
    'When less than 1000 items
    'Always download table to the from begin
    'This function cannot be used at the moment unless the dispense algorithm in Motion perfect was revised.
    Public Function DownloadToTable_Below1000(ByVal dispenselist As ArrayList) As Boolean
        If dispenselist.Count > 999 Then
            Return False
        End If
        Dim startTime As Long = DateTime.Now.Ticks
        If Not (m_Tri.RunTrioMotionProgram("CLEARTABLE", 2)) Then
            If Not (m_Tri.RunTrioMotionProgram("CLEARTABLE", 2)) Then Return False
        End If
        Dim firstSet = DateTime.Now.Ticks
        Dim Buffer(dispenselist.Count) As Single
        Buffer(0) = 1
        For i As Integer = 0 To dispenselist.Count - 1
            Buffer(i + 1) = dispenselist(i)
        Next

        If Not m_Tri.m_TriCtrl.SetTable(1001, dispenselist.Count + 1, Buffer) Then
            If Not m_Tri.m_TriCtrl.SetTable(1001, dispenselist.Count + 1, Buffer) Then
                Return False
            End If
        End If

        'Dim tail(0) As Double
        'tail(0) = 9999.0
        'Buffer(0) = 9999.0
        'If Not m_Tri.m_TriCtrl.SetTable(2000, 1, tail) Then
        '    If Not m_Tri.m_TriCtrl.SetTable(2000, 1, tail) Then
        '        Return False
        '    End If
        'End If
        Dim secSet As Long = DateTime.Now.Ticks
        'Console.WriteLine("#1" + ((firstSet - startTime) / 10000).ToString() + "  #2" + ((secSet - firstSet) / 10000).ToString())
        Dim counter As Integer = 0
        Dim rtn As Boolean
        Do
            counter += 1
            rtn = m_Tri.RunTrioMotionProgram("DISPENSE", 8)
        Loop Until rtn = True Or counter = 5 Or m_Tri.EStopActivated
        If rtn = False Or m_Tri.EStopActivated Then
            Return False
        End If
        SetLampsToRunningMode()
        LabelMessage("Start Dispensing")
        Return True
    End Function
End Class