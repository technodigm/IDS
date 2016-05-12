Imports Microsoft.DirectX
Imports Microsoft.DirectX.DirectInput
Imports DLL_Export_Service

'Execution global functions                                                                      '
'  Created by:                                                                                   ' 
'           Shen Jian  / Jiang Ting Ying 

Public Module ModuleGConst



    'Set system setup parameters to execution global varibles         '
    '                                                                 '
    Enum SettingsMode 'yy added to make sure gLeftNeedleOffset() wont be override when loading local settings file
        GlobalSettings
        LocalSettings
    End Enum

    Public Function SystemSetupDataRetrieve(ByVal mode As SettingsMode) As Integer
        'Read system setting parameters
        IDS.Data.OpenData()

        'Run mode = "Operator" for production; "Programmer" for programming
        gExeMode = IDS.Data.Admin.User.RunApplication

        'Gantry

        IDS.Data.Hardware.Gantry.MaxSpeedLimit = IDS.Data.Hardware.Gantry.MaxSpeedLimit
        IDS.Data.Hardware.Gantry.ElementXYSpeed = IDS.Data.Hardware.Gantry.ElementXYSpeed
        IDS.Data.Hardware.Gantry.ElementZSpeed = IDS.Data.Hardware.Gantry.ElementZSpeed
        IDS.Data.Hardware.Gantry.ServiceXYSpeed = IDS.Data.Hardware.Gantry.ServiceXYSpeed
        IDS.Data.Hardware.Gantry.ServiceZSpeed = IDS.Data.Hardware.Gantry.ServiceZSpeed

        ' Work origin and area wrf hard home
        gSystemOrigin(0) = IDS.Data.Hardware.Gantry.SystemOriginPos.X
        gSystemOrigin(1) = IDS.Data.Hardware.Gantry.SystemOriginPos.Y
        gSystemOrigin(2) = IDS.Data.Hardware.Gantry.SystemOriginPos.Z
        gWorkLimitXmax = IDS.Data.Hardware.Gantry.SystemOriginPos.X + IDS.Data.Hardware.Gantry.WorkArea.X
        gWorkLimitXmin = IDS.Data.Hardware.Gantry.SystemOriginPos.X

        gWorkLimitYmax = IDS.Data.Hardware.Gantry.SystemOriginPos.Y + IDS.Data.Hardware.Gantry.WorkArea.Y
        gWorkLimitYmin = IDS.Data.Hardware.Gantry.SystemOriginPos.Y
        gWorkLimitZmax = IDS.Data.Hardware.Gantry.WorkArea.Z.Max 'wrf hard home as workZ same as hardZ
        gWorkLimitZmin = IDS.Data.Hardware.Gantry.WorkArea.Z.Min

        'Production parameters
        gRejectAutoSkip = True

        If mode = SettingsMode.GlobalSettings Then 'yy Only override the following values if file is global settings file
            'Needle center offset wrt vision center
            gLeftNeedleOffs(0) = IDS.Data.Hardware.Needle.Left.NeedleCalibrationPosition.X
            gLeftNeedleOffs(1) = IDS.Data.Hardware.Needle.Left.NeedleCalibrationPosition.Y
            gLeftNeedleOffs(2) = IDS.Data.Hardware.Needle.Left.NeedleCalibrationPosition.Z
        End If
        'LASER/LVDT
        'Laser center offset wrt vision center
        gLaserOffX = -IDS.Data.Hardware.HeightSensor.Laser.OffsetPos.X
        gLaserOffY = -IDS.Data.Hardware.HeightSensor.Laser.OffsetPos.Y
        gLaserOffX = 74.668
        gLaserOffY = -0.553
        'Distance reading of height reference block

    End Function


    'Element dispaensing parameters structure                         '
    '                                                                 '

    Public Structure DispensePara
        Public DispenseOn As Boolean
        Public NeedleGap As Double
        Public Duration As Double
        Public TravelDelay As Double
        Public TravelSpeed As Double
        Public DeTailDist As Double
        Public ApproachHeight As Double
        Public RetractDelay As Double
        Public RetractSpeed As Double
        Public RetractHeight As Double
        Public ClearanceHeight As Double
        Public ArcRadius As Double
        Public Pitch As Double
        Public FillHeight As Double
        Public RotAngle As Double
        Public EdgeClearance As Double
        Public NoOfRun As Integer
        Public SprialIn As String
        Public Needle As String
        Public Rows As Double
        Public Columns As Double
    End Structure


    'Array data structure                                             '
    '                                                                 '

    Public Structure ArrayData
        Public FirstX As Double
        Public FirstY As Double
        Public SecondX As Double
        Public SecondY As Double
        Public DiagX As Double
        Public DiagY As Double
        Public RowNo As Integer
        Public ColumNo As Integer
        Public byColum As Boolean
        Public ByZigzag As Boolean
    End Structure


    'Fiducial compensation parameters structure                       '
    '                                                                 '

    Public Structure FidCompData
        'Main pattern fiducial 
        Public gparentFidFlag As Boolean  'true:compensation; false:no compensatation
        Public gparentFidPtX, gparentFidPtY, gparentFidPtZ As Double 'rotation point
        Public gparentOffsetX, gparentOffsetY, gparentOffsetZ As Double 'transform offset
        Public gparentCompAngle As Double  'rotation angle

        'Level1 sub pattern fiducial 
        Public parentFidFlag As Boolean   'true:compensation; false:no compensatation
        Public parentFidPtX, parentFidPtY, parentFidPtZ As Double  'rotation point
        Public parentOffsetX, parentOffsetY, parentOffsetZ As Double 'transform offset
        Public parentCompAngle As Double 'rotation angle

        'Level2 sub pattern fiducial 
        Public FidFlag As Boolean  'true:compensation; false:no compensatation
        Public FidPtX, FidPtY, FidPtZ As Double  'rotation point
        Public OffsetX, OffsetY, OffsetZ As Double  'transform offset
        Public CompAngle As Double  'rotation angle

        'Level1 sub pattern insertion transform 
        Public sub1Flag As Boolean 'true:transform; false:no transform
        Public sub1InsPtX, sub1InsPtY, sub1InsPtZ As Double  'insertion point
        Public sub1FirstPtX, sub1FirstPtY, sub1FirstPtZ As Double  'sub pattern's first dispensing point
        Public sub1InsAngle As Double  'insertion angle

        'Level2 sub pattern insertion transform 
        Public sub2Flag As Boolean  'true:transform; false:no transform
        Public sub2InsPtX, sub2InsPtY, sub2InsPtZ As Double  'insertion point
        Public sub2FirstPtX, sub2FirstPtY, sub2FirstPtZ As Double 'sub pattern's first dispensing point
        Public sub2InsAngle As Double


        'Reset all fiducial & sub  parameters                             '
        '                                                                 '

        Public Sub ClearAll()
            gparentFidFlag = False
            gparentFidPtX = 0.0
            gparentFidPtY = 0.0
            gparentFidPtZ = 0.0
            gparentOffsetX = 0.0
            gparentOffsetY = 0.0
            gparentOffsetZ = 0.0
            gparentCompAngle = 0.0
            parentFidFlag = False
            parentFidPtX = 0.0
            parentFidPtY = 0.0
            parentFidPtZ = 0.0
            parentOffsetX = 0.0
            parentOffsetY = 0.0
            parentOffsetZ = 0.0
            parentCompAngle = 0.0
            FidFlag = False
            FidPtX = 0.0
            FidPtY = 0.0
            FidPtZ = 0.0
            OffsetX = 0.0
            OffsetY = 0.0
            OffsetZ = 0.0
            CompAngle = 0.0
            sub1Flag = False
            sub1InsPtX = 0.0
            sub1InsPtY = 0.0
            sub1InsPtZ = 0.0
            sub1InsAngle = 0.0
            sub2Flag = False
            sub2InsPtX = 0.0
            sub2InsPtY = 0.0
            sub2InsPtZ = 0.0
            sub2InsAngle = 0.0
        End Sub


        'Reset main fiducial parameters                                   '
        '                                                                 '

        Public Sub ClearGparentFid()
            gparentFidFlag = False
            gparentFidPtX = 0.0
            gparentFidPtY = 0.0
            gparentFidPtZ = 0.0
            gparentOffsetX = 0.0
            gparentOffsetY = 0.0
            gparentOffsetZ = 0.0
            gparentCompAngle = 0.0
        End Sub


        'Reset level1 sub fiducial parameters                             '
        '                                                                 '

        Public Sub ClearParentFid()
            parentFidFlag = False
            parentFidPtX = 0.0
            parentFidPtY = 0.0
            parentFidPtZ = 0.0
            parentOffsetX = 0.0
            parentOffsetY = 0.0
            parentOffsetZ = 0.0
            parentCompAngle = 0.0
        End Sub


        'Reset level2 fiducial parameters                                 '
        '                                                                 '

        Public Sub ClearFid()
            FidFlag = False
            FidPtX = 0.0
            FidPtY = 0.0
            FidPtZ = 0.0
            OffsetX = 0.0
            OffsetY = 0.0
            OffsetZ = 0.0
            CompAngle = 0.0
        End Sub


        'Reset level1 sub parameters                                      '
        '                                                                 '

        Public Sub ClearSub1()
            sub1Flag = False
            sub1InsPtX = 0.0
            sub1InsPtY = 0.0
            sub1InsPtZ = 0.0
            sub1FirstPtX = 0.0
            sub1FirstPtY = 0.0
            sub1FirstPtZ = 0.0
            sub1InsAngle = 0.0
        End Sub


        'Reset level2 sub parameters                                      '

        Public Sub ClearSub2()
            sub2Flag = False
            sub2InsPtX = 0.0
            sub2InsPtY = 0.0
            sub2InsPtZ = 0.0
            sub2FirstPtX = 0.0
            sub2FirstPtY = 0.0
            sub2FirstPtZ = 0.0
            sub2InsAngle = 0.0
        End Sub


        'Set main fiducial parameters                                     '
        '                                                                 '

        Public Sub SetGparentFid(ByVal Fid() As Double, ByVal Offset() As Double, ByVal Angle As Double)
            gparentFidFlag = True
            gparentFidPtX = Fid(0)
            gparentFidPtY = Fid(1)
            gparentFidPtZ = Fid(2)
            gparentOffsetX = Offset(0)
            gparentOffsetY = Offset(1)
            gparentOffsetZ = Offset(2)
            gparentCompAngle = Angle
        End Sub


        'Set level1 sub fiducial parameters                               '
        '                                                                 '

        Public Sub SetParentFid(ByVal Fid() As Double, ByVal Offset() As Double, ByVal Angle As Double)
            parentFidFlag = True
            parentFidPtX = Fid(0)
            parentFidPtY = Fid(1)
            parentFidPtZ = Fid(2)
            parentOffsetX = Offset(0)
            parentOffsetY = Offset(1)
            parentOffsetZ = Offset(2)
            parentCompAngle = Angle
        End Sub


        'Set level2 sub fiducial parameters                               '
        '                                                                 '

        Public Sub SetFid(ByVal Fid() As Double, ByVal Offset() As Double, ByVal Angle As Double)
            FidFlag = True
            FidPtX = Fid(0)
            FidPtY = Fid(1)
            FidPtZ = Fid(2)
            OffsetX = Offset(0)
            OffsetY = Offset(1)
            OffsetZ = Offset(2)
            CompAngle = Angle
        End Sub


        'Set level1 sub parameters                                        '
        '                                                                 '

        Public Sub SetSub1(ByVal insPt() As Double, ByVal firstPt() As Double, ByVal insAngle As Double)
            sub1Flag = True
            sub1InsPtX = insPt(0)
            sub1InsPtY = insPt(1)
            sub1InsPtZ = insPt(2)
            sub1FirstPtX = firstPt(0)
            sub1FirstPtY = firstPt(1)
            sub1FirstPtZ = firstPt(2)
            sub1InsAngle = insAngle
        End Sub


        'Set level2 sub parameters                                        '
        '                                                                 '

        Public Sub SetSub2(ByVal insPt() As Double, ByVal firstPt() As Double, ByVal insAngle As Double)
            sub2Flag = True
            sub2InsPtX = insPt(0)
            sub2InsPtY = insPt(1)
            sub2InsPtZ = insPt(2)
            sub2FirstPtX = firstPt(0)
            sub2FirstPtY = firstPt(1)
            sub2FirstPtZ = firstPt(2)
            sub2InsAngle = insAngle
        End Sub

    End Structure

    Public gExeMode As String = "Operator"  ' "Operator": production mode; "Programmer": programming mode
    Public Const gPrecision As Double = 0.3 '0.05  'Arc line segment fitting tolerance 
    Public Const gMaxRows As Integer = 30000 ' max sheet rows
    Public Const gMaxColumns As Integer = 30 ' max sheet column

    Public gMinArcRadius As Double = 0.5
    Public gNeedleRadius As Double = 0.1  '0.5

    Public Const gMaxElementButtons As Integer = 21
    Public Const gMaxReferButtons As Integer = 3

    Public gPatternFileDir As String = "C:\IDS\Pattern_Dir"
    Public gFidFileName As String = ""
    Public gRejectAutoSkip As Boolean = False

    'Work area limits wrt hard home
    Public gWorkLimitXmax As Double
    Public gWorkLimitXmin As Double
    Public gWorkLimitYmax As Double
    Public gWorkLimitYmin As Double
    Public gWorkLimitZmax As Double
    Public gWorkLimitZmin As Double

    'Laser/LVDT
    Public gLaserOffX As Double = 74.668
    Public gLaserOffY As Double = -0.553

    Public gSystemOrigin(2) As Double
    Public gLeftNeedleOffs() As Double = {-64.5, -199.4, -73.8}     'x,y wrt camera origin
    'z component wrt hard_home Z = Zneedle + Zdist

    'Element dispensing parameter columns arrangement 
    Public gCommandNameColumn As Integer = 1
    Public gNeedleColumn As Integer = 2
    Public gDispensColumn As Integer = 3
    Public gFid1Column As Integer = 2
    Public gFid2Column As Integer = 3
    Public gSubnameColumn As Integer = 3
    Public gArraynameColumn As Integer = 3
    Public gSetIOOnOffColumn As Integer = 3
    Public gPos1XColumn As Integer = 4
    Public gPos1YColumn As Integer = 5
    Public gPos1ZColumn As Integer = 6
    Public gPos2XColumn As Integer = 7
    Public gPos2YColumn As Integer = 8
    Public gPos2ZColumn As Integer = 9
    Public gPos3XColumn As Integer = 10
    Public gPos3YColumn As Integer = 11
    Public gPos3ZColumn As Integer = 12

    Public gDotArrayRowsColumn As Integer = 13
    Public gTravelSpeedColumn As Integer = 13
    Public gNeedleGapColumn As Integer = 14
    Public gDurationColumn As Integer = 15
    Public gTravelDelayColumn As Integer = 16
    Public gRetractDelayColumn As Integer = 17
    Public gApproachHtColumn As Integer = 18
    Public gRetractSpeedColumn As Integer = 19
    Public gRetractHtColumn As Integer = 20

    Public gClearanceHtColumn As Integer = 21
    Public gDTaiDistColumn As Integer = 22
    Public gArcRadiusColumn As Integer = 23
    Public gPitchColumn As Integer = 24
    Public gFillHeightColumn As Integer = 25

    Public gNoOfRunColumn As Integer = 26
    Public gSprialColumn As Integer = 27
    Public gRtAngleColumn As Integer = 28
    Public gEdgeClearColumn As Integer = 29
    Public gIONumberColumn As Integer = 30

    'Vison parameter
    Public gOpenColumn As Integer = 31
    Public gCloseColumn As Integer = 32
    Public gBlackDotColumn As Integer = 33
    Public gBinarizedColumn As Integer = 34
    Public gBrightnessColumn As Integer = 35
    Public gCheckBoxColumn As Integer = 36
    Public gCwCCwColumn As Integer = 37
    Public gDispModelColumn As Integer = 38
    Public gInOutColumn As Integer = 39
    Public gMainEdgeColumn As Integer = 40
    Public gPointX1Column As Integer = 41
    Public gPointX2Column As Integer = 42
    Public gPointX3Column As Integer = 43
    Public gPointX4Column As Integer = 44
    Public gPointX5Column As Integer = 45
    Public gPointY1Column As Integer = 46
    Public gPointY2Column As Integer = 47
    Public gPointY3Column As Integer = 48
    Public gPointY4Column As Integer = 49
    Public gPointY5Column As Integer = 50
    Public gPosColumn As Integer = 51
    Public gPosXColumn As Integer = 52
    Public gPosYColumn As Integer = 53
    Public gROIColumn As Integer = 54
    Public gRotColumn As Integer = 55
    Public gSizeColumn As Integer = 56
    Public gSizeXColumn As Integer = 57
    Public gSizeYColumn As Integer = 58
    Public gThresholdColumn As Integer = 59 'For Fiducial no.2's brightness
    Public gVerticalColumn As Integer = 60
    'QC
    Public gCompactnessColumn As Integer = 61
    Public gMaxAreaColumn As Integer = 62
    Public gMinAreaColumn As Integer = 63
    Public gToleranceColumn As Integer = 64
    Public gDiameterColumn As Integer = 65
    Public gMRegionXColumn As Integer = 66
    Public gMRegionYColumn As Integer = 67
    Public gMROIxColumn As Integer = 68
    Public gMROIyColumn As Integer = 69
    Public gRoughnessColumn As Integer = 70
    Public gTypeColumn As Integer = 71
    'RM
    Public gAcceptRatioCoulumn As Integer = 72
    Public gBlackWithoutRMCoulumn As Integer = 73
    Public gBlackWithRMCoulumn As Integer = 74
    Public gWhiteWithoutRMCoulumn As Integer = 75
    Public gWhiteWithRMCoulumn As Integer = 76
    Public gWoRMCoulumn As Integer = 77
    Public gWRMCoulumn As Integer = 78

    'DIO configure 
    Public gIOXminusLimit As Integer = 0
    Public gIOXplusLimit As Integer = 1
    Public gIOYminusLimit As Integer = 2
    Public gIOYplusLimit As Integer = 3
    Public gIOZminusLimit As Integer = 5
    Public gIOZplusLimit As Integer = 4
    Public gIOXDatum As Integer = 0
    Public gIOYDatum As Integer = 2
    Public gIOZDatum As Integer = 5

    Public gIOLVDTProbe As Integer = 8
    Public gIORNeedleSlide As Integer = 17 '33 output
    Public gIOLNeedleSlide As Integer = 18 '34 output
    Public gIORNeedleSlideExt As Integer = 11 '33 intput
    Public gIOLNeedleSlideExt As Integer = 9 '34 input

    Public gIOClean As Integer = 19
    Public gIORightNeedle As Integer = 24
    Public gIOLeftNeedle As Integer = 25
    Public gIORightNeedleValve As Integer = 20
    Public gIOLeftNeedleValve As Integer = 21
    Public gIODispensingDone As Integer = 26
    Public gIOloadReady As Integer = 12

    Public gIOTrackBallXplus As Integer = 44 '28
    Public gIOTrackBallXminus As Integer = 45 '29
    Public gIOTrackBallYplus As Integer = 46 '30
    Public gIOTrackBallYminus As Integer = 47 '31

    'PC DIO
    Public gPCioPort As Integer = 1

    'AIO configure
    Public gAIOLaser As Integer = 32   'Lase or LVDT  share

    'Pattern Programmming commanddefine
    Public gFidCmdIndex As Integer = 1
    Public gReferenceCmdIndex As Integer = 2
    Public gHeightCmdIndex As Integer = 3
    Public gRejectCmdIndex As Integer = 4

    Public gDotCmdIndex As Integer = 1
    Public gLineCmdIndex As Integer = 2
    Public gArcCmdIndex As Integer = 3

    Public gRectangleCmdIndex As Integer = 4
    Public gCircleCmdIndex As Integer = 5
    Public gFillRectangleCmdIndex As Integer = 6
    Public gFillCircleCmdIndex As Integer = 7
    Public gLinkCmdIndex As Integer = 8

    Public gChipEdgeCmdIndex As Integer = 9
    Public gMoveCmdIndex As Integer = 10
    Public gWaitCmdIndex As Integer = 11
    Public gPurgeCmdIndex As Integer = 12
    Public gCleanCmdIndex As Integer = 13
    Public gQCCmdIndex As Integer = 14
    Public gSubPatternCmdIndex As Integer = 15
    Public gArrayCmdIndex As Integer = 16
    Public gSetIOCmdIndex As Integer = 17
    Public gGetIOCmdIndex As Integer = 18

    Public gSeperatorCmdIndex As Integer = 19
    Public gOffsetCmdIndex As Integer = 20
    Public gMeasurementCmdIndex As Integer = 21

End Module

'Math global functions                                                                           '
'  Created by:                                                                                   ' 
'           Shen Jian         

Public Module MathFunction


    '3D point structure                                               '
    '                                                                 '

    Public Structure Point3d
        Public DispenseOn As Boolean
        Public x, y, z As Double
    End Structure


    'Display pattern file compiling error message                     '
    '   sheetname:  sheet name that dispensing command line locates at'
    '   row:        row number                                        '
    '   error_id:   error id                                          ' 
    '                                                                 '

    Public Sub CompileErrorDisplay(ByVal sheetname As String, ByVal row As Integer, ByVal error_id As Integer)
        Dim ErrorMsg As String() = {"Out of range !", _
                                    "Missing subfilename !", _
                                    "Can't find subsheet !", _
                                    "Empty sheet !", _
                                    "No dispensing elements !", _
                                    "Fiducical  check  fail !", _
                                    "Height detection fail !", _
                                    "QC check fail !", _
                                    "Chip edge detection fail!", _
                                    "Two levels sub only!", _
                                    "Offset error!", _
                                    "Detail parameter error!", _
                                    "Arc or circle endpoints error!", _
                                    "Fill spiral error!", _
                                    "IO number must be between 32 and 47!" _
                                   }
        If ProgrammingMode() Then
            Dim errStr As String = "Sheet name: " + sheetname + "; Row: " + row.ToString + "; Error: " + ErrorMsg(error_id)
            MyMsgBox(errStr, MsgBoxStyle.OKOnly, "Compiling error")
        End If
    End Sub


    'Inch convert to mm                                               '
    '   Pos:  position to be converted                                '
    '                                                                 '

    Public Sub inchTomm(ByRef Pos() As Double)
        Pos(0) = Pos(0) * 25.4
        Pos(1) = Pos(1) * 25.4
        Pos(2) = Pos(2) * 25.4
    End Sub


    'mm convert to inch                                               '
    '   Pos:  position to be converted                                '
    '                                                                 '

    Public Sub mmToinch(ByRef Pos() As Double)
        Pos(0) = Pos(0) / 25.4
        Pos(1) = Pos(1) / 25.4
        Pos(2) = Pos(2) / 25.4
    End Sub


    'XY range checking                                                '
    '   Pt:  position to be checked                                   '
    '                                                                 '

    Public Function WorkAreaErrorCheckXY(ByVal Pt() As Double) As Boolean
        Dim errorON As Boolean = False

        If (Pt(0) > gWorkLimitXmax) Or (Pt(0) < gWorkLimitXmin) Then
            Return False
        ElseIf (Pt(1) > gWorkLimitYmax) Or (Pt(1) < gWorkLimitYmin) Then
            Return False
        End If
        Return True
    End Function


    'XYZ range checking                                               '
    '   Pt:  position to be checked                                   '
    '                                                                 '

    Public Function WorkAreaErrorCheckXYZ(ByVal runMode As Integer, ByVal Pt() As Double) As Boolean
        Dim errorON As Boolean = False

        If (Pt(0) > gWorkLimitXmax) Or (Pt(0) < gWorkLimitXmin) Then
            Return False
        ElseIf (Pt(1) > gWorkLimitYmax) Or (Pt(1) < gWorkLimitYmin) Then
            Return False
        End If
        If runMode <> 0 Then
            If (Pt(2) > gWorkLimitZmax) Or (Pt(2) < gWorkLimitZmin) Then
                Return False
                'jan 20 test - need to take into account revised z value for jetting
            End If
        End If
        Return True
    End Function


    'Z range checking                                                 '
    '   z:  position to be checked                                    '
    '                                                                 '
    Public Function WorkAreaErrorCheckZ(ByVal z As Double) As Boolean
        If (z > gWorkLimitZmax) Or (z < gWorkLimitZmin) Then
            Return False
            'jan 20 test
        End If
        Return True
    End Function


    '2d translation operation                                         '
    '   pt:     position to be translated                             '
    '   offset: translation values                                    '
    '   newPt:  translated position                                   '
    '                                                                 '

    Public Sub Translate2D(ByVal pt() As Double, ByVal offset() As Double, ByRef newPt() As Double)
        Dim i As Integer
        Dim tempt(3) As Double
        pt.CopyTo(tempt, 0)
        For i = 0 To 1
            newPt(i) = tempt(i) + offset(i)
        Next
    End Sub


    '3d translation operation                                         '
    '   pt:     position to be translated                             '
    '   offset: translation values                                    '
    '   newPt:  translated position                                   '
    '                                                                 '

    Public Sub Translate3D(ByVal pt() As Double, ByVal offset() As Double, ByRef newPt() As Double)
        Dim i As Integer
        Dim tempt(3) As Double
        pt.CopyTo(tempt, 0)
        For i = 0 To 2
            newPt(i) = tempt(i) + offset(i)
        Next
    End Sub


    '2d rotation operation                                            '
    '   pt:     position to be rotated                                '
    '   angle:  rotation angle                                        '
    '   newPt:  rotated position                                      '
    '                                                                 '

    Public Sub Rotate2D(ByVal pt() As Double, ByVal angle As Double, ByRef newPt() As Double)
        Dim x As Double = pt(0)
        Dim y As Double = pt(1)
        newPt(0) = x * Math.Cos(angle) - y * Math.Sin(angle)
        newPt(1) = x * Math.Sin(angle) + y * Math.Cos(angle)
    End Sub


    '3d rotation operation                                            '
    '   pt:     position to be rotated                                '
    '   Axis:   axis to rotate around, 0: X axis, 1: Y axis, 2: Z axis' 
    '   angle:  rotation angle                                        '
    '   newPt:  rotated position                                      '
    '                                                                 '

    Public Function Rotate3D(ByVal pt() As Double, ByVal angle As Double, ByVal Axis As Integer, ByRef newPt() As Double) As Integer
        Dim x As Double = pt(0)
        Dim y As Double = pt(1)
        Dim z As Double = pt(2)

        Select Case (Axis)
            Case 0 'X'
                newPt(0) = x
                newPt(1) = y * Math.Cos(angle) + z * Math.Sin(angle)
                newPt(2) = -y * Math.Sin(angle) + z * Math.Cos(angle)

            Case 1 'Y
                newPt(0) = x * Math.Cos(angle) - z * Math.Sin(angle)
                newPt(1) = y
                newPt(2) = x * Math.Sin(angle) + z * Math.Cos(angle)

            Case 2 'Z
                newPt(0) = x * Math.Cos(angle) + y * Math.Sin(angle)
                newPt(1) = -x * Math.Sin(angle) + y * Math.Cos(angle)
                newPt(2) = z

            Case Else
                Return 1
        End Select
        Return 0
    End Function


    'Coordinate transform from hard home to work origin (system)      '
    '   pt:     position to be transformed                            '
    '   newPt:  position under work coordinate system                 '

    Public Sub HardToSys(ByVal pt() As Double, ByRef newPt() As Double)
        Dim i As Integer
        Dim tempt(3) As Double
        pt.CopyTo(tempt, 0)
        For i = 0 To 2
            newPt(i) = tempt(i) - gSystemOrigin(i) ' - gReferencePt(i)
        Next
    End Sub


    'Coordinate transform from work origin (system) to hard home      '
    '   pt:     position to be transformed                            '
    '   newPt:  position under hard home coordinate system            '

    Public Sub SysToHard(ByVal pt() As Double, ByRef newPt() As Double)  'work to hard
        Dim i As Integer
        Dim tempt(3) As Double
        pt.CopyTo(tempt, 0)
        For i = 0 To 2
            newPt(i) = tempt(i) + gSystemOrigin(i)
        Next
    End Sub

    'Coordinate transform from reference to work coordinate           '
    '   pt:      position to be transformed                           '
    '   newPt:   position under work coordinate system                '
    '   referpt: reference point under work coordinate                '

    Public Sub ReferToSys(ByVal pt() As Double, ByRef newPt() As Double, ByVal referpt() As Double)
        Dim i As Integer
        Dim tempt(3) As Double
        pt.CopyTo(tempt, 0)
        For i = 0 To 2
            newPt(i) = tempt(i) + referpt(i)
        Next
    End Sub


    'Coordinate transform from work coordinate to reference coordinate'
    '   pt:      position to be transformed                           '
    '   newPt:   position under work coordinate system                '
    '   referpt: reference point under work coordinate                '
    '                                                                 '

    Public Sub SysToRefer(ByVal pt() As Double, ByRef newPt() As Double, ByVal referpt() As Double)
        Dim i As Integer
        Dim tempt(3) As Double
        pt.CopyTo(tempt, 0)
        For i = 0 To 2
            newPt(i) = tempt(i) - referpt(i)
        Next
    End Sub


    'Fiducial compensation                                            '
    '   pt:      position to be compensated                           '
    '   FdPt:    fiducial point under work  coordinate                '
    '   offset:  shifting value                                       '
    '   angle:   fiducial mismatching angle                           '
    '   newPt:   compensated position                                 '
    '                                                                 '

    Public Sub FiducialComp(ByVal pt() As Double, ByVal FdPt() As Double, ByVal offset() As Double, ByVal angle As Double, ByRef newPt() As Double)
        Dim tempPt1(2), tempPt2(2) As Double
        'pt under fiducial point coordinate system
        tempPt1(0) = pt(0) - FdPt(0)
        tempPt1(1) = pt(1) - FdPt(1)
        'rotate
        Rotate2D(tempPt1, angle, tempPt2)
        Dim allOffset(2) As Double
        allOffset(0) = offset(0) + FdPt(0) '+ gSystemOrigin(0)
        allOffset(1) = offset(1) + FdPt(1) '+ gSystemOrigin(1)
        'new pt under work coordinate system
        Translate2D(tempPt2, allOffset, newPt)
        newPt(2) = pt(2)
    End Sub


    'Calculate two line midpoints offset and intersection angle       '
    '   L1pt1:    Line 1 endpoint 1                                   '
    '   L1pt2:    Line 1 endpoint 2                                   '
    '   L2pt1:    Line 2 endpoint 1                                   '
    '   L2pt2:    Line 2 endpoint 2                                   '
    '   offset:   two line midpoints offset                           '
    '   angle:    intersection angle                                  '
    '                                                                 '

    Public Function Get2LineAngleOffset(ByVal L1pt1() As Double, ByVal L1pt2() As Double, ByVal L2pt1() As Double, ByVal L2pt2() As Double, _
                                   ByRef offset() As Double, ByRef angle As Double) As Integer
        Dim cntPt1(2), cntPt2(2), gv1(2), gv2(2) As Double
        Dim aa As Double
        cntPt1(0) = (L1pt1(0) + L1pt2(0)) / 2.0 'Line1 midpoint
        cntPt1(1) = (L1pt1(1) + L1pt2(1)) / 2.0
        cntPt2(0) = (L2pt1(0) + L2pt2(0)) / 2.0 'Line2 midpoint
        cntPt2(1) = (L2pt1(1) + L2pt2(1)) / 2.0
        offset(0) = cntPt2(0) - cntPt1(0)  'midpoint offset
        offset(1) = cntPt2(1) - cntPt1(1)

        gv1(0) = L1pt2(0) - L1pt1(0)    'Line1 vector
        gv1(1) = L1pt2(1) - L1pt1(1)
        gv2(0) = L2pt2(0) - L2pt1(0)    'Line2 vector
        gv2(1) = L2pt2(1) - L2pt1(1)

        'angle subtended by two vectors.  
        If Dot2(gv1, gv2, aa) < 0 Then Return -1
        If (Cross2(gv1, gv2) < 0.0) Then aa = -aa 'decide angle direction
        angle = aa
    End Function

    ''''
    'Calculate point on a line segment                                    '
    '   Pt1:    line endpoint 1                                           '
    '   Pt2:    line endpoint 2                                           '
    '   Dist:   detail distance from endpoint3, > 0: outside, < 0: inside '  
    '   newPt:  calculated point on a line segment                        '
    '                                                                     '
    ''''
    Public Function PointOnLine3d(ByVal Pt1() As Double, ByVal Pt2() As Double, ByVal Dist As Double, ByRef newPt() As Double) As Integer
        'line seg length
        Dim length As Double = System.Math.Sqrt((Pt1(0) - Pt2(0)) * (Pt1(0) - Pt2(0)) _
        + (Pt1(1) - Pt2(1)) * (Pt1(1) - Pt2(1)) + (Pt1(2) - Pt2(2)) * (Pt1(2) - Pt2(2)))
        Dim ddist As Double = Dist

        If (Dist <= -length) Then
            ddist = -length + 0.001
            'Return -1
        End If

        If length < 0.001 Then
            Return -1
        Else
            newPt(0) = ((length + ddist) * Pt2(0) - ddist * Pt1(0)) / length
            newPt(1) = ((length + ddist) * Pt2(1) - ddist * Pt1(1)) / length
            newPt(2) = ((length + ddist) * Pt2(2) - ddist * Pt1(2)) / length
        End If
        Return 0
    End Function


    'Check point whether inside a line segment                        '
    '   Pt1:     line endpoint 1                                      '
    '   Pt2:     line endpoint 2                                      '
    '   checkPt: point to be checked                                  '  
    '   return:  0: inside a line segment, -1: outside                '
    '                                                                 '

    Public Function CheckPointOnLine3d(ByVal Pt1() As Double, ByVal Pt2() As Double, ByVal checkPt() As Double) As Integer
        Dim axismin(3), axismax(3) As Double
        Dim i As Integer
        For i = 0 To 2
            If Pt1(i) >= Pt2(i) Then
                axismin(i) = Pt2(i)
                axismax(i) = Pt1(i)
            Else
                axismin(i) = Pt1(i)
                axismax(i) = Pt2(i)
            End If
            If checkPt(i) > axismax(i) Or checkPt(i) < axismin(i) Then Return -1
        Next i
        Return 0
    End Function

    ''''
    'Calculate point on a arc segment                                     '
    '   Pt1:    arc endpoint 1                                            '
    '   Pt2:    arc endpoint 2                                            '
    '   Pt3:    arc endpoint 3                                            '
    '   Dist:   detail distance from endpoint3, > 0: outside, < 0: inside '  
    '   centre: arc centre                                                '
    '   radius: arc radius                                                '
    '   direction: arc direction, 0: CCW, 1: CW                           '
    '   newPt:     calculated point on a arc segment                      '
    '                                                                     '
    ''''
    Public Function PointOnArc3d(ByVal Pt1() As Double, ByVal Pt2() As Double, ByVal Pt3() As Double, ByVal Dist As Double, _
                                 ByRef centre() As Double, ByRef radius As Double, ByRef direction As Integer, ByRef newPt() As Double)
        Dim normal(3) As Double
        Dim z1, z2, z3 As Double
        z1 = Pt1(2)
        z2 = Pt1(2)
        z3 = Pt3(2)
        Pt1(2) = 0.0
        Pt2(2) = 0.0
        Pt3(2) = 0.0
        'Calculate arc centre and radius
        If Get3dCircleCentre(Pt1, Pt2, Pt3, centre, radius) < 0 Then Return -1

        'Calculate arc residing plane's normal vector
        If Get3dNormal(Pt1, Pt2, Pt3, normal) < 0 Then Return -1

        Dim rAngle As Double = Dist / radius
        Dim offset() As Double = {-centre(0), -centre(1)}
        Dim tempPt1(3), tempPt2(3) As Double
        If normal(2) < 0.0 Then  'CW
            rAngle = -rAngle
            direction = 1
        Else                     'CCW
            direction = 0
        End If
        If (Dist < 0.0) Then ' when point inside arc seg, check detail dist whether larger than arc's length
            Dim start_angle As Double = Math.Atan2(Pt1(1), Pt1(0)) 'endpoint1's angle
            If start_angle < 0.0 Then start_angle = start_angle + 2.0 * Math.PI
            Dim end_angle As Double = Math.Atan2(Pt3(1), Pt3(0))   'endpoint3's angle
            If end_angle < 0.0 Then end_angle = end_angle + 2.0 * Math.PI
            Dim angle_span As Double = 0.0
            If (direction = 0) Then 'CCW
                If end_angle <= start_angle Then
                    angle_span = end_angle + 2.0 * Math.PI - start_angle
                Else
                    angle_span = end_angle - start_angle
                End If
            Else  'CW
                If end_angle <= start_angle Then
                    angle_span = start_angle - end_angle
                Else
                    angle_span = start_angle + 2.0 * Math.PI - end_angle
                End If
            End If
            If Math.Abs(rAngle) >= Math.Abs(angle_span) Then Return -1
        End If
        Translate2D(Pt3, offset, tempPt1)
        Rotate2D(tempPt1, rAngle, tempPt2)
        offset(0) = centre(0)
        offset(1) = centre(1)
        Translate2D(tempPt2, offset, newPt)

        newPt(2) = z3
        Pt1(2) = z1
        Pt2(2) = z2
        Pt3(2) = z3
        Return 0
    End Function

    ''''
    'Calculate point on a circle                                          '
    '   Pt1:    arc endpoint 1                                            '
    '   Pt2:    arc endpoint 2                                            '
    '   Pt3:    arc endpoint 3                                            '
    '   Dist:   detail distance from endpoint3, > 0: outside, < 0: inside '  
    '   centre: arc centre                                                '
    '   radius: arc radius                                                '
    '   direction: arc direction, 0: CCW, 1: CW                           '
    '   newPt:     calculated point on a arc segment                      '
    '                                                                     '
    ''''
    Public Function PointOnCircle3d(ByVal Pt1() As Double, ByVal Pt2() As Double, ByVal Pt3() As Double, ByVal Dist As Double, _
                                 ByRef centre() As Double, ByRef radius As Double, ByRef direction As Integer, ByRef newPt() As Double)
        Dim normal(3) As Double
        Dim z1, z2, z3 As Double
        z1 = Pt1(2)
        z2 = Pt1(2)
        z3 = Pt3(2)
        Pt1(2) = 0.0
        Pt2(2) = 0.0
        Pt3(2) = 0.0
        'Calculate arc centre and radius
        If Get3dCircleCentre(Pt1, Pt2, Pt3, centre, radius) < 0 Then Return -1

        'Calculate arc residing plane's normal vector
        If Get3dNormal(Pt1, Pt2, Pt3, normal) < 0 Then Return -1

        Dim rAngle As Double = Dist / radius
        If Dist < 0.0 Then 'detail distance can't larger than circle's length
            If -rAngle >= 2.0 * Math.PI Then Return -1
        End If
        Dim offset() As Double = {-centre(0), -centre(1)}
        Dim tempPt1(3), tempPt2(3) As Double
        If normal(2) < 0.0 Then  'CW
            rAngle = -rAngle
            direction = 1
        Else                     'CCW
            direction = 0
        End If
        Translate2D(Pt1, offset, tempPt1)
        Rotate2D(tempPt1, rAngle, tempPt2)
        offset(0) = centre(0)
        offset(1) = centre(1)
        Translate2D(tempPt2, offset, newPt)
        newPt(2) = z3
        Pt1(2) = z1
        Pt2(2) = z2
        Pt3(2) = z3
        Return 0
    End Function


    'Round value to decimal 3 digits                                  '
    '   value:     value to be rounded                                '
    '                                                                 '

    Public Sub RoundTo3DecimalPoints(ByRef value As Double)
        value = CInt(value * 1000.0) / 1000.0
    End Sub


    'Round point to decimal 3 digits                                  '
    '   pt:     point to be rounded                                   '
    '                                                                 '

    Public Sub RoundTo3DecimalPoints(ByRef pt() As Double)
        Dim i As Integer
        For i = 0 To 2
            pt(i) = CInt(pt(i) * 1000.0) / 1000.0
        Next
    End Sub


    'Calculate 3d circle's centre and radius                          '
    '   Pt1:    circle endpoint 1                                     '
    '   Pt2:    circle endpoint 2                                     '
    '   Pt3:    circle endpoint 3                                     '
    '   centre: circle centre                                         '
    '   radius: circle radius                                         '
    '                                                                 '

    Public Function Get3dCircleCentre(ByVal pt1() As Double, ByVal pt2() As Double, ByVal pt3() As Double, _
        ByRef centre() As Double, ByRef radius As Double) As Integer
        Dim a, b, c, d As Double
        '3 points' plane
        Planecoef(pt1, pt2, pt3, a, b, c, d)

        Dim midpointx, midpointy, midpointz As Double
        midpointx = (pt1(0) + pt2(0)) / 2.0
        midpointy = (pt1(1) + pt2(1)) / 2.0
        midpointz = (pt1(2) + pt2(2)) / 2.0
        Dim a1, b1, c1, d1 As Double
        a1 = pt2(0) - pt1(0)
        b1 = pt2(1) - pt1(1)
        c1 = pt2(2) - pt1(2)
        d1 = -a1 * midpointx - b1 * midpointy - c1 * midpointz
        midpointx = (pt2(0) + pt3(0)) / 2.0
        midpointy = (pt2(1) + pt3(1)) / 2.0
        midpointz = (pt2(2) + pt3(2)) / 2.0
        Dim a2, b2, c2, d2 As Double
        a2 = pt3(0) - pt2(0)
        b2 = pt3(1) - pt2(1)
        c2 = pt3(2) - pt2(2)
        d2 = -a2 * midpointx - b2 * midpointy - c2 * midpointz

        Dim v_Matrix As Double = a * b1 * c2 + a1 * b2 * c + b * c1 * a2 - a2 * b1 * c - c1 * b2 * a - a1 * b * c2
        If (v_Matrix < 0.0001) Then Return -1
        Dim a11, a12, a13, a21, a22, a23, a31, a32, a33 As Double
        a11 = (b1 * c2 - c1 * b2) / v_Matrix
        a12 = (c * b2 - b * c2) / v_Matrix
        a13 = (b * c1 - c * b1) / v_Matrix
        a21 = (c1 * a2 - a1 * c2) / v_Matrix
        a22 = (a * c2 - c * a2) / v_Matrix
        a23 = (c * a1 - a * c1) / v_Matrix
        a31 = (a1 * b2 - b1 * a2) / v_Matrix
        a32 = (b * a2 - a * b2) / v_Matrix
        a33 = (a * b1 - a1 * b) / v_Matrix

        centre(0) = -(a11 * d + a12 * d1 + a13 * d2)
        centre(1) = -(a21 * d + a22 * d1 + a23 * d2)
        centre(2) = -(a31 * d + a32 * d1 + a33 * d2)

        radius = Math.Sqrt((pt1(0) - centre(0)) * (pt1(0) - centre(0)) _
                   + (pt1(1) - centre(1)) * (pt1(1) - centre(1)) _
                   + (pt1(2) - centre(2)) * (pt1(2) - centre(2)))
        If (radius > 10000.0 Or radius < 0.005) Then Return -1

        Return 0

    End Function


    'Calculate plane's unit normal                                    '
    '   Pt1:    point 1                                               '
    '   Pt2:    point 2                                               '
    '   Pt3:    point 3                                               '
    '   normal: unit normal vector                                    '
    '                                                                 '

    Public Function Get3dNormal(ByVal pt1() As Double, ByVal pt2() As Double, ByVal pt3() As Double, ByRef normal() As Double) As Integer
        Dim a1, a2, a3, b1, b2, b3, c1, c2, c3, d As Double
        a1 = pt2(0) - pt1(0)
        a2 = pt2(1) - pt1(1)
        a3 = pt2(2) - pt1(2)

        b1 = pt3(0) - pt1(0)
        b2 = pt3(1) - pt1(1)
        b3 = pt3(2) - pt1(2)

        c1 = a2 * b3 - a3 * b2
        c2 = a3 * b1 - a1 * b3
        c3 = a1 * b2 - a2 * b1
        If (Math.Abs(c1) < 0.0001 And Math.Abs(c2) < 0.0001 And Math.Abs(c3) < 0.0001) Then _
            Return -1
        d = Math.Sqrt(c1 * c1 + c2 * c2 + c3 * c3)
        normal(0) = c1 / d
        normal(1) = c2 / d
        normal(2) = c3 / d
        Return 0
    End Function


    'Calculate unit normal angle                                      '
    '   normal:    normal vector                                      '
    '   angle:     angle(0): angle wrt X axis                         '
    '              angle(1): angle wrt Z axis                         '
    '                                                                 '

    Public Function Get3dNormalAngle(ByVal normal() As Double, ByRef angle() As Double) As Integer

        Dim r As Double = Math.Sqrt(normal(0) * normal(0) + normal(1) * normal(1) + normal(2) * normal(2))
        If (r < 0.0001) Then Return -1

        angle(0) = Math.Atan2(normal(1), normal(0))

        angle(1) = Math.Acos(normal(2) / r)

        Return 0
    End Function


    'Calculate two lines intersection                                 '
    '   lineAp1:    lineA endpoint1                                   '
    '   lineAp2:    lineA endpoint2                                   '
    '   lineBp1:    lineB endpoint1                                   '
    '   lineBp2:    lineB endpoint2                                   '
    '   interS:     intersection point                                '
    '                                                                 '

    Public Function LineIntersection(ByVal lineAp1() As Double, ByVal lineAp2() As Double, ByVal lineBp1() As Double, ByVal lineBp2() As Double, ByRef interS() As Double) As Integer
        Dim aM, bM, aB, bB, isX, isY As Double
        isX = 0.0
        isY = 0.0
        'check line segment length
        If (Math.Abs(lineAp2(0) - lineAp1(0)) < 0.0001) And (Math.Abs(lineAp2(1) - lineAp1(1)) < 0.0001) Then
            Return -1
        ElseIf (Math.Abs(lineBp2(0) - lineBp1(0)) < 0.0001) And (Math.Abs(lineBp2(1) - lineBp1(1)) < 0.0001) Then
            Return -1
        End If

        If (Math.Abs(lineAp2(0) - lineAp1(0)) < 0.0001) Then 'lineA is vertical
            If (Math.Abs(lineBp2(0) - lineBp1(0)) < 0.0001) Then Return -1
            isX = lineAp1(0)
            bM = (lineBp2(1) - lineBp1(1)) / (lineBp2(0) - lineBp1(0))
            bB = lineBp2(1) - bM * lineBp2(0)
            isY = bM * isX + bB
        ElseIf (Math.Abs(lineBp2(0) - lineBp1(0)) < 0.0001) Then 'lineB is vertical
            isX = lineBp1(0)
            aM = (lineAp2(1) - lineAp1(1)) / (lineAp2(0) - lineAp1(0))
            aB = lineAp2(1) - aM * lineAp2(0)
            isY = aM * isX + aB
        Else
            aM = (lineAp2(1) - lineAp1(1)) / (lineAp2(0) - lineAp1(0))
            bM = (lineBp2(1) - lineBp1(1)) / (lineBp2(0) - lineBp1(0))
            If (Math.Abs(aM - bM) < 0.0001) Then Return -1
            aB = lineAp2(1) - aM * lineAp2(0)
            bB = lineBp2(1) - bM * lineBp2(0)
            isX = (bB - aB) / (aM - bM)
            isY = aM * isX + aB
        End If

        interS(0) = isX
        interS(1) = isY
        Return 0
    End Function


    'Calculate line with plane's intersection point and plane's coef  '
    '   p1:    line endpoint1                                         '
    '   p2:    line endpoint2                                         '
    '   a:     coef of plane equation: ax+by+cz+d = 0                 '
    '   b:                                                            '
    '   c:                                                            '
    '   d:                                                            '
    '   interS:     intersection point                                '
    '                                                                 '

    Public Function LinePlaneIntersection(ByVal p1() As Double, ByVal p2() As Double, ByVal a As Double, ByVal b As Double, ByVal c As Double, ByVal d As Double, ByRef interS() As Double) As Integer
        Dim u1, u2, u As Double
        Dim i As Integer
        u1 = a * (p1(0) - p2(0)) + b * (p1(1) - p2(1)) + c * (p1(2) - p2(2))
        If u1 < 0.0001 Then Return -1
        u2 = a * p1(0) + b * p1(1) + c * p1(2) + d
        u = u2 / u1

        For i = 0 To 2
            interS(i) = p1(i) + u * (p2(i) - p1(i))
        Next
        Return 0
    End Function


    'Find nearest line endpoint of a given point                      '
    '   pt1:    line endpoint1                                        '
    '   pt2:    line endpoint2                                        '
    '   pt:     a given point                                         '
    '   whichPt: 1: endpoint1 is near, 2: endpoint2 is near           '
    '                                                                 '

    Public Sub GetNearestPoint(ByVal Pt1() As Double, ByVal Pt2() As Double, ByVal Pt() As Double, ByRef whichPt As Integer)
        Dim d1 As Double = (Pt1(0) - Pt(0)) * (Pt1(0) - Pt(0)) + (Pt1(1) - Pt(1)) * (Pt1(1) - Pt(1))
        Dim d2 As Double = (Pt2(0) - Pt(0)) * (Pt2(0) - Pt(0)) + (Pt2(1) - Pt(1)) * (Pt2(1) - Pt(1))
        If d1 > d2 Then
            whichPt = 2
        Else
            whichPt = 1
        End If
    End Sub


    'Get fillet minimun radius that makes smooth turning              '
    '   speed:   motion speed                                         '
    '   return:  turning radius                                         '
    '                                                                 '

    Public Function GetFilletRadius(ByVal speed As Double) As Double
        Dim recommended_radius As Double = (speed * 3.0 / 50.0)
        Dim min_radius As Double = 0.5
        If min_radius > recommended_radius Then
            Return min_radius
        Else
            Return recommended_radius
        End If
    End Function


    'Generate rectangle endpoint list                                 '
    '   pt1:    rectangle endpoint1                                   '
    '   pt2:    rectangle endpoint2                                   '
    '   pt3:    rectangle endpoint3                                   '
    '   ptlist: endpoint list                                         '
    '                                                                 '

    Public Function GetRectVertexlist(ByVal Pt1() As Double, ByVal Pt2() As Double, ByVal Pt3() As Double, ByRef Ptlist As ArrayList)
        If Ptlist.Count > 0 Then
            Ptlist.Clear()
        End If
        Dim p2(3), p4(3) As Double
        p2(0) = Pt1(0)
        p2(1) = Pt3(1)
        p2(2) = Pt1(2)
        p4(0) = Pt3(0)
        p4(1) = Pt1(1)
        p4(2) = Pt1(2)
        Ptlist.Add(Pt1)
        Dim whichPt As Integer
        GetNearestPoint(p2, p4, Pt2, whichPt)
        'add points to list
        If whichPt = 1 Then
            Ptlist.Add(p2)
            Ptlist.Add(Pt3)
            Ptlist.Add(p4)
        Else
            Ptlist.Add(p4)
            Ptlist.Add(Pt3)
            Ptlist.Add(p2)
        End If
        Ptlist.Add(Pt1)
    End Function


    'Generate diamond shape endpoint list                             '
    '   pt1:    rectangle endpoint1                                   '
    '   pt2:    rectangle endpoint2                                   '
    '   pt3:    rectangle endpoint3                                   '
    '   ptlist: endpoint list                                         '
    '                                                                 '

    Public Function GetDiaVertexlist(ByVal Pt1() As Double, ByVal Pt2() As Double, ByVal Pt3() As Double, ByRef Ptlist As ArrayList)
        If Ptlist.Count > 0 Then
            Ptlist.Clear()
        End If
        Dim p4(3), dx, dy, a, b, c, d As Double
        Planecoef(Pt1, Pt2, Pt3, a, b, c, d)
        dx = Pt2(0) - Pt1(0)
        dy = Pt2(1) - Pt1(1)
        p4(0) = Pt3(0) - dx
        p4(1) = Pt3(1) - dy
        If Math.Abs(c) > 0.0001 Then
            p4(2) = -(a * p4(0) + b * p4(1) + d) / c
        Else
            p4(2) = Pt1(2)
        End If
        'add points to list
        Ptlist.Add(Pt1)
        Ptlist.Add(Pt2)
        Ptlist.Add(Pt3)
        Ptlist.Add(p4)
        Ptlist.Add(Pt1)
        Return 0
    End Function


    'Generate diamond shape offset endpoint list (for chipedge)       '
    '   pt1:    rectangle endpoint1                                   '
    '   pt2:    rectangle endpoint2                                   '
    '   pt3:    rectangle endpoint3                                   '
    '   offset: offset value                                          '
    '   ptlist: offset endpoint list                                  '
    '                                                                 '

    Public Function GetDiaOffsetVertexlist(ByVal Pt1() As Double, ByVal Pt2() As Double, ByVal Pt3() As Double, ByVal offset As Double, ByRef Ptlist As ArrayList)
        If Ptlist.Count > 0 Then
            Ptlist.Clear()
        End If
        Dim Pt4(3), dx, dy, a, b, c, d As Double
        'plane's coef
        Planecoef(Pt1, Pt2, Pt3, a, b, c, d)
        dx = Pt2(0) - Pt1(0)
        dy = Pt2(1) - Pt1(1)
        Pt4(0) = Pt3(0) - dx
        Pt4(1) = Pt3(1) - dy
        If Math.Abs(c) > 0.0001 Then
            Pt4(2) = -(a * Pt4(0) + b * Pt4(1) + d) / c
        Else
            Pt4(2) = Pt1(2)
        End If
        Dim v1(3), v2(3) As Double
        v1(0) = Pt2(0) - Pt1(0)
        v1(1) = Pt2(1) - Pt1(1)
        v2(0) = Pt3(0) - Pt1(0)
        v2(1) = Pt3(1) - Pt1(1)
        'decide offset direction 
        Dim dir As Integer
        Dim vCross As Double = Cross2(v1, v2)
        If vCross < -0.0001 Then 'offset left
            dir = -1
        ElseIf vCross > 0.0001 Then 'offset right
            dir = 1
        Else
            Return -1
        End If
        Dim offS1(3), offS2(3), offS3(3), offS4(3), interS(3) As Double
        'pt1 pt2 offset
        If LineOffset(Pt1, Pt2, offset, dir, offS1, offS2) < 0 Then Return -1
        'pt2 pt3 offset
        If LineOffset(Pt2, Pt3, offset, dir, offS3, offS4) < 0 Then Return -1
        'two line's intersection
        If LineIntersection(offS1, offS2, offS3, offS4, interS) < 0 Then Return -1
        'add points to list
        Ptlist.Add(offS1)
        Ptlist.Add(offS2)
        Ptlist.Add(interS)
        Ptlist.Add(offS3)
        Ptlist.Add(offS4)
        Return 0
    End Function


    'Generate diamond shape offset endpoint list (for fill rectangle) '
    '   parentList:    parent endpoint list                           '
    '   offset: offset value                                          '
    '   dir:    list direction 0: can't offset further                '
    '                         -1: CCW                                 '
    '                          1: CW                                  '
    '   childList:     offset endpoint list                           '
    '                                                                 '

    Public Function Get4PtDiaOffsetVertexlist(ByVal parentList As ArrayList, ByVal offset As Double, ByRef dir As Integer, ByRef childList As ArrayList)
        If childList.Count > 0 Then
            childList.Clear()
        End If
        Dim Pt1(3), Pt2(3), Pt3(3), Pt4(3), z As Double

        If parentList.Count < 5 Then Return -1
        Dim point(7), interPoint, firstPt As Point3d
        point(0) = parentList(3)
        point(1) = parentList(0)
        point(2) = parentList(1)
        point(3) = parentList(2)
        point(4) = parentList(3)
        point(5) = parentList(4)
        z = point(0).z
        interPoint.DispenseOn = True
        interPoint.z = z
        ' calculate offset points 
        Dim offS1(3), offS2(3), offS3(3), offS4(3), interS(3) As Double
        Dim i As Integer
        For i = 0 To 3
            Pt1(0) = point(i).x
            Pt1(1) = point(i).y
            Pt1(2) = z
            Pt2(0) = point(i + 1).x
            Pt2(1) = point(i + 1).y
            Pt2(2) = z
            Pt3(0) = point(i + 2).x
            Pt3(1) = point(i + 2).y
            Pt3(2) = z
            If Math.Abs(Pt1(0) - Pt2(0)) < 0.001 And Math.Abs(Pt1(1) - Pt2(1)) < 0.001 Then
                dir = 0
                Return 0
            ElseIf Math.Abs(Pt3(0) - Pt2(0)) < 0.001 And Math.Abs(Pt3(1) - Pt2(1)) < 0.001 Then
                dir = 0
                Return 0
            End If
            If LineOffset(Pt1, Pt2, offset, dir, offS1, offS2) < 0 Then Return -1
            If LineOffset(Pt2, Pt3, offset, dir, offS3, offS4) < 0 Then Return -1
            If LineIntersection(offS1, offS2, offS3, offS4, interS) < 0 Then Return -1
            If CheckPointOnLine3d(offS1, offS2, interS) < 0 Then
                dir = 0
                Return 0
            End If
            interPoint.x = interS(0)
            interPoint.y = interS(1)
            childList.Add(interPoint)
            If i = 0 Then
                firstPt = interPoint
            End If
        Next i
        childList.Add(firstPt)
        Dim v1(3), v2(3) As Double
        v1(0) = childList(1).x - childList(0).x 'Pt2(0) - Pt1(0)
        v1(1) = childList(1).y - childList(0).y 'Pt2(1) - Pt1(1)
        v2(0) = childList(2).x - childList(0).x 'Pt3(0) - Pt1(0)
        v2(1) = childList(2).y - childList(0).y 'Pt3(1) - Pt1(1)

        'Check parent and child list's direction
        Dim vCross As Double = Cross2(v1, v2)
        If vCross < -0.0001 Then
            If dir <> 1 Then
                dir = 0
                Return 0
            End If
            dir = 1
        ElseIf vCross > 0.0001 Then
            If dir <> -1 Then
                dir = 0
                Return 0
            End If
            dir = -1
        Else
            dir = 0
        End If

        Return 0
    End Function



    'Calculate line offset endpoint                                   '
    '   Pt1:    line endpoint1                                        '
    '   Pt2:    line endpoint2                                        '
    '   offset: offset value                                          '
    '   dir:    -1: left, 1: right                                    '
    '   offS1:  offset endpoint1                                      '
    '   offS2:  offset endpoint2                                      '
    '                                                                 '

    Public Function LineOffset(ByVal Pt1() As Double, ByVal Pt2() As Double, ByVal offset As Double, ByVal dir As Integer, ByRef offS1() As Double, ByRef offS2() As Double) As Integer
        Dim newPt1(3), newPt2(3), angle As Double

        'calculate point on line seg with  offset distance from pt1
        If (PointOnLine3d(Pt2, Pt1, -offset, newPt1) < 0) Then
            Return -1
        Else 'calculate offset endpoint1
            Dim off() As Double = {-Pt1(0), -Pt1(1), 0.0}
            Dim temp(3), temp1(3) As Double

            Translate2D(newPt1, off, temp)
            If (dir > 0) Then
                angle = -Math.PI / 2.0
            Else
                angle = Math.PI / 2.0
            End If
            Rotate2D(temp, angle, temp1)
            Translate2D(temp1, Pt1, offS1)
        End If
        'calculate point on line seg with  offset distance from pt2
        If (PointOnLine3d(Pt1, Pt2, -offset, newPt1) < 0) Then
            Return -1
        Else 'calculate offset endpoint2
            Dim off() As Double = {-Pt2(0), -Pt2(1), 0.0}
            Dim temp(3), temp1(3) As Double

            Translate2D(newPt1, off, temp)
            If dir > 0 Then
                angle = Math.PI / 2.0
            Else
                angle = -Math.PI / 2.0
            End If
            Rotate2D(temp, angle, temp1)
            Translate2D(temp1, Pt2, offS2)
        End If

        Return 0
    End Function



    'Calculate line equation coef                                     '
    '   P1:     line endpoint1                                        '
    '   P2:     line endpoint2                                        '
    '   a:      coef of ax+by+c = 0                                   '
    '   b:                                                            '
    '   c:                                                            '
    '                                                                 '

    Public Sub Linecoef(ByVal p1() As Double, ByVal p2() As Double, ByRef a As Double, ByRef b As Double, ByRef c As Double)
        c = (p2(0) * p1(1)) - (p1(0) * p2(1))
        a = p2(1) - p1(1)
        b = p1(0) - p2(0)
    End Sub


    'Calculate plane equation coef                                    '
    '   P1:     point1 on a plane                                     '
    '   P2:     point2 on a plane                                     '
    '   P3:     point3 on a plane                                     '
    '   a:      coef of ax+by+cz+d = 0                                '
    '   b:                                                            '
    '   c:                                                            '
    '   d:                                                            '
    '                                                                 '

    Public Sub Planecoef(ByVal p1() As Double, ByVal p2() As Double, ByVal p3() As Double, ByRef a As Double, ByRef b As Double, ByRef c As Double, ByRef d As Double)
        Dim x11, x21, x31, y11, y21, y31, z11, z21, z31 As Double
        x11 = p1(0)
        x21 = p2(0) - p1(0)
        x31 = p3(0) - p1(0)
        y11 = p1(1)
        y21 = p2(1) - p1(1)
        y31 = p3(1) - p1(1)
        z11 = p1(2)
        z21 = p2(2) - p1(2)
        z31 = p3(2) - p1(2)

        a = y21 * z31 - y31 * z21
        b = z21 * x31 - x21 * z31
        c = x21 * y31 - x31 * y21
        d = (-y21 * z31 + y31 * z21) * x11 + (-z21 * x31 + x21 * z31) * y11 + (-x21 * y31 + x31 * y21) * z11
    End Sub



    'Calculate signed distance                                        '
    '   a:      coef of ax+by+c = 0                                   '
    '   b:                                                            '
    '   c:                                                            '
    '   p:      point                                                 '
    '   d:      signed distance from line to point p                  '
    '                                                                 '

    Public Function Linetopoint(ByVal a As Double, ByVal b As Double, ByVal c As Double, ByVal p() As Double, _
                                ByRef d As Double) As Integer

        Dim d1, lp As Double

        d1 = Math.Sqrt((a * a) + (b * b))
        If (Math.Abs(d1) < 0.0001) Then
            Return -1
        Else
            d = (a * p(0) + b * p(1) + c) / d1
        End If
        Return 0
    End Function



    'Calculate projection point of p onto a line                      '
    '   a:      coef of ax+by+c = 0                                   '
    '   b:                                                            '
    '   c:                                                            '
    '   p:      point                                                 '
    '   x:      projection point x                                    '
    '   y:      projection point y                                    '
    '                                                                 '

    Public Function Pointperp(ByVal a As Double, ByVal b As Double, ByVal c As Double, ByVal p() As Double, ByRef x As Double, ByRef y As Double) As Integer
        Dim d, cp As Double
        x = 0.0
        y = 0.0
        d = a * a + b * b
        cp = a * p(1) - b * p(0)
        If (Math.Abs(d) < 0.0001) Then
            Return -1
        Else
            x = (-a * c - b * cp) / d
            y = (a * cp - b * c) / d
        End If
        Return 0
    End Function


    'Two vectors' cross operation                                     '
    '   v1:      vector1                                              '
    '   v2:      vector2                                              '
    '   return:  cross result                                         '
    '                                                                 '

    Public Function Cross2(ByVal v1() As Double, ByVal v2() As Double) As Double
        Return (v1(0) * v2(1) - v2(0) * v1(1))
    End Function



    'Two vectors' dot operation                                       '
    '   v1:      vector1                                              '
    '   v2:      vector2                                              '
    '   angle:   angle subtended by two vectors                       '
    '                                                                 '

    Public Function Dot2(ByVal v1() As Double, ByVal v2() As Double, ByRef angle As Double) As Integer
        Dim d, t As Double
        d = Math.Sqrt(((v1(0) * v1(0)) + (v1(1) * v1(1))) * ((v2(0) * v2(0)) + (v2(1) * v2(1))))
        If (Math.Abs(d) < 0.0001) Then
            Return -1
        Else
            t = (v1(0) * v2(0) + v1(1) * v2(1)) / d
            If t > 1.0 Then
                t = 1.0
            ElseIf t < -1.0 Then
                t = -1.0
            End If
            'If Math.Abs(t) > 1.0 Then Return -1
            angle = Math.Acos(t)
        End If

        Return 0
    End Function



    'Calculate arc fillet between line l1 and l2                      '
    '   p1:      line1 endpoint1                                      '
    '   p2:      line1 endpoint2                                      '
    '   p3:      line2 endpoint1                                      '
    '   p4:      line2 endpoint2                                      '
    '   r:       fillet arc radius                                    '
    '   xc:      arc centre x                                         '
    '   yc:      arc centre y                                         '
    '   pa:      beginning angle of arc                               '
    '   aa:      angle subtended by arc                               '
    '   direction: direction of arc, 0: CCW, 1: CW                    '
    '                                                                 '

    Public Function Linelinefillet(ByVal p1() As Double, ByRef p2() As Double, ByRef p3() As Double, ByVal p4() As Double, ByVal r As Double, _
                                   ByRef xc As Double, ByRef yc As Double, ByRef pa As Double, ByRef aa As Double, ByRef direction As Integer) As Integer
        Dim a1, b1, c1, a2, b2, c2, c1p, _
            c2p, d1, d2, xa, xb, ya, yb, d, rr As Double
        Dim mp(3), pc(3), gv1(3), gv2(3) As Double

        Linecoef(p1, p2, a1, b1, c1)
        Linecoef(p3, p4, a2, b2, c2)

        If Math.Abs((a1 * b2) - (a2 * b1)) < 0.0001 Then  'Parallel or coincident lines 
            Return -1
        End If

        mp(0) = (p3(0) + p4(0)) / 2.0
        mp(1) = (p3(1) + p4(1)) / 2.0
        If Linetopoint(a1, b1, c1, mp, d1) < 0 Then 'Find distance mp to line1
            Return -1
        End If

        mp(0) = (p1(0) + p2(0)) / 2.0
        mp(1) = (p1(1) + p2(1)) / 2.0
        If Linetopoint(a2, b2, c2, mp, d2) < 0 Then 'Find distance mp to line2
            Return -1
        End If
        rr = r
        If (d1 <= 0.0) Then rr = -rr

        c1p = c1 - rr * Math.Sqrt((a1 * a1) + (b1 * b1)) 'Construct line parallel l1 at distance = rr
        rr = r

        If (d2 <= 0.0) Then rr = -rr

        c2p = c2 - rr * Math.Sqrt((a2 * a2) + (b2 * b2)) 'Construct line parallel l2 at distance = rr
        d = a1 * b2 - a2 * b1

        xc = (c2p * b1 - c1p * b2) / d   ' Intersect constructed lines 
        yc = (c1p * a2 - c2p * a1) / d   ' to find center of arc 
        pc(0) = xc
        pc(1) = yc

        Pointperp(a1, b1, c1, pc, xa, ya)  ' Clip or extend lines as required 
        Pointperp(a2, b2, c2, pc, xb, yb)

        Dim dx, dy As Double     'Check arc tangent point whether within line segment
        dx = p2(0) - xa
        dy = p2(1) - ya
        d1 = dx * dx + dy * dy
        dx = p2(0) - p1(0)
        dy = p2(1) - p1(1)
        d2 = dx * dx + dy * dy
        If (d1 >= d2) Then Return -1
        dx = p3(0) - xb
        dy = p3(1) - yb
        d1 = dx * dx + dy * dy
        dx = p3(0) - p4(0)
        dy = p3(1) - p4(1)
        d2 = dx * dx + dy * dy
        If (d1 >= d2) Then Return -1

        p2(0) = xa
        p2(1) = ya
        p3(0) = xb
        p3(1) = yb
        gv1(0) = xa - xc
        gv1(1) = ya - yc
        gv2(0) = xb - xc
        gv2(1) = yb - yc

        pa = Math.Atan2(gv1(1), gv1(0))       'Beginning angle for arc
        If Dot2(gv1, gv2, aa) < 0 Then Return -1
        If (Cross2(gv1, gv2) < 0.0) Then aa = -aa 'Angle subtended by arc
        If (aa < 0.0) Then  'direction = 0 CCW, = 1 CW
            direction = 1
        Else
            direction = 0
        End If
        Return 0
    End Function


    'Calculate 2.5D arc fillet between line l1 and l2                
    '   p1:      line1 endpoint1                                      
    '   p2:      line1 endpoint2                                      
    '   p3:      line2 endpoint1                                      
    '   p4:      line2 endpoint2                                      
    '   r:       fillet arc radius                                    
    '   pr:      previous fillet arc radius                           
    '   xc:      arc centre x                                         
    '   yc:      arc centre y                                         
    '   pa:      beginning angle of arc                               
    '   aa:      angle subtended by arc                               
    '   direction: direction of arc, 0: CCW, 1: CW                    
    '   plane:     arc plane, 1:YZ, 2: ZX                             
    '                                                                 

    Public Function Linelinefillet3d(ByVal p1() As Double, ByRef p2() As Double, ByRef p3() As Double, ByVal p4() As Double, ByVal r As Double, ByVal pr As Integer, _
                                   ByRef xc As Double, ByRef yc As Double, ByRef pa As Double, ByRef aa As Double, ByRef direction As Integer, ByRef plane As Integer) As Integer
        Dim a1, b1, c1, a2, b2, c2, c1p, _
            c2p, d1, d2, xa, xb, ya, yb, d, rr As Double
        Dim mp(3), pc(3), gv1(3), gv2(3) As Double
        Dim tmp1(3), tmp2(3), tmp3(3), tmp4(3) As Double
        plane = 0
        Dim dx As Double = p1(0) - p4(0)
        Dim dy As Double = p1(1) - p4(1)
        Dim dz As Double = p1(2) - p4(2)
        Dim dist As Double = Math.Sqrt(dx * dx + dy * dy)
        ' If (dist <= r + pr) Then Return -1 'can't insert two fillet arcs
        'Deside which plane to insert fillet arc
        If (Math.Abs(dx) > Math.Abs(dy)) Then 'Z-X plane
            If (Math.Abs(dx) <= r + pr Or Math.Abs(dz) <= r) Then Return -1 'can't insert two fillet arcs
            plane = 2
            tmp1(0) = p1(0)
            tmp1(1) = p1(2)
            tmp2(0) = p2(0)
            tmp2(1) = p2(2)
            tmp3(0) = p3(0)
            tmp3(1) = p3(2)
            tmp4(0) = p4(0)
            tmp4(1) = p4(2)
        Else  'Y-Z plane
            If (Math.Abs(dy) <= r + pr Or Math.Abs(dz) <= r) Then Return -1 'can't insert two fillet arcs
            plane = 1
            tmp1(0) = p1(1)
            tmp1(1) = p1(2)
            tmp2(0) = p2(1)
            tmp2(1) = p2(2)
            tmp3(0) = p3(1)
            tmp3(1) = p3(2)
            tmp4(0) = p4(1)
            tmp4(1) = p4(2)
        End If

        Linecoef(tmp1, tmp2, a1, b1, c1)
        Linecoef(tmp3, tmp4, a2, b2, c2)

        If Math.Abs((a1 * b2) - (a2 * b1)) < 0.0001 Then  'Parallel or coincident lines
            Return -1
        End If

        mp(0) = (tmp3(0) + tmp4(0)) / 2.0
        mp(1) = (tmp3(1) + tmp4(1)) / 2.0
        If Linetopoint(a1, b1, c1, mp, d1) < 0 Then 'Find distance p1p2 to p3 */
            Return -1
        End If

        mp(0) = (tmp1(0) + tmp2(0)) / 2.0
        mp(1) = (tmp1(1) + tmp2(1)) / 2.0
        If Linetopoint(a2, b2, c2, mp, d2) < 0 Then ' Find distance p3p4 to p2 */
            Return -1
        End If
        rr = r
        If (d1 <= 0.0) Then rr = -rr

        c1p = c1 - rr * Math.Sqrt((a1 * a1) + (b1 * b1)) ' Line parallel l1 at d */
        rr = r

        If (d2 <= 0.0) Then rr = -rr

        c2p = c2 - rr * Math.Sqrt((a2 * a2) + (b2 * b2)) 'Line parallel l2 at d */
        d = a1 * b2 - a2 * b1

        xc = (c2p * b1 - c1p * b2) / d   ' Intersect constructed lines */
        yc = (c1p * a2 - c2p * a1) / d   ' to find center of arc */
        pc(0) = xc
        pc(1) = yc

        Pointperp(a1, b1, c1, pc, xa, ya)  ' Clip or extend lines as required */
        Pointperp(a2, b2, c2, pc, xb, yb)

        dx = tmp2(0) - xa
        dy = tmp2(1) - ya
        d1 = dx * dx + dy * dy
        dx = tmp2(0) - tmp1(0)
        dy = tmp2(1) - tmp1(1)
        d2 = dx * dx + dy * dy
        If (d1 >= d2) Then Return -1 'tangent point outside line1
        dx = tmp3(0) - xb
        dy = tmp3(1) - yb
        d1 = dx * dx + dy * dy
        dx = tmp3(0) - tmp4(0)
        dy = tmp3(1) - tmp4(1)
        d2 = dx * dx + dy * dy
        If (d1 >= d2) Then Return -1 'tangent point outside line2
        Dim ca, cb, cc, cd As Double
        Planecoef(p1, p2, p4, ca, cb, cc, cd)
        If (plane = 2) Then 'XZ
            p2(0) = xa
            p2(1) = (-ca * xa - cc * ya - cd) / cb
            p2(2) = ya
            p3(0) = xb
            p3(1) = (-ca * xb - cc * yb - cd) / cb
            p3(2) = yb
        ElseIf (plane = 1) Then    'YZ
            p2(0) = (-cb * xa - cc * ya - cd) / ca
            p2(1) = xa
            p2(2) = ya
            p3(0) = (-cb * xb - cc * yb - cd) / ca
            p3(1) = xb
            p3(2) = yb
        End If

        gv1(0) = xa - xc
        gv1(1) = ya - yc
        gv2(0) = xb - xc
        gv2(1) = yb - yc

        pa = Math.Atan2(gv1(1), gv1(0))      '/* Beginning angle for arc */
        If Dot2(gv1, gv2, aa) < 0 Then Return -1
        If (Cross2(gv1, gv2) < 0.0) Then aa = -aa '/* Angle subtended by arc */
        If (aa < 0.0) Then
            direction = 1
        Else
            direction = 0
        End If

        Return 0                                  'direction = 1: CW, =0: CCW
    End Function



    'Calculate spiral point list                                      '
    '   pt1:      base circle point1                                  '
    '   pt2:      base circle point2                                  '
    '   pt3:      base circle point3                                  '
    '   r_pitch:  radius pitch                                        '
    '   noofrun:  number of runs (>=1)                                '
    '   height:   spiral height                                       '
    '   ddist:    detail distance                                     '
    '   InOut:    dispensing direction, "IN" or "OUT"                 '
    '   pointlist:   calculated point list                            '
    '                                                                 '

    Public Function GetSpiralPointList(ByVal pt1() As Double, ByVal pt2() As Double, ByVal pt3() As Double, ByVal r_pitch As Double, ByVal noofrun As Integer, ByVal height As Double, ByVal ddist As Double, ByVal InOut As String, ByRef pointlist As ArrayList) As Integer

        Dim centre(3), normal(3), angle(2), tmpt1(3), tmpt2(3), tmpt3(3), radius As Double
        Dim swapt(3) As Double
        Dim CW_CCW As Integer = 0
        If noofrun < 1 Then Return -1
        If Get3dCircleCentre(pt1, pt2, pt3, centre, radius) < 0 Then 'get circle centre and radius
            Return -1
        End If
        If Get3dNormal(pt1, pt2, pt3, normal) < 0 Then Return -1 'get plane normal
        If Get3dNormalAngle(normal, angle) < 0 Then Return -1 ' get normal angle
        If (angle(1) > Math.PI / 2.0) Or (angle(1) < -Math.PI / 2.0) Then 'CCW
            CW_CCW = 1
        Else
            CW_CCW = 0
        End If
        'rotating pt1,pt2,pt3 to align the plane normal with z axis
        Rotate3D(pt1, angle(0), 2, swapt)
        Rotate3D(swapt, angle(1), 1, tmpt1)
        Rotate3D(pt2, angle(0), 2, swapt)
        Rotate3D(swapt, angle(1), 1, tmpt2)
        Rotate3D(pt3, angle(0), 2, swapt)
        Rotate3D(swapt, angle(1), 1, tmpt3)
        'rotating centre to pt1,pt2,pt3's plane 
        Dim res_centre(3), top_centre(3) As Double
        centre.CopyTo(res_centre, 0)
        Rotate3D(centre, angle(0), 2, swapt)
        Rotate3D(swapt, angle(1), 1, centre)
        'transform to centre's coordinate
        Dim dist() As Double = {-centre(0), -centre(1), -centre(2)}
        Translate3D(tmpt1, dist, tmpt1)
        Translate3D(tmpt2, dist, tmpt2)
        Translate3D(tmpt3, dist, tmpt3)

        If (CW_CCW = 1) Then height = -height 'CCW
        Dim start_angle As Double = Math.Atan2(tmpt1(1), tmpt1(0)) 'circle starting angle
        If (gPrecision / radius > 1.0) Then Return -1 'fitting tolerance great then radius
        Dim est_angle_step As Double = 2.0 * Math.Acos(1.0 - gPrecision / radius)
        Dim step_no As Integer = CInt(2.0 * Math.PI / est_angle_step) + 1 'one run's steps or line segments
        If (step_no > 200) Then step_no = 200
        Dim actual_angle_step As Double = 2.0 * Math.PI / step_no 'one steps' actual angle
        Dim x, y As Double
        Dim z As Double = tmpt1(2)
        Dim r As Double = radius
        Dim start_z As Double = z
        Dim init_start_z As Double = start_z
        Dim end_r, end_z As Double
        Dim lastR As Double = radius - (noofrun) * r_pitch
        Dim z_pitch As Double
        Dim Rsegment As Integer
        'find relationship between z_pitch and r_pitch
        If (lastR < gNeedleRadius) Then
            'z_pitch = r_pitch * height / radius
            'Rsegment = CInt(r / r_pitch + 0.4999)
            Rsegment = CInt((r - gNeedleRadius) / r_pitch)
            If Rsegment < 1 Then
                Rsegment = 1
            End If
            z_pitch = height / Rsegment
        Else
            Rsegment = noofrun
            z_pitch = height / Rsegment

        End If


        Dim prev_x As Double = tmpt1(0)
        Dim prev_y As Double = tmpt1(1)
        Dim prev_z As Double = tmpt1(2)

        Dim i, j As Integer
        Dim p(3), lastPt(3) As Double
        Dim point3d As Point3d
        For i = 1 To Rsegment  'number of runs
            end_z = z + z_pitch
            end_r = r - r_pitch
            If (end_r < gNeedleRadius) Then end_r = gNeedleRadius
            If (height >= 0.0) And (end_z > height + init_start_z) Then
                end_z = height + init_start_z
            ElseIf (height < 0.0) And (end_z < height + init_start_z) Then
                end_z = height + init_start_z
            End If
            Dim ratio As Double = Math.Abs(end_r - r) / r_pitch  'for non full circle
            If (gPrecision / r > 1.0) Then Return -1
            Dim curr_est_angle_step As Double = 2.0 * Math.Acos(1.0 - gPrecision / r)
            Dim angle_ratio As Double = curr_est_angle_step / est_angle_step
            'test
            angle_ratio = 1.0
            Dim Angle_Step As Double = 2 * Math.PI * ratio * angle_ratio / step_no
            Dim curAngle As Double = start_angle
            For j = 1 To step_no  'one run 
                Dim t As Double = CDbl(j) / step_no
                Dim curR As Double = (1 - t) * r + end_r * t
                curAngle = curAngle + Angle_Step
                x = curR * Math.Cos(curAngle)
                y = curR * Math.Sin(curAngle)
                z = (1 - t) * start_z + end_z * t
                Dim prev_p() As Double = {prev_x, prev_y, prev_z}
                p(0) = x
                p(1) = y
                p(2) = z
                If (i = Rsegment) And (j = step_no) Then
                    lastPt(0) = x
                    lastPt(1) = y
                    lastPt(2) = z
                End If
                'transform to work coordinate
                Rotate3D(prev_p, -angle(1), 1, prev_p)
                Rotate3D(prev_p, -angle(0), 2, prev_p)
                Rotate3D(p, -angle(1), 1, p)
                Rotate3D(p, -angle(0), 2, p)
                Translate3D(prev_p, res_centre, prev_p)
                Translate3D(p, res_centre, p)
                point3d.DispenseOn = True
                point3d.x = prev_p(0)
                point3d.y = prev_p(1)
                point3d.z = prev_p(2)
                Dim dx As Double = x - prev_x 'problem here
                Dim dy As Double = y - prev_y
                Dim dz As Double = z - prev_z
                Dim distance As Double = Math.Sqrt(dx * dx + dy * dy + dz * dz)
                If distance > 1.0 * gPrecision Then
                    pointlist.Add(point3d) 'add to point list
                    prev_x = x
                    prev_y = y
                    prev_z = z
                End If
            Next j
            r = r - r_pitch
            start_z = z
        Next i
        point3d.DispenseOn = True
        point3d.x = p(0)
        point3d.y = p(1)
        point3d.z = p(2)
        pointlist.Add(point3d)
        Dim count As Integer = pointlist.Count
        If (InOut.ToUpper = "OUT") Then  'Detail operation
            pointlist.Reverse(0, count)

            If (ddist > 0.001) Then

                Dim detail_angle As Double = -ddist / radius
                If detail_angle <= -2.0 * Math.PI Then
                    detail_angle = -2.0 * Math.PI + Math.PI / 360.0 '0.5 degree tolerance
                End If
                Dim newPt(3) As Double
                Rotate2D(tmpt1, detail_angle, newPt)
                newPt(2) = tmpt1(2)
                detail_angle = Math.Atan2(newPt(1), newPt(0))
                If detail_angle < 0.0 Then detail_angle = detail_angle + 2.0 * Math.PI
                Dim end_angle As Double = Math.Atan2(tmpt1(1), tmpt1(0))
                If end_angle < 0.0 Then end_angle = end_angle + 2.0 * Math.PI
                Dim angle_span As Double
                If detail_angle <= end_angle Then
                    angle_span = end_angle - detail_angle
                Else
                    angle_span = end_angle + 2.0 * Math.PI - detail_angle
                End If
                step_no = CInt(angle_span / est_angle_step) + 1
                actual_angle_step = angle_span / step_no

                z = tmpt1(2)
                r = radius

                prev_x = tmpt1(0)
                prev_y = tmpt1(1)
                prev_z = tmpt1(2)

                Dim curAngle As Double = end_angle
                For j = 1 To step_no
                    curAngle = curAngle - actual_angle_step
                    x = r * Math.Cos(curAngle)
                    y = r * Math.Sin(curAngle)

                    Dim prev_p() As Double = {prev_x, prev_y, prev_z}
                    p(0) = x
                    p(1) = y
                    p(2) = z
                    Rotate3D(prev_p, -angle(1), 1, prev_p)
                    Rotate3D(prev_p, -angle(0), 2, prev_p)
                    Rotate3D(p, -angle(1), 1, p)
                    Rotate3D(p, -angle(0), 2, p)
                    Translate3D(prev_p, res_centre, prev_p)
                    Translate3D(p, res_centre, p)

                    point3d.DispenseOn = False
                    point3d.x = prev_p(0)
                    point3d.y = prev_p(1)
                    point3d.z = prev_p(2)
                    If (j > 1) Then pointlist.Add(point3d)
                    prev_x = x
                    prev_y = y
                    prev_z = z
                Next j
                point3d.DispenseOn = False
                point3d.x = p(0)
                point3d.y = p(1)
                point3d.z = p(2)
                pointlist.Add(point3d)
            ElseIf ddist < -0.001 Then
                Dim sumofdist, segdist As Double
                sumofdist = 0.0
                point3d = pointlist(0)
                prev_x = point3d.x
                prev_y = point3d.y
                prev_z = point3d.z
                For i = 1 To count - 1
                    point3d = pointlist(i)
                    point3d.DispenseOn = False
                    pointlist.RemoveAt(i)
                    pointlist.Insert(i, point3d)
                    x = point3d.x
                    y = point3d.y
                    z = point3d.z
                    segdist = Math.Sqrt((x - prev_x) * (x - prev_x) + (y - prev_y) * (y - prev_y) + (z - prev_z) * (z - prev_z))
                    sumofdist = sumofdist + segdist
                    If sumofdist >= Math.Abs(ddist) Then
                        Exit For
                    End If
                    prev_x = x
                    prev_y = y
                    prev_z = z
                Next
                If i = count Then Return -1
                pointlist.Reverse(0, count)
            Else
                pointlist.Reverse(0, count)
                point3d = pointlist(count - 1)
                point3d.DispenseOn = False
                pointlist.RemoveAt(count - 1)
                pointlist.Add(point3d)
            End If
        Else  'IN' 'problem with ddist
            If (ddist > 0.001) Then
                Dim detail_angle As Double = ddist / end_r
                If detail_angle >= 2.0 * Math.PI Then
                    detail_angle = 2.0 * Math.PI - Math.PI / 360.0 '0.5 degree tolerance
                End If
                Dim newPt(3) As Double
                Rotate2D(lastPt, detail_angle, newPt)
                newPt(2) = lastPt(2)
                detail_angle = Math.Atan2(newPt(1), newPt(0))
                If detail_angle < 0.0 Then detail_angle = detail_angle + 2.0 * Math.PI
                Dim end_angle As Double = Math.Atan2(lastPt(1), lastPt(0))
                If end_angle < 0.0 Then end_angle = end_angle + 2.0 * Math.PI
                Dim angle_span As Double
                If detail_angle >= end_angle Then
                    angle_span = detail_angle - end_angle
                Else
                    angle_span = detail_angle - end_angle + 2.0 * Math.PI
                End If
                step_no = CInt(angle_span / est_angle_step) + 1
                actual_angle_step = angle_span / step_no

                z = lastPt(2)
                r = end_r

                prev_x = lastPt(0)
                prev_y = lastPt(1)
                prev_z = lastPt(2)

                Dim curAngle As Double = end_angle
                For j = 1 To step_no
                    curAngle = curAngle + actual_angle_step
                    x = r * Math.Cos(curAngle)
                    y = r * Math.Sin(curAngle)

                    Dim prev_p() As Double = {prev_x, prev_y, prev_z}
                    p(0) = x
                    p(1) = y
                    p(2) = z
                    Rotate3D(prev_p, -angle(1), 1, prev_p)
                    Rotate3D(prev_p, -angle(0), 2, prev_p)
                    Rotate3D(p, -angle(1), 1, p)
                    Rotate3D(p, -angle(0), 2, p)
                    Translate3D(prev_p, res_centre, prev_p)
                    Translate3D(p, res_centre, p)

                    point3d.DispenseOn = False
                    point3d.x = prev_p(0)
                    point3d.y = prev_p(1)
                    point3d.z = prev_p(2)
                    If (j > 1) Then pointlist.Add(point3d)
                    prev_x = x
                    prev_y = y
                    prev_z = z
                Next j
                point3d.DispenseOn = False
                point3d.x = p(0)
                point3d.y = p(1)
                point3d.z = p(2)
                pointlist.Add(point3d)
            ElseIf ddist < -0.001 Then
                Dim sumofdist, segdist As Double
                sumofdist = 0.0
                point3d = pointlist(count - 1)
                prev_x = point3d.x
                prev_y = point3d.y
                prev_z = point3d.z
                For i = count - 2 To 0 Step -1
                    point3d = pointlist(i)
                    point3d.DispenseOn = False
                    pointlist.RemoveAt(i)
                    pointlist.Insert(i, point3d)
                    x = point3d.x
                    y = point3d.y
                    z = point3d.z
                    segdist = Math.Sqrt((x - prev_x) * (x - prev_x) + (y - prev_y) * (y - prev_y) + (z - prev_z) * (z - prev_z))
                    sumofdist = sumofdist + segdist
                    If sumofdist >= Math.Abs(ddist) Then
                        Exit For
                    End If
                    prev_x = x
                    prev_y = y
                    prev_z = z
                Next i
                If i = -1 Then Return -1
            Else
                point3d = pointlist(count - 1)
                point3d.DispenseOn = False
                pointlist.RemoveAt(count - 1)
                pointlist.Add(point3d)
            End If
        End If

        Return 0
    End Function


    'Calculate total length of a point list                           '
    '   plist:   a point list                                         '
    '   length:  total length of the point list                       '
    '                                                                 '

    Public Function CalculateTotalLength(ByVal plist As ArrayList, ByRef length As Double) As Integer
        Dim prevPoint, currPoint As Point3d
        Dim pointNum As Integer = plist.Count
        Dim i As Integer
        Dim oneLength As Double
        length = 0.0
        If pointNum < 1 Then Return -1
        prevPoint = plist(0)
        For i = 1 To pointNum - 1
            currPoint = plist(i)
            oneLength = Math.Sqrt((currPoint.x - prevPoint.x) * (currPoint.x - prevPoint.x) + (currPoint.y - prevPoint.y) * (currPoint.y - prevPoint.y))
            length = length + oneLength
            prevPoint = currPoint
        Next i
        Return 0
    End Function



    'Set fill rectangle point list's z value and transform it to work '
    '   coordinate system                                             '
    '   ttlength: point list's total length                           '
    '   rangle:   base rectangle's plane normal angle                 '
    '   height:   fill rectangle's height                             '
    '   pointlist: point list                                         '
    '                                                                 '

    Public Function ResetZnRotatePointList(ByVal ttlength As Double, ByVal rangle() As Double, ByVal height As Double, ByRef plist As ArrayList) As Integer
        Dim prevPoint, currPoint, tempPoint As Point3d
        Dim pointNum As Integer = plist.Count
        Dim i As Integer
        Dim t, z, oneLength, currLength, startZ, endZ, p(3) As Double
        If pointNum < 1 Then Return -1
        currLength = 0.0
        tempPoint.DispenseOn = True
        prevPoint = plist(0)
        startZ = prevPoint.z
        endZ = startZ + height

        p(0) = prevPoint.x
        p(1) = prevPoint.y
        p(2) = prevPoint.z
        'transform to work coordinate
        Rotate3D(p, -rangle(1), 1, p)
        Rotate3D(p, -rangle(0), 2, p)
        tempPoint.x = p(0)
        tempPoint.y = p(1)
        tempPoint.z = p(2)
        plist.Insert(0, tempPoint)
        plist.RemoveAt(1)

        For i = 1 To pointNum - 1
            currPoint = plist(i)
            oneLength = Math.Sqrt((currPoint.x - prevPoint.x) * (currPoint.x - prevPoint.x) + (currPoint.y - prevPoint.y) * (currPoint.y - prevPoint.y))
            currLength = currLength + oneLength
            t = currLength / ttlength
            z = (1 - t) * startZ + endZ * t  ' calculate z 
            p(0) = currPoint.x
            p(1) = currPoint.y
            p(2) = z
            'transform to work coordinate
            Rotate3D(p, -rangle(1), 1, p)
            Rotate3D(p, -rangle(0), 2, p)
            tempPoint.x = p(0)
            tempPoint.y = p(1)
            tempPoint.z = p(2)
            plist.Insert(i, tempPoint)
            plist.RemoveAt(i + 1)
            prevPoint = currPoint
        Next i
        Return 0
    End Function


    'Calculate rectangle spiral point list                            '
    '   pt1:      base circle point1                                  '
    '   pt2:      base circle point2                                  '
    '   pt3:      base circle point3                                  '
    '   r_pitch:  radius pitch                                        '
    '   noofrun:  number of runs (>=1)                                '
    '   rectheight:   rectangle spiral height                         '
    '   ddist:    detail distance                                     '
    '   InOut:    dispensing direction, "IN" or "OUT"                 '
    '   pointlist:   calculated point list                            '
    '                                                                 '

    Public Function GetSpiralRectPointList(ByVal pt1() As Double, ByVal pt2() As Double, ByVal pt3() As Double, ByVal r_pitch As Double, ByVal noofrun As Integer, ByVal rectheight As Double, ByVal ddist As Double, ByVal InOut As String, ByRef pointlist As ArrayList) As Integer

        Dim normal(3), angle(2), tmpt1(3), tmpt2(3), tmpt3(3), tmpt4(3), tmpt5(3) As Double
        Dim swapt(3), pt4(3), height As Double

        If noofrun < 1 Then Return -1
        If Get3dNormal(pt1, pt2, pt3, normal) < 0 Then Return -1 'get base plane's normal
        If Get3dNormalAngle(normal, angle) < 0 Then Return -1 'get normal's angle
        If (angle(1) > Math.PI / 2.0) Or (angle(1) < -Math.PI / 2.0) Then 'CCW
            height = -rectheight
        Else 'CW
            height = rectheight
        End If
        Dim Ptlist As New ArrayList
        GetDiaVertexlist(pt1, pt2, pt3, Ptlist) 'get base rectangle's point list
        'align base plane to xy plane
        Ptlist(3).CopyTo(pt4, 0)
        Rotate3D(pt1, angle(0), 2, swapt)
        Rotate3D(swapt, angle(1), 1, tmpt1)
        Rotate3D(pt2, angle(0), 2, swapt)
        Rotate3D(swapt, angle(1), 1, tmpt2)
        Rotate3D(pt3, angle(0), 2, swapt)
        Rotate3D(swapt, angle(1), 1, tmpt3)
        Rotate3D(pt4, angle(0), 2, swapt)
        Rotate3D(swapt, angle(1), 1, tmpt4)
        tmpt1.CopyTo(tmpt5, 0)

        Dim i, j, stNo, dir As Integer
        Dim p(3), lastPt(3) As Double
        Dim point3d, tempPoint As Point3d
        Dim parentList As New ArrayList
        Dim childList As New ArrayList
        tempPoint.DispenseOn = True
        tempPoint.x = tmpt1(0)
        tempPoint.y = tmpt1(1)
        tempPoint.z = tmpt1(2)
        parentList.Add(tempPoint)
        tempPoint.x = tmpt2(0)
        tempPoint.y = tmpt2(1)
        tempPoint.z = tmpt2(2)
        parentList.Add(tempPoint)
        tempPoint.x = tmpt3(0)
        tempPoint.y = tmpt3(1)
        tempPoint.z = tmpt3(2)
        parentList.Add(tempPoint)
        tempPoint.x = tmpt4(0)
        tempPoint.y = tmpt4(1)
        tempPoint.z = tmpt4(2)
        parentList.Add(tempPoint)
        tempPoint.x = tmpt5(0)
        tempPoint.y = tmpt5(1)
        tempPoint.z = tmpt5(2)
        parentList.Add(tempPoint)
        Dim v1(3), v2(3) As Double
        v1(0) = parentList(1).x - parentList(0).x 'Pt2(0) - Pt1(0)
        v1(1) = parentList(1).y - parentList(0).y 'Pt2(1) - Pt1(1)
        v2(0) = parentList(2).x - parentList(0).x 'Pt3(0) - Pt1(0)
        v2(1) = parentList(2).y - parentList(0).y 'Pt3(1) - Pt1(1)
        'set point list direction
        Dim vCross As Double = Cross2(v1, v2)
        If vCross < -0.0001 Then
            dir = 1
        ElseIf vCross > 0.0001 Then
            dir = -1
        Else
            dir = 0
            Return -1
        End If
        Dim rtn As Integer = 0
        Dim offS1(3), offS2(3), offS3(3), offS4(3), interS(3) As Double
        For i = 1 To noofrun
            rtn = Get4PtDiaOffsetVertexlist(parentList, r_pitch, dir, childList) 'offset rectangle
            If dir = 0 Then 'offset finished
                If i = 1 Then
                    stNo = 0
                Else
                    stNo = 1
                End If
                For j = stNo To parentList.Count - 1
                    pointlist.Add(parentList(j))
                Next j
                Exit For
            ElseIf rtn < 0 Then
                Return rtn
            End If
            Dim cv1(3) As Double
            cv1(0) = childList(1).x - childList(0).x
            cv1(1) = childList(1).y - childList(0).y
            Dim vangle As Double = 0.0
            If (Dot2(v1, cv1, vangle) <> 0) Then
                dir = 0
            ElseIf vangle > 0.5 Then
                dir = 0
            End If

            tempPoint = childList(0)
            offS1(0) = tempPoint.x
            offS1(1) = tempPoint.y
            offS1(2) = tempPoint.z
            tempPoint = childList(1)
            offS2(0) = tempPoint.x
            offS2(1) = tempPoint.y
            offS2(2) = tempPoint.z
            tempPoint = parentList(3)
            offS3(0) = tempPoint.x
            offS3(1) = tempPoint.y
            offS3(2) = tempPoint.z
            tempPoint = parentList(4)
            offS4(0) = tempPoint.x
            offS4(1) = tempPoint.y
            offS4(2) = tempPoint.z
            If LineIntersection(offS1, offS2, offS3, offS4, interS) < 0 Then Return -1
            tempPoint.x = interS(0)
            tempPoint.y = interS(1)
            tempPoint.z = interS(2)
            If dir <> 0 Then
                parentList.RemoveAt(4)
                parentList.Add(tempPoint)
            End If

            If i = 1 Then
                stNo = 0
            Else
                stNo = 1
            End If
            For j = stNo To parentList.Count - 1
                pointlist.Add(parentList(j))
            Next j

            If dir = 0 Then Exit For

            parentList = childList.Clone
            v1(0) = parentList(1).x - parentList(0).x
            v1(1) = parentList(1).y - parentList(0).y
        Next i

        Dim ttLength As Double
        If CalculateTotalLength(pointlist, ttLength) < 0 Then Return -1 'calculate total length of rectangle spiral
        'set rectangle spiral point list's height and transform it to work coordinate
        If ResetZnRotatePointList(ttLength, angle, height, pointlist) Then Return -1

        Dim count As Integer = pointlist.Count
        Dim detailPt(3) As Double
        'Detail operation
        If (InOut.ToUpper = "OUT") Then
            pointlist.Reverse(0, count)
        End If

        If (ddist > 0.001) Then
            point3d = pointlist(count - 2)
            tmpt1(0) = point3d.x
            tmpt1(1) = point3d.y
            tmpt1(2) = point3d.z
            point3d = pointlist(count - 1)
            tmpt2(0) = point3d.x
            tmpt2(1) = point3d.y
            tmpt2(2) = point3d.z
            If PointOnLine3d(tmpt1, tmpt2, ddist, detailPt) < 0 Then Return -1
            point3d.DispenseOn = True
            point3d.x = pointlist(count - 1).x
            point3d.y = pointlist(count - 1).y
            point3d.z = pointlist(count - 1).z
            pointlist.RemoveAt(count - 1)
            pointlist.Add(point3d)
            point3d.DispenseOn = False
            point3d.x = detailPt(0)
            point3d.y = detailPt(1)
            point3d.z = detailPt(2)
            pointlist.Add(point3d)
        ElseIf ddist < -0.001 Then
            point3d = pointlist(count - 2)
            tmpt1(0) = point3d.x
            tmpt1(1) = point3d.y
            tmpt1(2) = point3d.z
            point3d = pointlist(count - 1)
            tmpt2(0) = point3d.x
            tmpt2(1) = point3d.y
            tmpt2(2) = point3d.z
            If PointOnLine3d(tmpt1, tmpt2, ddist, detailPt) < 0 Then Return -1
            point3d.DispenseOn = False
            point3d.x = pointlist(count - 1).x
            point3d.y = pointlist(count - 1).y
            point3d.z = pointlist(count - 1).z
            pointlist.RemoveAt(count - 1)
            pointlist.Add(point3d)

            point3d.DispenseOn = True
            point3d.x = detailPt(0)
            point3d.y = detailPt(1)
            point3d.z = detailPt(2)
            pointlist.Insert(count - 1, point3d)
        Else
            point3d = pointlist(count - 2)
            tmpt1(0) = point3d.x
            tmpt1(1) = point3d.y
            tmpt1(2) = point3d.z
            point3d = pointlist(count - 1)
            tmpt2(0) = point3d.x
            tmpt2(1) = point3d.y
            tmpt2(2) = point3d.z
            If PointOnLine3d(tmpt1, tmpt2, -0.001, detailPt) < 0 Then Return -1
            point3d.DispenseOn = False
            point3d.x = pointlist(count - 1).x
            point3d.y = pointlist(count - 1).y
            point3d.z = pointlist(count - 1).z
            pointlist.RemoveAt(count - 1)
            pointlist.Add(point3d)

            point3d.DispenseOn = True
            point3d.x = detailPt(0)
            point3d.y = detailPt(1)
            point3d.z = detailPt(2)
            pointlist.Insert(count - 1, point3d)
        End If

        Return 0
    End Function



    'Calculate 3d arc point list                                      '
    '   pt1:      arc point1                                          '
    '   pt2:      arc point2                                          '
    '   pt3:      arc point3                                          '
    '   ddist:    detail distance                                     '
    '   pointlist:   calculated point list                            '
    '                                                                 '

    Public Function GetArc3dPointList(ByVal pt1() As Double, ByVal pt2() As Double, ByVal pt3() As Double, ByVal ddist As Double, ByRef pointlist As ArrayList) As Integer

        Dim centre(3), normal(3), angle(2), tmpt1(3), tmpt2(3), tmpt3(3), radius As Double

        If Get3dCircleCentre(pt1, pt2, pt3, centre, radius) < 0 Then Return -1 'get arc centre and radius
        If Get3dNormal(pt1, pt2, pt3, normal) < 0 Then Return -1 'get arc plane's normal
        If Get3dNormalAngle(normal, angle) < 0 Then Return -1 'get normal's angle
        'align arc plane to xy plane
        Rotate3D(pt1, angle(0), 2, tmpt1)
        Rotate3D(tmpt1, angle(1), 1, tmpt1)
        Rotate3D(pt2, angle(0), 2, tmpt2)
        Rotate3D(tmpt2, angle(1), 1, tmpt2)
        Rotate3D(pt3, angle(0), 2, tmpt3)
        Rotate3D(tmpt3, angle(1), 1, tmpt3)

        'transform to arc centre's coornidate
        Dim res_centre(3) As Double
        centre.CopyTo(res_centre, 0)
        Rotate3D(centre, angle(0), 2, centre)
        Rotate3D(centre, angle(1), 1, centre)
        Dim dist() As Double = {-centre(0), -centre(1), -centre(2)}
        Translate3D(tmpt1, dist, tmpt1)
        Translate3D(tmpt2, dist, tmpt2)
        Translate3D(tmpt3, dist, tmpt3)

        Dim dangle_span As Double = ddist / radius  'detail's angle
        Dim newPt(3) As Double

        Rotate2D(tmpt3, dangle_span, newPt)
        newPt(2) = tmpt1(2)

        Dim start_angle As Double = Math.Atan2(tmpt1(1), tmpt1(0))
        If start_angle < 0.0 Then start_angle = start_angle + 2.0 * Math.PI
        Dim detail_angle As Double = Math.Atan2(newPt(1), newPt(0))
        If detail_angle < 0.0 Then detail_angle = detail_angle + 2.0 * Math.PI
        Dim end_angle As Double = Math.Atan2(tmpt3(1), tmpt3(0))
        If end_angle < 0.0 Then end_angle = end_angle + 2.0 * Math.PI

        If Math.Abs(ddist) < 0.001 Then
            detail_angle = end_angle
        ElseIf ddist <= -0.001 Then

        Else
            Dim swap As Double = end_angle
            end_angle = detail_angle
            detail_angle = swap
        End If

        Dim angle_span As Double
        If end_angle <= start_angle Then
            angle_span = end_angle + 2.0 * Math.PI - start_angle
        Else
            angle_span = end_angle - start_angle
        End If

        If (ddist <= -0.001) Then
            If -dangle_span >= angle_span Then Return -1 'detail distance is too big
        End If

        If (gPrecision / radius > 1.0) Then Return -1 'radius is too small
        Dim est_angle_step As Double = 2.0 * Math.Acos(1.0 - gPrecision / radius)
        '''''''''''''''''''''''
        Dim step_no As Integer = CInt(angle_span / est_angle_step) + 1
        If (step_no > 500) Then step_no = 500
        Dim actual_angle_step As Double = angle_span / step_no
        Dim x, y As Double
        Dim z As Double = tmpt1(2)
        Dim r As Double = radius

        Dim prev_x As Double = tmpt1(0)
        Dim prev_y As Double = tmpt1(1)
        Dim prev_z As Double = tmpt1(2)

        Dim i, j As Integer
        Dim p(3) As Double
        Dim point3d As Point3d

        Dim curAngle As Double = start_angle
        For j = 1 To step_no

            curAngle = curAngle + actual_angle_step
            x = r * Math.Cos(curAngle)
            y = r * Math.Sin(curAngle)

            Dim prev_p() As Double = {prev_x, prev_y, prev_z}
            p(0) = x
            p(1) = y
            p(2) = z
            'transform to work coordinate
            Rotate3D(prev_p, -angle(1), 1, prev_p)
            Rotate3D(prev_p, -angle(0), 2, prev_p)
            Rotate3D(p, -angle(1), 1, p)
            Rotate3D(p, -angle(0), 2, p)
            Translate3D(prev_p, res_centre, prev_p)
            Translate3D(p, res_centre, p)
            If curAngle >= detail_angle + actual_angle_step Then
                point3d.DispenseOn = False
            Else
                point3d.DispenseOn = True
            End If
            point3d.x = prev_p(0)
            point3d.y = prev_p(1)
            point3d.z = prev_p(2)
            pointlist.Add(point3d) 'add to point list
            prev_x = x
            prev_y = y
            prev_z = z
        Next j

        point3d.DispenseOn = False
        point3d.x = p(0)
        point3d.y = p(1)
        point3d.z = p(2)
        pointlist.Add(point3d)
        Return 0
    End Function



    'Calculate 3d circle point list                                   '
    '   pt1:      base circle point1                                  '
    '   pt2:      base circle point2                                  '
    '   pt3:      base circle point3                                  '
    '   ddist:    detail distance                                     '
    '   pointlist:   calculated point list                            '
    '                                                                 '

    Public Function GetCircle3dPointList(ByVal pt1() As Double, ByVal pt2() As Double, ByVal pt3() As Double, ByVal ddist As Double, ByRef pointlist As ArrayList) As Integer

        Dim centre(3), normal(3), angle(2), tmpt1(3), tmpt2(3), tmpt3(3), radius As Double

        If Get3dCircleCentre(pt1, pt2, pt3, centre, radius) < 0 Then Return -1 'get circle's centre and radius
        If Get3dNormal(pt1, pt2, pt3, normal) < 0 Then Return -1 'get circle plane's normal
        If Get3dNormalAngle(normal, angle) < 0 Then Return -1 'get normal's angle
        'align circle's plane to xy plane
        Rotate3D(pt1, angle(0), 2, tmpt1)
        Rotate3D(tmpt1, angle(1), 1, tmpt1)
        Rotate3D(pt2, angle(0), 2, tmpt2)
        Rotate3D(tmpt2, angle(1), 1, tmpt2)
        Rotate3D(pt3, angle(0), 2, tmpt3)
        Rotate3D(tmpt3, angle(1), 1, tmpt3)

        'transform to circle centre's coordinate
        Dim res_centre(3) As Double
        centre.CopyTo(res_centre, 0)
        Rotate3D(centre, angle(0), 2, centre)
        Rotate3D(centre, angle(1), 1, centre)
        Dim dist() As Double = {-centre(0), -centre(1), -centre(2)}
        Translate3D(tmpt1, dist, tmpt1)
        Translate3D(tmpt2, dist, tmpt2)
        Translate3D(tmpt3, dist, tmpt3)

        Dim dangle_span As Double = ddist / radius 'detail's angle span
        Dim newPt(3) As Double

        Rotate2D(tmpt1, dangle_span, newPt)
        newPt(2) = tmpt1(2)

        Dim start_angle As Double = Math.Atan2(tmpt1(1), tmpt1(0))
        If start_angle < 0.0 Then start_angle = start_angle + 2.0 * Math.PI
        Dim detail_angle As Double = Math.Atan2(newPt(1), newPt(0))
        If detail_angle < 0.0 Then detail_angle = detail_angle + 2.0 * Math.PI
        Dim end_angle As Double = start_angle

        Dim angle_span As Double
        If Math.Abs(ddist) < 0.001 Then
            detail_angle = start_angle + 2.0 * Math.PI
            angle_span = 2.0 * Math.PI
        ElseIf ddist <= -0.001 Then
            angle_span = 2.0 * Math.PI
            If detail_angle < end_angle Then
                detail_angle = detail_angle + 2.0 * Math.PI
            End If
        Else
            Dim swap As Double = end_angle
            end_angle = detail_angle
            detail_angle = swap
            If end_angle < start_angle Then
                angle_span = end_angle + 2.0 * Math.PI - start_angle + 2.0 * Math.PI
            Else
                angle_span = end_angle - start_angle + 2.0 * Math.PI
            End If
            detail_angle = detail_angle + 2.0 * Math.PI
        End If

        If (ddist <= -0.001) Then
            If -dangle_span >= 2.0 * Math.PI Then Return -1 'detail distance is too long
        End If

        If (gPrecision / radius > 1.0) Then Return -1
        Dim est_angle_step As Double = 2.0 * Math.Acos(1.0 - gPrecision / radius)
        '''''''''''''''''''''''
        Dim step_no As Integer = CInt(angle_span / est_angle_step) + 1
        If (step_no > 500) Then step_no = 500
        Dim actual_angle_step As Double = angle_span / step_no
        Dim x, y As Double
        Dim z As Double = tmpt1(2)
        Dim r As Double = radius

        Dim prev_x As Double = tmpt1(0)
        Dim prev_y As Double = tmpt1(1)
        Dim prev_z As Double = tmpt1(2)

        Dim i, j As Integer
        Dim p(3) As Double
        Dim point3d As Point3d

        Dim curAngle As Double = start_angle
        For j = 1 To step_no
            curAngle = curAngle + actual_angle_step
            x = r * Math.Cos(curAngle)
            y = r * Math.Sin(curAngle)

            Dim prev_p() As Double = {prev_x, prev_y, prev_z}
            p(0) = x
            p(1) = y
            p(2) = z
            'transform to work coordinate
            Rotate3D(prev_p, -angle(1), 1, prev_p)
            Rotate3D(prev_p, -angle(0), 2, prev_p)
            Rotate3D(p, -angle(1), 1, p)
            Rotate3D(p, -angle(0), 2, p)
            Translate3D(prev_p, res_centre, prev_p)
            Translate3D(p, res_centre, p)
            If curAngle >= detail_angle + actual_angle_step Then
                point3d.DispenseOn = False
            Else
                point3d.DispenseOn = True
            End If
            point3d.x = prev_p(0)
            point3d.y = prev_p(1)
            point3d.z = prev_p(2)
            pointlist.Add(point3d) 'add to point list
            prev_x = x
            prev_y = y
            prev_z = z
        Next j

        point3d.DispenseOn = False
        point3d.x = p(0)
        point3d.y = p(1)
        point3d.z = p(2)
        pointlist.Add(point3d)
        Return 0
    End Function

End Module

Public Module SaveProgram

    Private savProgram As Boolean

    Public Property UnSave() As Boolean
        Get
            Return savProgram
        End Get
        Set(ByVal Value As Boolean)
            savProgram = Value
        End Set
    End Property
End Module