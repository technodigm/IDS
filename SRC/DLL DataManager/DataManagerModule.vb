Option Strict Off
Option Explicit On 
Imports System
Imports System.Data.OleDb
Imports Microsoft.VisualBasic
Imports System.Data.SqlClient
Imports System.IO
Imports System.Runtime.Serialization.Formatters.Soap
Imports System.Runtime.Serialization.Formatters.Binary
Imports System.Runtime.Serialization

<Serializable()> Public Class CIDSData

    Public SoftwareVersion As Double = 1.7              'current software version. - 26 sep 2007
    Public SystemAtLogin As Boolean


#Region "Base declaration and Function"

    Public Admin As New CIDSAdmin                       'create instance for administration objects
    Public Hardware As New CIDSHardWare                 'create instance for Hardware devices objects
    Public Execution As New CIDSPatternExecution        'create instance for execution related objects
    Public Shared ParameterID As New CIDSParameterID    'create instance for record/group 
    Public Shared MsgErr As String                      'declare Error Message

    Friend IsSaveAsFile As Boolean                      'flag
    Friend IsOpenFile As Boolean                        'flag
    Friend IsNewFile As Boolean                         'flag

    'to save file using soap method - not used
    Friend Sub SerializeData()

        Dim Savefile As FileStream
        Savefile = File.Create(CIDSData.ParameterID.RecordID + ".Soap")

        Dim Soapformatter As New SoapFormatter
        Soapformatter.Serialize(Savefile, IDSData)

        Savefile.Close()

    End Sub
    'to open file using soap method - not used
    Friend Sub DeserializeData()

        Dim ReadFile As FileStream
        ReadFile = File.OpenRead(CIDSData.ParameterID.RecordID + ".Soap")

        Try
            Dim Soapformatter As New SoapFormatter
            IDSData = CType(Soapformatter.Deserialize(ReadFile), CIDSData)
        Catch ex As Exception
            MessageBox.Show(ex.ToString, "Error!", MessageBoxButtons.OK)
        End Try

        ReadFile.Close()

    End Sub

    'to save .pat file that is opened as factorydefault stated in RecordID
    Public Function SaveData() As Boolean

        MsgErr = ""
        'Return SaveFile(ParameterID.RecordID)
        Return SaveFile(IDSData.ParameterID.RecordID)
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ''' end here
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    End Function
    Public Function SaveToDefaultFile() As Boolean
        Return SaveToDefault()
    End Function
    Public Function OpenDefaultFile() As Boolean
        Return OpenToDefault()
    End Function

    'to open .pat file that is opened as factorydefault stated in RecordID
    Public Function OpenData() As Boolean

        MsgErr = ""

        'Return openFile(ParameterID.RecordID)
        Return Openfile(IDSData.ParameterID.RecordID)

        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ''' end here
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    End Function

    'to save pathfile .pat file 
    Public Function SavePathFileData(ByRef MyFileName As String) As Boolean

        MsgErr = ""
        StrConvert(MyFileName)
        Return SavePathFileName(MyFileName)

        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ''' end here
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    End Function

    'extract filename and extension from myfilename
    Private Sub StrConvert(ByRef myfilename As String)

        Dim MyStr() As String
        MyStr = myfilename.Split("\")                                       'split to substrings if the char is "\" of the myfilename
        Dim Len As Integer = MyStr.Length - 1                               'total number of substring after split
        Dim Start As Integer = myfilename.Length - MyStr(Len).Length - 1    'the starting char of the filename
        Dim Count As Integer = MyStr(Len).Length + 1                        'get the number of char in the filename 
        IDSData.Admin.Folder.PatternPath = myfilename.Remove(Start, Count)  'get the path only of the pathfilename
        Dim Mystr1() As String = MyStr(Len).Split(".")                      'split the filename into substring of filename and extension
        IDSData.ParameterID.RecordID = Mystr1(0)                            'substring of filename
        IDSData.Admin.Folder.FileExtension = Mystr1(1)                      'substring of extension

    End Sub

    'to open pathfile .pat file 
    Public Function OpenPathFileData(ByRef MyFileName As String) As Boolean

        MsgErr = ""
        StrConvert(MyFileName)
        Return OpenPathFileName(MyFileName)

        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ''' end here
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    End Function
#End Region

#Region "Administration"

    '''''''''''''''''''''''''''''''''''''''''''''''
    '       System Administration Related Variables
    '''''''''''''''''''''''''''''''''''''''''''''''
    <Serializable()> Public Class CIDSAdmin
        Public Folder As CIDSFolder = New CIDSFolder                    'create instance of a folder object
        <Serializable()> Public Class CIDSFolder

            Public PatternPath As String = "C:\IDS\Pattern_Dir"
            Public DataPath As String = "C:\IDS\BIN"
            Public FileExtension As String = "Pat"
        End Class

        Public User As CIDSUser = New CIDSUser
        <Serializable()> Public Class CIDSUser                          'create instance of the login user
            Public ID As String
            Public Name As String
            Public PassWord As String
            Public Department As String
            Public ContactNo As String
            Public RunApplication As String
            Public Group As CIDSGroup = New CIDSGroup                   'create instance of the user login group 
        End Class

        <Serializable()> Public Class CIDSGroup
            Public ID As String
            Public Remark As String
            Public PrivilegeArray As New ArrayList                      'create instance of privilege given to login group
            Public SystemHardwareArray As New ArrayList                 'create instance of the hardware devices selected for the login group
        End Class

        Public ALLPrivileges As CIDSPrivileges = New CIDSPrivileges     'create instance of all privileges declared for IDS  
        <Serializable()> Public Class CIDSPrivileges
            Public IDArray As New ArrayList
            Public TypeArray As New ArrayList
        End Class
    End Class
#End Region

#Region "Hardware"

    '''''''''''''''''''''''''''''''''''''''''''''''
    '       System Related Hardware Variables
    '''''''''''''''''''''''''''''''''''''''''''''''

    <Serializable()> Public Class CIDSHardWare

#Region " "
        '
        'create instance of hardware devices
        '
        Public Gantry As New CIDSGantry
        Public Conveyor As New CIDSConveyor
        Public Thermal As New CIDSThermal
        'Public Lifter As New CIDSLifter
        Public Dispenser As New CIDSDispenser
        Public Regulator As New CIDSRegulator
        Public Camera As New CIDSCamera
        Public LightBox As New CIDSLightBox
        Public HeightSensor As New CIDSHeightSensor
        Public Needle As New CIDSNeedle
        Public Volume As New CIDSVolume
        Public SPC As New CIDSSPC
        Public SystemIO As New CIDSSystemIO
        'List all hardware iD of the IDS 
        Public Function ListAllHardWareID(ByRef RecordIDArray As ComboBox) As Boolean

            MsgErr = ""
            DBView = New DataView(DS.Tables("TableHardware"))
            Try
                If DBView.Count > 0 Then
                    Dim I As Integer = 0
                    RecordIDArray.Items.Clear()
                    For I = 0 To DBView.Count - 1
                        Dim Row As DataRow = DBView(I).Row
                        RecordIDArray.Items.Add(Row("ID"))
                    Next
                End If

            Catch ex As Exception
                MsgErr = ex.ToString
                Return False
            End Try
            Return True

        End Function
        'delete devices ID from hardware table
        Public Function DeleteHardWareID() As Boolean

            MsgErr = ""
            DBView = New DataView(DS.Tables("TableHardware"))
            DBView.RowFilter = "ID = '" + ParameterID.RecordID + "'"
            Try
                Dim Row As DataRow
                If DBView.Count <> 0 Then

                    If ParameterID.RecordID = "FactoryDefault" = False Then
                        Row = DBView(0).Row
                        Row.Delete()
                    Else
                        MsgErr = "Cant Delete Factory Default RecordID"
                        Return False
                    End If

                End If

                UpdateData("SELECT * FROM TableHardware", "TableHardware")
            Catch ex As Exception
                MsgErr = ex.ToString
                Return False
            End Try
            Return True
        End Function
        'saved devices ID to hardware table
        Public Shared Function SaveHardWareID() As Boolean
            MsgErr = ""

            DBView = New DataView(DS.Tables("TableHardware"))
            DBView.RowFilter = "ID ='" + ParameterID.RecordID + "'"
            Try
                Dim Row As DataRow
                If DBView.Count = 0 Then
                    Row = DBView.Table.NewRow()
                    Row("ID") = ParameterID.RecordID
                    DBView.Table.Rows.Add(Row)
                    UpdateData("SELECT * FROM TableHardware", "TableHardware")
                End If

            Catch ex As Exception
                MsgErr = ex.ToString
                Return False
            End Try

            Return True
        End Function
        'retrieve devices ID from hardware table
        Public Shared Function GetHardWareID() As Boolean
            MsgErr = ""
            DBView = New DataView(DS.Tables("TableHardware"))
            DBView.RowFilter = "ID ='" + ParameterID.RecordID + "'"


            If DBView.Count = 1 Then
                Dim Row As DataRow = DBView(0).Row
                ParameterID.RecordID = CStr(Row("ID"))
                Return True
            Else
                MsgErr = "Cant find Record ID " + ParameterID.RecordID
                Return False
            End If

        End Function

#End Region

#Region "Hardware Classes"

#Region " Gantry "
        <Serializable()> Public Class CIDSGantry

            Public SystemOriginPos As New CIDSPosXYZ        'DB

            'XY positions are taught in station settings 
            Public ParkPosition As New CIDSPosXYZ                'DB and Pat
            Public PurgePosition As New CIDSPosXYZ               'DB and Pat
            Public CleanPosition As New CIDSPosXYZ               'DB and Pat
            Public ChangeSyringePosition As New CIDSPosXYZ       'DB and Pat
            Public WeighingScalePosition As New CIDSPosXYZ       'DB
            Public WeighingScaleBottomRight As New CIDSPosXYZ       'DB

            'this defines the x,y,z work area. defined in system set-up
            Public WorkArea As New CIDSWorkArea             'DB

            'speed definitions
            Public ElementXYSpeed As Double                   'DB   'Pat
            Public ElementZSpeed As Double
            Public ServiceXYSpeed As Double                   'DB   'Pat
            Public ServiceZSpeed As Double
            Public MaxAccelerationLimit As Double           'DB    
            Public MaxSpeedLimit As Double                  'DB    

            Public GraphicDisplayXYLT As New CIDSPosXY      'Pat
            Public GraphicDisplayXYRB As New CIDSPosXY      'Pat
            Public GraphicDisplayRatio As Double            'Pat
            Public SystemUnit As Boolean                    'DB and Pat

            Public Function SaveData() As Boolean
                MsgErr = ""
                Return SaveFile(ParameterID.RecordID)

                ''' end here


                If SaveHardWareID() = False Then
                    Return False
                End If

                DBView = New DataView(DS.Tables("TableGantry"))
                DBView.RowFilter = "ID = '" + ParameterID.RecordID + "'"

                Try
                    Dim Row As DataRow
                    If DBView.Count = 0 Then
                        Row = DBView.Table.NewRow()
                        Row("ID") = ParameterID.RecordID
                        DBView.Table.Rows.Add(Row)
                    Else
                        Row = DBView(0).Row
                    End If

                    Row("SystemOriginPosX") = SystemOriginPos.X
                    Row("SystemOriginPosY") = SystemOriginPos.Y
                    Row("SystemOriginPosZ") = SystemOriginPos.Z

                    Row("ParkPosX") = ParkPosition.X
                    Row("ParkPosY") = ParkPosition.Y
                    Row("ParkPosZ") = ParkPosition.Z

                    Row("PurgePosX") = PurgePosition.X
                    Row("PurgePosY") = PurgePosition.Y
                    Row("PurgePosZ") = PurgePosition.Z

                    Row("CleanPosX") = CleanPosition.X
                    Row("CleanPosY") = CleanPosition.Y
                    Row("CleanPosZ") = CleanPosition.Z

                    Row("ChangeSyrPosX") = WeighingScalePosition.X
                    Row("ChangeSyrPosY") = WeighingScalePosition.Y
                    Row("ChangeSyrPosZ") = WeighingScalePosition.Z

                    Row("WorkAreaPosX") = WorkArea.X
                    Row("WorkAreaPosY") = WorkArea.Y
                    Row("WorkAreaPosZMin") = WorkArea.Z.Min
                    Row("WorkAreaPosZMax") = WorkArea.Z.Max

                    Row("ElementZSpeed") = ElementZSpeed
                    Row("ElementXYSpeed") = ElementXYSpeed
                    Row("MaxSpeedLimit") = MaxSpeedLimit
                    Row("MaxAccelerationLimit") = MaxAccelerationLimit

                    Row("SystemUnit") = SystemUnit

                    UpdateData("SELECT * FROM TableGantry", "TableGantry")

                Catch ex As Exception
                    MsgErr = ex.ToString
                    Return False
                End Try

                Return True
            End Function
            Public Function GetData() As Boolean

                MsgErr = ""
                Return Openfile(ParameterID.RecordID)

                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                ''' end here
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''


                DBView = New DataView(DS.Tables("TableGantry"))
                DBView.RowFilter = "ID = '" + ParameterID.RecordID + "'"
                If DBView.Count = 1 Then
                    Dim Row As DataRow = DBView(0).Row


                    SystemOriginPos.X = CDouble(Row("SystemOriginPosX"))
                    SystemOriginPos.Y = CDouble(Row("SystemOriginPosY"))
                    SystemOriginPos.Z = CDouble(Row("SystemOriginPosZ"))

                    ParkPosition.X = CDouble(Row("ParkPosX"))
                    ParkPosition.Y = CDouble(Row("ParkPosY"))
                    ParkPosition.Z = CDouble(Row("ParkPosZ"))

                    PurgePosition.X = CDouble(Row("PurgePosX"))
                    PurgePosition.Y = CDouble(Row("PurgePosY"))
                    PurgePosition.Z = CDouble(Row("PurgePosZ"))

                    CleanPosition.X = CDouble(Row("CleanPosX"))
                    CleanPosition.Y = CDouble(Row("CleanPosY"))
                    CleanPosition.Z = CDouble(Row("CleanPosZ"))

                    WeighingScalePosition.X = CDouble(Row("ChangeSyrPosX"))
                    WeighingScalePosition.Y = CDouble(Row("ChangeSyrPosY"))
                    WeighingScalePosition.Z = CDouble(Row("ChangeSyrPosZ"))

                    WorkArea.X = CDouble(Row("WorkAreaPosX"))
                    WorkArea.Y = CDouble(Row("WorkAreaPosY"))
                    WorkArea.Z.Min = CDouble(Row("WorkAreaPosZMin"))
                    WorkArea.Z.Max = CDouble(Row("WorkAreaPosZMax"))

                    ElementZSpeed = CDouble(Row("ElementZSpeed"))
                    ElementXYSpeed = CDouble(Row("ElementXYSpeed"))
                    MaxSpeedLimit = CDouble(Row("MaxSpeedLimit"))
                    MaxAccelerationLimit = CDouble(Row("MaxAccelerationLimit"))

                    SystemUnit = CBoolean(Row("SystemUnit"))

                    Return True
                Else
                    MsgErr = "Cant find Record ID " + ParameterID.RecordID
                    Return False
                End If

            End Function
            <Serializable()> Public Class CIDSXYZLimit
                Public X As New CIDSMinMax
                Public Y As New CIDSMinMax
                Public Z As New CIDSMinMax
            End Class



            <Serializable()> Public Class CIDSWorkArea
                Inherits CIDSPosXY
                Public Z As New CIDSMinMax
            End Class

        End Class
#End Region

#Region " Camera "
        <Serializable()> Public Class CIDSCamera


            Public SoftHomePos As New CIDSPosXYZ                    'DB   
            Public CalibrationPos As New CIDSPosXY                  'DB
            Public FiducialPos As New CIDSPosXYZ                    'PAT
            Public EdgeDetect As New CIDSEdgeDetect(4)              'declare Instance of 4 item for the array 
            <Serializable()> Public Class CIDSEdgeDetect
                Public Sub New(ByVal _Length As Integer)            'initial the size of the array
                    Call Length(_Length)
                End Sub

                Public NeedleOD As Integer                          'Pat                      
                Public ResultFlag As Boolean                        'Pat
                Public Point() As CIDSPosXYZ                        'Pat    'declare array without instance
                Public Sub Length(ByVal Length As Integer)
                    ReDim Point(Length) ' define the size 
                    Dim I As Integer
                    For I = 0 To Length
                        Point(I) = New CIDSPosXYZ
                    Next
                End Sub
            End Class

            Public Compensation As New CIDSCompansate(3)            'Pat 'declare Instance of 4 item for the array 
            <Serializable()> Public Class CIDSCompansate
                'CIDSMatrix as ctype
                Public Sub New(ByVal _Length As Integer)            'initial the size of the array
                    Call Length(_Length)
                End Sub
                Public Matrix() As CIDSMatrix                       'declare array without instance
                Public Sub Length(ByVal Number As Integer)
                    ReDim Matrix(Number) ' define the size 
                    Dim I As Integer
                    For I = 0 To Number
                        Matrix(I) = New CIDSMatrix
                    Next
                End Sub
                Public Function Length() As Integer
                    Return Matrix.Length
                End Function
                <Serializable()> Public Class CIDSMatrix
                    Inherits CIDSPosXYZ ' inherits from xyp 
                    Public W As Double = 0
                End Class
            End Class

            Public ChipDispenseType As New CIDSChipDispenseType(2)  'Pat'declare Instance of 3 item for the array 
            <Serializable()> Public Class CIDSChipDispenseType
                Public Sub New(ByVal _Length As Integer)            'initial the size of the array
                    Call Length(_Length)
                End Sub
                Public Point() As CIDSPosXYZ                         'declare array without instance
                Public Sub Length(ByVal Length As Integer)
                    ReDim Point(Length)
                    Dim I As Integer
                    For I = 0 To Length
                        Point(I) = New CIDSPosXYZ
                    Next
                End Sub
                Public Function Length() As Integer
                    Return Point.Length
                End Function
            End Class


            Public FileFiducial1 As String                      'Pat
            Public FileFiducial2 As String                      'Pat
            Public FileRejectMark As String                     'Pat
            Public ReferencePos As New CIDSPosXYZ               'Pat
            Public OffsetLNeedle As New CIDSPosXY               'Pat/DB
            Public OffsetRNeedle As New CIDSPosXY               'Pat/DB
            Public OffsetLaser As New CIDSPosXY                 'DB
            Public OffsetTouchProbe As New CIDSPosXY            'DB
            Public CalibrationFlag As Boolean                   'DB
            Public ImageResolution As Integer                   'DB
            Public PixelsToMM As New CIDSPosXY                  'DB
            Public PixelRatio As Double                         'DB
            Public GraphicsDisplayResolution As New CIDSPosXY   'DB

            Public CalibrationNumColumns As Integer
            Public CalibrationNumRows As Integer
            Public CalibrationSpacing As Integer
            Public CalibrationBlockPos As New CIDSPosXYZ
            Public BlockSize As New CIDSPosXY
            Public BlockSizeTolerance As Double
            Public BlockPosTolerance As Double
            Public BlockRotationTolerance As Double
            Public ImageFindDir As Boolean = True
            Public ImageVertical As Boolean = True
            Public Threshold As Integer
            Public Brightness As Integer
            Public Contrast As Integer
            Public ROISize As Double
            Public ResultSize As New CIDSPosXY
            Public ResultDia As Double
            Public ResultRotation As Double
            Public DotQCEnable As Boolean = False




            Public Function SaveData() As Boolean
                MsgErr = ""
                Return SaveFile(ParameterID.RecordID)

                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                ''' end here
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''


                If SaveHardWareID() = False Then
                    Return False
                End If


                DBView = New DataView(DS.Tables("TableCamera"))
                DBView.RowFilter = "ID = '" + ParameterID.RecordID + "'"
                Try
                    Dim Row As DataRow
                    If DBView.Count = 0 Then
                        Row = DBView.Table.NewRow()
                        Row("ID") = ParameterID.RecordID
                        DBView.Table.Rows.Add(Row)
                    Else
                        Row = DBView(0).Row
                    End If
                    Row("SoftHomePosX") = SoftHomePos.X
                    Row("SoftHomePosY") = SoftHomePos.Y
                    Row("SoftHomePosZ") = SoftHomePos.Z

                    Row("CalibrationPosX") = CalibrationPos.X
                    Row("CalibrationPosY") = CalibrationPos.Y

                    'Row("FiducialPosX") = FiducialPos.X
                    'Row("FiducialPosY") = FiducialPos.Y
                    'Row("FiducialPosZ") = FiducialPos.Z
                    'Row("FileFiducial1") = FileFiducial1
                    'Row("FileFiducial2") = FileFiducial2
                    'Row("FileRejectMark") = FileRejectMark
                    'Row("EdgeDetectNeedleOD") = EdgeDetect.NeedleOD
                    'Row("EdgeDetectResult") = EdgeDetect.ResultFlag

                    '''''' Compensation
                    'If Compensation.Matrix.Length > 0 Then
                    'Row("CompensationMatrix0W") = Compensation.Matrix(0).W
                    'Row("CompensationMatrix0X") = Compensation.Matrix(0).X
                    'Row("CompensationMatrix0Y") = Compensation.Matrix(0).Y
                    'Row("CompensationMatrix0Z") = Compensation.Matrix(0).Z

                    'ElseIf Compensation.Matrix.Length > 1 Then
                    '    Row("CompensationMatrix1W") = Compensation.Matrix(1).W
                    '    Row("CompensationMatrix1X") = Compensation.Matrix(1).X
                    '    Row("CompensationMatrix1Y") = Compensation.Matrix(1).Y
                    '    Row("CompensationMatrix1Z") = Compensation.Matrix(1).Z

                    'ElseIf Compensation.Matrix.Length > 2 Then
                    '    Row("CompensationMatrix2W") = Compensation.Matrix(2).W
                    '    Row("CompensationMatrix2X") = Compensation.Matrix(2).X
                    '    Row("CompensationMatrix2Y") = Compensation.Matrix(2).Y
                    '    Row("CompensationMatrix2Z") = Compensation.Matrix(2).Z

                    'ElseIf Compensation.Matrix.Length > 3 Then
                    '    Row("CompensationMatrix3W") = Compensation.Matrix(3).W
                    '    Row("CompensationMatrix3X") = Compensation.Matrix(3).X
                    '    Row("CompensationMatrix3Y") = Compensation.Matrix(3).Y
                    '    Row("CompensationMatrix3Z") = Compensation.Matrix(3).Z
                    'End If

                    'ChipDispenseType
                    'If ChipDispenseType.Point.Length > 0 Then
                    'Row("ChipDispenseTypePoint0X") = ChipDispenseType.Point(0).X
                    'Row("ChipDispenseTypePoint0Y") = ChipDispenseType.Point(0).Y
                    'Row("ChipDispenseTypePoint0Z") = ChipDispenseType.Point(0).Z

                    'ElseIf ChipDispenseType.Point.Length > 1 Then
                    '    Row("ChipDispenseTypePoint1X") = ChipDispenseType.Point(1).X
                    '    Row("ChipDispenseTypePoint1Y") = ChipDispenseType.Point(1).Y
                    '    Row("ChipDispenseTypePoint1Z") = ChipDispenseType.Point(1).Z
                    'End If


                    '''''' Edge detect
                    'If EdgeDetect.Point.Length > 0 Then
                    'Row("EdgeDetectPoint0X") = EdgeDetect.Point(0).X
                    'Row("EdgeDetectPoint0Y") = EdgeDetect.Point(0).Y
                    'Row("EdgeDetectPoint0Z") = EdgeDetect.Point(0).Z

                    'ElseIf EdgeDetect.Point.Length > 1 Then
                    '   Row("EdgeDetectPoint1X") = EdgeDetect.Point(1).X
                    '   Row("EdgeDetectPoint1Y") = EdgeDetect.Point(1).Y
                    '   Row("EdgeDetectPoint1Z") = EdgeDetect.Point(1).Z

                    'ElseIf EdgeDetect.Point.Length > 2 Then
                    '    Row("EdgeDetectPoint2X") = EdgeDetect.Point(2).X
                    '    Row("EdgeDetectPoint2Y") = EdgeDetect.Point(2).Y
                    '    Row("EdgeDetectPoint2Z") = EdgeDetect.Point(2).Z

                    'ElseIf EdgeDetect.Point.Length > 3 Then
                    '    Row("EdgeDetectPoint3X") = EdgeDetect.Point(3).X
                    '    Row("EdgeDetectPoint3Y") = EdgeDetect.Point(3).Y
                    '    Row("EdgeDetectPoint3Z") = EdgeDetect.Point(3).Z
                    'End If

                    'Row("ReferencePosX") = ReferencePos.X
                    'Row("ReferencePosY") = ReferencePos.Y
                    'Row("ReferencePosZ") = ReferencePos.Z

                    Row("OffsetLNeedleX") = OffsetLNeedle.X
                    Row("OffsetLNeedleY") = OffsetLNeedle.Y

                    Row("OffsetRNeedleX") = OffsetRNeedle.X
                    Row("OffsetRNeedleY") = OffsetRNeedle.Y

                    Row("OffsetLaserX") = OffsetLaser.X
                    Row("OffsetLaserY") = OffsetLaser.Y

                    Row("OffsetTouchProbeX") = OffsetTouchProbe.X
                    Row("OffsetTouchProbeY") = OffsetTouchProbe.Y


                    Row("CalibrationFlag") = CalibrationFlag
                    Row("ImageResolution") = ImageResolution
                    'Row("PixelsToMM") = PixelsToMM.X 
                    Row("GraphicsDisplayResolutionX") = GraphicsDisplayResolution.X
                    Row("GraphicsDisplayResolutionY") = GraphicsDisplayResolution.Y


                    UpdateData("SELECT * FROM TableCamera", "TableCamera")
                Catch ex As Exception
                    MsgErr = ex.ToString
                    Return False
                End Try

                Return True

            End Function
            Public Function GetData() As Boolean

                MsgErr = ""
                Return Openfile(ParameterID.RecordID)

                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                ''' end here
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''

                DBView = New DataView(DS.Tables("TableCamera"))
                DBView.RowFilter = "ID = '" + ParameterID.RecordID + "'"
                If DBView.Count = 1 Then
                    Dim Row As DataRow = DBView(0).Row

                    SoftHomePos.X = CDouble(Row("SoftHomePosX"))
                    SoftHomePos.Y = CDouble(Row("SoftHomePosy"))
                    SoftHomePos.Z = CDouble(Row("SoftHomePosz"))

                    CalibrationPos.X = CDouble(Row("CalibrationPosX"))
                    CalibrationPos.Y = CDouble(Row("CalibrationPosY"))

                    'FiducialPos.X = CDouble(Row("FiducialPosX"))
                    'FiducialPos.Y = CDouble(Row("FiducialPosY"))
                    'FiducialPos.Z = CDouble(Row("FiducialPosZ"))

                    'FileFiducial1 = CString(Row("FileFiducial1"))
                    'FileFiducial2 = CString(Row("FileFiducial2"))
                    'FileRejectMark = CString(Row("FileRejectMark"))

                    'EdgeDetect.NeedleOD = CInteger(Row("EdgeDetectNeedleOD"))
                    'EdgeDetect.ResultFlag = CBoolean(Row("EdgeDetectResult"))

                    '''''' Compensation
                    'If Compensation.Matrix.Length > 0 Then
                    'Compensation.Matrix(0).W = CDouble(Row("CompensationMatrix0W"))
                    'Compensation.Matrix(0).X = CDouble(Row("CompensationMatrix0X"))
                    'Compensation.Matrix(0).Y = CDouble(Row("CompensationMatrix0Y"))
                    'Compensation.Matrix(0).Z = CDouble(Row("CompensationMatrix0Z"))

                    'ElseIf Compensation.Matrix.Length > 1 Then
                    '   Compensation.Matrix(1).W = CDouble(Row("CompensationMatrix1W"))
                    '  Compensation.Matrix(1).X = CDouble(Row("CompensationMatrix1X"))
                    ' Compensation.Matrix(1).Y = CDouble(Row("CompensationMatrix1Y"))
                    'Compensation.Matrix(1).Z = CDouble(Row("CompensationMatrix1Z"))

                    'ElseIf Compensation.Matrix.Length > 2 Then
                    '   Compensation.Matrix(2).W = CDouble(Row("CompensationMatrix2W"))
                    '  Compensation.Matrix(2).X = CDouble(Row("CompensationMatrix2X"))
                    ' Compensation.Matrix(2).Y = CDouble(Row("CompensationMatrix2Y"))
                    'Compensation.Matrix(2).Z = CDouble(Row("CompensationMatrix2Z"))

                    'ElseIf Compensation.Matrix.Length > 3 Then
                    '   Compensation.Matrix(3).W = CDouble(Row("CompensationMatrix3W"))
                    'Compensation.Matrix(3).X = CDouble(Row("CompensationMatrix3X"))
                    '  Compensation.Matrix(3).Y = CDouble(Row("CompensationMatrix3Y"))
                    ' Compensation.Matrix(3).Z = CDouble(Row("CompensationMatrix3Z"))
                    'End If

                    'ChipDispenseType
                    'If ChipDispenseType.Point.Length > 0 Then
                    '    ChipDispenseType.Point(0).X = CDouble(Row("ChipDispenseTypePoint0X"))
                    '    ChipDispenseType.Point(0).Y = CDouble(Row("ChipDispenseTypePoint0Y"))
                    '    ChipDispenseType.Point(0).Z = CDouble(Row("ChipDispenseTypePoint0Z"))'

                    'ElseIf ChipDispenseType.Point.Length > 1 Then
                    '    ChipDispenseType.Point(1).X = CDouble(Row("ChipDispenseTypePoint1X"))
                    '    ChipDispenseType.Point(1).Y = CDouble(Row("ChipDispenseTypePoint1Y"))
                    '    ChipDispenseType.Point(1).Z = CDouble(Row("ChipDispenseTypePoint1Z"))
                    'End If


                    '''''' Edge detect
                    'If EdgeDetect.Point.Length > 0 Then
                    '    EdgeDetect.Point(0).X = CDouble(Row("EdgeDetectPoint0X"))
                    '    EdgeDetect.Point(0).Y = CDouble(Row("EdgeDetectPoint0Y"))
                    '    EdgeDetect.Point(0).Z = CDouble(Row("EdgeDetectPoint0Z"))
                    'ElseIf EdgeDetect.Point.Length > 1 Then
                    '   EdgeDetect.Point(1).X = CDouble(Row("EdgeDetectPoint1X"))
                    '   EdgeDetect.Point(1).Y = CDouble(Row("EdgeDetectPoint1Y"))
                    '   EdgeDetect.Point(1).Z = CDouble(Row("EdgeDetectPoint1Z"))
                    'ElseIf EdgeDetect.Point.Length > 2 Then
                    '   EdgeDetect.Point(2).X = CDouble(Row("EdgeDetectPoint2X"))
                    '   EdgeDetect.Point(2).Y = CDouble(Row("EdgeDetectPoint2Y"))
                    'EdgeDetect.Point(2).Z = CDouble(Row("EdgeDetectPoint2Z"))
                    'ElseIf EdgeDetect.Point.Length > 3 Then
                    '    EdgeDetect.Point(3).X = CDouble(Row("EdgeDetectPoint3X"))
                    '    EdgeDetect.Point(3).Y = CDouble(Row("EdgeDetectPoint3Y"))
                    '    EdgeDetect.Point(3).Z = CDouble(Row("EdgeDetectPoint3Z"))
                    'End If

                    '    ReferencePos.X = CDouble(Row("ReferencePosX"))
                    '    ReferencePos.Y = CDouble(Row("ReferencePosY"))
                    '    ReferencePos.Z = CDouble(Row("ReferencePosZ"))


                    OffsetLNeedle.X = CDouble(Row("OffsetLNeedleX"))
                    OffsetLNeedle.Y = CDouble(Row("OffsetLNeedleY"))

                    OffsetRNeedle.X = CDouble(Row("OffsetRNeedleX"))
                    OffsetRNeedle.Y = CDouble(Row("OffsetRNeedleY"))

                    OffsetLaser.X = CDouble(Row("OffsetLaserX"))
                    OffsetLaser.Y = CDouble(Row("OffsetLaserY"))

                    OffsetTouchProbe.X = CDouble(Row("OffsetTouchProbeX"))
                    OffsetTouchProbe.Y = CDouble(Row("OffsetTouchProbeY"))
                    CalibrationFlag = CBoolean(Row("CalibrationFlag"))
                    ImageResolution = CInteger(Row("ImageResolution"))
                    'PixelsToMM = CDouble(Row("PixelsToMM"))
                    GraphicsDisplayResolution.X = CDouble(Row("GraphicsDisplayResolutionX"))
                    GraphicsDisplayResolution.Y = CDouble(Row("GraphicsDisplayResolutionY"))


                    Return True
                Else
                    MsgErr = "Cant find Record ID " + ParameterID.RecordID
                    Return False
                End If
            End Function
        End Class
#End Region

#Region " LightBox "

        <Serializable()> Public Class CIDSLightBox
            'Public Contrast As Integer                         'pat
            Public Brightness As Integer                        'pat
            Public Function SaveData() As Boolean
                MsgErr = ""
                Return SaveFile(ParameterID.RecordID)

                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                ''' end here
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''

                If SaveHardWareID() = False Then
                    Return False
                End If

                DBView = New DataView(DS.Tables("TableLightBox"))
                DBView.RowFilter = "ID = '" + ParameterID.RecordID + "'"
                Try
                    Dim Row As DataRow
                    If DBView.Count = 0 Then
                        Row = DBView.Table.NewRow()
                        Row("ID") = ParameterID.RecordID
                        DBView.Table.Rows.Add(Row)
                    Else
                        Row = DBView(0).Row
                    End If
                    'Row("Contrast") = Contrast
                    'Row(" Brightness") = Brightness

                    UpdateData("SELECT * FROM TableLightBox", "TableLightBox")
                Catch ex As Exception
                    MsgErr = ex.ToString
                    Return False
                End Try

                Return True

            End Function
            Public Function GetData() As Boolean

                MsgErr = ""
                Return Openfile(ParameterID.RecordID)

                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                ''' end here
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''

                DBView = New DataView(DS.Tables("TableCamera"))
                DBView.RowFilter = "ID = '" + ParameterID.RecordID + "'"
                If DBView.Count = 1 Then
                    Dim Row As DataRow = DBView(0).Row
                    'Contrast = CInt(Row("Contrast"))
                    'Brightness = CInt(Row(" Brightness"))
                    Return True
                Else
                    MsgErr = "Cant find Record ID " + ParameterID.RecordID
                    Return False
                End If
            End Function
        End Class
#End Region

#Region " HeightSensor "

        <Serializable()> Public Class CIDSHeightSensor
            'False = Laser, true = TouchProbe
            Public SelectType As Boolean                        'DB
            Public BoardHeight As Double                        'Pat
            Public CalibratorHeightPos As New CIDSPosXY         'DB   

            Public Laser As New CIDSSensor                      'DB   'create instance for Laser sensor
            Public TP As New CIDSSensor                         'DB   'create instance for Touch probe
            <Serializable()> Public Class CIDSSensor
                Public CurrentPos As New CIDSPosXYZ             'Global variable 
                Public OffsetPos As New CIDSPosXY               'DB
                Public HeightReference As Double                'Db
            End Class

            Public LVDTCalBackGround As Boolean = True
            Public LVDTCalBrightness As Integer
            Public LVDTCalThreshold As Integer
            Public LVDTCalOpen As Integer
            Public LVDTCalClose As Integer
            Public LVDTCalMaxRadius As Double
            Public LVDTCalMinRadius As Double
            Public LVDTCalRoughness As Integer
            Public LVDTCalCompactness As Integer

            Public Function SaveData() As Boolean
                MsgErr = ""
                Return SaveFile(ParameterID.RecordID)

                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                ''' end here
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''

                If SaveHardWareID() = False Then
                    Return False
                End If

                DBView = New DataView(DS.Tables("TableHeightSensor"))
                DBView.RowFilter = "ID = '" + ParameterID.RecordID + "'"
                Try
                    Dim Row As DataRow
                    If DBView.Count = 0 Then
                        Row = DBView.Table.NewRow()
                        Row("ID") = ParameterID.RecordID
                        DBView.Table.Rows.Add(Row)
                    Else
                        Row = DBView(0).Row
                    End If
                    Row("SelectType") = SelectType
                    'Row("BoardHeight") = BoardHeight
                    Row("CalibratorHeightPosX") = CalibratorHeightPos.X
                    Row("CalibratorHeightPosY") = CalibratorHeightPos.Y

                    Row("LaserHeightReference") = Laser.HeightReference
                    Row("LaserOffsetPosX") = Laser.OffsetPos.X
                    Row("LaserOffsetPosY") = Laser.OffsetPos.Y

                    Row("TPHeightReference") = TP.HeightReference
                    Row("TPOffsetPosX") = TP.OffsetPos.X
                    Row("TPOffsetPosY") = TP.OffsetPos.Y

                    UpdateData("SELECT * FROM TableHeightSensor", "TableHeightSensor")
                Catch ex As Exception
                    MsgErr = ex.ToString
                    Return False
                End Try
                Return True

            End Function
            Public Function GetData() As Boolean

                MsgErr = ""
                Return Openfile(ParameterID.RecordID)

                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                ''' end here
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''

                DBView = New DataView(DS.Tables("TableHeightSensor"))
                DBView.RowFilter = "ID = '" + ParameterID.RecordID + "'"
                If DBView.Count = 1 Then
                    Dim Row As DataRow = DBView(0).Row
                    SelectType = CBoolean(Row("SelectType"))
                    'BoardHeight = CDouble(Row("BoardHeight"))
                    CalibratorHeightPos.X = CDouble(Row("CalibratorHeightPosX"))
                    CalibratorHeightPos.Y = CDouble(Row("CalibratorHeightPosY"))

                    Laser.HeightReference = CDouble(Row("LaserHeightReference"))
                    Laser.OffsetPos.X = CDouble(Row("LaserOffsetPosX"))
                    Laser.OffsetPos.Y = CDouble(Row("LaserOffsetPosY"))

                    TP.HeightReference = CDouble(Row("TPHeightReference"))
                    TP.OffsetPos.X = CDouble(Row("TPOffsetPosX"))
                    TP.OffsetPos.Y = CDouble(Row("TPOffsetPosY"))

                    Return True
                Else
                    MsgErr = "Cant find Record ID " + ParameterID.RecordID
                    Return False
                End If
            End Function
        End Class
#End Region

#Region " Needle "

        <Serializable()> Public Class CIDSNeedle
            Public Left As New CIDSNeedlesParameter("LeftNeedle")   'create instance for left needle
            Public Right As New CIDSNeedlesParameter("RightNeedle") 'create instance for right needle

            <Serializable()> Public Class CIDSNeedlesParameter
                Shared Selection As String
                Public Sub New(ByVal _Selection As String)           'select needle = left or right
                    Selection = _Selection
                    DotDiameter = 2.5
                End Sub

                Public CalibratorPos As New CIDSPosXYZ              'DB
                Public HeightApproach As Integer                    'DB AND PAT
                Public HeightRetract As Integer                     'DB AND PAT
                Public HeightClearance As Double                    'DB AND PAT
                Public ArcRadius As Double                          'DB AND PAT

                Public NeedleCalibrationPosition As New CIDSPosXYZ  'the z position is just above the touch sensor
                Public TouchSensorZPosition                         'the z position is the point at which the touch sensor registers an input.

                ''' [version 1.1]
                Public DispenseDot As New CIDSDispenseDot                       'DB AND PAT 'create instance 
                <Serializable()> Public Class CIDSDispenseDot
                    Public DispenseDuration As Double
                    Public NeedleGap As Double
                    Public ApproachHeight As Double
                    Public RetractDelay As Double
                    Public RetractSpeed As Double
                    Public RetractHeight As Double
                    Public ClearanceHeight As Double
                    Public ArcRadius As Double
                End Class
                ''' [version 1.1]

                Public ArrayDotPos1 As New CIDSPosXYZ                           'DB and Pat
                Public ArrayDotPos3 As New CIDSPosXYZ                           'DB and Pat

                'false for black, white for true
                Public CalBackground As Boolean = True
                Public CalBrightness As Integer
                Public CalThreshold As Integer
                Public CalOpen As Integer
                Public CalClose As Integer
                Public CalMaxRadius As Double
                Public CalMinRadius As Double
                Public CalRoughness As Double
                Public CalCompactness As Double

                Public DotDiameter As Double

                'Table Needle
                Public Function SaveData() As Boolean
                    MsgErr = ""
                    Return SaveFile(ParameterID.RecordID)

                    ''' end here

                    If SaveHardWareID() = False Then
                        Return False
                    End If

                    DBView = New DataView(DS.Tables("TableNeedle"))
                    DBView.RowFilter = "NeedleID = '" + ParameterID.RecordID + "' AND LeftRight = '" + Selection.ToString + "'"

                    Try
                        Dim Row As DataRow
                        If DBView.Count = 0 Then
                            Row = DBView.Table.NewRow()
                            Row("ID") = ParameterID.RecordID
                            Row("Selection") = Selection
                            DBView.Table.Rows.Add(Row)
                        Else
                            Row = DBView(0).Row
                        End If

                        Row("CalibratorPosX") = CalibratorPos.X
                        Row("CalibratorPosY") = CalibratorPos.Y
                        Row("CalibratorPosZ") = CalibratorPos.Z
                        Row("HeightApproach") = HeightApproach
                        Row("HeightRetract") = HeightRetract
                        Row("HeightClearance") = HeightClearance
                        Row("ArcRadius") = ArcRadius

                        UpdateData("SELECT * FROM TableNeedle", "TableNeedle")
                    Catch ex As Exception
                        MsgErr = ex.ToString
                        Return False
                    End Try

                    Return True

                End Function
                Public Function GetData() As Boolean

                    MsgErr = ""
                    Return Openfile(ParameterID.RecordID)

                    ''' end here

                    DBView = New DataView(DS.Tables("Tableneedle"))
                    DBView.RowFilter = "NeedleID = '" + ParameterID.RecordID + "' AND LeftRight = '" + Selection.ToString + "'"

                    If DBView.Count = 1 Then
                        Dim Row As DataRow = DBView(0).Row

                        CalibratorPos.X = CDouble(Row("CalibratorPosX"))
                        CalibratorPos.Y = CDouble(Row("CalibratorPosY"))
                        CalibratorPos.Z = CDouble(Row("CalibratorPosZ"))

                        HeightApproach = CInteger(Row("HeightApproach"))
                        HeightRetract = CInteger(Row("HeightRetract"))
                        HeightClearance = CDouble(Row("HeightClearance"))
                        ArcRadius = CDouble(Row("ArcRadius"))

                        Return True
                    Else
                        MsgErr = "Cant find Record ID " + ParameterID.RecordID
                        Return False
                    End If
                End Function
            End Class




        End Class
#End Region

#Region " Volume "

        <Serializable()> Public Class CIDSVolume

            Public WeighingScaleFlag As Boolean                 'DB

            Public Left As New CIDSVolumeParameter("Left")      'initialize left volume object   
            Public Right As New CIDSVolumeParameter("Right")    'initialize right volume object

            <Serializable()> Public Class CIDSVolumeParameter
                Public Selection As String
                Public Sub New(ByVal _Selection As String)      'initializw with selection = left or right
                    Selection = _Selection
                End Sub

                'calibration options
                Public OnOff As Boolean
                '1) let the user set desired weight and find out the pressure/RPM/dispense duration to obtain it
                '2) dispense with parameters obtained through trial and error in programming mode and then save the weight of the dispensed dot
                Public CalibrationType As String
                Public CalibrationInterval As Double 'minutes

                'setup parameters - display in calibration settings
                Public DesiredWeight As Double
                Public Tolerance As Double
                Public SetupDispenseDuration As Double
                Public SetupMaterialAirPressure As Double
                Public SetupRPM As Double
                Public NumberOfAttempts As Double 'before stopping auto-calibration
                Public RetractSpeed As Double
                Public RetractHeight As Double 'before stopping auto-calibration
                Public RetractDelay As Double


                Public PressureStepValue As Double
                Public RPMStepValue As Double
                Public DurationStepValue As Double

                'adjusted parameters - use in production
                Public AdjustedMaterialAirPressure As Double 'for time pressure valve/syringe
                Public AdjustedRPM As Double 'for auger valve
                Public AdjustedDispenseDuration As Double 'for slider, jetting valve

                Public PulseOnDuration As Double 'Pulse on and off duration consider as on cycle
                Public PulseOffDuration As Double 'Jetting valve only dispesing at pulse on duration
                Public JettingNoOfDispense As Double

            End Class

            Public Function SaveData() As Boolean
                MsgErr = ""
                Return SaveFile(ParameterID.RecordID)

                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                ''' end here
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''

                If SaveHardWareID() = False Then
                    Return False
                End If

                DBView = New DataView(DS.Tables("TableVolume"))
                DBView.RowFilter = "ID = '" + ParameterID.RecordID + "'"
                Try
                    Dim Row As DataRow
                    If DBView.Count = 0 Then
                        Row = DBView.Table.NewRow()
                        Row("ID") = ParameterID.RecordID
                        DBView.Table.Rows.Add(Row)
                    Else
                        Row = DBView(0).Row
                    End If
                    Row("WeighingScaleFlag") = WeighingScaleFlag


                    UpdateData("SELECT * FROM TableVolume", "TableVolume")
                Catch ex As Exception
                    MsgErr = ex.ToString
                    Return False
                End Try
                Return True

            End Function
            Public Function GetData() As Boolean

                MsgErr = ""
                Return Openfile(ParameterID.RecordID)

                ''' end here

                DBView = New DataView(DS.Tables("TableVolume"))
                DBView.RowFilter = "ID = '" + ParameterID.RecordID + "'"
                If DBView.Count = 1 Then
                    Dim Row As DataRow = DBView(0).Row
                    WeighingScaleFlag = CBoolean(Row("WeighingScaleFlag"))

                    Return True
                Else
                    MsgErr = "Cant find Record ID " + ParameterID.RecordID
                    Return False
                End If
            End Function
        End Class
#End Region

#Region " Conveyor "
        <Serializable()> Public Class CIDSConveyor

            Public Width As Double
            Public FullWidth As Integer                     'DB
            Public MinWidth As Integer                      'DB  
            Public Speed As Integer                         'PAT 
            Public TimeOut As Integer                       'PAT
            Public WidthMoveStep As Double                  'PAT
            Public upstreamTimeout As Integer
            Public downstreamTimeout As Integer

            Public Function SaveData() As Boolean
                MsgErr = ""
                Return SaveFile(ParameterID.RecordID)

                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                ''' end here
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''

                If SaveHardWareID() = False Then
                    Return False
                End If

                DBView = New DataView(DS.Tables("TableConveyor"))
                DBView.RowFilter = "ID = '" + ParameterID.RecordID + "'"

                Try
                    Dim Row As DataRow
                    If DBView.Count = 0 Then
                        Row = DBView.Table.NewRow()
                        Row("ID") = ParameterID.RecordID
                        DBView.Table.Rows.Add(Row)
                    Else
                        Row = DBView(0).Row
                    End If

                    Row("FullWidth") = FullWidth
                    Row("MinWidth") = MinWidth

                    UpdateData("SELECT * FROM TableConveyor", "TableConveyor")
                Catch ex As Exception
                    MsgErr = ex.ToString
                    Return False
                End Try

                Return True
            End Function
            Public Function GetData() As Boolean

                MsgErr = ""
                Return Openfile(ParameterID.RecordID)

                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                ''' end here
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''

                DBView = New DataView(DS.Tables("TableConveyor"))
                DBView.RowFilter = "ID = '" + ParameterID.RecordID + "'"
                If DBView.Count = 1 Then
                    Dim Row As DataRow = DBView(0).Row

                    FullWidth = CInteger(Row("FullWidth"))
                    MinWidth = CInteger(Row("MinWidth"))

                    Return True
                Else
                    MsgErr = "Cant find Record ID " + ParameterID.RecordID
                    Return False
                End If

            End Function
        End Class
#End Region


#Region "Heater"

        <Serializable()> Public Class CIDSThermal

            Public HeaterFeatureEnabled As Boolean
            Public Needle As New CIDSHeater("Needle")            'create instance for needle thermal
            Public Syringe As New CIDSHeater("Syringe")           'create instance for syringe thermal
            Public Station1 As New CIDSHeater("Station1")          'create instance for station1 thermal
            Public Station2 As New CIDSHeater("Station2")          'create instance for station2 thermal
            Public Station3 As New CIDSHeater("Station3")          'create instance for station3 thermal
            <Serializable()> Public Class CIDSHeater
                Public ControllerID As String
                Public Sub New(ByVal _HeaterName As String)         'selection = needle or syringe or staion 1 or 2 or 3
                    ControllerID = _HeaterName
                End Sub

                Public OperationTemp As Double                   'DB 'kr nov 21
                Public StandbyTemp As Double                     'DB 'kr nov 21
                Public AlarmThreshold As Double                  'DB and Pat
                Public OnOff As Boolean                          'DB

                Public Function SaveData() As Boolean
                    MsgErr = ""
                    Return SaveFile(ParameterID.RecordID)

                    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    ''' end here
                    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''

                    If SaveHardWareID() = False Then
                        Return False
                    End If

                    Try
                        DBView = New DataView(DS.Tables("TableThermal"))
                        DBView.RowFilter = "ID = '" + ParameterID.RecordID + " ' AND Name = '" + ControllerID.ToString + "'"

                        Dim Row As DataRow
                        If DBView.Count = 0 Then
                            Row = DBView.Table.NewRow()
                            Row("ID") = ParameterID.RecordID
                            Row("Name") = ControllerID.ToString
                            DBView.Table.Rows.Add(Row)
                        Else
                            Row = DBView(0).Row
                        End If

                        Row("Operation Temp") = OperationTemp
                        Row("Standby Temp") = StandbyTemp
                        Row("Threshold") = AlarmThreshold

                        UpdateData("SELECT * FROM TableThermal", "TableThermal")

                    Catch ex As Exception
                        MsgErr = ex.ToString
                        Return False
                    End Try

                    Return True
                End Function

                Public Function GetData() As Boolean

                    MsgErr = ""
                    Return Openfile(ParameterID.RecordID)

                    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    ''' end here
                    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''

                    DBView = New DataView(DS.Tables("TableThermal"))
                    DBView.RowFilter = "ID = '" + ParameterID.RecordID + " ' AND Name = '" + ControllerID.ToString + "'"
                    If DBView.Count = 1 Then

                        Dim Row As DataRow = DBView(0).Row

                        OperationTemp = CDouble(Row("Operation Temp"))
                        StandbyTemp = CDouble(Row("Standby Temp"))
                        AlarmThreshold = CDouble(Row("Threshold"))

                    Else

                        MsgErr = "Cant find Record ID " + ParameterID.RecordID + " Name = " + ControllerID
                        Return False

                    End If
                End Function
            End Class
        End Class
#End Region

#Region " Dispenser "

        <Serializable()> Public Class CIDSDispenser

            Public CurrentHeads As Integer

            Public Left As New CIDSDispenserParameter("Left")               'create instance for left dispenser
            Public Right As New CIDSDispenserParameter("Right")             'create instance for right dispenser

            <Serializable()> Public Class CIDSDispenserParameter
                Public Selection As String                                  'selection = left or right dispenser
                Public Sub New(ByVal _Selection As String)
                    Selection = _Selection
                End Sub

                'auger valve: AVC 1000
                'jetting valve: JVC 1000
                'slider valve: SVC 1000
                'from the above, you can get the idea of the pattern...
                'many of these variables are critical to the volume calibration process

                'dispensing of auger valve is controlled by RPM, which we send to the auger valve controller
                'for jetting, auger and slider valve, material air pressure remains as a constant as we perform volume calibration
                'for jetting and slider valve, dispensing time is a variable as we perform volume calibration
                'for time pressure syringe, material air pressure is a variable as we perform volume calibration
                Public MaterialAirPressure As Double
                Public SuckbackPressure As Double
                'for time auger valve, RPM is a variable as we perform volume calibration, which we send to the auger valve controller
                Public RPM As String
                Public RetractTime As String
                Public RetractDelay As String

                Public Pulse As String 'milliseconds
                Public Pause As String 'milliseconds
                Public Count As String 'number of times to repeat pulse cycle of on/off when IO triggers
                Public ValveTemperature As Double 'degree celsius

                'these are the information variables
                'operators use them to identify different dispensing needle tips quickly
                Public NeedleTipLength As String '0.25, 0.5, 1, 1.5 inches
                Public NeedleColor As String     'see technodigm brochure for available colors. they correspond to needle diameter
                Public MaterialInfo As String

                'maintenance variables
                Public AutoPurgingDuration As Integer 'milliseconds
                Public AutoCleaningDuration As Integer 'milliseconds
                Public AutoPurgingInterval As Integer 'minutes
                Public PotLifeDuration As Integer 'minutes
                'material pot life refers to the material's shelf life so to speak
                Public PotLifeOption As Boolean
                'auto purging refers to the process by which we conduct purging and cleaning in that order 
                'after a fixed amount of time in order to prevent the material from curing and clogging up the dispensing tip
                'such materials usually have a high alcohol content
                Public AutoPurgingOption As Boolean

                Public MaterialSensorEnabled As Boolean
                Public HeadType As String
                'Jetting Valve
                'Auger Valve
                'Slider Valve
                'Time Pressure Valve
                'Time Pressure Syringe
                'These 5 settings were used for safety checking when doing calibration
                'When doing calibration, there are 4 step. If for newly setup dispenser
                'user must start with step 1 until step 4
                'With these 5 settings, user are forced to start from step 1 to 4 when setup a
                'new dispenser but only once. These settings also can be a flage that indicate
                'that what type of dispenser had been used on this system before.
                Public JettingCalOnce As Boolean
                Public AugerCalOnce As Boolean
                Public SlideValveCalOnce As Boolean
                Public TimePressureSyringeCalOnce As Boolean
                Public TimePressureValveCalOnce As Boolean

                Public PressControlTable As New CIDSPressControlTable(99)   'create instance for control table array size of 100
                <Serializable()> Public Class CIDSPressControlTable
                    Public Sub New(ByVal _Length As Integer)
                        Call Length(_Length)
                    End Sub
                    Public PressControlElement() As CIDSPressControlElement 'declare array without instance
                    Public Sub Length(ByVal Number As Integer)
                        ReDim PressControlElement(Number)                   'initialize the array size
                        Dim I As Integer
                        For I = 0 To Number
                            PressControlElement(I) = New CIDSPressControlElement
                        Next
                    End Sub
                    Public Function Length() As Integer                     'return the length of array size
                        Return PressControlElement.Length
                    End Function

                    <Serializable()> Public Class CIDSPressControlElement
                        Public Index As Integer
                        Public Pressure As Double = 0
                        Public SuckBack As Double = 0
                        Public Duration As Integer = 0
                    End Class
                End Class
            End Class

            Public Function SaveData() As Boolean
                MsgErr = ""
                Return SaveFile(ParameterID.RecordID)

                If SaveHardWareID() = False Then
                    Return False
                End If

                DBView = New DataView(DS.Tables("TableDispenser"))
                DBView.RowFilter = "ID = '" + ParameterID.RecordID + "'"
                Try
                    Dim Row As DataRow
                    If DBView.Count = 0 Then
                        Row = DBView.Table.NewRow()
                        Row("ID") = ParameterID.RecordID
                        DBView.Table.Rows.Add(Row)
                    Else
                        Row = DBView(0).Row
                    End If
                    Row("NumOfHead") = CurrentHeads

                    UpdateData("SELECT * FROM TableDispenser", "TableDispenser")
                Catch ex As Exception
                    MsgErr = ex.ToString
                    Return False
                End Try
                Return True

            End Function
            Public Function GetData() As Boolean

                MsgErr = ""
                Return Openfile(ParameterID.RecordID)

                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                ''' end here
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''

                DBView = New DataView(DS.Tables("TableDispenser"))
                DBView.RowFilter = "ID = '" + ParameterID.RecordID + "'"
                If DBView.Count = 1 Then
                    Dim Row As DataRow = DBView(0).Row

                    CurrentHeads = CInteger(Row("CurrentHeads"))

                    Return True
                Else
                    MsgErr = "Cant find Record ID " + ParameterID.RecordID
                    Return False
                End If
            End Function
        End Class
#End Region

#Region " Regulator "


        <Serializable()> Public Class CIDSRegulator
            Public Left As New CIDSRegulatorParameter("Left")       'create instance for left regulator
            Public Right As New CIDSRegulatorParameter("Right")     'create instance for right regulator

            <Serializable()> Public Class CIDSRegulatorParameter
                Public Selection As String                          'selection = left or right regulator
                Public Sub New(ByVal _Selection As String)
                    Selection = _Selection
                End Sub
                Public Pressure As Double
                Public Vacuum As Double
            End Class
            Public Function SaveData() As Boolean
                MsgErr = ""
                Return SaveFile(ParameterID.RecordID)
            End Function
            Public Function GetData() As Boolean

                MsgErr = ""
                Return Openfile(ParameterID.RecordID)
            End Function

        End Class

#End Region

#Region " SPC "

        <Serializable()> Public Class CIDSSPC
            Public FiducialFailedAction As Boolean              'Pat
            Public HeightFailedAction As Boolean                'Pat
            Public ChipFailedAction As Boolean                  'Pat
            Public QCFailedAction As Boolean                    'Pat
            Public EnableSPCReport As Boolean                   'Pat
            Public CleanUpInterval As Integer                   'DB
            Public ItemsToBeReported As String                  'pat
            Public ReportFileName As String                     'pat
            Public ProgAuthorName As String                     'pat
            Public ProgAuthorContact As String                  'pat
            Public ProductionNote As String                     'pat

            Public Function SaveData() As Boolean
                MsgErr = ""
                Return SaveFile(ParameterID.RecordID)

                If SaveHardWareID() = False Then
                    Return False
                End If

                DBView = New DataView(DS.Tables("TableSPC"))
                DBView.RowFilter = "ID = '" + ParameterID.RecordID + "'"

                Try
                    Dim Row As DataRow
                    If DBView.Count = 0 Then
                        Row = DBView.Table.NewRow()
                        Row("ID") = ParameterID.RecordID
                        DBView.Table.Rows.Add(Row)
                    Else
                        Row = DBView(0).Row
                    End If

                    Row("CleanUpInterval") = CleanUpInterval


                    UpdateData("SELECT * FROM TableSPC", "TableSPC")
                Catch ex As Exception
                    MsgErr = ex.ToString
                    Return False
                End Try

                Return True
            End Function
            Public Function GetData() As Boolean
                MsgErr = ""
                Return Openfile(ParameterID.RecordID)

                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                ''' end here
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''

                DBView = New DataView(DS.Tables("TableSPC"))
                DBView.RowFilter = "ID = '" + ParameterID.RecordID + "'"
                If DBView.Count = 1 Then
                    Dim Row As DataRow = DBView(0).Row

                    CleanUpInterval = CInteger(Row("CleanUpInterval"))
                    Return True
                Else
                    MsgErr = "Cant find Record ID " + ParameterID.RecordID
                    Return False
                End If

            End Function


        End Class

#End Region

#Region " SystemIO "

        <Serializable()> Public Class CIDSSystemIO
            Public ItemArray As New ArrayList           'create array list
            Public Template As New CIDSIO               'create instance of systemio object
            ''''''''''''''''''''''''''''''''''''''''''''''''''''
            ''''''''''''''' not use
            ''''''''''''''''''''''''''''''''''''''''''''''''''''
            Public Sub Length(ByVal Number As Integer)
                ReDim Item(Number)
                Dim I As Integer
                For I = 0 To Number
                    Item(I) = New CIDSIO
                Next
            End Sub
            Public Function Length() As Integer
                Return Item.Length
            End Function
            Public Shared Item() As CIDSIO
            '''''''''''''''''''''''''''''''''''''''''''''''''''''
            <Serializable()> Public Class CIDSIO
                Public IOName As String                 'DB
                'Public IOAddress As String             'Db
                Public Type As String                   'DB
                Public Status As String                 'DB
                Public IO As String                     'DB
                Public CableName As String              'DB
                Public ModuleName As String             'DB
            End Class
            Public Function GetData() As Boolean
                MsgErr = ""
                Return Openfile(ParameterID.RecordID)

                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                ''' end here
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''

                DBView = New DataView(DS.Tables("TableSystemIO"))
                DBView.RowFilter = "ID = '" + ParameterID.RecordID + "'"

                If DBView.Count > 0 Then
                    Dim I As Integer = 0
                    Length(DBView.Count - 1)
                    For I = 0 To DBView.Count - 1
                        Dim Row As DataRow = DBView(I).Row

                        Item(I).IOName = CString(Row("IOName"))
                        Item(I).Type = CString(Row("Type"))
                        Item(I).Status = CString(Row("Status"))
                        Item(I).IO = CString(Row("IO"))
                        Item(I).CableName = CString(Row("CableName"))
                        Item(I).ModuleName = CString(Row("ModuleName"))


                        ItemArray.Add(Item(I))
                    Next
                    Return True
                Else
                    MsgErr = "Cant find Record ID " + ParameterID.RecordID
                    Return False
                End If

            End Function

            Public Function SaveData() As Boolean
                MsgErr = ""
                Return SaveFile(ParameterID.RecordID)

                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                ''' end here
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''

                If SaveHardWareID() = False Then
                    Return False
                End If

                Try
                    DBView = New DataView(DS.Tables("TableSystemIO"))
                    DBView.RowFilter = "ID = '" + ParameterID.RecordID + "'"

                    If DBView.Count > 0 Then
                        Dim J As Integer = 0
                        For J = 0 To DBView.Count - 1
                            Dim row As DataRow = DBView(0).Row
                            row.Delete()
                        Next
                    End If


                    DBView = New DataView(DS.Tables("TableSystemIO"))
                    Dim I As Integer = 0
                    For I = 0 To Item.Length - 1
                        Dim Row As DataRow
                        Row = DBView.Table.NewRow()
                        Row("ID") = ParameterID.RecordID
                        Row("IOName") = Item(I).IOName
                        Row("Type") = Item(I).Type
                        Row("Status") = Item(I).Status
                        Row("IO") = Item(I).IO
                        Row("CableName") = Item(I).CableName
                        Row("ModuleName") = Item(I).ModuleName
                        DBView.Table.Rows.Add(Row)
                    Next

                    UpdateData("SELECT * FROM TableSystemIO", "TableSystemIO")
                Catch ex As Exception
                    MsgErr = ex.ToString
                    Return False
                End Try
                Return True
            End Function
        End Class
#End Region

#End Region

    End Class

#End Region

#Region "Execution"

    <Serializable()> Public Class CIDSPatternExecution
        Public Prog As New ArrayList                                    'create instance for pattern program
        Public Pattern As New CIDSTemplate                              'create instance for pattern parameter template
        Public Setting As New CIDSExecutionSetting                      'create instance for patter setting related variable
#Region " Settings "
        <Serializable()> Public Class CIDSExecutionSetting
            Public PatternDir As String = "?IDSPatternFiles/"
            Public CurrentPatternName As String = ""
            Public DotTemplateDir As String = "?IDSPttnTempletDir/Dot/"
            Public LineTemplateDir As String = "?IDSPttnTempletDir/Line/"
            Public ArcTemplateDir As String = "?IDSPttnTempletDir/Arc/"
            Public RecTemplateDir As String = "?IDSPttnTempletDir/Rectangle/"
            Public CircleTemplateDir As String = "?IDSPttnTempletDir/Circle/"
            Public FilledRecTemplateDir As String = "?IDSPttnTempletDir/FilledRectangle/"
            Public FilledCircleTemplateDir As String = "?IDSPttnTempletDir/FilledCircle/"
            Public ChipEdgeTemplateDir As String = "?IDSPttnTempletDir/ChipEdge/"
            Public DefaultFileToOpen As String = ""

            Public Function SaveData() As Boolean
                MsgErr = ""
                Return SaveFile(ParameterID.RecordID)

                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                ''' end here
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            End Function
            Public Function GetData() As Boolean
                MsgErr = ""
                Return Openfile(ParameterID.RecordID)

                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                ''' end here
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            End Function
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '''  not used, for reference only
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Public Function ListAllExecutionRecordID(ByRef RecordIDArray As ComboBox) As Boolean
                MsgErr = ""
                DBView = New DataView(DS.Tables("TableExecution"))
                Try
                    If DBView.Count > 0 Then
                        Dim I As Integer = 0
                        RecordIDArray.Items.Clear()
                        For I = 0 To DBView.Count - 1
                            Dim Row As DataRow = DBView(I).Row
                            RecordIDArray.Items.Add(Row("ID"))
                        Next
                    End If

                Catch ex As Exception
                    MsgErr = ex.ToString
                    Return False
                End Try
                Return True

            End Function
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '''  not used, for reference only
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Public Function DeleteExecutionRecordID() As Boolean

                MsgErr = ""
                DBView = New DataView(DS.Tables("TableExecution"))
                DBView.RowFilter = "ID = '" + ParameterID.RecordID + "'"
                Try
                    Dim Row As DataRow
                    If DBView.Count <> 0 Then
                        If ParameterID.RecordID = "FactoryDefault" = False Then
                            Row = DBView(0).Row
                            Row.Delete()
                        Else
                        End If
                    End If

                    UpdateData("SELECT * FROM TableExecution", "TableExecution")
                Catch ex As Exception
                    MsgErr = ex.ToString
                    Return False
                End Try
                Return True
            End Function
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '''  not used, for reference only
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Public Function SaveExecutionRecordID() As Boolean
                MsgErr = ""
                DBView = New DataView(DS.Tables("TableExecution"))
                DBView.RowFilter = "ID = '" + ParameterID.RecordID + "'"
                Try
                    Dim Row As DataRow
                    If DBView.Count = 0 Then
                        Row = DBView.Table.NewRow()
                        Row("ID") = ParameterID.RecordID
                        DBView.Table.Rows.Add(Row)
                    Else
                        Row = DBView(0).Row
                    End If

                    Row("PatternDir") = PatternDir
                    Row("CurrentPatternName") = CurrentPatternName
                    Row("DotTemplateDir") = DotTemplateDir
                    Row("LineTemplateDir") = LineTemplateDir
                    Row("ArcTemplateDir") = ArcTemplateDir
                    Row("RecTemplateDir") = RecTemplateDir
                    Row("CircleTemplateDir") = CircleTemplateDir
                    Row("FilledRecTemplateDir") = FilledRecTemplateDir
                    Row("FilledCircleTemplateDir") = FilledCircleTemplateDir
                    Row("ChipEdgeTemplateDir") = ChipEdgeTemplateDir


                    UpdateData("SELECT * FROM TableExecution", "TableExecution")
                Catch ex As Exception
                    MsgErr = ex.ToString
                    Return False
                End Try
                Return True
            End Function
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '''  not used, for reference only
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Public Function ReadExecutionRecordID() As Boolean
                MsgErr = ""
                DBView = New DataView(DS.Tables("TableExecution"))
                DBView.RowFilter = "ID = '" + ParameterID.RecordID + "'"
                Dim Row As DataRow
                If DBView.Count <> 0 Then
                    Row = DBView(0).Row

                    PatternDir = CString(Row("PatternDir"))
                    CurrentPatternName = CString(Row("CurrentPatternName"))
                    DotTemplateDir = CString(Row("DotTemplateDir"))
                    LineTemplateDir = CString(Row("LineTemplateDir"))
                    ArcTemplateDir = CString(Row("ArcTemplateDir"))
                    RecTemplateDir = CString(Row("RecTemplateDir"))
                    CircleTemplateDir = CString(Row("CircleTemplateDir"))
                    FilledRecTemplateDir = CString(Row("FilledRecTemplateDir"))
                    FilledCircleTemplateDir = CString(Row("FilledCircleTemplateDir"))
                    ChipEdgeTemplateDir = CString(Row("ChipEdgeTemplateDir"))
                    Return True
                Else
                    MsgErr = "Cant find Record ID " + ParameterID.RecordID
                    Return False
                End If
            End Function
        End Class
#End Region

#Region " Pattern Service "

        <Serializable()> Public Class CIDSTemplate
            Public Item As Integer
            Public Name As String
            Public Pos1 As New CIDSPosXYZ
            Public Pos2 As New CIDSPosXYZ
            Public pos3 As New CIDSPosXYZ
            Public DispenseFlag As String = "On"
            Public NeedleGap As Double
            Public DispenseDuration As Integer
            Public ApproachDispHeight As Double
            Public TravelDelay As Integer
            Public TravelSpeed As Double
            Public DetailingDistance As Double
            Public RetractDelay As Integer
            Public RetractSpeed As Double
            Public RetractHeight As Double
            Public ClearanceHeight As Double
            Public ArcRadius As Double
            Public Pitch As Double
            Public FilledHeight As Double
            Public RotatingAngle As Integer
            Public SubFileName As String
            Public EdgeClearance As Double
            Public SpiralFlag As String = "IN"
            Public Needle As String = "LEFT"
            Public IONumber As Integer = 0
            Public NumofRun As Integer = 0

            Public Function SaveData() As Boolean
                MsgErr = ""
                Return SaveFile(ParameterID.RecordID)
            End Function

            Public Function GetData() As Boolean
                MsgErr = ""
                Return Openfile(ParameterID.RecordID)
            End Function
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '''  not used, for reference only
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Public Function ListAllPatternID(ByRef RecordIDArray As ComboBox) As Boolean
                MsgErr = ""
                DBView = New DataView(DS.Tables("TablePattern"))
                DBView.Sort = "ID ASC"

                Try
                    If DBView.Count > 0 Then
                        Dim I As Integer = 0
                        Dim Old As String = ""
                        RecordIDArray.Items.Clear()
                        For I = 0 To DBView.Count - 1
                            Dim Row As DataRow = DBView(I).Row

                            If Old <> CString(Row("id")) Then
                                RecordIDArray.Items.Add(Row("ID"))
                            End If
                            Old = CString(Row("id"))

                        Next
                    End If

                Catch ex As Exception
                    MsgErr = ex.ToString
                    Return False
                End Try
                Return True

            End Function
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '''  not used, for reference only
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Public Function DeletePatternRecordID() As Boolean

                MsgErr = ""
                DBView = New DataView(DS.Tables("TablePattern"))
                DBView.RowFilter = "ID = '" + ParameterID.RecordID + "'"
                Try
                    If DBView.Count > 0 Then
                        Dim I As Integer = 0
                        For I = 0 To DBView.Count - 1
                            Dim Row As DataRow = DBView(0).Row
                            Row.Delete()
                        Next
                    End If

                    UpdateData("SELECT * FROM TablePattern", "TablePattern")

                Catch ex As Exception
                    MsgErr = ex.ToString
                    Return False
                End Try
                Return True
            End Function
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '''  not used, for reference only
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Public Function AppendItem() As Boolean
                MsgErr = ""
                DBView = New DataView(DS.Tables("TablePattern"))
                DBView.RowFilter = "ID = '" + ParameterID.RecordID + "'"
                Try
                    Dim Row As DataRow

                    Row = DBView.Table.NewRow()
                    Row("ID") = ParameterID.RecordID

                    Dim _Item As Integer = DBView.Count
                    Row("item") = _Item
                    DBView.Table.Rows.Add(Row)

                    Call PatternRowEdit(Row, _Item)

                    UpdateData("SELECT * FROM TablePattern", "TablePattern")
                Catch ex As Exception
                    MsgErr = ex.ToString
                    Return False
                End Try
                Return True
            End Function
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '''  not used, for reference only
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Public Function WriteItem(ByRef _Item As Integer) As Boolean
                MsgErr = ""
                DBView = New DataView(DS.Tables("TablePattern"))
                DBView.RowFilter = "ID = '" + ParameterID.RecordID + "' AND Item = " + _Item.ToString
                Try
                    Dim Row As DataRow
                    If DBView.Count = 0 Then
                        Row = DBView.Table.NewRow()
                        Row("ID") = ParameterID.RecordID
                        Row("item") = _Item
                        DBView.Table.Rows.Add(Row)
                    Else
                        Row = DBView(0).Row
                    End If
                    Call PatternRowEdit(Row, _Item)

                Catch ex As Exception
                    MsgErr = ex.ToString
                    Return False
                End Try
                Return True
            End Function
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '''  not used, for reference only
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Private Sub PatternRowEdit(ByRef Row As DataRow, ByRef _Item As Integer)

                Row("item") = _Item
                Row("Name") = Name
                Row("Pos1X") = Pos1.X
                Row("Pos1Y") = Pos1.Y
                Row("Pos1Z") = Pos1.Z

                Row("Pos2X") = Pos2.X
                Row("Pos2Y") = Pos2.Y
                Row("Pos2Z") = Pos2.Z

                Row("Pos3X") = pos3.X
                Row("Pos3Y") = pos3.Y
                Row("Pos3Z") = pos3.Z

                Row("DispenseFlag") = DispenseFlag
                Row("NeedleGap") = NeedleGap
                Row("DispenseDuration") = DispenseDuration
                Row("ApproachDispHeight") = ApproachDispHeight

                Row("TravelDelay") = TravelDelay
                Row("TravelSpeed") = TravelSpeed
                Row("DetailingDistance") = DetailingDistance

                Row("RetractDelay") = RetractDelay
                Row("RetractSpeed") = RetractSpeed
                Row("RetractHeight") = RetractHeight

                Row("ClearanceHeight") = ClearanceHeight
                Row("ArcRadius") = ArcRadius
                Row("Pitch") = Pitch
                Row("FilledHeight") = FilledHeight
                Row("RotatingAngle") = RotatingAngle
                Row("SubFileName") = SubFileName
                Row("EdgeClearance") = EdgeClearance
                Row("SpiralFlag") = SpiralFlag
                Row("Needle") = Needle

            End Sub
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '''  not used, for reference only
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Public Function ReadItem(ByRef Item As Integer) As Boolean
                MsgErr = ""
                DBView = New DataView(DS.Tables("TablePattern"))
                DBView.RowFilter = "ID = '" + ParameterID.RecordID + "' AND Item = " + Item.ToString
                If DBView.Count <> 0 Then
                    Dim Row As DataRow
                    Row = DBView(0).Row

                    Name = CString(Row("Name"))
                    Pos1.X = CDouble(Row("Pos1X"))
                    Pos1.Y = CDouble(Row("Pos1Y"))
                    Pos1.Z = CDouble(Row("Pos1Z"))

                    Pos2.X = CDouble(Row("Pos2X"))
                    Pos2.Y = CDouble(Row("Pos2Y"))
                    Pos2.Z = CDouble(Row("Pos2Z"))

                    pos3.X = CDouble(Row("Pos3X"))
                    pos3.Y = CDouble(Row("Pos3Y"))
                    pos3.Z = CDouble(Row("Pos3Z"))

                    DispenseFlag = CString(Row("DispenseFlag"))
                    NeedleGap = CDouble(Row("NeedleGap"))
                    DispenseDuration = CInteger(Row("DispenseDuration"))
                    ApproachDispHeight = CDouble(Row("ApproachDispHeight"))

                    TravelDelay = CInteger(Row("TravelDelay"))
                    TravelSpeed = CDouble(Row("TravelSpeed"))
                    DetailingDistance = CDouble(Row("DetailingDistance"))

                    RetractDelay = CInteger(Row("RetractDelay"))
                    RetractSpeed = CDouble(Row("RetractSpeed"))
                    RetractHeight = CDouble(Row("RetractHeight"))

                    ClearanceHeight = CDouble(Row("ClearanceHeight"))
                    ArcRadius = CDouble(Row("ArcRadius"))
                    Pitch = CDouble(Row("Pitch"))
                    FilledHeight = CDouble(Row("FilledHeight"))
                    RotatingAngle = CInteger(Row("RotatingAngle"))
                    SubFileName = CString(Row("SubFileName"))
                    EdgeClearance = CDouble(Row("EdgeClearance"))
                    SpiralFlag = CString(Row("SpiralFlag"))
                    Needle = CString(Row("Needle"))
                    Return True
                Else
                    MsgErr = "Cant find Record ID " + ParameterID.RecordID + " / Item " + Item.ToString
                    Return False
                End If
            End Function
        End Class
#End Region
    End Class
#End Region

#Region "Commonly Used Classes"
    '''''''''''''''''''''''''''''''''''''
    ' commonly used Classes
    '''''''''''''''''''''''''''''''''''''
    <Serializable()> Public Class CIDSPosXYZ
        Inherits CIDSPosXY
        Public Z As Double = 0
    End Class
    <Serializable()> Public Class CIDSPosXY
        Public X As Double = 0
        Public Y As Double = 0
    End Class
    <Serializable()> Public Class CIDSParameterID
        'Public ModuleName As String
        Public GroupID As String
        Public RecordID As String = "FactoryDefault"
    End Class
    <Serializable()> Public Class CIDSMinMax
        Public Min As Double = 0
        Public Max As Double = 0
    End Class
#End Region

End Class

Public Module Module1

#Region "Global Instance and functions for other module to call upon"

    Public IDSData As New CIDSData          'create instance of IDS Global data object
    Public DS As New DataSet                'declare dataset for Access database
    Public DBView As New DataView           'declare dataview for Access database

#Region " Pattern File Utilities "

    Dim PatDisplayArray As New ArrayList    'temporary working variable
    Dim PatArray As New ArrayList           'temporary working variable
    'save global data into pat file
    Public Function SaveFile(ByRef FileName As String) As Boolean
        If FileName Is Nothing Then
            FileName = "factorydefault"
        End If
        Dim str As String = IDSData.Admin.Folder.PatternPath + "\" + FileName + "." + IDSData.Admin.Folder.FileExtension
        Return SavePathFileName(str)
    End Function
    Public Function SaveToDefault() As Boolean
        Dim str As String = "C:\IDS\Pattern_Dir\" + "factorydefault" + "." + IDSData.Admin.Folder.FileExtension
        Return SavePathFileName(str)
    End Function
    Public Function OpenToDefault() As Boolean
        Dim str As String = "C:\IDS\Pattern_Dir\" + "factorydefault" + "." + IDSData.Admin.Folder.FileExtension
        Return OpenPathFileName(str)
    End Function
    'open global data into pat file
    Public Function Openfile(ByRef FileName As String) As Boolean
        If FileName Is Nothing Then
            FileName = "factorydefault"
        End If
        Dim str As String = IDSData.Admin.Folder.PatternPath + "\" + FileName + "." + IDSData.Admin.Folder.FileExtension
        Return OpenPathFileName(str)
    End Function

    'Text encrypt                                                     '
    '   Text:     text to be encrypted                                ' 
    '   return:   encrypted text                                      '
    Public Function TextCrypt(ByVal Text As String) As String
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

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'save global data into pat file with path in binary format
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Friend Function SerializePathFileName(ByVal MyFileName As String) As Boolean
        Dim fs As New FileStream(MyFileName, FileMode.Create)
        Dim formatter As New BinaryFormatter
        Try
            formatter.Serialize(fs, IDSData)
        Catch e As SerializationException
            fs.Close()
            IDSData.MsgErr = "Failed to serialize. Reason: " & e.Message
            MessageBox.Show(IDSData.MsgErr)
            Return False
        End Try
        fs.Close()
        Return True
    End Function

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'open global data into pat file with path in binary format
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Friend Function DeserializePathFileName(ByVal MyFileName As String) As Boolean

        Dim fs As New FileStream(MyFileName, FileMode.Open)
        Try
            Dim formatter As New BinaryFormatter
            IDSData = DirectCast(formatter.Deserialize(fs), CIDSData)
        Catch e As SerializationException
            fs.Close()
            If e.Message = "BinaryFormatter Version incompatibility. Expected Version 1.0.  Received Version 224218703.1396984074." Then
                Return OpenPathFileName(MyFileName)
            Else
                IDSData.MsgErr = "Failed to serialize. Reason: " & e.Message
                MessageBox.Show(IDSData.MsgErr)
            End If
            Return False
        End Try
        fs.Close()
        Return True
    End Function

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '                                       kr                                       '
    ' savepathfilename is called to save all the parameters not pertaining to direct '
    ' commands in the pattern list, such as conveyor width and heater temperature    '
    ' save global data into pat file with path in text format
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Public Textflag As Boolean = False
    Public Function SavePathFileName(ByRef MyFileName As String) As Boolean
        Dim I As Integer
        Try
            PatDisplayArray.Clear()                                 'clear working description arraylist 
            PatArray.Clear()                                        'clear working value arraylist


            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            ''''''''''' Admin '''''''''''''''''
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

            PatDisplayArray.Add("")
            PatArray.Add("")
            PatDisplayArray.Add("~")
            PatArray.Add("[VERSION]")                               'database version
            PatDisplayArray.Add("IDSData.SoftwareVersion")
            PatArray.Add(IDSData.SoftwareVersion)

            PatDisplayArray.Add("")
            PatArray.Add("")
            PatDisplayArray.Add("~")                                'define header
            PatArray.Add("[ADMINISTRATION]")
            PatDisplayArray.Add("IDSData.Admin.Folder.DataPath")
            PatArray.Add(IDSData.Admin.Folder.DataPath)
            PatDisplayArray.Add("IDSData.Admin.Folder.PatternPath")
            PatArray.Add(IDSData.Admin.Folder.PatternPath)
            PatDisplayArray.Add("IDSData.Admin.Folder.FileExtension")
            PatArray.Add(IDSData.Admin.Folder.FileExtension)

            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            ''''''''''''''' ParameterID ''''''''''''''''
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            PatDisplayArray.Add("IDSData.ParameterID.GroupID")
            PatArray.Add(IDSData.ParameterID.GroupID)
            PatDisplayArray.Add("IDSData.ParameterID.RecordID")
            PatArray.Add(IDSData.ParameterID.RecordID)

            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            ''''''''''''' data from user and login group 
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

            PatDisplayArray.Add("IDSData.Admin.User.ContactNo")
            PatArray.Add(IDSData.Admin.User.ContactNo)
            PatDisplayArray.Add("IDSData.Admin.User.Department")
            PatArray.Add(IDSData.Admin.User.Department)
            PatDisplayArray.Add("IDSData.Admin.User.ID")
            PatArray.Add(IDSData.Admin.User.ID)
            PatDisplayArray.Add("IDSData.Admin.User.Name")
            PatArray.Add(IDSData.Admin.User.Name)
            PatDisplayArray.Add("IDSData.Admin.User.PassWord")
            PatArray.Add(IDSData.Admin.User.PassWord)
            PatDisplayArray.Add("IDSData.Admin.User.RunApplication")
            PatArray.Add(IDSData.Admin.User.RunApplication)
            PatDisplayArray.Add("IDSData.Admin.User.Group.ID")
            PatArray.Add(IDSData.Admin.User.Group.ID)
            PatDisplayArray.Add("IDSData.Admin.User.Group.Remark")
            PatArray.Add(IDSData.Admin.User.Group.Remark)

            I = 0
            If IDSData.Admin.User.Group.PrivilegeArray.Count > 0 Then
                For I = 0 To IDSData.Admin.User.Group.PrivilegeArray.Count - 1
                    PatDisplayArray.Add("IDSData.Admin.User.Group.PrivilegeArray (" + I.ToString + ")")
                    PatArray.Add(IDSData.Admin.User.Group.PrivilegeArray(I))

                Next
            End If

            I = 0
            If IDSData.Admin.User.Group.SystemHardwareArray.Count > 0 Then
                For I = 0 To IDSData.Admin.User.Group.SystemHardwareArray.Count - 1
                    PatDisplayArray.Add("IDSData.Admin.User.Group.SystemHardwareArray (" + I.ToString + ")")
                    PatArray.Add(IDSData.Admin.User.Group.SystemHardwareArray.Item(I))

                Next
            End If

            I = 0
            If IDSData.Admin.ALLPrivileges.IDArray.Count > 0 Then
                For I = 0 To IDSData.Admin.ALLPrivileges.IDArray.Count - 1
                    PatDisplayArray.Add("IDSData.Admin.ALLPrivileges.IDArray (" + I.ToString + ")")
                    PatArray.Add(IDSData.Admin.ALLPrivileges.IDArray.Item(I))
                Next
            End If
            I = 0
            If IDSData.Admin.ALLPrivileges.TypeArray.Count > 0 Then
                For I = 0 To IDSData.Admin.ALLPrivileges.TypeArray.Count - 1
                    PatDisplayArray.Add("IDSData.Admin.ALLPrivileges.TypeArray (" + I.ToString + ")")
                    PatArray.Add(IDSData.Admin.ALLPrivileges.TypeArray.Item(I))
                Next
            End If

            'camera
            PatDisplayArray.Add("")
            PatArray.Add("")
            PatDisplayArray.Add("~")
            PatArray.Add("[CAMERA]")
            PatDisplayArray.Add("IDSData.Hardware.Camera.CalibrationFlag")
            PatArray.Add(IDSData.Hardware.Camera.CalibrationFlag)
            PatDisplayArray.Add("IDSData.Hardware.Camera.CalibrationPos.X")
            PatArray.Add(IDSData.Hardware.Camera.CalibrationPos.X)
            PatDisplayArray.Add("IDSData.Hardware.Camera.CalibrationPos.Y")
            PatArray.Add(IDSData.Hardware.Camera.CalibrationPos.Y)

            I = 0
            If IDSData.Hardware.Camera.ChipDispenseType.Point.Length > 0 Then
                For I = 0 To IDSData.Hardware.Camera.ChipDispenseType.Point.Length - 1
                    PatDisplayArray.Add("IDSData.Hardware.Camera.ChipDispenseType.Point(" + I.ToString + ").X")
                    PatArray.Add(IDSData.Hardware.Camera.ChipDispenseType.Point(I).X)
                    PatDisplayArray.Add("IDSData.Hardware.Camera.ChipDispenseType.Point(" + I.ToString + ").Y")
                    PatArray.Add(IDSData.Hardware.Camera.ChipDispenseType.Point(I).Y)
                    PatDisplayArray.Add("IDSData.Hardware.Camera.ChipDispenseType.Point(" + I.ToString + ").Z")
                    PatArray.Add(IDSData.Hardware.Camera.ChipDispenseType.Point(I).Z)
                Next
            End If


            I = 0
            If IDSData.Hardware.Camera.Compensation.Matrix.Length > 0 Then
                For I = 0 To IDSData.Hardware.Camera.Compensation.Matrix.Length - 1
                    PatDisplayArray.Add("IDSData.Hardware.Camera.Compensation.Matrix(" + I.ToString + ").W")
                    PatArray.Add(IDSData.Hardware.Camera.Compensation.Matrix(I).W)
                    PatDisplayArray.Add("IDSData.Hardware.Camera.Compensation.Matrix(" + I.ToString + ").X")
                    PatArray.Add(IDSData.Hardware.Camera.Compensation.Matrix(I).X)
                    PatDisplayArray.Add("IDSData.Hardware.Camera.Compensation.Matrix(" + I.ToString + ").Y")
                    PatArray.Add(IDSData.Hardware.Camera.Compensation.Matrix(I).Y)
                    PatDisplayArray.Add("IDSData.Hardware.Camera.Compensation.Matrix(" + I.ToString + ").Z")
                    PatArray.Add(IDSData.Hardware.Camera.Compensation.Matrix(I).Z)
                Next
            End If

            PatDisplayArray.Add("IDSData.Hardware.Camera.EdgeDetect.NeedleOD")
            PatArray.Add(IDSData.Hardware.Camera.EdgeDetect.NeedleOD)
            PatDisplayArray.Add("IDSData.Hardware.Camera.EdgeDetect.ResultFlag")
            PatArray.Add(IDSData.Hardware.Camera.EdgeDetect.ResultFlag)

            I = 0
            If IDSData.Hardware.Camera.EdgeDetect.Point.Length > 0 Then
                For I = 0 To IDSData.Hardware.Camera.EdgeDetect.Point.Length - 1
                    PatDisplayArray.Add("IDSData.Hardware.Camera.EdgeDetect.Point(" + I.ToString + ").X")
                    PatArray.Add(IDSData.Hardware.Camera.EdgeDetect.Point(I).X)
                    PatDisplayArray.Add("IDSData.Hardware.Camera.EdgeDetect.Point(" + I.ToString + ").Y")
                    PatArray.Add(IDSData.Hardware.Camera.EdgeDetect.Point(I).Y)
                    PatDisplayArray.Add("IDSData.Hardware.Camera.EdgeDetect.Point(" + I.ToString + ").Z")
                    PatArray.Add(IDSData.Hardware.Camera.EdgeDetect.Point(I).Z)
                Next
            End If

            PatDisplayArray.Add("IDSData.Hardware.Camera.FiducialPos.X")
            PatArray.Add(IDSData.Hardware.Camera.FiducialPos.X)
            PatDisplayArray.Add("IDSData.Hardware.Camera.FiducialPos.Y")
            PatArray.Add(IDSData.Hardware.Camera.FiducialPos.Y)
            PatDisplayArray.Add("IDSData.Hardware.Camera.FiducialPos.Z")
            PatArray.Add(IDSData.Hardware.Camera.FiducialPos.Z)
            PatDisplayArray.Add("IDSData.Hardware.Camera.FileFiducial1")
            PatArray.Add(IDSData.Hardware.Camera.FileFiducial1)
            PatDisplayArray.Add("IDSData.Hardware.Camera.FileFiducial2")
            PatArray.Add(IDSData.Hardware.Camera.FileFiducial2)
            PatDisplayArray.Add("IDSData.Hardware.Camera.FileRejectMark")
            PatArray.Add(IDSData.Hardware.Camera.FileRejectMark)
            PatDisplayArray.Add("IDSData.Hardware.Camera.GraphicsDisplayResolution.X")
            PatArray.Add(IDSData.Hardware.Camera.GraphicsDisplayResolution.X)
            PatDisplayArray.Add("IDSData.Hardware.Camera.GraphicsDisplayResolution.Y")
            PatArray.Add(IDSData.Hardware.Camera.GraphicsDisplayResolution.Y)
            PatDisplayArray.Add("IDSData.Hardware.Camera.ImageResolution")
            PatArray.Add(IDSData.Hardware.Camera.ImageResolution)
            PatDisplayArray.Add("IDSData.Hardware.Camera.OffsetLaser.X")
            PatArray.Add(IDSData.Hardware.Camera.OffsetLaser.X)
            PatDisplayArray.Add("IDSData.Hardware.Camera.OffsetLaser.Y")
            PatArray.Add(IDSData.Hardware.Camera.OffsetLaser.Y)
            PatDisplayArray.Add("IDSData.Hardware.Camera.OffsetLNeedle.X")
            PatArray.Add(IDSData.Hardware.Camera.OffsetLNeedle.X)
            PatDisplayArray.Add("IDSData.Hardware.Camera.OffsetLNeedle.Y")
            PatArray.Add(IDSData.Hardware.Camera.OffsetLNeedle.Y)
            PatDisplayArray.Add("IDSData.Hardware.Camera.OffsetRNeedle.X")
            PatArray.Add(IDSData.Hardware.Camera.OffsetRNeedle.X)
            PatDisplayArray.Add("IDSData.Hardware.Camera.OffsetRNeedle.Y")
            PatArray.Add(IDSData.Hardware.Camera.OffsetRNeedle.Y)
            PatDisplayArray.Add("IDSData.Hardware.Camera.OffsetTouchProbe.X")
            PatArray.Add(IDSData.Hardware.Camera.OffsetTouchProbe.X)
            PatDisplayArray.Add("IDSData.Hardware.Camera.OffsetTouchProbe.Y")
            PatArray.Add(IDSData.Hardware.Camera.OffsetTouchProbe.Y)
            PatDisplayArray.Add("IDSData.Hardware.Camera.PixelsToMM.X")
            PatArray.Add(IDSData.Hardware.Camera.PixelsToMM.X)
            PatDisplayArray.Add("IDSData.Hardware.Camera.PixelsToMM.Y")
            PatArray.Add(IDSData.Hardware.Camera.PixelsToMM.Y)
            PatDisplayArray.Add("IDSData.Hardware.Camera.PixelRatio")
            PatArray.Add(IDSData.Hardware.Camera.PixelRatio)
            PatDisplayArray.Add("IDSData.Hardware.Camera.ReferencePos.X")
            PatArray.Add(IDSData.Hardware.Camera.ReferencePos.X)
            PatDisplayArray.Add("IDSData.Hardware.Camera.ReferencePos.Y")
            PatArray.Add(IDSData.Hardware.Camera.ReferencePos.Y)
            PatDisplayArray.Add("IDSData.Hardware.Camera.ReferencePos.Z")
            PatArray.Add(IDSData.Hardware.Camera.ReferencePos.Z)
            PatDisplayArray.Add("IDSData.Hardware.Camera.SoftHomePos.X")
            PatArray.Add(IDSData.Hardware.Camera.SoftHomePos.X)
            PatDisplayArray.Add("IDSData.Hardware.Camera.SoftHomePos.Y")
            PatArray.Add(IDSData.Hardware.Camera.SoftHomePos.Y)
            PatDisplayArray.Add("IDSData.Hardware.Camera.SoftHomePos.Z")
            PatArray.Add(IDSData.Hardware.Camera.SoftHomePos.Z)

            PatDisplayArray.Add("IDSData.Hardware.Camera.CalibrationNumColumns")
            PatArray.Add(IDSData.Hardware.Camera.CalibrationNumColumns)
            PatDisplayArray.Add("IDSData.Hardware.Camera.CalibrationNumRows")
            PatArray.Add(IDSData.Hardware.Camera.CalibrationNumRows)
            PatDisplayArray.Add("IDSData.Hardware.Camera.CalibrationSpacing")
            PatArray.Add(IDSData.Hardware.Camera.CalibrationSpacing)
            PatDisplayArray.Add("IDSData.Hardware.Camera.CalibrationBlockPos.X")
            PatArray.Add(IDSData.Hardware.Camera.CalibrationBlockPos.X)
            PatDisplayArray.Add("IDSData.Hardware.Camera.CalibrationBlockPos.Y")
            PatArray.Add(IDSData.Hardware.Camera.CalibrationBlockPos.Y)
            PatDisplayArray.Add("IDSData.Hardware.Camera.CalibrationBlockPos.Z")
            PatArray.Add(IDSData.Hardware.Camera.CalibrationBlockPos.Z)
            PatDisplayArray.Add("IDSData.Hardware.Camera.BlockSize.X")
            PatArray.Add(IDSData.Hardware.Camera.BlockSize.X)
            PatDisplayArray.Add("IDSData.Hardware.Camera.BlockSize.Y")
            PatArray.Add(IDSData.Hardware.Camera.BlockSize.Y)
            PatDisplayArray.Add("IDSData.Hardware.Camera.BlockSizeTolerance")
            PatArray.Add(IDSData.Hardware.Camera.BlockSizeTolerance)
            PatDisplayArray.Add("IDSData.Hardware.Camera.BlockPosTolerance")
            PatArray.Add(IDSData.Hardware.Camera.BlockPosTolerance)
            PatDisplayArray.Add("IDSData.Hardware.Camera.BlockRotationTolerance")
            PatArray.Add(IDSData.Hardware.Camera.BlockRotationTolerance)
            PatDisplayArray.Add("IDSData.Hardware.Camera.ImageFindDir")
            PatArray.Add(IDSData.Hardware.Camera.ImageFindDir)
            PatDisplayArray.Add("IDSData.Hardware.Camera.ImageVertical")
            PatArray.Add(IDSData.Hardware.Camera.ImageVertical)
            PatDisplayArray.Add("IDSData.Hardware.Camera.Threshold")
            PatArray.Add(IDSData.Hardware.Camera.Threshold)
            PatDisplayArray.Add("IDSData.Hardware.Camera.Brightness")
            PatArray.Add(IDSData.Hardware.Camera.Brightness)
            PatDisplayArray.Add("IDSData.Hardware.Camera.Contrast")
            PatArray.Add(IDSData.Hardware.Camera.Contrast)
            PatDisplayArray.Add("IDSData.Hardware.Camera.ROISize")
            PatArray.Add(IDSData.Hardware.Camera.ROISize)
            PatDisplayArray.Add("IDSData.Hardware.Camera.ResultSize.X")
            PatArray.Add(IDSData.Hardware.Camera.ResultSize.X)
            PatDisplayArray.Add("IDSData.Hardware.Camera.ResultSize.Y")
            PatArray.Add(IDSData.Hardware.Camera.ResultSize.Y)
            PatDisplayArray.Add("IDSData.Hardware.Camera.ResultDia")
            PatArray.Add(IDSData.Hardware.Camera.ResultDia)
            PatDisplayArray.Add("IDSData.Hardware.Camera.ResultRotation")
            PatArray.Add(IDSData.Hardware.Camera.ResultRotation)

            PatDisplayArray.Add("IDSData.Hardware.Camera.DotQCEnable")
            PatArray.Add(IDSData.Hardware.Camera.DotQCEnable)

            ''''''''''' Conveyor '''''''''''''''''
            PatDisplayArray.Add("")
            PatArray.Add("")
            PatDisplayArray.Add("~")
            PatArray.Add("[CONVEYOR]")
            PatDisplayArray.Add("IDSData.Hardware.Conveyor.Width")
            PatArray.Add(IDSData.Hardware.Conveyor.Width)
            PatDisplayArray.Add("IDSData.Hardware.Conveyor.FullWidth")
            PatArray.Add(IDSData.Hardware.Conveyor.FullWidth)
            PatDisplayArray.Add("IDSData.Hardware.Conveyor.MinWidth")
            PatArray.Add(IDSData.Hardware.Conveyor.MinWidth)
            PatDisplayArray.Add("IDSData.Hardware.Conveyor.Speed")
            PatArray.Add(IDSData.Hardware.Conveyor.Speed)
            PatDisplayArray.Add("IDSData.Hardware.Conveyor.TimeOut")
            PatArray.Add(IDSData.Hardware.Conveyor.TimeOut)
            PatDisplayArray.Add("IDSData.Hardware.Conveyor.WidthMoveStep")
            PatArray.Add(IDSData.Hardware.Conveyor.WidthMoveStep)

            PatDisplayArray.Add("IDSData.Hardware.Conveyor.upstreamTimeout")
            PatArray.Add(IDSData.Hardware.Conveyor.upstreamTimeout)
            PatDisplayArray.Add("IDSData.Hardware.Conveyor.downstreamTimeout")
            PatArray.Add(IDSData.Hardware.Conveyor.downstreamTimeout)

            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            ''''''''''' Dispenser '''''''''''''''''
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            PatDisplayArray.Add("")
            PatArray.Add("")
            PatDisplayArray.Add("~")
            PatArray.Add("[DISPENSER]")

            PatDisplayArray.Add("IDSData.Hardware.Dispenser.CurrentHeads")
            PatArray.Add(IDSData.Hardware.Dispenser.CurrentHeads)

            PatDisplayArray.Add("")
            PatArray.Add("")
            PatDisplayArray.Add("~")
            PatArray.Add("[DISPENSER@LEFT]")
            PatDisplayArray.Add("IDSData.Hardware.Dispenser.Left.MaterialAirPressure")
            PatArray.Add(IDSData.Hardware.Dispenser.Left.MaterialAirPressure)
            PatDisplayArray.Add("IDSData.Hardware.Dispenser.Left.SuckbackPressure")
            PatArray.Add(IDSData.Hardware.Dispenser.Left.SuckbackPressure)
            PatDisplayArray.Add("IDSData.Hardware.Dispenser.Left.RPM")
            PatArray.Add(IDSData.Hardware.Dispenser.Left.RPM)
            PatDisplayArray.Add("IDSData.Hardware.Dispenser.Left.RetractTime")
            PatArray.Add(IDSData.Hardware.Dispenser.Left.RetractTime)
            PatDisplayArray.Add("IDSData.Hardware.Dispenser.Left.RetractDelay")
            PatArray.Add(IDSData.Hardware.Dispenser.Left.RetractDelay)
            PatDisplayArray.Add("IDSData.Hardware.Dispenser.Left.Pulse")
            PatArray.Add(IDSData.Hardware.Dispenser.Left.Pulse)
            PatDisplayArray.Add("IDSData.Hardware.Dispenser.Left.Pause")
            PatArray.Add(IDSData.Hardware.Dispenser.Left.Pause)
            PatDisplayArray.Add("IDSData.Hardware.Dispenser.Left.Count")
            PatArray.Add(IDSData.Hardware.Dispenser.Left.Count)
            PatDisplayArray.Add("IDSData.Hardware.Dispenser.Left.ValveTemperature")
            PatArray.Add(IDSData.Hardware.Dispenser.Left.ValveTemperature)
            PatDisplayArray.Add("IDSData.Hardware.Dispenser.Left.NeedleTipLength")
            PatArray.Add(IDSData.Hardware.Dispenser.Left.NeedleTipLength)
            PatDisplayArray.Add("IDSData.Hardware.Dispenser.Left.NeedleColor")
            PatArray.Add(IDSData.Hardware.Dispenser.Left.NeedleColor)
            PatDisplayArray.Add("IDSData.Hardware.Dispenser.Left.MaterialInfo")
            PatArray.Add(IDSData.Hardware.Dispenser.Left.MaterialInfo)
            PatDisplayArray.Add("IDSData.Hardware.Dispenser.Left.AutoPurgingDuration")
            PatArray.Add(IDSData.Hardware.Dispenser.Left.AutoPurgingDuration)
            PatDisplayArray.Add("IDSData.Hardware.Dispenser.Left.AutoCleaningDuration")
            PatArray.Add(IDSData.Hardware.Dispenser.Left.AutoCleaningDuration)
            PatDisplayArray.Add("IDSData.Hardware.Dispenser.Left.AutoPurgingInterval")
            PatArray.Add(IDSData.Hardware.Dispenser.Left.AutoPurgingInterval)
            PatDisplayArray.Add("IDSData.Hardware.Dispenser.Left.PotLifeDuration")
            PatArray.Add(IDSData.Hardware.Dispenser.Left.PotLifeDuration)
            PatDisplayArray.Add("IDSData.Hardware.Dispenser.Left.PotLifeOption")
            PatArray.Add(IDSData.Hardware.Dispenser.Left.PotLifeOption)
            PatDisplayArray.Add("IDSData.Hardware.Dispenser.Left.AutoPurgingOption")
            PatArray.Add(IDSData.Hardware.Dispenser.Left.AutoPurgingOption)
            PatDisplayArray.Add("IDSData.Hardware.Dispenser.Left.HeadType")
            PatArray.Add(IDSData.Hardware.Dispenser.Left.HeadType)

            PatDisplayArray.Add("IDSData.Hardware.Dispenser.Left.JettingCalOnce")
            PatArray.Add(IDSData.Hardware.Dispenser.Left.JettingCalOnce)
            PatDisplayArray.Add("IDSData.Hardware.Dispenser.Left.AugerCalOnce")
            PatArray.Add(IDSData.Hardware.Dispenser.Left.AugerCalOnce)
            PatDisplayArray.Add("IDSData.Hardware.Dispenser.Left.SlideValveCalOnce")
            PatArray.Add(IDSData.Hardware.Dispenser.Left.SlideValveCalOnce)
            PatDisplayArray.Add("IDSData.Hardware.Dispenser.Left.TimePressureValveCalOnce")
            PatArray.Add(IDSData.Hardware.Dispenser.Left.TimePressureValveCalOnce)
            PatDisplayArray.Add("IDSData.Hardware.Dispenser.Left.TimePressureSyringeCalOnce")
            PatArray.Add(IDSData.Hardware.Dispenser.Left.TimePressureSyringeCalOnce)

            PatDisplayArray.Add("")
            PatArray.Add("")
            PatDisplayArray.Add("~")
            PatArray.Add("[DISPENSER@RIGHT]")
            PatDisplayArray.Add("IDSData.Hardware.Dispenser.Right.MaterialAirPressure")
            PatArray.Add(IDSData.Hardware.Dispenser.Right.MaterialAirPressure)
            PatDisplayArray.Add("IDSData.Hardware.Dispenser.Right.SuckbackPressure")
            PatArray.Add(IDSData.Hardware.Dispenser.Right.SuckbackPressure)
            PatDisplayArray.Add("IDSData.Hardware.Dispenser.Right.RPM")
            PatArray.Add(IDSData.Hardware.Dispenser.Right.RPM)
            PatDisplayArray.Add("IDSData.Hardware.Dispenser.Right.RetractTime")
            PatArray.Add(IDSData.Hardware.Dispenser.Right.RetractTime)
            PatDisplayArray.Add("IDSData.Hardware.Dispenser.Right.RetractDelay")
            PatArray.Add(IDSData.Hardware.Dispenser.Right.RetractDelay)
            PatDisplayArray.Add("IDSData.Hardware.Dispenser.Right.Pulse")
            PatArray.Add(IDSData.Hardware.Dispenser.Right.Pulse)
            PatDisplayArray.Add("IDSData.Hardware.Dispenser.Right.Pause")
            PatArray.Add(IDSData.Hardware.Dispenser.Right.Pause)
            PatDisplayArray.Add("IDSData.Hardware.Dispenser.Right.Count")
            PatArray.Add(IDSData.Hardware.Dispenser.Right.Count)
            PatDisplayArray.Add("IDSData.Hardware.Dispenser.Right.ValveTemperature")
            PatArray.Add(IDSData.Hardware.Dispenser.Right.ValveTemperature)
            PatDisplayArray.Add("IDSData.Hardware.Dispenser.Right.NeedleTipLength")
            PatArray.Add(IDSData.Hardware.Dispenser.Right.NeedleTipLength)
            PatDisplayArray.Add("IDSData.Hardware.Dispenser.Right.NeedleColor")
            PatArray.Add(IDSData.Hardware.Dispenser.Right.NeedleColor)
            PatDisplayArray.Add("IDSData.Hardware.Dispenser.Right.MaterialInfo")
            PatArray.Add(IDSData.Hardware.Dispenser.Right.MaterialInfo)
            PatDisplayArray.Add("IDSData.Hardware.Dispenser.Right.AutoPurgingDuration")
            PatArray.Add(IDSData.Hardware.Dispenser.Right.AutoPurgingDuration)
            PatDisplayArray.Add("IDSData.Hardware.Dispenser.Right.AutoCleaningDuration")
            PatArray.Add(IDSData.Hardware.Dispenser.Right.AutoCleaningDuration)
            PatDisplayArray.Add("IDSData.Hardware.Dispenser.Right.AutoPurgingInterval")
            PatArray.Add(IDSData.Hardware.Dispenser.Right.AutoPurgingInterval)
            PatDisplayArray.Add("IDSData.Hardware.Dispenser.Right.PotLifeDuration")
            PatArray.Add(IDSData.Hardware.Dispenser.Right.PotLifeDuration)
            PatDisplayArray.Add("IDSData.Hardware.Dispenser.Right.PotLifeOption")
            PatArray.Add(IDSData.Hardware.Dispenser.Right.PotLifeOption)
            PatDisplayArray.Add("IDSData.Hardware.Dispenser.Right.AutoPurgingOption")
            PatArray.Add(IDSData.Hardware.Dispenser.Right.AutoPurgingOption)
            PatDisplayArray.Add("IDSData.Hardware.Dispenser.Right.HeadType")
            PatArray.Add(IDSData.Hardware.Dispenser.Right.HeadType)

            PatDisplayArray.Add("IDSData.Hardware.Dispenser.Right.JettingCalOnce")
            PatArray.Add(IDSData.Hardware.Dispenser.Right.JettingCalOnce)
            PatDisplayArray.Add("IDSData.Hardware.Dispenser.Right.AugerCalOnce")
            PatArray.Add(IDSData.Hardware.Dispenser.Right.AugerCalOnce)
            PatDisplayArray.Add("IDSData.Hardware.Dispenser.Right.SlideValveCalOnce")
            PatArray.Add(IDSData.Hardware.Dispenser.Right.SlideValveCalOnce)
            PatDisplayArray.Add("IDSData.Hardware.Dispenser.Right.TimePressureValveCalOnce")
            PatArray.Add(IDSData.Hardware.Dispenser.Right.TimePressureValveCalOnce)
            PatDisplayArray.Add("IDSData.Hardware.Dispenser.Right.TimePressureSyringeCalOnce")
            PatArray.Add(IDSData.Hardware.Dispenser.Right.TimePressureSyringeCalOnce)

            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            ''''''''''' Gantry '''''''''''''''''
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            PatDisplayArray.Add("")
            PatArray.Add("")
            PatDisplayArray.Add("~")
            PatArray.Add("[GANTRY]")
            'If MyFileName.IndexOf("FactoryDefault") > 0 Then
            PatDisplayArray.Add("IDSData.Hardware.Gantry.WeighingScalePosition.X")
            PatArray.Add(IDSData.Hardware.Gantry.WeighingScalePosition.X)
            PatDisplayArray.Add("IDSData.Hardware.Gantry.WeighingScalePosition.Y")
            PatArray.Add(IDSData.Hardware.Gantry.WeighingScalePosition.Y)
            PatDisplayArray.Add("IDSData.Hardware.Gantry.WeighingScalePosition.Z")
            PatArray.Add(IDSData.Hardware.Gantry.WeighingScalePosition.Z)

            PatDisplayArray.Add("IDSData.Hardware.Gantry.CleanPosition.X")
            PatArray.Add(IDSData.Hardware.Gantry.CleanPosition.X)
            PatDisplayArray.Add("IDSData.Hardware.Gantry.CleanPosition.Y")
            PatArray.Add(IDSData.Hardware.Gantry.CleanPosition.Y)
            PatDisplayArray.Add("IDSData.Hardware.Gantry.CleanPosition.Z")
            PatArray.Add(IDSData.Hardware.Gantry.CleanPosition.Z)
            'End If


            PatDisplayArray.Add("IDSData.Hardware.Gantry.ElementZSpeed")
            PatArray.Add(IDSData.Hardware.Gantry.ElementZSpeed)
            PatDisplayArray.Add("IDSData.Hardware.Gantry.ElementXYSpeed")
            PatArray.Add(IDSData.Hardware.Gantry.ElementXYSpeed)
            PatDisplayArray.Add("IDSData.Hardware.Gantry.GraphicDisplayRatio")
            PatArray.Add(IDSData.Hardware.Gantry.GraphicDisplayRatio)
            PatDisplayArray.Add("IDSData.Hardware.Gantry.GraphicDisplayXYLT.X")
            PatArray.Add(IDSData.Hardware.Gantry.GraphicDisplayXYLT.X)
            PatDisplayArray.Add("IDSData.Hardware.Gantry.GraphicDisplayXYLT.Y")
            PatArray.Add(IDSData.Hardware.Gantry.GraphicDisplayXYLT.Y)
            PatDisplayArray.Add("IDSData.Hardware.Gantry.GraphicDisplayXYRB.X")
            PatArray.Add(IDSData.Hardware.Gantry.GraphicDisplayXYRB.X)
            PatDisplayArray.Add("IDSData.Hardware.Gantry.GraphicDisplayXYRB.Y")
            PatArray.Add(IDSData.Hardware.Gantry.GraphicDisplayXYRB.Y)
            PatDisplayArray.Add("IDSData.Hardware.Gantry.MaxAccelerationLimit")
            PatArray.Add(IDSData.Hardware.Gantry.MaxAccelerationLimit)
            PatDisplayArray.Add("IDSData.Hardware.Gantry.MaxSpeedLimit")
            PatArray.Add(IDSData.Hardware.Gantry.MaxSpeedLimit)

            'If MyFileName.IndexOf("FactoryDefault") > 0 Then
            PatDisplayArray.Add("IDSData.Hardware.Gantry.ParkPosition.X")
            PatArray.Add(IDSData.Hardware.Gantry.ParkPosition.X)
            PatDisplayArray.Add("IDSData.Hardware.Gantry.ParkPosition.Y")
            PatArray.Add(IDSData.Hardware.Gantry.ParkPosition.Y)
            PatDisplayArray.Add("IDSData.Hardware.Gantry.ParkPosition.Z")
            PatArray.Add(IDSData.Hardware.Gantry.ParkPosition.Z)
            PatDisplayArray.Add("IDSData.Hardware.Gantry.PurgePosition.X")
            PatArray.Add(IDSData.Hardware.Gantry.PurgePosition.X)
            PatDisplayArray.Add("IDSData.Hardware.Gantry.PurgePosition.Y")
            PatArray.Add(IDSData.Hardware.Gantry.PurgePosition.Y)
            PatDisplayArray.Add("IDSData.Hardware.Gantry.PurgePosition.Z")
            PatArray.Add(IDSData.Hardware.Gantry.PurgePosition.Z)
            'End If

            PatDisplayArray.Add("IDSData.Hardware.Gantry.SystemOriginPos.X")
            PatArray.Add(IDSData.Hardware.Gantry.SystemOriginPos.X)
            PatDisplayArray.Add("IDSData.Hardware.Gantry.SystemOriginPos.Y")
            PatArray.Add(IDSData.Hardware.Gantry.SystemOriginPos.Y)
            PatDisplayArray.Add("IDSData.Hardware.Gantry.SystemOriginPos.Z")
            PatArray.Add(IDSData.Hardware.Gantry.SystemOriginPos.Z)

            PatDisplayArray.Add("IDSData.Hardware.Gantry.SystemUnit")
            PatArray.Add(IDSData.Hardware.Gantry.SystemUnit)
            PatDisplayArray.Add("IDSData.Hardware.Gantry.WorkArea.X")
            PatArray.Add(IDSData.Hardware.Gantry.WorkArea.X)
            PatDisplayArray.Add("IDSData.Hardware.Gantry.WorkArea.Y")
            PatArray.Add(IDSData.Hardware.Gantry.WorkArea.Y)
            PatDisplayArray.Add("IDSData.Hardware.Gantry.WorkArea.Z.Max")
            PatArray.Add(IDSData.Hardware.Gantry.WorkArea.Z.Max)
            PatDisplayArray.Add("IDSData.Hardware.Gantry.WorkArea.Z.Min")
            PatArray.Add(IDSData.Hardware.Gantry.WorkArea.Z.Min)
            PatDisplayArray.Add("IDSData.Hardware.Gantry.ServiceXYSpeed")
            PatArray.Add(IDSData.Hardware.Gantry.ServiceXYSpeed)
            PatDisplayArray.Add("IDSData.Hardware.Gantry.ServiceZSpeed")
            PatArray.Add(IDSData.Hardware.Gantry.ServiceZSpeed)

            'If MyFileName.IndexOf("FactoryDefault") > 0 Then
            PatDisplayArray.Add("IDSData.Hardware.Gantry.ChangeSyringePosition.X")
            PatArray.Add(IDSData.Hardware.Gantry.ChangeSyringePosition.X)
            PatDisplayArray.Add("IDSData.Hardware.Gantry.ChangeSyringePosition.Y")
            PatArray.Add(IDSData.Hardware.Gantry.ChangeSyringePosition.Y)
            PatDisplayArray.Add("IDSData.Hardware.Gantry.ChangeSyringePosition.Z")
            PatArray.Add(IDSData.Hardware.Gantry.ChangeSyringePosition.Z)

            PatDisplayArray.Add("IDSData.Hardware.Gantry.WeighingScaleBottomRight.X")
            PatArray.Add(IDSData.Hardware.Gantry.WeighingScaleBottomRight.X)
            PatDisplayArray.Add("IDSData.Hardware.Gantry.WeighingScaleBottomRight.Y")
            PatArray.Add(IDSData.Hardware.Gantry.WeighingScaleBottomRight.Y)
            'End If

            'PatDisplayArray.Add("IDSData.Hardware.Gantry.WeighingScaleBottomRight.Z")
            'PatArray.Add(IDSData.Hardware.Gantry.WeighingScaleBottomRight.Z)

            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            ''''''''''' HEIGHTSENOR'''''''''''''''''
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            PatDisplayArray.Add("")
            PatArray.Add("")
            PatDisplayArray.Add("~")
            PatArray.Add("[HEIGHTSENOR]")
            PatDisplayArray.Add("IDSData.Hardware.HeightSensor.BoardHeight")
            PatArray.Add(IDSData.Hardware.HeightSensor.BoardHeight)
            PatDisplayArray.Add("IDSData.Hardware.HeightSensor.CalibratorHeightPos.X")
            PatArray.Add(IDSData.Hardware.HeightSensor.CalibratorHeightPos.X)
            PatDisplayArray.Add("IDSData.Hardware.HeightSensor.CalibratorHeightPos.Y")
            PatArray.Add(IDSData.Hardware.HeightSensor.CalibratorHeightPos.Y)
            PatDisplayArray.Add("IDSData.Hardware.HeightSensor.SelectType")
            PatArray.Add(IDSData.Hardware.HeightSensor.SelectType)

            PatDisplayArray.Add("")
            PatArray.Add("")
            PatDisplayArray.Add("~")
            PatArray.Add("[HEIGHTSENOR@LASER]")
            PatDisplayArray.Add("IDSData.Hardware.HeightSensor.Laser.CurrentPos.X")
            PatArray.Add(IDSData.Hardware.HeightSensor.Laser.CurrentPos.X)
            PatDisplayArray.Add("IDSData.Hardware.HeightSensor.Laser.CurrentPos.Y")
            PatArray.Add(IDSData.Hardware.HeightSensor.Laser.CurrentPos.Y)
            PatDisplayArray.Add("IDSData.Hardware.HeightSensor.Laser.CurrentPos.Z")
            PatArray.Add(IDSData.Hardware.HeightSensor.Laser.CurrentPos.Z)
            PatDisplayArray.Add("IDSData.Hardware.HeightSensor.Laser.HeightReference")
            PatArray.Add(IDSData.Hardware.HeightSensor.Laser.HeightReference)
            PatDisplayArray.Add("IDSData.Hardware.HeightSensor.Laser.OffsetPos.X")
            PatArray.Add(IDSData.Hardware.HeightSensor.Laser.OffsetPos.X)
            PatDisplayArray.Add("IDSData.Hardware.HeightSensor.Laser.OffsetPos.Y")
            PatArray.Add(IDSData.Hardware.HeightSensor.Laser.OffsetPos.Y)

            PatDisplayArray.Add("")
            PatArray.Add("")
            PatDisplayArray.Add("~")
            PatArray.Add("[HEIGHTSENOR@TOUCHPROBE]")
            PatDisplayArray.Add("IDSData.Hardware.HeightSensor.TP.CurrentPos.X")
            PatArray.Add(IDSData.Hardware.HeightSensor.TP.CurrentPos.X)
            PatDisplayArray.Add("IDSData.Hardware.HeightSensor.TP.CurrentPos.Y")
            PatArray.Add(IDSData.Hardware.HeightSensor.TP.CurrentPos.Y)
            PatDisplayArray.Add("IDSData.Hardware.HeightSensor.TP.CurrentPos.Z")
            PatArray.Add(IDSData.Hardware.HeightSensor.TP.CurrentPos.Z)
            PatDisplayArray.Add("IDSData.Hardware.HeightSensor.TP.HeightReference")
            PatArray.Add(IDSData.Hardware.HeightSensor.TP.HeightReference)
            PatDisplayArray.Add("IDSData.Hardware.HeightSensor.TP.OffsetPos.X")
            PatArray.Add(IDSData.Hardware.HeightSensor.TP.OffsetPos.X)
            PatDisplayArray.Add("IDSData.Hardware.HeightSensor.TP.OffsetPos.Y")
            PatArray.Add(IDSData.Hardware.HeightSensor.TP.OffsetPos.Y)


            PatDisplayArray.Add("")
            PatArray.Add("")
            PatDisplayArray.Add("~")
            PatArray.Add("[HEIGHTSENOR@LVDT]")
            PatDisplayArray.Add("IDSData.Hardware.HeightSensor.LVDTCalBackGround")
            PatArray.Add(IDSData.Hardware.HeightSensor.LVDTCalBackGround)
            PatDisplayArray.Add("IDSData.Hardware.HeightSensor.LVDTCalBrightness")
            PatArray.Add(IDSData.Hardware.HeightSensor.LVDTCalBrightness)
            PatDisplayArray.Add("IDSData.Hardware.HeightSensor.LVDTCalThreshold")
            PatArray.Add(IDSData.Hardware.HeightSensor.LVDTCalThreshold)
            PatDisplayArray.Add("IDSData.Hardware.HeightSensor.LVDTCalOpen")
            PatArray.Add(IDSData.Hardware.HeightSensor.LVDTCalOpen)
            PatDisplayArray.Add("IDSData.Hardware.HeightSensor.LVDTCalClose")
            PatArray.Add(IDSData.Hardware.HeightSensor.LVDTCalClose)
            PatDisplayArray.Add("IDSData.Hardware.HeightSensor.LVDTCalMaxRadius")
            PatArray.Add(IDSData.Hardware.HeightSensor.LVDTCalMaxRadius)
            PatDisplayArray.Add("IDSData.Hardware.HeightSensor.LVDTCalMinRadius")
            PatArray.Add(IDSData.Hardware.HeightSensor.LVDTCalMinRadius)
            PatDisplayArray.Add("IDSData.Hardware.HeightSensor.LVDTCalRoughness")
            PatArray.Add(IDSData.Hardware.HeightSensor.LVDTCalRoughness)
            PatDisplayArray.Add("IDSData.Hardware.HeightSensor.LVDTCalCompactness")
            PatArray.Add(IDSData.Hardware.HeightSensor.LVDTCalCompactness)

            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            ''''''''''' Lightbox ''''''''''''''''
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            PatDisplayArray.Add("")
            PatArray.Add("")
            PatDisplayArray.Add("~")
            PatArray.Add("[LIGHTBOX]")
            PatDisplayArray.Add("IDSData.Hardware.LightBox.Brightness")
            PatArray.Add(IDSData.Hardware.LightBox.Brightness)

            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            ''''''''''' needle ''''''''''''''''
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            PatDisplayArray.Add("")
            PatArray.Add("")
            PatDisplayArray.Add("~")
            PatArray.Add("[NEEDLE@LEFT]")
            PatDisplayArray.Add("IDSData.Hardware.Needle.Left.ArcRadius")
            PatArray.Add(IDSData.Hardware.Needle.Left.ArcRadius)
            PatDisplayArray.Add("IDSData.Hardware.Needle.Left.CalibratorPos.X")
            PatArray.Add(IDSData.Hardware.Needle.Left.CalibratorPos.X)
            PatDisplayArray.Add("IDSData.Hardware.Needle.Left.CalibratorPos.Y")
            PatArray.Add(IDSData.Hardware.Needle.Left.CalibratorPos.Y)
            PatDisplayArray.Add("IDSData.Hardware.Needle.Left.CalibratorPos.Z")
            PatArray.Add(IDSData.Hardware.Needle.Left.CalibratorPos.Z)

            PatDisplayArray.Add("IDSData.Hardware.Needle.Left.HeightApproach")
            PatArray.Add(IDSData.Hardware.Needle.Left.HeightApproach)
            PatDisplayArray.Add("IDSData.Hardware.Needle.Left.HeightClearance")
            PatArray.Add(IDSData.Hardware.Needle.Left.HeightClearance)
            PatDisplayArray.Add("IDSData.Hardware.Needle.Left.HeightRetract")
            PatArray.Add(IDSData.Hardware.Needle.Left.HeightRetract)
            PatDisplayArray.Add("IDSData.Hardware.Needle.Left.NeedleCalibrationPosition.X")
            PatArray.Add(IDSData.Hardware.Needle.Left.NeedleCalibrationPosition.X)
            PatDisplayArray.Add("IDSData.Hardware.Needle.Left.NeedleCalibrationPosition.Y")
            PatArray.Add(IDSData.Hardware.Needle.Left.NeedleCalibrationPosition.Y)
            PatDisplayArray.Add("IDSData.Hardware.Needle.Left.NeedleCalibrationPosition.Z")
            PatArray.Add(IDSData.Hardware.Needle.Left.NeedleCalibrationPosition.Z)
            PatDisplayArray.Add("IDSData.Hardware.Needle.Left.TouchSensorZPosition")
            PatArray.Add(IDSData.Hardware.Needle.Left.TouchSensorZPosition)

            PatDisplayArray.Add("IDSData.Hardware.Needle.Left.ArrayDotPos1.X")
            PatArray.Add(IDSData.Hardware.Needle.Left.ArrayDotPos1.X)
            PatDisplayArray.Add("IDSData.Hardware.Needle.Left.ArrayDotPos1.Y")
            PatArray.Add(IDSData.Hardware.Needle.Left.ArrayDotPos1.Y)
            PatDisplayArray.Add("IDSData.Hardware.Needle.Left.ArrayDotPos1.Z")
            PatArray.Add(IDSData.Hardware.Needle.Left.ArrayDotPos1.Z)

            PatDisplayArray.Add("IDSData.Hardware.Needle.Left.ArrayDotPos3.X")
            PatArray.Add(IDSData.Hardware.Needle.Left.ArrayDotPos3.X)
            PatDisplayArray.Add("IDSData.Hardware.Needle.Left.ArrayDotPos3.Y")
            PatArray.Add(IDSData.Hardware.Needle.Left.ArrayDotPos3.Y)
            PatDisplayArray.Add("IDSData.Hardware.Needle.Left.ArrayDotPos3.Z")
            PatArray.Add(IDSData.Hardware.Needle.Left.ArrayDotPos3.Z)

            PatDisplayArray.Add("IDSData.Hardware.Needle.Left.CalBackground")
            PatArray.Add(IDSData.Hardware.Needle.Left.CalBackground)
            PatDisplayArray.Add("IDSData.Hardware.Needle.Left.CalBrightness")
            PatArray.Add(IDSData.Hardware.Needle.Left.CalBrightness)
            PatDisplayArray.Add("IDSData.Hardware.Needle.Left.CalThreshold")
            PatArray.Add(IDSData.Hardware.Needle.Left.CalThreshold)
            PatDisplayArray.Add("IDSData.Hardware.Needle.Left.CalOpen")
            PatArray.Add(IDSData.Hardware.Needle.Left.CalOpen)
            PatDisplayArray.Add("IDSData.Hardware.Needle.Left.CalClose")
            PatArray.Add(IDSData.Hardware.Needle.Left.CalClose)
            PatDisplayArray.Add("IDSData.Hardware.Needle.Left.CalMaxRadius")
            PatArray.Add(IDSData.Hardware.Needle.Left.CalMaxRadius)
            PatDisplayArray.Add("IDSData.Hardware.Needle.Left.CalMinRadius")
            PatArray.Add(IDSData.Hardware.Needle.Left.CalMinRadius)
            PatDisplayArray.Add("IDSData.Hardware.Needle.Left.CalRoughness")
            PatArray.Add(IDSData.Hardware.Needle.Left.CalRoughness)
            PatDisplayArray.Add("IDSData.Hardware.Needle.Left.CalCompactness")
            PatArray.Add(IDSData.Hardware.Needle.Left.CalCompactness)

            PatDisplayArray.Add("IDSData.Hardware.Needle.Left.DispenseDot.ApproachHeight")
            PatArray.Add(IDSData.Hardware.Needle.Left.DispenseDot.ApproachHeight)
            PatDisplayArray.Add("IDSData.Hardware.Needle.Left.DispenseDot.ArcRadius")
            PatArray.Add(IDSData.Hardware.Needle.Left.DispenseDot.ArcRadius)
            PatDisplayArray.Add("IDSData.Hardware.Needle.Left.DispenseDot.ClearanceHeight")
            PatArray.Add(IDSData.Hardware.Needle.Left.DispenseDot.ClearanceHeight)
            PatDisplayArray.Add("IDSData.Hardware.Needle.Left.DispenseDot.DispenseDuration")
            PatArray.Add(IDSData.Hardware.Needle.Left.DispenseDot.DispenseDuration)
            PatDisplayArray.Add("IDSData.Hardware.Needle.Left.DispenseDot.NeedleGap")
            PatArray.Add(IDSData.Hardware.Needle.Left.DispenseDot.NeedleGap)
            PatDisplayArray.Add("IDSData.Hardware.Needle.Left.DispenseDot.RetractDelay")
            PatArray.Add(IDSData.Hardware.Needle.Left.DispenseDot.RetractDelay)
            PatDisplayArray.Add("IDSData.Hardware.Needle.Left.DispenseDot.RetractHeight")
            PatArray.Add(IDSData.Hardware.Needle.Left.DispenseDot.RetractHeight)
            PatDisplayArray.Add("IDSData.Hardware.Needle.Left.DispenseDot.RetractSpeed")
            PatArray.Add(IDSData.Hardware.Needle.Left.DispenseDot.RetractSpeed)

            PatDisplayArray.Add("IDSData.Hardware.Needle.Left.DotDiameter")
            PatArray.Add(IDSData.Hardware.Needle.Left.DotDiameter)


            PatDisplayArray.Add("")
            PatArray.Add("")
            PatDisplayArray.Add("~")
            PatArray.Add("[NEEDLE@RIGHT]")
            PatDisplayArray.Add("IDSData.Hardware.Needle.Right.ArcRadius")
            PatArray.Add(IDSData.Hardware.Needle.Right.ArcRadius)
            PatDisplayArray.Add("IDSData.Hardware.Needle.Right.CalibratorPos.X")
            PatArray.Add(IDSData.Hardware.Needle.Right.CalibratorPos.X)
            PatDisplayArray.Add("IDSData.Hardware.Needle.Right.CalibratorPos.Y")
            PatArray.Add(IDSData.Hardware.Needle.Right.CalibratorPos.Y)
            PatDisplayArray.Add("IDSData.Hardware.Needle.Right.CalibratorPos.Z")
            PatArray.Add(IDSData.Hardware.Needle.Right.CalibratorPos.Z)

            PatDisplayArray.Add("IDSData.Hardware.Needle.Right.HeightApproach")
            PatArray.Add(IDSData.Hardware.Needle.Right.HeightApproach)
            PatDisplayArray.Add("IDSData.Hardware.Needle.Right.HeightClearance")
            PatArray.Add(IDSData.Hardware.Needle.Right.HeightClearance)
            PatDisplayArray.Add("IDSData.Hardware.Needle.Right.HeightRetract")
            PatArray.Add(IDSData.Hardware.Needle.Right.HeightRetract)
            PatDisplayArray.Add("IDSData.Hardware.Needle.Right.NeedleCalibrationPosition.X")
            PatArray.Add(IDSData.Hardware.Needle.Right.NeedleCalibrationPosition.X)
            PatDisplayArray.Add("IDSData.Hardware.Needle.Right.NeedleCalibrationPosition.Y")
            PatArray.Add(IDSData.Hardware.Needle.Right.NeedleCalibrationPosition.Y)
            PatDisplayArray.Add("IDSData.Hardware.Needle.Right.NeedleCalibrationPosition.Z")
            PatArray.Add(IDSData.Hardware.Needle.Right.NeedleCalibrationPosition.Z)
            PatDisplayArray.Add("IDSData.Hardware.Needle.Right.TouchSensorZPosition")
            PatArray.Add(IDSData.Hardware.Needle.Right.TouchSensorZPosition)


            PatDisplayArray.Add("IDSData.Hardware.Needle.Right.ArrayDotPos1.X")
            PatArray.Add(IDSData.Hardware.Needle.Right.ArrayDotPos1.X)
            PatDisplayArray.Add("IDSData.Hardware.Needle.Right.ArrayDotPos1.Y")
            PatArray.Add(IDSData.Hardware.Needle.Right.ArrayDotPos1.Y)
            PatDisplayArray.Add("IDSData.Hardware.Needle.Right.ArrayDotPos1.Z")
            PatArray.Add(IDSData.Hardware.Needle.Right.ArrayDotPos1.Z)

            PatDisplayArray.Add("IDSData.Hardware.Needle.Right.ArrayDotPos3.X")
            PatArray.Add(IDSData.Hardware.Needle.Right.ArrayDotPos3.X)
            PatDisplayArray.Add("IDSData.Hardware.Needle.Right.ArrayDotPos3.Y")
            PatArray.Add(IDSData.Hardware.Needle.Right.ArrayDotPos3.Y)
            PatDisplayArray.Add("IDSData.Hardware.Needle.Right.ArrayDotPos3.Z")
            PatArray.Add(IDSData.Hardware.Needle.Right.ArrayDotPos3.Z)

            PatDisplayArray.Add("IDSData.Hardware.Needle.Right.CalBackground")
            PatArray.Add(IDSData.Hardware.Needle.Right.CalBackground)
            PatDisplayArray.Add("IDSData.Hardware.Needle.Right.CalBrightness")
            PatArray.Add(IDSData.Hardware.Needle.Right.CalBrightness)
            PatDisplayArray.Add("IDSData.Hardware.Needle.RightCalThreshold")
            PatArray.Add(IDSData.Hardware.Needle.Right.CalThreshold)
            PatDisplayArray.Add("IDSData.Hardware.Needle.Right.CalOpen")
            PatArray.Add(IDSData.Hardware.Needle.Right.CalOpen)
            PatDisplayArray.Add("IDSData.Hardware.Needle.Right.CalClose")
            PatArray.Add(IDSData.Hardware.Needle.Right.CalClose)
            PatDisplayArray.Add("IDSData.Hardware.Needle.Right.CalMaxRadius")
            PatArray.Add(IDSData.Hardware.Needle.Right.CalMaxRadius)
            PatDisplayArray.Add("IDSData.Hardware.Needle.Right.CalMinRadius")
            PatArray.Add(IDSData.Hardware.Needle.Right.CalMinRadius)
            PatDisplayArray.Add("IDSData.Hardware.Needle.Right.CalRoughness")
            PatArray.Add(IDSData.Hardware.Needle.Right.CalRoughness)
            PatDisplayArray.Add("IDSData.Hardware.Needle.Right.CalCompactness")
            PatArray.Add(IDSData.Hardware.Needle.Right.CalCompactness)

            PatDisplayArray.Add("IDSData.Hardware.Needle.Right.DispenseDot.ApproachHeight")
            PatArray.Add(IDSData.Hardware.Needle.Right.DispenseDot.ApproachHeight)
            PatDisplayArray.Add("IDSData.Hardware.Needle.Right.DispenseDot.ArcRadius")
            PatArray.Add(IDSData.Hardware.Needle.Right.DispenseDot.ArcRadius)
            PatDisplayArray.Add("IDSData.Hardware.Needle.Right.DispenseDot.ClearanceHeight")
            PatArray.Add(IDSData.Hardware.Needle.Right.DispenseDot.ClearanceHeight)
            PatDisplayArray.Add("IDSData.Hardware.Needle.Right.DispenseDot.DispenseDuration")
            PatArray.Add(IDSData.Hardware.Needle.Right.DispenseDot.DispenseDuration)
            PatDisplayArray.Add("IDSData.Hardware.Needle.Right.DispenseDot.NeedleGap")
            PatArray.Add(IDSData.Hardware.Needle.Right.DispenseDot.NeedleGap)
            PatDisplayArray.Add("IDSData.Hardware.Needle.Right.DispenseDot.RetractDelay")
            PatArray.Add(IDSData.Hardware.Needle.Right.DispenseDot.RetractDelay)
            PatDisplayArray.Add("IDSData.Hardware.Needle.Right.DispenseDot.RetractHeight")
            PatArray.Add(IDSData.Hardware.Needle.Right.DispenseDot.RetractHeight)
            PatDisplayArray.Add("IDSData.Hardware.Needle.Right.DispenseDot.RetractSpeed")
            PatArray.Add(IDSData.Hardware.Needle.Right.DispenseDot.RetractSpeed)

            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            ''''''''''' Regulator ''''''''''''''''
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            PatDisplayArray.Add("")
            PatArray.Add("")
            PatDisplayArray.Add("~")
            PatArray.Add("[REGULATOR]")
            PatDisplayArray.Add("IDSData.Hardware.Regulator.Left.Pressure")
            PatArray.Add(IDSData.Hardware.Regulator.Left.Pressure)
            PatDisplayArray.Add("IDSData.Hardware.Regulator.Left.Selection")
            PatArray.Add(IDSData.Hardware.Regulator.Left.Selection)
            PatDisplayArray.Add("IDSData.Hardware.Regulator.Left.Vacuum")
            PatArray.Add(IDSData.Hardware.Regulator.Left.Vacuum)

            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            ''''''''''' SPC ''''''''''''''''
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            PatDisplayArray.Add("")
            PatArray.Add("")
            PatDisplayArray.Add("~")
            PatArray.Add("[SPC]")
            PatDisplayArray.Add("IDSData.Hardware.SPC.ChipFailedAction")
            PatArray.Add(IDSData.Hardware.SPC.ChipFailedAction)
            PatDisplayArray.Add("IDSData.Hardware.SPC.CleanUpInterval")
            PatArray.Add(IDSData.Hardware.SPC.CleanUpInterval)
            PatDisplayArray.Add("IDSData.Hardware.SPC.EnableSPCReport")
            PatArray.Add(IDSData.Hardware.SPC.EnableSPCReport)
            PatDisplayArray.Add("IDSData.Hardware.SPC.FiducialFailedAction")
            PatArray.Add(IDSData.Hardware.SPC.FiducialFailedAction)
            PatDisplayArray.Add("IDSData.Hardware.SPC.HeightFailedAction")
            PatArray.Add(IDSData.Hardware.SPC.HeightFailedAction)
            PatDisplayArray.Add("IDSData.Hardware.SPC.QCFailedAction")
            PatArray.Add(IDSData.Hardware.SPC.QCFailedAction)
            PatDisplayArray.Add("IDSData.Hardware.SPC.ItemsToBeReported")
            PatArray.Add(IDSData.Hardware.SPC.ItemsToBeReported)
            PatDisplayArray.Add("IDSData.Hardware.SPC.ReportFileName")
            PatArray.Add(IDSData.Hardware.SPC.ReportFileName)

            PatDisplayArray.Add("IDSData.Hardware.SPC.ProgAuthorName")
            PatArray.Add(IDSData.Hardware.SPC.ProgAuthorName)
            PatDisplayArray.Add("IDSData.Hardware.SPC.ProgAuthorContact")
            PatArray.Add(IDSData.Hardware.SPC.ProgAuthorContact)
            PatDisplayArray.Add("IDSData.Hardware.SPC.ProductionNote")
            PatArray.Add(IDSData.Hardware.SPC.ProductionNote)

            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            ''''''''''' SystemIO ''''''''''''''''
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            PatDisplayArray.Add("")
            PatArray.Add("")
            PatDisplayArray.Add("~")
            PatArray.Add("[SYSTEMIO]")
            I = 0
            If IDSData.Hardware.SystemIO.ItemArray.Count > 0 Then
                For I = 0 To IDSData.Hardware.SystemIO.ItemArray.Count - 1
                    IDSData.Hardware.SystemIO.Template = CType(IDSData.Hardware.SystemIO.ItemArray(I), _
                                                                CIDSData.CIDSHardWare.CIDSSystemIO.CIDSIO)

                    PatDisplayArray.Add("(" + I.ToString + ")IDSData.Hardware.SystemIO.Template.CableName")
                    PatArray.Add(IDSData.Hardware.SystemIO.Template.CableName)
                    PatDisplayArray.Add("(" + I.ToString + ")IDSData.Hardware.SystemIO.Template.IO")
                    PatArray.Add(IDSData.Hardware.SystemIO.Template.IO)
                    PatDisplayArray.Add("(" + I.ToString + ")IDSData.Hardware.SystemIO.Template.IOName")
                    PatArray.Add(IDSData.Hardware.SystemIO.Template.IOName)
                    PatDisplayArray.Add("(" + I.ToString + ")IDSData.Hardware.SystemIO.Template.ModuleName")
                    PatArray.Add(IDSData.Hardware.SystemIO.Template.ModuleName)
                    PatDisplayArray.Add("(" + I.ToString + ")IDSData.Hardware.SystemIO.Template.Status")
                    PatArray.Add(IDSData.Hardware.SystemIO.Template.Status)
                    PatDisplayArray.Add("(" + I.ToString + ")IDSData.Hardware.SystemIO.Template.Type")
                    PatArray.Add(IDSData.Hardware.SystemIO.Template.Type)
                    PatDisplayArray.Add("")
                    PatArray.Add("")
                Next
            End If
            PatDisplayArray.Add("~")
            PatArray.Add("[SYSTEMIO END]")
            PatDisplayArray.Add("")
            PatArray.Add("")

            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            ''''''''''' Thermal ''''''''''''''''
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            PatDisplayArray.Add("")
            PatArray.Add("")
            PatDisplayArray.Add("~")
            PatArray.Add("[THERMAL]")
            PatDisplayArray.Add("IDSData.Hardware.Thermal.HeaterFeatureEnabled")
            PatArray.Add(IDSData.Hardware.Thermal.HeaterFeatureEnabled)
            PatDisplayArray.Add("IDSData.Hardware.Thermal.Needle.OperationTemp")
            PatArray.Add(IDSData.Hardware.Thermal.Needle.OperationTemp)
            PatDisplayArray.Add("IDSData.Hardware.Thermal.Needle.StandbyTemp")
            PatArray.Add(IDSData.Hardware.Thermal.Needle.StandbyTemp)
            PatDisplayArray.Add("IDSData.Hardware.Thermal.Needle.AlarmThreshold")
            PatArray.Add(IDSData.Hardware.Thermal.Needle.AlarmThreshold)
            PatDisplayArray.Add("IDSData.Hardware.Thermal.Needle.OnOff")
            PatArray.Add(IDSData.Hardware.Thermal.Needle.OnOff)

            PatDisplayArray.Add("IDSData.Hardware.Thermal.Syringe.OperationTemp")
            PatArray.Add(IDSData.Hardware.Thermal.Syringe.OperationTemp)
            PatDisplayArray.Add("IDSData.Hardware.Thermal.Syringe.StandbyTemp")
            PatArray.Add(IDSData.Hardware.Thermal.Syringe.StandbyTemp)
            PatDisplayArray.Add("IDSData.Hardware.Thermal.Syringe.AlarmThreshold")
            PatArray.Add(IDSData.Hardware.Thermal.Syringe.AlarmThreshold)
            PatDisplayArray.Add("IDSData.Hardware.Thermal.Syringe.OnOff")
            PatArray.Add(IDSData.Hardware.Thermal.Syringe.OnOff)

            PatDisplayArray.Add("IDSData.Hardware.Thermal.Station1.OperationTemp")
            PatArray.Add(IDSData.Hardware.Thermal.Station1.OperationTemp)
            PatDisplayArray.Add("IDSData.Hardware.Thermal.Station1.StandbyTemp")
            PatArray.Add(IDSData.Hardware.Thermal.Station1.StandbyTemp)
            PatDisplayArray.Add("IDSData.Hardware.Thermal.Station1.AlarmThreshold")
            PatArray.Add(IDSData.Hardware.Thermal.Station1.AlarmThreshold)
            PatDisplayArray.Add("IDSData.Hardware.Thermal.Station1.OnOff")
            PatArray.Add(IDSData.Hardware.Thermal.Station1.OnOff)

            PatDisplayArray.Add("IDSData.Hardware.Thermal.Station2.OperationTemp")
            PatArray.Add(IDSData.Hardware.Thermal.Station2.OperationTemp)
            PatDisplayArray.Add("IDSData.Hardware.Thermal.Station2.StandbyTemp")
            PatArray.Add(IDSData.Hardware.Thermal.Station2.StandbyTemp)
            PatDisplayArray.Add("IDSData.Hardware.Thermal.Station2.AlarmThreshold")
            PatArray.Add(IDSData.Hardware.Thermal.Station2.AlarmThreshold)
            PatDisplayArray.Add("IDSData.Hardware.Thermal.Station2.OnOff")
            PatArray.Add(IDSData.Hardware.Thermal.Station2.OnOff)

            PatDisplayArray.Add("IDSData.Hardware.Thermal.Station3.OperationTemp")
            PatArray.Add(IDSData.Hardware.Thermal.Station3.OperationTemp)
            PatDisplayArray.Add("IDSData.Hardware.Thermal.Station3.StandbyTemp")
            PatArray.Add(IDSData.Hardware.Thermal.Station3.StandbyTemp)
            PatDisplayArray.Add("IDSData.Hardware.Thermal.Station3.AlarmThreshold")
            PatArray.Add(IDSData.Hardware.Thermal.Station3.AlarmThreshold)
            PatDisplayArray.Add("IDSData.Hardware.Thermal.Station3.OnOff")
            PatArray.Add(IDSData.Hardware.Thermal.Station3.OnOff)

            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            ''''''''''' Volume  ''''''''''''''''
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            PatDisplayArray.Add("")
            PatArray.Add("")
            PatDisplayArray.Add("~")
            PatArray.Add("[VOLUME]")
            PatDisplayArray.Add("IDSData.Hardware.Volume.WeighingScaleFlag")
            PatArray.Add(IDSData.Hardware.Volume.WeighingScaleFlag)

            PatDisplayArray.Add("IDSData.Hardware.Volume.Left.OnOff")
            PatArray.Add(IDSData.Hardware.Volume.Left.OnOff)
            PatDisplayArray.Add("IDSData.Hardware.Volume.Left.CalibrationInterval") 'minutes
            PatArray.Add(IDSData.Hardware.Volume.Left.CalibrationInterval)
            PatDisplayArray.Add("IDSData.Hardware.Volume.Left.DesiredWeight")
            PatArray.Add(IDSData.Hardware.Volume.Left.DesiredWeight)
            PatDisplayArray.Add("IDSData.Hardware.Volume.Left.Tolerance")
            PatArray.Add(IDSData.Hardware.Volume.Left.Tolerance)
            PatDisplayArray.Add("IDSData.Hardware.Volume.Left.SetupDispenseDuration")
            PatArray.Add(IDSData.Hardware.Volume.Left.SetupDispenseDuration)
            PatDisplayArray.Add("IDSData.Hardware.Volume.Left.SetupMaterialAirPressure")
            PatArray.Add(IDSData.Hardware.Volume.Left.SetupMaterialAirPressure)
            PatDisplayArray.Add("IDSData.Hardware.Volume.Left.SetupRPM")
            PatArray.Add(IDSData.Hardware.Volume.Left.SetupRPM)
            PatDisplayArray.Add("IDSData.Hardware.Volume.Left.NumberOfAttempts")
            PatArray.Add(IDSData.Hardware.Volume.Left.NumberOfAttempts)
            PatDisplayArray.Add("IDSData.Hardware.Volume.Left.AdjustedMaterialAirPressure") 'for time pressure valve/syringe
            PatArray.Add(IDSData.Hardware.Volume.Left.AdjustedMaterialAirPressure)
            PatDisplayArray.Add("IDSData.Hardware.Volume.Left.AdjustedRPM") 'for auger valve
            PatArray.Add(IDSData.Hardware.Volume.Left.AdjustedRPM)
            PatDisplayArray.Add("IDSData.Hardware.Volume.Left.AdjustedDispenseDuration") 'for slider, jetting valve
            PatArray.Add(IDSData.Hardware.Volume.Left.AdjustedDispenseDuration)
            PatDisplayArray.Add("IDSData.Hardware.Volume.Left.PressureStepValue")
            PatArray.Add(IDSData.Hardware.Volume.Left.PressureStepValue)

            PatDisplayArray.Add("IDSData.Hardware.Volume.Left.DurationStepValue")
            PatArray.Add(IDSData.Hardware.Volume.Left.DurationStepValue)
            PatDisplayArray.Add("IDSData.Hardware.Volume.Left.RPMStepValue")
            PatArray.Add(IDSData.Hardware.Volume.Left.RPMStepValue)

            PatDisplayArray.Add("IDSData.Hardware.Volume.Left.RetractSpeed")
            PatArray.Add(IDSData.Hardware.Volume.Left.RetractSpeed)
            PatDisplayArray.Add("IDSData.Hardware.Volume.Left.RetractHeight")
            PatArray.Add(IDSData.Hardware.Volume.Left.RetractHeight)
            PatDisplayArray.Add("IDSData.Hardware.Volume.Left.RetractDelay")
            PatArray.Add(IDSData.Hardware.Volume.Left.RetractDelay)

            PatDisplayArray.Add("IDSData.Hardware.Volume.Left.PulseOnDuration")
            PatArray.Add(IDSData.Hardware.Volume.Left.PulseOnDuration)
            PatDisplayArray.Add("IDSData.Hardware.Volume.Left.PulseOffDuration")
            PatArray.Add(IDSData.Hardware.Volume.Left.PulseOffDuration)
            PatDisplayArray.Add("IDSData.Hardware.Volume.Left.JettingNoOfDispense")
            PatArray.Add(IDSData.Hardware.Volume.Left.JettingNoOfDispense)

            ' right
            PatDisplayArray.Add("IDSData.Hardware.Volume.Right.OnOff")
            PatArray.Add(IDSData.Hardware.Volume.Right.OnOff)
            PatDisplayArray.Add("IDSData.Hardware.Volume.Right.CalibrationInterval") 'minutes
            PatArray.Add(IDSData.Hardware.Volume.Right.CalibrationInterval)
            PatDisplayArray.Add("IDSData.Hardware.Volume.Right.DesiredWeight")
            PatArray.Add(IDSData.Hardware.Volume.Right.DesiredWeight)
            PatDisplayArray.Add("IDSData.Hardware.Volume.Right.Tolerance")
            PatArray.Add(IDSData.Hardware.Volume.Right.Tolerance)
            PatDisplayArray.Add("IDSData.Hardware.Volume.Right.SetupDispenseDuration")
            PatArray.Add(IDSData.Hardware.Volume.Right.SetupDispenseDuration)
            PatDisplayArray.Add("IDSData.Hardware.Volume.Right.SetupMaterialAirPressure")
            PatArray.Add(IDSData.Hardware.Volume.Right.SetupMaterialAirPressure)
            PatDisplayArray.Add("IDSData.Hardware.Volume.Right.SetupRPM")
            PatArray.Add(IDSData.Hardware.Volume.Right.SetupRPM)
            PatDisplayArray.Add("IDSData.Hardware.Volume.Right.NumberOfAttempts")
            PatArray.Add(IDSData.Hardware.Volume.Right.NumberOfAttempts)
            PatDisplayArray.Add("IDSData.Hardware.Volume.Right.AdjustedMaterialAirPressure") 'for time pressure valve/syringe
            PatArray.Add(IDSData.Hardware.Volume.Right.AdjustedMaterialAirPressure)
            PatDisplayArray.Add("IDSData.Hardware.Volume.Right.AdjustedRPM") 'for auger valve
            PatArray.Add(IDSData.Hardware.Volume.Right.AdjustedRPM)
            PatDisplayArray.Add("IDSData.Hardware.Volume.Right.AdjustedDispenseDuration") 'for slider, jetting valve
            PatArray.Add(IDSData.Hardware.Volume.Right.AdjustedDispenseDuration)
            PatDisplayArray.Add("IDSData.Hardware.Volume.Right.PressureStepValue")
            PatArray.Add(IDSData.Hardware.Volume.Right.PressureStepValue)

            PatDisplayArray.Add("IDSData.Hardware.Volume.Right.DurationStepValue")
            PatArray.Add(IDSData.Hardware.Volume.Right.DurationStepValue)
            PatDisplayArray.Add("IDSData.Hardware.Volume.Right.RPMStepValue")
            PatArray.Add(IDSData.Hardware.Volume.Right.RPMStepValue)

            PatDisplayArray.Add("IDSData.Hardware.Volume.Right.RetractSpeed")
            PatArray.Add(IDSData.Hardware.Volume.Right.RetractSpeed)
            PatDisplayArray.Add("IDSData.Hardware.Volume.Right.RetractHeight")
            PatArray.Add(IDSData.Hardware.Volume.Right.RetractHeight)
            PatDisplayArray.Add("IDSData.Hardware.Volume.Right.RetractDelay")
            PatArray.Add(IDSData.Hardware.Volume.Right.RetractDelay)

            PatDisplayArray.Add("IDSData.Hardware.Volume.Right.PulseOnDuration")
            PatArray.Add(IDSData.Hardware.Volume.Right.PulseOnDuration)
            PatDisplayArray.Add("IDSData.Hardware.Volume.Right.PulseOffDuration")
            PatArray.Add(IDSData.Hardware.Volume.Right.PulseOffDuration)
            PatDisplayArray.Add("IDSData.Hardware.Volume.Right.JettingNoOfDispense")
            PatArray.Add(IDSData.Hardware.Volume.Right.JettingNoOfDispense)


            '           Pattern Execution
            PatDisplayArray.Add("")
            PatArray.Add("")
            PatDisplayArray.Add("~")
            PatArray.Add("[PATTERNSETTING]")
            PatDisplayArray.Add("IDSData.Execution.Setting.ArcTemplateDir")
            PatArray.Add(IDSData.Execution.Setting.ArcTemplateDir)
            PatDisplayArray.Add("IDSData.Execution.Setting.ChipEdgeTemplateDir")
            PatArray.Add(IDSData.Execution.Setting.ChipEdgeTemplateDir)
            PatDisplayArray.Add("IDSData.Execution.Setting.CircleTemplateDir")
            PatArray.Add(IDSData.Execution.Setting.CircleTemplateDir)
            PatDisplayArray.Add("IDSData.Execution.Setting.CurrentPatternName")
            PatArray.Add(IDSData.Execution.Setting.CurrentPatternName)
            PatDisplayArray.Add("IDSData.Execution.Setting.DotTemplateDir")
            PatArray.Add(IDSData.Execution.Setting.DotTemplateDir)
            PatDisplayArray.Add("IDSData.Execution.Setting.FilledCircleTemplateDir")
            PatArray.Add(IDSData.Execution.Setting.FilledCircleTemplateDir)
            PatDisplayArray.Add("IDSData.Execution.Setting.FilledRecTemplateDir")
            PatArray.Add(IDSData.Execution.Setting.FilledRecTemplateDir)
            PatDisplayArray.Add("IDSData.Execution.Setting.LineTemplateDir")
            PatArray.Add(IDSData.Execution.Setting.LineTemplateDir)
            PatDisplayArray.Add("IDSData.Execution.Setting.PatternDir")
            PatArray.Add(IDSData.Execution.Setting.PatternDir)
            PatDisplayArray.Add("IDSData.Execution.Setting.RecTemplateDir")
            PatArray.Add(IDSData.Execution.Setting.RecTemplateDir)
            If IDSData.ParameterID.RecordID = "FactoryDefault" Then
                PatDisplayArray.Add("IDSData.Execution.Setting.DefaultFileToOpen")
                PatArray.Add(IDSData.Execution.Setting.DefaultFileToOpen)
            End If
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '''''''''''''''Programming '''''''''''''''''''''''
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            PatDisplayArray.Add("")
            PatArray.Add("")
            PatDisplayArray.Add("~")
            PatArray.Add("[PATTERNS ELEMENT]")
            I = 0
            If IDSData.Execution.Prog.Count > 0 Then

                For I = 0 To IDSData.Execution.Prog.Count - 1
                    IDSData.Execution.Pattern = CType(IDSData.Execution.Prog(I), _
                                                                CIDSData.CIDSPatternExecution.CIDSTemplate)

                    PatDisplayArray.Add("(" + I.ToString + ")Item")
                    PatArray.Add(IDSData.Execution.Pattern.Item)
                    PatDisplayArray.Add("(" + I.ToString + ")Name")
                    PatArray.Add(IDSData.Execution.Pattern.Name)
                    PatDisplayArray.Add("(" + I.ToString + ")ApproachDispHeight")
                    PatArray.Add(IDSData.Execution.Pattern.ApproachDispHeight)
                    PatDisplayArray.Add("(" + I.ToString + ")ArcRadius")
                    PatArray.Add(IDSData.Execution.Pattern.ArcRadius)
                    PatDisplayArray.Add("(" + I.ToString + ")ClearanceHeight")
                    PatArray.Add(IDSData.Execution.Pattern.ClearanceHeight)
                    PatDisplayArray.Add("(" + I.ToString + ")DetailingDistance")
                    PatArray.Add(IDSData.Execution.Pattern.DetailingDistance)
                    PatDisplayArray.Add("(" + I.ToString + ")DispenseDuration")
                    PatArray.Add(IDSData.Execution.Pattern.DispenseDuration)
                    PatDisplayArray.Add("(" + I.ToString + ")DispenseFlag")
                    PatArray.Add(IDSData.Execution.Pattern.DispenseFlag)
                    PatDisplayArray.Add("(" + I.ToString + ")EdgeClearance")
                    PatArray.Add(IDSData.Execution.Pattern.EdgeClearance)
                    PatDisplayArray.Add("(" + I.ToString + ")FilledHeight")
                    PatArray.Add(IDSData.Execution.Pattern.FilledHeight)
                    PatDisplayArray.Add("(" + I.ToString + ")Needle")
                    PatArray.Add(IDSData.Execution.Pattern.Needle)
                    PatDisplayArray.Add("(" + I.ToString + ")NeedleGap")
                    PatArray.Add(IDSData.Execution.Pattern.NeedleGap)
                    PatDisplayArray.Add("(" + I.ToString + ")Pitch")
                    PatArray.Add(IDSData.Execution.Pattern.Pitch)
                    PatDisplayArray.Add("(" + I.ToString + ")Pos1.X")
                    PatArray.Add(IDSData.Execution.Pattern.Pos1.X)
                    PatDisplayArray.Add("(" + I.ToString + ")Pos1.Y")
                    PatArray.Add(IDSData.Execution.Pattern.Pos1.Y)
                    PatDisplayArray.Add("(" + I.ToString + ")Pos1.Z")
                    PatArray.Add(IDSData.Execution.Pattern.Pos1.Z)
                    PatDisplayArray.Add("(" + I.ToString + ")Pos2.X")
                    PatArray.Add(IDSData.Execution.Pattern.Pos2.X)
                    PatDisplayArray.Add("(" + I.ToString + ")Pos2.Y")
                    PatArray.Add(IDSData.Execution.Pattern.Pos2.Y)
                    PatDisplayArray.Add("(" + I.ToString + ")Pos2.Z")
                    PatArray.Add(IDSData.Execution.Pattern.Pos2.Z)
                    PatDisplayArray.Add("(" + I.ToString + ")pos3.X")
                    PatArray.Add(IDSData.Execution.Pattern.pos3.X)
                    PatDisplayArray.Add("(" + I.ToString + ")pos3.Y")
                    PatArray.Add(IDSData.Execution.Pattern.pos3.Y)
                    PatDisplayArray.Add("(" + I.ToString + ")pos3.Z")
                    PatArray.Add(IDSData.Execution.Pattern.pos3.Z)
                    PatDisplayArray.Add("(" + I.ToString + ")RetractDelay")
                    PatArray.Add(IDSData.Execution.Pattern.RetractDelay)
                    PatDisplayArray.Add("(" + I.ToString + ")RetractHeight")
                    PatArray.Add(IDSData.Execution.Pattern.RetractHeight)
                    PatDisplayArray.Add("(" + I.ToString + ")RetractSpeed")
                    PatArray.Add(IDSData.Execution.Pattern.RetractSpeed)
                    PatDisplayArray.Add("(" + I.ToString + ")RotatingAngle")
                    PatArray.Add(IDSData.Execution.Pattern.RotatingAngle)
                    PatDisplayArray.Add("(" + I.ToString + ")SpiralFlag")
                    PatArray.Add(IDSData.Execution.Pattern.SpiralFlag)
                    PatDisplayArray.Add("(" + I.ToString + ")SubFileName")
                    PatArray.Add(IDSData.Execution.Pattern.SubFileName)
                    PatDisplayArray.Add("(" + I.ToString + ")TravelDelay")
                    PatArray.Add(IDSData.Execution.Pattern.TravelDelay)
                    PatDisplayArray.Add("(" + I.ToString + ")TravelSpeed")
                    PatArray.Add(IDSData.Execution.Pattern.TravelSpeed)
                    PatDisplayArray.Add("(" + I.ToString + ")RetractHeight")
                    PatArray.Add(IDSData.Execution.Pattern.RetractHeight)

                    PatDisplayArray.Add("(" + I.ToString + ")IONumber")
                    PatArray.Add(IDSData.Execution.Pattern.IONumber)

                    PatDisplayArray.Add("")
                    PatArray.Add("")
                Next

            End If
            PatDisplayArray.Add("~")
            PatArray.Add("[PATTERNS END]")
            PatDisplayArray.Add("")
            PatArray.Add("")

        Catch ex As Exception
            IDSData.MsgErr = ex.ToString
            MessageBox.Show(IDSData.MsgErr)
            Return False
        End Try




        Dim path As String = MyFileName
        Dim Sw As StreamWriter
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '''''' save data from memory to text file
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If Textflag = True Then
            Try
                Sw = New StreamWriter(path)

                I = 0
                For I = 0 To PatDisplayArray.Count - 1

                    If CStr(PatDisplayArray.Item(I)) = "~" Or CStr(PatDisplayArray.Item(I)) = "" Then
                        Sw.WriteLine("{0}{1}", PatDisplayArray.Item(I), PatArray.Item(I)) 'write devices header
                    Else
                        Sw.WriteLine("{0} ={1}", PatDisplayArray.Item(I), PatArray.Item(I)) 'write devices variables
                    End If
                Next

                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                '' end of text write
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Catch ex As Exception
                IDSData.MsgErr = ex.ToString
                MessageBox.Show(IDSData.MsgErr)
                Sw.Flush()  'clear the stream
                Sw.Close()  'close file
                Return False
            End Try
            Sw.Flush()  'clear the stream
            Sw.Close()  'close file
            Return True
        End If


        Dim Bw As BinaryWriter
        Dim Fs As FileStream

        If Textflag = False Then
            Try
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                '''''' save data from memory to binary file
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                Fs = New FileStream(path, FileMode.Create)
                Bw = New BinaryWriter(Fs)
                Bw.BaseStream.Seek(0, SeekOrigin.Begin)

                I = 0
                For I = 0 To PatDisplayArray.Count - 1
                    Dim str As String
                    If CStr(PatDisplayArray.Item(I)) = "~" Or CStr(PatDisplayArray.Item(I)) = "" Then
                        str = CStr(PatDisplayArray.Item(I)) + CStr(PatArray.Item(I)) + Environment.NewLine 'write devices header
                        Bw.Write(str)
                    Else
                        str = CStr(PatDisplayArray.Item(I)) + " =" + CStr(PatArray.Item(I)) + Environment.NewLine 'write devices variables
                        Bw.Write(str)
                    End If

                Next
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                '' end of binary write
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


            Catch ex As Exception
                IDSData.MsgErr = ex.ToString
                MessageBox.Show(IDSData.MsgErr)
                Bw.Close()
                Fs.Close()
                Return False
            End Try
            Bw.Close()
            Fs.Close()
            Return True
        End If


    End Function

    Private Function GetBooleanArrayValue(ByRef index As Integer)
        index += 1
        Return CBoolean(PatArray.Item(index))
    End Function

    Private Function GetIntegerArrayValue(ByRef index As Integer)
        index += 1
        Return CInteger(PatArray.Item(index))
    End Function

    Private Function GetDoubleArrayValue(ByRef index As Integer)
        index += 1
        Return CDouble(PatArray.Item(index))
    End Function

    Private Function GetStringArrayValue(ByRef index As Integer)
        index += 1
        Return CString(PatArray.Item(index))
    End Function

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'open global data into pat file with path in text format
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Public Function OpenPathFileName(ByRef MyFileName As String) As Boolean

        Dim path As String = MyFileName

        Dim SWVersion As Double                             'Global data version number
        Dim Index As Integer                                'Variable Record Index
        Dim I As Integer

        IDSData.MsgErr = ""


        If Textflag = False Then

            Dim Fs1 As FileStream = New FileStream(path, FileMode.Open)
            Dim Br As BinaryReader = New BinaryReader(Fs1)
            Try
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                '''''' save data from memory to binary file
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

                Br.BaseStream.Seek(0, SeekOrigin.Begin)
                I = 0
                PatDisplayArray.Clear()                         'clear the working arraylist
                PatArray.Clear()

                Dim c As Integer = Br.PeekChar
                While Fs1.Position < Fs1.Length
                    Dim str As String = Br.ReadString
                    str = str.Replace(vbCrLf, "")
                    Dim Retstr As String() = Splitstring(str)   'extract substrings from string line
                    PatDisplayArray.Add(Retstr(0))              'added to the arrylist (discription) 
                    PatArray.Add(Retstr(Retstr.Length - 1))     'added to the arraylist (value)
                    c = Br.PeekChar
                End While
                Br.Close()
                Fs1.Close()
            Catch ex As Exception
                Br.Close()
                Fs1.Close()
                IDSData.MsgErr = ""
                If ex.Message = "Unable to read beyond the end of the stream." Then
                    IDSData.MsgErr = ex.message
                End If
            End Try
        End If


        '"Unable to read beyond the end of the stream." 
        If Textflag = True Or IDSData.MsgErr <> "" Then
            Dim Sr As StreamReader = New StreamReader(path)

            Try
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                '''''' save data from memory to text file
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                'Create reader stream
                Sr = File.OpenText(path)                        'Create a file to write to.
                I = 0
                PatDisplayArray.Clear()                         'clear the working arraylist
                PatArray.Clear()
                Do While Sr.Peek() >= 0                         'is it the reader stream end?
                    Dim str As String = Sr.ReadLine()           'read string line from file
                    Dim Retstr As String() = Splitstring(str)   'extract substrings from string line
                    PatDisplayArray.Add(Retstr(0))              'added to the arrylist (discription) 
                    PatArray.Add(Retstr(Retstr.Length - 1))     'added to the arraylist (value)
                Loop
            Catch ex As Exception
                IDSData.MsgErr = "Reading Error found at << " & CStr(PatDisplayArray.Item(Index)) & " >> " + vbCrLf + vbCrLf + ex.ToString
                MessageBox.Show(IDSData.MsgErr)
                Sr.Close()
                Return False
            End Try
            Sr.Close()
        End If

        Try
            'admin
            'If path.IndexOf("FactoryDefault") > 0 Then
            Index = PatArray.IndexOf("[VERSION]")
            SWVersion = GetDoubleArrayValue(Index)

            'If IDSData.ParameterID.RecordID = "FactoryDefault" Then
            'If path.ToUpper().IndexOf("FACTORYDEFAULT") > 0 Then
            Index = PatArray.IndexOf("[ADMINISTRATION]")
            IDSData.Admin.Folder.DataPath = GetStringArrayValue(Index)
            IDSData.Admin.Folder.PatternPath = GetStringArrayValue(Index)
            IDSData.Admin.Folder.FileExtension = GetStringArrayValue(Index)

            ''''''''''''''' ParameterID ''''''''''''''''
            Index += 1
            'IDSData.ParameterID.GroupID = GetStringArrayValue(Index)
            Index += 1
            'IDSData.ParameterID.RecordID = GetStringArrayValue(Index)

            ''''''''''''''''' data from database 
            Index += 1
            'IDSData.Admin.User.ContactNo = GetStringArrayValue(Index)
            Index += 1
            'IDSData.Admin.User.Department = GetStringArrayValue(Index)
            IDSData.Admin.User.ID = GetStringArrayValue(Index)
            Index += 1
            'IDSData.Admin.User.Name = GetStringArrayValue(Index)
            Index += 1
            'IDSData.Admin.User.PassWord = GetStringArrayValue(Index)
            Index += 1
            'IDSData.Admin.User.RunApplication = GetStringArrayValue(Index)
            IDSData.Admin.User.Group.ID = GetStringArrayValue(Index)
            Index += 1
            'IDSData.Admin.User.Group.Remark = GetStringArrayValue(Index)


            'End If
            ' End If

            '''''''''''''''''''''''''''''''''''
            ''''''''''' Camer '''''''''''''''''
            '''''''''''''''''''''''''''''''''''

            Index = PatArray.IndexOf("[CAMERA]")
            IDSData.Hardware.Camera.CalibrationFlag = GetBooleanArrayValue(Index)
            IDSData.Hardware.Camera.CalibrationPos.X = GetDoubleArrayValue(Index)
            IDSData.Hardware.Camera.CalibrationPos.Y = GetDoubleArrayValue(Index)

            I = 0
            If IDSData.Hardware.Camera.ChipDispenseType.Point.Length > 0 Then
                For I = 0 To IDSData.Hardware.Camera.ChipDispenseType.Point.Length - 1
                    IDSData.Hardware.Camera.ChipDispenseType.Point(I).X = GetDoubleArrayValue(Index)
                    IDSData.Hardware.Camera.ChipDispenseType.Point(I).Y = GetDoubleArrayValue(Index)
                    IDSData.Hardware.Camera.ChipDispenseType.Point(I).Z = GetDoubleArrayValue(Index)
                Next
            End If


            I = 0
            If IDSData.Hardware.Camera.Compensation.Matrix.Length > 0 Then
                For I = 0 To IDSData.Hardware.Camera.Compensation.Matrix.Length - 1
                    IDSData.Hardware.Camera.Compensation.Matrix(I).W = GetDoubleArrayValue(Index)
                    IDSData.Hardware.Camera.Compensation.Matrix(I).X = GetDoubleArrayValue(Index)
                    IDSData.Hardware.Camera.Compensation.Matrix(I).Y = GetDoubleArrayValue(Index)
                    IDSData.Hardware.Camera.Compensation.Matrix(I).Z = GetDoubleArrayValue(Index)
                Next
            End If

            IDSData.Hardware.Camera.EdgeDetect.NeedleOD = GetIntegerArrayValue(Index)
            IDSData.Hardware.Camera.EdgeDetect.ResultFlag = GetBooleanArrayValue(Index)

            I = 0
            If IDSData.Hardware.Camera.EdgeDetect.Point.Length > 0 Then
                For I = 0 To IDSData.Hardware.Camera.EdgeDetect.Point.Length - 1
                    IDSData.Hardware.Camera.EdgeDetect.Point(I).X = GetDoubleArrayValue(Index)
                    IDSData.Hardware.Camera.EdgeDetect.Point(I).Y = GetDoubleArrayValue(Index)
                    IDSData.Hardware.Camera.EdgeDetect.Point(I).Z = GetDoubleArrayValue(Index)
                Next
            End If

            'pixels & real world info
            IDSData.Hardware.Camera.FiducialPos.X = GetDoubleArrayValue(Index)
            IDSData.Hardware.Camera.FiducialPos.Y = GetDoubleArrayValue(Index)
            IDSData.Hardware.Camera.FiducialPos.Z = GetDoubleArrayValue(Index)
            IDSData.Hardware.Camera.FileFiducial1 = GetStringArrayValue(Index)
            IDSData.Hardware.Camera.FileFiducial2 = GetStringArrayValue(Index)
            IDSData.Hardware.Camera.FileRejectMark = GetStringArrayValue(Index)
            IDSData.Hardware.Camera.GraphicsDisplayResolution.X = GetDoubleArrayValue(Index)
            IDSData.Hardware.Camera.GraphicsDisplayResolution.Y = GetDoubleArrayValue(Index)
            IDSData.Hardware.Camera.ImageResolution = GetIntegerArrayValue(Index)
            If IDSData.ParameterID.RecordID = "FactoryDefault" Then
                IDSData.Hardware.Camera.OffsetLaser.X = GetDoubleArrayValue(Index)
                IDSData.Hardware.Camera.OffsetLaser.Y = GetDoubleArrayValue(Index)
                IDSData.Hardware.Camera.OffsetLNeedle.X = GetDoubleArrayValue(Index)
                IDSData.Hardware.Camera.OffsetLNeedle.Y = GetDoubleArrayValue(Index)
                IDSData.Hardware.Camera.OffsetRNeedle.X = GetDoubleArrayValue(Index)
                IDSData.Hardware.Camera.OffsetRNeedle.Y = GetDoubleArrayValue(Index)
                IDSData.Hardware.Camera.OffsetTouchProbe.X = GetDoubleArrayValue(Index)
                IDSData.Hardware.Camera.OffsetTouchProbe.Y = GetDoubleArrayValue(Index)
                IDSData.Hardware.Camera.PixelsToMM.X = GetDoubleArrayValue(Index)
                IDSData.Hardware.Camera.PixelsToMM.Y = GetDoubleArrayValue(Index)
                IDSData.Hardware.Camera.PixelRatio = GetDoubleArrayValue(Index)
                IDSData.Hardware.Camera.ReferencePos.X = GetDoubleArrayValue(Index)
                IDSData.Hardware.Camera.ReferencePos.Y = GetDoubleArrayValue(Index)
                IDSData.Hardware.Camera.ReferencePos.Z = GetDoubleArrayValue(Index)
            Else
                Index += 14
            End If

            IDSData.Hardware.Camera.SoftHomePos.X = GetDoubleArrayValue(Index)
            IDSData.Hardware.Camera.SoftHomePos.Y = GetDoubleArrayValue(Index)
            IDSData.Hardware.Camera.SoftHomePos.Z = GetDoubleArrayValue(Index)

            'calibration info
            IDSData.Hardware.Camera.CalibrationNumColumns = GetIntegerArrayValue(Index)
            IDSData.Hardware.Camera.CalibrationNumRows = GetIntegerArrayValue(Index)
            IDSData.Hardware.Camera.CalibrationSpacing = GetIntegerArrayValue(Index)
            IDSData.Hardware.Camera.CalibrationBlockPos.X = GetDoubleArrayValue(Index)
            IDSData.Hardware.Camera.CalibrationBlockPos.Y = GetDoubleArrayValue(Index)
            IDSData.Hardware.Camera.CalibrationBlockPos.Z = GetDoubleArrayValue(Index)
            IDSData.Hardware.Camera.BlockSize.X = GetDoubleArrayValue(Index)
            IDSData.Hardware.Camera.BlockSize.Y = GetDoubleArrayValue(Index)
            IDSData.Hardware.Camera.BlockSizeTolerance = GetDoubleArrayValue(Index)
            IDSData.Hardware.Camera.BlockPosTolerance = GetDoubleArrayValue(Index)
            IDSData.Hardware.Camera.BlockRotationTolerance = GetDoubleArrayValue(Index)
            IDSData.Hardware.Camera.ImageFindDir = GetBooleanArrayValue(Index)
            IDSData.Hardware.Camera.ImageVertical = GetBooleanArrayValue(Index)
            IDSData.Hardware.Camera.Threshold = GetIntegerArrayValue(Index)
            IDSData.Hardware.Camera.Brightness = GetIntegerArrayValue(Index)
            IDSData.Hardware.Camera.Contrast = GetIntegerArrayValue(Index)
            IDSData.Hardware.Camera.ROISize = GetDoubleArrayValue(Index)
            IDSData.Hardware.Camera.ResultSize.X = GetDoubleArrayValue(Index)
            IDSData.Hardware.Camera.ResultSize.Y = GetDoubleArrayValue(Index)
            IDSData.Hardware.Camera.ResultDia = GetDoubleArrayValue(Index)
            IDSData.Hardware.Camera.ResultRotation = GetDoubleArrayValue(Index)
            If SWVersion > 1.2 Then
                IDSData.Hardware.Camera.DotQCEnable = GetBooleanArrayValue(Index)
            Else
                IDSData.Hardware.Camera.DotQCEnable = False
            End If

            'conveyor
            Index = PatArray.IndexOf("[CONVEYOR]")
            IDSData.Hardware.Conveyor.Width = GetDoubleArrayValue(Index)
            IDSData.Hardware.Conveyor.FullWidth = GetIntegerArrayValue(Index)
            IDSData.Hardware.Conveyor.MinWidth = GetIntegerArrayValue(Index)
            IDSData.Hardware.Conveyor.Speed = GetIntegerArrayValue(Index)
            IDSData.Hardware.Conveyor.TimeOut = GetIntegerArrayValue(Index)
            IDSData.Hardware.Conveyor.WidthMoveStep = GetDoubleArrayValue(Index)

            If SWVersion > 1.2 Then
                IDSData.Hardware.Conveyor.upstreamTimeout = GetIntegerArrayValue(Index)
                IDSData.Hardware.Conveyor.downstreamTimeout = GetIntegerArrayValue(Index)
            Else
                IDSData.Hardware.Conveyor.upstreamTimeout = 30
                IDSData.Hardware.Conveyor.downstreamTimeout = 30
            End If

            'dispenser
            Index = PatArray.IndexOf("[DISPENSER]")
            IDSData.Hardware.Dispenser.CurrentHeads = GetIntegerArrayValue(Index)

            Index = PatArray.IndexOf("[DISPENSER@LEFT]")
            IDSData.Hardware.Dispenser.Left.MaterialAirPressure = GetDoubleArrayValue(Index)
            IDSData.Hardware.Dispenser.Left.SuckbackPressure = GetDoubleArrayValue(Index)
            IDSData.Hardware.Dispenser.Left.RPM = GetStringArrayValue(Index)
            IDSData.Hardware.Dispenser.Left.RetractTime = GetStringArrayValue(Index)
            IDSData.Hardware.Dispenser.Left.RetractDelay = GetStringArrayValue(Index)
            IDSData.Hardware.Dispenser.Left.Pulse = GetStringArrayValue(Index)
            IDSData.Hardware.Dispenser.Left.Pause = GetStringArrayValue(Index)
            IDSData.Hardware.Dispenser.Left.Count = GetStringArrayValue(Index)
            IDSData.Hardware.Dispenser.Left.ValveTemperature = GetDoubleArrayValue(Index)

            IDSData.Hardware.Dispenser.Left.NeedleTipLength = GetStringArrayValue(Index)
            IDSData.Hardware.Dispenser.Left.NeedleColor = GetStringArrayValue(Index)
            IDSData.Hardware.Dispenser.Left.MaterialInfo = GetStringArrayValue(Index)
            IDSData.Hardware.Dispenser.Left.AutoPurgingDuration = GetIntegerArrayValue(Index)
            IDSData.Hardware.Dispenser.Left.AutoCleaningDuration = GetIntegerArrayValue(Index)
            IDSData.Hardware.Dispenser.Left.AutoPurgingInterval = GetIntegerArrayValue(Index)
            IDSData.Hardware.Dispenser.Left.PotLifeDuration = GetIntegerArrayValue(Index)
            IDSData.Hardware.Dispenser.Left.PotLifeOption = GetBooleanArrayValue(Index)
            IDSData.Hardware.Dispenser.Left.AutoPurgingOption = GetBooleanArrayValue(Index)
            IDSData.Hardware.Dispenser.Left.HeadType = GetStringArrayValue(Index)
            If SWVersion > 1.6 Then
                If IDSData.ParameterID.RecordID = "FactoryDefault" Then
                    IDSData.Hardware.Dispenser.Left.JettingCalOnce = GetBooleanArrayValue(Index)
                    IDSData.Hardware.Dispenser.Left.AugerCalOnce = GetBooleanArrayValue(Index)
                    IDSData.Hardware.Dispenser.Left.SlideValveCalOnce = GetBooleanArrayValue(Index)
                    IDSData.Hardware.Dispenser.Left.TimePressureValveCalOnce = GetBooleanArrayValue(Index)
                    IDSData.Hardware.Dispenser.Left.TimePressureSyringeCalOnce = GetBooleanArrayValue(Index)
                End If
            End If
            'left dispenser - pressure control

            Index = PatArray.IndexOf("[DISPENSER@RIGHT]")
            IDSData.Hardware.Dispenser.Right.MaterialAirPressure = GetDoubleArrayValue(Index)
            IDSData.Hardware.Dispenser.Right.SuckbackPressure = GetDoubleArrayValue(Index)
            IDSData.Hardware.Dispenser.Right.RPM = GetStringArrayValue(Index)
            IDSData.Hardware.Dispenser.Right.RetractTime = GetStringArrayValue(Index)
            IDSData.Hardware.Dispenser.Right.RetractDelay = GetStringArrayValue(Index)
            IDSData.Hardware.Dispenser.Right.Pulse = GetStringArrayValue(Index)
            IDSData.Hardware.Dispenser.Right.Pause = GetStringArrayValue(Index)
            IDSData.Hardware.Dispenser.Right.Count = GetStringArrayValue(Index)
            IDSData.Hardware.Dispenser.Right.ValveTemperature = GetDoubleArrayValue(Index)

            IDSData.Hardware.Dispenser.Right.NeedleTipLength = GetStringArrayValue(Index)
            IDSData.Hardware.Dispenser.Right.NeedleColor = GetStringArrayValue(Index)
            IDSData.Hardware.Dispenser.Right.MaterialInfo = GetStringArrayValue(Index)
            IDSData.Hardware.Dispenser.Right.AutoPurgingDuration = GetIntegerArrayValue(Index)
            IDSData.Hardware.Dispenser.Right.AutoCleaningDuration = GetIntegerArrayValue(Index)
            IDSData.Hardware.Dispenser.Right.AutoPurgingInterval = GetIntegerArrayValue(Index)
            IDSData.Hardware.Dispenser.Right.PotLifeDuration = GetIntegerArrayValue(Index)
            IDSData.Hardware.Dispenser.Right.PotLifeOption = GetBooleanArrayValue(Index)
            IDSData.Hardware.Dispenser.Right.AutoPurgingOption = GetBooleanArrayValue(Index)
            IDSData.Hardware.Dispenser.Right.HeadType = GetStringArrayValue(Index)

            If SWVersion > 1.6 Then
                If IDSData.ParameterID.RecordID = "FactoryDefault" Then
                    IDSData.Hardware.Dispenser.Right.JettingCalOnce = GetBooleanArrayValue(Index)
                    IDSData.Hardware.Dispenser.Right.AugerCalOnce = GetBooleanArrayValue(Index)
                    IDSData.Hardware.Dispenser.Right.SlideValveCalOnce = GetBooleanArrayValue(Index)
                    IDSData.Hardware.Dispenser.Right.TimePressureValveCalOnce = GetBooleanArrayValue(Index)
                    IDSData.Hardware.Dispenser.Right.TimePressureSyringeCalOnce = GetBooleanArrayValue(Index)
                End If
            End If
            'right dispenser - pressure control
            'gantry
            Index = PatArray.IndexOf("[GANTRY]")
            If IDSData.ParameterID.RecordID = "FactoryDefault" Then
                IDSData.Hardware.Gantry.WeighingScalePosition.X = GetDoubleArrayValue(Index)
                IDSData.Hardware.Gantry.WeighingScalePosition.Y = GetDoubleArrayValue(Index)
                IDSData.Hardware.Gantry.WeighingScalePosition.Z = GetDoubleArrayValue(Index)
                IDSData.Hardware.Gantry.CleanPosition.X = GetDoubleArrayValue(Index)
                IDSData.Hardware.Gantry.CleanPosition.Y = GetDoubleArrayValue(Index)
                IDSData.Hardware.Gantry.CleanPosition.Z = GetDoubleArrayValue(Index)
            Else
                Index += 6
            End If

            IDSData.Hardware.Gantry.ElementZSpeed = GetDoubleArrayValue(Index)
            IDSData.Hardware.Gantry.ElementXYSpeed = GetDoubleArrayValue(Index)
            IDSData.Hardware.Gantry.GraphicDisplayRatio = GetDoubleArrayValue(Index)
            IDSData.Hardware.Gantry.GraphicDisplayXYLT.X = GetDoubleArrayValue(Index)
            IDSData.Hardware.Gantry.GraphicDisplayXYLT.Y = GetDoubleArrayValue(Index)
            IDSData.Hardware.Gantry.GraphicDisplayXYRB.X = GetDoubleArrayValue(Index)
            IDSData.Hardware.Gantry.GraphicDisplayXYRB.Y = GetDoubleArrayValue(Index)
            IDSData.Hardware.Gantry.MaxAccelerationLimit = GetDoubleArrayValue(Index)
            IDSData.Hardware.Gantry.MaxSpeedLimit = GetDoubleArrayValue(Index)
            If IDSData.ParameterID.RecordID = "FactoryDefault" Then
                IDSData.Hardware.Gantry.ParkPosition.X = GetDoubleArrayValue(Index)
                IDSData.Hardware.Gantry.ParkPosition.Y = GetDoubleArrayValue(Index)
                IDSData.Hardware.Gantry.ParkPosition.Z = GetDoubleArrayValue(Index)
                IDSData.Hardware.Gantry.PurgePosition.X = GetDoubleArrayValue(Index)
                IDSData.Hardware.Gantry.PurgePosition.Y = GetDoubleArrayValue(Index)
                IDSData.Hardware.Gantry.PurgePosition.Z = GetDoubleArrayValue(Index)
                IDSData.Hardware.Gantry.SystemOriginPos.X = GetDoubleArrayValue(Index)
                IDSData.Hardware.Gantry.SystemOriginPos.Y = GetDoubleArrayValue(Index)
                IDSData.Hardware.Gantry.SystemOriginPos.Z = GetDoubleArrayValue(Index)
            Else
                Index += 9
            End If
            IDSData.Hardware.Gantry.SystemUnit = GetBooleanArrayValue(Index)
            If IDSData.ParameterID.RecordID = "FactoryDefault" Then
                IDSData.Hardware.Gantry.WorkArea.X = GetDoubleArrayValue(Index)
                IDSData.Hardware.Gantry.WorkArea.Y = GetDoubleArrayValue(Index)
                IDSData.Hardware.Gantry.WorkArea.Z.Max = GetDoubleArrayValue(Index)
                IDSData.Hardware.Gantry.WorkArea.Z.Min = GetDoubleArrayValue(Index)
            Else
                Index += 4
            End If


            IDSData.Hardware.Gantry.ServiceXYSpeed = GetDoubleArrayValue(Index)
            IDSData.Hardware.Gantry.ServiceZSpeed = GetDoubleArrayValue(Index)

            'Index = PatArray.IndexOf("IDSData.Hardware.Gantry.ChangeSyringePosition.X")
            'If PatArray.Contains("IDSData.Hardware.Gantry.ChangeSyringePosition.X") Then
            If IDSData.ParameterID.RecordID = "FactoryDefault" Then
                IDSData.Hardware.Gantry.ChangeSyringePosition.X = GetDoubleArrayValue(Index)
                IDSData.Hardware.Gantry.ChangeSyringePosition.Y = GetDoubleArrayValue(Index)
                IDSData.Hardware.Gantry.ChangeSyringePosition.Z = GetDoubleArrayValue(Index)
                If SWVersion > 1.3 Then
                    IDSData.Hardware.Gantry.WeighingScaleBottomRight.X = GetDoubleArrayValue(Index)
                    IDSData.Hardware.Gantry.WeighingScaleBottomRight.Y = GetDoubleArrayValue(Index)
                    'IDSData.Hardware.Gantry.WeighingScaleBottomRight.Z = GetDoubleArrayValue(Index)
                End If
            End If

            ' End If
            'If Index > 0 Then
            '    IDSData.Hardware.Gantry.ChangeSyringePosition.X = GetDoubleArrayValue(Index)
            '    IDSData.Hardware.Gantry.ChangeSyringePosition.Y = GetDoubleArrayValue(Index)
            '    IDSData.Hardware.Gantry.ChangeSyringePosition.Z = GetDoubleArrayValue(Index)
            'End If


            'height sensor
            Index = PatArray.IndexOf("[HEIGHTSENOR]")
            IDSData.Hardware.HeightSensor.BoardHeight = GetDoubleArrayValue(Index)
            IDSData.Hardware.HeightSensor.CalibratorHeightPos.X = GetDoubleArrayValue(Index)
            IDSData.Hardware.HeightSensor.CalibratorHeightPos.Y = GetDoubleArrayValue(Index)
            IDSData.Hardware.HeightSensor.SelectType = GetBooleanArrayValue(Index)

            Index = PatArray.IndexOf("[HEIGHTSENOR@LASER]")
            If IDSData.ParameterID.RecordID = "FactoryDefault" Then
                IDSData.Hardware.HeightSensor.Laser.CurrentPos.X = GetDoubleArrayValue(Index)
                IDSData.Hardware.HeightSensor.Laser.CurrentPos.Y = GetDoubleArrayValue(Index)
                IDSData.Hardware.HeightSensor.Laser.CurrentPos.Z = GetDoubleArrayValue(Index)
                IDSData.Hardware.HeightSensor.Laser.HeightReference = GetDoubleArrayValue(Index)
                IDSData.Hardware.HeightSensor.Laser.OffsetPos.X = GetDoubleArrayValue(Index)
                IDSData.Hardware.HeightSensor.Laser.OffsetPos.Y = GetDoubleArrayValue(Index)
            End If

            Index = PatArray.IndexOf("[HEIGHTSENOR@TOUCHPROBE]")
            IDSData.Hardware.HeightSensor.TP.CurrentPos.X = GetDoubleArrayValue(Index)
            IDSData.Hardware.HeightSensor.TP.CurrentPos.Y = GetDoubleArrayValue(Index)
            IDSData.Hardware.HeightSensor.TP.CurrentPos.Z = GetDoubleArrayValue(Index)
            IDSData.Hardware.HeightSensor.TP.HeightReference = GetDoubleArrayValue(Index)
            IDSData.Hardware.HeightSensor.TP.OffsetPos.X = GetDoubleArrayValue(Index)
            IDSData.Hardware.HeightSensor.TP.OffsetPos.Y = GetDoubleArrayValue(Index)

            Index = PatArray.IndexOf("[HEIGHTSENOR@LVDT]")
            IDSData.Hardware.HeightSensor.LVDTCalBackGround = GetBooleanArrayValue(Index)
            IDSData.Hardware.HeightSensor.LVDTCalBrightness = GetIntegerArrayValue(Index)
            IDSData.Hardware.HeightSensor.LVDTCalThreshold = GetIntegerArrayValue(Index)
            IDSData.Hardware.HeightSensor.LVDTCalOpen = GetIntegerArrayValue(Index)
            IDSData.Hardware.HeightSensor.LVDTCalClose = GetIntegerArrayValue(Index)
            IDSData.Hardware.HeightSensor.LVDTCalMaxRadius = GetDoubleArrayValue(Index)
            IDSData.Hardware.HeightSensor.LVDTCalMinRadius = GetDoubleArrayValue(Index)
            IDSData.Hardware.HeightSensor.LVDTCalRoughness = GetIntegerArrayValue(Index)
            IDSData.Hardware.HeightSensor.LVDTCalCompactness = GetIntegerArrayValue(Index)

            'ringlight
            Index = PatArray.IndexOf("[LIGHTBOX]")
            IDSData.Hardware.LightBox.Brightness = GetIntegerArrayValue(Index)

            'needle
            Index = PatArray.IndexOf("[NEEDLE@LEFT]")
            IDSData.Hardware.Needle.Left.ArcRadius = GetDoubleArrayValue(Index)
            If IDSData.ParameterID.RecordID = "FactoryDefault" Then
                IDSData.Hardware.Needle.Left.CalibratorPos.X = GetDoubleArrayValue(Index)
                IDSData.Hardware.Needle.Left.CalibratorPos.Y = GetDoubleArrayValue(Index)
                IDSData.Hardware.Needle.Left.CalibratorPos.Z = GetDoubleArrayValue(Index)
            Else
                Index += 3
            End If


            IDSData.Hardware.Needle.Left.HeightApproach = GetIntegerArrayValue(Index)
            IDSData.Hardware.Needle.Left.HeightClearance = GetDoubleArrayValue(Index)
            IDSData.Hardware.Needle.Left.HeightRetract = GetIntegerArrayValue(Index)

            If IDSData.ParameterID.RecordID = "FactoryDefault" Then
                IDSData.Hardware.Needle.Left.NeedleCalibrationPosition.X = GetDoubleArrayValue(Index)
                IDSData.Hardware.Needle.Left.NeedleCalibrationPosition.Y = GetDoubleArrayValue(Index)
                IDSData.Hardware.Needle.Left.NeedleCalibrationPosition.Z = GetDoubleArrayValue(Index)
                IDSData.Hardware.Needle.Left.TouchSensorZPosition = GetDoubleArrayValue(Index)
            Else
                Index += 4
            End If


            IDSData.Hardware.Needle.Left.ArrayDotPos1.X = GetDoubleArrayValue(Index)
            IDSData.Hardware.Needle.Left.ArrayDotPos1.Y = GetDoubleArrayValue(Index)
            IDSData.Hardware.Needle.Left.ArrayDotPos1.Z = GetDoubleArrayValue(Index)

            If SWVersion >= 1.2 Then

                IDSData.Hardware.Needle.Left.ArrayDotPos3.X = GetDoubleArrayValue(Index)
                IDSData.Hardware.Needle.Left.ArrayDotPos3.Y = GetDoubleArrayValue(Index)
                IDSData.Hardware.Needle.Left.ArrayDotPos3.Z = GetDoubleArrayValue(Index)

            End If

            IDSData.Hardware.Needle.Left.CalBackground = GetBooleanArrayValue(Index)
            IDSData.Hardware.Needle.Left.CalBrightness = GetIntegerArrayValue(Index)
            IDSData.Hardware.Needle.Left.CalThreshold = GetIntegerArrayValue(Index)
            IDSData.Hardware.Needle.Left.CalOpen = GetIntegerArrayValue(Index)
            IDSData.Hardware.Needle.Left.CalClose = GetIntegerArrayValue(Index)
            IDSData.Hardware.Needle.Left.CalMaxRadius = GetDoubleArrayValue(Index)
            IDSData.Hardware.Needle.Left.CalMinRadius = GetDoubleArrayValue(Index)
            IDSData.Hardware.Needle.Left.CalRoughness = GetDoubleArrayValue(Index)
            IDSData.Hardware.Needle.Left.CalCompactness = GetDoubleArrayValue(Index)

            ''' [version 1.1]
            If SWVersion >= 1.1 Then
                IDSData.Hardware.Needle.Left.DispenseDot.ApproachHeight = GetDoubleArrayValue(Index)
                IDSData.Hardware.Needle.Left.DispenseDot.ArcRadius = GetDoubleArrayValue(Index)
                IDSData.Hardware.Needle.Left.DispenseDot.ClearanceHeight = GetDoubleArrayValue(Index)
                IDSData.Hardware.Needle.Left.DispenseDot.DispenseDuration = GetDoubleArrayValue(Index)
                IDSData.Hardware.Needle.Left.DispenseDot.NeedleGap = GetDoubleArrayValue(Index)
                IDSData.Hardware.Needle.Left.DispenseDot.RetractDelay = GetDoubleArrayValue(Index)
                IDSData.Hardware.Needle.Left.DispenseDot.RetractHeight = (GetDoubleArrayValue(Index))
                IDSData.Hardware.Needle.Left.DispenseDot.RetractSpeed = GetDoubleArrayValue(Index)
            End If
            ''' [version 1.1]

            If SWVersion > 1.3 Then
                IDSData.Hardware.Needle.Left.DotDiameter = GetDoubleArrayValue(Index)
            End If

            Index = PatArray.IndexOf("[NEEDLE@RIGHT]")
            IDSData.Hardware.Needle.Right.ArcRadius = GetDoubleArrayValue(Index)
            If IDSData.ParameterID.RecordID = "FactoryDefault" Then
                IDSData.Hardware.Needle.Right.CalibratorPos.X = GetDoubleArrayValue(Index)
                IDSData.Hardware.Needle.Right.CalibratorPos.Y = GetDoubleArrayValue(Index)
                IDSData.Hardware.Needle.Right.CalibratorPos.Z = GetDoubleArrayValue(Index)
            Else
                Index += 3
            End If


            IDSData.Hardware.Needle.Right.HeightApproach = GetIntegerArrayValue(Index)
            IDSData.Hardware.Needle.Right.HeightClearance = GetDoubleArrayValue(Index)
            IDSData.Hardware.Needle.Right.HeightRetract = GetIntegerArrayValue(Index)
            If IDSData.ParameterID.RecordID = "FactoryDefault" Then
                IDSData.Hardware.Needle.Right.NeedleCalibrationPosition.X = GetDoubleArrayValue(Index)
                IDSData.Hardware.Needle.Right.NeedleCalibrationPosition.Y = GetDoubleArrayValue(Index)
                IDSData.Hardware.Needle.Right.NeedleCalibrationPosition.Z = GetDoubleArrayValue(Index)
                IDSData.Hardware.Needle.Right.TouchSensorZPosition = GetDoubleArrayValue(Index)
            Else
                Index += 4
            End If

            IDSData.Hardware.Needle.Right.ArrayDotPos1.X = GetDoubleArrayValue(Index)
            IDSData.Hardware.Needle.Right.ArrayDotPos1.Y = GetDoubleArrayValue(Index)
            IDSData.Hardware.Needle.Right.ArrayDotPos1.Z = GetDoubleArrayValue(Index)

            If SWVersion >= 1.2 Then

                IDSData.Hardware.Needle.Right.ArrayDotPos3.X = GetDoubleArrayValue(Index)
                IDSData.Hardware.Needle.Right.ArrayDotPos3.Y = GetDoubleArrayValue(Index)
                IDSData.Hardware.Needle.Right.ArrayDotPos3.Z = GetDoubleArrayValue(Index)

            End If

            IDSData.Hardware.Needle.Right.CalBackground = GetBooleanArrayValue(Index)
            IDSData.Hardware.Needle.Right.CalBrightness = GetIntegerArrayValue(Index)
            IDSData.Hardware.Needle.Right.CalThreshold = GetIntegerArrayValue(Index)
            IDSData.Hardware.Needle.Right.CalOpen = GetIntegerArrayValue(Index)
            IDSData.Hardware.Needle.Right.CalClose = GetIntegerArrayValue(Index)
            IDSData.Hardware.Needle.Right.CalMaxRadius = GetDoubleArrayValue(Index)
            IDSData.Hardware.Needle.Right.CalMinRadius = GetDoubleArrayValue(Index)
            IDSData.Hardware.Needle.Right.CalRoughness = GetDoubleArrayValue(Index)
            IDSData.Hardware.Needle.Right.CalCompactness = GetDoubleArrayValue(Index)

            ''' [version 1.1]
            If SWVersion >= 1.1 Then
                IDSData.Hardware.Needle.Right.DispenseDot.ApproachHeight = GetDoubleArrayValue(Index)
                IDSData.Hardware.Needle.Right.DispenseDot.ArcRadius = GetDoubleArrayValue(Index)
                IDSData.Hardware.Needle.Right.DispenseDot.ClearanceHeight = GetDoubleArrayValue(Index)
                IDSData.Hardware.Needle.Right.DispenseDot.DispenseDuration = GetDoubleArrayValue(Index)
                IDSData.Hardware.Needle.Right.DispenseDot.NeedleGap = GetDoubleArrayValue(Index)
                IDSData.Hardware.Needle.Right.DispenseDot.RetractDelay = GetDoubleArrayValue(Index)
                IDSData.Hardware.Needle.Right.DispenseDot.RetractHeight = GetDoubleArrayValue(Index)
                IDSData.Hardware.Needle.Right.DispenseDot.RetractSpeed = GetDoubleArrayValue(Index)
            End If
            ''' [version 1.1]

            'regulator
            Index = PatArray.IndexOf("[REGULATOR]")
            IDSData.Hardware.Regulator.Left.Pressure = GetDoubleArrayValue(Index)
            IDSData.Hardware.Regulator.Left.Selection = GetStringArrayValue(Index)
            IDSData.Hardware.Regulator.Left.Vacuum = GetDoubleArrayValue(Index)

            'spc
            Index = PatArray.IndexOf("[SPC]")
            IDSData.Hardware.SPC.ChipFailedAction = GetBooleanArrayValue(Index)
            IDSData.Hardware.SPC.CleanUpInterval = GetIntegerArrayValue(Index)
            IDSData.Hardware.SPC.EnableSPCReport = GetBooleanArrayValue(Index)
            IDSData.Hardware.SPC.FiducialFailedAction = GetBooleanArrayValue(Index)
            IDSData.Hardware.SPC.HeightFailedAction = GetBooleanArrayValue(Index)
            IDSData.Hardware.SPC.QCFailedAction = GetBooleanArrayValue(Index)
            IDSData.Hardware.SPC.ItemsToBeReported = GetStringArrayValue(Index)
            IDSData.Hardware.SPC.ReportFileName = GetStringArrayValue(Index)

            ''' [version 1.1]
            If SWVersion >= 1.1 Then
                IDSData.Hardware.SPC.ProgAuthorName = GetStringArrayValue(Index)
                IDSData.Hardware.SPC.ProgAuthorContact = GetStringArrayValue(Index)
                IDSData.Hardware.SPC.ProductionNote = GetStringArrayValue(Index)
            End If
            ''' [version 1.1]

            'system io
            Index = PatArray.IndexOf("[SYSTEMIO]")
            IDSData.Hardware.SystemIO.ItemArray.Clear()
            Do

                If GetStringArrayValue(Index) = "[SYSTEMIO END]" Then
                    Exit Do
                End If

                IDSData.Hardware.SystemIO.Template = New CIDSData.CIDSHardWare.CIDSSystemIO.CIDSIO
                IDSData.Hardware.SystemIO.Template.CableName = CStr(PatArray.Item(Index))
                IDSData.Hardware.SystemIO.Template.IO = GetStringArrayValue(Index)
                IDSData.Hardware.SystemIO.Template.IOName = GetStringArrayValue(Index)
                IDSData.Hardware.SystemIO.Template.ModuleName = GetStringArrayValue(Index)
                IDSData.Hardware.SystemIO.Template.Status = GetStringArrayValue(Index)
                IDSData.Hardware.SystemIO.Template.Type = GetStringArrayValue(Index)
                Index += 1 ' for the blank line
                IDSData.Hardware.SystemIO.ItemArray.Add(IDSData.Hardware.SystemIO.Template)
            Loop

            '                                       kr                                       '
            ' PAY VERY CAREFUL ATTENTION TO THE ORDER OF PATTERN! 
            ' previously it was 
            ' Needle -> Syringe -> Station1 -> Station2 -> Station3
            ' within each of these stations,
            ' AlarmThreshold -> Name -> CurrentTemp -> HeaterPresentFlag -> OnOffControl
            ' SetTemp -> TempRange.Max -> TempRange.Min

            'heater
            Index = PatArray.IndexOf("[THERMAL]")
            IDSData.Hardware.Thermal.HeaterFeatureEnabled = GetBooleanArrayValue(Index)
            IDSData.Hardware.Thermal.Needle.OperationTemp = GetDoubleArrayValue(Index)
            IDSData.Hardware.Thermal.Needle.StandbyTemp = GetDoubleArrayValue(Index)
            IDSData.Hardware.Thermal.Needle.AlarmThreshold = GetDoubleArrayValue(Index)
            IDSData.Hardware.Thermal.Needle.OnOff = GetBooleanArrayValue(Index)
            IDSData.Hardware.Thermal.Syringe.OperationTemp = GetDoubleArrayValue(Index)
            IDSData.Hardware.Thermal.Syringe.StandbyTemp = GetDoubleArrayValue(Index)
            IDSData.Hardware.Thermal.Syringe.AlarmThreshold = GetDoubleArrayValue(Index)
            IDSData.Hardware.Thermal.Syringe.OnOff = GetBooleanArrayValue(Index)
            IDSData.Hardware.Thermal.Station1.OperationTemp = GetDoubleArrayValue(Index)
            IDSData.Hardware.Thermal.Station1.StandbyTemp = GetDoubleArrayValue(Index)
            IDSData.Hardware.Thermal.Station1.AlarmThreshold = GetDoubleArrayValue(Index)
            IDSData.Hardware.Thermal.Station1.OnOff = GetBooleanArrayValue(Index)
            IDSData.Hardware.Thermal.Station2.OperationTemp = GetDoubleArrayValue(Index)
            IDSData.Hardware.Thermal.Station2.StandbyTemp = GetDoubleArrayValue(Index)
            IDSData.Hardware.Thermal.Station2.AlarmThreshold = GetDoubleArrayValue(Index)
            IDSData.Hardware.Thermal.Station2.OnOff = GetBooleanArrayValue(Index)
            IDSData.Hardware.Thermal.Station3.OperationTemp = GetDoubleArrayValue(Index)
            IDSData.Hardware.Thermal.Station3.StandbyTemp = GetDoubleArrayValue(Index)
            IDSData.Hardware.Thermal.Station3.AlarmThreshold = GetDoubleArrayValue(Index)
            IDSData.Hardware.Thermal.Station3.OnOff = GetBooleanArrayValue(Index)

            'volume
            Index = PatArray.IndexOf("[VOLUME]")
            IDSData.Hardware.Volume.WeighingScaleFlag = GetBooleanArrayValue(Index)

            ' left
            IDSData.Hardware.Volume.Left.OnOff = GetBooleanArrayValue(Index)
            IDSData.Hardware.Volume.Left.CalibrationInterval = GetDoubleArrayValue(Index)
            IDSData.Hardware.Volume.Left.DesiredWeight = GetDoubleArrayValue(Index)
            IDSData.Hardware.Volume.Left.Tolerance = GetDoubleArrayValue(Index)
            IDSData.Hardware.Volume.Left.SetupDispenseDuration = GetDoubleArrayValue(Index)
            IDSData.Hardware.Volume.Left.SetupMaterialAirPressure = GetDoubleArrayValue(Index)
            IDSData.Hardware.Volume.Left.SetupRPM = GetDoubleArrayValue(Index)
            IDSData.Hardware.Volume.Left.NumberOfAttempts = GetDoubleArrayValue(Index)
            IDSData.Hardware.Volume.Left.AdjustedMaterialAirPressure = GetDoubleArrayValue(Index)
            IDSData.Hardware.Volume.Left.AdjustedRPM = GetDoubleArrayValue(Index)
            IDSData.Hardware.Volume.Left.AdjustedDispenseDuration = GetDoubleArrayValue(Index)
            IDSData.Hardware.Volume.Left.PressureStepValue = GetDoubleArrayValue(Index)

            If SWVersion > 1.4 Then
                IDSData.Hardware.Volume.Left.DurationStepValue = GetDoubleArrayValue(Index)
                IDSData.Hardware.Volume.Left.RPMStepValue = GetDoubleArrayValue(Index)
            End If

            IDSData.Hardware.Volume.Left.RetractSpeed = GetDoubleArrayValue(Index)
            IDSData.Hardware.Volume.Left.RetractHeight = GetDoubleArrayValue(Index)
            IDSData.Hardware.Volume.Left.RetractDelay = GetDoubleArrayValue(Index)

            If SWVersion > 1.5 Then
                IDSData.Hardware.Volume.Left.PulseOnDuration = GetDoubleArrayValue(Index)
                IDSData.Hardware.Volume.Left.PulseOffDuration = GetDoubleArrayValue(Index)
                IDSData.Hardware.Volume.Left.JettingNoOfDispense = GetDoubleArrayValue(Index)
            End If

            ' right
            IDSData.Hardware.Volume.Right.OnOff = GetBooleanArrayValue(Index)
            IDSData.Hardware.Volume.Right.CalibrationInterval = GetDoubleArrayValue(Index)
            IDSData.Hardware.Volume.Right.DesiredWeight = GetDoubleArrayValue(Index)
            IDSData.Hardware.Volume.Right.Tolerance = GetDoubleArrayValue(Index)
            IDSData.Hardware.Volume.Right.SetupDispenseDuration = GetDoubleArrayValue(Index)
            IDSData.Hardware.Volume.Right.SetupMaterialAirPressure = GetDoubleArrayValue(Index)
            IDSData.Hardware.Volume.Right.SetupRPM = GetDoubleArrayValue(Index)
            IDSData.Hardware.Volume.Right.NumberOfAttempts = GetDoubleArrayValue(Index)
            IDSData.Hardware.Volume.Right.AdjustedMaterialAirPressure = GetDoubleArrayValue(Index)
            IDSData.Hardware.Volume.Right.AdjustedRPM = GetDoubleArrayValue(Index)
            IDSData.Hardware.Volume.Right.AdjustedDispenseDuration = GetDoubleArrayValue(Index)
            IDSData.Hardware.Volume.Right.PressureStepValue = GetDoubleArrayValue(Index)
            If SWVersion > 1.4 Then
                IDSData.Hardware.Volume.Right.DurationStepValue = GetDoubleArrayValue(Index)
                IDSData.Hardware.Volume.Right.RPMStepValue = GetDoubleArrayValue(Index)
            End If
            IDSData.Hardware.Volume.Right.RetractSpeed = GetDoubleArrayValue(Index)
            IDSData.Hardware.Volume.Right.RetractHeight = GetDoubleArrayValue(Index)
            IDSData.Hardware.Volume.Right.RetractDelay = GetDoubleArrayValue(Index)

            If SWVersion > 1.5 Then
                IDSData.Hardware.Volume.Right.PulseOnDuration = GetDoubleArrayValue(Index)
                IDSData.Hardware.Volume.Right.PulseOffDuration = GetDoubleArrayValue(Index)
                IDSData.Hardware.Volume.Right.JettingNoOfDispense = GetDoubleArrayValue(Index)
            End If


            'pattern
            Index = PatArray.IndexOf("[PATTERNSETTING]")
            IDSData.Execution.Setting.ArcTemplateDir = GetStringArrayValue(Index)
            IDSData.Execution.Setting.ChipEdgeTemplateDir = GetStringArrayValue(Index)
            IDSData.Execution.Setting.CircleTemplateDir = GetStringArrayValue(Index)
            IDSData.Execution.Setting.CurrentPatternName = GetStringArrayValue(Index)
            IDSData.Execution.Setting.DotTemplateDir = GetStringArrayValue(Index)
            IDSData.Execution.Setting.FilledCircleTemplateDir = GetStringArrayValue(Index)
            IDSData.Execution.Setting.FilledRecTemplateDir = GetStringArrayValue(Index)
            IDSData.Execution.Setting.LineTemplateDir = GetStringArrayValue(Index)
            IDSData.Execution.Setting.PatternDir = GetStringArrayValue(Index)
            IDSData.Execution.Setting.RecTemplateDir = GetStringArrayValue(Index)

            If SWVersion > 1.7 Then
                If IDSData.ParameterID.RecordID = "FactoryDefault" Then
                    IDSData.Execution.Setting.DefaultFileToOpen = GetStringArrayValue(Index)
                End If
            End If


            '''''''''''''''PATTERNSERVICE '''''''''''''''''''''''
            Index = PatArray.IndexOf("[PATTERNS ELEMENT]")
            IDSData.Execution.Prog.Clear()
            Do
                If GetStringArrayValue(Index) = "[PATTERNS END]" Then
                    Exit Do
                End If
                IDSData.Execution.Pattern = New CIDSData.CIDSPatternExecution.CIDSTemplate
                IDSData.Execution.Pattern.Item = GetIntegerArrayValue(Index)
                IDSData.Execution.Pattern.Name = GetStringArrayValue(Index)
                IDSData.Execution.Pattern.ApproachDispHeight = GetDoubleArrayValue(Index)
                IDSData.Execution.Pattern.ArcRadius = GetDoubleArrayValue(Index)
                IDSData.Execution.Pattern.ClearanceHeight = GetDoubleArrayValue(Index)
                IDSData.Execution.Pattern.DetailingDistance = GetDoubleArrayValue(Index)
                IDSData.Execution.Pattern.DispenseDuration = GetIntegerArrayValue(Index)
                IDSData.Execution.Pattern.DispenseFlag = GetStringArrayValue(Index)
                IDSData.Execution.Pattern.EdgeClearance = GetDoubleArrayValue(Index)
                IDSData.Execution.Pattern.FilledHeight = GetDoubleArrayValue(Index)
                IDSData.Execution.Pattern.Needle = GetStringArrayValue(Index)
                IDSData.Execution.Pattern.NeedleGap = GetDoubleArrayValue(Index)
                IDSData.Execution.Pattern.Pitch = GetDoubleArrayValue(Index)
                IDSData.Execution.Pattern.Pos1.X = GetDoubleArrayValue(Index)
                IDSData.Execution.Pattern.Pos1.Y = GetDoubleArrayValue(Index)
                IDSData.Execution.Pattern.Pos1.Z = GetDoubleArrayValue(Index)
                IDSData.Execution.Pattern.Pos2.X = GetDoubleArrayValue(Index)
                IDSData.Execution.Pattern.Pos2.Y = GetDoubleArrayValue(Index)
                IDSData.Execution.Pattern.Pos2.Z = GetDoubleArrayValue(Index)
                IDSData.Execution.Pattern.pos3.X = GetDoubleArrayValue(Index)
                IDSData.Execution.Pattern.pos3.Y = GetDoubleArrayValue(Index)
                IDSData.Execution.Pattern.pos3.Z = GetDoubleArrayValue(Index)
                IDSData.Execution.Pattern.RetractDelay = GetIntegerArrayValue(Index)
                IDSData.Execution.Pattern.RetractHeight = GetDoubleArrayValue(Index)
                IDSData.Execution.Pattern.RetractSpeed = GetDoubleArrayValue(Index)
                IDSData.Execution.Pattern.RotatingAngle = GetIntegerArrayValue(Index)
                IDSData.Execution.Pattern.SpiralFlag = GetStringArrayValue(Index)
                IDSData.Execution.Pattern.SubFileName = GetStringArrayValue(Index)
                IDSData.Execution.Pattern.TravelDelay = GetIntegerArrayValue(Index)
                IDSData.Execution.Pattern.TravelSpeed = GetDoubleArrayValue(Index)
                IDSData.Execution.Pattern.RetractHeight = GetDoubleArrayValue(Index)
                IDSData.Execution.Pattern.IONumber = GetIntegerArrayValue(Index)
                Index += 1 ' for the blank

                IDSData.Execution.Prog.Add(IDSData.Execution.Pattern)

            Loop

        Catch ex As Exception
            IDSData.MsgErr = "Reading Error found at << " & CStr(PatDisplayArray.Item(Index)) & " >> " + vbCrLf + vbCrLf + ex.ToString
            MessageBox.Show(IDSData.MsgErr)
            Return False
        End Try
        Return True

    End Function
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'to extract substrings from string
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Friend Function Splitstring(ByRef str As String) As String()
        Dim delimStr As String = "~="
        Dim delimiter As Char() = delimStr.ToCharArray()
        Return str.Split(delimiter)
    End Function

#End Region

#Region " Conversion of Cxxx with error checking "
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' convert object to double with error message check
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Public Function CDouble(ByRef X As Object) As Double

        If IsDBNull(X) = True Then
            Return 0
        Else
            Try
                Return CDbl(X)
            Catch ex As Exception
                MessageBox.Show(ex.ToString, "Application going to close")
                Return 0
            End Try
        End If

    End Function
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' convert object to integer with error message check
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Public Function CInteger(ByRef X As Object) As Integer

        If IsDBNull(X) = True Then
            Return 0
        Else
            Try
                Return CInt(X)
            Catch ex As Exception
                MessageBox.Show(ex.ToString, "Application going to close")
                Return 0
            End Try
        End If

    End Function
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' convert object to string with error message check    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Public Function CString(ByRef X As Object) As String

        If IsDBNull(X) = True Then
            Return ""
        Else
            Try
                Return CStr(X)
            Catch ex As Exception
                MessageBox.Show(ex.ToString, "Application going to close")
                Return ""
            End Try
        End If

    End Function
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' convert object to boolean with error message check
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Public Function CBoolean(ByRef X As Object) As Boolean

        If IsDBNull(X) = True Then
            Return False
        Else

            Try
                Return CBool(X)
            Catch ex As Exception
                MessageBox.Show(ex.ToString, "Application going to close")
                Return False
            End Try

        End If

    End Function

#End Region

#End Region

#Region "Access DataBase"

    '   declare the standard data component for database access
    '   the relationship of data tables is drawn in filename
    '   data relationship.jpg 
    Friend Builder As OleDbCommandBuilder
    Friend Adapter As OleDbDataAdapter = New OleDbDataAdapter
    Friend Reader As OleDbDataReader
    Public Conn As OleDbConnection
    Friend OleDbCom As OleDbCommand
    Friend Const TIMEOUT_VALUE As Integer = 0

    'Group and its associates ( SQL statement )
    Private SQLGroupTable As String = "SELECT * FROM GroupTable"
    Private SQLGroupUserTable As String = "SELECT * FROM GroupUserTable"
    Private SQLGroupPrivilegeTable As String = "SELECT * FROM GroupPrivilegeTable"
    Private SQLGroupHardwareTable As String = "SELECT * FROM GroupHardwareTable"

    'System Configuration   ( SQL statement )
    Private SQLSystemConfigureTable As String = "SELECT * FROM SystemConfigureTable"

    'Privilege   ( SQL statement )
    Private SQLPrivilegeTable As String = "SELECT * FROM PrivilegeTable"

    'User   ( SQL statement )
    Private SQLUserTable As String = "SELECT * FROM UserTable"

    'Hardware   ( SQL statement )
    Private SQLTableHardware As String = "SELECT * FROM TableHardware ORDER BY ID"
    Private SQLTableCamera As String = "SELECT * FROM TableCamera ORDER BY ID"
    Private SQLTableConveyor As String = "SELECT * FROM TableConveyor ORDER BY ID"
    Private SQLTableDispenser As String = "SELECT * FROM TableDispenser ORDER BY ID"
    Private SQLTableGantry As String = "SELECT * FROM TableGantry ORDER BY ID"
    Private SQLTableLightBox As String = "SELECT * FROM TableLightBox ORDER BY ID"
    Private SQLTableRegulator As String = "SELECT * FROM TableRegulator ORDER BY ID"
    Private SQLTableThermal As String = "SELECT * FROM TableThermal ORDER BY ID"
    Private SQLTableHeightSensor As String = "SELECT * FROM TableHeightSensor  ORDER BY ID"
    Private SQLTableVolume As String = "SELECT * FROM TableVolume ORDER BY ID"
    Private SQLTableZhead As String = "SELECT * FROM TableZhead ORDER BY ID"
    Private SQLTableNeedle As String = "SELECT * FROM TableNeedle ORDER BY ID"
    Private SQLTableNeedleDotArray As String = "SELECT * FROM TableNeedleDotArray ORDER BY NeedleID"
    Private SQLTableNeedleParameterArray As String = "SELECT * FROM TableNeedleParameterArray ORDER BY NeedleID"

    Private SQLTableSystemIO As String = "SELECT * FROM TableSystemIO ORDER BY ID, IOName"
    Private SQLTableSPC As String = "SELECT * FROM TableSPC"

    ' Execution and Pattern Programming    ( SQL statement ) - not used
    Private SQLTablePattern As String = "SELECT * FROM TablePattern ORDER BY ID, item"
    Friend SQLTableExecution As String = "SELECT * FROM TableExecution ORDER BY ID"



#Region " Administration "

    'load data to dataset and set primary key - system configuration table
    Private Sub LoadSystemConfiguration()

        Dim TablePk(1) As DataColumn    'define primary key
        Dim LocalTable As DataTable     'define temp working table
        Try
            Adapter.Fill(DS, "SystemConfigureTable")        'load data from database to the dataset
            LocalTable = DS.Tables("SystemConfigureTable")  'pass address from ds to datatable
            TablePk(0) = LocalTable.Columns("Hardware")     'set primary key for the column
            LocalTable.PrimaryKey = TablePk
        Catch ex As Exception
            MessageBox.Show(ex.ToString, "Error!", MessageBoxButtons.OK)
        End Try
    End Sub
    'load data to dataset and set primary key - user table
    Private Sub UserTablePriKey()

        Dim TablePk(1) As DataColumn    'define primary key
        Dim LocalTable As DataTable     'define temp working table

        Try
            Adapter.Fill(DS, "UserTable")                   'load data from database to the dataset
            LocalTable = DS.Tables("UserTable")             'pass address from ds to datatable
            TablePk(0) = LocalTable.Columns("UserID")       'set primary key for the column
            LocalTable.PrimaryKey = TablePk

        Catch ex As Exception
            MessageBox.Show(ex.ToString, "Error!", MessageBoxButtons.OK)
        End Try
    End Sub
    'load data to dataset and set primary key - group table
    Private Sub GroupTablePriKey()

        Dim TablePk(1) As DataColumn    'define primary key
        Dim LocalTable As DataTable     'define temp working table

        Try
            Adapter.Fill(DS, "GroupTable")                   'load data from database to the dataset
            LocalTable = DS.Tables("GroupTable")             'pass address from ds to datatable
            DS.Tables("GroupTable").TableName = "GroupTable" 'set tablename
            TablePk(0) = LocalTable.Columns("GroupID")       'set primary key for the column
            LocalTable.PrimaryKey = TablePk
            'LocalTable.Columns("GroupID").ReadOnly = True

        Catch ex As Exception
            MessageBox.Show(ex.ToString, "Error!", MessageBoxButtons.OK)
        End Try
    End Sub
    'load data to dataset and set primary key - privilege table
    Private Sub PrivilegeTablePriKey()

        Dim TablePk(2) As DataColumn    'define primary key
        Dim LocalTable As DataTable     'define temp working table

        Try
            Adapter.Fill(DS, "PrivilegeTable")                          'load data from databaset to the dataset
            LocalTable = DS.Tables("PrivilegeTable")                    'pass address from ds to datatable
            DS.Tables("PrivilegeTable").TableName = "PrivilegeTable"    'set tablename
            TablePk(0) = LocalTable.Columns("PrivilegeID")              'set primary key for the column
            LocalTable.PrimaryKey = TablePk
            ' LocalTable.Columns("PrivilegeID").ReadOnly = True

        Catch ex As Exception
            MessageBox.Show(ex.ToString, "Error!", MessageBoxButtons.OK)
        End Try
    End Sub
    'load data to dataset and set primary key - Groupuser table
    Private Sub GroupUserTablePriKey()

        Dim TablePk(2) As DataColumn    'define primary key
        Dim LocalTable As DataTable     'define temp working table

        Try
            Adapter.Fill(DS, "GroupUserTable")                          'load data from databaset to the dataset
            LocalTable = DS.Tables("GroupUserTable")                    'pass address from ds to datatable
            DS.Tables("GroupUserTable").TableName = "GroupUserTable"    'set tablename
            TablePk(0) = LocalTable.Columns("GroupID")                  'set primary key for the column
            TablePk(1) = LocalTable.Columns("UserID")                   'set primary key for the column
            LocalTable.PrimaryKey = TablePk

        Catch ex As Exception
            MessageBox.Show(ex.ToString, "Error!", MessageBoxButtons.OK)
        End Try
    End Sub
    'load data to dataset and set primary key - group privilege table
    Private Sub GroupPrivilegeTablePriKey()

        Dim TablePk(2) As DataColumn    'define primary key
        Dim LocalTable As DataTable     'define temp working table

        Try
            Adapter.Fill(DS, "GroupPrivilegeTable")                             'load data from databaset to the dataset
            LocalTable = DS.Tables("GroupPrivilegeTable")                       'pass address from ds to datatable
            DS.Tables("GroupPrivilegeTable").TableName = "GroupPrivilegeTable"  'set tablename
            TablePk(0) = LocalTable.Columns("GroupID")                          'set primary key for the column
            TablePk(1) = LocalTable.Columns("PrivilegeID")                      'set primary key for the column
            LocalTable.PrimaryKey = TablePk

        Catch ex As Exception
            MessageBox.Show(ex.ToString, "Error!", MessageBoxButtons.OK)
        End Try
    End Sub

    'load data to dataset and set primary key - group hardware table
    Private Sub GroupHardwareTablePriKey()

        Dim TablePk(2) As DataColumn    'define primary key
        Dim LocalTable As DataTable     'define temp working table

        Try
            Adapter.Fill(DS, "GroupHardwareTable")                              'load data from databaset to the dataset
            LocalTable = DS.Tables("GroupHardwareTable")                        'pass address from ds to datatable
            DS.Tables("GroupHardwareTable").TableName = "GroupHardwareTable"    'set tablename
            TablePk(0) = LocalTable.Columns("GroupID")                          'set primary key for the column
            TablePk(1) = LocalTable.Columns("HardwareID")                       'set primary key for the column
            LocalTable.PrimaryKey = TablePk

        Catch ex As Exception
            MessageBox.Show(ex.ToString, "Error!", MessageBoxButtons.OK)
        End Try
    End Sub
#End Region

#Region " HardWare and Execution & Pattern Tables & others "

    'load data to dataset and set primary key - subtables from hardware table
    Private Sub SetPrimaryKey(ByRef TableName As String)

        Dim TablePk(2) As DataColumn    'define primary key
        Dim LocalTable As DataTable     'define temp working table

        Try
            Adapter.Fill(DS, TableName)                             'load data from databaset to the dataset
            LocalTable = DS.Tables(TableName)                       'pass address from ds to datatable
            DS.Tables(TableName).TableName = TableName              'set tablename

            If TableName = "TableSystemIO" Then                     'set primary key for the selected column
                TablePk(0) = LocalTable.Columns("CableName")

            ElseIf TableName <> "TableSystemIO" Then

                TablePk(0) = LocalTable.Columns("ID")
                If TableName = "TablePattern" Then
                    TablePk(1) = LocalTable.Columns("Item")
                ElseIf TableName = "TableThermal" Then
                    TablePk(1) = LocalTable.Columns("Controller")
                End If

            End If
            LocalTable.PrimaryKey = TablePk


        Catch ex As Exception
            MessageBox.Show(ex.ToString, "Error!", MessageBoxButtons.OK)
        End Try
    End Sub

#End Region



    ' Load data to the dataset
    Public Function DataLoad() As Boolean

        Try
            Conn.Open()                     'connection on
            DS.CaseSensitive = True         'case sensitive is on
            DS.Reset()                      'reset DS

            '''''   load data from database to the dataset ''''''''''''''''''''''''
            'Group table
            OleDbCom = New OleDbCommand(SQLGroupTable + " ORDER BY GroupID", Conn)      'define SQL command
            OleDbCom.CommandTimeout = TIMEOUT_VALUE                                     'command timeout value
            Adapter.SelectCommand = OleDbCom                                            'assign command to data adapter
            GroupTablePriKey()                                                          'fill data from Database to Ds table 

            'User table
            OleDbCom = New OleDbCommand(SQLUserTable + " ORDER BY UserID", Conn)
            OleDbCom.CommandTimeout = TIMEOUT_VALUE
            Adapter.SelectCommand = OleDbCom
            UserTablePriKey()

            'privilege table
            OleDbCom = New OleDbCommand(SQLPrivilegeTable + " ORDER BY PrivilegeID", Conn)
            OleDbCom.CommandTimeout = TIMEOUT_VALUE
            Adapter.SelectCommand = OleDbCom
            PrivilegeTablePriKey()

            'system configuration table
            OleDbCom = New OleDbCommand(SQLSystemConfigureTable, Conn)
            OleDbCom.CommandTimeout = TIMEOUT_VALUE
            Adapter.SelectCommand = OleDbCom
            LoadSystemConfiguration()

            'GroupUser table
            OleDbCom = New OleDbCommand(SQLGroupUserTable + " ORDER BY GroupID, UserID", Conn)
            OleDbCom.CommandTimeout = TIMEOUT_VALUE
            Adapter.SelectCommand = OleDbCom
            GroupUserTablePriKey()

            'groupPrivilege table
            OleDbCom = New OleDbCommand(SQLGroupPrivilegeTable + " ORDER BY GroupID", Conn)
            OleDbCom.CommandTimeout = TIMEOUT_VALUE
            Adapter.SelectCommand = OleDbCom
            GroupPrivilegeTablePriKey()

            'Group hardware table
            OleDbCom = New OleDbCommand(SQLGroupHardwareTable + " ORDER BY GroupID", Conn)
            OleDbCom.CommandTimeout = TIMEOUT_VALUE
            Adapter.SelectCommand = OleDbCom
            GroupHardwareTablePriKey()

            ''''''''''''''
            '   Hardware
            ''''''''''''''

            'TableHardware
            OleDbCom = New OleDbCommand(SQLTableHardware, Conn)
            OleDbCom.CommandTimeout = TIMEOUT_VALUE
            Adapter.SelectCommand = OleDbCom
            Call SetPrimaryKey("TableHardware")

            'TableCamera
            OleDbCom = New OleDbCommand(SQLTableCamera, Conn)
            OleDbCom.CommandTimeout = TIMEOUT_VALUE
            Adapter.SelectCommand = OleDbCom
            Call SetPrimaryKey("TableCamera")

            'TableConveyor
            OleDbCom = New OleDbCommand(SQLTableConveyor, Conn)
            OleDbCom.CommandTimeout = TIMEOUT_VALUE
            Adapter.SelectCommand = OleDbCom
            Call SetPrimaryKey("TableConveyor")

            'TableDispenser
            OleDbCom = New OleDbCommand(SQLTableDispenser, Conn)
            OleDbCom.CommandTimeout = TIMEOUT_VALUE
            Adapter.SelectCommand = OleDbCom
            Call SetPrimaryKey("TableDispenser")

            'TableGantry
            OleDbCom = New OleDbCommand(SQLTableGantry, Conn)
            OleDbCom.CommandTimeout = TIMEOUT_VALUE
            Adapter.SelectCommand = OleDbCom
            Call SetPrimaryKey("TableGantry")

            'TableHeightSensor
            OleDbCom = New OleDbCommand(SQLTableHeightSensor, Conn)
            OleDbCom.CommandTimeout = TIMEOUT_VALUE
            Adapter.SelectCommand = OleDbCom
            Call SetPrimaryKey("TableHeightSensor")

            'TableLightBox
            OleDbCom = New OleDbCommand(SQLTableLightBox, Conn)
            OleDbCom.CommandTimeout = TIMEOUT_VALUE
            Adapter.SelectCommand = OleDbCom
            Call SetPrimaryKey("TableLightBox")

            'TableRegulator
            OleDbCom = New OleDbCommand(SQLTableRegulator, Conn)
            OleDbCom.CommandTimeout = TIMEOUT_VALUE
            Adapter.SelectCommand = OleDbCom
            Call SetPrimaryKey("TableRegulator")

            'TableThermal
            OleDbCom = New OleDbCommand(SQLTableThermal, Conn)
            OleDbCom.CommandTimeout = TIMEOUT_VALUE
            Adapter.SelectCommand = OleDbCom
            Call SetPrimaryKey("TableThermal")

            'TableVolume
            OleDbCom = New OleDbCommand(SQLTableVolume, Conn)
            OleDbCom.CommandTimeout = TIMEOUT_VALUE
            Adapter.SelectCommand = OleDbCom
            Call SetPrimaryKey("TableVolume")

            'TableZhead
            OleDbCom = New OleDbCommand(SQLTableZhead, Conn)
            OleDbCom.CommandTimeout = TIMEOUT_VALUE
            Adapter.SelectCommand = OleDbCom
            Call SetPrimaryKey("TableZhead")

            'TableNeedle
            OleDbCom = New OleDbCommand(SQLTableNeedle, Conn)
            OleDbCom.CommandTimeout = TIMEOUT_VALUE
            Adapter.SelectCommand = OleDbCom
            Call SetPrimaryKey("TableNeedle")

            'TableNeedleDot
            OleDbCom = New OleDbCommand(SQLTableNeedleDotArray, Conn)
            OleDbCom.CommandTimeout = TIMEOUT_VALUE
            Adapter.SelectCommand = OleDbCom
            Call SetPrimaryKey("TableNeedleDotArray")

            'TableNeedleParameter
            OleDbCom = New OleDbCommand(SQLTableNeedleParameterArray, Conn)
            OleDbCom.CommandTimeout = TIMEOUT_VALUE
            Adapter.SelectCommand = OleDbCom
            Call SetPrimaryKey("TableNeedleParameterArray")

            'TableSystemIO
            OleDbCom = New OleDbCommand(SQLTableSystemIO, Conn)
            OleDbCom.CommandTimeout = TIMEOUT_VALUE
            Adapter.SelectCommand = OleDbCom
            Call SetPrimaryKey("TableSystemIO")

            'TableSPC
            OleDbCom = New OleDbCommand(SQLTableSPC, Conn)
            OleDbCom.CommandTimeout = TIMEOUT_VALUE
            Adapter.SelectCommand = OleDbCom
            Call SetPrimaryKey("TableSPC")

            '   execution and pattern
            'TablePattern
            OleDbCom = New OleDbCommand(SQLTablePattern, Conn)
            OleDbCom.CommandTimeout = TIMEOUT_VALUE
            Adapter.SelectCommand = OleDbCom
            Call SetPrimaryKey("TablePattern")

            'TableExecution
            OleDbCom = New OleDbCommand(SQLTableExecution, Conn)
            OleDbCom.CommandTimeout = TIMEOUT_VALUE
            Adapter.SelectCommand = OleDbCom
            Call SetPrimaryKey("TableExecution")

            '''  Establish relationship between primary and foreign tables      ''

            Dim DsRelation As DataRelation
            '''''''''Below are the code for having UserTable as the primary table, others as the foreign tables
            Dim UserTableIDCol As DataColumn = DS.Tables("UserTable").Columns("UserID")                     'declare userid of Usertable as primary col
            Dim GroupUserTableUserIDCol As DataColumn = DS.Tables("GroupUserTable").Columns("UserID")       'declare userid of Groupusertable as foreign col
            DsRelation = New DataRelation("Groups", UserTableIDCol, GroupUserTableUserIDCol)                'establish their relationship as Group
            DS.Tables("GroupUserTable").ParentRelations.Add(DsRelation)                                     'add it to the ds

            '''''''''Below are the code for having PrivilegesTable as the primary table, others as the foreign tables
            Dim PrivilegeTableIDCol As DataColumn = DS.Tables("PrivilegeTable").Columns("PrivilegeID")
            Dim GroupPrivilegeTablePrivilegeIDCol As DataColumn = DS.Tables("GroupPrivilegeTable").Columns("PrivilegeID")
            DsRelation = New DataRelation("ToGroups", PrivilegeTableIDCol, GroupPrivilegeTablePrivilegeIDCol)
            DS.Tables("GroupPrivilegeTable").ParentRelations.Add(DsRelation)

            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '''''''''Below are the code for having GroupTable as the primary table, others as the foreign tables
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Dim GroupTableIDCol As DataColumn = DS.Tables("GroupTable").Columns("GroupID")
            Dim GroupUserTableGroupIDCol As DataColumn = DS.Tables("GroupUserTable").Columns("GroupID")
            Dim GroupPrivilegeTableGroupIDCol As DataColumn = DS.Tables("GroupPrivilegeTable").Columns("GroupID")
            Dim GroupHardwareTableGroupIDCol As DataColumn = DS.Tables("GroupHardwareTable").Columns("GroupID")

            DsRelation = New DataRelation("Users", GroupTableIDCol, GroupUserTableGroupIDCol)
            DS.Tables("GroupUserTable").ParentRelations.Add(DsRelation)
            DsRelation = New DataRelation("Privileges", GroupTableIDCol, GroupPrivilegeTableGroupIDCol)
            DS.Tables("GroupPrivilegeTable").ParentRelations.Add(DsRelation)
            DsRelation = New DataRelation("Hardwares", GroupTableIDCol, GroupHardwareTableGroupIDCol)
            DS.Tables("GroupHardwareTable").ParentRelations.Add(DsRelation)


            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '''''''''Below are the code for having Table Hardware as the primary table, others as the foreign tables
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

            Dim TableHardwareIDCol As DataColumn = DS.Tables("TableHardware").Columns("ID")
            Dim GroupHardwareTableIDCol As DataColumn = DS.Tables("GroupHardwareTable").Columns("HardwareID")
            Dim TableCameraIDCol As DataColumn = DS.Tables("TableCamera").Columns("ID")
            Dim TableConveyorIDCol As DataColumn = DS.Tables("TableConveyor").Columns("ID")
            Dim TableDispenserIDCol As DataColumn = DS.Tables("TableDispenser").Columns("ID")
            Dim TableGantryIDCol As DataColumn = DS.Tables("TableGantry").Columns("ID")
            Dim TableHeightSensorIDCol As DataColumn = DS.Tables("TableHeightSensor").Columns("ID")
            Dim TableLightBoxIDCol As DataColumn = DS.Tables("TableLightBox").Columns("ID")
            Dim TableRegulatorIDCol As DataColumn = DS.Tables("TableRegulator").Columns("ID")
            Dim TableThermalIDCol As DataColumn = DS.Tables("TableThermal").Columns("ID")
            Dim TableVolumeIDCol As DataColumn = DS.Tables("TableVolume").Columns("ID")
            Dim TableZheadIDCol As DataColumn = DS.Tables("TableZhead").Columns("ID")
            Dim TableNeedleIDCol As DataColumn = DS.Tables("TableNeedle").Columns("ID")
            Dim TableSystemIOIDCol As DataColumn = DS.Tables("TableSystemIO").Columns("ID")
            Dim TableSPCIDCol As DataColumn = DS.Tables("TableSPC").Columns("ID")


            DsRelation = New DataRelation("AtGroups", TableHardwareIDCol, GroupHardwareTableIDCol)
            DS.Tables("GroupHardwareTable").ParentRelations.Add(DsRelation)

            DsRelation = New DataRelation("Camera", TableHardwareIDCol, TableCameraIDCol)
            DS.Tables("TableCamera").ParentRelations.Add(DsRelation)
            DsRelation = New DataRelation("Conveyor", TableHardwareIDCol, TableConveyorIDCol)
            DS.Tables("TableConveyor").ParentRelations.Add(DsRelation)
            DsRelation = New DataRelation("Dispenser", TableHardwareIDCol, TableDispenserIDCol)
            DS.Tables("TableDispenser").ParentRelations.Add(DsRelation)
            DsRelation = New DataRelation("Gantry", TableHardwareIDCol, TableGantryIDCol)
            DS.Tables("TableGantry").ParentRelations.Add(DsRelation)
            DsRelation = New DataRelation("HeightSensor", TableHardwareIDCol, TableHeightSensorIDCol)
            DS.Tables("TableHeightSensor").ParentRelations.Add(DsRelation)
            DsRelation = New DataRelation("LightBox", TableHardwareIDCol, TableLightBoxIDCol)
            DS.Tables("TableLightBox").ParentRelations.Add(DsRelation)
            DsRelation = New DataRelation("Regulator", TableHardwareIDCol, TableRegulatorIDCol)
            DS.Tables("TableRegulator").ParentRelations.Add(DsRelation)
            DsRelation = New DataRelation("Thermal", TableHardwareIDCol, TableThermalIDCol)
            DS.Tables("TableThermal").ParentRelations.Add(DsRelation)
            DsRelation = New DataRelation("Volume", TableHardwareIDCol, TableVolumeIDCol)
            DS.Tables("TableVolume").ParentRelations.Add(DsRelation)
            DsRelation = New DataRelation("Zhead", TableHardwareIDCol, TableZheadIDCol)
            DS.Tables("TableZhead").ParentRelations.Add(DsRelation)
            DsRelation = New DataRelation("Needle", TableHardwareIDCol, TableNeedleIDCol)
            DS.Tables("TableNeedle").ParentRelations.Add(DsRelation)

            ' DsRelation = New DataRelation("SystemIO", TableHardwareIDCol, TableSystemIOIDCol)
            ' DS.Tables("TableSystemIO").ParentRelations.Add(DsRelation)
            DsRelation = New DataRelation("SPC", TableHardwareIDCol, TableSPCIDCol)
            DS.Tables("TableSPC").ParentRelations.Add(DsRelation)


            'Below is the code of having Table Needle as the primary table, others as the foreign tables
            Dim TableNeedleDotArrayNeedleID As DataColumn = DS.Tables("TableNeedleDotArray").Columns("NeedleID")
            Dim TableNeedleParameterArrayNeedleID As DataColumn = DS.Tables("TableNeedleParameterArray").Columns("NeedleID")

            DsRelation = New DataRelation("DotArray", TableNeedleIDCol, TableNeedleDotArrayNeedleID)
            DS.Tables("TableNeedleDotArray").ParentRelations.Add(DsRelation)
            DsRelation = New DataRelation("ParameterArray", TableNeedleIDCol, TableNeedleParameterArrayNeedleID)
            DS.Tables("TableNeedleParameterArray").ParentRelations.Add(DsRelation)

        Catch ex As Exception
            MessageBox.Show(ex.ToString, "Error!", MessageBoxButtons.OK)
            Conn.Close()
            Return False
        End Try
        Conn.Close()
        Return True
    End Function

    'update changes from dataset to the database
    Public Function UpdateData(ByRef SQLStr As String, ByRef TableStr As String) As Boolean

        Try
            Conn.Open()                                                     'open connection
            Dim OleDbCom As OleDbCommand = New OleDbCommand(SQLStr, Conn)   'declare command
            OleDbCom.CommandTimeout = TIMEOUT_VALUE                         'command timeout value
            Adapter.SelectCommand = OleDbCom                                'command assign to data adapter 

            Builder = New OleDbCommandBuilder(Adapter)                      'send command to database by commandbuilder
            Dim DsChanges As DataSet = DS.GetChanges()                      'get changes from ds

            If Not DsChanges Is Nothing Then
                Adapter.Update(DsChanges, TableStr)                        ' if there are changes at ds, update it to the adapter (database)                             
                DS.AcceptChanges()                                          'ds accept changes
            End If

        Catch ex As Exception                                               'if error
            Conn.Close()
            MessageBox.Show(ex.ToString)
            Return False
        End Try
        Conn.Close()                                                        'close con                

        Return True
    End Function

#End Region

End Module