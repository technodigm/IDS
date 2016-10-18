Imports System.IO
Imports System.Xml.Serialization

Public Class VolumeCalPostHandler
    Public Structure VolumeCalPostHandlerParam
        Public dotDiameter As Double
        Public pitch As Double
        'The following is the machine coordinate used to populate
        'the volume dispense dot position which in this case is the
        '2D matrix style in square shape
        Public topLeftX As Double
        Public topLeftY As Double
        Public bottomRightX As Double
        Public bottomRightY As Double
    End Structure

    Public vCParam As VolumeCalPostHandlerParam
    Private currentDot As Dot
    Private currentDotListPost As Integer
    Private dotList As New ArrayList
    Private filePath As String = "C:\TIDS\VolumeCalPostElement.txt"

    Public Sub New()
        currentDot = New Dot
        vCParam.dotDiameter = 2 'mm
        vCParam.pitch = 0.1 'mm
        vCParam.topLeftX = 0.0
        vCParam.topLeftY = 0.0
        vCParam.bottomRightX = 0.0
        vCParam.bottomRightY = 0.0
    End Sub
    Public Sub New(ByVal vCPParam As VolumeCalPostHandlerParam)
        currentDot = New Dot
        vCParam.dotDiameter = vCPParam.dotDiameter
        vCParam.pitch = vCPParam.pitch
        vCParam.topLeftX = vCPParam.topLeftX
        vCParam.topLeftY = vCPParam.topLeftY
        vCParam.bottomRightX = vCPParam.bottomRightX
        vCParam.bottomRightY = vCPParam.bottomRightY
        If Not Directory.Exists("C:\TIDS") Then
            Directory.CreateDirectory("C:\TIDS")
        End If
    End Sub
    'Populate the matrices of the volume calibration dots
    'nxn matrices
    Public Function PopulateNewPost() As Boolean
        Try
            Dim xLength As Double = Math.Abs(vCParam.topLeftX - vCParam.bottomRightX)
            Dim yLength As Double = Math.Abs(vCParam.topLeftY - vCParam.bottomRightY)
            Dim numOfDotPerXLength As Integer = xLength / (vCParam.dotDiameter + vCParam.pitch)
            Dim numOfDotPerYLength As Integer = yLength / (vCParam.dotDiameter + vCParam.pitch)
            Dim i As Integer = 0
            Dim j As Integer = 0
            dotList.Clear()
            Dim objStreamWriter As New StreamWriter(filePath, False)
            For j = 0 To numOfDotPerYLength
                For i = 0 To numOfDotPerXLength
                    dotList.Add(New Dot(j, i))
                    objStreamWriter.WriteLine(j.ToString() + "," + i.ToString() + "," + "0")
                Next
            Next
            objStreamWriter.Close()
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function
    Public Function OldPostFileExist() As Boolean
        Return File.Exists(filePath)
    End Function
    Public Function DeleteCalFile()
        If File.Exists(filePath) Then
            File.Delete(filePath)
        End If
    End Function
    'Get populated post from file
    Public Function GetAllPopulatedPost() As Boolean
        Try
            dotList.Clear()
            Dim sr As StreamReader = New StreamReader(filePath)
            Dim str As String = ""
            Do While sr.Peek() >= 0
                str = sr.ReadLine()
                If Not (str = "") Then
                    Dim strTemp() As String
                    strTemp = str.Split(",")
                    If Not (strTemp.Length = 3) Then
                        Return False
                    Else
                        dotList.Add(New Dot(strTemp(0), strTemp(1), Not (strTemp(2) = "0")))
                    End If
                End If
            Loop
            sr.Close()
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function
    'Return false means all available post were dispensed
    Public Function GetNextPost(ByRef postX As Double, ByRef postY As Double) As Boolean
        Dim i As Integer = -1
        For Each dot As Dot In dotList
            i += 1
            If Not dot.isDispensed Then
                currentDot = dot
                currentDotListPost = i
                postX = vCParam.topLeftX + (dot.column * (vCParam.dotDiameter + vCParam.pitch))
                postY = vCParam.topLeftY - (dot.row * (vCParam.dotDiameter + vCParam.pitch))
                Return True
                Exit For
            End If
        Next
        Return False
    End Function
    'Write to file
    Public Function SetCurrentDotDispensed()
        Dim dt As Dot = dotList.Item(currentDotListPost)
        dt.isDispensed = True
        Dim objStreamWriter As New StreamWriter(filePath, False)
        For Each dot As Dot In dotList
            If dot.isDispensed Then
                objStreamWriter.WriteLine(dot.row.ToString() + "," + dot.column.ToString() + "," + "1")
            Else
                objStreamWriter.WriteLine(dot.row.ToString() + "," + dot.column.ToString() + "," + "0")
            End If
        Next
        objStreamWriter.Close()
    End Function
End Class



Public Class Dot
    Public row As Integer = -1
    Public column As Integer = -1
    Public isDispensed As Boolean = False
    Public Sub New()

    End Sub
    Public Sub New(ByVal row As Integer, ByVal column As Integer)
        Me.row = row
        Me.column = column
        isDispensed = False
    End Sub
    Public Sub New(ByVal row As Integer, ByVal column As Integer, ByVal dipenseStatus As Boolean)
        Me.row = row
        Me.column = column
        isDispensed = dipenseStatus
    End Sub
End Class
