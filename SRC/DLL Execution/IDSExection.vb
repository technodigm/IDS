Option Explicit On 

Imports Microsoft.VisualBasic
Imports DLL_DataManager
Imports DLL_Export_Service
Imports Microsoft.DirectX
'Imports Microsoft.DirectX.DirectInput



 '''''''''''''''
'
'   Please Place Functions and attributs to be call by other DLL modules here
'
 ''''''''''''''

Public Class IDSExecution
    Inherits IDSService
    Public m_Pattern As New CIDSPattern
    Public m_ItemLUT As New CIDSItemsLUT
    Public m_Command As New CIDSCommand
    Public m_File As New CIDSFileHandler
    Public m_Undo As New CIDSUndo

    Public Sub New()
        LoadDefaultPatternParam()
    End Sub

    Public Function LoadDefaultPatternParam() As Integer
        Dim Templet As New CIDSData.CIDSPatternExecution.CIDSTemplate
        Dim Para As New CIDSData.CIDSParameterID

        Dim file As New CIDSFileHandler
        'should get from data_manager
        Dim dir As String = "C:\IDS\Appication Main\IDSPattnTemplet\"

        Dim ret As Integer = file.Open(dir + "default", "ptp")
        If (ret <> 0) Then
            'Error Message
            Return ret
        End If
        Dim eof As String = ""
        Dim LineStr(5) As String

        Do Until eof.Equals("EOF")
            eof = file.Read("ptp")
            If eof <> Nothing Or eof <> "" Then
                LineStr = eof.Split("=")
                If LineStr.Length = 2 Then
                    LineStr(0) = LineStr(0).Trim(" ")
                    LineStr(0) = LineStr(0).ToUpper()
                    LineStr(1) = LineStr(1).Trim(" ")
                    LineStr(1) = LineStr(1).ToUpper()
                    Dim value As Double = 0.0
                    Select Case LineStr(0)
                        Case "ONOFF"
                            If LineStr(1) = "TRUE" Then
                                Templet.DispenseFlag = True
                            Else
                                Templet.DispenseFlag = False
                            End If
                        Case "GAP"
                            value = CDbl(LineStr(1))
                            Templet.NeedleGap = value
                        Case "DTIME"
                            value = CDbl(LineStr(1))
                            Templet.DispenseDuration = value
                        Case "AHEIGHT"
                            value = CDbl(LineStr(1))
                            Templet.ApproachDispHeight = value
                        Case "TDELAY"
                            value = CDbl(LineStr(1))
                            Templet.TravelDelay = value
                        Case "TSPEED"
                            value = CDbl(LineStr(1))
                            Templet.TravelSpeed = value
                        Case "DTDIST"
                            value = CDbl(LineStr(1))
                            Templet.DetailingDistance = value
                        Case "RTDELAY"
                            value = CDbl(LineStr(1))
                            Templet.RetractDelay = value
                        Case "RTSPEED"
                            value = CDbl(LineStr(1))
                            Templet.RetractSpeed = value
                        Case "RTHEIGHT"
                            value = CDbl(LineStr(1))
                            Templet.RetractHeight = value
                        Case "CLHEIGHT"
                            value = CDbl(LineStr(1))
                            Templet.ClearanceHeight = value
                        Case "ARADIUS"
                            value = CDbl(LineStr(1))
                            Templet.ArcRadius = value
                        Case "PITCH"
                            value = CDbl(LineStr(1))
                            Templet.Pitch = value
                        Case "FHEIGHT"
                            value = CDbl(LineStr(1))
                            Templet.FilledHeight = value
                        Case "RANGLE"
                            value = CDbl(LineStr(1))
                            Templet.RotatingAngle = value
                        Case "SUBNAME"
                            Templet.SubFileName = LineStr(1)
                        Case "ECLEAR"
                            value = CDbl(LineStr(1))
                            Templet.EdgeClearance = value
                        Case "SPINOUT"
                            If LineStr(1) = "TRUE" Then
                                Templet.SpiralFlag = True
                            Else
                                Templet.SpiralFlag = False
                            End If
                        Case "NEEDLE"
                            value = CDbl(LineStr(1))
                            Templet.Needle = value

                        Case Else

                    End Select
                End If
            End If
        Loop
        'update data to data manager
        'Templet.SaveLocalData(Para)
        file.Close()
    End Function
End Class

Public Class IDSService

End Class


