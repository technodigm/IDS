Public Class CIDSDriverThermal
    Public FrmHeater As New Heater
End Class

Public Module SerialCommunicationModule
    Friend Declare Sub Sleep Lib "kernel32" Alias "Sleep" (ByVal dwMilliseconds As Long)  'delay function
End Module

Public Module Module1
    Public start_time As DateTime
    Public stop_time As DateTime
    Public elapsed_time As TimeSpan

    'weighting = COM 1
    'conveyor = COM 2
    'heater = COM 9 (USB)
    'laser = COM 5

    Public WeighingPort As Short = 1
    Public ConveyorPort As Short = 2
    Public DispenserPort As Short = 4
    Public LaserPort As Short = 5
    Public HeaterPort As Short = 9 'usb serial

    Public Sub ExceptionDisplay(ByVal ex As Exception)
        MsgBox(ex.ToString)
        Dim x As Integer = 1 + 1
        x += 2
    End Sub

End Module