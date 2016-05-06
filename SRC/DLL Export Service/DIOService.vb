Imports System.Threading

Public Class DIOService

    Public Sub TriggerUpstream()
        IDS.Devices.DIO.DIO_SetOutputBit(1, 1, True)
        Thread.Sleep(40)
        IDS.Devices.DIO.DIO_SetOutputBit(1, 1, False)
    End Sub

    Public Sub TriggerDownstream()
        IDS.Devices.DIO.DIO_SetOutputBit(1, 3, True)
        Thread.Sleep(40)
        IDS.Devices.DIO.DIO_SetOutputBit(1, 3, False)
    End Sub

    Public Sub Close_Tower_Lamp()
        Off_Green_Tower_Lamp()
        Off_Yellow_Tower_Lamp()
        Off_Red_Tower_Lamp()
    End Sub

    Public Sub Off_Red_Tower_Lamp()
        IDS.Devices.DIO.DIO_SetOutputBit(0, 3, False)
    End Sub

    Public Sub Off_Yellow_Tower_Lamp()
        IDS.Devices.DIO.DIO_SetOutputBit(0, 5, False)
    End Sub

    Public Sub Off_Green_Tower_Lamp()
        IDS.Devices.DIO.DIO_SetOutputBit(0, 7, False)
    End Sub

    Public Sub On_Red_Tower_Lamp()
        IDS.Devices.DIO.DIO_SetOutputBit(0, 3, True)
    End Sub

    Public Sub On_Yellow_Tower_Lamp()
        IDS.Devices.DIO.DIO_SetOutputBit(0, 5, True)
    End Sub

    Public Sub On_Green_Tower_Lamp()
        IDS.Devices.DIO.DIO_SetOutputBit(0, 7, True)
    End Sub

    Public Sub Off_Siren()
        IDS.Devices.DIO.DIO_SetOutputBit(1, 2, False)
    End Sub

    Public Sub On_Siren()
        IDS.Devices.DIO.DIO_SetOutputBit(1, 2, True)
    End Sub

End Class
