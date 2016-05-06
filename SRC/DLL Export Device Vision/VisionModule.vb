
Public Class CIDSVision
    Private Shared m_instance As CIDSVision
    Public Lightbox As New DeviceLightbox

    Public Shared ReadOnly Property Instance() As CIDSVision
        Get
            If m_instance Is Nothing Then
                m_instance = New CIDSVision
            Else
            End If
            Return m_instance
        End Get

    End Property

    Public Sub callFunctionHere()
        Dim x = MyVisionCounter
    End Sub

    Public Sub SetLaser(ByVal val As Boolean)
        Lightbox.SetLaser(val)
    End Sub

End Class

Public Class CIDSRegulator
    Public FrmRegulator As New DeviceRegulators
End Class

Public Class CIDSDIO
    Public DIO As New DeviceDIO

    'lsgoh
    Function SetOutputBit(ByRef PortNumber As Integer, ByRef BitNumber As Integer, ByRef SetValue As Boolean) As Boolean
        Return DIO.SetOutputBit(PortNumber, BitNumber, SetValue)
    End Function
    'lsgoh
    Public Function CheckInputByte(ByRef PortNumber As Integer, ByRef ReadByte As Byte) As Boolean
        Return DIO.CheckInputByte(PortNumber, ReadByte)
    End Function

End Class

Public Module Module1

    'utility functions
    Public Declare Sub Sleep Lib "kernel32" Alias "Sleep" (ByVal dwMilliseconds As Long)
    Public WaitRefresh As Boolean = True
    Public MyVisionCounter As Integer = 0
    Public Declare Function GetInputState Lib "user32" () As Int32

    Public Sub ExceptionDisplay(ByVal ex As Exception)
        MsgBox(ex.ToString)
        Dim x As Integer = 1 + 1
        x += 2
    End Sub


    'for the user to see the results
    Public Sub DelayClearingImage()
        Dim aaa As Integer
        For aaa = 1 To 10
            Sleep(10)
            TraceDoEvents()
        Next aaa
    End Sub

    Public Sub TraceDoEvents()
        Application.DoEvents()
    End Sub

    Public DisplayWidth As Integer = 768
    Public DisplayHeight As Integer = 576
    Public DisplayCenterXPosition As Integer = DisplayWidth / 2
    Public DisplayCenterYPosition As Integer = DisplayHeight / 2

#Region "MIL utilities"
    Public Sub FreeIfAllocated(ByVal control As AxMatrox.ActiveMIL.BlobAnalysis.AxMBlobAnalysis)
        If control.IsAllocated Then
            control.Free()
        End If
    End Sub

    Public Sub FreeIfAllocated(ByVal control As AxMatrox.ActiveMIL.AxMImage)
        If control.IsAllocated Then
            control.Free()
        End If
    End Sub

    Public Sub FreeIfAllocated(ByVal control As AxMatrox.ActiveMIL.ImageProcessing.AxMImageProcessing)
        If control.IsAllocated Then
            control.Free()
        End If
    End Sub

    Public Sub FreeIfAllocated(ByVal control As AxMatrox.ActiveMIL.AxMDisplay)
        If control.IsAllocated Then
            control.Free()
        End If
    End Sub

    Public Sub FreeIfAllocated(ByVal control As AxMatrox.ActiveMIL.AxMSystem)
        If control.IsAllocated Then
            control.Free()
        End If
    End Sub

    Public Sub FreeIfAllocated(ByVal control As AxMatrox.ActiveMIL.AxMGraphicContext)
        If control.IsAllocated Then
            control.Free()
        End If
    End Sub

    Public Sub FreeIfAllocated(ByVal control As AxMatrox.ActiveMIL.PatternMatching.AxMPatternMatching)
        If control.IsAllocated Then
            control.Free()
        End If
    End Sub

    Public Sub FreeIfAllocated(ByVal control As AxMatrox.ActiveMIL.Measurement.AxMMeasurement)
        If control.IsAllocated Then
            control.Free()
        End If
    End Sub

    Public Sub FreeIfAllocated(ByVal control As AxMatrox.ActiveMIL.Calibration.AxMCalibration)
        If control.IsAllocated Then
            control.Free()
        End If
    End Sub
#End Region

End Module
