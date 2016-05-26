Imports System.IO
Imports DLL_DataManager
Imports DLL_Export_Device_Vision
Imports DLL_Export_Device_Motor
Imports DLL_Serial_Communication

Public Class CIDSService

    Private Declare Sub Sleep Lib "kernel32" Alias "Sleep" (ByVal dwMilliseconds As Long)  'delay function
    Public Shared FrmExecution As FormExecution = FormExecution.Instance
    Public WithEvents T1 As New Timer
    Public __ID As String = "0"
    Public newOpen As Boolean = False                   'Flag for new and open files

    Public Function StopErrorCheck()
        T1.Stop()
    End Function

    Public Function StartErrorCheck()
        T1.Start()
    End Function

#Region "Using Inherits from base class method to trace all devices events "

    Public Devices As New CIDSServiceDevices
    Public Class CIDSServiceDevices

#Region "Conveyor Part"
        Public Shadows Conveyor As New CIDSServiceConveyor
        Public Class CIDSServiceConveyor
            Inherits Conveyor

            Shared No_Signal_time As Byte
            Shared com_error_time As Byte

            Private Shared m_instance As CIDSServiceConveyor

            Public Shared ReadOnly Property Instance() As CIDSServiceConveyor
                Get
                    If m_instance Is Nothing Then
                        m_instance = New CIDSServiceConveyor
                    Else
                        If m_instance.IsDisposed Then
                            m_instance = New CIDSServiceConveyor
                        End If
                    End If
                    m_instance.BringToFront()
                    Return m_instance
                End Get

            End Property

        End Class

#End Region

#Region "Weighting Machine"
        Public Shadows Weighting As New CIDSServiceWeighting
        Public Class CIDSServiceWeighting
            Inherits Weighting_Scale

            Private Shared m_instance As CIDSServiceWeighting

            Public Shared ReadOnly Property Instance() As CIDSServiceWeighting
                Get
                    If m_instance Is Nothing Then
                        m_instance = New CIDSServiceWeighting
                    Else
                        If m_instance.IsDisposed Then
                            m_instance = New CIDSServiceWeighting
                        End If
                    End If
                    m_instance.BringToFront()
                    Return m_instance
                End Get

            End Property

        End Class
#End Region

#Region "Thermal Part"
        Public Shadows Thermal As New CIDSServiceThermal
        Public Class CIDSServiceThermal
            Inherits CIDSDriverThermal

            Public Shared out_range_cycle As Integer = 0
            Public Shared ramping_cycle As Integer = 0
            Public Shared ramp_error_come As Boolean = False
            Shared Thermal_No_Signal_time As Byte
            Private Shared m_instance As CIDSServiceThermal

            Public Shared ReadOnly Property Instance() As CIDSServiceThermal
                Get
                    If m_instance Is Nothing Then
                        m_instance = New CIDSServiceThermal
                    Else
                        If m_instance.FrmHeater.IsDisposed Then
                            m_instance = New CIDSServiceThermal
                        End If
                    End If
                    m_instance.FrmHeater.BringToFront()
                    Return m_instance
                End Get

            End Property

            Public Sub getThermalError(ByRef ID As String, ByVal temperature_alarm As Boolean(), ByVal HeaterEnabled As Boolean(), ByVal ramping As Boolean)
                MyBase.FrmHeater.GetHeaterError(ID, HeaterEnabled)
                'If id = 1002201 Then '''''''Communication Error
                'form_service.DisplayErrorMessage(ID)
                'ElseIf id = 1002202 Then  '''''''Communication No Signal
                '    Thermal_No_Signal_time = Thermal_No_Signal_time + 1
                '    If Thermal_No_Signal_time = 2 Then
                '        Thermal_No_Signal_time = 0
                '        form_service.DisplayErrorMessage(ID)
                '    Else
                '        id = 1002200
                '    End If
                'ElseIf id > 1002205 And id < 1002211 And out_range_cycle = 0 Then
                '    form_service.DisplayErrorMessage(ID)
                'ElseIf id > 1002205 And id < 1002211 And out_range_cycle > 0 Then
                '    id = 1002200
                'ElseIf id > 1002210 And id < 1002216 And ramping_cycle = 0 Then
                '    form_service.DisplayErrorMessage(ID)
                '    ramp_error_come = True
                'ElseIf id > 1002210 And id < 1002216 And ramping_cycle > 0 Then
                '    id = 1002200
                'End If
            End Sub
        End Class
#End Region

#Region "Vision Events "
        Public Shadows Vision As New CIDSServiceVision
        ' track class vision
        Public Class CIDSServiceVision
            Inherits CIDSVision

            Private Shared m_instance As CIDSServiceVision
            Public Shared ReadOnly Property Minstance() As CIDSServiceVision
                Get
                    If m_instance Is Nothing Then
                        m_instance = New CIDSServiceVision
                    Else
                    End If
                    Return m_instance
                End Get
            End Property
        End Class
#End Region

        Public Shadows Regulator As New CIDSServiceRegulator
        ' track class vision
        Public Class CIDSServiceRegulator
            Inherits CIDSRegulator

            Private Shared m_instance As CIDSServiceRegulator
            Public Shared ReadOnly Property Instance() As CIDSServiceRegulator
                Get
                    If m_instance Is Nothing Then
                        m_instance = New CIDSServiceRegulator
                    Else
                    End If
                    Return m_instance
                End Get
            End Property
        End Class

#Region "Regulator Events"
#End Region

#Region "Trace DIO Events"
        Public Shadows DIO As New CIDSServiceDIO
        ' track class vision
        Public Class CIDSServiceDIO
            Inherits CIDSDIO

            'track function callfunctionhere

            Public Shadows Function DIO_SetOutputBit(ByVal PortNumber As Integer, ByVal Bitnumber As Integer, ByVal SetValue As Boolean)

                '' do my SPC event tracing here
                '  then call the base class
                MyBase.DIO.SetOutputBit(PortNumber, Bitnumber, SetValue)

            End Function
        End Class
#End Region

#Region "Dispenser"
        Public Shadows Dispenser As New CIDSServiceDispenser
        Public Class CIDSServiceDispenser
            Inherits Dispenser

            Private Shared m_instance As CIDSServiceDispenser

            Public Shared ReadOnly Property Instance() As CIDSServiceDispenser
                Get
                    If m_instance Is Nothing Then
                        m_instance = New CIDSServiceDispenser
                    Else
                        If m_instance.IsDisposed Then
                            m_instance = New CIDSServiceDispenser
                        End If
                    End If
                    m_instance.BringToFront()
                    Return m_instance
                End Get

            End Property
        End Class
#End Region

#Region "Laser"
        Public Shadows Laser As New CIDSServiceLaser
        Public Class CIDSServiceLaser
            Inherits Laser

            Private Shared m_instance As CIDSServiceLaser

            Public Shared ReadOnly Property Instance() As CIDSServiceLaser
                Get
                    If m_instance Is Nothing Then
                        m_instance = New CIDSServiceLaser
                    Else
                        If m_instance.IsDisposed Then
                            m_instance = New CIDSServiceLaser
                        End If
                    End If
                    m_instance.BringToFront()
                    Return m_instance
                End Get

            End Property
        End Class
#End Region

#Region "Motion Controller"
        Public Class CIDSMotor
            Inherits CIDSTrioController

            Private Shared m_instance As CIDSMotor
            Public Sub New()
                'InitializeComponent();
            End Sub

            Public Shared ReadOnly Property Instance() As CIDSMotor
                Get
                    If m_instance Is Nothing Then
                        m_instance = New CIDSMotor
                    End If
                    m_instance.BringToFront()
                    Return m_instance
                End Get
            End Property
        End Class
#End Region

    End Class
#End Region

#Region "Using Inherits from base class to Track Data Manager events "

    Public Data As New CIDSTraceData
    Public Class CIDSTraceData
        Inherits CIDSData
    End Class
#End Region

End Class

Public Module x
    Private Declare Sub Sleep Lib "kernel32" Alias "Sleep" (ByVal dwMilliseconds As Long)  'delay function
    Public Form_Service As New FormService
    Public Form_Execution As FormExecution = FormExecution.Instance
    Public DIO_Service As New DIOService
    Public IDS As New CIDSService

    'gui instantiation from other modules
    Public m_Tri As CIDSService.CIDSServiceDevices.CIDSMotor = m_Tri.Instance
    Public Conveyor As CIDSService.CIDSServiceDevices.CIDSServiceConveyor = Conveyor.Instance
    Public Weighting_Scale As CIDSService.CIDSServiceDevices.CIDSServiceWeighting = Weighting_Scale.Instance
    Public Heater As CIDSService.CIDSServiceDevices.CIDSServiceThermal = Heater.Instance
    Public Laser As CIDSService.CIDSServiceDevices.CIDSServiceLaser = Laser.Instance
    Public Dispenser As CIDSService.CIDSServiceDevices.CIDSServiceDispenser = Dispenser.Instance
    Public Vision As CIDSService.CIDSServiceDevices.CIDSServiceVision = Vision.Minstance

    Public Sub ExceptionDisplay(ByVal ex As Exception)
        MsgBox(ex.ToString)
    End Sub

End Module