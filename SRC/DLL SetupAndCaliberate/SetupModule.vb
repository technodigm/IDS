Imports DLL_Export_Service
Imports DLL_DataManager
Imports DLL_Export_Device_Motor

Public Class CIDSSetup

End Class

Public Module Module1

    'use these to track time-outs in do-while loops 
    Friend start_time As DateTime
    Friend stop_time As DateTime
    Friend elapsed_time As TimeSpan

    Public Declare Function GetInputState Lib "user32" () As Int32
    Public Declare Sub Sleep Lib "kernel32" Alias "Sleep" (ByVal dwMilliseconds As Long)  'delay function
    Public Para As New DLL_DataManager.CIDSData.CIDSParameterID

    Public string_result As String

    'placeholders for moving
    Public position(2) As Double 'for absolute moves
    Public distance(2) As Double 'for relative moves
    Public TableValues(5) As Double 'for reading from table
    Public XYZOrder() As Integer = {0, 1, 2}
    Public result As Boolean
    Public x As Double = 0
    Public y As Double = 0
    Public z As Double = 0
    Public offset_x As Double = 0
    Public offset_y As Double = 0
    Public offset_z As Double = 0

    'system setup
    Public m_Tri As CIDSService.CIDSServiceDevices.CIDSMotor = m_Tri.Instance
    Public MySysIO As New SystemIO
    Public MyGantrySetup As New GantrySetup
    Public MyLaser As New LaserCalibration
    Public MyNeedleCalibrationSetup1 As New NeedleCalibrationSetup1
    Public MyHardwareCommunicationSetup As New HardwareCommunicationSetup
    Public MySetup As New Setup

    'process setup
    Public MyDispenserSettings As New DispenserSettings
    Public MyConveyorSettings As New ConveyorSettings
    Public MyEventSettings As New EventSettings
    Public MySPCLogging As New SPCLogging
    Public MyGantrySettings As New GantrySettings
    Public MyHeaterSettings As New HeaterSettings
    Public MyNeedleCalibrationSettings As New NeedleCalibrationSettings
    Public MyVolumeCalibrationSettings As New VolumeCalibrationSettings
    Public MySettings As New Settings
    Public MyPSInformation As New PSInformation

    'outside of this DLL
    Public Conveyor As CIDSService.CIDSServiceDevices.CIDSServiceConveyor = Conveyor.Instance
    Public Weighting_Scale As CIDSService.CIDSServiceDevices.CIDSServiceWeighting = Weighting_Scale.Instance
    Public Heater As CIDSService.CIDSServiceDevices.CIDSServiceThermal = Heater.Instance
    Public Laser As CIDSService.CIDSServiceDevices.CIDSServiceLaser = Laser.Instance
    Public Dispenser As CIDSService.CIDSServiceDevices.CIDSServiceDispenser = Dispenser.Instance
    Public Vision As CIDSService.CIDSServiceDevices.CIDSServiceVision = Vision.Minstance

    'needs to be fixed. ids.data.hardware.heightsensor.laser.offsetoos.X/Y is WRONG.
    Public LaserOffX As Double = 74.668
    Public LaserOffY As Double = -0.553

    Public LeftNeedleOffsetX As Double = IDS.Data.Hardware.Needle.Left.CalibratorPos.X - IDS.Data.Hardware.Needle.Left.NeedleCalibrationPosition.X
    Public LeftNeedleOffsetY As Double = IDS.Data.Hardware.Needle.Left.CalibratorPos.Y - IDS.Data.Hardware.Needle.Left.NeedleCalibrationPosition.Y
    Public RightNeedleOffsetX As Double = IDS.Data.Hardware.Needle.Right.CalibratorPos.X - IDS.Data.Hardware.Needle.Right.NeedleCalibrationPosition.X
    Public RightNeedleOffsetY As Double = IDS.Data.Hardware.Needle.Right.CalibratorPos.Y - IDS.Data.Hardware.Needle.Right.NeedleCalibrationPosition.Y

    Friend WithEvents T1 As Timer = IDS.T1

    Public Sub ExceptionDisplay(ByVal ex As Exception)
        MsgBox(ex.ToString)
        Dim x As Integer = 1 + 1
        x += 2
    End Sub

    Public Sub MySleep(ByVal duration As Integer)
        Sleep(duration)
    End Sub

    Public Sub SetServiceSpeed()
        m_Tri.Set_XY_Speed(IDS.Data.Hardware.Gantry.ServiceXYSpeed)
        m_Tri.Set_Z_Speed(IDS.Data.Hardware.Gantry.ServiceZSpeed)
    End Sub

    Public Sub T1_Tick(ByVal sender As Object, ByVal e As System.EventArgs) Handles T1.Tick  '250 ms

        'list of errors are:

        'Dispense Alarms
        '1007201 Pot Life Alarm
        '1007202 Low Material Level

        'SPC Reports
        '1007401 Production Starts             
        '1007402 Board comes
        '1007403 Board goes
        '1007404 Pause
        '1007405 Resume
        '1007406 Needle Calibration
        '1007407 Volume Calibration
        '1007408 Board Failed (partially)
        '1007409 Board Failed (totally)
        '1007410 Production End

        'Errors
        '1004201 Volume Calibration (connection)
        '1004202 Volume Calibration (bad reading)

        'Conveyor errors
        '1003201 Width Adjustment Failed    
        '1003202 Lifter 1 Blocked
        '1003203 Lifter 2 Blocked
        '1003204 Lifter 3 Blocked
        '1003205 Retrieve Timeout
        '1003206 Stage 1 Travel Timeout
        '1003207 Release Travel Timeout
        '1003208 Stage 3 Travel Timeout
        '1003209 Board Not Aligned
        '1003210 PLC Communication Error
        '1003211 PLC No Signal
        '1003212 Low Incoming Air Pressure

        'Thermal Errors
        '1002201 Temperature Wrong 1
        '1002202 Temperature Wrong 2
        '1002204 Temperature Wrong 3
        '1002204 Temperature Wrong 4
        '1002205 Temperature Wrong 5
        '1002211 Thermal Com Error 1
        '1002212 Thermal Com Error 2
        '1002213 Thermal Com Error 3
        '1002214 Thermal Com Error 4
        '1002215 Thermal Com Error 5
        '1002221 Temperature Ramping (1 of the 5)
        '1002222 Temperature Ramping 2 
        '1002223 Temperature Ramping 3
        '1002224 Temperature Ramping 4
        '1002225 Temperature Ramping 5

        'Board check errors
        '1006201 Fiducial failed
        '1006202 Height Point failed
        '1006203 Chip Finding failed
        '1006204 QC Check failed

        'E-stop 
        '1006205 Robot Event

        'Left Buttons
        '6  IGNORE          + 60 000
        '7  CONTINUE        + 70 000
        '12 RESET           + 120 000
        '1  OK              + 10 000

        'Right Buttons
        '14 WAIT            + 140 000
        '8  CHANGE SYRINGE  + 80 000
        '5  SKIP CHIP       + 50 000
        '9  OFF HEATER      + 90 000
        '3  STOP            + 30 000

        'Middle Buttons
        '4  SKIP PATTERN    + 40 000
        '2  SKIP            + 20 000
        '10 STOP WAITING    + 100 000
        '13 RELEASE         + 130 000
        '15 REDO            + 150 000

        StopErrorCheck()
        MyVolumeCalibrationSettings.Weighting_T1_Tick()   'weighting machine

        'disable for exhibition
        'MyHeaterSettings.Thermal_T1_Tick()
        MyConveyorSettings.Regulator_T1_Tick()
        MyConveyorSettings.Pressure_T1_Tick() 'low pressure
        'MyConveyorSettings.e_stop_T1_Tick()
        'MyConveyorSettings.Conveyor_T1_Tick() 'conveyor
        StartErrorCheck()

    End Sub

    Public Function StopErrorCheck()
        T1.Stop()
    End Function

    Public Function StartErrorCheck()
        T1.Start()
    End Function

    Public Sub OnLaser()
        IDS.Devices.Vision.SetLaser(False) 'turn on
    End Sub

    Public Sub OffLaser()
        IDS.Devices.Vision.SetLaser(True) 'turn off
    End Sub

    Public Sub TraceDoEvents()
        Application.DoEvents()
    End Sub

    Public Function TrioMotionCalibrating()

        m_Tri.SetCalibrationFlag()
        m_Tri.RunTrioMotionProgram("CALIBRATIONS", 3)
        Do
            'm_Tri.GetIDSState()
            TraceDoEvents()
        Loop Until m_Tri.EStopActivated Or m_Tri.Calibrating

        While m_Tri.Calibrating And Not m_Tri.EStopActivated
            'm_Tri.GetIDSState()
            Sleep(1)
            TraceDoEvents()
        End While

        If m_Tri.EStopActivated Or m_Tri.MachineHoming Then Return False
        Return True

    End Function

    Public CurrentControl As System.Windows.Forms.Control

    Public Sub AddPanel(ByVal ParentControl As System.windows.forms.Control, ByVal ChildPanel As Panel)

        CurrentControl = ChildPanel

        ParentControl.Controls.Add(ChildPanel)
        ChildPanel.BringToFront()
        ChildPanel.Location = New Point(0, 0)
        ChildPanel.Show()

    End Sub

    Public Sub RemovePanel(ByVal ChildPanel As Panel)
        MySetup.PanelRight.Controls.Remove(ChildPanel)
        MySettings.PanelRight.Controls.Remove(ChildPanel)
    End Sub

    Public Function MachineRunning() As Boolean

        If m_Tri.MachineRunning Or m_Tri.MachineHoming Or m_Tri.Calibrating Then Return True
        Return False

    End Function

End Module