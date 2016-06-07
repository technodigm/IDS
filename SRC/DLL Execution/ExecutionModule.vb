Imports DLL_Export_Service
Imports DLL_Export_Device_Motor
Imports DLL_SetupAndCalibrate
Imports DLL_DataManager
Imports Microsoft.DirectX.DirectInput
Imports System.Threading

Public Class CIDSExe

    Public Function CALLME()
    End Function

    'Setting system setup parameters to execution moudle's global data
    Public Function SetExecutionSetupParam()
        SystemSetupDataRetrieve()
    End Function

End Class

Public Module ExecutionModule

    'threading declarations
    Public PreviousMachineState As String = "Idle"
    Public MachineState As String = "Idle"
    Public ThreadExecution As Threading.Thread
    Public ThreadMonitor As Threading.Thread
    Public ThreadExecutor As Threading.Thread

    Public InitThread As Threading.Thread

    Public VolumeCalibrationRunning As Boolean = False

    'external functions imported from DLLs
    Public Declare Sub MySleep Lib "kernel32" Alias "MySleep" (ByVal dwMilliseconds As Long)
    Public Declare Function InterlockedExchange Lib "kernel32" (ByRef Target As Integer, ByVal Value As Integer) As Integer
    Public Declare Function GetInputState Lib "user32" () As Int32

    'general frequently used variables
    Public position(2) As Double 'for absolute moves
    Public offset(2) As Double 'for offset
    Public distance(2) As Double 'for relative moves

    'downloading variables
    Public OnePageElements As Integer = 1000  '(change also in DISPENSE trio program)
    Public MaxPages As Integer = 30 '15 '5 (change also in DISPENSE trio program)
    Public StartDispPage As Integer = 2 '25 '5 (starting form 2) (no change in DISPENSE trio program)
    Public StartPosition As Integer = 1000 '(please do not change)

    'gui instantiation from other modules
    Public m_Tri As CIDSService.CIDSServiceDevices.CIDSMotor ' = m_Tri.Instance
    Public Conveyor As CIDSService.CIDSServiceDevices.CIDSServiceConveyor = Conveyor.Instance
    Public Weighting_Scale As CIDSService.CIDSServiceDevices.CIDSServiceWeighting = Weighting_Scale.Instance
    Public Heater As CIDSService.CIDSServiceDevices.CIDSServiceThermal = Heater.Instance
    Public Laser As CIDSService.CIDSServiceDevices.CIDSServiceLaser = Laser.Instance
    Public Dispenser As CIDSService.CIDSServiceDevices.CIDSServiceDispenser = Dispenser.Instance
    Public Vision As CIDSService.CIDSServiceDevices.CIDSServiceVision = Vision.Minstance

    'gui instantiation within this module
    Public Production As New FormProduction
    Public Programming As New FormProgramming
    Public IDSExe As New CIDSExe
    Public m_Execution As New IDSExecution
    Private MouseTimer As System.Threading.Timer

    'for z position
    Public SafePosition As Double = 0.0 'original -> Dim SafePosition As Double = gSoftHome(2) + gSystemOrigin(2)

#Region "useful utilities"

    Public Function ProductionMode() As Boolean
        Return (gExeMode = "Operator")
    End Function

    Public Function ProgrammingMode() As Boolean
        Return (gExeMode = "Programmer")
    End Function

    Public Function ShouldLog()
        Return IDS.Data.Hardware.SPC.EnableSPCReport And ProductionMode() And Production.ContinuousMode.Checked
    End Function

    Public Sub ClearDisplay()
        Programming.cross_num.Text = CStr(CInt(Programming.cross_num.Text) + 1)
        Vision.FrmVision.ClearDisplay()
    End Sub

    Public Sub ClearAndDisplayIndicator()
        If ProgrammingMode() Then
            Vision.FrmVision.DisplayIndicator()
        Else
            Vision.FrmVision.ClearDisplay()
        End If
    End Sub
#End Region

#Region "debugging"
    'for debugging
    Public start2 As DateTime
    Public end2 As DateTime
    Public elapsed2 As TimeSpan
    Public start1 As DateTime
    Public end1 As DateTime
    Public elapsed1 As TimeSpan

    Public Sub ExceptionDisplay(ByVal ex As Exception)
        MsgBox(ex.ToString)
        Dim x As Integer = 1 + 1
        x += 2
    End Sub

    Public Sub MySleep(ByVal duration As Integer)
        Sleep(duration)
    End Sub

    Public Sub MyMsgBox(ByVal str As String)
        Try
            MsgBox(str)
        Catch ex As ThreadAbortException
        Catch ex As Exception
            ExceptionDisplay(ex)
        End Try
    End Sub

    Public Function MyMsgBox(ByVal str As String, ByVal buttons As Microsoft.VisualBasic.MsgBoxStyle)
        Try
            Return MsgBox(str, buttons)
        Catch ex As ThreadAbortException
        Catch ex As Exception
            ExceptionDisplay(ex)
        End Try
    End Function

    Public Function MyMsgBox(ByVal str As String, ByVal buttons As Microsoft.VisualBasic.MsgBoxStyle, ByVal title As String)
        Try
            Return MsgBox(str, buttons, title)
        Catch ex As ThreadAbortException
        Catch ex As Exception
            ExceptionDisplay(ex)
        End Try
    End Function

#End Region

#Region "new states"

    Private Delegate Sub UpdateCoordinateDelegate(ByVal XCor As String, ByVal YCor As String, ByVal ZCor As String)
    Private Delegate Sub UpdateTeachPointDelegate(ByVal XCor As Double, ByVal YCor As Double, ByVal ZCor As Double)
    Private Delegate Sub PromptErrorDelegate()
    Private updateCoordinater As UpdateCoordinateDelegate = New UpdateCoordinateDelegate(AddressOf UpdateCoordinate)
    Private updateTeachPosition As UpdateTeachPointDelegate = New UpdateTeachPointDelegate(AddressOf UpdateTeachPoint)

    Private Sub UpdateCoordinate(ByVal XCor As String, ByVal YCor As String, ByVal ZCor As String)

        If ProgrammingMode() Then
            Programming.TextBoxRobotX.Text = "X: " & XCor
            Programming.TextBoxRobotY.Text = "Y: " & YCor
            Programming.TextBoxRobotZ.Text = "Z: " & ZCor
        Else
            Production.TextBoxRobotPos.Text = "X: " & XCor & ", Y: " & YCor & ", Z: " & ZCor
        End If

    End Sub

    Private Sub Callback(ByVal ar As IAsyncResult)
        ar.AsyncState.EndInvoke(ar)
    End Sub

    Private Sub UpdateMachineStatusCallback(ByVal ar As IAsyncResult)
        ar.AsyncState.EndInvoke(ar)
    End Sub

    Private Sub UpdateTeachPoint(ByVal XCor As Double, ByVal YCor As Double, ByVal ZCor As Double)
        Dim colmX, colmY, colmZ As Integer
        Try

            If m_TeachStepNumber = 1 Then
                colmX = gPos1XColumn
                colmY = gPos1YColumn
                colmZ = gPos1ZColumn
            ElseIf m_TeachStepNumber = 2 Then
                colmX = gPos2XColumn
                colmY = gPos2YColumn
                colmZ = gPos2ZColumn
            Else
                colmX = gPos3XColumn
                colmY = gPos3YColumn
                colmZ = gPos3ZColumn
            End If

            If Programming.DeletingRowFromExcel = False And m_PosUpdate = True Then
                SetCellValue(m_Row, colmX, XCor) 'm_CamPos(0) '
                SetCellValue(m_Row, colmY, YCor) 'm_CamPos(1) '
                If Programming.NeedleMode.Checked Or Programming.m_TeachMode = 2 Then 'LEFT Right
                    SetCellValue(m_Row, colmZ, ZCor) 'm_CamPos(1) 
                Else
                    SetCellValue(m_Row, colmZ, 0) 'm_CamPos(1) 
                End If
            End If

        Catch ex As Exception
            ExceptionDisplay(ex)
        End Try
    End Sub

    Private Sub UpdateTeachPointCallback(ByVal ar As IAsyncResult)
        ar.AsyncState.EndInvoke(ar)
    End Sub


    Public Sub GeneratePatternAndRunProgram()

        Dim rtn As Integer

        TraceGCCollect()
        LabelMessage("Carrying out reference checks.")
        m_Execution.m_Command.SetOptimizFlag(0)
        m_Tri.SetMachineRun()
        ChangeButtonState("Running")
        SwitchToTeachCamera()
        OnLaser()
        rtn = m_Execution.m_Command.ReadPattern(Programming.AxSpreadsheetProgramming)
        'to show cross
        Vision.FrmVision.DisplayIndicator()
        OffLaser()

        If rtn = 100 Then
            LabelMessage("Machine stopped.") 'pressed stop while doing reference check
            GoTo resetmachinestate
        ElseIf rtn < 0 Or rtn = 99 Then
            LabelMessage("Empty sheet or error in file.")
            GoTo resetmachinestate
        ElseIf rtn = 98 Then
            LabelMessage("Board rejected.")
            GoTo resetmachinestate
        End If

        LabelMessage("Compiling program.")

        If m_Execution.m_Command.Compile(Programming.m_RunMode) < 0 Then
            LabelMessage(ErrorMessage())
            ' LabelMessage("Compilation error.")
            GoTo resetmachinestate
        End If

        Try
            SetState("Run")
            ThreadExecution = New Threading.Thread(AddressOf DownloadProgram)
            ThreadExecution.Priority = Threading.ThreadPriority.Normal
            ThreadExecution.Start()
        Catch ex As ThreadAbortException
        Catch ex As Exception
            ExceptionDisplay(ex)
        End Try

        Exit Sub

ResetMachineState:
        LockMovementButtons()
        TravelToParkPosition()
        ResetToIdle()

    End Sub

    Public Function DispensingCtrl() As Integer

        Dim Dispensinglist As ArrayList
        Dim cmdBurn As New CIDSPattnBurn

        Try
            If m_Execution.m_Command.CompileStatus > 0 Then

                Dispensinglist = m_Execution.m_Command.DispenseList
                m_Tri.SetMachineRunMode(Programming.m_RunMode)

                If cmdBurn.BurnTable(Dispensinglist) = False Then
                    LabelMessage("Downloading failed or cancelled.")
                    Return -1
                Else
                    SetLampsToRunningMode()
                    LabelMessage("Download finished.")
                End If
            End If
        Catch ex As Exception
        End Try

        Return 0
    End Function

    Public EStop As Boolean = False 'boolean flag for triggering all the status/state updates only once when estop

    Dim MPos(3), off(3), ref(3), readbuf(4), Spos(3), Hpos(3) As Double
    Dim WorkX, WorkY, WorkZ, HomeX, HomeY, HomeZ As Double
    Dim EStopStatus, qcRequest As Integer

    'THREAD
    Public Sub StateMonitor()
        Do
            Try
                'how long the thread waits between updates
                MySleep(50)
                'this is to prevent e-stop window from freezing.
                'if may cause teaching problems.

                m_Tri.GetIDSState()

                'spc report
                If ProductionMode() And (m_Tri.MotionError Or m_Tri.EStopActivated) Then
                    If ProductionMode() Then Production.WriteSPCReport()
                    LabelMessage("Motion controller error. Please restart the motion controller and program.")
                End If

                Dim status_update As String = MyVolumeCalibrationSettings.VolumeCalibrationResult
                If VolumeCalibrationRunning And status_update <> "" Then
                    If ProgrammingMode() Then
                        If Programming.LabelMessege.Text <> status_update Then
                            LabelMessage(status_update)
                        End If
                    Else
                        If Production.LabelMessege.Text <> status_update Then
                            LabelMessage(status_update)
                        End If
                    End If
                End If

                Hpos(0) = m_Tri.XPosition
                Hpos(1) = m_Tri.YPosition
                Hpos(2) = m_Tri.ZPosition

                'Programming.openEventViewer = False
                HardToSys(Hpos, Spos)
                ref(0) = Programming.m_ReferPt(0)
                ref(1) = Programming.m_ReferPt(1)
                ref(2) = Programming.m_ReferPt(2)
                SysToRefer(Spos, Spos, ref)

                If Programming.NeedleMode.Checked Then
                    off(0) = gLeftNeedleOffs(0)
                    off(1) = gLeftNeedleOffs(1)
                    off(2) = -gLeftNeedleOffs(2)
                Else 'vision
                    off(0) = 0.0
                    off(1) = 0.0
                    off(2) = 0.0
                End If
                Spos(0) = Spos(0) + off(0)
                Spos(1) = Spos(1) + off(1)
                Spos(2) = Spos(2) + off(2) + Programming.m_BoardnRefBlkDist

                'here is where the co-ordinates are set. Rx, Ry and Rz are shown on spreadsheet
                'SPos = Work Coordinate System (use when teaching)
                'HPos = Hard Home Coordinate System (use when in basic setup)  

                RoundTo3DecimalPoints(Spos)
                RoundTo3DecimalPoints(Hpos)
                WorkX = Spos(0)
                WorkY = Spos(1)
                WorkZ = Spos(2)
                HomeX = Hpos(0)
                HomeY = Hpos(1)
                HomeZ = Hpos(2)
                Programming.m_MachinePos(0) = Hpos(0)
                Programming.m_MachinePos(1) = Hpos(1)
                Programming.m_MachinePos(2) = Hpos(2)

                'different display for basic setup and program editor (because of station settings)
                If Programming.CurrentMode = "Basic Setup" Then
                    updateCoordinater.BeginInvoke(HomeX, HomeY, HomeZ, AddressOf Callback, updateCoordinater)
                ElseIf Programming.CurrentMode = "Program Editor" Or ProductionMode() Then
                    updateCoordinater.BeginInvoke(WorkX, WorkY, WorkZ, AddressOf Callback, updateCoordinater)
                End If

                If ProgrammingMode() Then
                    If m_PosUpdate = True And Programming.DeletingRowFromExcel = False Then
                        updateTeachPosition.BeginInvoke(WorkX, WorkY, WorkZ, AddressOf UpdateTeachPointCallback, updateTeachPosition)
                    Else
                        If Programming.DeletingRowFromExcel = True Then
                            Programming.DeletingRowFinished = True
                        End If
                    End If
                Else
                End If

            Catch ex As InvalidCastException
            Catch ex As ThreadAbortException
            Catch ex As Exception
                ExceptionDisplay(ex)
            End Try
        Loop
    End Sub


    Public Sub EStopSequence()

        m_Tri.TrioStop()
        If Not ThreadExecution Is Nothing Then ThreadExecution.Abort()
        If VolumeCalibrationRunning = True Then MyVolumeCalibrationSettings.VolumeCalibrationState = "Stopped"

        VolumeCalibrationRunning = False

        SetState("Idle")
        SwitchToTeachCamera()
        LockMovementButtons()
        LabelMessage("Motion controller error. Release e-stop to re-home system.")

        Form_Service.DisplayErrorMessage("E-Stop")
        If ShouldLog() Then Form_Service.LogEventInSPCReport("E-Stop")

        If ProductionMode() Then
            Conveyor.Command("Release")
            'Production.WriteSPCReport()
        End If
    End Sub

    Public Sub SpreadsheetVolumeCalibration()

        m_Tri.ResetVolCalFlag()
        VolumeCalibrationRunning = True
        MyVolumeCalibrationSettings.VolumeCalibrationState = "Running"
        MyVolumeCalibrationSettings.VolumeCalibrationResult = "Doing volume calibration."

        Dim success As Boolean = False
        MyVolumeCalibrationSettings.VolumeCalibrationSetup()
        success = MyVolumeCalibrationSettings.VolumeCalibration(IDS.Data.Hardware.Volume.Left.NumberOfAttempts)
        'LabelMessage("Finished volume calibration.")
        VolumeCalibrationRunning = False
        MyVolumeCalibrationSettings.VolumeCalibrationResult = "Finished volume calibration."

        If success Then
            m_Tri.SetVolCalFinished()
            MyVolumeCalibrationSettings.DownloadAdjustedParameters()
            If ShouldLog() Then Form_Service.LogEventInSPCReport("Volume Calibration Passed")
        Else
            If CheckButtonState() = -1 Then Exit Sub
            Form_Service.DisplayErrorMessage("Volume Calibration Failed")
            WaitUntilErrorMessagesCleared()

            If Form_Service.VolCalPressedContinue Then
                m_Tri.SetVolCalFinished()
                If ShouldLog() Then Form_Service.LogEventInSPCReport("Volume Calibration Failed")
                If ShouldLog() Then Form_Service.LogEventInSPCReport("Board Partial Failure")
                If ShouldLog() Then Form_Service.LogEventInSPCReport("Resume")
            ElseIf Form_Service.VolCalPressedRedo Then
                'If ShouldLog() Then Form_Service.LogEventInSPCReport("Volume Calibration Redo")
                SpreadsheetVolumeCalibration()
            ElseIf Form_Service.VolCalPressedStop Then
                m_Tri.SetVolCalStop()
                LabelMessage("Machine stopped")
                LockMovementButtons()
                TravelToParkPosition()
                ResetToIdle()
                If ShouldLog() Then Form_Service.LogEventInSPCReport("Volume Calibration Failed")
                If ShouldLog() Then Form_Service.LogEventInSPCReport("Board Total Failure")
            Else
                'estop
            End If

        End If

    End Sub

    Public Sub QCcheck()

        Dim vParam As DLL_Export_Device_Vision.QC.QCParam
        Dim readbuf(15) As Double

        SwitchToTeachCamera()
        m_Tri.ResetQCFlag()

        m_Tri.m_TriCtrl.GetTable(210, 15, readbuf)
        vParam._Binarized = readbuf(0)
        If readbuf(1) = 1 Then
            vParam._BlackDot = True
        Else
            vParam._BlackDot = False
        End If
        vParam._Brightness = readbuf(2)
        'Set the brightness before doing the movement. This will affect vision(ActiveX) less.
        Vision.IDSV_SetBrightness(readbuf(2))
        vParam._Close = readbuf(3)
        vParam._Compactness = readbuf(4)
        vParam._MaxArea = readbuf(5)
        vParam._MinArea = readbuf(6)
        vParam._MRegionX = readbuf(7)
        vParam._MRegionY = readbuf(8)
        vParam._MROIx = readbuf(9)
        vParam._MROIy = readbuf(10)
        vParam._Open = readbuf(11)
        vParam._Roughness = readbuf(12)
        vParam._Diameter = readbuf(13)
        vParam._Tolerance = readbuf(14)

        Dim rtn As Boolean = Vision.IDSV_QC(vParam)
        ClearDisplay()
        If Programming.m_RunMode <> 0 Then SwitchToRealTimeCamera()
        If rtn Then
            If ShouldLog() Then Form_Service.LogEventInSPCReport("Dot Size Check Passed")
            m_Tri.SetQCFinished()
        Else
            If CheckButtonState() = -1 Then Exit Sub
            If IDS.Data.Hardware.SPC.QCFailedAction = False And ProductionMode() Then
                Form_Service.LogEventInSPCReport("Dot Size Check Failed")
            ElseIf IDS.Data.Hardware.SPC.QCFailedAction = True And ProductionMode() Then
                Form_Service.LogEventInSPCReport("Dot Size Check Failed")
                Form_Service.DisplayErrorMessage("Dot Size Check Failed")
            ElseIf IDS.Data.Hardware.SPC.QCFailedAction = True And ProgrammingMode() Then
                Form_Service.DisplayErrorMessage("Dot Size Check Failed")
            ElseIf IDS.Data.Hardware.SPC.QCFailedAction = False And ProgrammingMode() Then
                m_Tri.SetQCFinished()
                Exit Sub
            End If

            WaitUntilErrorMessagesCleared()

            If Form_Service.DotSizeCheckStop Then
                If ShouldLog() Then Form_Service.LogEventInSPCReport("Board Total Failure")
                m_Tri.SetQCStop()
                LabelMessage("Machine stopped.")
                LockMovementButtons()
                TravelToParkPosition()
                ResetToIdle()
            ElseIf Form_Service.DotSizeCheckContinue Then
                If ShouldLog() Then Form_Service.LogEventInSPCReport("Board Partial Failure")
                m_Tri.SetQCFinished()
            Else
                'estop
            End If

        End If

    End Sub

    'THREAD
    Public Sub DownloadProgram()
        Try
            Dim rtn As Integer = DispensingCtrl()
            If Not rtn = 0 Then ResetToIdle() 'error message
        Catch ex As ThreadAbortException
        Catch ex As Exception
            ExceptionDisplay(ex)
        End Try

    End Sub

    'THREAD
    'actions in this thread must reflect machine action at any time. 
    Public Sub StateChanger()
        Do
            Try

                Select Case MachineState
                    'run will compile the spreadsheet into an array of values to download into triomotion
                    'downloading of array values is done inside a new thread
                Case "Start"
                        If ProgrammingMode() Then
                            Programming.DispensingMode.Enabled = False
                            Programming.DisableTeachModeSwitching()
                            If VolumeCalibrationRunning = True Then MyVolumeCalibrationSettings.VolumeCalibrationState = "Running"
                            LockMovementButtons()
                            MoveZToSafePosition()
                            GeneratePatternAndRunProgram()
                        Else
                            If Programming.IsFilePresent Then
                                ResetToIdle()
                                LabelMessage("No file loaded.")
                                Exit Select
                            End If
                            LockMovementButtons()
                            If IDS.Data.Hardware.Dispenser.Left.HeadType = "Jetting Valve" Or IDS.Data.Hardware.Dispenser.Left.HeadType = "Auger Valve" Then
                                m_Tri.TurnOn("Material Air")
                            End If
                            If Production.ContinuousMode.Checked = True Then ' continuous run
                                Form_Service.LogEventInSPCReport("Production Start")
                                Production.Conveyor_Check()
                            Else
                                Production.Start()
                            End If
                        End If

                    Case "Run"

                        MySleep(50)
                        TraceDoEvents()
                        If m_Tri.QCRequested Then
                            QCcheck()
                        ElseIf m_Tri.VolCalRequested Then
                            SpreadsheetVolumeCalibration()
                        ElseIf m_Tri.DispensingFinished() Then
                            m_Tri.ResetDispensingFlag()
                            If ProgrammingMode() Then
                                LabelMessage("Dispensing finished.")
                                LockMovementButtons()
                                TravelToParkPosition()
                                ResetToIdle()
                                Programming.NeedleMode.Enabled = True
                                Programming.VisionMode.Enabled = True
                                TraceGCCollect()
                            Else
                                If Production.ContinuousMode.Checked = True Then
                                    Production.DispensingFinishCallback.Invoke() 'synchronous
                                Else
                                    LockMovementButtons()
                                    TravelToParkPosition()
                                    ResetToIdle()
                                End If
                            End If
                        End If

                    Case "Park"
                        LockMovementButtons()
                        TravelToParkPosition()
                        ResetToIdle()

                    Case "Change Syringe"
                        DoChangeSyringe()

                    Case "Volume Calibration"
                        DoVolumeCalibration()

                    Case "Needle Calibration"
                        DoNeedleCalibration()

                    Case "Jogging"
                        LockMovementButtonsOnce()
                        Do
                            MySleep(25)
                            TraceDoEvents()
                        Loop Until Not m_Tri.MachineJogging Or m_Tri.EStopActivated

                    Case "Homing"
                        DoHoming()
                        MySleep(500) 'to wait for ids.getstate in statemonitor, otherwise it will exit
                        'the loop immediately.
                        Do
                            MySleep(25)
                            TraceDoEvents()
                        Loop Until m_Tri.HomingFinished Or m_Tri.EStopActivated

                        ResetToIdle()
                        LabelMessage("Homing finished.")

                    Case "AutoPurge"
                        Production.ProdPurgeClean()

                    Case "Purge", "Clean"
                        For i As Integer = 1 To 5
                            TraceDoEvents()
                            MySleep(100)
                        Next

                    Case "Idle", "Stop", "Pause"

                End Select

            Catch ex_abort As ThreadAbortException
            Catch ex As Exception
                ExceptionDisplay(ex)
            End Try
        Loop
    End Sub

    Private startInit As AutoResetEvent = New AutoResetEvent(False)
    Private abortInit As AutoResetEvent = New AutoResetEvent(False)
    Private waitHandles() As WaitHandle = {abortInit, startInit}
    Public isInited As Boolean = False
    Public isExited As Boolean = False
    Public Sub Init()
        InitThread = New System.Threading.Thread(AddressOf ControllerThread)
        InitThread.Priority = ThreadPriority.Normal
        InitThread.Start()
        startInit.Set()
    End Sub
    Public Sub ExitInit()
        abortInit.Set()
    End Sub
    'Thread
    Public Sub ControllerThread()
        While (1)
            Dim i As Integer = WaitHandle.WaitAny(waitHandles)
            Select Case i
                Case 0
                    'Abort
                    Console.WriteLine("Thread exit #1")
                    abortInit.Reset()
                    isExited = True
                    Return
                Case 1
                    Console.WriteLine("Thread create Motor instance")
                    If m_Tri Is Nothing Then
                        m_Tri = New CIDSService.CIDSServiceDevices.CIDSMotor
                    End If
                    m_Tri.Connect_Controller()
                    isInited = True
                    startInit.Reset()
                    'Init
            End Select
        End While
    End Sub

#End Region

#Region "states"

    Public Function SetState(ByVal str As String) As Boolean

        'If MachineState = str Then
        '    Return False
        'End If

        'check to see whether it is a valid command
        Select Case str
            Case "Start"
                If IsIdle() Then
                    str = "Start"
                ElseIf IsRunning() Then
                    Return False
                ElseIf IsPaused() Then
                    ResumeDispensing() 'this function does not take too long so it's okay to put it here.
                    str = "Run"
                End If
            Case "Run"
                'start -> run or resume -> run. otherwise invalid
                If Not IsStart() Then Return False
            Case "Purge"
                If MachineState = "Purge" Then
                    Return True
                ElseIf IsBusy() Then
                    Return False
                End If
            Case "Clean"
                If MachineState = "Clean" Then
                    Return True
                ElseIf IsBusy() Then
                    Return False
                End If
            Case "Park"
                If IsBusy() Then Return False
            Case "Change Syringe"
                If IsBusy() Then Return False
            Case "Volume Calibration"
                If IsBusy() Then Return False
            Case "Needle Calibration"
                If IsBusy() Then Return False
            Case "Jogging"
            Case "AutoPurge"
            Case "Homing"
                If IsBusy() Then Return False
            Case "Idle"
        End Select

        PreviousMachineState = MachineState
        MachineState = str

        If ProductionMode() Then
            Production.TextBox1.Text = PreviousMachineState
            Production.TextBox2.Text = MachineState
        Else
            Programming.TextBox1.Text = PreviousMachineState
            Programming.TextBox2.Text = MachineState
        End If
        Return True
    End Function
    Public Function WasRunning()
        If PreviousMachineState = "Run" Then Return True
        Return False
    End Function
    Public Function WasIdle()
        If PreviousMachineState = "Idle" Then Return True
        Return False
    End Function
    Public Function WasStart()
        If PreviousMachineState = "Start" Then Return True
        Return False
    End Function
    Public Function IsRunning()
        If MachineState = "Run" Then Return True
        Return False
    End Function
    Public Function IsJogging()
        If MachineState = "Jogging" Then Return True
        Return False
    End Function
    Public Function IsBusy()
        If IsIdle() Then Return False
        Return True
    End Function
    'why would machinestate be "idle" yet calibrating or have a motion error
    Public Function IsIdle()
        If m_Tri Is Nothing Then Return False
        If MachineState = "Idle" And Not m_Tri.Calibrating And Not m_Tri.MotionError Then Return True
        Return False

    End Function
    Public Function IsStart()
        If MachineState = "Start" Then Return True
        Return False
    End Function
    Public Function IsPaused()
        If MachineState = "Pause" Then Return True
        Return False
    End Function
    Public Function IsHoming()
        If MachineState = "Homing" Then Return True
        Return False
    End Function

    Public Function CheckButtonState()

        If IsRunning() Or IsStart() Then
            Return 0
        ElseIf IsPaused() Then

            Do
                MySleep(1)
                TraceDoEvents()
            Loop Until Not IsPaused() Or m_Tri.EStopActivated
            If IsRunning() Or IsStart() Then
                Return 0
            Else
                Return -1
            End If
        Else
            Return -1
        End If

    End Function

    Public Sub ResetToIdle()
        If m_Tri Is Nothing Then Return
        'machinestate
        SetState("Idle")
        m_Tri.SetMachineStop()

        'IO turn off
        m_Tri.TurnOff("Clean")
        m_Tri.TurnOff("Left Needle IO")
        m_Tri.TurnOff("Material Air")

        'camera
        SwitchToTeachCamera()

        'GUI
        SetLampsToReadyMode()
        UnlockMovementButtons()
        ChangeButtonState("Idle")
        If ProgrammingMode() Then
            Programming.DispensingMode.Enabled = True
            Programming.VisionMode.Enabled = True
            Programming.NeedleMode.Enabled = True
        End If

    End Sub

#End Region

#Region "trio motion routines"

    Public Sub DoHoming()

        If m_Tri.m_TriCtrl.IsOpen(0) Then
            LabelMessage("Homing..")
            m_Tri.SetMachineRun()
            SetLampsToRunningMode()
            LockMovementButtons()
            SetState("Homing")
            m_Tri.m_TriCtrl.Execute("RUN SETDATUM")
        End If

    End Sub

    Public Sub MoveZToSafePosition()

        m_Tri.Set_Z_Speed(IDS.Data.Hardware.Gantry.ElementZSpeed)
        If Not m_Tri.Move_Z(SafePosition) Then Exit Sub

    End Sub

    Public Sub TravelToParkPosition()

        'warning! we do not set run mode in this function.
        position(0) = IDS.Data.Hardware.Gantry.ParkPosition.X
        position(1) = IDS.Data.Hardware.Gantry.ParkPosition.Y
        m_Tri.Set_XY_Speed(IDS.Data.Hardware.Gantry.ElementXYSpeed)
        MoveZToSafePosition()
        If Not m_Tri.Move_XY(position) Then Exit Sub

    End Sub

#End Region

#Region "gui"

    Public Function WaitUntilErrorMessagesCleared() As Boolean

        While IDS.__ID > 1006200 And IDS.__ID < 1006300
            If EStop = True Then
                Form_Service.Hide()
                Form_Service.DisplayErrorMessage("E-Stop")
                Return False
            End If
            TraceDoEvents()
        End While
        Return True

    End Function

    Public Sub EnableProgrammingBrightnessToggle()
        Vision.FrmVision.SetBrightness(Programming.ValueBrightness.Value)
        Programming.ValueBrightness.Enabled = True
    End Sub

    Public Sub DisableProgrammingBrightnessToggle()
        Programming.ValueBrightness.Enabled = False
    End Sub

    Public Sub LockButtonsForPurge()
        LockMovementButtons()
        If ProgrammingMode() Then
            Programming.ButtonPurge.Enabled = True
        ElseIf ProductionMode() Then
            Programming.ButtonPurge.Enabled = True
        End If
    End Sub

    Public Sub LockButtonsForClean()
        LockMovementButtons()
        If ProgrammingMode() Then
            Programming.ButtonClean.Enabled = True
        ElseIf ProductionMode() Then
            Programming.ButtonClean.Enabled = True
        End If
    End Sub

    'we add this as a hack to prevent button refreshing when machine state is set to jogging
    Dim button_lock_flag As Boolean = False
    Public Sub LockMovementButtonsOnce()
        If button_lock_flag Then
            Exit Sub
        End If
        button_lock_flag = True
        LockMovementButtons()
    End Sub

    Public Sub LockMovementButtons()
        ChangeButtonState("Disabled")
        If ProgrammingMode() Then
            m_Tri.SteppingButtons.Enabled = False 'stepping buttons
            With Programming
                .ButtonHome.Enabled = False
                .ButtonNeedleCal.Enabled = False
                .ButtonVolCal.Enabled = False
                .ButtonPurge.Enabled = False
                .ButtonClean.Enabled = False
                .DispensingMode.Enabled = False
            End With
        ElseIf ProductionMode() Then
            With Production
                .ButtonHome.Enabled = False
                .ButtonChgSyringe.Enabled = False
                .ButtonClean.Enabled = False
                .ButtonVolCalib.Enabled = False
                .ButtonNdlCalib.Enabled = False
                .ButtonPurge.Enabled = False

                .ConveyorBox.Enabled = False
            End With
        End If
    End Sub

    Public Sub UnlockMovementButtons()
        button_lock_flag = False
        If ProgrammingMode() Then
            m_Tri.SteppingButtons.Enabled = True
            With Programming
                .ButtonHome.Enabled = True
                .ButtonNeedleCal.Enabled = True
                .ButtonVolCal.Enabled = True
                .ButtonPurge.Enabled = True
                .ButtonClean.Enabled = True
                .DispensingMode.Enabled = True
            End With
        ElseIf ProductionMode() Then
            With Production
                .ButtonHome.Enabled = True
                .ButtonChgSyringe.Enabled = True
                .ButtonClean.Enabled = True
                .ButtonVolCalib.Enabled = True
                .ButtonNdlCalib.Enabled = True
                .ButtonPurge.Enabled = True
                .ConveyorBox.Enabled = True
            End With
        End If
    End Sub

    Public Sub LabelMessage(ByVal message As String)
        Try
            If gExeMode <> "Operator" Then
                Programming.LabelMessege.Text = message
                Programming.LabelMessege.Refresh()
            Else
                Production.LabelMessege.Text = message
                Production.LabelMessege.Refresh()
            End If
        Catch ex As Exception
        End Try
    End Sub

    Public Sub TraceGCCollect()
        GC.Collect()
    End Sub

    Public Sub TraceDoEvents()
        Application.DoEvents()
    End Sub

    'machine running -> green
    'machine idle/ready to run -> yellow
    'machine error -> red
    Public Sub SetLampsToRunningMode()
        DIO_Service.Off_Red_Tower_Lamp()
        DIO_Service.On_Green_Tower_Lamp()
        DIO_Service.Off_Yellow_Tower_Lamp()
        If gExeMode <> "Operator" Then
            Programming.PBRed.Hide()
            Programming.PBYellow.Hide()
            Programming.PBGreen.Show()
        Else
            Production.PBRed.Hide()
            Production.PBYellow.Hide()
            Production.PBGreen.Show()
        End If
    End Sub

    Public Sub SetLampsToReadyMode()
        DIO_Service.Off_Red_Tower_Lamp()
        DIO_Service.Off_Green_Tower_Lamp()
        DIO_Service.On_Yellow_Tower_Lamp()
        If gExeMode <> "Operator" Then
            Programming.PBRed.Hide()
            Programming.PBYellow.Show()
            Programming.PBGreen.Hide()
        Else
            Production.PBRed.Hide()
            Production.PBYellow.Show()
            Production.PBGreen.Hide()
        End If
    End Sub

    Public Sub SetLampsToErrorMode()
        DIO_Service.On_Red_Tower_Lamp()
        DIO_Service.Off_Green_Tower_Lamp()
        DIO_Service.Off_Yellow_Tower_Lamp()
        If gExeMode <> "Operator" Then
            Programming.PBRed.Show()
            Programming.PBYellow.Hide()
            Programming.PBGreen.Hide()
        Else
            Production.PBRed.Show()
            Production.PBYellow.Hide()
            Production.PBGreen.Hide()
        End If
    End Sub

    Public Sub ChangeButtonState(ByVal state As String) 'kr

        If state = "Running" Then 'running so pause or stop only
            Programming.TBOperation.Buttons(0).Enabled = False
            Programming.TBOperation.Buttons(1).Enabled = True
            Programming.TBOperation.Buttons(2).Enabled = True
            Production.TBOperation.Buttons(0).Enabled = False
            Production.TBOperation.Buttons(1).Enabled = True
            Production.TBOperation.Buttons(2).Enabled = True
        ElseIf state = "Idle" Then 'idle so run only
            Programming.TBOperation.Buttons(0).Enabled = True
            Programming.TBOperation.Buttons(1).Enabled = False
            Programming.TBOperation.Buttons(2).Enabled = False
            Production.TBOperation.Buttons(0).Enabled = True
            Production.TBOperation.Buttons(1).Enabled = False
            Production.TBOperation.Buttons(2).Enabled = False
        ElseIf state = "Disabled" Then 'disable all
            Programming.TBOperation.Buttons(0).Enabled = False
            Programming.TBOperation.Buttons(1).Enabled = False
            Programming.TBOperation.Buttons(2).Enabled = False
            Production.TBOperation.Buttons(0).Enabled = False
            Production.TBOperation.Buttons(1).Enabled = False
            Production.TBOperation.Buttons(2).Enabled = False
        ElseIf state = "Paused" Then 'pause
            Programming.TBOperation.Buttons(0).Enabled = True
            Programming.TBOperation.Buttons(1).Enabled = False
            Programming.TBOperation.Buttons(2).Enabled = True
            Production.TBOperation.Buttons(0).Enabled = True
            Production.TBOperation.Buttons(1).Enabled = False
            Production.TBOperation.Buttons(2).Enabled = True
        ElseIf state = "Only Stop Displayed" Then 'stop only
            Programming.TBOperation.Buttons(0).Enabled = False
            Programming.TBOperation.Buttons(1).Enabled = False
            Programming.TBOperation.Buttons(2).Enabled = True
            Production.TBOperation.Buttons(0).Enabled = False
            Production.TBOperation.Buttons(1).Enabled = False
            Production.TBOperation.Buttons(2).Enabled = True
        End If
        Programming.TBOperation.Refresh()
        Production.TBOperation.Refresh()

    End Sub

#End Region

#Region "calibration routines"

    Public Sub DoPurge()

        Dim state As String
        If ProductionMode() Then
            If Production.DoorCheck() = False Then
                LabelMessage("Close the door first.")
                GoTo Reset
            End If
            state = Production.ButtonPurge.Text
        Else
            state = Programming.ButtonPurge.Text
            If Programming.IsVisionTeachMode Then
                LabelMessage("Wrong mode. Please switch to needle mode first.")
                GoTo Reset
            End If
        End If

        If state = "Purge On" Then
            m_Tri.SetMachineRun()
            LockButtonsForPurge()
            Production.ButtonPurge.Enabled = False
            Programming.ButtonPurge.Enabled = False
            Programming.ButtonPurge.Text = "Purge Off"
            Production.ButtonPurge.Text = "Purge Off"
            m_Tri.Set_XY_Speed(IDS.Data.Hardware.Gantry.ServiceXYSpeed)
            m_Tri.Set_Z_Speed(IDS.Data.Hardware.Gantry.ServiceZSpeed)
            offset(0) = gLeftNeedleOffs(0)
            offset(1) = gLeftNeedleOffs(1)
            offset(2) = gLeftNeedleOffs(2) + IDS.Data.Hardware.Gantry.PurgePosition.Z
            position(0) = IDS.Data.Hardware.Gantry.PurgePosition.X - offset(0)
            position(1) = IDS.Data.Hardware.Gantry.PurgePosition.Y - offset(1)
            '1) move to safe z
            '2) move to x and y coords for purging
            '3) move to purge position 
            '4) turn on valves
            If Not m_Tri.Move_Z(SafePosition) Then GoTo Reset
            If Not m_Tri.Move_XY(position) Then GoTo Reset
            If Not m_Tri.Move_Z(offset(2)) Then GoTo Reset
            LockButtonsForPurge()
            m_Tri.TurnOn("Left Needle IO")
            Exit Sub
        Else
            m_Tri.SetMachineStop()
            Programming.ButtonPurge.Text = "Purge On"
            Production.ButtonPurge.Text = "Purge On"
            m_Tri.TurnOff("Left Needle IO")
            If Not m_Tri.Move_Z(SafePosition) Then GoTo Reset
        End If

Reset:
        ResetToIdle()

    End Sub

    Public Sub DoClean()

        Dim state As String
        If ProductionMode() Then
            state = Production.ButtonClean.Text
            If Production.DoorCheck() = False Then
                LabelMessage("Close the door first.")
                GoTo Reset
            End If
        Else
            state = Programming.ButtonClean.Text
            If Programming.IsVisionTeachMode Then
                LabelMessage("Wrong mode. Please switch to needle mode first.")
                GoTo Reset
            End If
        End If

        If state = "Clean On" Then
            m_Tri.SetMachineRun()
            LockButtonsForClean()
            Production.ButtonClean.Enabled = False
            Programming.ButtonClean.Enabled = False
            Programming.ButtonClean.Text = "Clean Off"
            Production.ButtonClean.Text = "Clean Off"
            offset(0) = gLeftNeedleOffs(0)
            offset(1) = gLeftNeedleOffs(1)
            offset(2) = gLeftNeedleOffs(2) + IDS.Data.Hardware.Gantry.CleanPosition.Z
            position(0) = IDS.Data.Hardware.Gantry.CleanPosition.X - offset(0)
            position(1) = IDS.Data.Hardware.Gantry.CleanPosition.Y - offset(1)
            m_Tri.Set_XY_Speed(IDS.Data.Hardware.Gantry.ServiceXYSpeed)
            m_Tri.Set_Z_Speed(IDS.Data.Hardware.Gantry.ServiceZSpeed)
            If Not m_Tri.Move_Z(SafePosition) Then GoTo Reset
            If Not m_Tri.Move_XY(position) Then GoTo Reset
            If Not m_Tri.Move_Z(offset(2)) Then GoTo Reset
            LockButtonsForClean()
            m_Tri.TurnOn("Clean")
            Exit Sub
        Else
            m_Tri.SetMachineStop()
            Production.ButtonClean.Text = "Clean On"
            Programming.ButtonClean.Text = "Clean On"
            m_Tri.TurnOff("Clean")
            If Not m_Tri.Move_Z(SafePosition) Then GoTo Reset
        End If

Reset:
        ResetToIdle()

    End Sub

    Public Sub DoNeedleCalibration()

        'checks
        If ProductionMode() Then
            Form_Service.LogEventInSPCReport("Needle Calibration")
            If Production.DoorCheck() = False Then
                LabelMessage("Close the door first.")
                GoTo Reset
            End If
        Else
            If Programming.IsVisionTeachMode Then
                LabelMessage("Wrong mode. Please switch to needle mode first.")
                GoTo Reset
            End If
        End If

        'execute
        LabelMessage("Needle calibration ...")
        LockMovementButtons()
        OnLaser()
        Dim rtn As Boolean = MyNeedleCalibrationSettings.NeedleCalibration()
        OffLaser()

        If rtn = True Then
            LabelMessage("Needle calibration successful.")
        Else
            LabelMessage("Needle calibration failed.")
        End If

Reset:
        ResetToIdle()

    End Sub

    Public Sub DoChangeSyringe()

        'checks
        If ProductionMode() Then
            If Production.DoorCheck() = False Then
                LabelMessage("Close the door first.")
                GoTo Reset
            End If
        Else
            If Programming.IsVisionTeachMode Then
                LabelMessage("Wrong mode. Please switch to needle mode first.")
                GoTo Reset
            End If
        End If

        offset(0) = gLeftNeedleOffs(0)
        offset(1) = gLeftNeedleOffs(1)
        offset(2) = gLeftNeedleOffs(2) + IDS.Data.Hardware.Gantry.ChangeSyringePosition.Z
        position(0) = IDS.Data.Hardware.Gantry.ChangeSyringePosition.X - offset(0)
        position(1) = IDS.Data.Hardware.Gantry.ChangeSyringePosition.Y - offset(1)

        'execute
        LockMovementButtons()
        m_Tri.SetMachineRun()
        m_Tri.Set_XY_Speed(IDS.Data.Hardware.Gantry.ServiceXYSpeed)
        m_Tri.Set_Z_Speed(IDS.Data.Hardware.Gantry.ServiceZSpeed)
        LabelMessage("Moving to change syringe position...")
        If Not m_Tri.Move_Z(SafePosition) Then GoTo Reset
        If Not m_Tri.Move_XY(position) Then GoTo Reset
        If Not m_Tri.Move_Z(offset(2)) Then GoTo Reset

Reset:
        ResetToIdle()

    End Sub

    Public Sub DoVolumeCalibration()

        If ProductionMode() Then
            If Production.DoorCheck() = False Then
                LabelMessage("Close the door first.")
                GoTo Reset
            End If
            'Form_Service.LogEventInSPCReport("Volume Calibration")
        Else
            If Programming.IsVisionTeachMode Then
                LabelMessage("Wrong mode. Please switch to needle mode first.")
                GoTo Reset
            End If
        End If

        LockMovementButtons()
        LabelMessage("Doing volume calibration.")
        MyVolumeCalibrationSettings.VolumeCalibrationSetup()
        VolumeCalibrationRunning = True
        m_Tri.SetMachineRun()
        Dim rtn As Boolean = MyVolumeCalibrationSettings.VolumeCalibration(IDS.Data.Hardware.Volume.Left.NumberOfAttempts)
        If rtn = True Then
            LabelMessage("Volume calibration successful.")
        Else
            LabelMessage("Volume calibration failed.")
        End If

Reset:
        ResetToIdle()

    End Sub

    Public Sub PauseDispensing()
        SetState("Pause")
        m_Tri.SetMachinePause()
        ChangeButtonState("Paused")
        LabelMessage("Dispensing Pause!")
        If VolumeCalibrationRunning = True Then MyVolumeCalibrationSettings.VolumeCalibrationState = "Paused"
        If ShouldLog() Then Form_Service.LogEventInSPCReport("Pause")
    End Sub

    Public Sub StopDispensing()
        SetState("Stop")
        m_Tri.TrioStop()
        If Not ThreadExecution Is Nothing Then ThreadExecution.Abort()
        If VolumeCalibrationRunning = True Then MyVolumeCalibrationSettings.VolumeCalibrationState = "Stopped"
        VolumeCalibrationRunning = False
        LabelMessage("Dispensing Stop!")
        If Not WasStart() Then
            LockMovementButtons()
            TravelToParkPosition()
            ResetToIdle()
        End If
        If ProductionMode() Then
            Conveyor.Command("Release")
            Production.WriteSPCReport()
        End If
    End Sub

    Public Sub ResumeDispensing()
        If VolumeCalibrationRunning = True Then MyVolumeCalibrationSettings.VolumeCalibrationState = "Running"
        ChangeButtonState("Running")
        m_Tri.SetMachineRun()
        If ShouldLog() Then Form_Service.LogEventInSPCReport("Resume")
        LabelMessage("Dispensing Resume..")
        SetState("Run")
    End Sub

#End Region

#Region "io/hardware"

    Public Sub SwitchToTeachCamera()
        If ProductionMode() And Production.ContinuousMode.Checked = True Then
        Else
            Vision.FrmVision.SetBrightness(Programming.ValueBrightness.Value)
        End If
        ClearAndDisplayIndicator()
        Vision.FrmVision.SwitchCamera("Teach Camera")
    End Sub

    Public Sub SwitchToRealTimeCamera()
        Vision.FrmVision.ClearDisplay()
        Vision.FrmVision.SwitchCamera("View Camera")
    End Sub

    Public Sub OffTowerLamp()
        DIO_Service.Off_Red_Tower_Lamp()
        DIO_Service.Off_Green_Tower_Lamp()
        DIO_Service.Off_Yellow_Tower_Lamp()
    End Sub

    Public Function UnlockDoor()
        IDS.Devices.DIO.DIO.DoorLock(True) 'unlock is true
    End Function

    Public Function LockDoor()
        IDS.Devices.DIO.DIO.DoorLock(False) 'lock is false
    End Function

    Public Sub OnLaser()
        Vision.FrmVision.SetLaser(False) 'turn on
    End Sub

    Public Sub OffLaser()
        Vision.FrmVision.SetLaser(True) 'turn offset
    End Sub

#End Region

End Module