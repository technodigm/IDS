Imports DLL_Export_Service
Imports DLL_Export_Device_Motor
Imports DLL_SetupAndCalibrate
Imports DLL_DataManager
Imports Microsoft.DirectX.DirectInput
Imports System.Threading
Imports System.IO
Imports PLC
Imports Laser

Public Class CIDSExe

    Public Function CALLME()
    End Function

    'Setting system setup parameters to execution moudle's global data
    Public Function SetExecutionSetupParam()
        SystemSetupDataRetrieve()
    End Function

End Class

Public Module ExecutionModule
#Region "Declaration"
    'threading declarations
    Public PreviousMachineState As String = "Idle"
    Public MachineState As String = "Idle"
    Public ThreadExecution As Threading.Thread
    Public ThreadMonitor As Threading.Thread
    Public ThreadExecutor As Threading.Thread

    Public InitThread As Threading.Thread

    Public VolumeCalibrationRunning As Boolean = False
    Public NeedleCalibrationRunning As Boolean = False
    Public AbortPurgeCleanChgSyringe As Boolean = False

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
    Public Conveyor As PLC.IPLC = CIDSService.CIDSServiceDevices.CIDSServiceConveyor.Instance
    'Public Weighting_Scale As CIDSService.CIDSServiceDevices.CIDSServiceWeighting = Weighting_Scale.Instance
    Public Weighting_Scale As WeightScale.IWeightingScale = CIDSService.CIDSServiceDevices.CIDSServiceWeighting.Instance
    Public Heater As CIDSService.CIDSServiceDevices.CIDSServiceThermal = Heater.Instance
    Public Laser As ILaser = CIDSService.CIDSServiceDevices.CIDSServiceLaser.Instance
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


#End Region
    Private thisLock As Object = New Object
    Sub New()
        ConfigureLog()
    End Sub
    Private Function ConfigureLog()
        Dim repo As log4net.Repository.ILoggerRepository = log4net.LogManager.GetRepository()
        Dim lstAppender() As log4net.Appender.IAppender = repo.GetAppenders()
        For Each itemAppender As log4net.Appender.IAppender In lstAppender
            Dim fileappender As log4net.Appender.RollingFileAppender = itemAppender
            fileappender.AppendToFile = True
            fileappender.MaximumFileSize = "2MB"
            fileappender.MaxSizeRollBackups = -1
            fileappender.Threshold = log4net.Core.Level.Debug
            fileappender.File = "C:\Log\logfile.txt"
            'fileappender.DatePattern = "ddMMyyyy"
            fileappender.DatePattern = "yyyy\\\\MM\\\\dd\\\\yyyyMMdd"
            fileappender.StaticLogFileName = False
            fileappender.Layout = New log4net.Layout.PatternLayout("%date %-5level - [%logger] %message%newline")
            fileappender.RollingStyle = log4net.Appender.RollingFileAppender.RollingMode.Composite
            fileappender.ActivateOptions()
            log4net.Config.BasicConfigurator.Configure(fileappender)
        Next
        DeleteEmptyLog()
    End Function

#Region "useful utilities"
    Private Function DeleteEmptyLog()
        Dim path As String = Directory.GetCurrentDirectory() + "\(null)" + DateTime.Now.ToString("ddMMyyyy") + ".txt"
        If File.Exists(path) Then
            File.Delete(path)
        End If
    End Function
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
        Vision.FrmVision.ClearDisplay()
    End Sub

    Public Sub ClearAndDisplayIndicator()
        Vision.FrmVision.DisplayIndicator()
        'If ProgrammingMode() Then
        '    Vision.FrmVision.DisplayIndicator()
        'Else
        '    Vision.FrmVision.ClearDisplay()
        'End If
    End Sub
    Public Sub DisplayCrossHair()
        Vision.FrmVision.DisplayIndicator()
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
            Programming.TextBoxRobotX.Refresh()
            Programming.TextBoxRobotY.Text = "Y: " & YCor
            Programming.TextBoxRobotY.Refresh()
            Programming.TextBoxRobotZ.Text = "Z: " & ZCor
            Programming.TextBoxRobotZ.Refresh()
        Else
            Production.TextBoxRobotPos.Text = " X: " & XCor
            Production.tbYPost.Text = " Y: " & YCor
            Production.tbZPost.Text = " Z: " & ZCor
        End If

    End Sub

    Private Sub Callback(ByVal ar As IAsyncResult)
        ar.AsyncState.EndInvoke(ar)
    End Sub

    Private Sub UpdateMachineStatusCallback(ByVal ar As IAsyncResult)
        ar.AsyncState.EndInvoke(ar)
    End Sub

    Private Sub UpdateTeachPoint(ByVal XCor As Double, ByVal YCor As Double, ByVal ZCor As Double)
        If Not (m_PosUpdate) Then Return
        Dim colmX, colmY, colmZ As Integer
        Try

            If m_TeachStepNumber = 1 Then
                colmX = gPos1XColumn
                colmY = gPos1YColumn
                colmZ = gPos1ZColumn
                If ProgrammingMode() Then
                    If Not (Programming.ArrayDlg Is Nothing) Then
                        If Programming.ArrayDlg.needPositionUpdate Then
                            Programming.ArrayDlg.SetPoint(XCor, YCor, ZCor)
                        End If
                    End If
                End If

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
                If Programming.teachingMode = "Needle" Then 'NeedleMode.Checked Or Programming.m_TeachMode = 2 Then 'LEFT Right
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
    Public localRunning As Boolean = False
    Public Sub GeneratePatternAndRunProgram()

        Dim rtn As Integer
        Vision.FrmVision.DisplayIndicator()
        TraceGCCollect()
        LabelMessage("Carrying out reference checks.")
        m_Execution.m_Command.SetOptimizFlag(0)
        m_Tri.SetMachineRun()
        ChangeButtonState("Running")
        SwitchToTeachCamera()
        'OnLaser()
        OffLaser()
        localRunning = True
        rtn = m_Execution.m_Command.ReadPattern(Programming.AxSpreadsheetProgramming)

        'to show cross
        'Vision.FrmVision.DisplayIndicator()
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
        localRunning = False
        LabelMessage("Compiling program.")
        Vision.FrmVision.DisplayIndicator()
        If m_Execution.m_Command.Compile(Programming.m_RunMode, Programming.GlobalQCEnabled) < 0 Then
            LabelMessage(ErrorMessage())
            ' LabelMessage("Compilation error.")
            GoTo resetmachinestate
        End If
        Try
            SetState("Run")
            LabelMessage("Downloading")
            ThreadExecution = New Threading.Thread(AddressOf DownloadProgram)
            ThreadExecution.Priority = Threading.ThreadPriority.Normal
            ThreadExecution.Start()
        Catch ex As ThreadAbortException
        Catch ex As Exception
            ExceptionDisplay(ex)
        End Try

        Exit Sub

ResetMachineState:
        Vision.FrmVision.DisplayIndicator()
        If ProductionMode() Then
            Production.btExit.Enabled = True
        End If
        If Not stopping Then
            LockMovementButtons()
            LabelMessage("#Generate pattern", False, False)
            TravelToParkPosition()
            ResetToIdle()
        End If
        stopping = False
        LabelMessage("#Exit GP", False, False)
        localRunning = False
    End Sub

    Public Function DispensingCtrl() As Integer

        Dim Dispensinglist As ArrayList
        Dim cmdBurn As New CIDSPattnBurn

        Try
            If m_Execution.m_Command.CompileStatus > 0 Then

                Dispensinglist = m_Execution.m_Command.DispenseList
                m_Tri.SetMachineRunMode(Programming.m_RunMode)
                LabelMessage("Downloading data......")
                If cmdBurn.BurnTable(Dispensinglist) = False Then
                    LabelMessage("Downloading failed or cancelled.")
                    Return -1
                Else
                    SetLampsToRunningMode()
                    'LabelMessage("Download finished.")
                    LabelMessage("Dispensing")
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
    Public trioMotionConnection As Boolean = False
    Public Sub StateMonitor()
        Dim timeStart As DateTime = DateTime.Now
        Do
            Try
                'how long the thread waits between updates
                MySleep(50)
                'this is to prevent e-stop window from freezing.
                'if may cause teaching problems.

                If m_Tri.GetIDSState() Then
                    If ProductionMode() Then
                        If Not trioMotionConnection Then
                            Production.tbRobotStatus.Text = "Good"
                            Production.tbRobotStatus.BackColor = System.Drawing.SystemColors.Control
                            trioMotionConnection = True
                        End If
                    Else

                    End If
                Else
                    If trioMotionConnection Then
                        Production.tbRobotStatus.Text = "Lost connection"
                        Production.tbRobotStatus.BackColor = System.Drawing.Color.Red
                        trioMotionConnection = False
                    End If
                End If

                'spc report
                If ProductionMode() And (m_Tri.MotionError Or m_Tri.EStopActivated) Then
                    If ProductionMode() Then Production.WriteSPCReport()
                    LabelMessage("Motion controller error. Please restart the motion controller and program.")
                End If

                'Dim status_update As String = MyVolumeCalibrationSettings.VolumeCalibrationResult
                'If VolumeCalibrationRunning And status_update <> "" Then
                '    If ProgrammingMode() Then
                '        If Programming.LabelMessege.Text <> status_update Then
                '            LabelMessage(status_update)
                '        End If
                '    Else
                '        If Production.LabelMessege.Text <> status_update Then
                '            LabelMessage(status_update)
                '        End If
                '    End If
                'End If

                Hpos(0) = m_Tri.XPosition
                Hpos(1) = m_Tri.YPosition
                Hpos(2) = m_Tri.ZPosition
                'Programming.openEventViewer = False
                HardToSys(Hpos, Spos)
                ref(0) = Programming.m_ReferPt(0)
                ref(1) = Programming.m_ReferPt(1)
                ref(2) = Programming.m_ReferPt(2)
                SysToRefer(Spos, Spos, ref)

                If Programming.teachingMode = "Needle" Then 'NeedleMode.Checked Then
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
                        'Update post every 200 ms
                        If (DateTime.Now.Ticks - timeStart.Ticks) / 10000 > 200 Then
                            timeStart = DateTime.Now
                            updateTeachPosition.BeginInvoke(WorkX, WorkY, WorkZ, AddressOf UpdateTeachPointCallback, updateTeachPosition)
                        End If

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
        'MyVolumeCalibrationSettings.VolumeCalibrationSetup()
        MyVolumeCalibrationSettings.VolumeCalibrationParamSetup()
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
    Public Function HandlingGlobalQCCheck() As Integer

        Return 0
    End Function
    'This function is to handle a standalone Global QC check.
    'Global QC check only need to set once and all the dots will be 
    'checked after all dispensing was done.
    'Return 1 is abort, 2 is continue, 0 is QC passed
    Public Function GlobalQCCheck() As Integer
        SwitchToTeachCamera()
        Dim vParam As DLL_Export_Device_Vision.QC.QCParam
        vParam._Binarized = m_Execution.m_Command.globalQCParam._Binarized
        vParam._BlackDot = m_Execution.m_Command.globalQCParam._BlackDot

        vParam._Brightness = m_Execution.m_Command.globalQCParam._Brightness
        'Set the brightness before doing the movement. This will affect vision(ActiveX) less.
        Vision.IDSV_SetBrightness(vParam._Brightness)
        vParam._Close = m_Execution.m_Command.globalQCParam._Close
        vParam._Compactness = m_Execution.m_Command.globalQCParam._Compactness
        vParam._MaxArea = m_Execution.m_Command.globalQCParam._MaxArea
        vParam._MinArea = m_Execution.m_Command.globalQCParam._MinArea
        vParam._MRegionX = m_Execution.m_Command.globalQCParam._MRegionX
        vParam._MRegionY = m_Execution.m_Command.globalQCParam._MRegionY
        vParam._MROIx = m_Execution.m_Command.globalQCParam._MROIx
        vParam._MROIy = m_Execution.m_Command.globalQCParam._MROIy
        vParam._Open = m_Execution.m_Command.globalQCParam._Open
        vParam._Roughness = m_Execution.m_Command.globalQCParam._Roughness
        vParam._Diameter = m_Execution.m_Command.globalQCParam._Diameter
        vParam._Tolerance = m_Execution.m_Command.globalQCParam._Tolerance
        Dim rtn As Boolean = Vision.IDSV_QC(vParam)
        If rtn Then
            If ShouldLog() Then Form_Service.LogEventInSPCReport("Dot Size Check Passed")
            Programming.LogFile("Dot QC passed.")
            Return True
        Else
            If CheckButtonState() = -1 Then Exit Function
            If IDS.Data.Hardware.SPC.QCFailedAction = False And ProductionMode() Then
                Form_Service.LogEventInSPCReport("Dot Size Check Failed")
            ElseIf IDS.Data.Hardware.SPC.QCFailedAction = True And ProductionMode() Then
                Form_Service.LogEventInSPCReport("Dot Size Check Failed")
                Form_Service.DisplayErrorMessage("Dot Size Check Failed")
            ElseIf IDS.Data.Hardware.SPC.QCFailedAction = True And ProgrammingMode() Then
                Form_Service.DisplayErrorMessage("Dot Size Check Failed")
            ElseIf IDS.Data.Hardware.SPC.QCFailedAction = False And ProgrammingMode() Then
                Programming.LogFile("Dot size check failed! #QC failed action alarm bypassed.")
                Return 0
            End If

            WaitUntilErrorMessagesCleared()
            If Form_Service.DotSizeCheckStop Then
                If ShouldLog() Then Form_Service.LogEventInSPCReport("Board Total Failure")
                LabelMessage("Dot size failed. Process aborted!", True)
                Return 1
            ElseIf Form_Service.DotSizeCheckContinue Then
                If ShouldLog() Then Form_Service.LogEventInSPCReport("Board Partial Failure")
                LabelMessage("Dot size failed alarm ignored", True)
                Return 2
            Else
                'estop
                Return 1
            End If
        End If
        Return 1
    End Function

    Public Sub QCcheck()

        Dim vParam As DLL_Export_Device_Vision.QC.QCParam
        Dim readbuf(16) As Double
        LabelMessage("QC Checking")
        SwitchToTeachCamera()
        m_Tri.ResetQCFlag()

        m_Tri.m_TriCtrl.GetTable(210, 16, readbuf)
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
        LabelMessage("QC Checked")
        ClearDisplay()
        LabelMessage("#2 QC Checked")
        If Programming.m_RunMode <> 0 And readbuf(15) = 0 Then SwitchToRealTimeCamera()
        If rtn Then
            If ShouldLog() Then Form_Service.LogEventInSPCReport("Dot Size Check Passed")
            LabelMessage("Dot size check passed! Diameter(mm): " + Vision.diameterResult.ToString("0.000"))
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

            'WaitUntilErrorMessagesCleared()

            If Form_Service.DotSizeCheckStop Then
                If ShouldLog() Then Form_Service.LogEventInSPCReport("Board Total Failure")
                m_Tri.SetQCStop()
                LabelMessage("Machine stopped.")
                LockMovementButtons()
                TravelToParkPosition()
                ResetToIdle()
                Vision.FrmVision.DisplayIndicator()
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
    Private sTime As Long = 0
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
                            Programming.AxSpreadsheetProgramming.Enabled = False
                            Programming.DisableTeachModeSwitching()
                            If VolumeCalibrationRunning = True Then MyVolumeCalibrationSettings.VolumeCalibrationState = "Running"
                            LockMovementButtons()
                            MoveZToSafePosition()
                            If IDS.Data.Hardware.Dispenser.Left.HeadType = "Jetting Valve" Or IDS.Data.Hardware.Dispenser.Left.HeadType = "Auger Valve" Then
                                'This is to always on the Valve if using jetting valve
                                'This is Output 27 of Trio Controller
                                'Output 25 used as trigger signal and is connected to microcontroller that used
                                'to send PWM to jetting driver in order to dispense. All the dispensing    
                                'parameter like How many dispenses or dots are control by micro controller.
                                If Programming.m_RunMode = 4 Then
                                    m_Tri.TurnOn("Material Air")
                                End If

                            End If
                            SetLampsToRunningMode()
                            GeneratePatternAndRunProgram()
                        Else
                            If Programming.IsFilePresent Then
                                ResetToIdle()
                                LabelMessage("No file loaded.")
                                Exit Select
                            End If
                            LockMovementButtons()
                            Production.ContinuousMode.Enabled = False
                            If Production.ContinuousMode.Checked Then
                                PortLifeAction()
                                'AutoPurgeAction()
                            End If
                            If IDS.Data.Hardware.Dispenser.Left.HeadType = "Jetting Valve" Or IDS.Data.Hardware.Dispenser.Left.HeadType = "Auger Valve" Then
                                'This is to always on the Valve if using jetting valve
                                'This is Output 27 of Trio Controller
                                'Output 25 used as trigger signal and is connected to microcontroller that used
                                'to send PWM to jetting driver in order to dispense. All the dispensing    
                                'parameter like How many dispenses or dots are control by micro controller.
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
                                LabelMessage("Moving to park position")
                                LockMovementButtons()
                                TravelToParkPosition()
                                ResetToIdle()
                                ClearAndDisplayIndicator()
                                LabelMessage("Dispensing done")
                                Programming.NeedleMode.Enabled = True
                                Programming.VisionMode.Enabled = True
                                TraceGCCollect()
                                If Programming.cbContinueTest.Checked Then
                                    MachineState = "Start"
                                End If
                            Else
                                If Production.ContinuousMode.Checked = True Then
                                    Production.DispensingFinishCallback.Invoke() 'synchronous
                                    ClearAndDisplayIndicator()
                                Else
                                    LabelMessage("Moving to park position")
                                    LockMovementButtons()
                                    TravelToParkPosition()
                                    Production.btExit.Enabled = True
                                    LabelMessage("Dispensing done")
                                    ResetToIdle()
                                    ClearAndDisplayIndicator()
                                    'PortLifeAction()
                                    'AutoPurgeAction()
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
                        MyVolumeCalibrationSettings.attemptCount = 0
                        Vision.FrmVision.SwitchCamera("View Camera")
                        Vision.FrmVision.ClearDisplay()
                        DoVolumeCalibration()
                        Do
                            MySleep(25)
                            TraceDoEvents()
                        Loop Until Not (MyVolumeCalibrationSettings.VolumeCalibrationState = "Running") Or m_Tri.EStopActivated
                        VCResult()
                        Vision.FrmVision.SwitchCamera("Teach Camera")
                        Vision.FrmVision.DisplayIndicator()
                        IDS.Devices.Vision.IDSV_SetBrightness(IDS.Data.Hardware.Camera.Brightness)
                    Case "Needle Calibration"
                        DoNeedleCalibration()


                    Case "Jogging"
                        LockMovementButtonsOnce()
                        Do
                            MySleep(25)
                            TraceDoEvents()
                        Loop Until Not m_Tri.MachineJogging Or m_Tri.EStopActivated

                    Case "Homing"
                        MySleep(500)
                        DoHoming()
                        MySleep(500) 'to wait for ids.getstate in statemonitor, otherwise it will exit
                        'the loop immediately.
                        sTime = DateTime.Now.Ticks
                        Do
                            MySleep(25)
                            TraceDoEvents()
                        Loop Until m_Tri.HomingFinished Or m_Tri.EStopActivated Or m_Tri.IDSIsHomed() Or ((DateTime.Now.Ticks - sTime) / 10000) > 60000
                        If (DateTime.Now.Ticks - sTime) / 10000 > 60000 Then
                            Dim fm As InfoForm = New InfoForm
                            If m_Tri.IsOpen Then
                                fm.AddNewLine("Homing timout after 60 seconds.")
                            Else
                                fm.AddNewLine("Homing timeout! Connection to motion controller dropped!")
                                fm.AddNewLine("Please make network connection to motion controller is established")
                            End If

                            fm.OkOnly()
                            fm.TopMost = True
                            fm.ShowDialog()
                        End If
                        ResetToIdle()

                        If ProductionMode() Then
                            Production.ButtonOpenFile.Enabled = True
                            Production.btPlay.Enabled = Production.fileLoaded
                            Production.ButtonVolCalib.Enabled = Production.btPlay.Enabled
                            Production.ButtonClean.Enabled = Production.btPlay.Enabled
                            Production.ButtonNdlCalib.Enabled = Production.btPlay.Enabled
                            Production.ButtonChgSyringe.Enabled = Production.btPlay.Enabled
                            Production.ButtonPurge.Enabled = Production.btPlay.Enabled
                        Else
                            Programming.ButtonToggleMode.Enabled = False
                            LabelMessage("Opening default file......")
                            Programming.Enabled = False
                            Programming.OpenDefaultFile()
                            Programming.Enabled = True
                            Programming.EnableClickToMove()
                            Programming.ButtonToggleMode.Enabled = True
                        End If
                        LabelMessage("System Ready")
                        If ProductionMode() Then
                            If Production.firstLoad Then
                                LabelMessage("Please select a pattern file.")
                                Production.firstLoad = False
                            End If
                        End If
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
                StopDispensing()
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
                    'Console.WriteLine("Thread exit #1")
                    abortInit.Reset()
                    isExited = True
                    Return
                Case 1
                    'Console.WriteLine("Thread create Motor instance")
                    If m_Tri Is Nothing Then
                        m_Tri = New CIDSService.CIDSServiceDevices.CIDSMotor
                    End If
                    If Not m_Tri.Connect_Controller() Then
                        If ProductionMode() Then
                            Production.tbRobotStatus.Text = "No Connection"
                            Production.tbRobotStatus.BackColor = System.Drawing.Color.Red
                            isInited = False
                        End If
                    Else
                        isInited = True
                    End If
                    'Assigne motor class to setup when running programming mode or production mode
                    DLL_SetupAndCalibrate.m_Tri = m_Tri
                    'isInited = True
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
                    If ProductionMode() Then
                        Production.btExit.Enabled = False
                    End If
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

        'If ProductionMode() Then
        '    Production.TextBox1.Text = PreviousMachineState
        '    Production.TextBox2.Text = MachineState
        'Else
        '    Programming.TextBox1.Text = PreviousMachineState
        '    Programming.TextBox2.Text = MachineState
        'End If
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
    'This function is use to track if the system is started and on standby for production
    Public Function IsStopped()
        If MachineState = "stop" Then Return True
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
        'LabelMessage("System Idle")
        If ProgrammingMode() Then
            Programming.AxSpreadsheetProgramming.Enabled = True
            Programming.DispensingMode.Enabled = True
            Programming.VisionMode.Enabled = True
            Programming.NeedleMode.Enabled = True
            Programming.MenuItem1.Enabled = True
            Programming.btExit.Enabled = True
            Programming.btReset.Enabled = True
            Programming.btEject.Enabled = True
        Else
            If Not (Production.ContinuousMode.Checked) Then
                Production.ConveyorBox.Enabled = True
                Production.ButtonCV_Prod_Retrieve.Enabled = True
                Production.ButtonStartFirstStage.Enabled = True
                Production.ButtonCV_Prod_Release.Enabled = True
                Production.ResetPLCLogic.Enabled = True
                Production.btExit.Enabled = True
            End If
            Production.ContinuousMode.Enabled = True
        End If

    End Sub

    Public Sub PortLifeAction()
        If Production.m_PotLifeExpire Then
            Dim fm As InfoForm = New InfoForm(True)
            fm.AddNewLine("Pot Life Expired!!")
            fm.AddNewLine("Click OK to change the syringe now!")
            fm.AddNewLine("Click Cancel to ignore this message and pot life checking")
            fm.AddNewLine("will be disabled.")
            If fm.ShowDialog() = DialogResult.OK Then
                SetLampsToRunningMode()
                DoChangeSyringe()
                SetLampsToReadyMode()
            Else
                Production.CheckBoxPotOn.Text = "Pot Life On"
                Production.CheckBoxPotOn.Checked = False
                Production.m_PotLifeExpire = False
            End If
        End If
    End Sub

    Public Sub AutoPurgeAction()
        If Production.m_AutoPurgeRequired Then
            Production.tbAutoPurgeCountDown.Text = "Auto Purging Now"
            SetLampsToRunningMode()
            Production.ProdPurgeClean()
            SetLampsToReadyMode()
            Production.tbAutoPurgeCountDown.Text = "Auto Purging Done"
            Production.m_AutoPurgeRequired = False
            Production.ResetTimer("Reset Autopurging Timers")
        End If
    End Sub

#End Region

#Region "trio motion routines"
    Public forceHome As Boolean = False
    Public Sub DoHoming()
        If m_Tri.m_TriCtrl.IsOpen(0) Then
            m_Tri.GetIDSIsHomeState()
            m_Tri.GetIDSIsHomeState()
            If Not (forceHome) Then
                If m_Tri.HomingFinished Then
                    Return
                End If
            End If
            forceHome = False
            LabelMessage("Homing..")
            m_Tri.SetMachineRun()
            SetLampsToRunningMode()
            LockMovementButtons()
            If ProgrammingMode() Then
                Programming.btExit.Enabled = False
            Else
                Production.btExit.Enabled = False
            End If
            'SetState("Homing")
            m_Tri.m_TriCtrl.Execute("RUN SETDATUM")
        End If
    End Sub

    Public Function MoveZToSafePosition() As Boolean
        m_Tri.Set_Z_Speed(IDS.Data.Hardware.Gantry.ElementZSpeed)
        If Not m_Tri.Move_Z(SafePosition) Then
            Return False
        End If
        Return True
    End Function

    Public Function TravelToParkPosition() As Boolean
        LabelMessage("Move to parking position")
        'warning! we do not set run mode in this function.
        position(0) = IDS.Data.Hardware.Gantry.ParkPosition.X
        position(1) = IDS.Data.Hardware.Gantry.ParkPosition.Y
        m_Tri.Set_XY_Speed(IDS.Data.Hardware.Gantry.ElementXYSpeed)
        LabelMessage("Move to Z safe position", False, False)
        If Not MoveZToSafePosition() Then
            LabelMessage("Travel to park: Move Z to safe position failed")
            Return False
        End If
        LabelMessage("Move to XY parking position", False, False)
        If Not m_Tri.Move_XY(position) Then
            LabelMessage("Move XY to parking position failed")
            Return False
        End If
        LabelMessage("Stationed at parking position")
        Return True
    End Function

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
            Production.ButtonPurge.Enabled = True
        End If
    End Sub

    Public Sub LockButtonsForClean()
        LockMovementButtons()
        If ProgrammingMode() Then
            Programming.ButtonClean.Enabled = True
        ElseIf ProductionMode() Then
            Production.ButtonClean.Enabled = True
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
                If NeedleCalibrationRunning = False Then
                    .ButtonNeedleCal.Enabled = False
                End If
                If VolumeCalibrationRunning = False Then
                    .ButtonVolCal.Enabled = False
                End If
                .ButtonPurge.Enabled = False
                .ButtonClean.Enabled = False
                .DispensingMode.Enabled = False
                .btRelease.Enabled = False
                .btRetrieve.Enabled = False
                .ButtonStartFirstStage.Enabled = False
                .ButtonToggleMode.Enabled = False
                .VisionMode.Enabled = False
                .NeedleMode.Enabled = False
                .btChangeSyringe.Enabled = False
                .btResetVolCalSettings.Enabled = False
                Programming.btReset.Enabled = False
                Programming.btEject.Enabled = False
                Programming.ElementsCommandBlock.Enabled = False
                Programming.ReferenceCommandBlock.Enabled = False
                Programming.btResetVolumeSettings.Enabled = False
            End With
        ElseIf ProductionMode() Then
            With Production
                .ButtonHome.Enabled = False
                .ButtonChgSyringe.Enabled = False
                .ButtonClean.Enabled = False
                .DoorLock.Enabled = False
                If VolumeCalibrationRunning = False Then
                    .ButtonVolCalib.Enabled = False
                End If
                If NeedleCalibrationRunning = False Then
                    .ButtonNdlCalib.Enabled = False
                End If
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
                .btRelease.Enabled = True
                .btRetrieve.Enabled = True
                .ButtonStartFirstStage.Enabled = True
                .ButtonToggleMode.Enabled = True
                .VisionMode.Enabled = True
                .NeedleMode.Enabled = True
                .btChangeSyringe.Enabled = True
                .btResetVolCalSettings.Enabled = True
                Programming.btReset.Enabled = True
                Programming.btEject.Enabled = True
                Programming.ElementsCommandBlock.Enabled = True
                Programming.ReferenceCommandBlock.Enabled = True
                Programming.btResetVolumeSettings.Enabled = True
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
                .DoorLock.Enabled = True
            End With
        End If
    End Sub
    Public Sub LabelMessage(ByVal message As String, Optional ByVal isError As Boolean = False, Optional ByVal logToScreen As Boolean = True)
        Try
            If thisLock Is Nothing Then Exit Sub
            Monitor.Enter(thisLock)
            If gExeMode <> "Operator" Then
                If logToScreen Then
                    Programming.LabelMessege.Text = message
                    Programming.LabelMessege.Refresh()
                End If
                If isError Then
                    Programming.LogFile(message, 2)
                Else
                    Programming.LogFile(message)
                End If
            Else
                Production.LabelMessege.Text = message
                Production.LabelMessege.Refresh()
                If logToScreen Then
                    If isError Then
                        Production.LogScreen(message, True, 2)
                    Else
                        Production.LogScreen(message)
                    End If
                End If
            End If
            Monitor.Exit(thisLock)
        Catch ex As Exception
            Monitor.Exit(thisLock)
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
            Programming.GreenLight.Image = Programming.TowerLightImageList.Images.Item(2)
            Programming.AmberLight.Image = Programming.TowerLightImageList.Images.Item(3)
            Programming.RedLight.Image = Programming.TowerLightImageList.Images.Item(3)
        Else
            Production.GreenLight.Image = Production.TowerLightImageList.Images.Item(2)
            Production.AmberLight.Image = Production.TowerLightImageList.Images.Item(3)
            Production.RedLight.Image = Production.TowerLightImageList.Images.Item(3)
        End If
    End Sub

    Public Sub SetLampsToReadyMode()
        DIO_Service.Off_Red_Tower_Lamp()
        DIO_Service.Off_Green_Tower_Lamp()
        DIO_Service.On_Yellow_Tower_Lamp()
        If gExeMode <> "Operator" Then
            Programming.GreenLight.Image = Programming.TowerLightImageList.Images.Item(3)
            Programming.AmberLight.Image = Programming.TowerLightImageList.Images.Item(1)
            Programming.RedLight.Image = Programming.TowerLightImageList.Images.Item(3)
        Else
            Production.GreenLight.Image = Production.TowerLightImageList.Images.Item(3)
            Production.AmberLight.Image = Production.TowerLightImageList.Images.Item(1)
            Production.RedLight.Image = Production.TowerLightImageList.Images.Item(3)
        End If
    End Sub

    Public Sub SetLampsToErrorMode()
        DIO_Service.On_Red_Tower_Lamp()
        DIO_Service.Off_Green_Tower_Lamp()
        DIO_Service.Off_Yellow_Tower_Lamp()
        If gExeMode <> "Operator" Then
            Programming.GreenLight.Image = Programming.TowerLightImageList.Images.Item(3)
            Programming.AmberLight.Image = Programming.TowerLightImageList.Images.Item(3)
            Programming.RedLight.Image = Programming.TowerLightImageList.Images.Item(0)
        Else
            Production.GreenLight.Image = Production.TowerLightImageList.Images.Item(3)
            Production.AmberLight.Image = Production.TowerLightImageList.Images.Item(3)
            Production.RedLight.Image = Production.TowerLightImageList.Images.Item(0)
        End If
    End Sub

    Public Sub ChangeButtonState(ByVal state As String) 'kr

        If state = "Running" Then 'running so pause or stop only
            Programming.btPlay.Enabled = False
            Programming.btPause.Enabled = True
            Programming.btStop.Enabled = True
            Production.btPlay.Enabled = False
            Production.btPause.Enabled = True
            Production.btStop.Enabled = True
        ElseIf state = "Idle" Then 'idle so run only
            Programming.btPlay.Enabled = True
            Programming.btPause.Enabled = False
            Programming.btStop.Enabled = False
            Production.btPlay.Enabled = True
            Production.btPause.Enabled = False
            Production.btStop.Enabled = False
        ElseIf state = "Disabled" Then 'disable all
            Programming.btPlay.Enabled = False
            Programming.btPause.Enabled = False
            Programming.btStop.Enabled = False
            Production.btPlay.Enabled = False
            Production.btPause.Enabled = False
            Production.btStop.Enabled = False
        ElseIf state = "Paused" Then 'pause
            Programming.btPlay.Enabled = True
            Programming.btPause.Enabled = False
            Programming.btStop.Enabled = True
            Production.btPlay.Enabled = True
            Production.btPause.Enabled = False
            Production.btStop.Enabled = True
        ElseIf state = "Only Stop Displayed" Then 'stop only
            Programming.btPlay.Enabled = False
            Programming.btPause.Enabled = False
            Programming.btStop.Enabled = True
            Production.btPlay.Enabled = False
            Production.btPause.Enabled = False
            Production.btStop.Enabled = True
        End If
    End Sub

#End Region

#Region "calibration routines"
    Private stopPurge As Boolean = False
    Private PurgingRunning As Boolean = False
    Private purgeLock As Boolean = False
    Public Sub DoPurge(Optional ByVal notMoveAfterPurge As Boolean = False)

        Dim state As String
        If ProductionMode() Then
            If Production.DoorCheck() = False Then
                LabelMessage("Close the door first.")
                Return
                GoTo Reset
            End If
            state = Production.ButtonPurge.Text
        Else
            state = Programming.ButtonPurge.Text
            '    If Programming.IsVisionTeachMode Then
            '        LabelMessage("Wrong mode. Please switch to needle mode first.")
            '        GoTo Reset
            '    End If
        End If

        If state = "Purge On" Then
            stopPurge = False
            PurgingRunning = True
            Production.ButtonPurge.Enabled = False
            Programming.ButtonPurge.Enabled = False
            m_Tri.SetMachineRun()
            LockButtonsForPurge()
            DLL_SetupAndCalibrate.MyDispenserSettings.DownloadMaterialAirPressure(IDS.Data.Hardware.Dispenser.Left.MaterialAirPressure, IDS.Data.Hardware.Dispenser.Left.SuckbackPressure)
            Programming.ButtonPurge.Text = "Purge Off"
            Production.ButtonPurge.Text = "Purge Off"
            m_Tri.Set_XY_Speed(IDS.Data.Hardware.Gantry.ServiceXYSpeed)
            m_Tri.Set_Z_Speed(IDS.Data.Hardware.Gantry.ServiceZSpeed)
            offset(0) = gLeftNeedleOffs(0)
            offset(1) = gLeftNeedleOffs(1)
            offset(2) = gLeftNeedleOffs(2) + IDS.Data.Hardware.Gantry.PurgePosition.Z
            position(0) = IDS.Data.Hardware.Gantry.PurgePosition.X - offset(0)
            position(1) = IDS.Data.Hardware.Gantry.PurgePosition.Y - offset(1)
            Production.ButtonPurge.Enabled = True
            Programming.ButtonPurge.Enabled = True
            '1) move to safe z
            '2) move to x and y coords for purging
            '3) move to purge position 
            '4) turn on valves
            Vision.FrmVision.SwitchCamera("View Camera")
            purgeLock = True
            If stopPurge Then
                GoTo Reset
            End If
            If Not m_Tri.Move_Z(SafePosition) Then GoTo Reset
            If stopPurge Then
                GoTo Reset
            End If
            If Not m_Tri.Move_XY(position) Then GoTo Reset
            If stopPurge Then
                GoTo Reset
            End If
            If Not m_Tri.Move_Z(offset(2)) Then GoTo Reset
            If stopPurge Then
                GoTo Reset
            End If
            LockButtonsForPurge()
            m_Tri.TurnOn("Left Needle IO")
            If stopPurge Then
                m_Tri.TurnOff("Left Needle IO")
                GoTo Reset
            End If
            purgeLock = False
            Exit Sub
        Else
            Production.ButtonPurge.Enabled = False
            Programming.ButtonPurge.Enabled = False
            stopPurge = True
            PurgingRunning = False
            m_Tri.AbortMotionDone()
            m_Tri.TurnOff("Left Needle IO")
            m_Tri.TrioStop()
            m_Tri.SetMachineStop()
            m_Tri.AbortMotionDone()
            'If Not m_Tri.Move_Z(SafePosition) Then GoTo Reset
            If Not (notMoveAfterPurge) Then
                If Not purgeLock Then
                    LockMovementButtons()
                    TravelToParkPosition()
                End If
            Else
                MoveZToSafePosition()
            End If

            Programming.ButtonPurge.Text = "Purge On"
            Production.ButtonPurge.Text = "Purge On"
            If Not (notMoveAfterPurge) Then
                ResetToIdle()
            End If
            SwapToTeachCam()
            Exit Sub

        End If

Reset:
        If Not (notMoveAfterPurge) Then
            LockMovementButtons()
            TravelToParkPosition()
            ResetToIdle()
            purgeLock = False
        End If
        SwapToTeachCam()

    End Sub
    Private Sub SwapToTeachCam()
        Vision.FrmVision.SwitchCamera("Teach Camera")
        Vision.FrmVision.DisplayIndicator()
        IDS.Devices.Vision.IDSV_SetBrightness(IDS.Data.Hardware.Camera.Brightness)
    End Sub
    Private Sub SwapToViewCam()
        Vision.FrmVision.SwitchCamera("View Camera")
        Vision.FrmVision.ClearDisplay()
    End Sub

    Private stopClean As Boolean = False
    Public CleaningRunning As Boolean = False
    'Private cleaning As Boolean = False
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
            'Else
            '    state = Programming.ButtonClean.Text
            '    If Programming.IsVisionTeachMode Then
            '        LabelMessage("Wrong mode. Please switch to needle mode first.")
            '        GoTo Reset
            '    End If
        End If

        If state = "Clean On" Then
            CleaningRunning = True
            stopClean = False
            Production.ButtonClean.Enabled = False
            Programming.ButtonClean.Enabled = False
            m_Tri.SetMachineRun()
            LockButtonsForClean()
            'Production.ButtonClean.Enabled = False
            'Programming.ButtonClean.Enabled = False
            Programming.ButtonClean.Text = "Clean Off"
            Production.ButtonClean.Text = "Clean Off"
            offset(0) = gLeftNeedleOffs(0)
            offset(1) = gLeftNeedleOffs(1)
            offset(2) = gLeftNeedleOffs(2) + IDS.Data.Hardware.Gantry.CleanPosition.Z
            position(0) = IDS.Data.Hardware.Gantry.CleanPosition.X - offset(0)
            position(1) = IDS.Data.Hardware.Gantry.CleanPosition.Y - offset(1)
            m_Tri.Set_XY_Speed(IDS.Data.Hardware.Gantry.ServiceXYSpeed)
            m_Tri.Set_Z_Speed(IDS.Data.Hardware.Gantry.ServiceZSpeed)
            Production.ButtonClean.Enabled = True
            Programming.ButtonClean.Enabled = True
            SwapToViewCam()
            If stopClean Then Exit Sub
            LabelMessage("Cleaning Moving Z Safety")
            If Not m_Tri.Move_Z(SafePosition) Then GoTo Reset
            LabelMessage("Cleaning Moving Z Done")
            If stopClean Then
                'cleaning = False
                LabelMessage("Exit Cleaning After Move Z Safety")
                Exit Sub
            End If
            LabelMessage("Cleaning Moving XY")
            If Not m_Tri.Move_XY(position) Then GoTo Reset
            LabelMessage("Cleaning Moving XY Done")
            If stopClean Then
                'cleaning = False
                LabelMessage("Exit Cleaning After Move XY")
                Exit Sub
            End If
            LabelMessage("Cleaning Moving Z Down")
            If Not m_Tri.Move_Z(offset(2)) Then GoTo Reset
            LabelMessage("Cleaning Moving Z Down Done")
            If stopClean Then
                'cleaning = False
                LabelMessage("Exit Cleaning After Move Z Down")
                Exit Sub
            End If
            LockButtonsForClean()
            m_Tri.TurnOn("Clean")
            If stopClean Then
                m_Tri.TurnOff("Clean")
                SwapToTeachCam()
                Exit Sub
            End If
            Exit Sub
        Else
            Production.ButtonClean.Enabled = False
            Programming.ButtonClean.Enabled = False
            stopClean = True
            CleaningRunning = False
            m_Tri.TrioStop()
            m_Tri.SetMachineStop()
            m_Tri.TurnOff("Clean")
            'If Not m_Tri.Move_Z(SafePosition) Then GoTo Reset
            LockMovementButtons()
            TravelToParkPosition()
            Production.ButtonClean.Text = "Clean On"
            Programming.ButtonClean.Text = "Clean On"
            SwapToTeachCam()
            ResetToIdle()
        End If

Reset:
        ResetToIdle()
        Production.ButtonClean.Enabled = True
        Programming.ButtonClean.Enabled = True
        SwapToTeachCam()
    End Sub

    Public Sub DoNeedleCalibration()
        NeedleCalibrationRunning = True
        'checks
        If ProductionMode() Then
            'Form_Service.LogEventInSPCReport("Needle Calibration")
            If Production.DoorCheck() = False Then
                LabelMessage("Close the door first.")
                GoTo Reset
            End If
        Else
            'If Programming.IsVisionTeachMode Then
            '    LabelMessage("Wrong mode. Please switch to needle mode first.")
            '    GoTo Reset
            'End If
        End If
        MyNeedleCalibrationSettings.NeedleCalibrationState = "Running"
        'execute
        LabelMessage("Needle calibration ...")
        LockMovementButtons()
        If ProgrammingMode() Then
            Programming.ButtonNeedleCal.Text = "Stop Cal. Need. "
        Else
            Production.ButtonNdlCalib.Text = "Stop Cal. Need."
        End If
        OnLaser()
        SwapToViewCam()
        Dim rtn As Boolean = MyNeedleCalibrationSettings.NeedleCalibration(False)
        OffLaser()
        If MyNeedleCalibrationSettings.NeedleCalibrationState = "Running" Then
            If rtn = True Then
                LabelMessage("Needle calibration successful.")
            Else
                m_Tri.Set_Z_Speed(IDS.Data.Hardware.Gantry.ServiceZSpeed)
                m_Tri.Move_Z(0)
                LabelMessage("Needle calibration failed.")
            End If
Reset:
            ResetToIdle()
            NeedleCalibrationRunning = False
            MyNeedleCalibrationSettings.NeedleCalibrationState = "Stopped"
        End If
        SwapToTeachCam()
        If ProgrammingMode() Then
            Programming.ButtonNeedleCal.Text = "Calibrate Needle"
        Else
            Production.ButtonNdlCalib.Text = "Calibrate Needle"
        End If
    End Sub
    Private changingSyringe As Boolean = False
    Public Sub DoChangeSyringe()

        'checks
        Dim isChangeSyringe As Boolean = False
        If ProductionMode() Then
            If Production.DoorCheck() = False Then
                LabelMessage("Close the door first.")
                GoTo Reset
            End If
            'If Production.ButtonChgSyringe.Text = "Change Syringe" Then
            '    isChangeSyringe = True
            '    Production.ButtonChgSyringe.Text = "Chg Syr Done"
            'Else
            '    isChangeSyringe = False
            '    Production.ButtonChgSyringe.Text = "Change Syringe"
            'End If
        Else
            'If Programming.IsVisionTeachMode Then
            '    LabelMessage("Wrong mode. Please switch to needle mode first.")
            '    GoTo Reset
            'End If
            isChangeSyringe = True
        End If
        changingSyringe = True
        'If isChangeSyringe Then
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
        Console.WriteLine("Bef Change Syringe Move Z")
        If Not m_Tri.Move_Z(SafePosition) Then GoTo Reset
        Console.WriteLine("After Change Syringe Move Z")
        If Not m_Tri.Move_XY(position) Then GoTo Reset
        If Not m_Tri.Move_Z(offset(2)) Then GoTo Reset
        LabelMessage("Stationed at change syringe position")
        Dim fm As InfoForm = New InfoForm
        fm.AddNewLine("Stationed at change syring position. Door was unlocked!")
        fm.AddNewLine("Only click Done button after changed syringe.")
        If IDS.Data.Hardware.Dispenser.Left.AutoPurgingOption Then
            fm.AddNewLine("Only click Done button after changed syringe.")
        End If
        fm.AddNewLine("Click Abort button if not to change syringe")
        fm.SetOKButtonText("Done")
        fm.SetCancelButtonText("Abort")
        UnlockDoor()
        Dim dR As DialogResult = fm.ShowDialog()
        m_Tri.SetMachineStop()
        If dR = DialogResult.Cancel Then
            LockDoor()
            LockMovementButtons()
            TravelToParkPosition()
        Else
            If IDS.Data.Hardware.Dispenser.Left.AutoPurgingOption Then
                LabelMessage("Auto purging enabled: Auto purging timer was reset after syringe changed.")
                Production.ResetTimer("Reset Autopurging Timers")
            End If
            Dim db As InfoForm = New InfoForm
            db.SetTitle("Needle Calibration")
            db.AddNewLine("Would you like to perform purge and needle calibration?")
            db.AddNewLine("Click Ok button to perform purge first")
            db.AddNewLine("Click Cancel button if not to calibrate needle.")
            If db.ShowDialog() = DialogResult.OK Then
                Do
                    If IDS.Devices.DIO.DIO.doorclose_flag Then Exit Do
                    db.SetMessage("")
                    db.AddNewLine("Please Close the door!")
                    db.OkOnly()
                    db.ShowDialog()
                Loop Until IDS.Devices.DIO.DIO.doorclose_flag
                LockDoor()
                'ResetToIdle()
                LabelMessage("Purging before Needle calibration")
                If SetState("Do Purge") Then
                    DoPurge()
                End If
                Dim purgeFm As InfoForm = New InfoForm
                purgeFm.AddNewLine("Performing purging now")
                purgeFm.AddNewLine("Click Off to stop purge otherwise cancel needle calibration")
                purgeFm.OkOnly()
                purgeFm.SetOKButtonText("Off")
                dR = purgeFm.ShowDialog()
                If SetState("Do Purge") Then
                    DoPurge(True)
                End If
                If Not (dR = DialogResult.OK) Then
                    LabelMessage("Needle calibration cancled after purging")
                    LockMovementButtons()
                    TravelToParkPosition()
                    ResetToIdle()
                    changingSyringe = False
                    Exit Sub
                Else
                    LabelMessage("Purging off")
                End If
                Dim cleanNeedle As InfoForm = New InfoForm
                cleanNeedle.AddNewLine("Would you like to clean needle (manually) first before proceed")
                cleanNeedle.AddNewLine("to do needle calibration?")
                cleanNeedle.SetOKButtonText("Yes")
                cleanNeedle.SetCancelButtonText("No")
                dR = cleanNeedle.ShowDialog()
                If dR = DialogResult.OK Then
                    'unlock the lock
                    UnlockDoor()
                    Dim doorForm As InfoForm = New InfoForm
                    doorForm.AddNewLine("Door is unlocked!")
                    doorForm.AddNewLine("Please open the door and clean the needle tip.")
                    doorForm.AddNewLine("Please close the door before click done")
                    doorForm.SetOKButtonText("Done")
                    doorForm.OkOnly()
                    doorForm.ShowDialog()
                    Do
                        If IDS.Devices.DIO.DIO.doorclose_flag Then Exit Do
                        doorForm.SetMessage("")
                        doorForm.AddNewLine("Please close the door!")
                        doorForm.ShowDialog()
                    Loop Until IDS.Devices.DIO.DIO.doorclose_flag
                    LockDoor()
                End If
                Dim calibFm As InfoForm = New InfoForm
                calibFm.AddNewLine("Continue with needle calibration?")
                calibFm.AddNewLine("Click Continue for needle calibration.")
                calibFm.AddNewLine("Click Cancel to abort.")
                purgeFm.SetOKButtonText("Continue")
                dR = calibFm.ShowDialog()
                If dR = DialogResult.OK Then
                    If ProgrammingMode() Then
                        ResetToIdle()
                        SetState("Needle Calibration")
                        Programming.ButtonNeedleCal.Text = "Stop Cal. Need. "
                        Programming.ButtonNeedleCal.Enabled = True
                        changingSyringe = False
                        Return
                    Else
                        ResetToIdle()
                        SetState("Needle Calibration")
                        Production.ButtonNdlCalib.Text = "Stop Cal. Need."
                        Production.ButtonNdlCalib.Enabled = True
                        changingSyringe = False
                        Return
                    End If
                Else
                    LockMovementButtons()
                    TravelToParkPosition()
                End If

            Else
                LockMovementButtons()
                TravelToParkPosition()
            End If
        End If

        'ResetToIdle()
        'Production.ButtonChgSyringe.Enabled = True
        'Return
        'End If
        'If Not (isChangeSyringe) Then
        '    m_Tri.SetMachineStop()
        '    LockMovementButtons()
        '    TravelToParkPosition()
        '    ResetToIdle()
        'End If

Reset:
        ResetToIdle()
        changingSyringe = False
        If ProductionMode() Then
            Production.ButtonVolCalib.Text = "Vol. Cal"
        End If
    End Sub

    Public Sub DoVolumeCalibration()
        VolumeCalibrationRunning = True
        MyVolumeCalibrationSettings.VolumeCalibrationState = "Running"
        MyVolumeCalibrationSettings.VCRunning = True
        If ProductionMode() Then
            If Production.DoorCheck() = False Then
                LabelMessage("Close the door first.")
                GoTo Reset
            End If
            Form_Service.LogEventInSPCReport("Volume Calibration")
        End If

        LockMovementButtons()
        If ProgrammingMode() Then
            Programming.ButtonVolCal.Text = "Stop Cal. Vol."
        Else
            Production.ButtonVolCalib.Text = "Stop Cal. Vol."
        End If
        MyVolumeCalibrationSettings.VCStatusUpdate = New VolumeCalibrationSettings.VCStatusUpdateDelegate(AddressOf updateLabelMessage)
        MyVolumeCalibrationSettings.VolumeCalibrationResult = ""
        LabelMessage("Doing volume calibration.")
        ' MyVolumeCalibrationSettings.VolumeCalibrationSetup()
        If Not (MyVolumeCalibrationSettings.VolumeCalibrationState = "Stopped") Then
            'VolumeCalibrationRunning = True
            m_Tri.SetMachineRun()
            MyVolumeCalibrationSettings.VolumeCalibrationParamSetup()
            Dim rtn As Boolean = MyVolumeCalibrationSettings.VolumeCalibration(IDS.Data.Hardware.Volume.Left.NumberOfAttempts)
            'If rtn = True Then
            '    LabelMessage("Volume calibration successful.")
            '    If ProductionMode() Then
            '        Production.DisplayAdjustedPressure()
            '    End If
            'Else
            '    If Not (MyVolumeCalibrationSettings.VolumeCalibrationState = "Stopped") Then
            '        LabelMessage("Volume calibration failed.")
            '    Else
            '        LabelMessage("Volume calibration canceled.")
            '    End If
            'End If
            'If ProgrammingMode() Then
            '    Programming.ButtonVolCal.Text = "Vol. Cal."
            'Else
            '    Production.ButtonVolCalib.Text = "Vol. Cal."
            'End If
            'Reset:
            '            If Not (MyVolumeCalibrationSettings.VolumeCalibrationState = "Stopped") Then
            '                ResetToIdle()
            '            Else
            '                SetState("Idle")
            '            End If
            '            If ProgrammingMode() Then
            '                Programming.ButtonVolCal.Text = "Vol. Cal."
            '            Else
            '                Production.ButtonVolCalib.Text = "Vol. Cal."
            '            End If
            '            MyVolumeCalibrationSettings.VolumeCalibrationState = "Stopped"
            '            MyVolumeCalibrationSettings.VCRunning = False
            '            VolumeCalibrationRunning = False
        Else
            LabelMessage("Volume calibration canceled.")
            Weighting_Scale.RequestWeightUpdate = False
            MyVolumeCalibrationSettings.VCRunning = False
            VolumeCalibrationRunning = False
            If ProgrammingMode() Then
                Programming.ButtonVolCal.Text = "Calibrate Volume"
            Else
                Production.ButtonVolCalib.Text = "Calibrate Volume"
            End If
        End If
        Exit Sub

Reset:
        If Not (MyVolumeCalibrationSettings.VolumeCalibrationState = "Stopped") Then
            ResetToIdle()
        Else
            SetState("Idle")
        End If
        If ProgrammingMode() Then
            Programming.ButtonVolCal.Text = "Calibrate Volume"
        Else
            Production.ButtonVolCalib.Text = "Calibrate Volume"
        End If
        MyVolumeCalibrationSettings.VolumeCalibrationState = "Stopped"
        MyVolumeCalibrationSettings.VCRunning = False
        VolumeCalibrationRunning = False
    End Sub

    Public Sub VCResult()
        If MyVolumeCalibrationSettings.VolumeCalibrationState.ToUpper = "SUCCESS" Then
            'LabelMessage("Volume calibration successful.")
            If ProductionMode() Then
                Production.DisplayAdjustedPressure()
            End If
        Else
            If Not (MyVolumeCalibrationSettings.VolumeCalibrationState = "Stopped") Then
                LabelMessage("Volume calibration failed.")
            Else
                LabelMessage("Volume calibration canceled.")
            End If
        End If
        If ProgrammingMode() Then
            Programming.ButtonVolCal.Text = "Calibrate Volume"
        Else
            Production.ButtonVolCalib.Text = "Calibrate Volume"
        End If
Reset:
        If Not (MyVolumeCalibrationSettings.VolumeCalibrationState = "Stopped") Then
            ResetToIdle()
        Else
            SetState("Idle")
        End If
        If ProgrammingMode() Then
            Programming.ButtonVolCal.Text = "Calibrate Volume"
        Else
            Production.ButtonVolCalib.Text = "Calibrate Volume"
        End If
        MyVolumeCalibrationSettings.VolumeCalibrationState = "Stopped"
        MyVolumeCalibrationSettings.VCRunning = False
        VolumeCalibrationRunning = False
    End Sub

    Public Sub PauseDispensing()
        If IsIdle() Then Return
        If VolumeCalibrationRunning = True Then Return
        SetState("Pause")
        m_Tri.SetMachinePause()
        ChangeButtonState("Paused")
        LabelMessage("Dispensing Pause!")
        If VolumeCalibrationRunning = True Then MyVolumeCalibrationSettings.VolumeCalibrationState = "Paused"
        If ShouldLog() Then Form_Service.LogEventInSPCReport("Pause")
    End Sub
    Public stopping As Boolean = False
    Public Sub StopDispensing()
        If changingSyringe Then Exit Sub
        If IsIdle() Then Exit Sub 'Prevent hardware stop button to trigger stop when system idle
        'm_Tri.AbortDispensing()
        stopping = True
        If MyVolumeCalibrationSettings.VolumeCalibrationState = "Running" Then
            LabelMessage("Stopping volume calibration")
            MyVolumeCalibrationSettings.VolumeCalibrationState = "Stopped"
            If ProgrammingMode() Then
                Programming.ButtonVolCal.Enabled = False
                Programming.ButtonVolCal.Text = "Calibrate Volume"
            Else
                Production.ButtonVolCalib.Enabled = False
                Production.ButtonVolCalib.Text = "Calibrate Volume"
            End If
        End If
        If CleaningRunning Then
            stopClean = True
            CleaningRunning = False
            m_Tri.TurnOff("Clean")
            LabelMessage("Stopping Cleaning Process")
            If ProgrammingMode() Then
                Programming.ButtonClean.Enabled = False
                Programming.ButtonClean.Text = "Clean On"
            Else
                Production.ButtonClean.Enabled = False
                Production.ButtonClean.Text = "Clean On"
            End If
        End If
        If PurgingRunning Then
            stopPurge = True
            PurgingRunning = False
            m_Tri.TurnOff("Left Needle IO")
            LabelMessage("Stopping Purging Process")
            If ProgrammingMode() Then
                Programming.ButtonPurge.Enabled = False
                Programming.ButtonPurge.Text = "Purge On"
            Else
                Production.ButtonPurge.Enabled = False
                Production.ButtonPurge.Text = "Purge On"
            End If
        End If
        'VolumeCalibrationRunning = False
        'LabelMessage("Dispensing Stop!")
        If MyNeedleCalibrationSettings.NeedleCalibrationState = "Running" Then 'NeedleCalibrationRunning Then
            MyNeedleCalibrationSettings.NeedleCalibrationState = "Stopped"
            LabelMessage("Stopping needle calibration")
            Dim pos(0) As Integer
            pos(0) = 1
            m_Tri.m_TriCtrl.SetTable(126, 1, pos)
            'm_Tri.AbortMotionDone()
            'LabelMessage("Stopping Needle Calibration")
            If ProgrammingMode() Then
                Programming.ButtonNeedleCal.Enabled = False
                Programming.ButtonNeedleCal.Text = "Calibrate Needle"
            Else
                Production.ButtonNdlCalib.Enabled = False
                Production.ButtonNdlCalib.Text = "Calibrate Needle"
            End If
        End If
        While MyVolumeCalibrationSettings.VCRunning
            Application.DoEvents()
        End While
        While MyNeedleCalibrationSettings.NCRunning
            Application.DoEvents()
        End While
        'SetState("Stop")
        m_Tri.AbortMotionDone()
        'If Not (m_Tri.WaitAbortDispensing()) Then
        '    MessageBox.Show("Abort dispensing timeout")
        'End If
        m_Tri.m_TriCtrl.Execute("OP(25,0)")
        If Not (m_Tri.TrioStop()) Then
            MessageBox.Show("Unable to stop the dispensing process now. Try again later!")
            Return
        End If
        SetState("Stop")
        m_Tri.AbortMotionDone()
        If Not ThreadExecution Is Nothing Then ThreadExecution.Abort()
        'If VolumeCalibrationRunning Then
        'Turn of dispensing output

        'End If
        m_Tri.SetQCStop()
        m_Tri.ResetQCFlag()

        Dim enterTime As Long = DateTime.Now.Ticks
        While ((DateTime.Now.Ticks - enterTime) / 1000) > 1000
            Application.DoEvents()
        End While
        'Console.WriteLine("After abort motion done")
        If Not IsIdle() Then 'If Not WasStart() Or VolumeCalibrationRunning Or NeedleCalibrationRunning Then
            enterTime = DateTime.Now.Ticks
            If localRunning Then
                LabelMessage("Waiting for local process to be abort", False, False)
            End If
            While localRunning And ((DateTime.Now.Ticks - enterTime) / 1000) < 10000
                Application.DoEvents()
            End While
            localRunning = False
            LockMovementButtons()
            'Console.WriteLine("#c")
            LabelMessage("#Stop Dispense", False, False)
            TravelToParkPosition()
            'Console.WriteLine("#d")
            ResetToIdle()
            'Console.WriteLine("#e")
            Programming.ButtonVolCal.Text = "Calibrate Volume"
            Programming.ButtonNeedleCal.Text = "Calibrate Needle"
        End If
        m_Tri.ResetAbortDispensingFlag()
        VolumeCalibrationRunning = False
        NeedleCalibrationRunning = False
        'Console.WriteLine("#a")
        If ProductionMode() Then
            If Production.ContinuousMode.Checked Then
                Conveyor.Command("Release")
            End If
            Production.WriteSPCReport()
        End If
        stopping = False
        'Console.WriteLine("#b")
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
        Vision.FrmVision.SwitchCamera("Teach Camera")
        'ClearAndDisplayIndicator()
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
        If ProductionMode() Then
            Production.DoorLock.Text = "Lock Door"
        Else
            Programming.CBDoorLock.Text = "Lock Door"
        End If
    End Function

    Public Function LockDoor()
        IDS.Devices.DIO.DIO.DoorLock(False) 'lock is false
        If ProductionMode() Then
            Production.DoorLock.Text = "Unlock Door"
        Else
            Programming.CBDoorLock.Text = "Unlock Door"
        End If
    End Function

    Public Sub OnLaser()
        If Laser Is Nothing Then Return
        Vision.FrmVision.SetLaser(False) 'turn on
        'Laser.ClearComBuffer()
        Laser.SendCommand("TurnOnMeasurement")
    End Sub

    Public Sub OffLaser()
        If Laser Is Nothing Then Return
        Vision.FrmVision.SetLaser(True) 'turn offset
        Laser.SendCommand("TurnOffMeasurement")
        'Laser.ClearComBuffer()
    End Sub

#End Region

    Public Sub updateLabelMessage(ByVal status As String)
        LabelMessage(status)
    End Sub

End Module