
Public Class CIDSVision
    Public FrmVision As FormVision = FrmVision.Instance
    Public Shared LightBox = _LightBox
    Public QC_form As QC = QC.Instance
    Public ChipEdgePoints_form As ChipEdgePoints = ChipEdgePoints.Instance
    Public RejectPoint_form As RejectPoint = RejectPoint.Instance
    Public FiducialMark_form As FiducialForm = FiducialMark_form.Instance

    Private Shared m_instance As CIDSVision

    Public Shared ReadOnly Property Instance() As CIDSVision
        Get
            If m_instance Is Nothing Then
                m_instance = New CIDSVision
            Else
            End If
            Return m_instance
        End Get

    End Property

#Region "SystemSetup"
    Public Function IDSV_SS(ByVal ColumnsNo As Integer, ByVal RowsNo As Integer, ByVal PitchSpacing As Double) As Integer
        FrmVision.CameraCalibration(ColumnsNo, RowsNo, PitchSpacing)
    End Function

#End Region

#Region "Production"
    Public Function IDSV_FI(ByVal Filename As String, ByRef FI_OffX As Double, ByRef FI_OffY As Double, Optional ByVal SkipWaitStable As Boolean = False) As Boolean
        Return FrmVision.IDSV_FI(Filename, FI_OffX, FI_OffY, SkipWaitStable)
    End Function
    Public Sub IDSV_FIOutput(ByRef DelX As Double, ByRef DelY As Double) 'useless
        FrmVision.IDSV_FI(DelX, DelY)
    End Sub
    Public lastError As String = ""
    Public Function IDSV_CE(ByVal vParam As DLL_Export_Device_Vision.ChipEdgePoints.ChipEdgeParam) As Boolean
        Dim rtn As Boolean = FrmVision.IDSV_CE(vParam)
        lastError = FrmVision.lastError
        Return rtn
    End Function
    Public Sub IDSV_CEOutput(ByRef PointX1, ByRef PointY1, ByRef PointX2, ByRef PointY2, ByRef PointX3, ByRef PointY3, ByRef PointX4, ByRef PointY4)
        FrmVision.IDSV_CE(PointX1, PointY1, PointX2, PointY2, PointX3, PointY3, PointX4, PointY4)
    End Sub
    Public Function IDSV_NC(ByVal BlackDot As Boolean, ByVal Binarized As Integer, ByVal MaxArea As Double, ByVal MinArea As Double, ByVal Close As Integer, ByVal Open As Integer, ByVal Roughness As Double, ByVal Compactness As Double, ByVal _DisplayCenterXPosition As Decimal, ByVal _DisplayCenterYPosition As Decimal, ByVal _MRoiX As Decimal, ByVal _MRoiY As Decimal, ByRef NC_OffX As Double, ByRef NC_OffY As Double) As Boolean
        Return FrmVision.IDSV_NC(BlackDot, Binarized, MaxArea, MinArea, Close, Open, Roughness, Compactness, _DisplayCenterXPosition, _DisplayCenterYPosition, _MRoiX, _MRoiY, NC_OffX, NC_OffY)
    End Function
    Public diameterResult As Double = 0
    Public Function IDSV_QC(ByVal VParam As DLL_Export_Device_Vision.QC.QCParam, Optional ByVal skipWaitStable As Boolean = False) As Boolean
        Dim rtn = FrmVision.IDSV_QC(VParam, skipWaitStable)
        If rtn Then
            diameterResult = FrmVision.diameterResult
        Else
            diameterResult = -1
        End If
        Return rtn
    End Function
    Public Function IDSV_RM(ByVal VParam As DLL_Export_Device_Vision.RejectPoint.RMParam) As Boolean
        Return FrmVision.IDSV_RM(VParam)
    End Function
#End Region

#Region "Edit"

    Public Function IDSV_Form_FI_Edit(ByVal FI_No As Integer, ByVal Filename As String, ByVal brightness As Integer)
        Return FrmVision.Form_FI_Edit(FI_No, Filename, brightness)
    End Function
    Public Function IDSV_Form_CE_Edit(ByVal vParam As DLL_Export_Device_Vision.ChipEdgePoints.ChipEdgeParam) As Boolean
        Return FrmVision.Form_CE_Edit(vParam)
    End Function
    Public Function IDSV_Form_QC_Edit(ByVal VParam As DLL_Export_Device_Vision.QC.QCParam) As Boolean
        Return FrmVision.Form_QC_Edit(VParam)
    End Function
    Public Function IDSV_Form_RM_Edit(ByVal VParam As DLL_Export_Device_Vision.RejectPoint.RMParam) As Boolean
        Return FrmVision.Form_RM_Edit(VParam)
    End Function
#End Region

#Region "Programming"
#Region "Fi"
    Public Sub IDSV_Form_FI(ByVal FI_No As Integer, ByVal brightness As Integer)
        FrmVision.ShowFiducialForm(FI_No, brightness)
    End Sub
    Function GetFiducialStatus() As Integer
        Return FiducialMark_form.GetFiducialStatus()
    End Function
    Function SetFiducialFilename(ByVal Filename As String)
        FrmVision.SetFiducialFilename(Filename)
    End Function
    Function GetFiducialFilename() As String
        Return FiducialMark_form.GetFiducialFilename
    End Function
    Function GetFiducialBrightness() As Integer
        Return FiducialMark_form.BrightnessValue.Value()
    End Function
#End Region
#Region "CE"
    Public Sub IDSV_Form_CE(ByVal brightness As Double)
        FrmVision.Form_CE(brightness)
    End Sub
    Function GetChipEdgeStatus() As Integer
        Return ChipEdgePoints_form.GetChipEdgeStatus()
    End Function
    Function GetChipEdgeParameters(ByRef param As DLL_Export_Device_Vision.ChipEdgePoints.ChipEdgeParam)
        FrmVision.GetChipEdgeParameters(param)
    End Function
    'Function GetChipEdgeParameters(ByRef _SizeX As Double, ByRef _SizeY As Double, ByRef _PosX As Double, ByRef _PosY As Double, ByRef _Size As Double, ByRef _Pos As Double, ByRef _Rot As Double, ByRef _Inside_out As Boolean, ByVal _Cw_CCw As Boolean, ByRef _Threshold As Double, ByRef _ROI As Double, ByRef _Brightness As Integer, ByRef _Vertical As Boolean, ByRef _MainEdge As Integer, ByRef _PointX1 As Double, ByRef _PointY1 As Double, ByRef _PointX2 As Double, ByRef _PointY2 As Double, ByRef _PointX3 As Double, ByRef _PointY3 As Double, ByRef _PointX4 As Double, ByRef _PointY4 As Double, ByRef _PointX5 As Double, ByRef _PointY5 As Double, ByRef _DispenseModel As Integer, ByRef _EdgeClearance As Double, ByRef _CheckBox_ChipRec_Enable As Boolean)
    '    FrmVision.GetChipEdgeParameters(_SizeX, _SizeY, _PosX, _PosY, _Size, _Pos, _Rot, _Inside_out, _Cw_CCw, _Threshold, _ROI, _Brightness, _Vertical, _MainEdge, _PointX1, _PointY1, _PointX2, _PointY2, _PointX3, _PointY3, _PointX4, _PointY4, _PointX5, _PointY5, _DispenseModel, _EdgeClearance, _CheckBox_ChipRec_Enable)
    'End Function
    Function RobotMotionOffset(ByRef _MoveOffsetXmm As Double, ByRef _MoveOffsetYmm As Double) As Boolean
        Return FrmVision.RobotMotionOffset(_MoveOffsetXmm, _MoveOffsetYmm)
    End Function
    Sub SetCEReset()
        ChipEdgePoints_form.ResetVariables()
    End Sub
#End Region
#Region "QC"
    Public Sub IDSV_Form_QC(ByVal brightness As Double)
        FrmVision.Form_QC(brightness)
    End Sub
    Function GetQCStatus() As Integer
        Return QC_form.GetQCStatus()
    End Function
    Function GetIsGlobalQC() As Boolean
        Return QC_form.GetIsGlobalQC()
    End Function
    Function SetAllowGlobalQC(ByVal allow As Boolean)
        QC_form.SetAllowGlobalMode(allow)
    End Function
    Function SetGlobalQCMode(ByVal isGlobal As Boolean)
        QC_form.SetGlobalMode(isGlobal)
    End Function
    Function GetQCParameters(ByRef param As DLL_Export_Device_Vision.QC.QCParam)
        Return QC_form.GetQCParameters(param)
    End Function
    Function GetQCSuccess() As String
        Return QC_form.QC_success
    End Function
    Function ReportQCFailure() As String
        Return QC_form.QC_failure
    End Function
    Sub SetQCReset()
        QC_form.SetQCReset()
    End Sub
#End Region
#Region "Reject Mark"
    Public Sub IDSV_Form_RM(ByVal brightness As Double)
        FrmVision.RejectPoint_flag = True
        RejectPoint_form.ResetRMStatus()
        FrmVision.RejectPoints()
        SetRMReset()
        FrmVision.ClearDisplay()
        FrmVision.modelregionDrawing()
        FrmVision.SetBrightness(brightness)
        RejectPoint_form.ValueBrightness1.Value = brightness
        RejectPoint_form.ValueBrightness1.Text = brightness
        RejectPoint_form.Button_Next.Enabled = True
        If RejectPoint_form.Visible = False Then
            RejectPoint_form.Visible = True
        End If
    End Sub
    Function GetRMStatus() As Integer
        Return FrmVision.GetRMStatus
    End Function
    Function GetRMParameters(ByRef param As DLL_Export_Device_Vision.RejectPoint.RMParam)
        Return FrmVision.GetRMParameters(param)
    End Function
    Sub SetRMReset()
        FrmVision.SetRMReset()
    End Sub
#End Region
    Public Sub IDSV_Brightness()
        FrmVision.Brightness()
    End Sub
    Public Sub IDSV_SetBrightness(ByVal a As Integer)
        FrmVision.SetBrightness(a)
    End Sub
    Public Sub IDSV_SetLaser(ByVal a As Boolean)
        FrmVision.SetLaser(a) 'true is to turn off, false is to turn on
    End Sub
    Public Sub IDSV_SaveImage()
        FrmVision.SaveImage()
    End Sub
    Public Sub IDSV_DisplayIndicator()
        FrmVision.DisplayIndicator()
    End Sub
#End Region

    Public Sub callFunctionHere()
        Dim x = MyVisionCounter
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
    Public FrmVision As FormVision = CIDSVision.Instance.FrmVision.Instance
    Public QC_form As QC = CIDSVision.Instance.QC_form.Instance
    Public ChipEdgePoints_form As ChipEdgePoints = CIDSVision.Instance.ChipEdgePoints_form.Instance
    Public RejectPoint_form As RejectPoint = CIDSVision.Instance.RejectPoint_form.Instance
    Public FiducialMark_form As FiducialForm = CIDSVision.Instance.FiducialMark_form.Instance
    Friend _LightBox As New DeviceLighbox
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
