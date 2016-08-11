Public Class CIDSTrioController
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents BTStepZdown As System.Windows.Forms.Button
    Friend WithEvents BTStepZup As System.Windows.Forms.Button
    Friend WithEvents BTStepXminus As System.Windows.Forms.Button
    Friend WithEvents BTStepXplus As System.Windows.Forms.Button
    Friend WithEvents BTStepYminus As System.Windows.Forms.Button
    Friend WithEvents BTStepYplus As System.Windows.Forms.Button
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents LabelStep As System.Windows.Forms.Label
    Friend WithEvents ComboBoxFineStep As System.Windows.Forms.NumericUpDown
    Public WithEvents SteppingButtons As System.Windows.Forms.Panel
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(CIDSTrioController))
        Me.SteppingButtons = New System.Windows.Forms.Panel
        Me.ComboBoxFineStep = New System.Windows.Forms.NumericUpDown
        Me.BTStepZdown = New System.Windows.Forms.Button
        Me.BTStepZup = New System.Windows.Forms.Button
        Me.BTStepXminus = New System.Windows.Forms.Button
        Me.BTStepXplus = New System.Windows.Forms.Button
        Me.BTStepYminus = New System.Windows.Forms.Button
        Me.BTStepYplus = New System.Windows.Forms.Button
        Me.Label6 = New System.Windows.Forms.Label
        Me.LabelStep = New System.Windows.Forms.Label
        Me.SteppingButtons.SuspendLayout()
        CType(Me.ComboBoxFineStep, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'SteppingButtons
        '
        Me.SteppingButtons.BackColor = System.Drawing.SystemColors.Control
        Me.SteppingButtons.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.SteppingButtons.Controls.Add(Me.ComboBoxFineStep)
        Me.SteppingButtons.Controls.Add(Me.BTStepZdown)
        Me.SteppingButtons.Controls.Add(Me.BTStepZup)
        Me.SteppingButtons.Controls.Add(Me.BTStepXminus)
        Me.SteppingButtons.Controls.Add(Me.BTStepXplus)
        Me.SteppingButtons.Controls.Add(Me.BTStepYminus)
        Me.SteppingButtons.Controls.Add(Me.BTStepYplus)
        Me.SteppingButtons.Controls.Add(Me.Label6)
        Me.SteppingButtons.Controls.Add(Me.LabelStep)
        Me.SteppingButtons.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.SteppingButtons.Location = New System.Drawing.Point(20, 23)
        Me.SteppingButtons.Name = "SteppingButtons"
        Me.SteppingButtons.Size = New System.Drawing.Size(336, 280)
        Me.SteppingButtons.TabIndex = 23
        '
        'ComboBoxFineStep
        '
        Me.ComboBoxFineStep.DecimalPlaces = 3
        Me.ComboBoxFineStep.Location = New System.Drawing.Point(152, 240)
        Me.ComboBoxFineStep.Maximum = New Decimal(New Integer() {200, 0, 0, 0})
        Me.ComboBoxFineStep.Name = "ComboBoxFineStep"
        Me.ComboBoxFineStep.Size = New System.Drawing.Size(88, 27)
        Me.ComboBoxFineStep.TabIndex = 9
        Me.ComboBoxFineStep.Value = New Decimal(New Integer() {5, 0, 0, 131072})
        '
        'BTStepZdown
        '
        Me.BTStepZdown.BackColor = System.Drawing.Color.Linen
        Me.BTStepZdown.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.5!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.BTStepZdown.Image = CType(resources.GetObject("BTStepZdown.Image"), System.Drawing.Image)
        Me.BTStepZdown.ImageAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.BTStepZdown.Location = New System.Drawing.Point(264, 128)
        Me.BTStepZdown.Name = "BTStepZdown"
        Me.BTStepZdown.Size = New System.Drawing.Size(48, 72)
        Me.BTStepZdown.TabIndex = 8
        Me.BTStepZdown.Text = "Dn"
        Me.BTStepZdown.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'BTStepZup
        '
        Me.BTStepZup.BackColor = System.Drawing.Color.Linen
        Me.BTStepZup.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.5!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.BTStepZup.Image = CType(resources.GetObject("BTStepZup.Image"), System.Drawing.Image)
        Me.BTStepZup.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.BTStepZup.Location = New System.Drawing.Point(264, 40)
        Me.BTStepZup.Name = "BTStepZup"
        Me.BTStepZup.Size = New System.Drawing.Size(48, 72)
        Me.BTStepZup.TabIndex = 7
        Me.BTStepZup.Text = "Up"
        Me.BTStepZup.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'BTStepXminus
        '
        Me.BTStepXminus.BackColor = System.Drawing.Color.LightGoldenrodYellow
        Me.BTStepXminus.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.5!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.BTStepXminus.Image = CType(resources.GetObject("BTStepXminus.Image"), System.Drawing.Image)
        Me.BTStepXminus.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BTStepXminus.Location = New System.Drawing.Point(24, 96)
        Me.BTStepXminus.Name = "BTStepXminus"
        Me.BTStepXminus.Size = New System.Drawing.Size(80, 48)
        Me.BTStepXminus.TabIndex = 6
        Me.BTStepXminus.Text = "X-"
        Me.BTStepXminus.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'BTStepXplus
        '
        Me.BTStepXplus.BackColor = System.Drawing.Color.LightGoldenrodYellow
        Me.BTStepXplus.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.5!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.BTStepXplus.Image = CType(resources.GetObject("BTStepXplus.Image"), System.Drawing.Image)
        Me.BTStepXplus.ImageAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.BTStepXplus.Location = New System.Drawing.Point(152, 96)
        Me.BTStepXplus.Name = "BTStepXplus"
        Me.BTStepXplus.Size = New System.Drawing.Size(80, 48)
        Me.BTStepXplus.TabIndex = 5
        Me.BTStepXplus.Text = "X+"
        Me.BTStepXplus.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'BTStepYminus
        '
        Me.BTStepYminus.BackColor = System.Drawing.Color.LightGoldenrodYellow
        Me.BTStepYminus.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.5!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.BTStepYminus.Image = CType(resources.GetObject("BTStepYminus.Image"), System.Drawing.Image)
        Me.BTStepYminus.ImageAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.BTStepYminus.Location = New System.Drawing.Point(104, 144)
        Me.BTStepYminus.Name = "BTStepYminus"
        Me.BTStepYminus.Size = New System.Drawing.Size(48, 80)
        Me.BTStepYminus.TabIndex = 4
        Me.BTStepYminus.Text = "Y-"
        Me.BTStepYminus.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'BTStepYplus
        '
        Me.BTStepYplus.BackColor = System.Drawing.Color.LightGoldenrodYellow
        Me.BTStepYplus.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.5!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.BTStepYplus.Image = CType(resources.GetObject("BTStepYplus.Image"), System.Drawing.Image)
        Me.BTStepYplus.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.BTStepYplus.Location = New System.Drawing.Point(104, 16)
        Me.BTStepYplus.Name = "BTStepYplus"
        Me.BTStepYplus.Size = New System.Drawing.Size(48, 80)
        Me.BTStepYplus.TabIndex = 3
        Me.BTStepYplus.Text = "Y+"
        Me.BTStepYplus.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'Label6
        '
        Me.Label6.BackColor = System.Drawing.SystemColors.Control
        Me.Label6.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label6.Location = New System.Drawing.Point(248, 240)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(40, 24)
        Me.Label6.TabIndex = 2
        Me.Label6.Text = "mm"
        Me.Label6.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'LabelStep
        '
        Me.LabelStep.BackColor = System.Drawing.SystemColors.Control
        Me.LabelStep.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.LabelStep.Location = New System.Drawing.Point(56, 240)
        Me.LabelStep.Name = "LabelStep"
        Me.LabelStep.Size = New System.Drawing.Size(88, 24)
        Me.LabelStep.TabIndex = 0
        Me.LabelStep.Text = "Fine Step:"
        Me.LabelStep.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'CIDSTrioController
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(376, 326)
        Me.Controls.Add(Me.SteppingButtons)
        Me.Name = "CIDSTrioController"
        Me.Text = "Axes Relative Move"
        Me.SteppingButtons.ResumeLayout(False)
        CType(Me.ComboBoxFineStep, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region


    'DIO list
    Public gIOXminusLimit As Integer = 0
    Public gIOXplusLimit As Integer = 1
    Public gIOYminusLimit As Integer = 2
    Public gIOYplusLimit As Integer = 3
    Public gIOZminusLimit As Integer = 5
    Public gIOZplusLimit As Integer = 4
    Public gIOXDatum As Integer = 0
    Public gIOYDatum As Integer = 2
    Public gIOZDatum As Integer = 5

    Public gIOLVDTProbe As Integer = 8
    Public gIORNeedleSlide As Integer = 17 '33 output
    Public gIOLNeedleSlide As Integer = 18 '34 output
    Public gIORNeedleSlideExt As Integer = 11 '33 intput
    Public gIOLNeedleSlideExt As Integer = 9 '34 input

    Public gIOClean As Integer = 19
    Public gIORightNeedle As Integer = 24
    Public gIOLeftNeedle As Integer = 25
    Public gIORightNeedleValve As Integer = 20
    Public gIOLeftNeedleValve As Integer = 21
    Public gIODispensingDone As Integer = 26
    Public gIOBoardReady As Integer = 12
    Public gMaterialAir As Integer = 27

    Public Declare Sub Sleep Lib "kernel32" Alias "Sleep" (ByVal dwMilliseconds As Long)  'delay function
    Public Const gTrioLowerInput As Integer = 0    'Trio onboard digital IO input lower bound no.
    Public Const gTrioUpperInput As Integer = 15   'Trio onboard digital IO input upper bound no.
    Public Const gCANLowerInput As Integer = 16    'Trio CAN digital IO input lower bound no.
    Public Const gCANUpperInput As Integer = 31    'Trio CAN digital IO input upper bound no.
    Public Const gTrioLowerOutput As Integer = 8   'Trio onboard digital IO output lower bound no.
    Public Const gTrioUpperOutput As Integer = 15  'Trio onboard digital IO output upper bound no.
    Public Const gCANLowerOutput As Integer = 16   'Trio CAN digital IO output lower bound no.
    Public Const gCANUpperOutput As Integer = 31   'Trio CAN digital IO output upper bound no.

    Public m_TriCtrl As New TrioPCLib.TrioPCClass
    Public StateContainer(250) As Double
    Public HomeStateContainer(0) As Double


#Region " Connection "

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Connecting TrioMotion controller via ethernet                    '
    '   address:  TrioMotion controller's IP address                  '
    '   pmode:    0: Synchronous Mode, 1:Asynchronous Mode,           '
    '          3240:Synchronous Mode (for Ethernet connections only)  ' 
    '                                                                 '
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Public Function TrioConnectEthernet(ByVal address As String, ByVal pmode As Short) As Boolean ' input for ip 
        Try
            m_TriCtrl.CmdProtocol = 0
            m_TriCtrl.SetHost(address)
            Dim type As Short = 2
            m_TriCtrl.Open(type, pmode)
            Dim mode As Integer = pmode
            If m_TriCtrl.IsOpen(mode) Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception

        End Try
    End Function

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Connecting TrioMotion controller via USB                         '
    '   pmode:  0: Synchronous Mode, 1:Asynchronous Mode,             '
    '           3240:Synchronous Mode (for Ethernet connections only) ' 
    '                                                                 '
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Public Function TrioConnectUsb(ByVal pmode As Short) As Boolean
        m_TriCtrl.Open(0, pmode)
        If m_TriCtrl.IsOpen(-1) Then
            Return True
        Else
            Return False

        End If
    End Function

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Connecting TrioMotion controller via PCI                         '
    '   pmode:  0: Synchronous Mode, 1:Asynchronous Mode,             '
    '           3240:Synchronous Mode (for Ethernet connections only) ' 
    '                                                                 '
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Public Function TrioConnectPci(ByVal pmode As Short) As Boolean
        m_TriCtrl.Open(3, pmode)
        If m_TriCtrl.IsOpen(-1) Then
            Return True
        Else
            Return False
        End If
    End Function

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Disconnecting TrioMotion controller                              '
    '                                                                 '
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Public Function TrioDisconnect() As Boolean
        m_TriCtrl.Close(-1)
        If m_TriCtrl.IsOpen(-1) Then
            Return False
        Else
            Return True
        End If
    End Function
#End Region

    'Set DIOs ON/OFF                                                  '
    '   ioPort:  0: controller DIOs, else: CAN DIOs                   '
    '   ioNo:    DIO number                                           '
    '   OnOff:   0: ON, 1: OFF                                        '
    Public Function SetDIOs(ByVal ioPort As Integer, ByVal ioNo As Integer, ByVal OnOff As Integer) As Boolean
        If ioPort = 0 Then 'controller DIOs
            If ioNo > gTrioUpperOutput Or ioNo < gTrioLowerOutput Then Return False
        Else 'CAN DIOs
            If ioNo > gCANUpperOutput Or ioNo < gCANLowerOutput Then Return False
        End If
        Return m_TriCtrl.Op(ioNo, OnOff)
    End Function

    'Set all DIOs OFF                                                 '
    Public Function SetAllDIOsOff() As Boolean
        Return m_TriCtrl.Op(0)
    End Function

    'Checking DIOs inputs                                             '
    ' ioPort:     0: controller DIOs, else: CAN DIOs                  '
    ' startIONo:  Starting DIO number                                 '
    ' endIONo:    Ending DIO number                                   ' 
    ' ioBitMap:   DIO status                                          '
    Public Function GetDIOs(ByVal ioPort As Integer, ByVal startIONo As Short, ByVal endIONo As Short, ByRef ioBitMap As Integer) As Boolean
        If ioPort = 0 Then 'controller DIOos
            If startIONo > gTrioUpperInput Or startIONo < gTrioLowerInput Then Return False
            If endIONo > gTrioUpperInput Or endIONo < gTrioLowerInput Then Return False
        Else 'CAN DIOs
            If startIONo > gCANUpperInput Or startIONo < gCANLowerInput Then Return False
            If endIONo > gCANUpperInput Or endIONo < gCANLowerInput Then Return False
        End If
        Return m_TriCtrl.In(startIONo, endIONo, ioBitMap)
    End Function

    'Detecting distance from top surface of PCB board by LVDT         '
    '   ioNo:       AIO's number                                      '
    '   range:      Distance (0mm to 20mm)                            '
    Public Function GetLVDTRange(ByVal ioNo As Short, ByRef range As Double) As Integer
        Dim reading As Double
        m_TriCtrl.Ain(ioNo, reading)
        If (CInt(reading) < -2046 Or CInt(reading) > 2046) Then
            range = 0
            Return -1
        End If
        range = 20.0 * (reading + 2046) / 4092.0
        Return 0
    End Function

    'Waiting for reaching demand positions on 3 axes                  '
    Public Sub WaitForEndOfMove()
        Dim dRemain(3) As Double
        Dim nAxis As Integer
        Dim bWaiting As Boolean
        Dim esp As Double = 0.01

        bWaiting = True
        While bWaiting  ' waiting for 3 axes all reaching demand postions
            For nAxis = 0 To 2
                If m_TriCtrl.Base(1, nAxis) Then  ' set base axis
                    m_TriCtrl.GetVariable("REMAIN", dRemain(nAxis))  'get remaining distance for base axis
                End If
            Next nAxis

            Sleep(25)
            Application.DoEvents()

            bWaiting = (Math.Abs(dRemain(0)) > esp) Or (Math.Abs(dRemain(1)) > esp) Or (Math.Abs(dRemain(2)) > esp)
            ' as long any axis more than 0.05 continue waiting
        End While
    End Sub

    'Waiting for reaching demand positions on 3 axes                  '
    Public Sub WaitMotionDone()
        Dim dRemain(3) As Double
        Dim nAxis As Integer
        Dim bWaiting As Boolean
        Dim esp As Double = 0.001

        bWaiting = True
        While bWaiting  ' waiting for 3 axes all reaching demand postions
            For nAxis = 0 To 2
                If m_TriCtrl.Base(1, nAxis) Then  ' set base axis
                    m_TriCtrl.GetVariable("REMAIN", dRemain(nAxis))  'get remaining distance for base axis
                End If
            Next nAxis

            Sleep(25)
            Application.DoEvents()

            bWaiting = (Math.Abs(dRemain(0)) > esp) Or (Math.Abs(dRemain(1)) > esp) Or (Math.Abs(dRemain(2)) > esp)
            ' as long any axis more than 0.05 continue waiting
        End While
    End Sub

    Public Sub WaitForNoError()
        Dim dError(3) As Double
        Dim nAxis As Integer
        Dim bWaiting As Boolean
        Dim esp As Double = 0.005

        bWaiting = True
        While bWaiting  ' waiting for 3 axes all reaching demand postions
            For nAxis = 0 To 2
                If m_TriCtrl.Base(1, nAxis) Then  ' set base axis
                    m_TriCtrl.GetVariable("FE", dError(nAxis))  'get remaining distance for base axis
                End If
            Next nAxis

            Application.DoEvents()

            bWaiting = (Math.Abs(dError(0)) > esp) Or (Math.Abs(dError(1)) > esp) Or (Math.Abs(dError(2)) > esp)
            ' as long any axis more than 0.05 continue waiting
        End While
    End Sub

    'Waiting for no moving commands in buffers on 3 axes
    Public Sub WaitNextFree()
        StopJogging()
        WaitNType()
        WaitMType()
    End Sub
    Public Sub WaitNType()
        Dim dNType(3) As Double
        Dim nAxis As Integer
        Dim bWaiting As Boolean
        Dim esp As Double = 0.1
        bWaiting = True
        While bWaiting   ' waiting for 3 axes with no more buffered moving commands
            For nAxis = 0 To 2
                If m_TriCtrl.Base(1, nAxis) Then
                    m_TriCtrl.GetVariable("NTYPE", dNType(nAxis))
                End If
            Next nAxis
            bWaiting = (Math.Abs(dNType(0)) > esp) And (Math.Abs(dNType(1)) > esp) And (Math.Abs(dNType(2)) > esp)
        End While
    End Sub
    Public Sub WaitMType()
        Dim dNType(3) As Double
        Dim nAxis As Integer
        Dim bWaiting As Boolean
        Dim esp As Double = 0.1
        bWaiting = True
        While bWaiting   ' waiting for 3 axes with no more buffered moving commands
            For nAxis = 0 To 2
                If m_TriCtrl.Base(1, nAxis) Then
                    m_TriCtrl.GetVariable("MTYPE", dNType(nAxis))
                End If
            Next nAxis
            bWaiting = (Math.Abs(dNType(0)) > esp) And (Math.Abs(dNType(1)) > esp) And (Math.Abs(dNType(2)) > esp)
        End While
    End Sub

#Region "Motion Commands"
    Public Sub SetBase(ByVal Order() As Integer)
        'The BASE command is used to direct all subsequent motion commands and axis parameter read/writes to a particular axis, or group of axes. The default setting is a sequence: 0, 1, 2, 3... 
        m_TriCtrl.Base(3, Order)
    End Sub
    Public Sub SetXYZBase()
        Dim order(2) As Integer
        order(0) = 0
        order(1) = 1
        order(2) = 2
        m_TriCtrl.Base(3, order)
    End Sub
    Public Sub SetZXYBase()
        Dim order(2) As Integer
        order(0) = 2
        order(1) = 0
        order(2) = 1
        m_TriCtrl.Base(3, order)
    End Sub
    Public Function Move_Z(ByVal Height As Double)
        WaitForEndOfMove()
        DoStepXZero()
        WaitForEndOfMove()
        If MachineHoming() Or MachineJogging() Or EStopActivated() Then
            Return False
        End If
        m_TriCtrl.MoveAbs(1, Height, 2)
        WaitForEndOfMove()
        Return True
    End Function
    Public Function MoveRelative_Z(ByVal Height As Double)
        WaitForEndOfMove()
        DoStepXZero()
        WaitForEndOfMove()
        If MachineHoming() Or MachineJogging() Or EStopActivated() Then
            Return False
        End If
        m_TriCtrl.MoveRel(1, Height, 2)
        WaitForEndOfMove()
        Return True
    End Function
    Public Function Move_XY(ByVal Position() As Double)
        WaitForEndOfMove()
        'DoStepXZero()
        'WaitForEndOfMove()
        If MachineHoming() Or MachineJogging() Or EStopActivated() Then
            Return False
        End If
        m_TriCtrl.MoveAbs(2, Position, 0)
        WaitForEndOfMove()
        Return True
    End Function
    Public Function Move_XY(ByVal x As Double, ByVal y As Double)
        WaitForEndOfMove()
        DoStepXZero()
        Dim Position(2) As Double
        Position(0) = x
        Position(1) = y
        'WaitForEndOfMove()
        If MachineHoming() Or MachineJogging() Or EStopActivated() Then
            Return False
        End If
        m_TriCtrl.MoveAbs(2, Position, 0)
        WaitForEndOfMove()
        Return True
    End Function
    Public Function MoveRelative_XY(ByVal Position() As Double)
        WaitForEndOfMove()
        DoStepXZero()
        WaitForEndOfMove()
        If MachineHoming() Or MachineJogging() Or EStopActivated() Then
            Return False
        End If
        m_TriCtrl.MoveRel(2, Position, 0)
        'WaitForEndOfMove()
        WaitMotionDone()
        Return True
    End Function
    'Only move but not wait for motion done
    Public Function MoveRelOnly(ByVal Position() As Double)
        WaitForEndOfMove()
        DoStepXZero()
        WaitForEndOfMove()
        If MachineHoming() Or MachineJogging() Or EStopActivated() Then
            Return False
        End If
        m_TriCtrl.MoveRel(2, Position, 0)
        Return True
    End Function
    Public Function MoveRelative_XYZ(ByVal Position() As Double)
        WaitForEndOfMove()
        DoStepXZero()
        WaitForEndOfMove()
        If MachineHoming() Or MachineJogging() Or EStopActivated() Then
            Return False
        End If
        m_TriCtrl.MoveRel(3, Position, 0)
        WaitForEndOfMove()
        Return True
    End Function

    Public Sub RunTrioMotionProgram(ByVal Program As String)
        m_TriCtrl.Run(Program)
    End Sub
    Public Function RunTrioMotionProgram(ByVal Program As String, ByVal Priority As Integer) As Boolean
        Dim status, attempts As Double
        Dim start_time As DateTime
        Dim stop_time As DateTime
        Dim elapsed_time As TimeSpan
        Dim rtn As Boolean
        start_time = Now

        'Do
        '    rtn = m_TriCtrl.GetProcVariable("PROC_STATUS", Priority, status)
        '    Sleep(1)
        '    stop_time = Now
        '    elapsed_time = stop_time.Subtract(start_time)
        'Loop Until rtn = 0 Or elapsed_time.TotalSeconds > 5
        'If elapsed_time.TotalSeconds > 5 Then
        '    Return False
        'End If

        Do
            status = 0
            m_TriCtrl.Run(Program, Priority)
            Sleep(25)
            rtn = m_TriCtrl.GetProcVariable("PROC_STATUS", Priority, status)
            stop_time = Now
            elapsed_time = stop_time.Subtract(start_time)
        Loop Until (rtn And status = 1) Or elapsed_time.TotalSeconds > 5
        If elapsed_time.TotalSeconds > 5 Then
            Return False
        End If
        Return True
    End Function
    Public Sub StopTrioMotionProgram(ByVal Program As String)
        m_TriCtrl.Stop(Program)
    End Sub
    Public Sub Set_XY_Speed(ByVal Speed As String)
        SetXYZBase()
        'Dim cmdStr As String
        'cmdStr = "SPEED AXIS(0)=" & Speed
        'If m_TriCtrl.Execute(cmdStr) Then
        '    Sleep(500)
        '    SetXYZBase()
        '    cmdStr = "SPEED AXIS(0)=" & Speed
        '    m_TriCtrl.Execute(cmdStr)
        'End If
        Dim sp As Double = Convert.ToDouble(Speed)
        m_TriCtrl.SetAxisVariable("SPEED", 0, sp)
        m_TriCtrl.SetAxisVariable("SPEED", 1, sp)
    End Sub
    Public Sub Set_Z_Speed(ByVal Speed As String)
        SetXYZBase()
        Dim sp As Double = Convert.ToDouble(Speed)
        m_TriCtrl.SetAxisVariable("SPEED", 2, sp)
        'Dim cmdStr As String
        'cmdStr = "SPEED AXIS(2)=" & Speed
        'm_TriCtrl.Execute(cmdStr)
    End Sub
    'This function exist bacause of the bug in trio active x
    'Everytime after mouse jogging and jogging program runned, the moverel and moveabs
    'was not working properly, command can be called and return true but the actual axis not moving
    'and if check the remain in motion perfect, there is always value, so the waitmotiondone is always
    'waiting for this value to reduce to 0 but this is nv happen
    Private Sub DoStepXZero()
        dStepVal(0) = 0
        dStepVal(1) = gStepCtrlSpeed
        dStepVal(2) = 0
        dStepVal(3) = 0
        dStepVal(4) = 0
        DoStep(dStepVal)
    End Sub
#End Region

#Region "Get Values"
    Public Function GetIDSIsHomeState()
        Return m_TriCtrl.GetTable(301, 1, HomeStateContainer)
    End Function
    Public Function IDSIsHomed() As Boolean
        If HomeStateContainer(0) = 1 Then
            Return True
        End If
        Return False
    End Function
    Public Function GetIDSState()
        Return m_TriCtrl.GetTable(0, 251, StateContainer)
    End Function
    Public Function XPosition()
        Return StateContainer(0)
    End Function
    Public Function YPosition()
        Return StateContainer(1)
    End Function
    Public Function ZPosition()
        Return StateContainer(2)
    End Function
    Public Function EStopActivated()
        If StateContainer(3) = 1 Then Return True
        Return False
    End Function
    Public Function MachineJogging()
        If StateContainer(5) = 1 Then Return True
        Return False
    End Function
    Public Function MachineRunning()
        If StateContainer(21) = 2 Then Return True
        Return False
    End Function
    Public Function MachinePaused()
        If StateContainer(21) = 1 Then Return True
        Return False
    End Function
    Public Function MachineStopped()
        If StateContainer(21) = 0 Then Return True
        Return False
    End Function
    Public Function MotionError()
        If StateContainer(22) = 1 Then Return True
        Return False
    End Function
    Public Function DispensingFinished()
        If StateContainer(23) = 1 Then Return True
        Return False
    End Function
    'set TABLE(25) to 2 to stop triggering "homing is finished label"
    Public Function HomingFinished()
        If StateContainer(25) = 1 Then
            m_TriCtrl.SetTable(25, 1, 2)
            Do
                Sleep(25)
                GetIDSState()
            Loop Until StateContainer(25) = 2
            Return True
        ElseIf StateContainer(25) = 2 Then
            Return True
        End If
    End Function
    Public Function VolCalRequested()
        If StateContainer(199) = 1 Then Return True
        Return False
    End Function
    Public Function QCRequested()
        If StateContainer(200) = 1 Then Return True
        Return False
    End Function
    Public Function MachineHoming()
        If StateContainer(25) = 0 Then Return True
        Return False
    End Function
    Public Function Stepping()
        If StateContainer(50) = 1 Then Return True
        Return False
    End Function
    Public Function Calibrating()
        If StateContainer(100) = 1 Then Return True
        Return False
    End Function
    Public Function GetCalibrationFlag()
        Return StateContainer(100)
    End Function
    Public Function ReadIO(ByVal Sensor As String, ByVal Value As Integer)
        Select Case Sensor
            Case "Dispensing Done"
                Return ReadIO(26, Value)
            Case "Board Ready"
                Return ReadIO(12, Value)
        End Select
    End Function
    Public Function ReadIO(ByVal IOnum As Integer, ByVal Value As Integer)
        m_TriCtrl.In(IOnum, IOnum, Value)
        Return Value
    End Function
#End Region

#Region "Set Values"
    Public Sub StopJogging()
        Dim values(2) As Double
        For i As Integer = 0 To 2
            values(i) = 0
        Next
        m_TriCtrl.SetTable(5, 3, values)
    End Sub
    Public Sub ResetDispensingFlag()
        Do
            m_TriCtrl.SetTable(23, 1, 0)
            Sleep(25)
            GetIDSState()
        Loop Until StateContainer(23) = 0
    End Sub
    Public Function ResetSteppingFlag(Optional ByVal timeOut As Double = 5000)
        Dim dt As DateTime = DateTime.Now
        Dim diff As Long = 0
        Do
            m_TriCtrl.SetTable(50, 1, 1)
            Sleep(25)
            GetIDSState()
            diff = (DateTime.Now.Ticks - dt.Now.Ticks) / 10000
        Loop Until StateContainer(50) = 0 Or diff > timeOut
        If diff > timeOut Then
            Return False
        End If
        Return True
    End Function
    Public Sub ResetCalibrationFlag()
        Do
            m_TriCtrl.SetTable(100, 1, 0)
            Sleep(25)
            GetIDSState()
        Loop Until StateContainer(100) = 0
    End Sub
    Public Sub ResetVolCalFlag()
        Do
            m_TriCtrl.SetTable(199, 1, 0)
            Sleep(25)
            GetIDSState()
        Loop Until StateContainer(199) = 0
    End Sub
    Public Sub ResetQCFlag()
        Do
            m_TriCtrl.SetTable(200, 1, 0)
            Sleep(25)
            GetIDSState()
        Loop Until StateContainer(200) = 0
    End Sub
    Public Sub SetQCFinished()
        Do
            m_TriCtrl.SetTable(201, 1, 1)
            Sleep(25)
            GetIDSState()
        Loop Until StateContainer(201) = 1
    End Sub
    Public Sub SetVolCalFinished()
        Do
            m_TriCtrl.SetTable(198, 1, 1)
            Sleep(25)
            GetIDSState()
        Loop Until StateContainer(198) = 1
    End Sub
    Public Sub SetVolCalStop()
        Do
            m_TriCtrl.SetTable(198, 1, 2)
            Sleep(25)
            GetIDSState()
        Loop Until StateContainer(198) = 2
    End Sub
    Public Sub SetQCStop()
        Do
            m_TriCtrl.SetTable(201, 1, 2)
            Sleep(25)
            GetIDSState()
        Loop Until StateContainer(201) = 2
    End Sub
    Public Sub SetMachineRun()
        Do
            m_TriCtrl.SetTable(21, 1, 2)
            Sleep(25)
            GetIDSState()
        Loop Until StateContainer(21) = 2
    End Sub
    Public Sub SetCalibrationFlag()
        Do
            m_TriCtrl.SetTable(100, 1, 1)
            Sleep(25)
            GetIDSState()
        Loop Until StateContainer(100) = 1
    End Sub
    Public Sub SetCalibrationType(ByVal command As String)
        Dim flag_num As Integer
        Select Case command
            Case "Needle XY Calibration"
                flag_num = 1
            Case "Needle Z Calibration"
                flag_num = 2
            Case "Volume Calibration"
                flag_num = 3
            Case "Laser XY Calibration"
                flag_num = 4
        End Select

        Do
            m_TriCtrl.SetTable(109, 1, flag_num)
            Sleep(25)
            GetIDSState()
        Loop Until StateContainer(109) = flag_num
    End Sub
    Public Sub SetMachinePause()
        Do
            m_TriCtrl.SetTable(21, 1, 1)
            Sleep(25)
            GetIDSState()
        Loop Until StateContainer(21) = 1
    End Sub
    Public Sub SetMachineStop()
        Do
            m_TriCtrl.SetTable(21, 1, 0)
            Sleep(25)
            GetIDSState()
        Loop Until StateContainer(21) = 0
    End Sub
    Public Sub SetMachineRunMode(ByVal runmode As Integer)
        Do
            m_TriCtrl.SetTable(20, 1, runmode)
            Sleep(25)
            GetIDSState()
        Loop Until StateContainer(20) = runmode
    End Sub
    Public Sub SetTrioMotionValues(ByVal Command As String)
        Dim table_index, value As Integer
        Select Case Command
            Case "Lock X"
                table_index = 10
                value = 1
            Case "Lock Y"
                table_index = 11
                value = 1
            Case "Lock Z"
                table_index = 12
                value = 1
            Case "Unlock X"
                table_index = 10
                value = 0
            Case "Unlock Y"
                table_index = 11
                value = 0
            Case "Unlock Z"
                table_index = 12
                value = 0
        End Select
        Do
            m_TriCtrl.SetTable(table_index, 1, value)
            Sleep(25)
            GetIDSState()
        Loop Until StateContainer(table_index) = value
    End Sub
    Public Sub SetTrioMotionValues(ByVal Command As String, ByVal values As Object)
        Select Case Command
            Case "Jogging"
                m_TriCtrl.SetTable(5, 3, values)
        End Select
    End Sub
    Public Sub DoStep(ByVal values() As Single)
        m_TriCtrl.SetTable(50, 5, values)
        ResetSteppingFlag()
    End Sub
    Public Function TurnOn(ByVal Sensor As String)
        Select Case Sensor
            Case "Clean" 'suckback
                Return WriteToIO(19, 1)
            Case "Left Needle IO"
                Return WriteToIO(25, 1)
            Case "Material Air"
                Return WriteToIO(27, 1)
            Case "Left Needle Slide"
                Return WriteToIO(18, 1)
            Case "Right Needle Slide"
                Return WriteToIO(17, 1)
        End Select
    End Function

    Public Function TurnOff(ByVal Sensor As String)
        Select Case Sensor
            Case "Clean" 'suckback
                Return WriteToIO(19, 0)
            Case "Left Needle IO"
                Return WriteToIO(25, 0)
            Case "Material Air"
                Return WriteToIO(27, 0)
            Case "Left Needle Slide"
                Return WriteToIO(18, 0)
            Case "Right Needle Slide"
                Return WriteToIO(17, 0)
        End Select
    End Function
    Public Function WriteToIO(ByVal IOnum As Integer, ByVal OnOff As Integer)
        Return m_TriCtrl.Op(IOnum, OnOff)
    End Function
    Public Sub TrioStop()
        m_TriCtrl.Stop("SETDATUM")
        If m_TriCtrl.Stop("CALIBRATIONS") Then
            ResetCalibrationFlag()
        End If
        m_TriCtrl.Stop("DISPENSE")
        m_TriCtrl.RapidStop()
        m_TriCtrl.Cancel(1, 0)
        m_TriCtrl.Cancel(1, 1)
        m_TriCtrl.Cancel(1, 2)
        m_TriCtrl.Cancel(0, 0)
        m_TriCtrl.Cancel(0, 1)
        m_TriCtrl.Cancel(0, 2)
        m_TriCtrl.Op(27, 0)
        m_TriCtrl.Op(25, 0)
    End Sub
#End Region

    Public Sub Disconnect_Controller()

        If m_TriCtrl.IsOpen(0) Then
            TrioStop()
            StopTrioMotionProgram("LOGPOS")
            StopTrioMotionProgram("JOGGING")
            RunTrioMotionProgram("EXITIDS")
            Array.Clear(StateContainer, 0, StateContainer.Length)
            m_TriCtrl.Close(0)
        End If

    End Sub
    'yy
    Public Sub Disconnect()

        If m_TriCtrl.IsOpen(0) Then
            m_TriCtrl.Close(0)
        End If

    End Sub

    Public Function Connect()
        Return TrioConnectEthernet("192.168.0.250", 3240)
    End Function

    Public Function Connect_Controller()

        'StopTrioMotionProgram("LOGPOS")
        'StopTrioMotionProgram("JOGGING")
        'StopTrioMotionProgram("DISPENSE")
        'StopTrioMotionProgram("SETDATUM")
        'StopTrioMotionProgram("CALIBRATIONS")
        If TrioConnectEthernet("192.168.0.250", 3240) Then
            RunTrioMotionProgram("STARTUP")
            Sleep(500)
            RunTrioMotionProgram("JOGGING", 5)
            RunTrioMotionProgram("LOGPOS", 4)
            Sleep(500)
        Else
            MsgBox("Can't connect to controller")
            Return False
        End If

        Return True

    End Function

#Region "GUI for stepper"
    Dim dStepVal(5) As Single
    Dim gStepCtrlSpeed As Double = 100.0

    Private Sub BTStepXminus_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTStepXminus.Click

        If Get_Limit("X-") Or MachineRunning() Or MachineHoming() Or Calibrating() Then
            Exit Sub
        End If

        dStepVal(0) = 0
        dStepVal(1) = gStepCtrlSpeed
        dStepVal(2) = -ComboBoxFineStep.Value
        dStepVal(3) = 0
        dStepVal(4) = 0
        DoStep(dStepVal)

    End Sub

    Private Sub BTStepXplus_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTStepXplus.Click

        If Get_Limit("X+") Or MachineRunning() Or MachineHoming() Or Calibrating() Then
            Exit Sub
        End If

        dStepVal(0) = 0
        dStepVal(1) = gStepCtrlSpeed
        dStepVal(2) = ComboBoxFineStep.Value
        dStepVal(3) = 0
        dStepVal(4) = 0
        DoStep(dStepVal)

    End Sub

    Private Sub BTStepYminus_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTStepYminus.Click

        If Get_Limit("Y-") Or MachineRunning() Or MachineHoming() Or Calibrating() Then
            Exit Sub
        End If

        dStepVal(0) = 0
        dStepVal(1) = gStepCtrlSpeed
        dStepVal(2) = 0
        dStepVal(3) = -ComboBoxFineStep.Value
        dStepVal(4) = 0
        DoStep(dStepVal)

    End Sub

    Private Sub BTStepYplus_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTStepYplus.Click

        If Get_Limit("Y+") Or MachineRunning() Or MachineHoming() Or Calibrating() Then
            Exit Sub
        End If

        dStepVal(0) = 0
        dStepVal(1) = gStepCtrlSpeed
        dStepVal(2) = 0
        dStepVal(3) = ComboBoxFineStep.Value
        dStepVal(4) = 0
        DoStep(dStepVal)


    End Sub

    Private Sub BTStepZup_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTStepZup.Click

        If Get_Limit("Up") Or MachineRunning() Or MachineHoming() Or Calibrating() Then
            Exit Sub
        End If

        dStepVal(0) = 0
        dStepVal(1) = gStepCtrlSpeed
        dStepVal(2) = 0
        dStepVal(3) = 0
        dStepVal(4) = ComboBoxFineStep.Value
        DoStep(dStepVal)

    End Sub

    Private Sub BTStepZdown_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTStepZdown.Click

        If Get_Limit("Dn") Or MachineRunning() Or MachineHoming() Or Calibrating() Then
            Exit Sub
        End If

        dStepVal(0) = 0
        dStepVal(1) = gStepCtrlSpeed
        dStepVal(2) = 0
        dStepVal(3) = 0
        dStepVal(4) = -ComboBoxFineStep.Value
        DoStep(dStepVal)

    End Sub

    Private Function Get_Limit(ByVal Direction As String) As Boolean

        Dim Read_Value As Integer = 0

        m_TriCtrl.In(0, 5, Read_Value)

        Select Case Direction
            Case "X+"
                Read_Value = Read_Value And &H2
            Case "X-"
                Read_Value = Read_Value And &H1
            Case "Y+"
                Read_Value = Read_Value And &H4
            Case "Y-"
                Read_Value = Read_Value And &H8
            Case "Up"
                Read_Value = Read_Value And &H10
            Case "Dn"
                Read_Value = Read_Value And &H20
            Case Else
                Return False
        End Select

        If (Read_Value = 0) Then
            MsgBox("Machine gets the limit position!", MsgBoxStyle.OKOnly, "Limit Error")
            Return True
        Else
            Return False
        End If

    End Function


#End Region

End Class
