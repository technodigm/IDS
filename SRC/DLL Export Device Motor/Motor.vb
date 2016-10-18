
Imports System.Threading

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
    Friend WithEvents ComboBoxFineStep As System.Windows.Forms.NumericUpDown
    Public WithEvents SteppingButtons As System.Windows.Forms.Panel
    Friend WithEvents rbtSteps As System.Windows.Forms.RadioButton
    Friend WithEvents rbt100micron As System.Windows.Forms.RadioButton
    Friend WithEvents rbt10micron As System.Windows.Forms.RadioButton
    Friend WithEvents rbt1micron As System.Windows.Forms.RadioButton
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents rtb20micron As System.Windows.Forms.RadioButton
    Friend WithEvents rbt50micron As System.Windows.Forms.RadioButton
    Friend WithEvents rbt2micron As System.Windows.Forms.RadioButton
    Friend WithEvents rbt5micron As System.Windows.Forms.RadioButton
    Friend WithEvents rbt200micron As System.Windows.Forms.RadioButton
    Friend WithEvents rbt500 As System.Windows.Forms.RadioButton
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(CIDSTrioController))
        Me.SteppingButtons = New System.Windows.Forms.Panel
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.rtb20micron = New System.Windows.Forms.RadioButton
        Me.rbt10micron = New System.Windows.Forms.RadioButton
        Me.ComboBoxFineStep = New System.Windows.Forms.NumericUpDown
        Me.rbtSteps = New System.Windows.Forms.RadioButton
        Me.rbt500 = New System.Windows.Forms.RadioButton
        Me.rbt200micron = New System.Windows.Forms.RadioButton
        Me.rbt100micron = New System.Windows.Forms.RadioButton
        Me.rbt50micron = New System.Windows.Forms.RadioButton
        Me.rbt5micron = New System.Windows.Forms.RadioButton
        Me.rbt2micron = New System.Windows.Forms.RadioButton
        Me.rbt1micron = New System.Windows.Forms.RadioButton
        Me.BTStepZdown = New System.Windows.Forms.Button
        Me.BTStepZup = New System.Windows.Forms.Button
        Me.BTStepXminus = New System.Windows.Forms.Button
        Me.BTStepXplus = New System.Windows.Forms.Button
        Me.BTStepYminus = New System.Windows.Forms.Button
        Me.BTStepYplus = New System.Windows.Forms.Button
        Me.SteppingButtons.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        CType(Me.ComboBoxFineStep, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'SteppingButtons
        '
        Me.SteppingButtons.BackColor = System.Drawing.SystemColors.Control
        Me.SteppingButtons.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.SteppingButtons.Controls.Add(Me.GroupBox1)
        Me.SteppingButtons.Controls.Add(Me.BTStepZdown)
        Me.SteppingButtons.Controls.Add(Me.BTStepZup)
        Me.SteppingButtons.Controls.Add(Me.BTStepXminus)
        Me.SteppingButtons.Controls.Add(Me.BTStepXplus)
        Me.SteppingButtons.Controls.Add(Me.BTStepYminus)
        Me.SteppingButtons.Controls.Add(Me.BTStepYplus)
        Me.SteppingButtons.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.SteppingButtons.Location = New System.Drawing.Point(20, 23)
        Me.SteppingButtons.Name = "SteppingButtons"
        Me.SteppingButtons.Size = New System.Drawing.Size(332, 280)
        Me.SteppingButtons.TabIndex = 23
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.rtb20micron)
        Me.GroupBox1.Controls.Add(Me.rbt10micron)
        Me.GroupBox1.Controls.Add(Me.ComboBoxFineStep)
        Me.GroupBox1.Controls.Add(Me.rbtSteps)
        Me.GroupBox1.Controls.Add(Me.rbt500)
        Me.GroupBox1.Controls.Add(Me.rbt200micron)
        Me.GroupBox1.Controls.Add(Me.rbt100micron)
        Me.GroupBox1.Controls.Add(Me.rbt50micron)
        Me.GroupBox1.Controls.Add(Me.rbt5micron)
        Me.GroupBox1.Controls.Add(Me.rbt2micron)
        Me.GroupBox1.Controls.Add(Me.rbt1micron)
        Me.GroupBox1.Location = New System.Drawing.Point(8, 184)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(320, 88)
        Me.GroupBox1.TabIndex = 14
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "mm"
        '
        'rtb20micron
        '
        Me.rtb20micron.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.rtb20micron.Location = New System.Drawing.Point(256, 56)
        Me.rtb20micron.Name = "rtb20micron"
        Me.rtb20micron.Size = New System.Drawing.Size(59, 24)
        Me.rtb20micron.TabIndex = 14
        Me.rtb20micron.Text = "0.02"
        '
        'rbt10micron
        '
        Me.rbt10micron.Location = New System.Drawing.Point(200, 56)
        Me.rbt10micron.Name = "rbt10micron"
        Me.rbt10micron.Size = New System.Drawing.Size(64, 24)
        Me.rbt10micron.TabIndex = 12
        Me.rbt10micron.Text = "0.01"
        '
        'ComboBoxFineStep
        '
        Me.ComboBoxFineStep.DecimalPlaces = 3
        Me.ComboBoxFineStep.Enabled = False
        Me.ComboBoxFineStep.Location = New System.Drawing.Point(224, 24)
        Me.ComboBoxFineStep.Maximum = New Decimal(New Integer() {200, 0, 0, 0})
        Me.ComboBoxFineStep.Name = "ComboBoxFineStep"
        Me.ComboBoxFineStep.Size = New System.Drawing.Size(88, 27)
        Me.ComboBoxFineStep.TabIndex = 9
        Me.ComboBoxFineStep.Value = New Decimal(New Integer() {1, 0, 0, 0})
        '
        'rbtSteps
        '
        Me.rbtSteps.Location = New System.Drawing.Point(208, 24)
        Me.rbtSteps.Name = "rbtSteps"
        Me.rbtSteps.Size = New System.Drawing.Size(16, 24)
        Me.rbtSteps.TabIndex = 10
        '
        'rbt500
        '
        Me.rbt500.Location = New System.Drawing.Point(160, 24)
        Me.rbt500.Name = "rbt500"
        Me.rbt500.Size = New System.Drawing.Size(56, 24)
        Me.rbt500.TabIndex = 19
        Me.rbt500.Text = "0.5"
        '
        'rbt200micron
        '
        Me.rbt200micron.Location = New System.Drawing.Point(112, 24)
        Me.rbt200micron.Name = "rbt200micron"
        Me.rbt200micron.Size = New System.Drawing.Size(56, 24)
        Me.rbt200micron.TabIndex = 18
        Me.rbt200micron.Text = "0.2"
        '
        'rbt100micron
        '
        Me.rbt100micron.Location = New System.Drawing.Point(64, 24)
        Me.rbt100micron.Name = "rbt100micron"
        Me.rbt100micron.Size = New System.Drawing.Size(56, 24)
        Me.rbt100micron.TabIndex = 11
        Me.rbt100micron.Text = "0.1"
        '
        'rbt50micron
        '
        Me.rbt50micron.Checked = True
        Me.rbt50micron.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.rbt50micron.Location = New System.Drawing.Point(8, 24)
        Me.rbt50micron.Name = "rbt50micron"
        Me.rbt50micron.Size = New System.Drawing.Size(64, 24)
        Me.rbt50micron.TabIndex = 15
        Me.rbt50micron.TabStop = True
        Me.rbt50micron.Text = "0.05"
        '
        'rbt5micron
        '
        Me.rbt5micron.Location = New System.Drawing.Point(136, 56)
        Me.rbt5micron.Name = "rbt5micron"
        Me.rbt5micron.Size = New System.Drawing.Size(72, 24)
        Me.rbt5micron.TabIndex = 17
        Me.rbt5micron.Text = "0.005"
        '
        'rbt2micron
        '
        Me.rbt2micron.Location = New System.Drawing.Point(72, 56)
        Me.rbt2micron.Name = "rbt2micron"
        Me.rbt2micron.Size = New System.Drawing.Size(72, 24)
        Me.rbt2micron.TabIndex = 16
        Me.rbt2micron.Text = "0.002"
        '
        'rbt1micron
        '
        Me.rbt1micron.Location = New System.Drawing.Point(8, 56)
        Me.rbt1micron.Name = "rbt1micron"
        Me.rbt1micron.Size = New System.Drawing.Size(72, 24)
        Me.rbt1micron.TabIndex = 13
        Me.rbt1micron.Text = "0.001"
        '
        'BTStepZdown
        '
        Me.BTStepZdown.BackColor = System.Drawing.Color.Linen
        Me.BTStepZdown.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.5!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.BTStepZdown.Image = CType(resources.GetObject("BTStepZdown.Image"), System.Drawing.Image)
        Me.BTStepZdown.ImageAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.BTStepZdown.Location = New System.Drawing.Point(264, 104)
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
        Me.BTStepZup.Location = New System.Drawing.Point(264, 16)
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
        Me.BTStepXminus.Location = New System.Drawing.Point(16, 72)
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
        Me.BTStepXplus.Location = New System.Drawing.Point(160, 72)
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
        Me.BTStepYminus.Location = New System.Drawing.Point(104, 104)
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
        Me.BTStepYplus.Location = New System.Drawing.Point(104, 8)
        Me.BTStepYplus.Name = "BTStepYplus"
        Me.BTStepYplus.Size = New System.Drawing.Size(48, 80)
        Me.BTStepYplus.TabIndex = 3
        Me.BTStepYplus.Text = "Y+"
        Me.BTStepYplus.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'CIDSTrioController
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(376, 326)
        Me.Controls.Add(Me.SteppingButtons)
        Me.Name = "CIDSTrioController"
        Me.Text = "Axes Relative Move"
        Me.SteppingButtons.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
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

    Public Shared m_TriCtrl As New TrioPCLib.TrioPCClass
    Public WithEvents robot_Asyn As New TrioPCLib.TrioPCClass
    Public StateContainer(251) As Double
    Public HomeStateContainer(0) As Double
    Private startTime As Long = DateTime.Now.Ticks
    Private Shared motionDoneEvent(1) As ManualResetEvent
    Private Shared checkRemainEvent(2) As ManualResetEvent
    Private stepSize As Single

#Region " Connection "

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Connecting TrioMotion controller via ethernet                    '
    '   address:  TrioMotion controller's IP address                  '
    '   pmode:    0: Synchronous Mode, 1:Asynchronous Mode,           '
    '          3240:Synchronous Mode (for Ethernet connections only)  ' 
    '                                                                 '
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim asynConnect As Boolean = False
    Public Function TrioConnectEthernet(ByVal address As String, ByVal pmode As Short) As Boolean ' input for ip 
        Try
            

            m_TriCtrl.CmdProtocol = 0
            m_TriCtrl.SetHost(address)
            Dim type As Short = 2
            m_TriCtrl.Open(type, pmode)
            Dim mode As Integer = pmode
            If m_TriCtrl.IsOpen(mode) Then
                'robot_Asyn.CmdProtocol = 0
                'robot_Asyn.SetHost(address)
                'robot_Asyn.Open(2, 23)
                'If robot_Asyn.IsOpen(23) Then
                '    asynConnect = True
                '    Console.WriteLine("Asyn Port 23 opened")
                'End If
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
    'Public Sub WaitForEndOfMove()
    '    Dim dRemain(3) As Double
    '    Dim nAxis As Integer
    '    Dim bWaiting As Boolean
    '    'Dim esp As Double = 0.01
    '    Dim esp As Double = 0.0

    '    bWaiting = True
    '    While bWaiting  ' waiting for 3 axes all reaching demand postions
    '        For nAxis = 0 To 2
    '            If m_TriCtrl.Base(1, nAxis) Then  ' set base axis
    '                m_TriCtrl.GetVariable("REMAIN", dRemain(nAxis))  'get remaining distance for base axis
    '            End If
    '        Next nAxis

    '        Sleep(25)
    '        Application.DoEvents()

    '        bWaiting = (Math.Abs(dRemain(0)) > esp) Or (Math.Abs(dRemain(1)) > esp) Or (Math.Abs(dRemain(2)) > esp)
    '        ' as long any axis more than 0.05 continue waiting
    '    End While
    'End Sub

    'Waiting for reaching demand positions on 3 axes   
    Shared waitMotionTimeout As Integer = 60000
    Shared thd As Thread
    'Pool inside thread
    Dim startWaitTime As Long = 0
    Shared bWaitingWMD As Boolean = False
    Shared bWaitingCheckRemain As Boolean = False
    Public Function WaitMotionDone() As Boolean

        bWaitingWMD = True
        'If Not (thd Is Nothing) Then
        'thd.Abort()
        'breakCheckRemain = True
        'While breakCheckRemain
        '    Application.DoEvents()
        'End While
        'Sleep(5)
        'End If
        'motionDoneEvent(0).Reset()
        'motionDoneEvent(1).Reset()
        If isCheckingRemain Then
            checkRemainEvent(1).Set()
            While isCheckingRemain
                Application.DoEvents()
                Sleep(5)
            End While
        End If
        If thd Is Nothing Then
            thd = New Thread(AddressOf CheckAxesRemain)
            'Console.WriteLine("#Start motion done check")
            Sleep(100)
            thd.Start()
            checkRemainEvent(2).Set()
        Else
            checkRemainEvent(2).Set()
            'Sleep(50)
        End If

        startWaitTime = DateTime.Now.Ticks
        'Console.WriteLine("#Start Motion done wait")
        While True
            Dim rtn As Integer = WaitHandle.WaitAny(motionDoneEvent, 100, False)
            'motionDoneEvent(0).Reset()
            'motionDoneEvent(1).Reset()
            If rtn = 0 Then
                motionDoneEvent(0).Reset()
                motionDoneEvent(1).Reset()
                Try
                    If thd.IsAlive Then
                        checkRemainEvent(1).Reset()
                        'thd.Abort()
                        'While thd.IsAlive
                        '    Application.DoEvents()
                        'End While
                    End If
                    Console.WriteLine("#2TH motion done event received")
                    bWaitingWMD = False
                    Return True
                Catch ex As Exception
                End Try
                Return True
            ElseIf rtn = 1 Then
                Try
                    'motionDoneEvent(0).Reset()
                    motionDoneEvent(1).Reset()
                    AbortStatus = False
                    If thd.IsAlive Then
                        checkRemainEvent(1).Reset()
                        'thd.Abort()
                        'While thd.IsAlive
                        '    Application.DoEvents()
                        'End While
                    End If
                    bWaitingWMD = False
                    Console.WriteLine("#2TH Wait motion done cancel")
                Catch ex As Exception
                End Try
                Return False
            ElseIf rtn = WaitHandle.WaitTimeout Then
                Try
                    If ((DateTime.Now.Ticks - startWaitTime) / 10000) > waitMotionTimeout Then
                        'thd.Abort()
                        checkRemainEvent(2).Reset()
                        bWaitingWMD = False
                        Return False
                        'Else keep waiting
                    End If
                    Application.DoEvents()
                    'Console.WriteLine("#1TH Abort Thd")
                    'Return False
                Catch ex As Exception
                End Try
                'Return False
            Else
                MessageBox.Show("Unknow event in thread")
            End If
        End While

    End Function
    Public Shared AbortStatus As Boolean = False
    Public Shared Sub AbortMotionDone()
        If Not (thd Is Nothing) Then
            AbortStatus = True
            If bWaitingWMD Then
                motionDoneEvent(1).Set()
            End If

            'thd.Abort()
            If isCheckingRemain Then
                checkRemainEvent(1).Set()
            End If
            'Console.WriteLine("Abort motion done thread")
        End If
        'If Not (motionDoneEvent(1) Is Nothing) Then
        '    Console.WriteLine("Trigger Abort motion done event")
        '    AbortStatus = True
        '    motionDoneEvent(1).Set()
        'End If
    End Sub

    Public Shared breakCheckRemain As Boolean = False
    Public Shared isCheckingRemain As Boolean = False
    Private Shared Sub CheckAxesRemain()
        Dim wait As Boolean
        Dim nAxis As Integer
        Dim dRemain(3) As Double
        wait = True

        While True
            Dim rtn As Integer = WaitHandle.WaitAny(checkRemainEvent)
            'motionDoneEvent(0).Reset()
            If rtn = 0 Then ' Abort this thread
                checkRemainEvent(0).Reset()
                checkRemainEvent(1).Reset()
                checkRemainEvent(2).Reset()
                Console.WriteLine("#Check Remain exit")
                Exit Sub
            ElseIf rtn = 1 Then 'Reset
                checkRemainEvent(0).Reset()
                checkRemainEvent(1).Reset()
                checkRemainEvent(2).Reset()
                isCheckingRemain = False
                Console.WriteLine("#Check Remain reset")
            ElseIf rtn = 2 Then 'Check remain
                isCheckingRemain = True
                For nAxis = 0 To 2
                    If m_TriCtrl.Base(1, nAxis) Then  ' set base axis
                        m_TriCtrl.GetVariable("REMAIN", dRemain(nAxis))  'get remaining distance for base axis
                    End If
                    If breakCheckRemain Then
                        breakCheckRemain = False
                        Exit Sub
                    End If
                Next nAxis
                Sleep(25)
                If (Math.Abs(dRemain(0)) = 0) And (Math.Abs(dRemain(1)) = 0) And (Math.Abs(dRemain(2)) = 0) Then
                    checkRemainEvent(0).Reset()
                    checkRemainEvent(1).Reset()
                    checkRemainEvent(2).Reset()
                    isCheckingRemain = False
                    motionDoneEvent(0).Set()
                    Console.WriteLine("#MotionDone Event Set")
                End If
            Else
                MessageBox.Show("Unknow event in check remain")
            End If
        End While
        Console.WriteLine("#Check Remain exit")
        'Dim enterTime As Long = DateTime.Now.Ticks
        ''Console.WriteLine("#TH Check remain start")
        'Do
        '    If breakCheckRemain Then
        '        breakCheckRemain = False
        '        Exit Sub
        '    End If
        '    For nAxis = 0 To 2
        '        If m_TriCtrl.Base(1, nAxis) Then  ' set base axis
        '            m_TriCtrl.GetVariable("REMAIN", dRemain(nAxis))  'get remaining distance for base axis
        '        End If
        '        If breakCheckRemain Then
        '            breakCheckRemain = False
        '            Exit Sub
        '        End If
        '    Next nAxis
        '    Sleep(25)
        '    wait = (Math.Abs(dRemain(0)) > 0) Or (Math.Abs(dRemain(1)) > 0) Or (Math.Abs(dRemain(2)) > 0)
        'Loop Until Not (wait) Or (DateTime.Now.Ticks - enterTime) / 1000 > (waitMotionTimeout)
        'If Not wait Then
        '    motionDoneEvent(0).Set()
        '    Console.WriteLine("#MotionDone Event Set")
        'Else
        '    'Console.WriteLine("#TH Time out")
        'End If
        ' Console.WriteLine("#TH Check remain Ended")
    End Sub


    Public Sub WaitForNoError()
        Dim dError(3) As Double
        Dim nAxis As Integer
        Dim bWaiting As Boolean
        Dim esp As Double = 0.005

        bWaiting = True
        While bWaiting   ' waiting for 3 axes all reaching demand postions
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
        'WaitMotionDone()
        ' DoStepXZero()
        Console.WriteLine("Wait MD before Move Z")
        If Not WaitMotionDone() Then Return False
        Console.WriteLine("After Wait MD before Move Z")
        If MachineHoming() Or MachineJogging() Or EStopActivated() Then
            Return False
        End If
        Console.WriteLine("Move Z")
        m_TriCtrl.MoveAbs(1, Height, 2)
        Console.WriteLine("After Move Z")
        If Not WaitMotionDone() Then Return False
        Return True
    End Function
    Public Function MoveRelative_Z(ByVal Height As Double)
        WaitMotionDone()
        DoStepXZero()
        WaitMotionDone()
        If MachineHoming() Or MachineJogging() Or EStopActivated() Then
            Return False
        End If
        m_TriCtrl.MoveRel(1, Height, 2)
        WaitMotionDone()
        Return True
    End Function
    Public Function Move_XY(ByVal Position() As Double)
        ' WaitMotionDone()
        'DoStepXZero()
        'WaitForEndOfMove()
        If MachineHoming() Or MachineJogging() Or EStopActivated() Then
            Return False
        End If
        m_TriCtrl.MoveAbs(2, Position, 0)
        WaitMotionDone()
        Return True
    End Function
    Public Function Move_XY(ByVal x As Double, ByVal y As Double)
        WaitMotionDone()
        DoStepXZero()
        Dim Position(2) As Double
        Position(0) = x
        Position(1) = y
        WaitMotionDone()
        If MachineHoming() Or MachineJogging() Or EStopActivated() Then
            Return False
        End If
        m_TriCtrl.MoveAbs(2, Position, 0)
        WaitMotionDone()
        Return True
    End Function
    Public Function MoveRelative_XY(ByVal Position() As Double)
        WaitMotionDone()
        DoStepXZero()
        WaitMotionDone()
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
        WaitMotionDone()
        DoStepXZero()
        WaitMotionDone()
        If MachineHoming() Or MachineJogging() Or EStopActivated() Then
            Return False
        End If
        m_TriCtrl.MoveRel(2, Position, 0)
        WaitMotionDone()
        Return True
    End Function
    Public Function MoveRelative_XYZ(ByVal Position() As Double)
        WaitMotionDone()
        DoStepXZero()
        WaitMotionDone()
        If MachineHoming() Or MachineJogging() Or EStopActivated() Then
            Return False
        End If
        m_TriCtrl.MoveRel(3, Position, 0)
        WaitMotionDone()
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
        Return m_TriCtrl.GetTable(0, 252, StateContainer)
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
            'm_TriCtrl.SetTable(25, 1, 2)
            Do
                If m_TriCtrl.SetTable(25, 1, 2) Then Exit Function
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
        'GetIDSState()
        If StateContainer(100) = 1 Then Return True
        Return False
    End Function
    Public Function GetCalibrationFlag()
        Return StateContainer(100)
    End Function
    Public Function GetVCTable100(ByRef table100Value As Object) As Boolean
        Return m_TriCtrl.GetTable(100, 1, table100Value)
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
        'Retry 3 time
        If m_TriCtrl.SetTable(5, 3, values) Then Exit Sub
        If m_TriCtrl.SetTable(5, 3, values) Then Exit Sub
        If m_TriCtrl.SetTable(5, 3, values) Then Exit Sub
    End Sub
    Private Function IfTimeout() As Boolean
        If ((DateTime.Now.Ticks - startTime) / 1000) > 5000 Then Return True
        Return False
    End Function

    Public Sub ResetDispensingFlag()
        startTime = DateTime.Now.Ticks
        Do
            If m_TriCtrl.SetTable(23, 1, 0) Then Exit Sub
            Sleep(25)
            GetIDSState()
        Loop Until (StateContainer(23) = 0 Or IfTimeout())
    End Sub
    Public Function ResetSteppingFlag(Optional ByVal timeOut As Double = 5000)
        Dim dt As DateTime = DateTime.Now
        Dim diff As Long = 0
        Do
            If m_TriCtrl.SetTable(50, 1, 1) Then Exit Function
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
        startTime = DateTime.Now.Ticks
        Do
            If m_TriCtrl.SetTable(100, 1, 0) Then
                'StateContainer(100) = 0
                Exit Sub
            End If
            Sleep(25)
            GetIDSState()
        Loop Until StateContainer(100) = 0 Or IfTimeout()
    End Sub
    Public Sub ResetVolCalFlag()
        startTime = DateTime.Now.Ticks
        Do
            If m_TriCtrl.SetTable(199, 1, 0) Then Exit Sub
            Sleep(25)
            GetIDSState()
        Loop Until StateContainer(199) = 0 Or IfTimeout()
    End Sub
    Public Sub ResetQCFlag()
        startTime = DateTime.Now.Ticks
        Do
            If m_TriCtrl.SetTable(200, 1, 0) Then Exit Sub
            Sleep(25)
            GetIDSState()
        Loop Until StateContainer(200) = 0 Or IfTimeout()
    End Sub
    Public Sub SetQCFinished()
        startTime = DateTime.Now.Ticks
        Do
            If m_TriCtrl.SetTable(201, 1, 1) Then Exit Sub
            Sleep(25)
            GetIDSState()
        Loop Until StateContainer(201) = 1 Or IfTimeout()
    End Sub
    Public Sub SetVolCalFinished()
        startTime = DateTime.Now.Ticks
        Do
            If m_TriCtrl.SetTable(198, 1, 1) Then Exit Sub
            Sleep(25)
            GetIDSState()
        Loop Until StateContainer(198) = 1 Or IfTimeout()
    End Sub
    Public Sub SetVolCalStop()
        startTime = DateTime.Now.Ticks
        Do
            If m_TriCtrl.SetTable(198, 1, 2) Then Exit Sub
            Sleep(25)
            GetIDSState()
        Loop Until StateContainer(198) = 2 Or IfTimeout()
    End Sub
    Public Sub SetQCStop()
        startTime = DateTime.Now.Ticks
        Do
            If m_TriCtrl.SetTable(201, 1, 2) Then Exit Sub
            Sleep(25)
            GetIDSState()
        Loop Until StateContainer(201) = 2 Or IfTimeout()
    End Sub
    'To make sure dispensing process is stopped
    Public Sub AbortDispensing()
        startTime = DateTime.Now.Ticks
        Do
            If m_TriCtrl.SetTable(251, 1, 1) Then Exit Sub
            Sleep(25)
            GetIDSState()
        Loop Until StateContainer(251) = 1 Or IfTimeout()
    End Sub
    Public Function WaitAbortDispensing() As Boolean
        startTime = DateTime.Now.Ticks
        Do
            Sleep(25)
            GetIDSState()
            If IfTimeout() Then Return False
        Loop Until StateContainer(251) = 0
        Return True
    End Function
    'Reset the flag if start a new dispensing process
    Public Sub ResetAbortDispensingFlag()
        startTime = DateTime.Now.Ticks
        Do
            If m_TriCtrl.SetTable(251, 1, 0) Then Exit Sub
            Sleep(25)
            GetIDSState()
        Loop Until StateContainer(251) = 0 Or IfTimeout()
    End Sub
    Public Sub SetMachineRun()
        startTime = DateTime.Now.Ticks
        Do
            If m_TriCtrl.SetTable(21, 1, 2) Then Exit Sub
            Sleep(25)
            GetIDSState()
        Loop Until StateContainer(21) = 2 Or IfTimeout()
    End Sub
    Public Function SetCalibrationFlag() As Boolean
        startTime = DateTime.Now.Ticks
        Do
            If m_TriCtrl.SetTable(100, 1, 1) Then
                ' GetIDSState()
                StateContainer(100) = 1
                'Exit Sub
            End If
            Sleep(25)
            GetIDSState()
        Loop Until StateContainer(100) = 1 Or IfTimeout()
        If IfTimeout() Then Return False
        Return True
    End Function

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
        startTime = DateTime.Now.Ticks
        Do
            If m_TriCtrl.SetTable(109, 1, flag_num) Then Exit Sub
            Sleep(25)
            GetIDSState()
        Loop Until StateContainer(109) = flag_num Or IfTimeout()
    End Sub
    Public Sub SetMachinePause()
        startTime = DateTime.Now.Ticks
        Do
            If m_TriCtrl.SetTable(21, 1, 1) Then Exit Sub
            Sleep(25)
            GetIDSState()
        Loop Until StateContainer(21) = 1 Or IfTimeout()
    End Sub
    Public Sub SetMachineStop()
        startTime = DateTime.Now.Ticks
        Do
            If m_TriCtrl.SetTable(21, 1, 0) Then Exit Sub
            Sleep(25)
            GetIDSState()
        Loop Until StateContainer(21) = 0 Or IfTimeout()
    End Sub
    Public Sub SetMachineRunMode(ByVal runmode As Integer)
        startTime = DateTime.Now.Ticks
        Do
            If m_TriCtrl.SetTable(20, 1, runmode) Then Exit Sub
            Sleep(25)
            GetIDSState()
        Loop Until StateContainer(20) = runmode Or IfTimeout()
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
        startTime = DateTime.Now.Ticks
        Do
            If m_TriCtrl.SetTable(table_index, 1, value) Then Exit Sub
            Sleep(25)
            GetIDSState()
        Loop Until StateContainer(table_index) = value Or IfTimeout()
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
    Public Function TrioStop() As Boolean
        Dim timeOut As Long = 5000
        If Not (m_TriCtrl.Stop("DISPENSE")) Then
            If Not (m_TriCtrl.Stop("DISPENSE")) Then
                If Not (m_TriCtrl.Stop("DISPENSE")) Then
                    MessageBox.Show("Failed to stop dispense program")
                    Return False
                End If
            End If
        End If
        Dim status_value As Double = -1
        Dim sTime As Long = DateTime.Now.Ticks
        While Not (status_value = 0) And (DateTime.Now.Ticks - sTime) / 10000 < timeOut
            If m_TriCtrl.GetProcVariable("PROC_STATUS", 8, status_value) Then
                If status_value = 0 Then
                    ' the process has stopped
                    Exit While
                Else
                    Sleep(50)
                    Application.DoEvents()
                End If
            End If
        End While
        If (DateTime.Now.Ticks - sTime) / 10000 > timeOut Then
            MessageBox.Show("Timeout while wait for dispense process to stop")
            Return False
        End If

        If Not m_TriCtrl.Stop("CALIBRATIONS") Then
            If Not m_TriCtrl.Stop("CALIBRATIONS") Then
                If Not m_TriCtrl.Stop("CALIBRATIONS") Then
                    MessageBox.Show("Failed to stop calibration program")
                    Return False
                End If
            End If
        End If
        ResetCalibrationFlag()
        sTime = DateTime.Now.Ticks
        While Not (status_value = 0) And (DateTime.Now.Ticks - sTime) / 10000 < timeOut
            If m_TriCtrl.GetProcVariable("PROC_STATUS", 3, status_value) Then
                If status_value = 0 Then
                    ' the process has stopped
                    Exit While
                Else
                    Sleep(50)
                    Application.DoEvents()
                End If
            End If
        End While
        If (DateTime.Now.Ticks - sTime) / 10000 > timeOut Then
            MessageBox.Show("Timeout while wait for calibration process to stop")
            Return False
        End If

        If Not m_TriCtrl.Stop("SETDATUM") Then
            If Not m_TriCtrl.Stop("SETDATUM") Then
                If Not m_TriCtrl.Stop("SETDATUM") Then
                    MessageBox.Show("Failed to stop homing program")
                    Return False
                End If
            End If
        End If

        sTime = DateTime.Now.Ticks
        While Not (status_value = 0) And (DateTime.Now.Ticks - sTime) / 10000 < timeOut
            If m_TriCtrl.GetProcVariable("PROC_STATUS", 3, status_value) Then
                If status_value = 0 Then
                    ' the process has stopped
                    Exit While
                Else
                    Sleep(50)
                    Application.DoEvents()
                End If
            End If
        End While
        If (DateTime.Now.Ticks - sTime) / 10000 > timeOut Then
            MessageBox.Show("Timeout while wait for homing process to stop")
            Return False
        End If
        'If m_TriCtrl.Stop("CALIBRATIONS") Then
        '    ResetCalibrationFlag()
        'End If
        m_TriCtrl.RapidStop() 'Then Console.WriteLine("Rapid stop Done")
        m_TriCtrl.Cancel(0, 0) 'Then Console.WriteLine("Cancel 0 Done")
        m_TriCtrl.Cancel(0, 1) 'Then Console.WriteLine("Cancel 1 Done")
        m_TriCtrl.Cancel(0, 2) 'Then Console.WriteLine("Cancel 2 Done")
        m_TriCtrl.Cancel(1, 0) 'Then Console.WriteLine("Cancel 1,0 Done")
        m_TriCtrl.Cancel(1, 1) 'Then Console.WriteLine("Cancel 1,1 Done")
        m_TriCtrl.Cancel(1, 2) 'Then Console.WriteLine("Cancel 1,2 Done")
        m_TriCtrl.Op(27, 0)
        m_TriCtrl.Op(25, 0)
        'm_TriCtrl.Stop("SETDATUM")
        'If m_TriCtrl.Stop("CALIBRATIONS") Then
        '    ResetCalibrationFlag()
        'End If
        'If Not (m_TriCtrl.Stop("DISPENSE")) Then
        '    If Not (m_TriCtrl.Stop("DISPENSE")) Then
        '        m_TriCtrl.Stop("DISPENSE")
        '    End If
        'End If
        Return True
    End Function
#End Region
    Public Function IsOpen() As Boolean
        Return m_TriCtrl.IsOpen(0)
    End Function
    Public Sub Disconnect_Controller()
        If asynConnect Then
            robot_Asyn.Close(-1)
        End If
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
        If motionDoneEvent(0) Is Nothing Then
            motionDoneEvent(0) = New ManualResetEvent(False)
            motionDoneEvent(1) = New ManualResetEvent(False)
            checkRemainEvent(0) = New ManualResetEvent(False)
            checkRemainEvent(1) = New ManualResetEvent(False)
            checkRemainEvent(2) = New ManualResetEvent(False)
        End If
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

    Private Sub Received0() Handles robot_Asyn.OnReceiveChannel0
        'Console.WriteLine("#Channel 0 Triggered")
        Dim str As String
        If robot_Asyn.GetData(0, str) Then
            'Console.WriteLine("#Channel 0 Data: " + str)
        End If
    End Sub
    Private Sub Received5() Handles robot_Asyn.OnReceiveChannel5
        'Console.WriteLine("#Channel 5 Triggered")
        Dim str As String
        If robot_Asyn.GetData(5, str) Then
            'Console.WriteLine("#Channel 5 Data: " + str)
        End If
    End Sub
    Private Sub Received6() Handles robot_Asyn.OnReceiveChannel6
        'Console.WriteLine("#Channel 6 Triggered")
        Dim str As String
        If robot_Asyn.GetData(6, str) Then
            'Console.WriteLine("#Channel 6 Data: " + str)
        End If
    End Sub
    Private Sub Received7() Handles robot_Asyn.OnReceiveChannel7
        'Console.WriteLine("#Channel 7 Triggered")
        Dim str As String
        If robot_Asyn.GetData(7, str) Then
            'Console.WriteLine("#Channel 7 Data: " + str)
        End If
    End Sub
    Private Sub Received9() Handles robot_Asyn.OnReceiveChannel9
        'Console.WriteLine("#Channel 9 Triggered")
        Dim str As String
        If robot_Asyn.GetData(9, str) Then
            'Console.WriteLine("#Channel 9 Data: " + str)
        End If
    End Sub

    Public Sub SetOutputOn()
        If robot_Asyn.IsOpen(23) Then
            If robot_Asyn.SendData(0, "OP(25,1)") Then
                'Console.WriteLine("#1 Channel 0 Data Send")
            End If
        End If
    End Sub
    Public Sub SetOutputOff()
        If robot_Asyn.IsOpen(23) Then
            If robot_Asyn.SendData(0, "OP(25,0)") Then
                'Console.WriteLine("#2 Channel 0 Data Send")
            End If
        End If
    End Sub
#Region "GUI for stepper"
    Dim dStepVal(5) As Single
    Dim gStepCtrlSpeed As Double = 100.0

    Private Sub BTStepXminus_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTStepXminus.Click

        If Get_Limit("X-") Or MachineRunning() Or MachineHoming() Or Calibrating() Then
            Exit Sub
        End If

        dStepVal(0) = 0
        dStepVal(1) = gStepCtrlSpeed
        'dStepVal(2) = -ComboBoxFineStep.Value
        If rbtSteps.Checked Then
            dStepVal(2) = -ComboBoxFineStep.Value
        Else
            dStepVal(2) = -stepSize
        End If
        dStepVal(3) = 0
        dStepVal(4) = 0
        DoStep(dStepVal)
        'Me.SetOutputOff()
    End Sub

    Private Sub BTStepXplus_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTStepXplus.Click

        If Get_Limit("X+") Or MachineRunning() Or MachineHoming() Or Calibrating() Then
            Exit Sub
        End If

        dStepVal(0) = 0
        dStepVal(1) = gStepCtrlSpeed
        'dStepVal(2) = ComboBoxFineStep.Value
        If rbtSteps.Checked Then
            dStepVal(2) = ComboBoxFineStep.Value
        Else
            dStepVal(2) = stepSize
        End If
        dStepVal(3) = 0
        dStepVal(4) = 0
        DoStep(dStepVal)
        'Me.SetOutputOn()
    End Sub

    Private Sub BTStepYminus_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTStepYminus.Click
        If Get_Limit("Y-") Or MachineRunning() Or MachineHoming() Or Calibrating() Then
            Exit Sub
        End If

        dStepVal(0) = 0
        dStepVal(1) = gStepCtrlSpeed
        dStepVal(2) = 0
        'dStepVal(3) = -ComboBoxFineStep.Value
        If rbtSteps.Checked Then
            dStepVal(3) = -ComboBoxFineStep.Value
        Else
            dStepVal(3) = -stepSize
        End If
        dStepVal(4) = 0
        DoStep(dStepVal)

    End Sub

    Private Sub BTStepYplus_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTStepYplus.Click
        'Me.m_TriCtrl.AboutBox()
        If Get_Limit("Y+") Or MachineRunning() Or MachineHoming() Or Calibrating() Then
            Exit Sub
        End If

        dStepVal(0) = 0
        dStepVal(1) = gStepCtrlSpeed
        dStepVal(2) = 0
        If rbtSteps.Checked Then
            dStepVal(3) = ComboBoxFineStep.Value
        Else
            dStepVal(3) = stepSize
        End If

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
        'dStepVal(4) = ComboBoxFineStep.Value
        If rbtSteps.Checked Then
            dStepVal(4) = ComboBoxFineStep.Value
        Else
            dStepVal(4) = stepSize
        End If
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
        'dStepVal(4) = -ComboBoxFineStep.Value
        If rbtSteps.Checked Then
            dStepVal(4) = -ComboBoxFineStep.Value
        Else
            dStepVal(4) = -stepSize
        End If
        DoStep(dStepVal)

    End Sub

    Private Function Get_Limit(ByVal Direction As String) As Boolean

        Dim Read_Value As Integer = 0

        m_TriCtrl.In(0, 5, Read_Value)

        Select Case Direction
            Case "X+"
                Read_Value = Read_Value And &H2
                If (Read_Value = 0) Then
                    MsgBox("Machine reach the X+ limit sensor position!", MsgBoxStyle.OKOnly, "Limit Error")
                    Return True
                Else
                    Return False
                End If
            Case "X-"
                Read_Value = Read_Value And &H1
                If (Read_Value = 0) Then
                    MsgBox("Machine reach the X- limit sensor position!", MsgBoxStyle.OKOnly, "Limit Error")
                    Return True
                Else
                    Return False
                End If
            Case "Y+"
                Read_Value = Read_Value And &H4
                If (Read_Value = 0) Then
                    MsgBox("Machine reach the Y+ limit sensor position!", MsgBoxStyle.OKOnly, "Limit Error")
                    Return True
                Else
                    Return False
                End If
            Case "Y-"
                Read_Value = Read_Value And &H8
                If (Read_Value = 0) Then
                    MsgBox("Machine reach the Y- limit sensor position!", MsgBoxStyle.OKOnly, "Limit Error")
                    Return True
                Else
                    Return False
                End If
            Case "Up"
                Read_Value = Read_Value And &H10
                If (Read_Value = 0) Then
                    MsgBox("Machine reach the Z+ limit sensor position!", MsgBoxStyle.OKOnly, "Limit Error")
                    Return True
                Else
                    Return False
                End If
            Case "Dn"
                Read_Value = Read_Value And &H20
                If (Read_Value = 0) Then
                    MsgBox("Machine reach the Z- limit sensor position!", MsgBoxStyle.OKOnly, "Limit Error")
                    Return True
                Else
                    Return False
                End If
            Case Else
                Return False
        End Select



    End Function


#End Region

    Private Sub rbtSteps_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbtSteps.CheckedChanged
        ComboBoxFineStep.Enabled = True
        stepSize = ComboBoxFineStep.Value
    End Sub

    Private Sub rbt50micron_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbt50micron.CheckedChanged
        ComboBoxFineStep.Enabled = False
        stepSize = 0.05
    End Sub

    Private Sub rbt100micron_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbt100micron.CheckedChanged
        ComboBoxFineStep.Enabled = False
        stepSize = 0.1
    End Sub

    Private Sub rbt200micron_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbt200micron.CheckedChanged
        ComboBoxFineStep.Enabled = False
        stepSize = 0.2
    End Sub

    Private Sub rbt500_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbt500.CheckedChanged
        ComboBoxFineStep.Enabled = False
        stepSize = 0.5
    End Sub

    Private Sub rbt1micron_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbt1micron.CheckedChanged
        ComboBoxFineStep.Enabled = False
        stepSize = 0.001
    End Sub

    Private Sub rbt2micron_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbt2micron.CheckedChanged
        ComboBoxFineStep.Enabled = False
        stepSize = 0.002
    End Sub

    Private Sub rbt5micron_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbt5micron.CheckedChanged
        ComboBoxFineStep.Enabled = False
        stepSize = 0.005
    End Sub

    Private Sub rbt10micron_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbt10micron.CheckedChanged
        ComboBoxFineStep.Enabled = False
        stepSize = 0.01
    End Sub

    Private Sub rtb20micron_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rtb20micron.CheckedChanged
        ComboBoxFineStep.Enabled = False
        stepSize = 0.02
    End Sub
End Class
