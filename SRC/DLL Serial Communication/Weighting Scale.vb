Imports System.Text
Imports System.io

Public Class Weighting_Scale
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
    Friend WithEvents AxMSComm1 As AxMSCommLib.AxMSComm
    Friend WithEvents trimmed_result_display As System.Windows.Forms.TextBox
    Friend WithEvents VeryStable As System.Windows.Forms.Button
    Friend WithEvents Stable As System.Windows.Forms.Button
    Friend WithEvents VeryUnstable As System.Windows.Forms.Button
    Friend WithEvents Unstable As System.Windows.Forms.Button
    Friend WithEvents Button6 As System.Windows.Forms.Button
    Friend WithEvents Button7 As System.Windows.Forms.Button
    Friend WithEvents Button8 As System.Windows.Forms.Button
    Friend WithEvents number_of_result_display As System.Windows.Forms.TextBox
    Friend WithEvents Button10 As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents raw_result_display As System.Windows.Forms.TextBox
    Friend WithEvents DisplayRaw As System.Windows.Forms.TextBox
    Friend WithEvents DisplayStable As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents number_of_missed_result_display As System.Windows.Forms.TextBox
    Friend WithEvents Counter As System.Timers.Timer
    Friend WithEvents ButtonGetWeight As System.Windows.Forms.Button
    Friend WithEvents ButtonTare As System.Windows.Forms.Button
    Friend WithEvents ButtonReset As System.Windows.Forms.Button
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents ButtonOpenPort As System.Windows.Forms.Button
    Friend WithEvents ButtonClosePort As System.Windows.Forms.Button
    Friend WithEvents tbComData As System.Windows.Forms.TextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(Weighting_Scale))
        Me.AxMSComm1 = New AxMSCommLib.AxMSComm
        Me.ButtonGetWeight = New System.Windows.Forms.Button
        Me.trimmed_result_display = New System.Windows.Forms.TextBox
        Me.ButtonTare = New System.Windows.Forms.Button
        Me.DisplayRaw = New System.Windows.Forms.TextBox
        Me.VeryStable = New System.Windows.Forms.Button
        Me.Stable = New System.Windows.Forms.Button
        Me.VeryUnstable = New System.Windows.Forms.Button
        Me.Unstable = New System.Windows.Forms.Button
        Me.Button6 = New System.Windows.Forms.Button
        Me.Button7 = New System.Windows.Forms.Button
        Me.Button8 = New System.Windows.Forms.Button
        Me.number_of_result_display = New System.Windows.Forms.TextBox
        Me.Button10 = New System.Windows.Forms.Button
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.raw_result_display = New System.Windows.Forms.TextBox
        Me.ButtonReset = New System.Windows.Forms.Button
        Me.DisplayStable = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.ButtonOpenPort = New System.Windows.Forms.Button
        Me.ButtonClosePort = New System.Windows.Forms.Button
        Me.number_of_missed_result_display = New System.Windows.Forms.TextBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.Counter = New System.Timers.Timer
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.Panel2 = New System.Windows.Forms.Panel
        Me.tbComData = New System.Windows.Forms.TextBox
        CType(Me.AxMSComm1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Counter, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel1.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.SuspendLayout()
        '
        'AxMSComm1
        '
        Me.AxMSComm1.Enabled = True
        Me.AxMSComm1.Location = New System.Drawing.Point(792, 0)
        Me.AxMSComm1.Name = "AxMSComm1"
        Me.AxMSComm1.OcxState = CType(resources.GetObject("AxMSComm1.OcxState"), System.Windows.Forms.AxHost.State)
        Me.AxMSComm1.Size = New System.Drawing.Size(38, 38)
        Me.AxMSComm1.TabIndex = 0
        '
        'ButtonGetWeight
        '
        Me.ButtonGetWeight.Location = New System.Drawing.Point(96, 40)
        Me.ButtonGetWeight.Name = "ButtonGetWeight"
        Me.ButtonGetWeight.Size = New System.Drawing.Size(128, 24)
        Me.ButtonGetWeight.TabIndex = 1
        Me.ButtonGetWeight.Text = "Get Weight"
        '
        'trimmed_result_display
        '
        Me.trimmed_result_display.Location = New System.Drawing.Point(232, 40)
        Me.trimmed_result_display.Name = "trimmed_result_display"
        Me.trimmed_result_display.Size = New System.Drawing.Size(104, 27)
        Me.trimmed_result_display.TabIndex = 2
        Me.trimmed_result_display.Text = ""
        '
        'ButtonTare
        '
        Me.ButtonTare.Location = New System.Drawing.Point(96, 8)
        Me.ButtonTare.Name = "ButtonTare"
        Me.ButtonTare.Size = New System.Drawing.Size(128, 24)
        Me.ButtonTare.TabIndex = 1
        Me.ButtonTare.Text = "Tare and Zero"
        '
        'DisplayRaw
        '
        Me.DisplayRaw.AutoSize = False
        Me.DisplayRaw.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DisplayRaw.Location = New System.Drawing.Point(8, 112)
        Me.DisplayRaw.Multiline = True
        Me.DisplayRaw.Name = "DisplayRaw"
        Me.DisplayRaw.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.DisplayRaw.Size = New System.Drawing.Size(336, 248)
        Me.DisplayRaw.TabIndex = 2
        Me.DisplayRaw.Text = "All weight values displayed here"
        '
        'VeryStable
        '
        Me.VeryStable.Location = New System.Drawing.Point(8, 56)
        Me.VeryStable.Name = "VeryStable"
        Me.VeryStable.Size = New System.Drawing.Size(128, 24)
        Me.VeryStable.TabIndex = 1
        Me.VeryStable.Text = "Very Stable"
        '
        'Stable
        '
        Me.Stable.Location = New System.Drawing.Point(8, 88)
        Me.Stable.Name = "Stable"
        Me.Stable.Size = New System.Drawing.Size(128, 24)
        Me.Stable.TabIndex = 1
        Me.Stable.Text = "Stable"
        '
        'VeryUnstable
        '
        Me.VeryUnstable.Location = New System.Drawing.Point(8, 152)
        Me.VeryUnstable.Name = "VeryUnstable"
        Me.VeryUnstable.Size = New System.Drawing.Size(128, 24)
        Me.VeryUnstable.TabIndex = 1
        Me.VeryUnstable.Text = "Very Unstable"
        '
        'Unstable
        '
        Me.Unstable.Location = New System.Drawing.Point(8, 120)
        Me.Unstable.Name = "Unstable"
        Me.Unstable.Size = New System.Drawing.Size(128, 24)
        Me.Unstable.TabIndex = 1
        Me.Unstable.Text = "Unstable"
        '
        'Button6
        '
        Me.Button6.Location = New System.Drawing.Point(8, 432)
        Me.Button6.Name = "Button6"
        Me.Button6.Size = New System.Drawing.Size(152, 24)
        Me.Button6.TabIndex = 1
        Me.Button6.Text = "Restart"
        '
        'Button7
        '
        Me.Button7.Location = New System.Drawing.Point(8, 464)
        Me.Button7.Name = "Button7"
        Me.Button7.Size = New System.Drawing.Size(152, 24)
        Me.Button7.TabIndex = 1
        Me.Button7.Text = "Calibrate"
        '
        'Button8
        '
        Me.Button8.Location = New System.Drawing.Point(8, 496)
        Me.Button8.Name = "Button8"
        Me.Button8.Size = New System.Drawing.Size(152, 24)
        Me.Button8.TabIndex = 1
        Me.Button8.Text = "Internal Calibrate"
        '
        'number_of_result_display
        '
        Me.number_of_result_display.AutoSize = False
        Me.number_of_result_display.Location = New System.Drawing.Point(16, 240)
        Me.number_of_result_display.Name = "number_of_result_display"
        Me.number_of_result_display.Size = New System.Drawing.Size(104, 40)
        Me.number_of_result_display.TabIndex = 2
        Me.number_of_result_display.Text = ""
        '
        'Button10
        '
        Me.Button10.Location = New System.Drawing.Point(8, 528)
        Me.Button10.Name = "Button10"
        Me.Button10.Size = New System.Drawing.Size(152, 24)
        Me.Button10.TabIndex = 1
        Me.Button10.Text = "Print model"
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(16, 192)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(104, 40)
        Me.Label1.TabIndex = 4
        Me.Label1.Text = "Successful Readings"
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(16, 8)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(104, 40)
        Me.Label2.TabIndex = 4
        Me.Label2.Text = "Ambient Conditions"
        '
        'Label3
        '
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(32, 384)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(104, 40)
        Me.Label3.TabIndex = 4
        Me.Label3.Text = "Misc Functions"
        '
        'raw_result_display
        '
        Me.raw_result_display.Location = New System.Drawing.Point(232, 8)
        Me.raw_result_display.Name = "raw_result_display"
        Me.raw_result_display.Size = New System.Drawing.Size(104, 27)
        Me.raw_result_display.TabIndex = 2
        Me.raw_result_display.Text = ""
        '
        'ButtonReset
        '
        Me.ButtonReset.Location = New System.Drawing.Point(8, 40)
        Me.ButtonReset.Name = "ButtonReset"
        Me.ButtonReset.Size = New System.Drawing.Size(88, 24)
        Me.ButtonReset.TabIndex = 1
        Me.ButtonReset.Text = "Reset"
        '
        'DisplayStable
        '
        Me.DisplayStable.AutoSize = False
        Me.DisplayStable.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DisplayStable.Location = New System.Drawing.Point(352, 112)
        Me.DisplayStable.Multiline = True
        Me.DisplayStable.Name = "DisplayStable"
        Me.DisplayStable.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.DisplayStable.Size = New System.Drawing.Size(336, 544)
        Me.DisplayStable.TabIndex = 2
        Me.DisplayStable.Text = "All stable weight values displayed here"
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(112, 88)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(88, 24)
        Me.Label4.TabIndex = 5
        Me.Label4.Text = "Raw values"
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(464, 88)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(128, 24)
        Me.Label5.TabIndex = 5
        Me.Label5.Text = "Stable readings"
        '
        'ButtonOpenPort
        '
        Me.ButtonOpenPort.Location = New System.Drawing.Point(464, 16)
        Me.ButtonOpenPort.Name = "ButtonOpenPort"
        Me.ButtonOpenPort.Size = New System.Drawing.Size(128, 24)
        Me.ButtonOpenPort.TabIndex = 1
        Me.ButtonOpenPort.Text = "Open Port"
        '
        'ButtonClosePort
        '
        Me.ButtonClosePort.Location = New System.Drawing.Point(464, 48)
        Me.ButtonClosePort.Name = "ButtonClosePort"
        Me.ButtonClosePort.Size = New System.Drawing.Size(128, 24)
        Me.ButtonClosePort.TabIndex = 1
        Me.ButtonClosePort.Text = "Close Port"
        '
        'number_of_missed_result_display
        '
        Me.number_of_missed_result_display.AutoSize = False
        Me.number_of_missed_result_display.Location = New System.Drawing.Point(16, 336)
        Me.number_of_missed_result_display.Name = "number_of_missed_result_display"
        Me.number_of_missed_result_display.Size = New System.Drawing.Size(104, 40)
        Me.number_of_missed_result_display.TabIndex = 2
        Me.number_of_missed_result_display.Text = ""
        '
        'Label6
        '
        Me.Label6.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.Location = New System.Drawing.Point(16, 288)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(104, 40)
        Me.Label6.TabIndex = 4
        Me.Label6.Text = "Missed Readings"
        '
        'Counter
        '
        Me.Counter.Interval = 50
        Me.Counter.SynchronizingObject = Me
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.Button6)
        Me.Panel1.Controls.Add(Me.Button10)
        Me.Panel1.Controls.Add(Me.Label2)
        Me.Panel1.Controls.Add(Me.Label3)
        Me.Panel1.Controls.Add(Me.number_of_missed_result_display)
        Me.Panel1.Controls.Add(Me.Label6)
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Controls.Add(Me.VeryStable)
        Me.Panel1.Controls.Add(Me.Stable)
        Me.Panel1.Controls.Add(Me.VeryUnstable)
        Me.Panel1.Controls.Add(Me.Unstable)
        Me.Panel1.Controls.Add(Me.Button7)
        Me.Panel1.Controls.Add(Me.Button8)
        Me.Panel1.Controls.Add(Me.number_of_result_display)
        Me.Panel1.Location = New System.Drawing.Point(696, 96)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(160, 560)
        Me.Panel1.TabIndex = 6
        '
        'Panel2
        '
        Me.Panel2.Controls.Add(Me.trimmed_result_display)
        Me.Panel2.Controls.Add(Me.raw_result_display)
        Me.Panel2.Controls.Add(Me.ButtonReset)
        Me.Panel2.Controls.Add(Me.ButtonTare)
        Me.Panel2.Controls.Add(Me.ButtonGetWeight)
        Me.Panel2.Location = New System.Drawing.Point(8, 8)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(352, 72)
        Me.Panel2.TabIndex = 7
        '
        'tbComData
        '
        Me.tbComData.AutoSize = False
        Me.tbComData.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tbComData.Location = New System.Drawing.Point(8, 376)
        Me.tbComData.Multiline = True
        Me.tbComData.Name = "tbComData"
        Me.tbComData.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.tbComData.Size = New System.Drawing.Size(336, 272)
        Me.tbComData.TabIndex = 8
        Me.tbComData.Text = ""
        '
        'Weighting_Scale
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(8, 20)
        Me.ClientSize = New System.Drawing.Size(864, 662)
        Me.Controls.Add(Me.tbComData)
        Me.Controls.Add(Me.DisplayRaw)
        Me.Controls.Add(Me.DisplayStable)
        Me.Controls.Add(Me.Panel2)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.AxMSComm1)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.ButtonOpenPort)
        Me.Controls.Add(Me.ButtonClosePort)
        Me.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Name = "Weighting_Scale"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Weighting_Scale"
        CType(Me.AxMSComm1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Counter, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel1.ResumeLayout(False)
        Me.Panel2.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region
    'private
    Dim WeightCommError As Boolean = False
    Dim WeightOverLoad As Boolean = False
    Dim expected_reading_count As Integer = 0
    Dim successful_reading_count As Integer = 0
    Private Declare Sub Sleep Lib "kernel32" Alias "Sleep" (ByVal dwMilliseconds As Long)  'delay function
    Dim newline_count As Integer
    Dim raw_result_store As String
    Dim stable_result_store As String
    Dim TimeOfLastReading As DateTime
    Dim CurrentTime As DateTime
    Dim TimeSinceLastReading As TimeSpan
    'number of readings to test for. get the median value of these readings
    Const number_of_values As Integer = 5
    Public Values(number_of_values - 1) As Double
    Public ValuesIndex As Integer = 0
    'exposed variables
    Public WeightReading As Double = 0
    Public ValueUpdated As Boolean = False 'when you turn on, use this flag to check whether the axmscomm has updated the reading
    Public Taring As Boolean = False 'flag to read values until volume calibration cell is tared
    Private errorMessage As String = ""
    Private errorCode As Integer = 0
    Private tempCnt As Integer = 0
    Public portOpened As Boolean = False
    Public ReadOnly Property returnString() As String
        Get
            Return WeightReading.ToString()
        End Get
    End Property

    Public Delegate Sub WeightingScaleReturnedDel()
    Public WeightingScaleReturned As WeightingScaleReturnedDel

    Public Function OpenPort() As Boolean
        If AxMSComm1.PortOpen = True Then
            EnableButtons()
            Return True
        End If
        With AxMSComm1
            .CommPort = WeighingPort
            .Settings = "1200,O,7,1"
            .InputMode = MSCommLib.InputModeConstants.comInputModeText
            .InBufferCount = 0
            .InBufferSize = 16
            .InputLen = 16
            .RThreshold = 16
        End With
        Try
            If Not AxMSComm1.PortOpen Then
                AxMSComm1.PortOpen = True
            End If
            portOpened = True
        Catch ex As System.Runtime.InteropServices.COMException
            ExceptionDisplay(ex)
            Return False
        End Try
        CommandFormat1("Very Unstable")
        ResetValues()
        DoTare()
        EnableButtons()
        Return True
    End Function

    Public Sub ClosePort()
        If AxMSComm1.PortOpen = True Then
            AxMSComm1.PortOpen = False
            portOpened = False
        End If
        DisableButtons()
        DisablePrintingWatcher()
    End Sub

    Private Sub EnablePrintingWatcher()
        Counter.Enabled = True
    End Sub

    Private Sub DisablePrintingWatcher()
        Counter.Enabled = False
    End Sub

    Private Sub ClearBuffer()
        AxMSComm1.InBufferCount = 0
    End Sub

    Private Function WithinTolerance(ByVal val1 As Double, ByVal val2 As Double) As Boolean
        '100 mg / 0.1 g tolerance
        Return Math.Abs(val2 - val1) < 0.3
    End Function

    Public Sub DoTare()
        Taring = True
        CommandFormat1("Tare1")
        Taring = False
        'raw_result_store = "Taring Now" + vbCrLf + raw_result_store
        'DisplayRaw.Text += "Taring Now" + vbCrLf

        'start_time = Now
        'successful_reading_count = 0
        'expected_reading_count = 1
        'Dim missed_results As Integer = expected_reading_count - successful_reading_count
        'number_of_result_display.Text = successful_reading_count.ToString
        'number_of_missed_result_display.Text = missed_results.ToString

        'ValueUpdated = False
        'CommandFormat1("Print")
    End Sub

    Public Sub Restart()
        CommandFormat1("Restart")
    End Sub

    Public Sub Zero()
        CommandFormat1("Zero")
    End Sub

    'Private Sub AxMSComm1_OnComm(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AxMSComm1.OnComm

    '    Dim raw_result As Object
    '    Dim result As Double
    '    Dim string_result As String
    '    Dim char_result As Char()

    '    'If missed_readings <= 0 Then
    '    '    ResetValues()
    '    '    ClearBuffer()
    '    '    DisablePrintingWatcher()
    '    '    Exit Sub
    '    'End If

    '    If Not Taring And ValueUpdated Then
    '        ClearBuffer()
    '    End If
    '    tbComData.Text += "####Trigger####" + vbLf
    '    Try
    '        With AxMSComm1

    '            If AxMSComm1.InBufferCount > 0 Then
    '                .RThreshold = 0 'prevent generation of events
    '                Dim valid_reading As Boolean = False 'whether the reading is acceptable
    '                raw_result = AxMSComm1.Input
    '                char_result = CType(raw_result, Char())
    '                string_result = CType(raw_result, String)
    '                tempCnt += 1
    '                tbComData.Text += "#" + tempCnt.ToString() + string_result + vbLf
    '                tbComData.Select(0, tbComData.Text.Length)
    '                tbComData.ScrollToCaret()
    '                If char_result.Length <> 16 Then
    '                    Exit Sub
    '                End If

    '                Dim start_index As Integer
    '                Dim positive As Boolean
    '                For i As Integer = 0 To 15
    '                    If char_result(i) = "+" Then
    '                        start_index = i
    '                        positive = True
    '                        valid_reading = True
    '                    ElseIf char_result(i) = "-" Then
    '                        start_index = i
    '                        positive = False
    '                        valid_reading = True
    '                    End If
    '                Next

    '                'if start index = 0, index 4 = 1, index 5 = decimal point, index 6 = 0.1, index 7 = 
    '                Dim hundreds_index As Integer = start_index + 2
    '                Dim tens_index As Integer = start_index + 3
    '                Dim ones_index As Integer = start_index + 4
    '                Dim point_tens_index As Integer = start_index + 6
    '                Dim point_hundreds_index As Integer = start_index + 7
    '                Dim point_thousands_index As Integer = start_index + 8
    '                Dim point_ten_thousands_index As Integer = start_index + 9

    '                If char_result(5) = "H" Then
    '                    result = 999.999            'means overloaded
    '                ElseIf char_result(5) = "L" Then
    '                    result = -999.999           'means underloaded
    '                ElseIf char_result(3) = "E" Then
    '                    result = 888.0              'means weighing machine internal error
    '                ElseIf char_result(0) = Nothing Then
    '                    result = 777.777 'it means time out error
    '                End If

    '                If valid_reading = True Then

    '                    'this is for reading results
    '                    If hundreds_index > 15 Then hundreds_index = hundreds_index - 16
    '                    If tens_index > 15 Then tens_index = tens_index - 16
    '                    If ones_index > 15 Then ones_index = ones_index - 16
    '                    If point_tens_index > 15 Then point_tens_index = point_tens_index - 16
    '                    If point_hundreds_index > 15 Then point_hundreds_index = point_hundreds_index - 16
    '                    If point_thousands_index > 15 Then point_thousands_index = point_thousands_index - 16
    '                    If point_ten_thousands_index > 15 Then point_ten_thousands_index = point_ten_thousands_index - 16
    '                    'see the sartorius weighing scale communication protocol to understand.
    '                    result = Val(char_result(hundreds_index)) * 100 + Val(char_result(tens_index)) * 10 + Val(char_result(ones_index)) + Val(char_result(point_tens_index)) / 10 + Val(char_result(point_hundreds_index)) / 100 + Val(char_result(point_thousands_index)) / 1000 + Val(char_result(point_ten_thousands_index)) / 10000
    '                    If Not positive Then
    '                        result = -result
    '                    End If
    '                    'display the result in mg
    '                    result = result * 1000

    '                    'this is for interface window (displaying of results)
    '                    raw_result_display.Text = string_result
    '                    trimmed_result_display.Text = result.ToString
    '                    stop_time = Now
    '                    elapsed_time = stop_time.Subtract(start_time)
    '                    raw_result_store = "Time: " + elapsed_time.TotalSeconds.ToString + " seconds. Reading: " + result.ToString + " mg" + vbCrLf + raw_result_store
    '                    DisplayRaw.Text = raw_result_store
    '                    newline_count += 1
    '                    If newline_count = 25 Then
    '                        raw_result_store = ""
    '                        DisplayRaw.Text = raw_result_store
    '                        newline_count = 0
    '                    End If
    '                    successful_reading_count += 1
    '                    number_of_result_display.Text = successful_reading_count.ToString
    '                    Dim missed_readings As Integer = expected_reading_count - successful_reading_count
    '                    number_of_missed_result_display.Text = CStr(missed_readings)
    '                    TimeOfLastReading = Now

    '                    'this is for outside modules (somewhat mixed with display)
    '                    If Taring Then
    '                        If WithinTolerance(result, 0) Then
    '                            WeightReading = 0
    '                            raw_result_store = "Taring Done" + vbCrLf + raw_result_store
    '                            DisplayRaw.Text = raw_result_store
    '                            Taring = False
    '                            stable_result_store = "Taring done in " + elapsed_time.TotalSeconds.ToString + " seconds." + vbCrLf + stable_result_store
    '                        Else
    '                            DoTare()
    '                        End If
    '                    ElseIf ValueUpdated = False Then
    '                        If ValuesIndex = number_of_values - 1 Then
    '                            Values(ValuesIndex) = result
    '                            Array.Sort(Values)
    '                            WeightReading = Values(number_of_values - 1)
    '                            raw_result_store = "Final reading of Weight Done" + vbCrLf + raw_result_store
    '                            DisplayRaw.Text = raw_result_store
    '                            ValueUpdated = True
    '                            ValuesIndex = 0
    '                            stable_result_store = "Final reading: " + CStr(WeightReading) + " mg obtained in " + elapsed_time.TotalSeconds.ToString + " seconds." + vbCrLf + stable_result_store
    '                        Else
    '                            stable_result_store = CStr(result) + " mg obtained in " + elapsed_time.TotalSeconds.ToString + " seconds." + vbCrLf + stable_result_store
    '                            raw_result_store = "Reading of Weight Done" + vbCrLf + raw_result_store
    '                            Values(ValuesIndex) = result
    '                            ValuesIndex += 1
    '                            Sleep(50)
    '                            GetWeight()
    '                        End If
    '                    End If
    '                Else
    '                    CommandFormat1("Print")
    '                End If


    '                'DTE (data terminal equipment) -> computer
    '                'DCE (data circuit-terminating equipment) -> calibration cell
    '                'CTS: DCE is ready to accept data from the DTE
    '                'DSR: DCE is ready to receive commands or data.
    '            Else
    '                'CommandFormat1("Print1")
    '            End If
    '        End With
    '        AxMSComm1.RThreshold = 16      'resume generation of events

    '    Catch ex As Exception
    '        MsgBox(ex.Message)
    '    End Try

    'End Sub
    Private Sub AxMSComm1_OnComm(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AxMSComm1.OnComm
        Dim raw_result As Object
        Dim result As Double
        Dim string_result As String
        Dim tempStr As String
        If Not Taring And ValueUpdated Then
            raw_result = AxMSComm1.Input
            string_result = CType(raw_result, String)
            ClearBuffer()
        End If
        'tbComData.Text += "####Trigger####" + vbLf
        Try
            With AxMSComm1

                If AxMSComm1.InBufferCount > 0 Then
                    .RThreshold = 0 'prevent generation of events
                    Dim valid_reading As Boolean = False 'whether the reading is acceptable
                    raw_result = AxMSComm1.Input
                    string_result = CType(raw_result, String)
                    tempCnt += 1
                    'tbComData.Text += "#" + tempCnt.ToString() + string_result + vbLf
                    'tbComData.Select(0, tbComData.Text.Length)
                    'tbComData.ScrollToCaret()
                    If string_result.Length <> 16 Then
                        Exit Sub
                    End If
                    Dim start_index As Integer
                    Dim positive As Boolean
                    tempStr = string_result.Substring(0, 1)
                    If tempStr = "+" Then
                        positive = True
                        valid_reading = True
                    ElseIf tempStr = "-" Then
                        positive = False
                        valid_reading = True
                    End If

                    If valid_reading Then
                        tempStr = string_result.Substring(3, 7)
                        tempStr = tempStr.Trim()
                        result = Convert.ToDouble(tempStr) * 1000 'Convert to mg
                        If Not positive Then
                            result = -result
                        End If
                        WeightReading = result
                        errorCode = 0
                        errorMessage = "Working fine"
                        stable_result_store = "Weight reading: " + CStr(result) + " mg obtained" + vbCrLf
                        Taring = False
                        ValueUpdated = True
                        'Console.WriteLine("Reading " + WeightReading.ToString())
                    Else
                        tempStr = string_result.Substring(5, 1)
                        If tempStr = "H" Then
                            errorCode = -1
                            errorMessage = "Overloaded"
                        ElseIf tempStr = "L" Then
                            errorCode = -2
                            errorMessage = "Underloaded"
                        ElseIf tempStr = "E" Then
                            errorCode = -3
                            errorMessage = "Internal Error"
                        Else
                            errorCode = -100
                            errorMessage = "Unknown Error"
                        End If
                        'stable_result_store = "Weight reading Failed: " + errorMessage + vbCrLf
                    End If

                End If
                'DTE (data terminal equipment) -> computer
                'DCE (data circuit-terminating equipment) -> calibration cell
                'CTS: DCE is ready to accept data from the DTE
                'DSR: DCE is ready to receive commands or data.
            End With
            AxMSComm1.RThreshold = 16      'resume generation of events
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Public Sub ResetValues()
        'for interface display
        'expected_reading_count = 0
        'successful_reading_count = 0
        'DisplayRaw.Text = "All weight values displayed here"
        'DisplayStable.Text = "All stable weight values displayed here"
        'raw_result_store = "Raw result"
        'stable_result_store = "Weight result"

        'for outside module
        ValuesIndex = 0
        Array.Clear(Values, 0, number_of_values - 1)
        ValueUpdated = False
        ClearBuffer()
        'CommandFormat1("Tare1")
    End Sub

    Sub CommandFormat1(ByVal command As String)
        Dim output_str As String
        output_str = Chr(27)
        Select Case command
            Case "Very Stable"
                output_str = output_str + Chr(75)
            Case "Stable"
                output_str = output_str + Chr(76)
            Case "Unstable"
                output_str = output_str + Chr(77)
            Case "Very Unstable"
                output_str = output_str + Chr(78)
            Case "Block Keys"
                output_str = output_str + Chr(79)
            Case "Print"
                output_str = output_str + Chr(80)
            Case "Acoustic Signal"
                output_str = output_str + Chr(81)
            Case "Unblock Keys"
                output_str = output_str + Chr(82)
            Case "Restart"
                output_str = output_str + Chr(83)
            Case "Tare1" 'tare and zero
                output_str = output_str + Chr(84)
            Case "Tare2" 'tare only
                output_str = output_str + Chr(85)
            Case "Zero"
                output_str = output_str + Chr(86)
            Case "Settings Calibrate"
                output_str = output_str + Chr(87)
            Case "Internal Calibrate"
                output_str = output_str + Chr(90)
            Case "Weighing Mode 1"
                output_str = output_str + Chr(90)
            Case "Weighing Mode 2"
                output_str = output_str + Chr(90)
            Case "Weighing Mode 3"
                output_str = output_str + Chr(90)
            Case "Weighing Mode 4"
                output_str = output_str + Chr(90)
        End Select
        output_str = output_str + vbCrLf
        Try
            AxMSComm1.Output = output_str
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonTare.Click
        DoTare()
    End Sub

    Private Sub VeryStable_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles VeryStable.Click
        CommandFormat1("Very Stable")
    End Sub

    Private Sub Stable_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Stable.Click
        CommandFormat1("Stable")
    End Sub

    Private Sub Unstable_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Unstable.Click
        CommandFormat1("Unstable")
    End Sub

    Private Sub VeryUnstable_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles VeryUnstable.Click
        CommandFormat1("Very Unstable")
    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        CommandFormat1("Block Keys")
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        CommandFormat1("Unblock Keys")
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        CommandFormat1("Restart")
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        CommandFormat1("Settings Calibrate")
    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        CommandFormat1("Internal Calibrate")
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        DisplayRaw.Text = "All weight values displayed here"
        DisplayStable.Text = "All stable weight values displayed here"
        raw_result_store = "Raw result"
        DisplayRaw.Text = raw_result_store
        stable_result_store = "Weight result"
    End Sub

    Public Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonGetWeight.Click
        GetWeight()
    End Sub

    Private Sub Button10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button10.Click
        AxMSComm1.Output = Chr(27) + Chr(120) + Chr(49) + Chr(95) + vbCrLf
    End Sub

    Public Sub GetWeight()
        'for interface display
        raw_result_store = "Reading Weight Now" + vbCrLf + raw_result_store
        DisplayRaw.Text = raw_result_store
        start_time = Now
        successful_reading_count = 0
        expected_reading_count = 1
        Dim missed_results As Integer = expected_reading_count - successful_reading_count
        number_of_result_display.Text = successful_reading_count.ToString
        number_of_missed_result_display.Text = missed_results.ToString

        'for outside module
        ValueUpdated = False
        CommandFormat1("Print")

    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonReset.Click
        expected_reading_count = 0
        successful_reading_count = 0
        DisplayRaw.Text = "All weight values displayed here"
        DisplayStable.Text = "All stable weight values displayed here"
        raw_result_store = "Raw result"
        DisplayRaw.Text = raw_result_store
        stable_result_store = "Weight result"
    End Sub

    Private Sub Button5_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs)
        CommandFormat1("Weighing Mode 1")
    End Sub

    Private Sub Button9_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs)
        CommandFormat1("Weighing Mode 2")
    End Sub

    Private Sub Button12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        CommandFormat1("Weighing Mode 3")
    End Sub

    Private Sub Button11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        CommandFormat1("Weighing Mode 4")
    End Sub

    Private Sub Button13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonOpenPort.Click
        OpenPort()
    End Sub

    Private Sub Button14_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonClosePort.Click
        ClosePort()
    End Sub

    'Private Sub Counter_Elapsed(ByVal sender As System.Object, ByVal e As System.Timers.ElapsedEventArgs) Handles Counter.Elapsed
    '    Dim missed_readings As Integer = expected_reading_count - successful_reading_count
    '    If missed_readings > 0 Then
    '        CurrentTime = Now
    '        TimeSinceLastReading = CurrentTime.Subtract(TimeOfLastReading)
    '        If TimeSinceLastReading.TotalSeconds > 1 Then
    '            AxMSComm1.RThreshold = 16
    '            CommandFormat1("Print")
    '        End If
    '    End If
    'End Sub

    Private Sub Weighting_Scale_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        If portOpened Then
            'If AxMSComm1.PortOpen Then
            EnableButtons()
        Else
            DisableButtons()
        End If
    End Sub

    'enable buttons for open port
    Private Sub EnableButtons()
        Panel1.Enabled = True
        Panel2.Enabled = True
        ButtonOpenPort.Enabled = False
        ButtonClosePort.Enabled = True
    End Sub

    'disable buttons for closed port
    Private Sub DisableButtons()
        Panel1.Enabled = False
        Panel2.Enabled = False
        ButtonOpenPort.Enabled = True
        ButtonClosePort.Enabled = False
    End Sub

    Private Sub Weighting_Scale_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        e.Cancel = True
        Hide()
    End Sub
End Class
