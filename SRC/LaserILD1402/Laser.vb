Imports System.Char
Imports System.Windows.Forms
Imports Laser

Public Class Laser
    Inherits System.Windows.Forms.Form
    Implements ILaser

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
    Friend WithEvents Display As System.Windows.Forms.TextBox
    Friend WithEvents Highest As System.Windows.Forms.TextBox
    Friend WithEvents Lowest As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Status As System.Windows.Forms.TextBox
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents ContinuousReadButton As System.Windows.Forms.Button
    Friend WithEvents lbReading As System.Windows.Forms.Label
    Friend WithEvents btTurnOffMeasurement As System.Windows.Forms.Button
    Friend WithEvents btTurnOnMeasurement As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(Laser))
        Me.AxMSComm1 = New AxMSCommLib.AxMSComm
        Me.Display = New System.Windows.Forms.TextBox
        Me.Highest = New System.Windows.Forms.TextBox
        Me.Lowest = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.Status = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.Button1 = New System.Windows.Forms.Button
        Me.ContinuousReadButton = New System.Windows.Forms.Button
        Me.lbReading = New System.Windows.Forms.Label
        Me.btTurnOffMeasurement = New System.Windows.Forms.Button
        Me.btTurnOnMeasurement = New System.Windows.Forms.Button
        CType(Me.AxMSComm1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'AxMSComm1
        '
        Me.AxMSComm1.Enabled = True
        Me.AxMSComm1.Location = New System.Drawing.Point(424, 280)
        Me.AxMSComm1.Name = "AxMSComm1"
        Me.AxMSComm1.OcxState = CType(resources.GetObject("AxMSComm1.OcxState"), System.Windows.Forms.AxHost.State)
        Me.AxMSComm1.Size = New System.Drawing.Size(38, 38)
        Me.AxMSComm1.TabIndex = 1
        '
        'Display
        '
        Me.Display.AutoSize = False
        Me.Display.HideSelection = False
        Me.Display.Location = New System.Drawing.Point(8, 8)
        Me.Display.Multiline = True
        Me.Display.Name = "Display"
        Me.Display.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.Display.Size = New System.Drawing.Size(384, 576)
        Me.Display.TabIndex = 3
        Me.Display.Text = ""
        '
        'Highest
        '
        Me.Highest.Location = New System.Drawing.Point(488, 40)
        Me.Highest.Name = "Highest"
        Me.Highest.TabIndex = 6
        Me.Highest.Text = "0"
        '
        'Lowest
        '
        Me.Lowest.Location = New System.Drawing.Point(488, 8)
        Me.Lowest.Name = "Lowest"
        Me.Lowest.TabIndex = 7
        Me.Lowest.Text = "99999999999999"
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(416, 8)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(64, 23)
        Me.Label1.TabIndex = 8
        Me.Label1.Text = "Lowest"
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(416, 48)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(72, 23)
        Me.Label2.TabIndex = 8
        Me.Label2.Text = "Highest"
        '
        'Status
        '
        Me.Status.Location = New System.Drawing.Point(488, 72)
        Me.Status.Name = "Status"
        Me.Status.TabIndex = 6
        Me.Status.Text = ""
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(416, 80)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(56, 23)
        Me.Label3.TabIndex = 8
        Me.Label3.Text = "Status"
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(416, 168)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(168, 23)
        Me.Button1.TabIndex = 9
        Me.Button1.Text = "Read Once"
        '
        'ContinuousReadButton
        '
        Me.ContinuousReadButton.Location = New System.Drawing.Point(416, 128)
        Me.ContinuousReadButton.Name = "ContinuousReadButton"
        Me.ContinuousReadButton.Size = New System.Drawing.Size(168, 23)
        Me.ContinuousReadButton.TabIndex = 9
        Me.ContinuousReadButton.Text = "Enable Continuous Read"
        '
        'lbReading
        '
        Me.lbReading.Location = New System.Drawing.Point(752, 80)
        Me.lbReading.Name = "lbReading"
        Me.lbReading.TabIndex = 10
        '
        'btTurnOffMeasurement
        '
        Me.btTurnOffMeasurement.Location = New System.Drawing.Point(416, 240)
        Me.btTurnOffMeasurement.Name = "btTurnOffMeasurement"
        Me.btTurnOffMeasurement.Size = New System.Drawing.Size(168, 23)
        Me.btTurnOffMeasurement.TabIndex = 11
        Me.btTurnOffMeasurement.Text = "Turn Off Measurement"
        '
        'btTurnOnMeasurement
        '
        Me.btTurnOnMeasurement.Location = New System.Drawing.Point(416, 200)
        Me.btTurnOnMeasurement.Name = "btTurnOnMeasurement"
        Me.btTurnOnMeasurement.Size = New System.Drawing.Size(168, 23)
        Me.btTurnOnMeasurement.TabIndex = 12
        Me.btTurnOnMeasurement.Text = "Turn On Measurement"
        '
        'Laser
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(8, 20)
        Me.ClientSize = New System.Drawing.Size(610, 592)
        Me.Controls.Add(Me.btTurnOnMeasurement)
        Me.Controls.Add(Me.btTurnOffMeasurement)
        Me.Controls.Add(Me.lbReading)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Highest)
        Me.Controls.Add(Me.Lowest)
        Me.Controls.Add(Me.Display)
        Me.Controls.Add(Me.Status)
        Me.Controls.Add(Me.AxMSComm1)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.ContinuousReadButton)
        Me.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Name = "Laser"
        Me.Text = "Laser"
        Me.TopMost = True
        CType(Me.AxMSComm1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

#Region "Declaration"
    'private variables for show in the interface window
    Private raw_result_store As String
    Private last_line As String
    Private newline_count As Integer
    Private DoRead As Boolean
    Private ContinuousRead As Boolean = False
    Private readCount As Integer = 0
    'exposed to other modules
    Private ls_Reading As Double = 0
    Private VUpdated As Boolean 'when you turn on, use this flag to check wether the axmscomm has updated the reading.
    Private LaserPort As Integer = 5
    Private start_time As DateTime
    Private stop_time As DateTime
    Private elapsed_time As TimeSpan
    Private afterCommand As Boolean = False
    Private needReturn As Boolean = False
    Friend Declare Sub Sleep Lib "kernel32" Alias "Sleep" (ByVal dwMilliseconds As Long)  'delay function
#End Region

#Region "Private"
    Private Sub Laser_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        OpenPort()
    End Sub

    Private Function ConvertReadingToMM(ByVal reading As String) As Double
        Dim double_val As Double = CDbl(reading)
        double_val = ((double_val * (1.02 / 16368)) - 0.01) * 20
        Return double_val
    End Function

    Private Function WithinTolerance(ByVal val1 As Double, ByVal val2 As Double) As Boolean
        '100 um / 0.1 mm tolerance
        Return Math.Abs(val2 - val1) < 0.1
    End Function

    Private Sub AxMSComm1_OnComm(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AxMSComm1.OnComm

        Dim raw_result As Object 'raw char values from serial port'
        Dim string_result As String = "" 'coverted into a string
        Dim valid_reading As Boolean 'whether the reading is acceptable
        ' Dim averaged_out As Boolean 'set to true after getting 5 readings that are close in range to each other.
        Dim string_iterator As IEnumerator = string_result.GetEnumerator
        Dim digit_count As Integer = 0
        Dim string_store As String

        'If DoRead = False And ContinuousRead = False Then
        '    Exit Sub
        'End If
        'If Not DoRead Then Exit Sub
        'If ValueUpdated Then Exit Sub
        'If retrieving Then Exit Sub
        'retrieving = True
        Dim comStatus As Short = 0
        Try
            With AxMSComm1
                comStatus = .CommEvent
                Select Case comStatus
                    Case 2
                        .RThreshold = 0 'prevent generation of events
                        'If Not DoRead Then
                        '    raw_result = AxMSComm1.Input
                        '    '.InBufferCount = 0
                        '    .RThreshold = 12
                        '    Exit Sub
                        'End If
                        raw_result = AxMSComm1.Input
                        string_result = CType(raw_result, String)
                        string_iterator = string_result.GetEnumerator
                        digit_count = 0
                        string_store = ""
                        valid_reading = False
                        'If needReturn Then
                        '    If string_result.Length = 12 Then
                        '        'Console.WriteLine(DateTime.Now.ToLongTimeString() + " command returned.")
                        '        afterCommand = True

                        '    End If
                        '    AxMSComm1.RThreshold = 12
                        '    needReturn = False
                        '    Exit Sub
                        'End If
                        'this is to read sequences of consecutive 4-5 digits to get the proper reading
                        'If afterCommand Then
                        'afterCommand = False
                        ' Console.WriteLine(DateTime.Now.ToLongTimeString() + " Received data after command")
                        'End If
                        While string_iterator.MoveNext()
                            If IsDigit(CChar(string_iterator.Current)) Then
                                string_store = string_store + CStr(string_iterator.Current)
                                digit_count += 1
                            Else
                                'if there there are spaces in betwen then use as a valid reading.
                                If digit_count = 4 Then
                                    ls_Reading = ConvertReadingToMM(string_store)
                                    string_result = FormatNumber(LASER_Reading, 3)
                                    If readCount >= 10 Then
                                        VUpdated = True
                                        DoRead = False
                                        'Console.WriteLine("Laser Updated #" & readCount)
                                        '.InBufferCount = 0
                                        readCount = 0
                                        'Display.Text = ""
                                        Exit While

                                    Else
                                        readCount += 1
                                        If ContinuousRead Then
                                            Display.AppendText(DateTime.Now.ToLongTimeString() & "Ls Read: " & ls_Reading & vbCrLf)
                                            'Display.Text += DateTime.Now.ToLongTimeString() & "Ls Read: " & ls_Reading & vbCrLf
                                        End If

                                        'Console.WriteLine("Laser #" & MM_Reading)
                                    End If
                                    'Console.WriteLine("Laser OnCom reading: " & MM_Reading)

                                    'Exit While
                                Else
                                    VUpdated = False
                                    string_store = ""
                                    digit_count = 0
                                End If
                            End If
                        End While

                        'If ValueUpdated = True Then
                        'If string_store.Length <> 4 Then
                        '    Status.Text = "OK"
                        'End If
                        'Status.Text = "OK"
                        'apply formula and format to 3 decimal places
                        'MM_Reading = ConvertReadingToMM(string_store)
                        'lbReading.Text = FormatNumber(MM_Reading, 3)
                        'string_result = FormatNumber(MM_Reading, 3)

                        'everything below is for interface window only.
                        'If MM_Reading > CDbl(Highest.Text) Then
                        '    Highest.Text = MM_Reading.ToString
                        'End If
                        'If MM_Reading < CDbl(Lowest.Text) Then Lowest.Text = MM_Reading.ToString
                        'If last_line = string_result Then
                        'Else
                        'display only up to 40 lines at any time


                        'last_line = string_result
                        'raw_result_store = raw_result_store + "Reading: " + string_store + " at: " + DateTime.Now.ToShortTimeString + vbCrLf
                        'newline_count += 1
                        'If newline_count > 30 Then
                        '    raw_result_store = ""
                        '    newline_count = 0
                        'End If


                        'End If
                        'Else
                        'Status.Text = "Bad Reading: " + string_result
                        'End If

                        .RThreshold = 12
                        'Display.Text = raw_result_store
                End Select
            End With
            'retrieving = False
        Catch ex As Exception

        End Try
    End Sub

    Sub ILD1402CommandFormat(ByVal command As String)
        '+ + + CR I L D 1
        Dim output_str As String = Chr(&H2B) + Chr(&H2B) + Chr(&H2B) + Chr(&HD) + Chr(&H49) + Chr(&H4C) + Chr(&H44) + Chr(&H31)

        Select Case command
            Case "Read"
                output_str = output_str + Chr(&H20) + Chr(&H77) + Chr(&H0) + Chr(&H2)
        End Select

        Try
            AxMSComm1.Output = output_str
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If WaitForReadingToStabilize() Then
            'Display.Text += "Laser Reading:" + Me.ls_Reading.ToString("0.00") + " mm." + vbCrLf
            Display.AppendText("Laser Reading:" + Me.ls_Reading.ToString("0.00") + " mm." + vbCrLf)
        Else
            'Display.Text += "Laser Reading failed." + vbCrLf
            Display.AppendText("Laser Reading failed." + vbCrLf)
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ContinuousReadButton.Click
        If ContinuousReadButton.Text = "Enable Continuous Read" Then
            EnableContinuousRead()
            ContinuousReadButton.Text = "Disable Continuous Read"
        ElseIf ContinuousReadButton.Text = "Disable Continuous Read" Then
            DisableContinuousRead()
            ContinuousReadButton.Text = "Enable Continuous Read"
        End If
    End Sub

    Private Sub Laser_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        DisableContinuousRead()
        ContinuousReadButton.Text = "Enable Continuous Read"
        e.Cancel = True
        Hide()
    End Sub

    Private Function WaitForReply(ByVal timeout As Double) As Boolean
        Dim timeStart As Long = DateTime.Now.Ticks
        Do
            Application.DoEvents()
        Loop Until Not needReturn Or (((DateTime.Now.Ticks - timeStart) / 10000) > timeout)
        If needReturn Then
            Return False
        End If
        Return True
    End Function

    Private Sub btTurnOffMeasurement_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btTurnOffMeasurement.Click
        Me.SendCommand("TurnOffMeasurement")
    End Sub

    Private Sub btTurnOnMeasurement_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btTurnOnMeasurement.Click
        Me.SendCommand("TurnOnMeasurement")
    End Sub
#End Region

#Region "Public Implements ILaser Interface"
    Public Property LASER_Reading() As Double Implements ILaser.LASER_Reading
        Get
            Return ls_Reading
        End Get
        Set(ByVal Value As Double)
            ls_Reading = Value
        End Set
    End Property

    Public Property ValueUpdated() As Boolean Implements ILaser.ValueUpdated
        Get
            Return Me.VUpdated
        End Get
        Set(ByVal Value As Boolean)
            VUpdated = Value
        End Set
    End Property

    Public Function SendCommand(ByVal cmd As String) As Boolean Implements ILaser.SendCommand
        needReturn = False
        OpenPort()
        Dim output_str As String
        Dim header As String = Chr(&H2B) & Chr(&H2B) & Chr(&H2B) & Chr(&HD)
        Dim subHeader1 As String = Chr(&H49) & Chr(&H4C) & Chr(&H44) & Chr(&H31)
        Dim subHeader2 As String = Chr(&H20) & Chr(&H90) & Chr(&H0) & Chr(&H3)
        Dim command As String = Chr(&H20) & Chr(&H76) & Chr(&H0) & Chr(&H2)
        If cmd = "TurnOnMeasurement" Then
            command = Chr(&H20) & Chr(&H77) & Chr(&H0) & Chr(&H2)
            output_str = header + subHeader1 + command
            'Console.WriteLine("Turn On Laser")
        ElseIf cmd = "TurnOffMeasurement" Then
            command = Chr(&H20) & Chr(&H76) & Chr(&H0) & Chr(&H2)
            output_str = header + subHeader1 + command
            'Console.WriteLine("Turn Off Laser")
        Else : Return False
        End If
        Try
            DoRead = True
            needReturn = True
            afterCommand = False
            AxMSComm1.RThreshold = 12
            AxMSComm1.Output = output_str
        Catch ex As Exception
            MsgBox(ex.Message)
            Return False
        End Try
        'WaitForReply(1000)
        Return True
    End Function

    Public Sub EnableContinuousRead() Implements ILaser.EnableContinuousRead
        ContinuousRead = True
    End Sub

    Public Sub DisableContinuousRead() Implements ILaser.DisableContinuousRead
        ContinuousRead = False
    End Sub

    Public Sub ClearComBuffer() Implements ILaser.ClearComBuffer
        AxMSComm1.InBufferCount = 0
    End Sub

    Public Function WaitForReadingToStabilize() As Boolean Implements ILaser.WaitForReadingToStabilize
        readCount = 0
        VUpdated = False
        DoRead = True
        'Sleep(250)
        start_time = Now
        Do
            Sleep(25)
            Application.DoEvents()
            stop_time = Now
            elapsed_time = stop_time.Subtract(start_time)
        Loop Until VUpdated Or elapsed_time.TotalSeconds > 10
        If elapsed_time.TotalSeconds > 10 Then
            Return False
        End If
        'start_time = Now
        'Do
        '    Application.DoEvents()
        '    stop_time = Now
        '    elapsed_time = stop_time.Subtract(start_time)
        'Loop Until elapsed_time.TotalSeconds > 0.5
        DoRead = False
        Return True
    End Function

    Public Function OpenPort() As Boolean Implements ILaser.OpenPort
        If AxMSComm1.PortOpen = True Then
            Exit Function
        End If
        With AxMSComm1
            .CommPort = LaserPort
            .Settings = "115200,N,8,1"
            .InBufferCount = 0
            .InputLen = 12
            .RThreshold = 12
            .SThreshold = 1
            .InputMode = MSCommLib.InputModeConstants.comInputModeText
        End With
        Try
            If Not AxMSComm1.PortOpen Then
                AxMSComm1.PortOpen = True
            End If
        Catch ex As System.Runtime.InteropServices.COMException
            Return False
        End Try
        VUpdated = False
        Return True
    End Function

    Public Function ClosePort() As Boolean Implements ILaser.ClosePort
        If AxMSComm1.PortOpen = True Then
            AxMSComm1.PortOpen = False
        End If
        Return True
    End Function

    Public Function ShowForm() Implements ILaser.ShowForm
        Me.Show()
    End Function
    Public Function HideForm() Implements ILaser.HideForm
        Me.Hide()
    End Function
#End Region
End Class
