Imports DLL_Export_Service
Imports DLL_DataManager

Public Class HeaterSettings
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
    Friend WithEvents ButtonRevert As System.Windows.Forms.Button
    Friend WithEvents PanelToBeAdded As System.Windows.Forms.Panel
    Friend WithEvents Label17 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents ThermalPSReadTemp As System.Windows.Forms.Timer
    Friend WithEvents ButtonSave As System.Windows.Forms.Button
    Friend WithEvents ButtonExit As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(HeaterSettings))
        Me.ButtonSave = New System.Windows.Forms.Button
        Me.ButtonRevert = New System.Windows.Forms.Button
        Me.PanelToBeAdded = New System.Windows.Forms.Panel
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label17 = New System.Windows.Forms.Label
        Me.ButtonExit = New System.Windows.Forms.Button
        Me.ThermalPSReadTemp = New System.Windows.Forms.Timer(Me.components)
        Me.PanelToBeAdded.SuspendLayout()
        Me.SuspendLayout()
        '
        'ButtonSave
        '
        Me.ButtonSave.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonSave.Location = New System.Drawing.Point(312, 704)
        Me.ButtonSave.Name = "ButtonSave"
        Me.ButtonSave.Size = New System.Drawing.Size(75, 40)
        Me.ButtonSave.TabIndex = 16
        Me.ButtonSave.Text = "Save"
        '
        'ButtonRevert
        '
        Me.ButtonRevert.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonRevert.Location = New System.Drawing.Point(408, 704)
        Me.ButtonRevert.Name = "ButtonRevert"
        Me.ButtonRevert.Size = New System.Drawing.Size(75, 40)
        Me.ButtonRevert.TabIndex = 15
        Me.ButtonRevert.Text = "Revert"
        Me.ButtonRevert.Visible = False
        '
        'PanelToBeAdded
        '
        Me.PanelToBeAdded.BackColor = System.Drawing.SystemColors.Control
        Me.PanelToBeAdded.Controls.Add(Me.Label1)
        Me.PanelToBeAdded.Controls.Add(Me.Label17)
        Me.PanelToBeAdded.Controls.Add(Me.ButtonSave)
        Me.PanelToBeAdded.Controls.Add(Me.ButtonRevert)
        Me.PanelToBeAdded.Controls.Add(Me.ButtonExit)
        Me.PanelToBeAdded.Location = New System.Drawing.Point(16, 16)
        Me.PanelToBeAdded.Name = "PanelToBeAdded"
        Me.PanelToBeAdded.Size = New System.Drawing.Size(512, 944)
        Me.PanelToBeAdded.TabIndex = 24
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(40, 48)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(320, 32)
        Me.Label1.TabIndex = 35
        Me.Label1.Text = "Temperature Units: Degree Celcius"
        '
        'Label17
        '
        Me.Label17.Font = New System.Drawing.Font("Microsoft Sans Serif", 16.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label17.ForeColor = System.Drawing.SystemColors.ActiveCaption
        Me.Label17.Location = New System.Drawing.Point(0, 0)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(296, 32)
        Me.Label17.TabIndex = 34
        Me.Label17.Text = "Thermal Controller Setup"
        '
        'ButtonExit
        '
        Me.ButtonExit.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonExit.Image = CType(resources.GetObject("ButtonExit.Image"), System.Drawing.Image)
        Me.ButtonExit.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.ButtonExit.Location = New System.Drawing.Point(392, 24)
        Me.ButtonExit.Name = "ButtonExit"
        Me.ButtonExit.Size = New System.Drawing.Size(75, 50)
        Me.ButtonExit.TabIndex = 47
        Me.ButtonExit.TabStop = False
        Me.ButtonExit.Text = "Exit"
        Me.ButtonExit.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'ThermalPSReadTemp
        '
        Me.ThermalPSReadTemp.Interval = 5000
        '
        'HeaterSettings
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(664, 912)
        Me.Controls.Add(Me.PanelToBeAdded)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Name = "HeaterSettings"
        Me.Text = "ThermalController"
        Me.PanelToBeAdded.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

     


    Dim Heater_Enabled_check As Integer = 0
    Dim Heater_Enabled_check2 As Integer = 0
    Public Heater_Enabled As Boolean() = {False, False, False, False, False}
    Dim AlarmValue As Integer() = {0, 0, 0, 0, 0}
    Dim set_temp As Integer() = {0, 0, 0, 0, 0}
    Dim auto_click As Boolean = False

    Dim Read_Device_Cycle As Integer = 0
    Private Shared check_cycle As Integer = 0

#Region " lsgoh please stop T1, T2 and readTemp timer before operating thermal controller!!!!"

    Dim rowData As DataRow
    Dim I As Integer


    '                                       kr                                       '
    ' we only need one value for threshold. this is because the alarm threshold is
    ' triggered through deviation i.e. a set temperature point of 65 and threshold
    ' of 1, will cause alarm trigger at 63/67


    Public Sub RevertData(Optional ByVal hideexit As Boolean = False)
        ButtonExit.Visible = Not hideexit
        IDS.Data.OpenData()
    End Sub
    Private Function DataIsNumeric(ByRef Threshold As Object, ByRef OperationTemp As Object) As Boolean

        If IsNumeric(Threshold) = False Then
            MessageBox.Show("Threshold value is not a numeric number, value will not be saved")

            Return False
        End If

        If IsNumeric(OperationTemp) = False Then
            MessageBox.Show("Set Temperature value is not a numeric number, value will not be saved")
            Return False
        End If

        Return True
    End Function

    ' save data to DB
    Private Sub ButtonSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonSave.Click
        SaveData()
    End Sub

    Public Sub SaveData()

    End Sub

    Private Sub ButtonRevert_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonRevert.Click
        RevertData()
    End Sub
#End Region

    ''' do the other function here

    ' set temp readings to the heater
    Private Sub RS232SetHeater()
    End Sub

    Dim current As String
    Dim wait_list As New SortedList
    Dim error_type As String
    Dim node_num As String

    Friend Sub Thermal_T1_Tick()

        'Dim temp_alarm As Boolean()

        'If wait_list.Count <> 0 Then
        '    For r As Integer = 0 To wait_list.Count - 1
        '        If wait_list.GetByIndex(r) = 50 Then 'confirm 50 cycles is how many seconds later
        '            wait_list.RemoveAt(r)
        '            Exit Sub 'this is necessary, otherwise if you remove an index and move on to the next one (which will be non existent after the removal, it will throw up a run time error
        '        Else
        '            wait_list.SetByIndex(r, wait_list.GetByIndex(r) + 1) 'key is error type, value is counter no.
        '        End If
        '    Next
        'End If

        'If Form_Service.NoActionToExecute Then
        '    IDS.__ID = "1002200"
        '    Heater.FrmHeater.GetHeaterError(IDS.__ID, Heater_Enabled)
        '    error_type = IDS.__ID
        '    node_num = error_type.Substring(6)

        '    If wait_list.ContainsKey(Convert.ToDouble(error_type)) = True Then 'if error type is already in wait list then just +1 to counter
        '        IDS.__ID = "1002200"
        '        Exit Sub
        '    End If

        '    If Not IDS.__ID = "1002200" Then
        '        form_service.DisplayErrorMessage(IDS.__ID)
        '    End If
        'End If

        'error_type = IDS.__ID
        'node_num = error_type.Substring(6)

        ''90 000 is off heater, 60 000 is ignore

        'If error_type = 1062201 Or error_type = 1062202 Or error_type = 1062203 Or error_type = 1062204 Or error_type = 1062205 Then
        '    wait_list.Add(error_type - 60000, 0) 'the 60k is added when you press the button. won't match the fresh error case when checking otherwise, so remove it.
        '    IDS.__ID = "1002200"
        'End If

        'If error_type = 1092201 Or error_type = 1092202 Or error_type = 1092203 Or error_type = 1092204 Or error_type = 1092205 Or error_type = 1092211 Or error_type = 1092212 Or error_type = 1092213 Or error_type = 1092214 Or error_type = 1092215 Then

        '    wait_list.Add(error_type - 90000, 30) 'when turning off, 20 cycles of delay
        '    IDS.__ID = "1002200"
        '    'if error_type is currently in waitlist then remove it
        'End If

    End Sub

    ' to test communication, open com port, send command for echoback test,
    ' read received bytes and determine if any error message

    Private Sub ButtonExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonExit.Click
        RemovePanel(CurrentControl)
    End Sub
End Class