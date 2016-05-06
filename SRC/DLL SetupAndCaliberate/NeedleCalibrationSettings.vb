Imports System.IO
Imports DLL_Export_Service

Public Class NeedleCalibrationSettings
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
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents SaveFileDialog1 As System.Windows.Forms.SaveFileDialog
    Friend WithEvents PanelToBeAdded As System.Windows.Forms.Panel
    Friend WithEvents ButtonExit As System.Windows.Forms.Button
    Friend WithEvents Label21 As System.Windows.Forms.Label
    Friend WithEvents ButtonCalibrate As System.Windows.Forms.Button
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents ZOffset As System.Windows.Forms.Label
    Friend WithEvents YOffset As System.Windows.Forms.Label
    Friend WithEvents XOffset As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Button1 As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(NeedleCalibrationSettings))
        Me.PanelToBeAdded = New System.Windows.Forms.Panel
        Me.Label4 = New System.Windows.Forms.Label
        Me.ZOffset = New System.Windows.Forms.Label
        Me.YOffset = New System.Windows.Forms.Label
        Me.XOffset = New System.Windows.Forms.Label
        Me.Label7 = New System.Windows.Forms.Label
        Me.Label6 = New System.Windows.Forms.Label
        Me.Label21 = New System.Windows.Forms.Label
        Me.ButtonExit = New System.Windows.Forms.Button
        Me.ButtonCalibrate = New System.Windows.Forms.Button
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog
        Me.SaveFileDialog1 = New System.Windows.Forms.SaveFileDialog
        Me.Button1 = New System.Windows.Forms.Button
        Me.PanelToBeAdded.SuspendLayout()
        Me.SuspendLayout()
        '
        'PanelToBeAdded
        '
        Me.PanelToBeAdded.BackColor = System.Drawing.SystemColors.Control
        Me.PanelToBeAdded.Controls.Add(Me.Label4)
        Me.PanelToBeAdded.Controls.Add(Me.ZOffset)
        Me.PanelToBeAdded.Controls.Add(Me.YOffset)
        Me.PanelToBeAdded.Controls.Add(Me.XOffset)
        Me.PanelToBeAdded.Controls.Add(Me.Label7)
        Me.PanelToBeAdded.Controls.Add(Me.Label6)
        Me.PanelToBeAdded.Controls.Add(Me.Label21)
        Me.PanelToBeAdded.Controls.Add(Me.ButtonExit)
        Me.PanelToBeAdded.Controls.Add(Me.ButtonCalibrate)
        Me.PanelToBeAdded.Controls.Add(Me.Button1)
        Me.PanelToBeAdded.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.PanelToBeAdded.Location = New System.Drawing.Point(24, 16)
        Me.PanelToBeAdded.Name = "PanelToBeAdded"
        Me.PanelToBeAdded.Size = New System.Drawing.Size(512, 911)
        Me.PanelToBeAdded.TabIndex = 68
        '
        'Label4
        '
        Me.Label4.ImageAlign = System.Drawing.ContentAlignment.TopLeft
        Me.Label4.Location = New System.Drawing.Point(128, 385)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(136, 25)
        Me.Label4.TabIndex = 91
        Me.Label4.Text = "X offset"
        '
        'ZOffset
        '
        Me.ZOffset.Location = New System.Drawing.Point(296, 446)
        Me.ZOffset.Name = "ZOffset"
        Me.ZOffset.Size = New System.Drawing.Size(88, 24)
        Me.ZOffset.TabIndex = 96
        Me.ZOffset.Text = "0.000"
        '
        'YOffset
        '
        Me.YOffset.Location = New System.Drawing.Point(296, 414)
        Me.YOffset.Name = "YOffset"
        Me.YOffset.Size = New System.Drawing.Size(88, 24)
        Me.YOffset.TabIndex = 95
        Me.YOffset.Text = "0.000"
        '
        'XOffset
        '
        Me.XOffset.Location = New System.Drawing.Point(296, 385)
        Me.XOffset.Name = "XOffset"
        Me.XOffset.Size = New System.Drawing.Size(88, 24)
        Me.XOffset.TabIndex = 94
        Me.XOffset.Text = "0.000"
        '
        'Label7
        '
        Me.Label7.ImageAlign = System.Drawing.ContentAlignment.TopLeft
        Me.Label7.Location = New System.Drawing.Point(128, 446)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(160, 26)
        Me.Label7.TabIndex = 93
        Me.Label7.Text = "Z offset"
        '
        'Label6
        '
        Me.Label6.ImageAlign = System.Drawing.ContentAlignment.TopLeft
        Me.Label6.Location = New System.Drawing.Point(128, 414)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(136, 30)
        Me.Label6.TabIndex = 92
        Me.Label6.Text = "Y offset"
        '
        'Label21
        '
        Me.Label21.Font = New System.Drawing.Font("Microsoft Sans Serif", 16.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label21.ForeColor = System.Drawing.SystemColors.ActiveCaption
        Me.Label21.Location = New System.Drawing.Point(0, 0)
        Me.Label21.Name = "Label21"
        Me.Label21.Size = New System.Drawing.Size(296, 32)
        Me.Label21.TabIndex = 90
        Me.Label21.Text = "Needle Calibration Settings"
        '
        'ButtonExit
        '
        Me.ButtonExit.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonExit.Image = CType(resources.GetObject("ButtonExit.Image"), System.Drawing.Image)
        Me.ButtonExit.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.ButtonExit.Location = New System.Drawing.Point(432, 8)
        Me.ButtonExit.Name = "ButtonExit"
        Me.ButtonExit.Size = New System.Drawing.Size(75, 50)
        Me.ButtonExit.TabIndex = 73
        Me.ButtonExit.TabStop = False
        Me.ButtonExit.Text = "Exit"
        Me.ButtonExit.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'ButtonCalibrate
        '
        Me.ButtonCalibrate.Location = New System.Drawing.Point(160, 486)
        Me.ButtonCalibrate.Name = "ButtonCalibrate"
        Me.ButtonCalibrate.Size = New System.Drawing.Size(176, 40)
        Me.ButtonCalibrate.TabIndex = 61
        Me.ButtonCalibrate.Text = "Calibrate"
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.DefaultExt = "txt"
        Me.OpenFileDialog1.Filter = "(*.txt)*|"
        '
        'SaveFileDialog1
        '
        Me.SaveFileDialog1.DefaultExt = "txt"
        Me.SaveFileDialog1.Filter = "(*.txt)|*"
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(160, 536)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(176, 48)
        Me.Button1.TabIndex = 61
        Me.Button1.Text = "Move to Calibrator Position"
        '
        'NeedleCalibrationSettings
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(1000, 912)
        Me.Controls.Add(Me.PanelToBeAdded)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Name = "NeedleCalibrationSettings"
        Me.Text = "NeedleCalibration1"
        Me.PanelToBeAdded.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub ButtonExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonExit.Click
        OffLaser()
        RemovePanel(CurrentControl)
        m_Tri.Move_Z(0)
    End Sub

    Private Sub ButtonCalibratePosition_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonCalibrate.Click
        With IDS.Data.Hardware.Needle.Left
            XOffset.Text = .CalibratorPos.X - m_Tri.XPosition()
            YOffset.Text = .CalibratorPos.Y - m_Tri.YPosition()
            ZOffset.Text = .CalibratorPos.Z - m_Tri.ZPosition()
            .NeedleCalibrationPosition.X = .CalibratorPos.X - m_Tri.XPosition()
            .NeedleCalibrationPosition.Y = .CalibratorPos.Y - m_Tri.YPosition()
            .NeedleCalibrationPosition.Z = -(.CalibratorPos.Z - m_Tri.ZPosition())
        End With
        IDS.Data.SaveLocalData()
        IDS.Data.SaveGlobalData()
    End Sub

    Public Sub Revert()
        IDS.Data.OpenData()
        With IDS.Data.Hardware.Needle.Left
            XOffset.Text = .NeedleCalibrationPosition.X
            YOffset.Text = .NeedleCalibrationPosition.Y
            ZOffset.Text = .NeedleCalibrationPosition.Z
        End With
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim pos(1) As Double
        With IDS.Data.Hardware.Needle.Left
            pos(0) = .CalibratorPos.X - .NeedleCalibrationPosition.X
            pos(1) = .CalibratorPos.Y - .NeedleCalibrationPosition.Y
            If Not m_Tri.Move_Z(0) Then Exit Sub
            If Not m_Tri.Move_XY(pos) Then Exit Sub
            If Not m_Tri.Move_Z(.CalibratorPos.Z + .NeedleCalibrationPosition.Z) Then Exit Sub
        End With
    End Sub

End Class
