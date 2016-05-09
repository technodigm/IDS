Imports DLL_Export_Service

Public Class NeedleCalibrationSetup1
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
    Friend WithEvents ButtonExit As System.Windows.Forms.Button
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents PanelToBeAdded As System.Windows.Forms.Panel
    Friend WithEvents ButtonSetCalibratorPosition As System.Windows.Forms.Button
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents ZPosition As System.Windows.Forms.Label
    Friend WithEvents YPosition As System.Windows.Forms.Label
    Friend WithEvents XPosition As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Button1 As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(NeedleCalibrationSetup1))
        Me.PanelToBeAdded = New System.Windows.Forms.Panel
        Me.Label4 = New System.Windows.Forms.Label
        Me.ZPosition = New System.Windows.Forms.Label
        Me.YPosition = New System.Windows.Forms.Label
        Me.XPosition = New System.Windows.Forms.Label
        Me.Label7 = New System.Windows.Forms.Label
        Me.Label6 = New System.Windows.Forms.Label
        Me.ButtonSetCalibratorPosition = New System.Windows.Forms.Button
        Me.Label9 = New System.Windows.Forms.Label
        Me.ButtonExit = New System.Windows.Forms.Button
        Me.Button1 = New System.Windows.Forms.Button
        Me.PanelToBeAdded.SuspendLayout()
        Me.SuspendLayout()
        '
        'PanelToBeAdded
        '
        Me.PanelToBeAdded.Controls.Add(Me.Label4)
        Me.PanelToBeAdded.Controls.Add(Me.ZPosition)
        Me.PanelToBeAdded.Controls.Add(Me.YPosition)
        Me.PanelToBeAdded.Controls.Add(Me.XPosition)
        Me.PanelToBeAdded.Controls.Add(Me.Label7)
        Me.PanelToBeAdded.Controls.Add(Me.Label6)
        Me.PanelToBeAdded.Controls.Add(Me.ButtonSetCalibratorPosition)
        Me.PanelToBeAdded.Controls.Add(Me.Label9)
        Me.PanelToBeAdded.Controls.Add(Me.ButtonExit)
        Me.PanelToBeAdded.Controls.Add(Me.Button1)
        Me.PanelToBeAdded.Location = New System.Drawing.Point(8, 24)
        Me.PanelToBeAdded.Name = "PanelToBeAdded"
        Me.PanelToBeAdded.Size = New System.Drawing.Size(512, 944)
        Me.PanelToBeAdded.TabIndex = 39
        '
        'Label4
        '
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.ImageAlign = System.Drawing.ContentAlignment.TopLeft
        Me.Label4.Location = New System.Drawing.Point(184, 379)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(40, 24)
        Me.Label4.TabIndex = 97
        Me.Label4.Text = "X"
        '
        'ZPosition
        '
        Me.ZPosition.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ZPosition.Location = New System.Drawing.Point(248, 443)
        Me.ZPosition.Name = "ZPosition"
        Me.ZPosition.Size = New System.Drawing.Size(88, 24)
        Me.ZPosition.TabIndex = 102
        Me.ZPosition.Text = "0.000"
        '
        'YPosition
        '
        Me.YPosition.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.YPosition.Location = New System.Drawing.Point(248, 411)
        Me.YPosition.Name = "YPosition"
        Me.YPosition.Size = New System.Drawing.Size(88, 24)
        Me.YPosition.TabIndex = 101
        Me.YPosition.Text = "0.000"
        '
        'XPosition
        '
        Me.XPosition.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.XPosition.Location = New System.Drawing.Point(248, 377)
        Me.XPosition.Name = "XPosition"
        Me.XPosition.Size = New System.Drawing.Size(88, 24)
        Me.XPosition.TabIndex = 100
        Me.XPosition.Text = "0.000"
        '
        'Label7
        '
        Me.Label7.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.ImageAlign = System.Drawing.ContentAlignment.TopLeft
        Me.Label7.Location = New System.Drawing.Point(184, 443)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(40, 24)
        Me.Label7.TabIndex = 99
        Me.Label7.Text = "Z"
        '
        'Label6
        '
        Me.Label6.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.ImageAlign = System.Drawing.ContentAlignment.TopLeft
        Me.Label6.Location = New System.Drawing.Point(184, 411)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(40, 24)
        Me.Label6.TabIndex = 98
        Me.Label6.Text = "Y"
        '
        'ButtonSetCalibratorPosition
        '
        Me.ButtonSetCalibratorPosition.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonSetCalibratorPosition.Location = New System.Drawing.Point(160, 527)
        Me.ButtonSetCalibratorPosition.Name = "ButtonSetCalibratorPosition"
        Me.ButtonSetCalibratorPosition.Size = New System.Drawing.Size(192, 40)
        Me.ButtonSetCalibratorPosition.TabIndex = 62
        Me.ButtonSetCalibratorPosition.Text = "Set Calibrator Position"
        '
        'Label9
        '
        Me.Label9.Font = New System.Drawing.Font("Microsoft Sans Serif", 16.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label9.ForeColor = System.Drawing.SystemColors.ActiveCaption
        Me.Label9.Location = New System.Drawing.Point(0, 0)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(272, 32)
        Me.Label9.TabIndex = 46
        Me.Label9.Text = "Needle Calibration Setup"
        '
        'ButtonExit
        '
        Me.ButtonExit.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonExit.Image = CType(resources.GetObject("ButtonExit.Image"), System.Drawing.Image)
        Me.ButtonExit.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.ButtonExit.Location = New System.Drawing.Point(416, 16)
        Me.ButtonExit.Name = "ButtonExit"
        Me.ButtonExit.Size = New System.Drawing.Size(75, 50)
        Me.ButtonExit.TabIndex = 45
        Me.ButtonExit.TabStop = False
        Me.ButtonExit.Text = "Exit"
        Me.ButtonExit.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'Button1
        '
        Me.Button1.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button1.Location = New System.Drawing.Point(160, 479)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(192, 40)
        Me.Button1.TabIndex = 62
        Me.Button1.Text = "Move to Position"
        '
        'NeedleCalibrationSetup1
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(576, 912)
        Me.Controls.Add(Me.PanelToBeAdded)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Name = "NeedleCalibrationSetup1"
        Me.Text = "NeedleCalibrationSetup1"
        Me.PanelToBeAdded.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub ButtonSetCalibratorPosition_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonSetCalibratorPosition.Click
        XPosition.Text = m_Tri.XPosition()
        YPosition.Text = m_Tri.YPosition()
        ZPosition.Text = m_Tri.ZPosition()
        IDS.Data.Hardware.Needle.Left.CalibratorPos.X = m_Tri.XPosition()
        IDS.Data.Hardware.Needle.Left.CalibratorPos.Y = m_Tri.YPosition()
        IDS.Data.Hardware.Needle.Left.CalibratorPos.Z = m_Tri.ZPosition()
        IDS.Data.SaveLocalData()
    End Sub

    Private Sub ButtonExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonExit.Click
        If Not m_Tri.Move_Z(0) Then Exit Sub
        RemovePanel(CurrentControl)
    End Sub

    Public Sub Revert()
        IDS.Data.OpenData()
        XPosition.Text = IDS.Data.Hardware.Needle.Left.CalibratorPos.X
        YPosition.Text = IDS.Data.Hardware.Needle.Left.CalibratorPos.Y
        ZPosition.Text = IDS.Data.Hardware.Needle.Left.CalibratorPos.Z
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim pos(1) As Double
        pos(0) = IDS.Data.Hardware.Needle.Left.CalibratorPos.X
        pos(1) = IDS.Data.Hardware.Needle.Left.CalibratorPos.Y
        If Not m_Tri.Move_Z(0) Then Exit Sub
        If Not m_Tri.Move_XY(pos) Then Exit Sub
        If Not m_Tri.Move_Z(IDS.Data.Hardware.Needle.Left.CalibratorPos.Z) Then Exit Sub
    End Sub

    Private Sub NeedleCalibrationSetup1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub
End Class
