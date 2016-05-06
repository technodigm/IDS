Public Class FormArraySetting
    Inherits System.Windows.Forms.Form
    Private m_ArrayData As ArrayData
#Region " Windows Form Designer generated code "

    Public Sub New(ByVal data As ArrayData)
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        m_ArrayData = Data
        m_ArrayData.ColumNo = 2
        m_ArrayData.RowNo = 2
        m_ArrayData.byColum = False
        m_ArrayData.ByZigzag = False
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
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBoxSize As System.Windows.Forms.GroupBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents TextFirstX As System.Windows.Forms.TextBox
    Friend WithEvents TextFirstY As System.Windows.Forms.TextBox
    Friend WithEvents TextDiagY As System.Windows.Forms.TextBox
    Friend WithEvents TextDiagX As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents NumericUpDownRows As System.Windows.Forms.NumericUpDown
    Friend WithEvents NumericUpDownColum As System.Windows.Forms.NumericUpDown
    Friend WithEvents CheckBoxColum As System.Windows.Forms.CheckBox
    Friend WithEvents ImageListYesNo As System.Windows.Forms.ImageList
    Friend WithEvents TBYesNo As System.Windows.Forms.ToolBar
    Friend WithEvents OK As System.Windows.Forms.ToolBarButton
    Friend WithEvents Cancel As System.Windows.Forms.ToolBarButton
    Friend WithEvents CheckBoxZigzag As System.Windows.Forms.CheckBox
    Friend WithEvents LabelMessage As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(FormArraySetting))
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.TextDiagY = New System.Windows.Forms.TextBox
        Me.TextDiagX = New System.Windows.Forms.TextBox
        Me.TextFirstY = New System.Windows.Forms.TextBox
        Me.TextFirstX = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.GroupBoxSize = New System.Windows.Forms.GroupBox
        Me.CheckBoxZigzag = New System.Windows.Forms.CheckBox
        Me.CheckBoxColum = New System.Windows.Forms.CheckBox
        Me.NumericUpDownColum = New System.Windows.Forms.NumericUpDown
        Me.NumericUpDownRows = New System.Windows.Forms.NumericUpDown
        Me.Label7 = New System.Windows.Forms.Label
        Me.Label6 = New System.Windows.Forms.Label
        Me.LabelMessage = New System.Windows.Forms.Label
        Me.ImageListYesNo = New System.Windows.Forms.ImageList(Me.components)
        Me.TBYesNo = New System.Windows.Forms.ToolBar
        Me.OK = New System.Windows.Forms.ToolBarButton
        Me.Cancel = New System.Windows.Forms.ToolBarButton
        Me.GroupBox1.SuspendLayout()
        Me.GroupBoxSize.SuspendLayout()
        CType(Me.NumericUpDownColum, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.NumericUpDownRows, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.TextDiagY)
        Me.GroupBox1.Controls.Add(Me.TextDiagX)
        Me.GroupBox1.Controls.Add(Me.TextFirstY)
        Me.GroupBox1.Controls.Add(Me.TextFirstX)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Location = New System.Drawing.Point(24, 184)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(368, 32)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Position"
        Me.GroupBox1.Visible = False
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(256, 16)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(16, 16)
        Me.Label4.TabIndex = 7
        Me.Label4.Text = "Y"
        Me.Label4.Visible = False
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(152, 16)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(16, 16)
        Me.Label3.TabIndex = 6
        Me.Label3.Text = "X"
        Me.Label3.Visible = False
        '
        'TextDiagY
        '
        Me.TextDiagY.Location = New System.Drawing.Point(224, 72)
        Me.TextDiagY.Name = "TextDiagY"
        Me.TextDiagY.Size = New System.Drawing.Size(72, 20)
        Me.TextDiagY.TabIndex = 5
        Me.TextDiagY.Text = "0"
        Me.TextDiagY.Visible = False
        '
        'TextDiagX
        '
        Me.TextDiagX.Location = New System.Drawing.Point(120, 72)
        Me.TextDiagX.Name = "TextDiagX"
        Me.TextDiagX.Size = New System.Drawing.Size(72, 20)
        Me.TextDiagX.TabIndex = 4
        Me.TextDiagX.Text = "0"
        Me.TextDiagX.Visible = False
        '
        'TextFirstY
        '
        Me.TextFirstY.Enabled = False
        Me.TextFirstY.Location = New System.Drawing.Point(224, 32)
        Me.TextFirstY.Name = "TextFirstY"
        Me.TextFirstY.Size = New System.Drawing.Size(72, 20)
        Me.TextFirstY.TabIndex = 3
        Me.TextFirstY.Text = "0"
        Me.TextFirstY.Visible = False
        '
        'TextFirstX
        '
        Me.TextFirstX.Enabled = False
        Me.TextFirstX.Location = New System.Drawing.Point(120, 32)
        Me.TextFirstX.Name = "TextFirstX"
        Me.TextFirstX.Size = New System.Drawing.Size(72, 20)
        Me.TextFirstX.TabIndex = 2
        Me.TextFirstX.Text = "0"
        Me.TextFirstX.Visible = False
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(32, 72)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(88, 16)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "Diagonal Point"
        Me.Label2.Visible = False
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(40, 40)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(56, 16)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "First Point"
        Me.Label1.Visible = False
        '
        'GroupBoxSize
        '
        Me.GroupBoxSize.Controls.Add(Me.CheckBoxZigzag)
        Me.GroupBoxSize.Controls.Add(Me.CheckBoxColum)
        Me.GroupBoxSize.Controls.Add(Me.NumericUpDownColum)
        Me.GroupBoxSize.Controls.Add(Me.NumericUpDownRows)
        Me.GroupBoxSize.Controls.Add(Me.Label7)
        Me.GroupBoxSize.Controls.Add(Me.Label6)
        Me.GroupBoxSize.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBoxSize.Location = New System.Drawing.Point(16, 8)
        Me.GroupBoxSize.Name = "GroupBoxSize"
        Me.GroupBoxSize.Size = New System.Drawing.Size(376, 96)
        Me.GroupBoxSize.TabIndex = 1
        Me.GroupBoxSize.TabStop = False
        Me.GroupBoxSize.Text = "Size / Mode"
        '
        'CheckBoxZigzag
        '
        Me.CheckBoxZigzag.Location = New System.Drawing.Point(208, 96)
        Me.CheckBoxZigzag.Name = "CheckBoxZigzag"
        Me.CheckBoxZigzag.Size = New System.Drawing.Size(104, 28)
        Me.CheckBoxZigzag.TabIndex = 6
        Me.CheckBoxZigzag.Text = "Zigzag ON"
        Me.CheckBoxZigzag.Visible = False
        '
        'CheckBoxColum
        '
        Me.CheckBoxColum.Location = New System.Drawing.Point(48, 96)
        Me.CheckBoxColum.Name = "CheckBoxColum"
        Me.CheckBoxColum.Size = New System.Drawing.Size(104, 28)
        Me.CheckBoxColum.TabIndex = 5
        Me.CheckBoxColum.Text = "Dispense colum by colum"
        Me.CheckBoxColum.Visible = False
        '
        'NumericUpDownColum
        '
        Me.NumericUpDownColum.Location = New System.Drawing.Point(304, 48)
        Me.NumericUpDownColum.Minimum = New Decimal(New Integer() {2, 0, 0, 0})
        Me.NumericUpDownColum.Name = "NumericUpDownColum"
        Me.NumericUpDownColum.Size = New System.Drawing.Size(56, 26)
        Me.NumericUpDownColum.TabIndex = 4
        Me.NumericUpDownColum.Value = New Decimal(New Integer() {2, 0, 0, 0})
        '
        'NumericUpDownRows
        '
        Me.NumericUpDownRows.Location = New System.Drawing.Point(104, 48)
        Me.NumericUpDownRows.Minimum = New Decimal(New Integer() {2, 0, 0, 0})
        Me.NumericUpDownRows.Name = "NumericUpDownRows"
        Me.NumericUpDownRows.Size = New System.Drawing.Size(56, 26)
        Me.NumericUpDownRows.TabIndex = 3
        Me.NumericUpDownRows.Value = New Decimal(New Integer() {2, 0, 0, 0})
        '
        'Label7
        '
        Me.Label7.Location = New System.Drawing.Point(184, 48)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(120, 16)
        Me.Label7.TabIndex = 2
        Me.Label7.Text = "No. of columns"
        '
        'Label6
        '
        Me.Label6.Location = New System.Drawing.Point(8, 48)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(88, 16)
        Me.Label6.TabIndex = 1
        Me.Label6.Text = "No. of rows"
        '
        'LabelMessage
        '
        Me.LabelMessage.BackColor = System.Drawing.Color.PaleGoldenrod
        Me.LabelMessage.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.LabelMessage.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelMessage.ForeColor = System.Drawing.Color.Blue
        Me.LabelMessage.Location = New System.Drawing.Point(16, 128)
        Me.LabelMessage.Name = "LabelMessage"
        Me.LabelMessage.Size = New System.Drawing.Size(296, 42)
        Me.LabelMessage.TabIndex = 2
        '
        'ImageListYesNo
        '
        Me.ImageListYesNo.ImageSize = New System.Drawing.Size(30, 30)
        Me.ImageListYesNo.ImageStream = CType(resources.GetObject("ImageListYesNo.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.ImageListYesNo.TransparentColor = System.Drawing.Color.Transparent
        '
        'TBYesNo
        '
        Me.TBYesNo.BackColor = System.Drawing.Color.FromArgb(CType(192, Byte), CType(255, Byte), CType(255, Byte))
        Me.TBYesNo.Buttons.AddRange(New System.Windows.Forms.ToolBarButton() {Me.OK, Me.Cancel})
        Me.TBYesNo.ButtonSize = New System.Drawing.Size(42, 42)
        Me.TBYesNo.Divider = False
        Me.TBYesNo.Dock = System.Windows.Forms.DockStyle.None
        Me.TBYesNo.DropDownArrows = True
        Me.TBYesNo.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.TBYesNo.ImageList = Me.ImageListYesNo
        Me.TBYesNo.Location = New System.Drawing.Point(320, 128)
        Me.TBYesNo.Name = "TBYesNo"
        Me.TBYesNo.ShowToolTips = True
        Me.TBYesNo.Size = New System.Drawing.Size(84, 46)
        Me.TBYesNo.TabIndex = 3
        '
        'OK
        '
        Me.OK.ImageIndex = 0
        Me.OK.ToolTipText = "OK"
        '
        'Cancel
        '
        Me.Cancel.ImageIndex = 1
        Me.Cancel.ToolTipText = "Cancel"
        '
        'FormArraySetting
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(408, 182)
        Me.Controls.Add(Me.TBYesNo)
        Me.Controls.Add(Me.LabelMessage)
        Me.Controls.Add(Me.GroupBoxSize)
        Me.Controls.Add(Me.GroupBox1)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "FormArraySetting"
        Me.Text = "Array Setup"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBoxSize.ResumeLayout(False)
        CType(Me.NumericUpDownColum, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.NumericUpDownRows, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

    Public Sub GetArrayData(ByRef data As ArrayData)
        data = m_ArrayData
    End Sub

    Private Sub TBYesNo_ButtonClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ToolBarButtonClickEventArgs) Handles TBYesNo.ButtonClick
        If e.Button.ToolTipText = "OK" Then
            Dim value As Integer
            Try
                value = System.Int32.Parse(NumericUpDownRows.Text)
                If value < 2 Or value > 100 Then
                    Me.LabelMessage.Text = "No. of rows must be between 2 to 100"
                    Me.LabelMessage.Update()
                    Exit Sub
                End If
                m_ArrayData.RowNo = value
                value = System.Int32.Parse(NumericUpDownColum.Text)
                If value < 2 Or value > 100 Then
                    Me.LabelMessage.Text = "No. of colums must be between 2 to 100"
                    Me.LabelMessage.Update()
                    Exit Sub
                End If
                m_ArrayData.ColumNo = value

            Catch ex As Exception
            End Try
            Me.DialogResult = DialogResult.OK
        Else
            Me.DialogResult = DialogResult.Cancel
        End If
        Me.Close()
    End Sub
    Private Sub TextFirstX_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextFirstX.TextChanged
        Try
            m_ArrayData.FirstX = System.Double.Parse(TextFirstX.Text)

        Catch ex As Exception
            TextFirstX.Text = m_ArrayData.FirstX.ToString
            TextFirstX.Update()
        End Try
    End Sub


    Private Sub TextFirstY_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextFirstY.TextChanged
        Try
            m_ArrayData.FirstY = System.Double.Parse(TextFirstY.Text)

        Catch ex As Exception
            TextFirstY.Text = m_ArrayData.FirstY.ToString
            TextFirstY.Update()
        End Try
    End Sub

    Private Sub TextDiagX_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextDiagX.TextChanged
        Try
            m_ArrayData.DiagX = System.Double.Parse(TextDiagX.Text)

        Catch ex As Exception
            TextDiagX.Text = m_ArrayData.DiagX.ToString
            TextDiagX.Update()
        End Try
    End Sub

    Private Sub TextDiagY_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextDiagY.TextChanged
        Try
            m_ArrayData.DiagY = System.Double.Parse(TextDiagY.Text)

        Catch ex As Exception
            TextDiagY.Text = m_ArrayData.DiagY.ToString
            TextDiagY.Update()
        End Try
    End Sub


    Private Sub FormArraySetting_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.TextFirstX.Text = m_ArrayData.FirstX.ToString
        Me.TextFirstY.Text = m_ArrayData.FirstY.ToString
    End Sub


    Private Sub CheckBoxColum_CheckStateChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles CheckBoxColum.CheckStateChanged
        m_ArrayData.byColum = Not m_ArrayData.byColum
    End Sub

    Private Sub CheckBoxZigzag_CheckStateChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles CheckBoxZigzag.CheckStateChanged
        m_ArrayData.ByZigzag = Not m_ArrayData.ByZigzag
    End Sub


End Class
