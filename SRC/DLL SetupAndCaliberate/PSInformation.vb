Imports DLL_Export_Service

Public Class PSInformation
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
    Friend WithEvents PanelToBeAdded As System.Windows.Forms.Panel
    Friend WithEvents Label17 As System.Windows.Forms.Label
    Friend WithEvents ButtonSave As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents ButtonRevert As System.Windows.Forms.Button
    Friend WithEvents TB_Author As System.Windows.Forms.TextBox
    Friend WithEvents TB_Contact As System.Windows.Forms.TextBox
    Friend WithEvents RTB_Infor As System.Windows.Forms.RichTextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.TB_Author = New System.Windows.Forms.TextBox
        Me.TB_Contact = New System.Windows.Forms.TextBox
        Me.RTB_Infor = New System.Windows.Forms.RichTextBox
        Me.PanelToBeAdded = New System.Windows.Forms.Panel
        Me.ButtonRevert = New System.Windows.Forms.Button
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label17 = New System.Windows.Forms.Label
        Me.ButtonSave = New System.Windows.Forms.Button
        Me.PanelToBeAdded.SuspendLayout()
        Me.SuspendLayout()
        '
        'TB_Author
        '
        Me.TB_Author.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.TB_Author.Location = New System.Drawing.Point(32, 152)
        Me.TB_Author.Name = "TB_Author"
        Me.TB_Author.Size = New System.Drawing.Size(448, 27)
        Me.TB_Author.TabIndex = 133
        Me.TB_Author.Text = "Created by Xu Long"
        '
        'TB_Contact
        '
        Me.TB_Contact.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.TB_Contact.Location = New System.Drawing.Point(32, 264)
        Me.TB_Contact.Name = "TB_Contact"
        Me.TB_Contact.Size = New System.Drawing.Size(448, 27)
        Me.TB_Contact.TabIndex = 134
        Me.TB_Contact.Text = "+65-88888888"
        '
        'RTB_Infor
        '
        Me.RTB_Infor.Font = New System.Drawing.Font("Microsoft Sans Serif", 18.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.RTB_Infor.Location = New System.Drawing.Point(32, 384)
        Me.RTB_Infor.Name = "RTB_Infor"
        Me.RTB_Infor.Size = New System.Drawing.Size(448, 304)
        Me.RTB_Infor.TabIndex = 135
        Me.RTB_Infor.Text = "Important message or reminder to be displayed in the production environment"
        '
        'PanelToBeAdded
        '
        Me.PanelToBeAdded.Controls.Add(Me.ButtonRevert)
        Me.PanelToBeAdded.Controls.Add(Me.Label3)
        Me.PanelToBeAdded.Controls.Add(Me.Label2)
        Me.PanelToBeAdded.Controls.Add(Me.Label1)
        Me.PanelToBeAdded.Controls.Add(Me.Label17)
        Me.PanelToBeAdded.Controls.Add(Me.RTB_Infor)
        Me.PanelToBeAdded.Controls.Add(Me.TB_Contact)
        Me.PanelToBeAdded.Controls.Add(Me.TB_Author)
        Me.PanelToBeAdded.Controls.Add(Me.ButtonSave)
        Me.PanelToBeAdded.Location = New System.Drawing.Point(24, 16)
        Me.PanelToBeAdded.Name = "PanelToBeAdded"
        Me.PanelToBeAdded.Size = New System.Drawing.Size(512, 911)
        Me.PanelToBeAdded.TabIndex = 139
        '
        'ButtonRevert
        '
        Me.ButtonRevert.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonRevert.Location = New System.Drawing.Point(376, 784)
        Me.ButtonRevert.Name = "ButtonRevert"
        Me.ButtonRevert.Size = New System.Drawing.Size(88, 48)
        Me.ButtonRevert.TabIndex = 144
        Me.ButtonRevert.Text = "Cancel"
        '
        'Label3
        '
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(8, 344)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(200, 23)
        Me.Label3.TabIndex = 143
        Me.Label3.Text = "Production Information:"
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(8, 224)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(144, 23)
        Me.Label2.TabIndex = 142
        Me.Label2.Text = "Contact Number:"
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(8, 112)
        Me.Label1.Name = "Label1"
        Me.Label1.TabIndex = 141
        Me.Label1.Text = "Author:"
        '
        'Label17
        '
        Me.Label17.Font = New System.Drawing.Font("Microsoft Sans Serif", 16.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label17.ForeColor = System.Drawing.SystemColors.ActiveCaption
        Me.Label17.Location = New System.Drawing.Point(0, 0)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(320, 32)
        Me.Label17.TabIndex = 34
        Me.Label17.Text = "Pattern Program Information"
        '
        'ButtonSave
        '
        Me.ButtonSave.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonSave.Location = New System.Drawing.Point(264, 784)
        Me.ButtonSave.Name = "ButtonSave"
        Me.ButtonSave.Size = New System.Drawing.Size(88, 48)
        Me.ButtonSave.TabIndex = 140
        Me.ButtonSave.Text = "OK"
        '
        'PSInformation
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(576, 961)
        Me.Controls.Add(Me.PanelToBeAdded)
        Me.Name = "PSInformation"
        Me.Text = "Form3"
        Me.PanelToBeAdded.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub ButtonSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonSave.Click
        'xu long 190707 /*

        IDS.Data.Hardware.SPC.ProgAuthorName = TB_Author.Text
        IDS.Data.Hardware.SPC.ProgAuthorContact = TB_Contact.Text
        IDS.Data.Hardware.SPC.ProductionNote = RTB_Infor.Text

        IDS.Data.SaveLocalData()

    End Sub

    Private Sub ButtonRevert_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonRevert.Click
        'IDS.Data.OpenData()

        RevertData()
    End Sub

    Public Sub RevertData()
        TB_Author.Text = IDS.Data.Hardware.SPC.ProgAuthorName
        TB_Contact.Text = IDS.Data.Hardware.SPC.ProgAuthorContact
        RTB_Infor.Text = IDS.Data.Hardware.SPC.ProductionNote
        'xu long 190707 */
    End Sub
End Class
