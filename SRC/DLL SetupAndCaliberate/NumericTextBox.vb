imports System.Globalization

Public Class NumericTextBox
    Inherits TextBox

    Private SpaceOK As Boolean = False
    Private DecimalOK As Boolean = True
    Private NumDecimalPlace As Integer = 2



    ' Restricts the entry of characters to digits (including hex),
    ' the negative sign, the e decimal point, and editing keystrokes (backspace).
    Protected Overrides Sub OnKeyPress(ByVal e As KeyPressEventArgs)
        MyBase.OnKeyPress(e)

        Dim numberFormatInfo As NumberFormatInfo = System.Globalization.CultureInfo.CurrentCulture.NumberFormat
        Dim decimalSeparator As String = numberFormatInfo.NumberDecimalSeparator
        Dim groupSeparator As String = numberFormatInfo.NumberGroupSeparator
        Dim negativeSign As String = numberFormatInfo.NegativeSign

        Dim keyInput As String = e.KeyChar.ToString()

        Dim alreadyGotDecimal As Boolean = False

        If Me.Text.IndexOf("."c) >= 0 Then
            alreadyGotDecimal = True
        End If


        If [Char].IsDigit(e.KeyChar) Then
            ' Digits are OK
        ElseIf (keyInput.Equals(decimalSeparator) AndAlso Not alreadyGotDecimal AndAlso DecimalOK) OrElse keyInput.Equals(negativeSign) Then
            ' Decimal separator is OK
        ElseIf e.KeyChar = vbBack Then
            ' Backspace key is OK
            '    else if ((ModifierKeys & (Keys.Control | Keys.Alt)) != 0)
            '    {
            '     // Let the edit control handle control and alt key combinations
            '    }
        ElseIf Me.SpaceOK AndAlso e.KeyChar = " "c Then

        Else
            ' Swallow this invalid key and beep
            e.Handled = True
        End If

    End Sub


    Public ReadOnly Property IntValue() As Integer
        Get
            Return Int32.Parse(Me.Text)
        End Get
    End Property


    Public ReadOnly Property DecimalValue() As Decimal
        Get
            Return [Decimal].Parse(Me.Text)
        End Get
    End Property


    Public Property AllowSpace() As Boolean
        Get
            Return Me.SpaceOK
        End Get
        Set(ByVal value As Boolean)
            Me.SpaceOK = value
        End Set
    End Property

    Public Property AllowDecimal() As Boolean
        Get
            Return Me.DecimalOK
        End Get
        Set(ByVal value As Boolean)
            Me.DecimalOK = value
        End Set
    End Property

    Public Property DecimalPlaces() As Integer
        Get
            Return Me.NumDecimalPlace
        End Get
        Set(ByVal value As Integer)
            Me.NumDecimalPlace = value
        End Set
    End Property


    Protected Overrides Sub OnLeave(ByVal e As System.EventArgs)
        Dim counter As Integer
        If DecimalOK Then                                                   'Textbox allow decimal entry
            If Me.Text.IndexOf("."c) = 0 Then                               'Check if first digit = decimal point
                If Me.TextLength = Me.MaxLength Then                        'Adjust length to allow insertion later
                    Me.MaxLength = Me.MaxLength + 1
                End If
                If Me.TextLength - 1 >= NumDecimalPlace Then                '1st digit = decimal point & no.dp > allowed
                    Me.Text = Me.Text.Substring(0, NumDecimalPlace + 1)     'Extract to the dp allowed
                Else                                                        '1st digit = dp & no.dp < allowed
                    For counter = 1 To NumDecimalPlace - (Me.TextLength - 1) 'Append 0 to get the desired dp
                        Me.AppendText("0")
                    Next
                End If
                Me.Text = Me.Text.Insert(0, "0")                                                        'Handle decimal point display
            ElseIf Me.Text.IndexOf("."c) > 0 Then                                                       'first digit is not a dp
                If Me.TextLength - (Me.Text.IndexOf("."c) + 1) > NumDecimalPlace Then                   'dp > allowed. Delete extra dp
                    Me.Text = Me.Text.Substring(0, (Me.Text.IndexOf("."c) + 1) + NumDecimalPlace)
                ElseIf Me.TextLength - (Me.Text.IndexOf("."c) + 1) < NumDecimalPlace Then               'dp < allowed. Append 0
                    For counter = Me.TextLength - Me.Text.IndexOf("."c) To NumDecimalPlace
                        If Me.TextLength = Me.MaxLength Then
                            Me.MaxLength = Me.MaxLength + 1
                        End If
                        Me.AppendText("0")
                    Next

                End If
            Else                                            'No dp when expected
                Me.AppendText(".")                          'Append dp and 0
                For counter = 1 To NumDecimalPlace
                    Me.AppendText("0")
                Next
            End If
            'check for invalid zeros
            While (Me.Text.Substring(0, 1) = "0" AndAlso Not Me.Text.Substring(1, 1) = ".")
                Me.Text = Me.Text.Substring(1, Me.TextLength - 1)
            End While

        End If
    End Sub
End Class

