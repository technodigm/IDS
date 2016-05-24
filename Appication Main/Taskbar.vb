Public Class Taskbar
    Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Integer, ByVal hWndInsertAfter As Integer, ByVal x As Integer, ByVal y As Integer, ByVal cx As Integer, ByVal cy As Integer, ByVal wFlags As Integer) As Integer
    Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Integer
    Const SWP_HIDEWINDOW = &H80
    Const SWP_SHOWWINDOW = &H40

    Public Shared Sub ShowTaskBar(ByVal toShow As Boolean)

        If toShow Then
            SetWindowPos(FindWindow("Shell_traywnd", ""), 0&, 0&, 0&, 0&, 0&, SWP_SHOWWINDOW)
        Else
            SetWindowPos(FindWindow("Shell_traywnd", ""), 0&, 0&, 0&, 0&, 0&, SWP_HIDEWINDOW)
        End If
    End Sub
End Class
