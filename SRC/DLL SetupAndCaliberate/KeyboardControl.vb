Imports System.Runtime.InteropServices
Imports System.Reflection
Imports System.Reflection.Assembly
Public Class KeyboardControl
    Private Delegate Function HookCallback(ByVal nCode As Integer, ByVal wParam As Integer, ByVal lParam As IntPtr) As Integer
    Private Shared HookDelegate As HookCallback
    Private Shared HookId As Integer
    Private Const Wh_Keyboard_LL As Integer = 13
    Private Const Vk_Tab As Integer = 9
    Private Const Vk_Escape As Integer = 27
    Private Const VK_LWinKey As Integer = 91
    Private Const VK_RWinKey As Integer = 92

    Private Const WM_KEYDOWN As Integer = 256
    Private Const WM_KEYUP As Integer = 257

    Private Const VK_LControl As Integer = 162
    Private Const VK_RControl As Integer = 163

    Public Shared ControlKeyPressed As Boolean = False
    Private Shared Function KeyBoardHookProc(ByVal nCode As Integer, ByVal wParam As Integer, ByVal lParam As IntPtr) As Integer
        'All keyboard events will be sent here.
        'Don't process just pass along.
        If nCode < 0 Then
            Return CallNextHookEx(HookId, nCode, New IntPtr(wParam), lParam)
        End If
        'Extract the keyboard structure from the lparam
        'This will contain the virtual key and any flags.
        'This is using the my.computer.keyboard to get the
        'flags instead
        Dim KeyboardSruct As KBDLLHOOKSTRUCT = Marshal.PtrToStructure(lParam, GetType(KBDLLHOOKSTRUCT))
        'Send the message along  
        If wParam = WM_KEYDOWN And (KeyboardSruct.vkCode = VK_RControl Or KeyboardSruct.vkCode = VK_LControl) Then
            ControlKeyPressed = True
            'Console.WriteLine("Programming Key Down" & Hex(KeyboardSruct.vkCode))
        End If
        If wParam = WM_KEYUP And (KeyboardSruct.vkCode = VK_RControl Or KeyboardSruct.vkCode = VK_LControl) Then
            ControlKeyPressed = False
            'Console.WriteLine("Programming Key Up" & Hex(KeyboardSruct.vkCode))
        End If
        Return CallNextHookEx(HookId, nCode, New IntPtr(wParam), lParam)
    End Function
    Public Shared Sub GainControls()
        'Add the low level keyboard hook
        If HookId = 0 Then
            HookDelegate = AddressOf KeyBoardHookProc
            HookId = SetWindowsHookEx(Wh_Keyboard_LL, HookDelegate, Marshal.GetHINSTANCE(System.Reflection.Assembly.GetExecutingAssembly.GetModules()(0)), 0)
            If HookId = 0 Then
                MessageBox.Show("Error when enable keyboard aceess control")
            End If
        End If
    End Sub
    Public Shared Sub ReleaseControls()
        'Remove the hook
        UnhookWindowsHookEx(HookId)
        HookId = 0
    End Sub
    <DllImport("user32.dll", CharSet:=CharSet.Auto, CallingConvention:=CallingConvention.StdCall)> _
    Private Shared Function CallNextHookEx(ByVal idHook As Integer, ByVal nCode As Integer, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As Integer
    End Function
    <DllImport("user32.dll", CharSet:=CharSet.Auto, CallingConvention:=CallingConvention.StdCall, SetLastError:=True)> _
    Private Shared Function SetWindowsHookEx(ByVal idHook As Integer, ByVal HookProc As HookCallback, ByVal hInstance As IntPtr, ByVal wParam As Integer) As Integer
    End Function
    <DllImport("user32.dll", CharSet:=CharSet.Auto, CallingConvention:=CallingConvention.StdCall, SetLastError:=True)> _
    Private Shared Function UnhookWindowsHookEx(ByVal idHook As Integer) As Integer
    End Function
    Private Structure KBDLLHOOKSTRUCT
        Public vkCode As Integer
        Public scanCode As Integer
        Public flags As Integer
        Public time As Integer
        Public dwExtraInfo As IntPtr
    End Structure
End Class
