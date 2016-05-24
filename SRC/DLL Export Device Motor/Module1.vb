Imports Microsoft.DirectX
Imports Microsoft.DirectX.DirectInput
Imports DLL_DataManager

'Public Class CIDSDeviceMotor
'    Inherits CIDSTrioController      ' TrioMotion controller device class

'    Public Sub OtherFunctions()
'    End Sub

'End Class

Public Module Module1

    'Trio Motion activeX instance

End Module

Public Module Jogging
    Class Mouse
        Implements IDisposable

        'open device for aquiring
        Public Sub New(ByVal form As Form)
            _device = New Device(SystemGuid.Mouse)
            _device.SetCooperativeLevel(form, CooperativeLevelFlags.Background Or CooperativeLevelFlags.NonExclusive)
            _device.SetDataFormat(DeviceDataFormat.Mouse)
            Try
                _device.Acquire()
            Catch dex As DirectXException
                '  Console.WriteLine(dex.Message)
            End Try
        End Sub


        'polling mouse's states
        Public Sub Poll()
            Try
                _device.Poll()
                _state = _device.CurrentMouseState
                _buttonBuffer = _state.GetMouseButtons
            Catch generatedExceptionVariable0 As NotAcquiredException
                Try
                    _device.Acquire()
                Catch iex As InputException
                    '     Console.WriteLine(iex.Message)
                End Try
            Catch ex2 As InputException
                'Console.WriteLine(ex2.Message)
            End Try
        End Sub


        'Get mouse's states
        Public ReadOnly Property State() As MouseState
            Get
                Return _state
            End Get
        End Property


        'Get mouse's buttons' states
        Public ReadOnly Property MouseButtons() As Byte()
            Get
                Return _buttonBuffer
            End Get
        End Property


        'Get mouse's x direction moving count
        Public ReadOnly Property MouseX() As Integer
            Get
                Return _state.X
            End Get
        End Property


        'Get mouse's y direction moving count
        Public ReadOnly Property MouseY() As Integer
            Get
                Return _state.Y
            End Get
        End Property


        'Get mouse's z direction moving count
        Public ReadOnly Property MouseZ() As Integer
            Get
                Return _state.Z
            End Get
        End Property

        Public Sub Dispose() Implements IDisposable.Dispose
            Dispose(True)
            GC.SuppressFinalize(Me)
        End Sub

        Protected Overridable Sub Dispose(ByVal disposing As Boolean)
            If Not Me._disposed Then
                If disposing Then
                    If Not (_device Is Nothing) Then
                        _device.Unacquire()
                    End If
                End If
            End If
            _disposed = True
        End Sub

        Private _disposed As Boolean
        Private _device As Device
        Private _state As MouseState
        Private _buttonBuffer As Byte()

    End Class

    Class Keyboard
        Implements IDisposable


        'open keyboard device for polling its states
        Public Sub New(ByVal form As Form)
            '_device = New Device(SystemGuid.Keyboard)
            '_device.SetCooperativeLevel(form, CooperativeLevelFlags.Background Or CooperativeLevelFlags.NonExclusive)
            '_device.SetDataFormat(DeviceDataFormat.Keyboard)
            'Try
            '    _device.Acquire()
            'Catch ex As DirectXException
            '    'Console.WriteLine(ex.Message)
            'End Try
        End Sub


        'polling keyboard's states  
        Public Sub Poll()
            'Try
            '    _device.Poll()
            '    _state = _device.GetCurrentKeyboardState
            'Catch generatedExceptionVariable0 As NotAcquiredException
            '    Try
            '        _device.Acquire()
            '    Catch iex As InputException
            '        '   Console.WriteLine(iex.Message)
            '    End Try
            'Catch ex2 As InputException
            '    'Console.WriteLine(ex2.Message)
            'End Try
        End Sub


        'Get keyboard's states
        Public ReadOnly Property State() As KeyboardState
            Get
                Return _state
            End Get
        End Property

        Public Sub Dispose() Implements IDisposable.Dispose
            Dispose(True)
            GC.SuppressFinalize(Me)
        End Sub

        Overridable Sub Dispose(ByVal disposing As Boolean)
            If Not Me._disposed Then
                If disposing Then
                    If Not (_device Is Nothing) Then
                        _device.Unacquire()
                    End If
                End If
            End If
            _disposed = True
        End Sub

        Private _disposed As Boolean
        Private _device As Device
        Private _state As KeyboardState
    End Class
End Module