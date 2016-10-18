Public Interface ILaser
    'Laser reading
    Property LASER_Reading() As Double
    'Flag to check if laser reading is updated
    Property ValueUpdated() As Boolean

    Function ShowForm()
    Function HideForm()
    Function OpenPort() As Boolean
    Function ClosePort() As Boolean
    Sub EnableContinuousRead()
    Sub DisableContinuousRead()
    Sub ClearComBuffer()
    Function WaitForReadingToStabilize() As Boolean
    Function SendCommand(ByVal cmd As String) As Boolean
End Interface
