Public Interface IPLC
    Sub ShowForm()
    Sub HideForm()
    Function ClosePort() As Boolean
    Function OpenPort() As Boolean
    Function Command(ByVal cmd As String) As Boolean
    Function MoveTo(ByVal width As Double) As Boolean
    Function GetError() As String
    Sub ActivePositionTimer(ByVal start As Boolean)
    Sub Reset()
    Function SetCommand(ByVal cmd As String, ByVal value As Decimal) As Boolean
    Function ReportWidthAdjustFailure() As String
    Property ConveyorError() As String
    Property MoveToWidth() As Double
    Property WidthPosition() As Double
End Interface
