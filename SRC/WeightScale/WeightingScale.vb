Public Interface IWeightingScale
   
    Delegate Sub ReadWeightDel(ByVal returnedWeight As Double)
    WriteOnly Property ReadWeightReturned() As ReadWeightDel

    Property IsEnable() As Boolean

    Property RequestWeightUpdate() As Boolean

    Property ValueUpdated() As Boolean

    Property WeightReading() As Double

    Function OpenPort() As Boolean

    Function ClosePort() As Boolean

    Function DoTare() As Boolean

    Function Zero() As Boolean

    Function Restart() As Boolean

    Function ResetValues() As Boolean

    Function GetWeight() As Boolean

    Function SetAmbient(ByVal mode As String)

End Interface
