Option Explicit On 

'Class CIDSElements                                                                              '
'  Despription:                                                                                  '  
'           Dispensing element data handling base class                                          '    
'  Created by:                                                                                   ' 
'           Shen Jian            

Public Class CIDSElements
    Protected Type As String    'element type
    Protected LineNo As Integer 'row number in the sheet
    Protected Sheet As String   'sheet name
    Protected HeightComp As Double 'height compensation value wrt refernce block
    Protected dpara As DispensePara 'dispensing parameters

     
    'Property of dispensing element type                              '
    '                                                                 '
     
    Public Property CmdType() As String
        Get
            Return Type
        End Get
        Set(ByVal Value As String)
            Type = Value
        End Set
    End Property

     
    'Property of dispensing element row number                        '
    '                                                                 '
     
    Public Property CmdLineNo() As Integer
        Get
            Return LineNo
        End Get
        Set(ByVal Value As Integer)
            LineNo = Value
        End Set
    End Property

     
    'Property of sheet name                                           '
    '                                                                 '
     
    Public Property SheetName() As String
        Get
            Return Sheet
        End Get
        Set(ByVal Value As String)
            Sheet = Value
        End Set
    End Property


     
    'Property of height compensation                                  '
    '                                                                 '
     
    Public Property HeightCompS() As Double
        Get
            Return HeightComp
        End Get
        Set(ByVal Value As Double)
            HeightComp = Value
        End Set
    End Property

     
    'Property of dispensing element parameters                        '
    '                                                                 '
     
    Public Property Param() As DispensePara
        Get
            Return dpara
        End Get
        Set(ByVal Value As DispensePara)
            dpara = Value
        End Set
    End Property

End Class

'
'Class CIDSFiducial                                                                              '
'  Despription:                                                                                  '  
'           Fiducial element data handling class                                                 '    
'  Created by:                                                                                   ' 
'           Shen Jian      

Public Class CIDSFiducial
    Inherits CIDSElements

    Protected X1, X2 As Double
    Protected Y1, Y2 As Double
    Protected Z1, Z2 As Double
    Protected NumOfFiducial As Integer
    Protected FidFile1 As String
    Protected FidFile2 As String

    'Property of the first position x
     
    Public Property PosX1() As Double
        Get
            Return X1
        End Get
        Set(ByVal Value As Double)
            X1 = Value
        End Set
    End Property

    'Property of the first position y                                 '
     
    Public Property PosY1() As Double
        Get
            Return Y1
        End Get
        Set(ByVal Value As Double)
            Y1 = Value
        End Set
    End Property

     
    'Property of the first position z                                 '
     
    Public Property PosZ1() As Double
        Get
            Return Z1
        End Get
        Set(ByVal Value As Double)
            Z1 = Value
        End Set
    End Property

     
    'Property of the second position x                                '
     
    Public Property PosX2() As Double
        Get
            Return X2
        End Get
        Set(ByVal Value As Double)
            X2 = Value
        End Set
    End Property

     
    'Property of the second position y                                '
    '                                                                 '
     
    Public Property PosY2() As Double
        Get
            Return Y2
        End Get
        Set(ByVal Value As Double)
            Y2 = Value
        End Set
    End Property

     
    'Property of the second position z                                '
    '                                                                 '
     
    Public Property PosZ2() As Double
        Get
            Return Z2
        End Get
        Set(ByVal Value As Double)
            Z2 = Value
        End Set
    End Property

     
    'Property of number of fiducial                                   '
    '                                                                 '
     
    Public Property NumOfFid() As Integer
        Get
            Return NumOfFiducial
        End Get
        Set(ByVal Value As Integer)
            NumOfFiducial = Value
        End Set
    End Property

     
    'Property of the first fiducial file name                         '
    '                                                                 '
     
    Public Property FirstFid() As String
        Get
            Return FidFile1
        End Get
        Set(ByVal Value As String)
            FidFile1 = Value
        End Set
    End Property

     
    'Property of the second fiducial file name                        '
    '                                                                 '
     
    Public Property SecondFid() As String
        Get
            Return FidFile2
        End Get
        Set(ByVal Value As String)
            FidFile2 = Value
        End Set
    End Property

End Class

'Class CIDSDotArray                                                                             '
'  Despription:                                                                                  '  
'           Dot Array  element data handling class                                                  
'  Created by:                                                                                   ' 
'           Kang Ren          
Public Class CIDSDotArray
    Inherits CIDSElements

    Protected X1, X2, X3 As Double
    Protected Y1, Y2, Y3 As Double
    Protected Z1, Z2, Z3 As Double


    'Property of the first endpoint x                                 '
    '                                                                 '

    Public Property PosX1() As Double
        Get
            Return X1
        End Get
        Set(ByVal Value As Double)
            X1 = Value
        End Set
    End Property


    'Property of the first endpoint y                                 '
    '                                                                 '

    Public Property PosY1() As Double
        Get
            Return Y1
        End Get
        Set(ByVal Value As Double)
            Y1 = Value
        End Set
    End Property


    'Property of the first endpoint z                                 '
    '                                                                 '

    Public Property PosZ1() As Double
        Get
            Return Z1
        End Get
        Set(ByVal Value As Double)
            Z1 = Value
        End Set
    End Property


    'Property of the second endpoint x                                '
    '                                                                 '

    Public Property PosX2() As Double
        Get
            Return X2
        End Get
        Set(ByVal Value As Double)
            X2 = Value
        End Set
    End Property


    'Property of the second endpoint y                                 '
    '                                                                 '

    Public Property PosY2() As Double
        Get
            Return Y2
        End Get
        Set(ByVal Value As Double)
            Y2 = Value
        End Set
    End Property


    'Property of the second endpoint z                                '
    '                                                                 '

    Public Property PosZ2() As Double
        Get
            Return Z2
        End Get
        Set(ByVal Value As Double)
            Z2 = Value
        End Set
    End Property


    'Property of the third endpoint x                                 '
    '                                                                 '

    Public Property PosX3() As Double
        Get
            Return X3
        End Get
        Set(ByVal Value As Double)
            X3 = Value
        End Set
    End Property


    'Property of the third endpoint y                                 '
    '                                                                 '

    Public Property PosY3() As Double
        Get
            Return Y3
        End Get
        Set(ByVal Value As Double)
            Y3 = Value
        End Set
    End Property


    'Property of the third endpoint z                                 '
    '                                                                 '

    Public Property PosZ3() As Double
        Get
            Return Z3
        End Get
        Set(ByVal Value As Double)
            Z3 = Value
        End Set
    End Property

End Class

'Class CIDSDot                                                                                   '
'  Despription:                                                                                  '  
'           Dot element data handling class                                                      '    
'  Created by:                                                                                   ' 
'           Shen Jian 
Public Class CIDSDot
    Inherits CIDSElements

    Protected X As Double
    Protected Y As Double
    Protected Z As Double


    'Property of position x                                           '
    '                                                                 '

    Public Property PosX() As Double
        Get
            Return X
        End Get
        Set(ByVal Value As Double)
            X = Value
        End Set
    End Property


    'Property of position y                                           '
    '                                                                 '

    Public Property PosY() As Double
        Get
            Return Y
        End Get
        Set(ByVal Value As Double)
            Y = Value
        End Set
    End Property


    'Property of position z                                           '
    '                                                                 '

    Public Property PosZ() As Double
        Get
            Return Z
        End Get
        Set(ByVal Value As Double)
            Z = Value
        End Set
    End Property

End Class

'Class CIDSLinkS                                                                                 '
'  Despription:                                                                                  '  
'           Link start element data  class                                                       '    
'  Created by:                                                                                   ' 
'           Shen Jian           
Public Class CIDSLinkS
    Inherits CIDSElements
End Class

'Class CIDSLinkE                                                                                 '
'  Despription:                                                                                  '  
'           Link end element data  class                                                         '    
'  Created by:                                                                                   ' 
'           Shen Jian              
Public Class CIDSLinkE
    Inherits CIDSElements
End Class

'Class CIDSLine                                                                                  '
'  Despription:                                                                                  '  
'           Line element data handling class                                                     '    
'  Created by:                                                                                   ' 
'           Shen Jian             
Public Class CIDSLine
    Inherits CIDSElements

    Protected X1, X2 As Double
    Protected Y1, Y2 As Double
    Protected Z1, Z2 As Double


    'Property of the first endpoint x                                 '
    '                                                                 '

    Public Property PosX1() As Double
        Get
            Return X1
        End Get
        Set(ByVal Value As Double)
            X1 = Value
        End Set
    End Property


    'Property of the first endpoint y                                 '
    '                                                                 '

    Public Property PosY1() As Double
        Get
            Return Y1
        End Get
        Set(ByVal Value As Double)
            Y1 = Value
        End Set
    End Property


    'Property of the first endpoint z                                 '
    '                                                                 '

    Public Property PosZ1() As Double
        Get
            Return Z1
        End Get
        Set(ByVal Value As Double)
            Z1 = Value
        End Set
    End Property


    'Property of the second endpoint x                                '
    '                                                                 '

    Public Property PosX2() As Double
        Get
            Return X2
        End Get
        Set(ByVal Value As Double)
            X2 = Value
        End Set
    End Property


    'Property of the second endpoint y                                '
    '                                                                 '

    Public Property PosY2() As Double
        Get
            Return Y2
        End Get
        Set(ByVal Value As Double)
            Y2 = Value
        End Set
    End Property


    'Property of the second endpoint z                                '
    '                                                                 '

    Public Property PosZ2() As Double
        Get
            Return Z2
        End Get
        Set(ByVal Value As Double)
            Z2 = Value
        End Set
    End Property
End Class


'''''''''''''''''''''''''''''''
'                                                                                                '
'Class CIDSArc                                                                                   '
'  Despription:                                                                                  '  
'           Arc element data handling class                                                      '    
'  Created by:                                                                                   ' 
'           Shen Jian                                                                            '
'                                                                                                '
'''''''''''''''''''''''''''''''
Public Class CIDSArc
    Inherits CIDSElements

    Protected X1, X2, X3 As Double
    Protected Y1, Y2, Y3 As Double
    Protected Z1, Z2, Z3 As Double


    'Property of the first endpoint x                                 '
    '                                                                 '

    Public Property PosX1() As Double
        Get
            Return X1
        End Get
        Set(ByVal Value As Double)
            X1 = Value
        End Set
    End Property


    'Property of the first endpoint y                                 '
    '                                                                 '

    Public Property PosY1() As Double
        Get
            Return Y1
        End Get
        Set(ByVal Value As Double)
            Y1 = Value
        End Set
    End Property


    'Property of the first endpoint z                                 '
    '                                                                 '

    Public Property PosZ1() As Double
        Get
            Return Z1
        End Get
        Set(ByVal Value As Double)
            Z1 = Value
        End Set
    End Property


    'Property of the second endpoint x                                '
    '                                                                 '

    Public Property PosX2() As Double
        Get
            Return X2
        End Get
        Set(ByVal Value As Double)
            X2 = Value
        End Set
    End Property


    'Property of the second endpoint y                                '
    '                                                                 '

    Public Property PosY2() As Double
        Get
            Return Y2
        End Get
        Set(ByVal Value As Double)
            Y2 = Value
        End Set
    End Property


    'Property of the second endpoint z                                '
    '                                                                 '

    Public Property PosZ2() As Double
        Get
            Return Z2
        End Get
        Set(ByVal Value As Double)
            Z2 = Value
        End Set
    End Property


    'Property of the third endpoint x                                 '
    '                                                                 '

    Public Property PosX3() As Double
        Get
            Return X3
        End Get
        Set(ByVal Value As Double)
            X3 = Value
        End Set
    End Property


    'Property of the third endpoint y                                 '
    '                                                                 '

    Public Property PosY3() As Double
        Get
            Return Y3
        End Get
        Set(ByVal Value As Double)
            Y3 = Value
        End Set
    End Property


    'Property of the third endpoint z                                 '
    '                                                                 '

    Public Property PosZ3() As Double
        Get
            Return Z3
        End Get
        Set(ByVal Value As Double)
            Z3 = Value
        End Set
    End Property

End Class


'''''''''''''''''''''''''''''''
'                                                                                                '
'Class CIDSCircle                                                                                '
'  Despription:                                                                                  '  
'           Circle element data handling class                                                   '    
'  Created by:                                                                                   ' 
'           Shen Jian                                                                            '
'                                                                                                '
'''''''''''''''''''''''''''''''
Public Class CIDSCircle
    Inherits CIDSElements

    Protected X1, X2, X3 As Double
    Protected Y1, Y2, Y3 As Double
    Protected Z1, Z2, Z3 As Double


    'Property of the first endpoint x                                 '
    '                                                                 '

    Public Property PosX1() As Double
        Get
            Return X1
        End Get
        Set(ByVal Value As Double)
            X1 = Value
        End Set
    End Property


    'Property of the first endpoint y                                 '
    '                                                                 '

    Public Property PosY1() As Double
        Get
            Return Y1
        End Get
        Set(ByVal Value As Double)
            Y1 = Value
        End Set
    End Property


    'Property of the first endpoint z                                 '
    '                                                                 '

    Public Property PosZ1() As Double
        Get
            Return Z1
        End Get
        Set(ByVal Value As Double)
            Z1 = Value
        End Set
    End Property


    'Property of the second endpoint x                                '
    '                                                                 '

    Public Property PosX2() As Double
        Get
            Return X2
        End Get
        Set(ByVal Value As Double)
            X2 = Value
        End Set
    End Property


    'Property of the second endpoint y                                 '
    '                                                                 '

    Public Property PosY2() As Double
        Get
            Return Y2
        End Get
        Set(ByVal Value As Double)
            Y2 = Value
        End Set
    End Property


    'Property of the second endpoint z                                '
    '                                                                 '

    Public Property PosZ2() As Double
        Get
            Return Z2
        End Get
        Set(ByVal Value As Double)
            Z2 = Value
        End Set
    End Property


    'Property of the third endpoint x                                 '
    '                                                                 '

    Public Property PosX3() As Double
        Get
            Return X3
        End Get
        Set(ByVal Value As Double)
            X3 = Value
        End Set
    End Property


    'Property of the third endpoint y                                 '
    '                                                                 '

    Public Property PosY3() As Double
        Get
            Return Y3
        End Get
        Set(ByVal Value As Double)
            Y3 = Value
        End Set
    End Property


    'Property of the third endpoint z                                 '
    '                                                                 '

    Public Property PosZ3() As Double
        Get
            Return Z3
        End Get
        Set(ByVal Value As Double)
            Z3 = Value
        End Set
    End Property

End Class


'''''''''''''''''''''''''''''''
'                                                                                                '
'Class CIDSFillCircle                                                                            '
'  Despription:                                                                                  '  
'           Fill circle element data handling class                                              '    
'  Created by:                                                                                   ' 
'           Shen Jian                                                                            '
'                                                                                                '
'''''''''''''''''''''''''''''''
Public Class CIDSFillCircle
    Inherits CIDSElements

    Protected X1, X2, X3 As Double
    Protected Y1, Y2, Y3 As Double
    Protected Z1, Z2, Z3 As Double


    'Property of the first endpoint x                                 '
    '                                                                 '

    Public Property PosX1() As Double
        Get
            Return X1
        End Get
        Set(ByVal Value As Double)
            X1 = Value
        End Set
    End Property


    'Property of the first endpoint y                                 '
    '                                                                 '

    Public Property PosY1() As Double
        Get
            Return Y1
        End Get
        Set(ByVal Value As Double)
            Y1 = Value
        End Set
    End Property


    'Property of the first endpoint z                                 '
    '                                                                 '

    Public Property PosZ1() As Double
        Get
            Return Z1
        End Get
        Set(ByVal Value As Double)
            Z1 = Value
        End Set
    End Property


    'Property of the second endpoint x                                '
    '                                                                 '

    Public Property PosX2() As Double
        Get
            Return X2
        End Get
        Set(ByVal Value As Double)
            X2 = Value
        End Set
    End Property


    'Property of the second endpoint y                                '
    '                                                                 '

    Public Property PosY2() As Double
        Get
            Return Y2
        End Get
        Set(ByVal Value As Double)
            Y2 = Value
        End Set
    End Property


    'Property of the second endpoint z                                '
    '                                                                 '

    Public Property PosZ2() As Double
        Get
            Return Z2
        End Get
        Set(ByVal Value As Double)
            Z2 = Value
        End Set
    End Property


    'Property of the third endpoint x                                 '
    '                                                                 '

    Public Property PosX3() As Double
        Get
            Return X3
        End Get
        Set(ByVal Value As Double)
            X3 = Value
        End Set
    End Property


    'Property of the third endpoint y                                 '
    '                                                                 '

    Public Property PosY3() As Double
        Get
            Return Y3
        End Get
        Set(ByVal Value As Double)
            Y3 = Value
        End Set
    End Property


    'Property of the third endpoint z                                 '
    '                                                                 '

    Public Property PosZ3() As Double
        Get
            Return Z3
        End Get
        Set(ByVal Value As Double)
            Z3 = Value
        End Set
    End Property

End Class


'''''''''''''''''''''''''''''''
'                                                                                                '
'Class CIDSRectangle                                                                             '
'  Despription:                                                                                  '  
'           Rectangle element data handling class                                                '    
'  Created by:                                                                                   ' 
'           Shen Jian                                                                            '
'                                                                                                '
'''''''''''''''''''''''''''''''
Public Class CIDSRectangle
    Inherits CIDSElements

    Protected X1, X2, X3 As Double
    Protected Y1, Y2, Y3 As Double
    Protected Z1, Z2, Z3 As Double


    'Property of the first endpoint x                                 '
    '                                                                 '

    Public Property PosX1() As Double
        Get
            Return X1
        End Get
        Set(ByVal Value As Double)
            X1 = Value
        End Set
    End Property


    'Property of the first endpoint y                                 '
    '                                                                 '

    Public Property PosY1() As Double
        Get
            Return Y1
        End Get
        Set(ByVal Value As Double)
            Y1 = Value
        End Set
    End Property


    'Property of the first endpoint z                                 '
    '                                                                 '

    Public Property PosZ1() As Double
        Get
            Return Z1
        End Get
        Set(ByVal Value As Double)
            Z1 = Value
        End Set
    End Property


    'Property of the second endpoint x                                '
    '                                                                 '

    Public Property PosX2() As Double
        Get
            Return X2
        End Get
        Set(ByVal Value As Double)
            X2 = Value
        End Set
    End Property


    'Property of the second endpoint y                                '
    '                                                                 '

    Public Property PosY2() As Double
        Get
            Return Y2
        End Get
        Set(ByVal Value As Double)
            Y2 = Value
        End Set
    End Property


    'Property of the second endpoint z                                 '
    '                                                                 '

    Public Property PosZ2() As Double
        Get
            Return Z2
        End Get
        Set(ByVal Value As Double)
            Z2 = Value
        End Set
    End Property


    'Property of the third endpoint x                                 '
    '                                                                 '

    Public Property PosX3() As Double
        Get
            Return X3
        End Get
        Set(ByVal Value As Double)
            X3 = Value
        End Set
    End Property


    'Property of the third endpoint y                                 '
    '                                                                 '

    Public Property PosY3() As Double
        Get
            Return Y3
        End Get
        Set(ByVal Value As Double)
            Y3 = Value
        End Set
    End Property


    'Property of the third endpoint z                                 '
    '                                                                 '

    Public Property PosZ3() As Double
        Get
            Return Z3
        End Get
        Set(ByVal Value As Double)
            Z3 = Value
        End Set
    End Property

End Class


'''''''''''''''''''''''''''''''
'                                                                                                '
'Class CIDSFillRectangle                                                                         '
'  Despription:                                                                                  '  
'           Fill rectangle element data handling class                                           '    
'  Created by:                                                                                   ' 
'           Shen Jian                                                                            '
'                                                                                                '
'''''''''''''''''''''''''''''''
Public Class CIDSFillRectangle
    Inherits CIDSElements

    Protected X1, X2, X3 As Double
    Protected Y1, Y2, Y3 As Double
    Protected Z1, Z2, Z3 As Double


    'Property of the first endpoint x                                 '
    '                                                                 '

    Public Property PosX1() As Double
        Get
            Return X1
        End Get
        Set(ByVal Value As Double)
            X1 = Value
        End Set
    End Property


    'Property of the first endpoint y                                 '
    '                                                                 '

    Public Property PosY1() As Double
        Get
            Return Y1
        End Get
        Set(ByVal Value As Double)
            Y1 = Value
        End Set
    End Property


    'Property of the first endpoint z                                 '
    '                                                                 '

    Public Property PosZ1() As Double
        Get
            Return Z1
        End Get
        Set(ByVal Value As Double)
            Z1 = Value
        End Set
    End Property


    'Property of the second endpoint x                                '
    '                                                                 '

    Public Property PosX2() As Double
        Get
            Return X2
        End Get
        Set(ByVal Value As Double)
            X2 = Value
        End Set
    End Property


    'Property of the second endpoint y                                '
    '                                                                 '

    Public Property PosY2() As Double
        Get
            Return Y2
        End Get
        Set(ByVal Value As Double)
            Y2 = Value
        End Set
    End Property


    'Property of the second endpoint z                                '
    '                                                                 '

    Public Property PosZ2() As Double
        Get
            Return Z2
        End Get
        Set(ByVal Value As Double)
            Z2 = Value
        End Set
    End Property


    'Property of the third endpoint x                                 '
    '                                                                 '

    Public Property PosX3() As Double
        Get
            Return X3
        End Get
        Set(ByVal Value As Double)
            X3 = Value
        End Set
    End Property


    'Property of the third endpoint y                                 '
    '                                                                 '

    Public Property PosY3() As Double
        Get
            Return Y3
        End Get
        Set(ByVal Value As Double)
            Y3 = Value
        End Set
    End Property


    'Property of the third endpoint z                                 '
    '                                                                 '

    Public Property PosZ3() As Double
        Get
            Return Z3
        End Get
        Set(ByVal Value As Double)
            Z3 = Value
        End Set
    End Property

End Class


'''''''''''''''''''''''''''''''
'                                                                                                '
'Class CIDSChipEdge                                                                              '
'  Despription:                                                                                  '  
'           Chip edge element data handling class                                                '    
'  Created by:                                                                                   ' 
'           Shen Jian                                                                            '
'                                                                                                '
'''''''''''''''''''''''''''''''
Public Class CIDSChipEdge
    Inherits CIDSElements

    Protected X, Y, Z As Double
    Protected visionParam As DLL_Export_Device_Vision.ChipEdgePoints.ChipEdgeParam 'vision parameters


    'Property of vision parameters for chip detection                 '
    '                                                                 '

    Public Property vParam() As DLL_Export_Device_Vision.ChipEdgePoints.ChipEdgeParam
        Get
            Return visionParam
        End Get
        Set(ByVal Value As DLL_Export_Device_Vision.ChipEdgePoints.ChipEdgeParam)
            visionParam = Value
        End Set
    End Property


    'Property of chip position x                                      '
    '                                                                 '

    Public Property PosX() As Double
        Get
            Return X
        End Get
        Set(ByVal Value As Double)
            X = Value
        End Set
    End Property


    'Property of chip position y                                      '
    '                                                                 '

    Public Property PosY() As Double
        Get
            Return Y
        End Get
        Set(ByVal Value As Double)
            Y = Value
        End Set
    End Property


    'Property of chip position z                                      '
    '                                                                 '

    Public Property PosZ() As Double
        Get
            Return Z
        End Get
        Set(ByVal Value As Double)
            Z = Value
        End Set
    End Property


End Class


'''''''''''''''''''''''''''''''
'                                                                                                '
'Class CIDSQC                                                                                    '
'  Despription:                                                                                  '  
'           QC element data handling class                                                       '    
'  Created by:                                                                                   ' 
'           Shen Jian                                                                            '
'                                                                                                '
'''''''''''''''''''''''''''''''
Public Class CIDSQC
    Inherits CIDSElements

    Protected X, Y, Z As Double
    Protected visionParam As DLL_Export_Device_Vision.QC.QCParam


    'Property of vision parameters for QC inspection                  '
    '                                                                 '

    Public Property vParam() As DLL_Export_Device_Vision.QC.QCParam
        Get
            Return visionParam
        End Get
        Set(ByVal Value As DLL_Export_Device_Vision.QC.QCParam)
            visionParam = Value
        End Set
    End Property


    'Property of QC checking position x                               '
    '                                                                 '

    Public Property PosX() As Double
        Get
            Return X
        End Get
        Set(ByVal Value As Double)
            X = Value
        End Set
    End Property


    'Property of QC checking position y                               '
    '                                                                 '

    Public Property PosY() As Double
        Get
            Return Y
        End Get
        Set(ByVal Value As Double)
            Y = Value
        End Set
    End Property


    'Property of QC checking position z                               '
    '                                                                 '

    Public Property PosZ() As Double
        Get
            Return Z
        End Get
        Set(ByVal Value As Double)
            Z = Value
        End Set
    End Property

End Class


'''''''''''''''''''''''''''''''
'                                                                                                '
'Class CIDSSetIO                                                                                 '
'  Despription:                                                                                  '  
'           Set DIO element data handling class                                                  '    
'  Created by:                                                                                   ' 
'           Shen Jian                                                                            '
'                                                                                                '
'''''''''''''''''''''''''''''''
Public Class CIDSSetIO
    Inherits CIDSElements

    Protected IONo As Integer
    Protected IOOnOff As Integer


    'Property of digital output IO number                             '
    '                                                                 '

    Public Property SetIONo() As Integer
        Get
            Return IONo
        End Get
        Set(ByVal Value As Integer)
            IONo = Value
        End Set
    End Property


    'Property of digital output ON/OFF flag                           '
    '                                                                 '

    Public Property SetIOOnOffFlag() As Integer
        Get
            Return IOOnOff
        End Get
        Set(ByVal Value As Integer)
            IOOnOff = Value
        End Set
    End Property
End Class


'''''''''''''''''''''''''''''''
'                                                                                                '
'Class CIDSGetIO                                                                                 '
'  Despription:                                                                                  '  
'           Get DIO element data handling class                                                  '    
'  Created by:                                                                                   ' 
'           Shen Jian                                                                            '
'                                                                                                '
'''''''''''''''''''''''''''''''
Public Class CIDSGetIO
    Inherits CIDSElements

    Protected IONo As Integer


    'Property of digital input IO number                              '
    '                                                                 '

    Public Property GetIONo() As Integer
        Get
            Return IONo
        End Get
        Set(ByVal Value As Integer)
            IONo = Value
        End Set
    End Property
End Class


'''''''''''''''''''''''''''''''
'                                                                                                '
'Class CIDSSub                                                                                   '
'  Despription:                                                                                  '  
'           Call sub record data  class                                                          '    
'  Created by:                                                                                   ' 
'           Shen Jian                                                                            '
'                                                                                                '
'''''''''''''''''''''''''''''''
Public Class CIDSSub
    Inherits CIDSElements
    'Public SEflag As Integer  '1: start; 2: End.
    Public subName As String
    Public level As Integer
End Class


'''''''''''''''''''''''''''''''
'                                                                                                '
'Class CIDSWait                                                                                  '
'  Despription:                                                                                  '  
'           Wait element data handling class                                                     '    
'  Created by:                                                                                   ' 
'           Shen Jian                                                                            '
'                                                                                                '
'''''''''''''''''''''''''''''''
Public Class CIDSWait
    Inherits CIDSElements
    Protected X As Double
    Protected Y As Double
    Protected Z As Double


    'Property of waiting position x                                   '
    '                                                                 '

    Public Property PosX() As Double
        Get
            Return X
        End Get
        Set(ByVal Value As Double)
            X = Value
        End Set
    End Property


    'Property of waiting position y                                   '
    '                                                                 '

    Public Property PosY() As Double
        Get
            Return Y
        End Get
        Set(ByVal Value As Double)
            Y = Value
        End Set
    End Property


    'Property of waiting position z                                   '
    '                                                                 '

    Public Property PosZ() As Double
        Get
            Return Z
        End Get
        Set(ByVal Value As Double)
            Z = Value
        End Set
    End Property

End Class


'''''''''''''''''''''''''''''''
'                                                                                                '
'Class CIDSClean                                                                                '
'  Despription:                                                                                  '  
'           Clean element data handling class                                                   '    
'  Created by:                                                                                   ' 
'           Shen Jian                                                                            '
'                                                                                                '
'''''''''''''''''''''''''''''''
Public Class CIDSClean
    Inherits CIDSElements

End Class

'''''''''''''''''''''''''''''''
'                                                                                                '
'Class CIDSPurge                                                                                 '
'  Despription:                                                                                  '  
'           Purge element data handling class                                                    '    
'  Created by:                                                                                   ' 
'           Shen Jian                                                                            '
'                                                                                                '
'''''''''''''''''''''''''''''''
Public Class CIDSVolumeCalibration
    Inherits CIDSElements

End Class

'''''''''''''''''''''''''''''''
'                                                                                                '
'Class CIDSPurge                                                                                 '
'  Despription:                                                                                  '  
'           Purge element data handling class                                                    '    
'  Created by:                                                                                   ' 
'           Shen Jian                                                                            '
'                                                                                                '
'''''''''''''''''''''''''''''''
Public Class CIDSPurge
    Inherits CIDSElements

End Class


'''''''''''''''''''''''''''''''
'                                                                                                '
'Class CIDSHeight                                                                                '
'  Despription:                                                                                  '  
'           Height element data handling class                                                   '    
'  Created by:                                                                                   ' 
'           Shen Jian                                                                            '
'                                                                                                '
'''''''''''''''''''''''''''''''
Public Class CIDSHeight
    Inherits CIDSElements

    Protected X As Double
    Protected Y As Double
    Protected Z As Double


    'Property of height measuring position x                          '
    '                                                                 '

    Public Property PosX() As Double
        Get
            Return X
        End Get
        Set(ByVal Value As Double)
            X = Value
        End Set
    End Property


    'Property of height measuring position y                          '
    '                                                                 '

    Public Property PosY() As Double
        Get
            Return Y
        End Get
        Set(ByVal Value As Double)
            Y = Value
        End Set
    End Property


    'Property of height measuring position z                          '
    '                                                                 '

    Public Property PosZ() As Double
        Get
            Return Z
        End Get
        Set(ByVal Value As Double)
            Z = Value
        End Set
    End Property

End Class


'''''''''''''''''''''''''''''''
'                                                                                                '
'Class CIDSMove                                                                                '
'  Despription:                                                                                  '  
'           Move element data handling class                                                   '    
'  Created by:                                                                                   ' 
'           Shen Jian                                                                            '
'                                                                                                '
'''''''''''''''''''''''''''''''
Public Class CIDSMove
    Inherits CIDSElements

    Protected X As Double
    Protected Y As Double
    Protected Z As Double


    'Property of move to position x                                   '
    '                                                                 '

    Public Property PosX() As Double
        Get
            Return X
        End Get
        Set(ByVal Value As Double)
            X = Value
        End Set
    End Property


    'Property of move to position y                                   '
    '                                                                 '

    Public Property PosY() As Double
        Get
            Return Y
        End Get
        Set(ByVal Value As Double)
            Y = Value
        End Set
    End Property


    'Property of move to position z                                   '
    '                                                                 '

    Public Property PosZ() As Double
        Get
            Return Z
        End Get
        Set(ByVal Value As Double)
            Z = Value
        End Set
    End Property

End Class