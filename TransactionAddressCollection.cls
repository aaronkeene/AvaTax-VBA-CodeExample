VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TransactionAddressCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private aShipFrom As Address
Private aShipTo As Address
Private aPointOfOrderAcceptance As Address
Private aPointOfOrderOrigin As Address
Private aSingleLocation As Address

Property Get ShipFrom() As Address

    Set ShipFrom = aShipFrom
    
End Property

Property Set ShipFrom(Address As Address)

    If Not aSingleLocation Is Nothing Then
    
        Err.Raise vbObjectError + 1000, Description:="ShipFrom cannot be set if SingleLocation is set."
    
    Else
    
        Set aShipFrom = Address
    
    End If

End Property


Property Get ShipTo() As Address

    Set ShipTo = aShipTo
    
End Property

Property Set ShipTo(Address As Address)

    If Not aSingleLocation Is Nothing Then
    
        Err.Raise vbObjectError + 1000, Description:="ShipTo cannot be set if SingleLocation is set."
    
    Else
    
        Set aShipTo = Address
    
    End If

End Property


Property Get PointOfOrderAcceptance() As Address

    Set PointOfOrderAcceptance = aPointOfOrderAcceptance
    
End Property

Property Set PointOfOrderAcceptance(Address As Address)

    If Not aSingleLocation Is Nothing Then
    
        Err.Raise vbObjectError + 1000, Description:="PointOfOrderAcceptance cannot be set if SingleLocation is set."
    
    Else
    
        Set aPointOfOrderAcceptance = Address
    
    End If

End Property


Property Get PointOfOrderOrigin() As Address

    Set PointOfOrderOrigin = aPointOfOrderOrigin
    
End Property

Property Set PointOfOrderOrigin(Address As Address)

    If Not aSingleLocation Is Nothing Then
    
        Err.Raise vbObjectError + 1000, Description:="PointOfOrderOrigin cannot be set if SingleLocation is set."
    
    Else
    
        Set aShipTo = Address
    
    End If

End Property


Property Get SingleLocation() As Address

    Set SingleLocation = aSingleLocation

End Property

Property Set SingleLocation(Address As Address)

    If Not aShipFrom Is Nothing _
        Or Not aShipTo Is Nothing _
        Or Not aPointOfOrderAcceptance Is Nothing _
        Or Not aPointOfOrderOrigin Is Nothing Then
    
        Err.Raise vbObjectError + 1000, Description:="SingleLocation cannot be set if other types of addresses are set."
    
    Else
    
        Set aSingleLocation = Address
    
    End If

End Property


Function ToJson() As Object

    Dim addresses As Object
    Set addresses = JsonConverter.ParseJson("{}")
    
    If Not Me.ShipFrom Is Nothing Then
    
        If Not Me.ShipTo Is Nothing Then
        
            Set addresses("ShipFrom") = Me.ShipFrom.ToJson()
            Set addresses("ShipTo") = Me.ShipTo.ToJson()
        
        Else
        
            Err.Raise vbObjectError + 1000, Description:="ShipFrom must have ShipTo set as well."
            
        End If
        
        If Not Me.PointOfOrderAcceptance Is Nothing Then
    
            Set addresses("PointOfOrderAcceptance") = Me.PointOfOrderAcceptance.ToJson()
            
        End If
        
        If Not Me.PointOfOrderOrigin Is Nothing Then
            
            Set addresses("PointOfOrderOrigin") = Me.PointOfOrderOrigin.ToJson()
            
        End If
        
    ElseIf Not Me.SingleLocation Is Nothing Then
    
        Set addresses("SingleLocation") = Me.SingleLocation.ToJson()
    
    ElseIf Not Me.ShipTo Is Nothing Then
    
        Err.Raise vbObjectError + 1000, Description:="ShipFrom must be set if ShipTo is set."
        
    End If
        
    Set ToJson = addresses

End Function
