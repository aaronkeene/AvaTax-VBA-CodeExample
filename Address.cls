VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Address"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Private cId As Long
Private fLongitude As Double
Private fLatitude As Double
Private lTaxRegionId As Long
Private sCountry As String
Private sPostalCode As String
Private sRegion As String
Private sCity As String
Private sLine3 As String
Private sLine2 As String
Private sLine1 As String
Private sBoundaryLevel As String
Private cTransactionId As Long

Property Get Id() As Long

    Let Id = cId

End Property

Property Let Id(Id As Long)

    Let cId = Id

End Property



Property Get TransactionId() As Long

    Let TransactionId = cTransactionId

End Property

Property Let TransactionId(TransactionId As Long)

    Let cTransactionId = TransactionId

End Property



Property Get BoundaryLevel() As String

    Let BoundaryLevel = sBoundaryLevel

End Property

Property Let BoundaryLevel(BoundaryLevel As String)

    Let sBoundaryLevel = BoundaryLevel

End Property


Property Get Line1() As String

    Let Line1 = sLine1

End Property

Property Let Line1(Line1 As String)

    Let sLine1 = Line1

End Property


Property Get Line2() As String

    Let Line2 = sLine2

End Property

Property Let Line2(Line2 As String)

    Let sLine2 = Line2

End Property


Property Get Line3() As String

    Let Line3 = sLine3

End Property

Property Let Line3(Line3 As String)

    Let sLine3 = Line3

End Property



Property Get City() As String

    Let City = sCity

End Property

Property Let City(City As String)

    Let sCity = City

End Property



Property Get Region() As String

    Let Region = sRegion

End Property

Property Let Region(Region As String)

    Let sRegion = Region

End Property



Property Get PostalCode() As String

    Let PostalCode = sPostalCode

End Property

Property Let PostalCode(PostalCode As String)

    Let sPostalCode = PostalCode

End Property



Property Get Country() As String

    Let Country = sCountry

End Property

Property Let Country(Country As String)

    Let sCountry = Country

End Property


Property Get TaxRegionId() As Long

    Let TaxRegionId = lTaxRegionId

End Property

Property Let TaxRegionId(TaxRegionId As Long)

    Let lTaxRegionId = TaxRegionId

End Property




Property Get Latitude() As Double

    Let Latitude = fLatitude

End Property

Property Let Latitude(Latitude As Double)

    Let fLatitude = Latitude

End Property



Property Get Longitude() As Double

    Let Longitude = fLongitude

End Property

Property Let Longitude(Longitude As Double)

    Let fLongitude = Longitude

End Property


Function ToJson() As Object

    Dim Json As Object
    Set Json = JsonConverter.ParseJson("{}")

    Json("line1") = Me.Line1
    Json("line2") = Me.Line2
    Json("line3") = Me.Line3
    Json("city") = Me.City
    Json("region") = Me.Region
    Json("country") = Me.Country
    Json("postalCode") = Me.PostalCode
    
    If Round(Me.Latitude, 6) <> Round(0, 6) Or Round(Me.Longitude, 6) <> Round(0, 6) Then
        Json("latitude") = Me.Latitude
        Json("longitude") = Me.Longitude
    End If
    
    Set ToJson = Json
    
End Function
