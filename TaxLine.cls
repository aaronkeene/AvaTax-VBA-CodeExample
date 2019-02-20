VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TaxLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private cId As Long
Private sBusinessIdentificationNumber As String
Private sRevenueAccount As String
Private bDiscounted As Boolean
Private sExemptionCode As String
Private cAmount As Currency
Private tacAddresses As TransactionAddressCollection
Private aDetails() As TaxLineDetail
Private bTaxIncluded As Boolean
Private sTaxOverrideReason As String
Private cTaxOverrideAmount As Currency
Private sTaxOverrideType As String
Private sTaxEngine As String
Private dTaxDate As Date
Private cTaxCodeId As Long
Private sTaxCode As String
Private cTaxCalculated As Currency
Private cTaxableAmount As Currency
Private cTax As Currency
Private sSourcing As String
Private sRevAccount As String
Private dReportingDate As Date
Private sRef2 As String
Private sRef1 As String
Private cQuantity As Currency
Private cOriginAddressId As Long
Private cLineAmount As Currency
Private sItemCode As String
Private bIsSSTP As Boolean
Private bIsItemTaxable As Boolean
Private sExemptNo As String
Private cExemptCertId As Long
Private cExemptAmount As Currency
Private cDiscountTypeId As Long
Private cDiscountAmount As Currency
Private cDestinationAddressId As Long
Private sDescription As String
Private sCustomerUsageType As String
Private cBoundaryOverrideId As Long
Private sMethod As String
Private sLineNumber As String
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

Property Get LineNumber() As String

    Let LineNumber = sLineNumber

End Property

Property Let LineNumber(LineNumber As String)

    Let sLineNumber = LineNumber

End Property


Property Get Method() As String

    Let Method = sMethod

End Property

Property Let Method(Method As String)

    Let sMethod = Method

End Property



Property Get BoundaryOverrideId() As Long

    Let BoundaryOverrideId = cBoundaryOverrideId

End Property

Property Let BoundaryOverrideId(BoundaryOverrideId As Long)

    Let cBoundaryOverrideId = BoundaryOverrideId

End Property



Property Get CustomerUsageType() As String

    Let CustomerUsageType = sCustomerUsageType

End Property

Property Let CustomerUsageType(CustomerUsageType As String)

    Let sCustomerUsageType = CustomerUsageType

End Property



Property Get Description() As String

    Let Description = sDescription

End Property

Property Let Description(Description As String)

    Let sDescription = Description

End Property



Property Get DestinationAddressId() As Long

    Let DestinationAddressId = cDestinationAddressId

End Property

Property Let DestinationAddressId(DestinationAddressId As Long)

    Let cDestinationAddressId = DestinationAddressId

End Property



Property Get DiscountAmount() As Currency

    Let DiscountAmount = cDiscountAmount

End Property

Property Let DiscountAmount(DiscountAmount As Currency)

    Let cDiscountAmount = DiscountAmount

End Property



Property Get DiscountTypeId() As Long

    Let DiscountTypeId = cDiscountTypeId

End Property

Property Let DiscountTypeId(DiscountTypeId As Long)

    Let cDiscountTypeId = DiscountTypeId

End Property



Property Get ExemptAmount() As Currency

    Let ExemptAmount = cExemptAmount

End Property

Property Let ExemptAmount(ExemptAmount As Currency)

    Let cExemptAmount = ExemptAmount

End Property



Property Get ExemptCertId() As Long

    Let ExemptCertId = cExemptCertId

End Property

Property Let ExemptCertId(ExemptCertId As Long)

    Let cExemptCertId = ExemptCertId

End Property



Property Get ExemptNo() As String

    Let ExemptNo = sExemptNo

End Property

Property Let ExemptNo(ExemptNo As String)

    Let sExemptNo = ExemptNo

End Property



Property Get IsItemTaxable() As Boolean

    Let IsItemTaxable = bIsItemTaxable

End Property

Property Let IsItemTaxable(IsItemTaxable As Boolean)

    Let bIsItemTaxable = IsItemTaxable

End Property



Property Get IsSSTP() As Boolean

    Let IsSSTP = bIsSSTP

End Property

Property Let IsSSTP(IsSSTP As Boolean)

    Let bIsSSTP = IsSSTP

End Property



Property Get ItemCode() As String

    Let ItemCode = sItemCode

End Property

Property Let ItemCode(ItemCode As String)

    Let sItemCode = ItemCode

End Property



Property Get LineAmount() As Currency

    Let LineAmount = cLineAmount

End Property

Property Let LineAmount(LineAmount As Currency)

    Let cLineAmount = LineAmount

End Property



Property Get OriginAddressId() As Long

    Let OriginAddressId = cOriginAddressId

End Property

Property Let OriginAddressId(OriginAddressId As Long)

    Let cOriginAddressId = OriginAddressId

End Property



Property Get Quantity() As Currency

    Let Quantity = cQuantity

End Property

Property Let Quantity(Quantity As Currency)

    Let cQuantity = Quantity

End Property


Property Get Ref1() As String

    Let Ref1 = sRef1

End Property

Property Let Ref1(Ref1 As String)

    Let sRef1 = Ref1

End Property



Property Get Ref2() As String

    Let Ref2 = sRef2

End Property

Property Let Ref2(Ref2 As String)

    Let sRef2 = Ref2

End Property


Property Get ReportingDate() As Date

    Let ReportingDate = dReportingDate

End Property

Property Let ReportingDate(ReportingDate As Date)

    Let dReportingDate = ReportingDate

End Property



Property Get RevAccount() As String

    Let RevAccount = sRevAccount

End Property

Property Let RevAccount(RevAccount As String)

    Let sRevAccount = RevAccount

End Property



Property Get Sourcing() As String

    Let Sourcing = sSourcing

End Property

Property Let Sourcing(Sourcing As String)

    Let sSourcing = Sourcing

End Property



Property Get Tax() As Currency

    Let Tax = cTax

End Property

Property Let Tax(Tax As Currency)

    Let cTax = Tax

End Property



Property Get TaxableAmount() As Currency

    Let TaxableAmount = cTaxableAmount

End Property

Property Let TaxableAmount(TaxableAmount As Currency)

    Let cTaxableAmount = TaxableAmount

End Property



Property Get TaxCalculated() As Currency

    Let TaxCalculated = cTaxCalculated

End Property

Property Let TaxCalculated(TaxCalculated As Currency)

    Let cTaxCalculated = TaxCalculated

End Property



Property Get TaxCode() As String

    Let TaxCode = sTaxCode

End Property

Property Let TaxCode(TaxCode As String)

    Let sTaxCode = TaxCode

End Property



Property Get TaxCodeId() As Long

    Let TaxCodeId = cTaxCodeId

End Property

Property Let TaxCodeId(TaxCodeId As Long)

    Let cTaxCodeId = TaxCodeId

End Property



Property Get TaxDate() As Date

    Let TaxDate = dTaxDate

End Property

Property Let TaxDate(TaxDate As Date)

    Let dTaxDate = TaxDate

End Property



Property Get TaxEngine() As String

    Let TaxEngine = sTaxEngine

End Property

Property Let TaxEngine(TaxEngine As String)

    Let sTaxEngine = TaxEngine

End Property



Property Get TaxOverrideType() As String

    Let TaxOverrideType = sTaxOverrideType

End Property

Property Let TaxOverrideType(TaxOverrideType As String)

    Let sTaxOverrideType = TaxOverrideType

End Property



Property Get TaxOverrideAmount() As Currency

    Let TaxOverrideAmount = cTaxOverrideAmount

End Property

Property Let TaxOverrideAmount(TaxOverrideAmount As Currency)

    Let cTaxOverrideAmount = TaxOverrideAmount

End Property



Property Get TaxOverrideReason() As String

    Let TaxOverrideReason = sTaxOverrideReason

End Property

Property Let TaxOverrideReason(TaxOverrideReason As String)

    Let sTaxOverrideReason = TaxOverrideReason

End Property



Property Get TaxIncluded() As Boolean

    Let TaxIncluded = bTaxIncluded

End Property

Property Let TaxIncluded(TaxIncluded As Boolean)

    Let bTaxIncluded = TaxIncluded

End Property


Property Get Details(index As Long) As TaxLineDetail

    Set Details = aDetails(index)
    
End Property

Property Set Details(index As Long, detail As TaxLineDetail)

    Set aDetails(index) = detail

End Property


Property Get addresses() As TransactionAddressCollection

    Set addresses = tacAddresses

End Property

Property Set addresses(addresses As TransactionAddressCollection)

    Set tacAddresses = addresses

End Property



Property Get Amount() As Currency

    Let Amount = cAmount
    
End Property

Property Let Amount(Amount As Currency)

    Let cAmount = Amount

End Property



Property Get ExemptionCode() As String

    Let ExemptionCode = sExemptionCode

End Property

Property Let ExemptionCode(ExemptionCode As String)

    Let sExemptionCode = ExemptionCode

End Property


Property Get Discounted() As Boolean

    Let Discounted = bDiscounted

End Property

Property Let Discounted(Discounted As Boolean)

    Let bDiscounted = Discounted

End Property


Property Get RevenueAccount() As String

    Let RevenueAccount = sRevenueAccount

End Property

Property Let RevenueAccount(RevenueAccount As String)

    Let sRevenueAccount = RevenueAccount

End Property


Property Get BusinessIdentificationNumber() As String

    Let BusinessIdentificationNumber = sBusinessIdentificationNumber
    
End Property

Property Let BusinessIdentificationNumber(BusinessIdentificationNumber As String)

    Let sBusinessIdentificationNumber = BusinessIdentificationNumber

End Property

Function ToJson() As Object
    Dim Json As Object
    Set Json = JsonConverter.ParseJson("{}")
    
    If Me.LineNumber = "" Then
        Err.Raise vbObjectError + 1000, Description:="A line number must be provided."
    End If
    
    Json("number") = Me.LineNumber
    
    Json("quantity") = Me.Quantity
    
    Json("amount") = Me.Amount
    
    If Not Me.addresses Is Nothing Then
    
        Set Json("addresses") = Me.addresses.ToJson()
    
    End If
    
    If Me.TaxCode <> "" Then
    
        Json("taxCode") = Me.TaxCode
    
    End If
    
    If Me.CustomerUsageType <> "" Then
    
        Json("customerUsageType") = Me.CustomerUsageType
    
    End If
    
    If Me.ItemCode <> "" Then
        
        Json("itemCode") = Me.ItemCode
        
    End If
    
    If Me.ExemptionCode <> "" Then
    
        Json("exemptionCode") = Me.ExemptionCode
    
    End If
    
    Json("discounted") = Me.Discounted
        
    Json("taxIncluded") = Me.TaxIncluded
    
    If Me.RevenueAccount <> "" Then
    
        Json("revenueAccount") = Me.RevenueAccount
    
    End If
    
    If Me.Ref1 <> "" Then
        
        Json("ref1") = Me.Ref1
        
    End If
    
    If Me.Ref2 <> "" Then
    
        Json("ref2") = Me.Ref2
    
    End If
    
    If Me.Description <> "" Then
        
        Json("description") = Me.Description
        
    End If
    
    If Me.BusinessIdentificationNumber <> "" Then
    
        Json("businessIdentificationNo") = Me.BusinessIdentificationNumber
        
    End If
    
    If Me.TaxOverrideAmount <> 0 Then
    
        Json("taxOverride")("taxAmount") = Me.TaxOverrideAmount
        Json("taxOverride")("type") = Me.TaxOverrideType
        Json("taxOverride")("reason") = Me.TaxOverrideReason
        Json("taxOverride")("taxDate") = Me.TaxDate
    
    End If
    
    Set ToJson = Json
End Function

