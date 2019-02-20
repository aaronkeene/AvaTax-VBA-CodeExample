VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TransactionRequest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private sCompanyCode As String
Private dTransactionDate As Date
Private sCode As String
Private sTransactionType As String
Private sCustomerCode As String
Private sSalespersonCode As String
Private tacAddresses As TransactionAddressCollection
Private aLines As Collection
Private sReferenceCode As String
Private cHashCode As Long
Private sReportingLocationCode As String
Private bCommit As Boolean
Private sBatchCode As String
Private sTaxOverrideType As String
Private cTaxOverrideAmount As Currency
Private dTaxOverrideDate As Date
Private sTaxOverrideReason As String
Private dTaxDate As Date
Private sCurrencyCode As String
Private sServiceMode As String
Private dPaymentDate As Date
Private cExchangeRate As Currency
Private dExchangeRateEffectiveDate As Date
Private sPosLaneCode As String
Private sBusinessIdentificationNo As String
Private bIsSellerImporterOfRecord As Boolean
Private sDescription As String
Private sEmail As String
Private sExemptionNo As String
Private sPurchaseOrderNo As String
Private cDiscount As Currency
Private sCustomerUsageType As String




Property Get CompanyCode() As String

    CompanyCode = sCompanyCode

End Property

Property Let CompanyCode(CompanyCode As String)

    sCompanyCode = CompanyCode

End Property



Property Get TransactionType() As String

    TransactionType = sTransactionType

End Property

Property Let TransactionType(TransactionType As String)

    sTransactionType = TransactionType

End Property



Property Get Code() As String

    Code = sCode

End Property

Property Let Code(Code As String)

    sCode = Code

End Property




Property Get TransactionDate() As Date

    TransactionDate = dTransactionDate

End Property

Property Let TransactionDate(TransactionDate As Date)

    dTransactionDate = TransactionDate

End Property


Property Get SalespersonCode() As String

    SalespersonCode = sSalespersonCode

End Property

Property Let SalespersonCode(SalespersonCode As String)

    sSalespersonCode = SalespersonCode

End Property




Property Get customerCode() As String

    customerCode = sCustomerCode

End Property

Property Let customerCode(customerCode As String)

    sCustomerCode = customerCode

End Property



Property Get CustomerUsageType() As String

    CustomerUsageType = sCustomerUsageType

End Property

Property Let CustomerUsageType(CustomerUsageType As String)

    sCustomerUsageType = CustomerUsageType

End Property



Property Get Discount() As Currency

    Discount = cDiscount

End Property

Property Let Discount(Discount As Currency)

    cDiscount = Discount

End Property



Property Get PurchaseOrderNo() As String

    PurchaseOrderNo = sPurchaseOrderNo

End Property

Property Let PurchaseOrderNo(PurchaseOrderNo As String)

    sPurchaseOrderNo = PurchaseOrderNo

End Property



Property Get ExemptionNo() As String

    ExemptionNo = sExemptionNo

End Property

Property Let ExemptionNo(ExemptionNo As String)

    sExemptionNo = ExemptionNo

End Property


Property Get addresses() As TransactionAddressCollection

    addresses = tacAddresses

End Property

Property Set addresses(addresses As TransactionAddressCollection)

    Set tacAddresses = addresses

End Property

Property Get Lines() As Collection
    
    Set Lines = aLines

End Property

Property Set Lines(txLines As Collection)

    Set aLines = txLines

End Property

Property Get ReferenceCode() As String

    ReferenceCode = sReferenceCode

End Property

Property Let ReferenceCode(refCode As String)

    sReferenceCode = refCode

End Property



Property Get HashCode() As Long

    HashCode = cHashCode

End Property

Property Let HashCode(HashCode As Long)

    cHashCode = HashCode

End Property



Property Get ReportingLocationCode() As String

    ReportingLocationCode = sReportingLocationCode

End Property

Property Let ReportingLocationCode(ReportingLocationCode As String)

    sReportingLocationCode = ReportingLocationCode

End Property



Property Get Commit() As Boolean

    Commit = bCommit

End Property

Property Let Commit(Commit As Boolean)

    bCommit = Commit

End Property



Property Get BatchCode() As String

    BatchCode = sBatchCode

End Property

Property Let BatchCode(BatchCode As String)

    sBatchCode = BatchCode

End Property


Property Get TaxOverrideType() As String

    TaxOverrideType = sTaxOverrideType

End Property

Property Let TaxOverrideType(TaxOverrideType As String)

    sTaxOverrideType = TaxOverrideType

End Property



Property Get TaxOverrideAmount() As Currency

    TaxOverrideAmount = cTaxOverrideAmount

End Property

Property Let TaxOverrideAmount(TaxOverrideAmount As Currency)

    cTaxOverrideAmount = TaxOverrideAmount

End Property



Property Get TaxOverrideDate() As Date

    TaxOverrideDate = dTaxOverrideDate

End Property

Property Let TaxOverrideDate(TaxOverrideDate As Date)

    dTaxOverrideDate = TaxOverrideDate

End Property


Property Get TaxOverrideReason() As String

    TaxOverrideReason = sTaxOverrideReason

End Property

Property Let TaxOverrideReason(TaxOverrideReason As String)

    sTaxOverrideReason = TaxOverrideReason

End Property


Property Get TaxDate() As Date

    TaxDate = dTaxDate

End Property

Property Let TaxDate(TaxDate As Date)

    dTaxDate = TaxDate

End Property



Property Get CurrencyCode() As String

    CurrencyCode = sCurrencyCode

End Property

Property Let CurrencyCode(CurrencyCode As String)

    sCurrencyCode = CurrencyCode

End Property



Property Get ServiceMode() As String

    ServiceMode = sServiceMode

End Property

Property Let ServiceMode(ServiceMode As String)

    sServiceMode = ServiceMode

End Property



Property Get PaymentDate() As Date

    PaymentDate = dPaymentDate

End Property

Property Let PaymentDate(PaymentDate As Date)

    dPaymentDate = PaymentDate

End Property



Property Get ExchangeRate() As Currency

    ExchangeRate = cExchangeRate

End Property

Property Let ExchangeRate(ExchangeRate As Currency)

    cExchangeRate = ExchangeRate

End Property



Property Get ExchangeRateEffectiveDate() As Date

    ExchangeRateEffectiveDate = dExchangeRateEffectiveDate

End Property

Property Let ExchangeRateEffectiveDate(ExchangeRateEffectiveDate As Date)

    dExchangeRateEffectiveDate = ExchangeRateEffectiveDate

End Property



Property Get PosLaneCode() As String

    PosLaneCode = sPosLaneCode

End Property

Property Let PosLaneCode(PosLaneCode As String)

    sPosLaneCode = PosLaneCode

End Property



Property Get BusinessIdentificationNo() As String

    BusinessIdentificationNo = sBusinessIdentificationNo

End Property

Property Let BusinessIdentificationNo(BusinessIdentificationNo As String)

    sBusinessIdentificationNo = BusinessIdentificationNo

End Property



Property Get IsSellerImporterOfRecord() As Boolean

    IsSellerImporterOfRecord = bIsSellerImporterOfRecord

End Property

Property Let IsSellerImporterOfRecord(IsSellerImporterOfRecord As Boolean)

    bIsSellerImporterOfRecord = IsSellerImporterOfRecord

End Property



Property Get Description() As String

    Description = sDescription

End Property

Property Let Description(Description As String)

    sDescription = Description

End Property



Property Get Email() As String

    Email = sEmail

End Property

Property Let Email(Email As String)

    sEmail = Email

End Property

Function ToJson() As Object

    Dim Json As Object
    Set Json = JsonConverter.ParseJson("{}")
    
    If Me.CompanyCode = "" Then
        Err.Raise vbObjectError + 1000, Description:="CompanyCode is required."
    End If
    
    Json("companyCode") = Me.CompanyCode
    
    If Me.TransactionType = "" Then
        Json("type") = "SalesOrder"
    Else
        Json("type") = Me.TransactionType
    End If
    
    If Me.Code = "" Then
        Let Me.Code = Mid$(CreateObject("Scriptlet.TypeLib").GUID, 2, 36)
    End If
    
    Json("code") = Me.Code
    
    If Me.customerCode = "" Then
        Err.Raise vbObjectError + 1000, Description:="CustomerCode is required."
    End If
    
    If Me.TransactionDate = CDate(0) Then
        Json("date") = Now
    Else
        Json("date") = Me.TransactionDate
    End If
    
    If Me.SalespersonCode <> "" Then
        Json("salespersonCode") = Me.SalespersonCode
    End If
    
    If Me.customerCode <> "" Then
        Json("customerCode") = Me.customerCode
    End If
    
    If Me.CustomerUsageType <> "" Then
        Json("customerUsageType") = Me.CustomerUsageType
    End If
    
    If Me.Discount = 0 Then
        Json("discount") = Me.Discount
    End If
    
    If Me.PurchaseOrderNo <> "" Then
        Json("purchaseOrderNo") = Me.PurchaseOrderNo
    End If
    
    If Me.ExemptionNo <> "" Then
        Json("exemptionNo") = Me.ExemptionNo
    End If
    
    Dim line As TaxLine
    
    If tacAddresses Is Nothing Then
        Dim noAddresses As Boolean
        Let noAddresses = True
        
        For Each line In Me.Lines
            If Not line Is Nothing Then
                If Not line.addresses Is Nothing Then
                    Let noAddresses = False
                End If
            End If
        Next
        
        If noAddresses Then
            Err.Raise vbObjectError + 1000, Description:="An address is required at either the transaction or tax line level."
        End If
    Else
        Set Json("addresses") = tacAddresses.ToJson()
    End If
    
    Dim taxLines As New Collection
    
    For Each line In Me.Lines
        If Not line Is Nothing Then
            Dim lineJson As Object
            Set lineJson = line.ToJson()
            Debug.Print JsonConverter.ConvertToJson(lineJson)
            taxLines.Add lineJson
        End If
    Next
    
    Set Json("lines") = taxLines
    
    Set ToJson = Json
    
End Function
