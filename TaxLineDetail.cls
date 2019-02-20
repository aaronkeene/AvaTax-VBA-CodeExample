VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TaxLineDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Private dReportingDate As Date

Property Get ReportingDate() As Date

    Set ReportingDate = dReportingDate

End Property

Property Let ReportingDate(ReportingDate As Date)

    Set dReportingDate = ReportingDate

End Property


Private sRevAccount As String

Property Get RevAccount() As String

    Set RevAccount = sRevAccount

End Property

Property Let RevAccount(RevAccount As String)

    Set sRevAccount = RevAccount

End Property


Private sSourcing As String

Property Get Sourcing() As String

    Set Sourcing = sSourcing

End Property

Property Let Sourcing(Sourcing As String)

    Set sSourcing = Sourcing

End Property


Private cTax As Currency

Property Get Tax() As Currency

    Set Tax = cTax

End Property

Property Let Tax(Tax As Currency)

    Set cTax = Tax

End Property


Private cTaxableAmount As Currency

Property Get TaxableAmount() As Currency

    Set TaxableAmount = cTaxableAmount

End Property

Property Let TaxableAmount(TaxableAmount As Currency)

    Set cTaxableAmount = TaxableAmount

End Property


Private cTaxCalculated As Currency

Property Get TaxCalculated() As Currency

    Set TaxCalculated = cTaxCalculated

End Property

Property Let TaxCalculated(TaxCalculated As Currency)

    Set cTaxCalculated = TaxCalculated

End Property


Private sTaxCode As String

Property Get TaxCode() As String

    Set TaxCode = sTaxCode

End Property

Property Let TaxCode(TaxCode As String)

    Set sTaxCode = TaxCode

End Property


Private cTaxCodeId As Long

Property Get TaxCodeId() As Long

    Set TaxCodeId = cTaxCodeId

End Property

Property Let TaxCodeId(TaxCodeId As Long)

    Set cTaxCodeId = TaxCodeId

End Property


Private dTaxDate As Date

Property Get TaxDate() As Date

    Set TaxDate = dTaxDate

End Property

Property Let TaxDate(TaxDate As Date)

    Set dTaxDate = TaxDate

End Property


Private sTaxEngine As String

Property Get TaxEngine() As String

    Set TaxEngine = sTaxEngine

End Property

Property Let TaxEngine(TaxEngine As String)

    Set sTaxEngine = TaxEngine

End Property


Private sTaxOverrideType As String

Property Get TaxOverrideType() As String

    Set TaxOverrideType = sTaxOverrideType

End Property

Property Let TaxOverrideType(TaxOverrideType As String)

    Set sTaxOverrideType = TaxOverrideType

End Property


Private cTaxOverrideAmount As Currency

Property Get TaxOverrideAmount() As Currency

    Set TaxOverrideAmount = cTaxOverrideAmount

End Property

Property Let TaxOverrideAmount(TaxOverrideAmount As Currency)

    Set cTaxOverrideAmount = TaxOverrideAmount

End Property


Private sTaxOverrideReason As String

Property Get TaxOverrideReason() As String

    Set TaxOverrideReason = sTaxOverrideReason

End Property

Property Let TaxOverrideReason(TaxOverrideReason As String)

    Set sTaxOverrideReason = TaxOverrideReason

End Property


Private bTaxIncluded As Boolean

Property Get TaxIncluded() As Boolean

    Set TaxIncluded = bTaxIncluded

End Property

Property Let TaxIncluded(TaxIncluded As Boolean)

    Set bTaxIncluded = TaxIncluded

End Property

Private sRateType As String

Property Get RateType() As String

    Set RateType = sRateType

End Property

Property Let RateType(RateType As String)

    Set sRateType = RateType
    
End Property
End Property
