Attribute VB_Name = "AvaTaxForExcel"
Option Explicit

Private sBaseUrl As String
Private sUserPass As String

Property Get BaseUrl() As String

    If sBaseUrl = "" Then
        
        Let sBaseUrl = "https://sandbox-rest.avatax.com"
        
    End If
    
    BaseUrl = sBaseUrl
    
End Property

Property Let BaseUrl(sUrl As String)

    sBaseUrl = sUrl

End Property

Property Get UserPass() As String

    If sUserPass = "" Then
    
        Let sUserPass = InputBox("Please Enter Your Base64 Encoded AvaTax Username and Password", "Credentials")
    
    End If
    

    Let UserPass = sUserPass

End Property

Property Let UserPass(up As String)

    Let sUserPass = up

End Property

Private Function HTTPGet(sPath As String, sQuery As String) As String

    Dim sResult As String
    Dim sUrl As String
    Let sUrl = BaseUrl & sPath
    
    Dim objHTTP As Object
    Set objHTTP = CreateObject("WinHttp.WinHttpRequest.5.1")
    
    objHTTP.Open "GET", sUrl, False
    
    objHTTP.send (sQuery)
    Debug.Print objHTTP.Status
    Debug.Print objHTTP.ResponseText
    
    sResult = objHTTP.ResponseText
        
    HTTPGet = sResult

End Function

Private Function HTTPPost(sPath As String, sPostData As String, sHeader1 As String, sHeader2 As String, sHeader3 As String) As String
    
    Dim sResult As String
    Dim sUrl As String
    Let sUrl = BaseUrl & sPath
    
    Dim objHTTP As Object
    Set objHTTP = CreateObject("WinHttp.WinHttpRequest.5.1")
    
    objHTTP.Open "POST", sUrl, False
    
    If sHeader1 <> "" Then
        objHTTP.setRequestHeader Split(sHeader1, ": ", 2)(0), Split(sHeader1, ": ", 2)(1)
    End If
    
    If sHeader2 <> "" Then
        objHTTP.setRequestHeader Split(sHeader2, ": ", 2)(0), Split(sHeader2, ": ", 2)(1)
    End If
    
    If sHeader3 <> "" Then
        objHTTP.setRequestHeader Split(sHeader3, ": ", 2)(0), Split(sHeader3, ": ", 2)(1)
    End If
    
    objHTTP.send (sPostData)
    Debug.Print objHTTP.Status
    Debug.Print objHTTP.ResponseText
    
    sResult = objHTTP.ResponseText
    
    HTTPPost = sResult

End Function


Function GetTax(docDate As Date, customerCode As String, shipFromLine1 As String, shipFromLine2 As String, shipFromLine3 As String, shipFromCity As String, shipFromRegion As String, shipFromCountry As String, shipFromPostalCode As String, _
    shipToLine1 As String, shipToLine2 As String, shipToLine3 As String, shipToCity As String, shipToRegion As String, shipToCountry As String, shipToPostalCode As String, _
    lineHeaders As Range, lineItems As Range) As String

    Dim transaction As TransactionRequest
    Set transaction = New TransactionRequest
    
    On Error GoTo errCantConvert
    transaction.CompanyCode = "default"
    transaction.customerCode = customerCode
    transaction.TransactionDate = docDate
    ' TODO: Change this to SalesOrder if you do not want these transactions to post to the ledger
    transaction.TransactionType = "SalesInvoice"
    
    Dim shipFromLocation As New Address
    Dim shipToLocation As New Address
    Dim addressCollection As New TransactionAddressCollection
    
    shipFromLocation.Line1 = shipFromLine1
    shipFromLocation.Line2 = shipFromLine2
    shipFromLocation.Line3 = shipFromLine3
    shipFromLocation.City = shipFromCity
    shipFromLocation.Region = shipFromRegion
    shipFromLocation.Country = shipFromCountry
    shipFromLocation.PostalCode = shipFromPostalCode
    
    shipToLocation.Line1 = shipToLine1
    shipToLocation.Line2 = shipToLine2
    shipToLocation.Line3 = shipToLine3
    shipToLocation.City = shipToCity
    shipToLocation.Region = shipToRegion
    shipToLocation.Country = shipToCountry
    shipToLocation.PostalCode = shipToPostalCode
    
    Dim usedLineCount As Integer
    
    Set transaction.Lines = New Collection
    
    usedLineCount = 0
    
    Dim row As Long
    Dim column As Long
    
    For row = LBound(lineItems.Value2, 1) To UBound(lineItems.Value2, 1)
        
        Dim line As TaxLine
        Set line = New TaxLine
        Let line.LineNumber = row
    
        For column = 1 To lineHeaders.Count
        
            Dim header As Variant
            Let header = lineHeaders.Value2(1, column)
            
            Dim cell As Variant
            Let cell = lineItems.Value2(row, column)
            
            If cell <> Empty Then
            
                If header = "Item" Then
                
                    line.Description = cell
                
                End If
                
                If header = "Type" Then
                
                    If cell = "Taxable Item" Then
                    
                        line.TaxCode = "P0000000"
                    
                    ElseIf cell = "Non-Taxable Item" Then
                    
                        line.TaxCode = "NT"
                    
                    ElseIf cell = "Shipping" Then
                    
                        line.TaxCode = "FR020100"
                        
                    End If
                    
                End If
                
                If header = "Qty" Then
                
                    line.Quantity = cell
                    
                End If
                
                If header = "Amount" Then
                
                    line.Amount = cell
                
                End If
                
            Else
                
                If header = "Amount" Then
                    
                    line.TaxCode = ""
                
                End If
                
            End If
            
            Debug.Print header
        
        Next column
        
        If line.TaxCode <> "" Then
            
            transaction.Lines.Add line
            
        End If
        
    Next row
    
    Set addressCollection.ShipFrom = shipFromLocation
    Set addressCollection.ShipTo = shipToLocation
    Set transaction.addresses = addressCollection
    
    Dim transactionJson(0 To 0) As Object
    Set transactionJson(0) = transaction.ToJson()
    
    Dim Json As String
    Json = JsonConverter.ConvertToJson(transactionJson)
    Debug.Print Json
    Dim ResponseJson As String
    Let ResponseJson = HTTPPost("/api/v2/transactions/create", Json, "Authorization: Basic " & UserPass, "Content-Type: application/json", "")
    
    Dim response As Variant
    Set response = JsonConverter.ParseJson(ResponseJson)
    
    If TypeName(response) = "Collection" Then
    
        GetTax = response(1)("totalTax")
    
    Else
    
        Dim errs As String
        Dim i As Integer
        Dim actualCount As Integer
        Let actualCount = 1
        
        For i = 1 To response("error")("details").Count
            
            Dim summary As String
            Let summary = response("error")("details")(i)("Summary")
            
            If summary <> "" Then
            
                Let errs = errs & actualCount & ". " & summary & vbNewLine
                Let actualCount = actualCount + 1
                
            End If
            
        Next i
        
        Err.Raise vbObjectError + 1000, Description:=(errs)
    
    End If
    
    GoTo noErr

errCantConvert:
    GetTax = "ERR!"
    MsgBox Err.Description
    
    
noErr:
    
End Function

Public Sub CalculateTax()
    
    Range("J38:J38").Value = GetTax(Range("J3").Value, Range("J4").Value, Range("A5").Value, Range("A6").Value, Range("A7").Value, Range("A8").Value, Range("C8").Value, Range("A9").Value, _
                                Range("D8").Value, Range("G12").Value, Range("G13").Value, Range("G14").Value, Range("G15").Value, Range("I15").Value, _
                                Range("G16").Value, Range("J15").Value, Range("A18:J18"), Range("A19:J36"))
    
End Sub

