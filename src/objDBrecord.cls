VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "objDBrecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Franck Binard, October 2017
'Class Module objDBrecord, processes and stores the information
'contained in the CSV attachment
 
Dim csvInfo As objCSVrecord
 
'BOL Database connection
Dim BOLcn As Object
Dim BOLrs As Object

'Company related variables
Dim BOLCompanyID As Long
Dim companyname As String
Dim CompStreet As String
Dim CompCity As String
Dim CompProv As String
Dim CompPostalCode As String
Dim CompPhone As String
Dim custPhoneDay

Dim BOLHeaderID As Long
Dim InvoiceNumber As String
Dim StoreNumber As String

'customer information
Dim CustomerName As String
Dim CustStreet As String
Dim custcity As String
Dim custprov As String
Dim custPostal As String


Public Sub init(ByRef csvRecord As objCSVrecord, _
    emailSender As String)
    Dim strConnection As String
    On Error GoTo initError
    Set csvInfo = csvRecord
    
     'Initialized Database connection
    strConnection = GetSetting("BOL Loader", "Settings", "DefaultODBC", _
        Const_defaultODBCName)
    Set BOLcn = CreateObject("ADODB.Connection")
    BOLcn.Open strConnection
    If Not BOLcn.State = adStateOpen Then
        errCall (1001)
    End If
    
    
    Set BOLrs = CreateObject("ADODB.Recordset")
    Call getCompanyInfo
        If Not companyInDB Then
            Call companyAdd(companyname, CompStreet, CompCity, CompProv, CompPostalCode, CompPhone)
            If Not companyInDB Then Err.Raise "Unable to add company to company directory"
        End If
        Call getBOLHeaderInfo
    Exit Sub
initError:
    Select Case Err
    Case -2147467259
    Err.Raise 1001, "objdbrecord.init", "Unable to connect to database"
    End Select
End Sub
 
Public Sub remove()
    
    BOLcn.Close
End Sub


Public Sub writeToDB()
    Dim products As objProductGroup
    Dim sku As String
    Dim skuDescription As String
    Dim qty As String
    
    Call insertBOLRecord
    Set products = New objProductGroup
    products.init (BOLHeaderID)
    With csvInfo.recordset
        .MoveFirst
        While Not .EOF
        If Not IsNull(.fields("SKU")) Then
            sku = .fields("SKU")
            skuDescription = .fields("SKU_Description")
            qty = .fields("Quantity")
            Call products.addNewProduct(sku, skuDescription, qty)
        End If
        .MoveNext
        Wend
    End With
  
    Call products.insertDBProductDetail(BOLcn)
    
    BOLrs.Open "bolinvoicedetails", BOLcn, 1, 3
    With BOLrs
        .AddNew
        .fields("bol_header_id") = BOLHeaderID
        .Update
        .Close
    End With
    
End Sub

Private Sub getCompanyInfo()
    With csvInfo.recordset
        .MoveFirst
        
        If Not IsNull(.fields("Origin_Customer_First_Name")) And _
            Not IsNull(.fields("Origin_Customer_Last_Name")) Then
            companyname = .fields("Origin_Customer_First_Name") & .fields("Origin_Customer_Last_Name")
        End If
        If Not IsNull(.fields("Origin_Street")) Then
            CompStreet = .fields("Origin_Street")
        End If
        If Not IsNull(.fields("Origin_City")) Then
            CompCity = .fields("Origin_City")
        End If
        If Not IsNull(.fields("Origin_ProvinceID")) Then
            CompProv = PID(.fields("Origin_ProvinceID"))
        End If
        If Not IsNull(.fields("Origin_Postal_Code")) Then
            CompPostalCode = formatPC(.fields("Origin_Postal_Code"))
        End If
        If Not IsNull(.fields("Origin_Phone_Number")) Then
            CompPhone = .fields("Origin_Phone_Number")
        End If
    End With
End Sub

Private Function companyInDB() As Boolean

    If Trim(CompPostalCode) = "" Then Err.Raise 1111, "Invalide company postal code"
    
    BOLrs.Open "SELECT BOLCompanyID from bolcompany " & _
    "where comppostalcode = '" & CompPostalCode & "'", BOLcn, 1, 3
    If BOLrs.EOF Then
            companyInDB = False
            BOLrs.Close
            Exit Function
    End If
    companyInDB = True
    BOLCompanyID = BOLrs.fields("BOLCompanyID")
    BOLrs.Close
End Function

Private Sub newCompanyRecord()
'inserts a new company record in BOLdb
'sets the value for BOLCompanyID
    Dim rs As Object
    Set rs = CreateObject("ADODB.Recordset")
    'Set rs = New ADODB.Recordset
    rs.Open "bolcompany", BOLcn, 1, 3
    'BOLrs.Open "bolcompany", BOLcn
    With rs
        .AddNew
       ' lngNewID = rs!MyTableID
        .fields("CompanyName").Value = companyname
        .fields("CompStreet").Value = CompStreet
        .fields("CompCity").Value = CompCity
        .fields("CompProv").Value = CompProv
        .fields("CompPostalCode").Value = CompPostalCode
        .fields("CompPhone").Value = CompPhone
        .fields("CreatedBy").Value = "Loader 1.8"
        .Update
        .Close
    End With
    Set rs = BOLcn.Execute("SELECT @@Identity", , adCmdText)
    Debug.Print rs.fields(0).Value
    BOLCompanyID = rs.fields(0).Value
    rs.Close
End Sub

Private Sub getBOLHeaderInfo()
    Dim invoiceNumberElts() As String
    Dim cFirstName As String
    Dim cLastName As String
    
  
    
    
    With csvInfo.recordset
        .MoveFirst
        If IsNull(.fields("Invoice_Number")) Then errCall (1041)
        If Not isInvoiceNumber(.fields("Invoice_Number")) Then errCall (1040)
        invoiceNumberElts = Split(.fields("Invoice_Number"), "-")
        InvoiceNumber = invoiceNumberElts(1)
        StoreNumber = invoiceNumberElts(0)
        
        If debug_on Then
            logNewEntry (InvoiceNumber)
        End If
        
    
       
        If Not IsNull(.fields("Destination_Customer_First_Name")) Then
            cFirstName = .fields("Destination_Customer_First_Name")
        End If
        If Not IsNull(.fields("Destination_Customer_Last_Name")) Then
            cLastName = .fields("Destination_Customer_Last_Name")
        End If
        CustomerName = cFirstName & " " & cLastName
        If Len(CustomerName) = 0 Then errCall 1010
        If Not IsNull(.fields("Destination_Street")) Then _
            CustStreet = .fields("Destination_Street")
        If Not IsNull(.fields("Destination_City")) Then _
            custcity = .fields("Destination_City")
        If Not IsNull(.fields("Destination_ProvinceID")) Then _
            custprov = PID(.fields("Destination_ProvinceID"))
        If Not IsNull(.fields("Destination_Postal_Code")) Then _
            custPostal = .fields("Destination_Postal_Code")
        If Not IsNull(.fields("Destination_Phone_Number")) Then _
            custPhoneDay = formatPhone(.fields("Destination_Phone_Number"))
        
    End With
End Sub

Private Sub insertBOLRecord()
    On Error GoTo insertBOLRecordError
    BOLrs.Open "bolheader", BOLcn, 1, 3
    Dim rs As ADODB.recordset
    Dim creationDate As String
       
    With BOLrs
        .AddNew
        If debug_on Then
            InvoiceNumber = "Test" & InvoiceNumber
        End If
        .fields("InvoiceNumber") = InvoiceNumber
        .fields("StoreNumber") = StoreNumber
        .fields("BOLCompanyID") = BOLCompanyID
        .fields("CustomerName") = Left(Trim(CustomerName), 35)
        .fields("CreationDate") = Now()
        Debug.Print (Len(CustStreet))
        .fields("CustStreet") = Left(CustStreet, 45)
        .fields("CustCity") = custcity
        .fields("CustProv") = custprov
        .fields("CustPostal") = custPostal
        .fields("CustPhoneDay") = custPhoneDay
        .fields("ModifiedDate").Value = Now()
        .fields("Status").Value = "Loader 1.7"
        .Update
        '.MoveLast
        'get the ID of the header just entered
        'BOLHeaderID = .fields("BOLHeaderID").Value
        .Close
    End With
   
    Set rs = BOLcn.Execute("SELECT @@Identity", , adCmdText)
    BOLHeaderID = rs.fields(0).Value
    If debug_on Then
        Debug.Print "New Bolheader record, ID:" & BOLHeaderID
        logNewEntry ("BolHeaderID|" & BOLHeaderID)
    End If
    rs.Close
    Exit Sub
insertBOLRecordError:
    Select Case Err
    Case -2147217887
        Resume Next
    Case Else:
    MsgBox "Unhandled Error:" & Err.Number
    Debug.Print "Error:" & Err.Number
    Debug.Print Err.Description
    End Select
End Sub





