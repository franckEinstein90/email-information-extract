VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "objProductGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Franck Binard, October 2017
'Class Module objProductGroup,
'processes and stores the information
'contained in the CSV attachment
'related to products

Dim BOLHeaderID As Long
Dim ProductValues As Scripting.Dictionary
Const valSeparator = "||"

Public Sub init(ByVal BOLHID As Long)
    BOLHeaderID = BOLHID
    Set ProductValues = New Scripting.Dictionary
End Sub


Public Sub addNewProduct(ByVal sku As String, ByVal skuDescription As String, ByVal qty As String)
    If ProductValues.Exists(sku) Then Err.Raise 100014, "Product already in listing"
    ProductValues(sku) = qty & valSeparator & skuDescription
End Sub

Public Sub debugOut()
    Dim key As Variant
    For Each key In ProductValues.keys
        Debug.Print key, ProductValues(key)
    Next key
End Sub

Public Sub insertDBProductDetail(ByRef BOLcn As Object)
    Dim key As Variant
    
    Dim BOLrs As Object
    Dim rs As Object
    
    Dim QTYOrdered As Integer
    Dim ProductSKU As Long
    Dim ProductDescription As String
    Dim productInfo() As String
    Dim weight As String
    Dim dimensions As String
   
    Set BOLrs = CreateObject("ADODB.Recordset")
    BOLrs.Open "boldetails", BOLcn, 1, 3
    With BOLrs
        For Each key In ProductValues.keys
            ProductSKU = CLng(key)
            Set rs = CreateObject("ADODB.recordset")
            rs.Open "SELECT weight,dimensions From Product where code='" & ProductSKU & "'", BOLcn, 1, 3
            If Not rs.EOF Then
                weight = rs.fields("weight").Value
                dimensions = rs.fields("dimensions").Value
            Else
                weight = "???"
                dimensions = "???"
            End If
            
            rs.Close
            productInfo = Split(ProductValues(key), valSeparator)
            QTYOrdered = CInt(productInfo(0))
            ProductDescription = Left(productInfo(1), 45)
            .AddNew
            .fields("BOLHeaderID").Value = BOLHeaderID
            .fields("Weight").Value = weight
            .fields("Dimension").Value = dimensions
            .fields("ProductSKU").Value = key
            .fields("QTYOrdered").Value = QTYOrdered
            .fields("ProductDescription").Value = ProductDescription
            .Update
        Next key
        .Close
    End With
    
    Set rs = BOLcn.Execute("SELECT @@Identity", , adCmdText)
    Debug.Print rs.fields(0).Value
    rs.Close
End Sub


