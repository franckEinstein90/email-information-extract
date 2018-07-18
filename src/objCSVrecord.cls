VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "objCSVrecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Franck Binard, this module extracts the information
'from a csv file for processing


Dim fieldDict As Object
Dim csvRS As ADODB.recordset

Property Get recordset() As ADODB.recordset
    Set recordset = csvRS
End Property



Public Sub init(ByRef csvAttachment As Outlook.Attachment, ByVal bolnumber As String)
    Dim strTempCopyPath As String
    Dim key As Variant
    strTempCopyPath = copyFileToTempLocation(csvAttachment, bolnumber)
    
    Set fieldDict = CreateObject("Scripting.Dictionary")
    Set csvRS = extractEmailAttachementInformation(strTempCopyPath, fieldDict)
    
    csvRS.MoveFirst
    If csvRS.EOF Then
        Err.Raise 100014, , "Invalide BOL information"
    End If
    
    If debug_on Then
    With csvRS
        .MoveFirst
        While Not .EOF
            For Each key In fieldDict.keys
                Debug.Print .fields(key)
            Next key
            .MoveNext
        Wend
        .MoveFirst
    End With
    End If
End Sub


Private Function extractEmailAttachementInformation(ByVal filePath As String, fieldDict As Object) As ADODB.recordset
  
    Dim directory As String
    Dim fileName As String
   
    directory = Left(filePath, InStrRev(filePath, "\"))
    fileName = Right(filePath, Len(filePath) - InStrRev(filePath, "\"))
   
    Set extractEmailAttachementInformation = readCSVInformation(directory, fileName, fieldDict)
End Function

Private Sub readCSVHeaderFile(ByVal directory As String, ByVal fileName As String, ByRef headerDict As Object)
    Dim headerLine As String
    Dim fields() As String
    Dim fileNo As Integer
    Dim fieldCtr As Long
    
    fileNo = FreeFile 'Get first free file number
    Open directory & "\" & fileName For Input As #fileNo
    Line Input #fileNo, headerLine
    Close #fileNo
    
    fields = Split(headerLine, ",")
    For fieldCtr = 0 To UBound(fields)
         headerDict.Add key:=fields(fieldCtr), Item:=fieldCtr
    Next fieldCtr
End Sub

Private Function readCSVInformation(ByVal directory As String, ByVal fileName As String, fieldDict As Object) As ADODB.recordset
    'read the information contained in an attached csv files and returns an ADODB recordset
    'with the corresponding information
    Dim rs As ADODB.recordset
    Dim connectionString As String
    Dim strSql As String

    Dim fieldnum As Variant
    
    On Error GoTo readCSVInformationError
    
    '1. check that file exists
    If Dir(directory & "\" & fileName) = "" Then
        errCall (1031)
    End If
    
    '2. get file headers and store with corresponding indexes in fieldDict
    Call readCSVHeaderFile(directory, fileName, fieldDict)
    
    With fieldDict
        'For Each fieldnum In fieldDict.keys
            'Debug.Print fieldnum & " - " & .Item(fieldnum)
        'Next fieldnum
    End With
    
    '3. read the rest of the information using recordset
    connectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & directory & ";" & "Extended Properties=""text;HDR=Yes;FMT=Delimited"";"
    strSql = "SELECT * FROM [" & fileName & "]"
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open strSql, connectionString, 3, 3
    Set readCSVInformation = rs
    Exit Function
readCSVInformationError:
    Debug.Print "Error: " & Err.Number & "::" & Err.Description
End Function

