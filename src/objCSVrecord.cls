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
'Franck Binard - 2018
'Module to extract and read information
'from a csv file for processing

Private Type TCSVRecordSet
    fieldDict As Object
    csvRs As ADODB.recordset
End Type

Dim this As TCSVRecordSet

Property Get recordset() As ADODB.recordset
    Set recordset = this.csvRs
End Property


Private Sub debugPrintFieldValuePairs()
    Dim key As Variant
    With this.csvRs
        .MoveFirst
        While Not .EOF
            For Each key In this.fieldDict.keys
                Debug.Print key
                Debug.Print this.fieldDict.fields(key)
            Next key
            .MoveNext
        Wend
        .MoveFirst
    End With
End Sub


Private Function copyFile( _
        objAtt As Outlook.Attachment, _
        ByVal fileNameWithPath As String) _
    As String
    
    'copies the attachment to location specified and
    'returns the full path and name of the file
    objAtt.SaveAsFile fileNameWithPath
    copyFile = fileNameWithPath
    'logNewEntry (copyFileToTempLocation)
End Function


Public Sub init( _
    ByRef csvAttachment As Outlook.Attachment, _
    ByVal copyNameWithPath As String)
    
    Dim copyFilePath As String: copyFilePath = copyFile(csvAttachment, copyNameWithPath)
    
    With this
        Set .fieldDict = CreateObject("Scripting.Dictionary")
        Set .csvRs = extractEmailAttachementInformation(copyFilePath, .fieldDict)
        .csvRs.MoveFirst
        If .csvRs.EOF Then
            Err.Raise 100014, , "Invalide BOL information"
        End If
    End With
    Call debugPrintFieldValuePairs
  
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

Private Function readCSVInformation( _
    ByVal directory As String, _
    ByVal fileName As String, _
    fieldDict As Object) As ADODB.recordset
    
    'reads information in an attached csv file and
    'returns an ADODB recordset with corresponding information
    
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

