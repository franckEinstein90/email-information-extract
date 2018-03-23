Option Explicit

'Franck Binard Email Filing Script
'Update 19/02/2018
'Add seconds to title of file saved



Public Const garbageWords = "[Tt][Hh][Ee],[Tt][Oo],[Aa][Ss],[aA][Nn][Dd],[Oo][fF],[fF][Ww]:?,[oO][rR],[Ii][nN],[Ll][Ee],[Ll][aA],[Dd][Ee],[Dd][Uu],[Dd][Ee][s],[Ee][t],[Uu][Nn][Ee],[Oo][Uu],[Dd][Aa][Nn][Ss]"
Public Const garbageParticules = "[Ll],[Dd]"


Public Const validExtensionsFile = "C:\Garbage\01 - Work\08 - OutlookFiler\validExtensions.txt"
Public Const invalidExtensionsFile = "C:\Garbage\01 - Work\08 - OutlookFiler\invalidExtensions.txt"

Public Const defaultPath = "C:\tmp"


Public Sub processEmailV3()
    'internals directly related to email itself
    Dim myOlExp As Outlook.Explorer
    Dim myOlSel As Outlook.Selection
    Dim omail As Object
    Dim selObjectCtr As Long
    Dim strFolderPath As String
    
    Dim objAtt As Outlook.Attachment
    
    Set myOlExp = Application.ActiveExplorer
    Set myOlSel = myOlExp.Selection
    
    
    strFolderPath = browseForFolder(defaultPath)
    For selObjectCtr = 1 To myOlSel.Count
      '  If Not (myOlSel.Item(selObjectCtr).Class = OlObjectClass.olMail) _
       ' And Not (myOlSel.Item(selObjectCtr).Class = OlObjectClass.olMeetingRequest) Then
        '    MsgBox ("This script can only be applied to emails and meeting requests")
        '    GoTo NEXTSELECTION
        'End If
        Set omail = myOlSel.Item(selObjectCtr)
        Call fileEmail(omail, strFolderPath)
NEXTSELECTION:
    Next selObjectCtr
End Sub





Private Sub fileEmail(omail As Object, ByVal strFolderPath As String)

    Dim strSubject As String
    Dim strSender As String
    Dim dtDate As Date
    Dim attachmentName As String
    
    'name and path of copy
    Dim strFileName As String
    Dim fNamePrefix As String
    Dim fNameSuffix As String
    Dim iDuplicateCounter As Integer
    
    Dim objItem As Object
    Dim objAtt As Outlook.Attachment
    
    Dim sPath As String
    
    strSubject = cleanSubjectLine(omail.subject)
    strSubject = formatSubject(strSubject)
        
    Debug.Print "The file subject is now: " & strSubject
        
    strSender = omail.SenderName
    formatSender strSender, ""
        
    dtDate = omail.ReceivedTime
        
    For Each objAtt In omail.Attachments
        attachmentName = objAtt.fileName
        If isConsideredAttachment(attachmentName) Then
            fNameSuffix = "_A"
            GoTo CONSTRUCTFILENAME
            End If
        Next objAtt
                
CONSTRUCTFILENAME:
        fNamePrefix = _
            Format(dtDate, "yyyy-mm-dd hh\hnn\mss", vbUseSystemDayOfWeek, vbUseSystem) _
            & " - " & strSender & " - "
            
        If Len(fNamePrefix) + Len(fNameSuffix) + Len(strSubject) > 90 Then
            strSubject = Left(strSubject, 90 - Len(fNamePrefix) - Len(fNameSuffix))
        End If
        
        strFileName = fNamePrefix & strSubject & fNameSuffix
       

COPYEMAILTODRIVE:
    sPath = strFolderPath & "\" & strFileName
    'iDuplicateCounter = 1
    'Do While Dir(sPath & ".msg") <> ""
     '   If iDuplicateCounter = 1 Then
      '      sPath = sPath & iDuplicateCounter
      '  Else
     '       sPath = Left(sPath, Len(sPath) - 1) & iDuplicateCounter
      '  End If
      '  iDuplicateCounter = iDuplicateCounter + 1
    'Loop
    sPath = sPath & ".msg"
    Debug.Print "The file will be saved as: " & sPath
    omail.SaveAs sPath, olMSG

QUERYFORDELETE:
        If askForDelete() Then
            omail.Delete
        End If
End Sub

Private Function isConsideredAttachment(ByVal fileName As String) As Boolean
    Dim extension As String
    Dim txtLine As String
    Dim a As Integer
    
    
    extension = Right(fileName, Len(fileName) - InStrRev(fileName, "."))
    If Dir(validExtensionsFile) = "" Then
        MsgBox "Unable to find valid extension file at path" & validExtensionsFile
        a = MsgBox("Would you like to mark that this file has an attachment in the saved version?", vbYesNo)
        If a = 6 Then
            isConsideredAttachment = True
        Else
            isConsideredAttachment = False
        End If
        GoTo CLEAN_AND_RETURN
    End If
    Open validExtensionsFile For Input As #1
    Do Until EOF(1)
        Line Input #1, txtLine
        If UCase(txtLine) = UCase(extension) Then
            isConsideredAttachment = True
            Close #1
            GoTo CLEAN_AND_RETURN
        End If
    Loop
    Close #1
    
    If Dir(invalidExtensionsFile) = "" Then
        MsgBox "Unable to find invalid extension file at path" & invalidExtensionsFile
        a = MsgBox("Would you like to mark that this file has an attachment in the saved version?", vbYesNo)
        If a = 6 Then
            isConsideredAttachment = True
        Else
            isConsideredAttachment = False
        End If
        GoTo CLEAN_AND_RETURN
    End If
    
    Open invalidExtensionsFile For Input As #2
    Do Until EOF(2)
        Line Input #2, txtLine
        If UCase(txtLine) = UCase(extension) Then
            isConsideredAttachment = False
            Close #2
            GoTo CLEAN_AND_RETURN
        End If
    Loop
    Close #2
    
    
   'Message Box with title, yes no and cancel Butttons
    a = MsgBox("There is no rule for attachments of type " & extension & " - Would you like to flag emails that contains attachment of that type?", vbYesNo)
    If a = 6 Then
      Open validExtensionsFile For Append As #1
      Print #1, extension
      Close #1
      isConsideredAttachment = True
    Else
        Open invalidExtensionsFile For Append As #2
        Print #2, extension
        Close #2
        isConsideredAttachment = False
    End If
    
CLEAN_AND_RETURN:
End Function




Public Function askForDelete() As Boolean
    Dim a As Integer
   
   'Message Box with title, yes no and cancel Butttons
   a = MsgBox("Delete this message?", vbYesNo)
    If a = 6 Then
        askForDelete = True
    Else
        askForDelete = False
    End If
End Function

Function browseForFolder(Optional OpenAt As Variant) As Variant
  Dim ShellApp As Object
  Set ShellApp = CreateObject("Shell.Application").browseForFolder(0, "Please choose a folder", 0, OpenAt)
  
 On Error Resume Next
    browseForFolder = ShellApp.self.Path
On Error GoTo 0
  
 Set ShellApp = Nothing
    Select Case Mid(browseForFolder, 2, 1)
        Case Is = ":"
            If Left(browseForFolder, 1) = ":" Then GoTo Invalid
        Case Is = "\"
            If Not Left(browseForFolder, 1) = "\" Then GoTo Invalid
        Case Else
            GoTo Invalid
    End Select
Exit Function
  
Invalid:
browseForFolder = False
End Function

Public Function cleanSubjectLine(ByVal subjectLine As String) As String
    Dim wArray() As String
    Dim particuleArray() As String
    
    Dim wordKey As Variant
    cleanSubjectLine = subjectLine
    wArray = Split(garbageWords, ",")
    particuleArray = Split(garbageParticules, ",")
    
    For Each wordKey In wArray
        cleanSubjectLine = removeBadWord(cleanSubjectLine, wordKey)
    Next wordKey
    For Each wordKey In particuleArray
        cleanSubjectLine = removeBadParticule(cleanSubjectLine, wordKey)
    Next wordKey
    cleanSubjectLine = finalClean(cleanSubjectLine)
End Function

Private Function removeBadWord(ByVal subject As String, ByVal strToRemove) As String

    
    Dim strPattern As String: strPattern = "(\s+|^|\()" & strToRemove & "\s+|" & "\s+" & strToRemove & "$"
    Dim strReplace As String: strReplace = " "
    Dim regEx As New RegExp
    
    
    With regEx
            .Global = True
            .MultiLine = True
            .IgnoreCase = False
            .Pattern = strPattern
    End With
        
    removeBadWord = regEx.Replace(subject, strReplace)
    
End Function

Private Function removeBadParticule(ByVal subject As String, ByVal strToRemove) As String

    
    Dim strPattern As String: strPattern = "(\s+|^)" & strToRemove & "'"
    Dim strReplace As String: strReplace = " "
    Dim regEx As New RegExp
    
    
    With regEx
            .Global = True
            .MultiLine = True
            .IgnoreCase = False
            .Pattern = strPattern
    End With
        
    removeBadParticule = regEx.Replace(subject, strReplace)
    
End Function
   
 Private Function finalClean(ByVal subject As String) As String

    
    Dim strPattern As String: strPattern = "\s+"
    Dim strReplace As String: strReplace = " "
    Dim regEx As New RegExp
    
    
    With regEx
            .Global = True
            .MultiLine = True
            .IgnoreCase = False
            .Pattern = strPattern
    End With
        
    finalClean = regEx.Replace(subject, strReplace)
    
End Function
      
    
Public Function formatSubject(ByVal subject As String) As String

Dim sChr As String

sChr = ""
formatSubject = Replace(subject, "'", sChr)
formatSubject = Replace(formatSubject, "*", sChr)
formatSubject = Replace(formatSubject, "/", sChr)
formatSubject = Replace(formatSubject, "\", sChr)
formatSubject = Replace(formatSubject, ":", sChr)
formatSubject = Replace(formatSubject, "?", sChr)
formatSubject = Replace(formatSubject, Chr(34), sChr)
formatSubject = Replace(formatSubject, "<", sChr)
formatSubject = Replace(formatSubject, ">", sChr)
formatSubject = Replace(formatSubject, "|", sChr)
formatSubject = Replace(formatSubject, "(", sChr)
formatSubject = Replace(formatSubject, ")", sChr)
formatSubject = Replace(formatSubject, ".", sChr)
formatSubject = Replace(formatSubject, "!", sChr)
formatSubject = Replace(formatSubject, ",", sChr)
formatSubject = Replace(formatSubject, "…", sChr)
formatSubject = Replace(formatSubject, "'", sChr)
End Function


Private Sub formatSender(sSenderName As String, _
sChr As String _
)
sSenderName = Replace(sSenderName, "/", sChr)
sSenderName = Replace(sSenderName, "\", sChr)
sSenderName = Replace(sSenderName, ":", sChr)
sSenderName = Replace(sSenderName, Chr(34), sChr)
sSenderName = Replace(sSenderName, "<", sChr)
sSenderName = Replace(sSenderName, ">", sChr)
sSenderName = Replace(sSenderName, "|", sChr)
sSenderName = Replace(sSenderName, "HCSC", sChr)
sSenderName = Replace(sSenderName, "(", sChr)
sSenderName = Replace(sSenderName, ")", sChr)

 
End Sub




