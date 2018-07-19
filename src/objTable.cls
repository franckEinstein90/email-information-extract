VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "objTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Public fieldDicts As Scripting.Dictionary

Public Sub debugToStd()
    Dim key As Variant
    For Each key In fieldDicts.keys
        Debug.Print key, fieldDicts(key)
    Next key
End Sub

Public Sub init(ByVal fieldStr As String)
    Dim fields() As String
    Dim ctr As Integer
    Set fieldDicts = New Scripting.Dictionary
    
    fields = Split(fieldStr, "|")
    For ctr = 0 To UBound(fields)
        fieldDicts.Add ctr, fields(ctr)
    Next ctr
    
End Sub

