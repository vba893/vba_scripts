VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsReplaceMessageFields"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private ws As Worksheet
Private fieldNames As Collection
Private fieldColumnIndexes As Collection

Public Function ParseTemplateMessage(ByRef wsSourceFiels As Worksheet, ByVal message As String, ByVal rowSourceValues As Integer) As String
    Set ws = wsSourceFiels
    Set fieldNames = New Collection
    Set fieldColumnIndexes = New Collection
    
    getFieldNames
    
    ParseTemplateMessage = replaceFieldsWithValues(message, rowSourceValues)
End Function

Private Function getFieldNames() As String()
    Dim c As Range
    For Each c In ws.UsedRange.Columns
        fieldNames.Add CStr(c.Value2(1, 1))
        fieldColumnIndexes.Add c.Column
    Next c
End Function

Private Function replaceFieldsWithValues(ByVal message As String, ByVal rowSourceValues As Integer) As String
    Dim i As Integer
    Dim value As String
    Dim cell As Range
    
    Dim regEx As New RegExp
    regEx.MultiLine = True
    regEx.IgnoreCase = True
    
    For i = 1 To fieldNames.Count
        regEx.pattern = "<" & fieldNames.Item(i) & ">"
        Set cell = ws.UsedRange.cells(rowSourceValues, fieldColumnIndexes.Item(i))
        value = CStr(Trim(cell.text))
        
        message = regEx.Replace(message, value)
    Next i
    
    message = Replace(message, "  ", " ")
    
    replaceFieldsWithValues = message
End Function

Private Sub debugFieldNames()
    Dim i As Integer
    For i = 1 To fieldNames.Count
        Debug.Print "Index : " & fieldColumnIndexes.Item(i) & " > Name : " & fieldNames.Item(i)
    Next i
End Sub

