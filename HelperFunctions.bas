Attribute VB_Name = "HelperFunctions"
Public Declare Function GetTickCount Lib "kernel32.dll" () As Long

Sub Button1_Click()
    frmTableFilter.Show False
End Sub

Public Sub GenerateNumber()
    Dim r As Range
    Dim count As Long
    Set r = Worksheets("Sheet1").Range("A1:A100000")
    For Each c In r
        count = count + 1
        c.Value = count
    Next c
End Sub

Public Function QuickFind(ByVal item As String) As String
    Dim varray1 As Variant
    Dim varray2 As Variant
    Dim r1 As Range
    Dim r2 As Range
    
    StartTick = GetTickCount()

    Set r1 = Range("A2:A" & Cells(Rows.count, "A").End(xlUp).Row)
    Set r2 = Range("B2:B" & Cells(Rows.count, "B").End(xlUp).Row)

    varray1 = r1.Value
    varray2 = r2.Value
    
    For i = 1 To UBound(varray1, 1)
        If item = varray1(i, 1) Then
            QuickFind = varray2(i, 1)
        End If
    Next
    
    EndTick = GetTickCount()
    Debug.Print "Elapsed tick : " & Str(EndTick - StartTick)

End Function




