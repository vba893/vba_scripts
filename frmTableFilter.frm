VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmTableFilter 
   Caption         =   "UserForm1"
   ClientHeight    =   2310
   ClientLeft      =   0
   ClientTop       =   130
   ClientWidth     =   7280
   OleObjectBlob   =   "frmTableFilter.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmTableFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Implements IFormTableFilter

Private controlsGroup(1 To 4) As clsTableFilter
Attribute controlsGroup.VB_VarHelpID = -1
Const sheetName As String = "Sheet Name With Spaces"

Private Function GetRange() As Range
    Dim ws As Worksheet
    Dim r As Range
    Set ws = Worksheets(sheetName)
    Set GetRange = ws.Range("B1:E1")
End Function

Public Sub IFormTableFilter_TextboxChange(txt As MSForms.TextBox, ByVal control_index As Integer, ByVal column_index As Integer)
    ApplyFilter txt.Text, control_index, column_index
End Sub

Public Sub IFormTableFilter_CheckboxClick(chk As MSForms.CheckBox, ByVal control_index As Integer, ByVal column_index As Integer)
    ToggleCheckbox control_index
End Sub

Private Sub UserForm_Initialize()
    Me.Caption = "Table Filter (" + sheetName + ")"
    
    Dim isSheetNameValid As Boolean
    validateSheetname sheetName
    If Not isSheetNameValid Then
        Exit Sub
    End If

    Dim i As Integer
    For i = LBound(controlsGroup) To UBound(controlsGroup)
        Set controlsGroup(i) = New clsTableFilter
        controlsGroup(i).control_index_ = i
    Next i
    
    controlsGroup(1).Init Me, lblItem1, txtItem1, chkItem1, 1
    controlsGroup(2).Init Me, lblItem2, txtItem2, chkItem2, 2
    controlsGroup(3).Init Me, lblItem3, txtItem3, chkItem3, 3
    controlsGroup(4).Init Me, lblItem4, txtItem4, chkItem4, 4
   
    Dim r As Range
    Set r = GetRange()
    For i = LBound(controlsGroup) To UBound(controlsGroup)
        controlsGroup(i).lbl_.Caption = r.Cells(1, controlsGroup(i).column_index_)
    Next i
        
End Sub

Private Function validateSheetname(ByVal sheetName As String) As Boolean
    Dim sheet As Worksheet
    Dim sheetFound As Boolean
    sheetFound = False
    
    For Each sheet In ThisWorkbook.Sheets
        If sheet.Name = sheetName Then
            sheetFound = True
        End If
    Next sheet
    
    If Not sheetFound Then
        MsgBox "Sheet not found : " + sheetName, vbCritical
    End If
End Function

Private Sub cmdClear_Click()
    ClearFilters
End Sub

Public Sub ClearFilters()
    Dim ws As Worksheet
    Set ws = Worksheets(sheetName)

    On Error Resume Next
    ws.ShowAllData
    On Error GoTo 0
    
    Dim i As Integer
    For i = LBound(controlsGroup) To UBound(controlsGroup)
        controlsGroup(i).Clear
    Next i
    
End Sub

Private Sub ToggleCheckbox(ByVal control_index As Integer)
    Dim column_index As Integer
    column_index = controlsGroup(control_index).column_index_
    
    If controlsGroup(control_index).chk_.Value = False Then
        ApplyFilter "", control_index, column_index
    Else
        ApplyFilter controlsGroup(control_index).txt_.Text, control_index, column_index
    End If
End Sub

Private Sub ApplyFilter(ByVal filter As String, ByVal control_index, ByVal field_index As Integer)
    Dim r As Range
    Set r = GetRange()
    
    If Trim(filter) = "" Then
        On Error Resume Next
        r.AutoFilter Field:=field_index
        controlsGroup(control_index).chk_.Value = False
        On Error GoTo 0
    Else
        filter = "*" & filter & "*"
        controlsGroup(control_index).chk_.Value = True
        On Error Resume Next
        r.AutoFilter Field:=field_index, Criteria1:=filter, Operator:=xlFilterValues
        On Error GoTo 0
    End If
End Sub

