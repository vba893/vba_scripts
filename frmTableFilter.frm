VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmTableFilter 
   Caption         =   "UserForm1"
   ClientHeight    =   3020
   ClientLeft      =   80
   ClientTop       =   430
   ClientWidth     =   5370
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

Const MAX_FILTERS As Integer = 6
'Const SHEETNAME As String = "Sheet Name With Spaces"
Const ANY_CELL_IN_TABLE As String = "C1"

Private controlsGroup(1 To MAX_FILTERS) As clsFormControlGroup
Private columnFilters() As Collection
Private sourceWorksheet As Worksheet
Public Sub Initialize(ByRef ws As Worksheet, ByRef columnsMapping() As Collection)
    Set sourceWorksheet = ws
    
    Dim i As Integer
    ReDim columnFilters(LBound(columnsMapping) To UBound(columnsMapping))
    For i = LBound(columnsMapping) To UBound(columnsMapping)
        Set columnFilters(i) = columnsMapping(i)
    Next i
    
    Me.Caption = "Table Filter (" + ws.name + ")"
    
    ' Initialize controls groups
    For i = LBound(controlsGroup) To UBound(controlsGroup)
        Set controlsGroup(i) = CreateControlGroup(i)
    Next i
    
    For i = UBound(columnFilters) To UBound(controlsGroup)
        controlsGroup(i).CheckBox.Locked = True
        controlsGroup(i).TextBox.Locked = True
    Next i
    
    HideUnusedControls
    InitializeFiltersName
    RedimForm
End Sub

Private Sub RedimForm()
    Dim lastControlIndex As Integer
    lastControlIndex = UBound(columnFilters) + 1
    
    Const txtPrefix As String = "txtItem"
    Dim txt As Control
    
    Dim ctrl As Control
    Dim name As String
    For Each ctrl In Me.Controls
        name = (txtPrefix + CStr(lastControlIndex))
        If ctrl.name = name Then
            Set txt = ctrl
        End If
    Next ctrl

    Me.Height = txt.Top + txt.Height + 15
End Sub

Private Sub HideUnusedControls()
    Dim lastControlIndex As Integer
    lastControlIndex = UBound(columnFilters) + 1
    
    Const lblPrefix As String = "lblItem"
    Const txtPrefix As String = "txtItem"
    Const chkPrefix As String = "chkItem"
    
    Dim ctrl As Control
    Dim ctrlFound As Boolean
    Dim i As Integer
    
    For i = lastControlIndex To UBound(controlsGroup)
        For Each ctrl In Me.Controls
            ctrlFound = _
                ctrl.name = (lblPrefix + CStr(i)) Or _
                ctrl.name = (txtPrefix + CStr(i)) Or _
                ctrl.name = (chkPrefix + CStr(i))
            If ctrlFound Then
                ctrl.Visible = False
            End If
        Next ctrl
    Next i
End Sub

Private Function GetRange() As Range
    Dim r As Range
    
    ' CurrentRegion is a range bounded by any combination of blank rows and blank columns
    Set GetRange = sourceWorksheet.Range(ANY_CELL_IN_TABLE).CurrentRegion
    
End Function

Private Sub UserForm_Initialize()
    
End Sub

Private Sub InitializeFiltersName()
    Dim r As Range
    Set r = GetRange()
    Dim i, j As Integer
    
    Dim filterName As String
    For i = LBound(columnFilters) To UBound(columnFilters)
        filterName = ""
        For j = 1 To columnFilters(i).count
            If j > 1 Then
                filterName = filterName & " & "
            End If
            filterName = filterName & r.Cells(1, columnFilters(i).item(j))
        Next j
        controlsGroup(i).Label.Caption = filterName
    Next i
End Sub



Private Function CreateControlGroup(ByVal index As Integer) As clsFormControlGroup
    Set CreateControlGroup = New clsFormControlGroup
    
    Const lblPrefix As String = "lblItem"
    Const txtPrefix As String = "txtItem"
    Const chkPrefix As String = "chkItem"
    
    Dim lbl As MSForms.Label
    Dim txt As MSForms.TextBox
    Dim chk As MSForms.CheckBox
    
    Dim ctrl As Control
    Dim name As String
    For Each ctrl In Me.Controls
        If ctrl.name = (lblPrefix + CStr(index)) Then
            Set lbl = ctrl
        End If
        If ctrl.name = (txtPrefix + CStr(index)) Then
            Set txt = ctrl
        End If
        If ctrl.name = (chkPrefix + CStr(index)) Then
            Set chk = ctrl
        End If
    Next ctrl
    CreateControlGroup.Init Me, lbl, txt, chk, index
End Function

Private Function validateSheetname(ByVal SHEETNAME As String) As Boolean
    Dim sheet As Worksheet
    Dim sheetFound As Boolean
    sheetFound = False
    
    For Each sheet In ThisWorkbook.Sheets
        If sheet.name = SHEETNAME Then
            sheetFound = True
        End If
    Next sheet
        
    validateSheetname = sheetFound
    If Not sheetFound Then
        MsgBox "Sheet not found : " + SHEETNAME, vbCritical
    End If
End Function

Public Sub IFormTableFilter_TextboxChange(obj As clsFormControlGroup)
    ApplyFilter obj.TextBox.Text, obj.ControlIndex
End Sub

Public Sub IFormTableFilter_CheckboxClick(obj As clsFormControlGroup)
    ToggleCheckbox obj
End Sub

Private Sub cmdClear_Click()
    ClearFilters
End Sub

Public Sub ClearFilters()
    On Error Resume Next
    sourceWorksheet.ShowAllData
    On Error GoTo 0

    Dim i As Integer
    For i = LBound(columnFilters) To UBound(columnFilters)
        controlsGroup(i).Clear
    Next i
    
End Sub

Function collectionToArray(c As Collection) As Variant()
    Dim a() As Variant: ReDim a(1 To c.count)
    Dim i As Integer
    For i = 1 To c.count
        a(i) = c.item(i)
    Next
    collectionToArray = a
End Function

Private Sub ToggleCheckbox(ByRef obj As clsFormControlGroup)
   
    Dim filter As String
    If obj.CheckBox.Value Then
        filter = obj.TextBox.Text
    Else
        filter = ""
    End If
    
    ApplyFilter filter, obj.ControlIndex
   
End Sub

Private Sub ApplyFilter(ByVal filter As String, ByVal control_index As Integer)
    Dim i As Integer
    
    Dim r As Range
    Set r = GetRange()
    
    Dim ctrl As clsFormControlGroup
    Set ctrl = controlsGroup(control_index)
    
    Dim coll As Collection
    Set coll = columnFilters(control_index)
    
    Dim field_indexes() As Variant
    field_indexes = collectionToArray(coll)
        
    If Trim(filter) = "" Then
        On Error Resume Next
        For i = LBound(field_indexes) To UBound(field_indexes)
            r.AutoFilter Field:=field_indexes(i)
        Next i
        ctrl.CheckBox.Value = False
        On Error GoTo 0
    Else
        Dim filters() As String
        filters = Split(controlsGroup(control_index).TextBox.Text)
        RemoveEmptyField filters
        
        For i = LBound(filters) To UBound(filters)
            If filters(i) <> "" Then
                filters(i) = filters(i) & "*"
            End If
        Next i
    
        ctrl.CheckBox.Value = True
        On Error Resume Next
        For i = LBound(filters) To UBound(filters)
            r.AutoFilter Field:=field_indexes(i), Criteria1:=filters(i), Operator:=xlFilterValues
        Next i
        On Error GoTo 0
    End If
End Sub

Private Function RemoveEmptyField(ByRef arr() As String)
    Dim i As Integer
    Dim tmpcoll As Collection
    Set tmpcoll = New Collection
    
    For i = LBound(arr) To UBound(arr)
        If arr(i) <> "" Then
            tmpcoll.Add arr(i)
        End If
    Next i
    
    ReDim arr(1 To tmpcoll.count)
    For i = 1 To tmpcoll.count
        arr(i) = tmpcoll.item(i)
    Next i
    
End Function
