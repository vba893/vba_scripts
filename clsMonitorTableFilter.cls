VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMonitorTableFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private WithEvents frm As frmTableFilter
Attribute frm.VB_VarHelpID = -1

Private sourceTableWS_ As Worksheet

Public Sub Initialize(ByRef tableWorksheet As Worksheet)
    ' Associate each controls group to sheet column
    ' If multiple columns are added to the same controls group
    ' filter is splitted on space char. First part of the filter is applied to the first specified column
    ' and second part of the filter is applied to the second specified column
    Dim columnsMapping(1 To 1) As Collection
    Set columnsMapping(1) = New Collection

    columnsMapping(1).Add 1
    columnsMapping(1).Add 2

    Set sourceTableWS_ = tableWorksheet
    Set frm = New frmTableFilter
    
    frm.Initialize sourceTableWS_, columnsMapping
    frm.Show False
End Sub

Private Sub frm_AcceptSelection(ByVal selected_row As Integer)
    Dim rng As Range
    Set rng = Application.ActiveCell
    
    If selected_row = 0 Then
        Exit Sub
    End If

    rng.Value = sourceTableWS_.Cells(selected_row, 5)
    rng.Offset(1, 0).Activate
End Sub

Private Sub frm_FilterChanged(rng As Range)

End Sub
