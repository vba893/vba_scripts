Option Explicit

Public WithEvents lbl_ As MSForms.Label
Public WithEvents txt_ As MSForms.TextBox
Public WithEvents chk_ As MSForms.CheckBox
Private frm_ As UserForm1

Public control_index_ As Integer
Public column_index_ As Integer

Public Sub Init(ByRef frm As UserForm1, ByRef lbl As MSForms.Label, ByRef txt As MSForms.TextBox, ByRef chk As MSForms.CheckBox, ByVal column_index As Integer)
    Set frm_ = frm
    Set lbl_ = lbl
    Set txt_ = txt
    Set chk_ = chk
    
    column_index_ = column_index
End Sub

Public Sub Clear()
    txt_.Text = ""
    chk_.Value = False
End Sub

Private Sub Class_Initialize()
    Static instances As Integer
    instances = instances + 1
    control_index_ = instances
End Sub

Private Sub txt__Change()
    frm_.TextboxChange txt_, control_index_, column_index_
End Sub

Private Sub chk__Click()
    frm_.CheckboxClick chk_, control_index_, column_index_
End Sub
