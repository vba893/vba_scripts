VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTableFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents lbl_ As MSForms.Label
Attribute lbl_.VB_VarHelpID = -1
Public WithEvents txt_ As MSForms.TextBox
Attribute txt_.VB_VarHelpID = -1
Public WithEvents chk_ As MSForms.CheckBox
Attribute chk_.VB_VarHelpID = -1
Private table_filter_ As IFormTableFilter
Attribute table_filter_.VB_VarHelpID = -1

Public control_index_ As Integer
Public column_index_ As Integer

Public Sub Init(ByRef table_filter As IFormTableFilter, ByRef lbl As MSForms.Label, ByRef txt As MSForms.TextBox, ByRef chk As MSForms.CheckBox, ByVal column_index As Integer)
    Set table_filter_ = table_filter
    Set lbl_ = lbl
    Set txt_ = txt
    Set chk_ = chk
    
    column_index_ = column_index
End Sub

Public Sub Clear()
    txt_.Text = ""
    chk_.Value = False
End Sub

Private Sub txt__Change()
    table_filter_.TextboxChange txt_, control_index_, column_index_
End Sub

Private Sub chk__Click()
    table_filter_.CheckboxClick chk_, control_index_, column_index_
End Sub
