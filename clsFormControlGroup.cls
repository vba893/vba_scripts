VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFormControlGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents lbl_ As MSForms.Label
Attribute lbl_.VB_VarHelpID = -1
Private WithEvents txt_ As MSForms.TextBox
Attribute txt_.VB_VarHelpID = -1
Private WithEvents chk_ As MSForms.CheckBox
Attribute chk_.VB_VarHelpID = -1

Private control_index_ As Integer
Private table_filter_ As IFormTableFilter

Public Sub Init(ByRef table_filter As IFormTableFilter, ByRef lbl As MSForms.Label, ByRef txt As MSForms.TextBox, ByRef chk As MSForms.CheckBox, ByVal control_index As Integer)
    Set table_filter_ = table_filter
    Set lbl_ = lbl
    Set txt_ = txt
    Set chk_ = chk
    
    control_index_ = control_index
End Sub

Public Property Get TextBox() As MSForms.TextBox
    Set TextBox = txt_
End Property

Public Property Get Label() As MSForms.Label
    Set Label = lbl_
End Property

Public Property Get CheckBox() As MSForms.CheckBox
    Set CheckBox = chk_
End Property

Public Property Get ControlIndex() As Integer
    ControlIndex = control_index_
End Property

Public Sub Clear()
    txt_.Text = ""
    chk_.Value = False
End Sub

Private Sub txt__Change()
    table_filter_.TextboxChange Me
End Sub

Private Sub chk__Click()
    table_filter_.CheckboxClick Me
End Sub
