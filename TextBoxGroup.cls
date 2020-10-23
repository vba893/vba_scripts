VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TextBoxGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents txt As MSForms.TextBox
Attribute txt.VB_VarHelpID = -1

Public Event Change(ByVal index As Integer)

Private index As Integer

Sub SetTextBox(ctl As MSForms.TextBox)
    Static instance_index As Integer
    instance_index = instance_index + 1
    index = instance_index
    Set txt = ctl
End Sub

Private Sub txt_Change()
    raisevent Change(index)
End Sub

