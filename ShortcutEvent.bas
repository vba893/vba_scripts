Attribute VB_Name = "ShortcutEvent"
Sub CreateShortcutListener()
    Application.OnKey "^~", "CtrlEnterKeyPressed"
End Sub

Public Sub CtrlEnterKeyPressed()
    frmTableFilter.Show
    
End Sub
