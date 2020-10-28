Attribute VB_Name = "ShortcutEvent"
Private monitor As clsMonitorTableFilter

Sub CreateShortcutListener()
    Application.OnKey "^~", "CtrlEnterKeyPressed"
End Sub

Public Sub CtrlEnterKeyPressed()
    Set monitor = New clsMonitorTableFilter
    monitor.Initialize Application.Worksheets("Sheet Name With Spaces")
End Sub

