VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsKeyPressEventListener"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Type MSG
    hwnd As Long
    Message As Long
    wParam As Long
    lParam As Long
    time As Long
    pt As POINTAPI
End Type

Private Declare Function WaitMessage Lib "user32" () As Long

Private Declare Function PeekMessage Lib "user32" Alias "PeekMessageA" _
    (ByRef lpMsg As MSG, ByVal hwnd As Long, _
     ByVal wMsgFilterMin As Long, _
     ByVal wMsgFilterMax As Long, _
     ByVal wRemoveMsg As Long) As Long

Private Declare Function TranslateMessage Lib "user32" _
    (ByRef lpMsg As MSG) As Long

Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" _
    (ByVal hwnd As Long, _
     ByVal wMsg As Long, _
     ByVal wParam As Long, _
     lParam As Any) As Long

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" _
    (ByVal lpClassName As String, _
     ByVal lpWindowName As String) As Long

Private Const WM_KEYDOWN As Long = &H100
Private Const PM_REMOVE  As Long = &H1
Private Const WM_CHAR    As Long = &H102
Private bExitLoop As Boolean

Public Event KeyPressed _
    (ByVal KeyAscii As Integer, _
     ByVal KeyCode As Integer, _
     ByVal Target As Range, _
     ByRef Cancel As Boolean)

Public Sub StartKeyPressInit()
    Dim msgMessage As MSG
    Dim bCancel As Boolean
    Dim iMessage As Integer
    Dim iKeyCode As Integer
    Dim lXLhwnd As Long

    On Error GoTo errHandler
    Application.EnableCancelKey = xlErrorHandler
    'Initialize this boolean flag.
    bExitLoop = False
    'Get the app hwnd.
    lXLhwnd = FindWindow("XLMAIN", Application.Caption)

    Do
        WaitMessage

        'Exit the loop if we were aborted
        If bExitLoop Then Exit Do

        'Check for a key press and remove it from the msg queue.
        If PeekMessage(msgMessage, lXLhwnd, WM_KEYDOWN, WM_KEYDOWN, PM_REMOVE) Then
            'Store the virtual key code for later use.
            iMessage = msgMessage.Message
            iKeyCode = msgMessage.wParam

            'Translate the virtual key code into a char msg.
            TranslateMessage msgMessage
            PeekMessage msgMessage, lXLhwnd, WM_CHAR, WM_CHAR, PM_REMOVE

            bCancel = False
            RaiseEvent KeyPressed(msgMessage.wParam, iKeyCode, Selection, bCancel)

            'If not handled, post back to the window using the original values
            If Not bCancel Then
                PostMessage lXLhwnd, iMessage, iKeyCode, 0
            End If
        End If
errHandler:
        'Allow the processing of other msgs.
        DoEvents
    Loop Until bExitLoop
End Sub

Public Sub StopKeyPressWatch()
    'Set this boolean flag to exit the above loop.
    bExitLoop = True
End Sub

