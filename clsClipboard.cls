VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsClipboard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
#If Mac Then
    ' do nothing
#Else
    #If VBA7 Then
        Private Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
        Private Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
        Private Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As LongPtr) As LongPtr

        Private Declare PtrSafe Function CloseClipboard Lib "User32" () As Long
        Private Declare PtrSafe Function OpenClipboard Lib "User32" (ByVal hwnd As LongPtr) As LongPtr
        Private Declare PtrSafe Function EmptyClipboard Lib "User32" () As Long

        Private Declare PtrSafe Function lstrcpy Lib "kernel32" (ByVal lpString1 As Any, ByVal lpString2 As Any) As LongPtr

        Private Declare PtrSafe Function SetClipboardData Lib "User32" (ByVal wFormat As Long, ByVal hMem As LongPtr) As LongPtr
    #Else
        Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
        Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
        Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long

        Private Declare Function CloseClipboard Lib "User32" () As Long
        Private Declare Function OpenClipboard Lib "User32" (ByVal hwnd As Long) As Long
        Private Declare Function EmptyClipboard Lib "User32" () As Long

        Private Declare Function lstrcpy Lib "kernel32" (ByVal lpString1 As Any, ByVal lpString2 As Any) As Long

        Private Declare Function SetClipboardData Lib "User32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
    #End If
#End If

Const GHND = &H42
Const CF_TEXT = 1
Const MAXSIZE = 4096

Public Sub ClipBoard_SetData(MyString As String)
    #If Mac Then
        With New MSForms.DataObject
            .SetText MyString
            .PutInClipboard
        End With
    #Else
        #If VBA7 Then
            Dim hGlobalMemory As LongPtr
            Dim hClipMemory   As LongPtr
            Dim lpGlobalMemory    As LongPtr
        #Else
            Dim hGlobalMemory As Long
            Dim hClipMemory   As Long
            Dim lpGlobalMemory    As Long
        #End If

        Dim x                 As Long

        ' Allocate moveable global memory.
       '-------------------------------------------
       hGlobalMemory = GlobalAlloc(GHND, Len(MyString) + 1)

        ' Lock the block to get a far pointer
       ' to this memory.
       lpGlobalMemory = GlobalLock(hGlobalMemory)

        ' Copy the string to this global memory.
       lpGlobalMemory = lstrcpy(lpGlobalMemory, MyString)

        ' Unlock the memory.
       If GlobalUnlock(hGlobalMemory) <> 0 Then
            MsgBox "Could not unlock memory location. Copy aborted."
            GoTo PrepareToClose
        End If

        ' Open the Clipboard to copy data to.
       If OpenClipboard(0&) = 0 Then
            MsgBox "Could not open the Clipboard. Copy aborted."
            Exit Sub
        End If

        ' Clear the Clipboard.
       x = EmptyClipboard()

        ' Copy the data to the Clipboard.
       hClipMemory = SetClipboardData(CF_TEXT, hGlobalMemory)

PrepareToClose:

        If CloseClipboard() = 0 Then
            MsgBox "Could not close Clipboard."
        End If
    #End If

End Sub

