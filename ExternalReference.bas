Attribute VB_Name = "ExternalReference"


'Credits: Ken Puls
Sub AddReference()
     'Macro purpose:  To add a reference to the project using the GUID for the
     'reference library

    Dim strGUID As String, theRef As Variant, i As Long

     'Update the GUID you need below.
    strGUID = "{00020905-0000-0000-C000-000000000046}"

     'Set to continue in case of error
    On Error Resume Next

     'Remove any missing references
    For i = ThisWorkbook.VBProject.References.count To 1 Step -1
        Set theRef = ThisWorkbook.VBProject.References.item(i)
        If theRef.IsBroken = True Then
            ThisWorkbook.VBProject.References.Remove theRef
        End If
    Next i

     'Clear any errors so that error trapping for GUID additions can be evaluated
    Err.Clear

     'Add the reference
    ThisWorkbook.VBProject.References.AddFromGuid _
    GUID:=strGUID, Major:=1, Minor:=0

     'If an error was encountered, inform the user
    Select Case Err.Number
    Case Is = 32813
         'Reference already in use.  No action necessary
    Case Is = vbNullString
         'Reference added without issue
    Case Else
         'An unknown error was encountered, so alert the user
        MsgBox "A problem was encountered trying to" & vbNewLine _
        & "add or remove a reference in this file" & vbNewLine & "Please check the " _
        & "references in your VBA project!", vbCritical + vbOKOnly, "Error!"
    End Select
    On Error GoTo 0
End Sub
Private Function AddScriptingLibrary() As Boolean
    ' Add Microsoft Scripting Runtime
    Const GUID As String = "{420B2830-E718-11CF-893D-00A0C9054228}"
    
    AddReferenceByGUID (GUID)
End Function

Private Sub AddRefGuid()
    'Add VBIDE (Microsoft Visual Basic for Applications Extensibility 5.3)
    Const GUID As String = "{0002E157-0000-0000-C000-000000000046}"
    
    AddReferenceByGUID (GUID)
End Sub

Public Sub AddReferenceByGUID(ByVal GUID As String)
    On Error GoTo errHandler
    ThisWorkbook.VBProject.References.AddFromGuid GUID, 1, 0
    
errHandler:
    MsgBox Err.Description
End Sub

' Credits: https://social.msdn.microsoft.com/Forums/azure/en-US/57813453-9a21-4080-9d4a-e548e715d7ca/add-visual-basic-extensibility-library-through-code?forum=isvvba
Public Sub ListProjectReferences()
     'Macro purpose:  To determine full path and Globally Unique Identifier (GUID)
     'to each referenced library.  Select the reference in the Tools\References
     'window, then run this code to get the information on the reference's library
    Dim i As Long
   
    For i = 1 To ThisWorkbook.VBProject.References.count
        With ThisWorkbook.VBProject.References(i)
            Debug.Print "--------------------------------------------"
            Debug.Print .Description
            Debug.Print "     GUID  : " & .GUID
            Debug.Print "     Name  : " & .name
            Debug.Print "  Fullpath : " & .FullPath
        End With
    Next i
End Sub

