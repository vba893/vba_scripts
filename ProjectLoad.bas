Attribute VB_Name = "ProjectLoad"
Option Explicit
Public Sub LoadProject()
    ' OpenDialog
    Dim strFolderExists As String
    Dim fd As Office.FileDialog
    Dim importPath As String, myFile As String
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    
    ' Check if all references exists
    AddAllExternalReferences
        
    ' Select starting folder
    Dim userDocuments As String
    If strFolderExists = "" Then
        userDocuments = Environ("OneDriveCommercial") & "\Documents"
        strFolderExists = Dir(userDocuments, vbDirectory)
    End If
    If strFolderExists = "" Then
        userDocuments = Environ("USERPROFILE") & "\Documents"
        strFolderExists = Dir(userDocuments, vbDirectory)
    End If
    If strFolderExists = "" Then
        userDocuments = "C:\"
    End If
    
    With fd
        .Filters.Clear
        .Filters.Add "Config Files", "*.cfg", 1
        .Title = "Choose a project config file"
        .AllowMultiSelect = False
     
        .InitialFileName = userDocuments
     
        If .Show = True Then
            myFile = .SelectedItems(1)
        End If
    End With
    
    ' Exit if no file selected or user canceled
    If fd.SelectedItems.Count = 0 Then Return
    
    importPath = GetDirectoryName(myFile)
    
    ' ReadFileContent
    Dim lines() As String
    lines = GetFileContent(myFile)
    
    ' Import VBA files or execute command
    ImportFileOrExecuteCommand importPath, lines
    
    MsgBox "Project '" & myFile & "' loaded !", vbInformation
End Sub

Private Function GetDirectoryName(ByVal fullname As String) As String
    Dim objFSO As Scripting.FileSystemObject
    Set objFSO = New Scripting.FileSystemObject
    
    GetDirectoryName = objFSO.GetParentFolderName(fullname)
End Function

Private Function ImportFileOrExecuteCommand(ByVal importPath As String, ByRef lines() As String)
    ' Import VBA files
    Dim i As Integer
    Dim line As String
    For i = LBound(lines) To UBound(lines)
        line = lines(i)
        If IsCommandLine(line) Then
            ' Execute specified commands
            Application.Run ExtractCommand(line)
        Else
            ' Import the specified file
            ImportModule importPath, line
        End If
    Next i
End Function

Private Function IsCommandLine(ByVal line As String) As Boolean
    Const command_prefix As String = "Call"
    Dim line_prefix As String
    line_prefix = Mid(line, 1, Len(command_prefix))
    
    IsCommandLine = StringEquals(command_prefix, line_prefix)
End Function

Private Function ExtractCommand(ByVal line As String) As String
    Const command_prefix As String = "Call"
    Dim cmd As String
    If Not IsCommandLine(line) Then Exit Function
    
    cmd = Trim(Mid(line, Len(command_prefix) + 1))
    
    ExtractCommand = cmd
End Function

Private Function StringEquals(ByVal str1 As String, ByVal str2 As String) As Boolean
     StringEquals = UCase(str1) = UCase(str2)
End Function

Private Function GetFileContent(ByVal filename As String) As String()
    ' ReadFileContent
    Dim lineCount As Integer, line As String, lines() As String
    
    ReDim lines(1 To 65535)
    lineCount = 0
    Open filename For Input As #1
    Do Until EOF(1)
        Line Input #1, line
        If Trim(line & vbNullString) <> vbNullString Then
            lineCount = lineCount + 1
            lines(lineCount) = line
        End If
    Loop
    Close #1
    ReDim Preserve lines(1 To lineCount)
    
    GetFileContent = lines
End Function

Private Sub ImportModule(ByVal import_path As String, ByVal filename As String)
    Dim wkbTarget As Excel.Workbook
    Dim szTargetWorkbook As String
    Dim szImportPath As String
    Dim szImportFile As String
    Dim szFileName As String
    Dim moduleAlreadyImported As Boolean
    
    ' Early binding
    Dim objFSO As Scripting.FileSystemObject
    Dim objFile As Scripting.File
    Dim cmpComponents As VBIDE.VBComponents
    Set wkbTarget = Application.ThisWorkbook
    Set cmpComponents = wkbTarget.VBProject.VBComponents
    Set objFSO = New Scripting.FileSystemObject
    
    ''' NOTE: This workbook must be open in Excel.
    szTargetWorkbook = ThisWorkbook.Name
    
    If wkbTarget.VBProject.Protection = 1 Then
    MsgBox "The VBA in this workbook is protected," & _
        "not possible to Import the code"
    Exit Sub
    End If

    ''' NOTE: Path where the code modules are located.
    szImportPath = import_path & "\"
    szImportFile = szImportPath & filename
    
    ' Check if file exist
    If Not objFSO.FileExists(szImportFile) Then
        MsgBox "The specified file '" & szImportFile & "' was not found", vbExclamation
        Exit Sub
    End If
    
    ' Load file content to get name attribut
    Dim moduleName As String
    moduleName = GetModuleName(szImportFile)
    
    ' Check if module name is already loaded
    moduleAlreadyImported = IsModuleAlreadyImported(moduleName)
    If moduleAlreadyImported Then
        MsgBox "Module name '" & moduleName & "' is already in project", vbInformation
        Exit Sub
    End If
    
    If (objFSO.GetExtensionName(szImportFile) = "cls") Or _
        (objFSO.GetExtensionName(szImportFile) = "frm") Or _
        (objFSO.GetExtensionName(szImportFile) = "bas") Then
        cmpComponents.Import szImportFile
    End If
End Sub

Private Function GetModuleName(ByVal filename As String) As String
    ' ReadFileContent
    Dim lines() As String
    lines = GetFileContent(filename)
    
    ' Import VBA files
    Dim i As Integer
    Dim line As String
    Const namePrefix As String = "Attribute VB_Name = "
    For i = LBound(lines) To UBound(lines)
        line = lines(i)
        If InStr(1, line, namePrefix, vbTextCompare) > 0 Then
            GetModuleName = Mid(line, Len(namePrefix) + 1)
            GetModuleName = Replace(GetModuleName, Chr(34), "")
            Exit For
        End If
    Next i
    
End Function

Private Function IsModuleAlreadyImported(ByVal moduleName As String) As Boolean
    Dim wb As Excel.Workbook
    Set wb = Application.ThisWorkbook
    Dim szFileName As String
    Dim cmpComponent As VBIDE.VBComponent
    
    For Each cmpComponent In wb.VBProject.VBComponents
        szFileName = cmpComponent.Name

        ''' Concatenate the correct filename for export.
        Select Case cmpComponent.Type
            Case vbext_ct_ClassModule
                szFileName = szFileName & ".cls"
            Case vbext_ct_MSForm
                szFileName = szFileName & ".frm"
            Case vbext_ct_StdModule
                szFileName = szFileName & ".bas"
            Case vbext_ct_Document

        End Select
        
        If cmpComponent.Name = moduleName Then
            IsModuleAlreadyImported = True
            Exit For
        End If
   
    Next cmpComponent

End Function

'******************************************************************************
'AddLib: Adds a library reference to this script programmatically, so that
'        libraries do not need to be added manually.
'******************************************************************************
Private Function AddReference(libName As String, GUID As String, major As Long, minor As Long)

    Dim exObj As Object: Set exObj = GetObject(, "Excel.Application")
    Dim vbProj As Object: Set vbProj = exObj.ActiveWorkbook.VBProject
    Dim chkRef As Object

    ' Check if the library has already been added
    For Each chkRef In vbProj.References
        Debug.Print chkRef.Name
        If chkRef.Name = libName Then
            GoTo CleanUp
        End If
    Next

    vbProj.References.AddFromGuid GUID, major, minor

CleanUp:
    Set vbProj = Nothing
End Function

Private Sub AddAllExternalReferences()
    AddScriptingLibraryReference
    AddVBELibraryReference
End Sub

Private Function AddScriptingLibraryReference() As Boolean
    ' Add Microsoft Scripting Runtime
    Const GUID As String = "{420B2830-E718-11CF-893D-00A0C9054228}"
    
    AddReference "Scripting", GUID, 1, 0
End Function

Private Sub AddVBELibraryReference()
    'Add VBIDE (Microsoft Visual Basic for Applications Extensibility 5.3)
    Const GUID As String = "{0002E157-0000-0000-C000-000000000046}"
    
    AddReference "VBIDE", GUID, 1, 0
End Sub

