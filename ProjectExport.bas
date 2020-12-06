Attribute VB_Name = "ProjectExport"
Public Sub ExportModules()
    Dim bExport As Boolean
    Dim wkbSource As Excel.Workbook
    Dim szSourceWorkbook As String
    Dim szExportPath As String
    Dim szFileName As String
    Dim cmpComponent As VBIDE.VBComponent

    ' OpenDialog
    Dim strFolderExists As String
    Dim fd As Office.FileDialog
    Dim myFolder As String
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    
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
        .Title = "Choose a folder where to export project files"
        .AllowMultiSelect = False
     
        .InitialFileName = userDocuments
     
        If .Show = True Then
            myFolder = .SelectedItems(1)
        End If
    End With
    
    ' Exit if no folder selected or user canceled
    If fd.SelectedItems.Count = 0 Then Return

    ''' NOTE: This workbook must be open in Excel.
    Set wkbSource = Application.ThisWorkbook
    
    If wkbSource.VBProject.Protection = 1 Then
    MsgBox "The VBA in this workbook is protected," & _
        "not possible to export the code"
    Exit Sub
    End If
    
    szExportPath = myFolder & "\"
    
    For Each cmpComponent In wkbSource.VBProject.VBComponents
        
        bExport = True
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
                ''' This is a worksheet or workbook object.
                ''' Don't try to export.
'                szFileName = szFileName & ".bas"
                bExport = False
        End Select
        
        If bExport Then
            ''' Export the component to a text file.
            cmpComponent.Export szExportPath & szFileName
            
        ''' remove it from the project if you want
        '''wkbSource.VBProject.VBComponents.Remove cmpComponent
        
        End If
   
    Next cmpComponent

    MsgBox "Export is ready"
End Sub

