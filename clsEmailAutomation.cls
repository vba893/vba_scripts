VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsEmailAutomation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Credit : http://learnexcelmacro.com/wp/2016/12/image-in-signature-not-displayed-mail-sent-by-vba/
Option Explicit
Private WithEvents oMailItem As Outlook.MailItem
Attribute oMailItem.VB_VarHelpID = -1
Private WithEvents oOutlook As Outlook.Application

Enum MailUserAction
    Unknown = 0
    Canceled = 1
    Send = 2
End Enum

Enum FieldType
    eRecipient_To = 0
    eRecipient_CC
    eRecipient_BCC
    eSubject
    eBody
    eAttachments
    eStatus
    eCustomField
End Enum

Private Type StateData
    CityCode(1 To 100) As Integer     ' Declare a static array.
    County As String * 30
End Type

Private columnHeaderTypes As Collection

Private signatureHtml As String
Private ws As Worksheet

Private currentMailUserAction As MailUserAction
Private selectedAccount As Outlook.Account

Public Function SetSignature(ByVal signatureName As String, ByVal confirmSignature As Boolean) As Boolean
    ' Load Signature

    SetSignature = LoadSignature(signatureName)
    If SetSignature = False Then
        Exit Function
    End If
     
    If confirmSignature Then
        SetSignature = ConfirmEmailSignature()
    End If
End Function

Public Function SendUsingSpecificAccount(ByVal accountName As String, Optional showError As Boolean = True) As Boolean
    Dim oAccount As Outlook.Account
    If oOutlook Is Nothing Then
        Set oOutlook = CreateObject("Outlook.Application")
    End If
    
    For Each oAccount In oOutlook.Session.Accounts
        If oAccount.DisplayName = accountName Then
            Set selectedAccount = oAccount
            SendUsingSpecificAccount = True
            Exit Function
        End If
    Next
    
    If showError Then
        MsgBox "Unable to find account name '" & accountName & "'", vbCritical + vbOKOnly, "Invalid account name"
    End If
    
    SendUsingSpecificAccount = False
End Function

Public Sub PrintAvailableAccounts()
    Dim oAccount As Outlook.Account
    If oOutlook Is Nothing Then
        Set oOutlook = CreateObject("Outlook.Application")
    End If
    Dim itemCount As Integer
    
    Debug.Print "---------------------------------------------------"
    Debug.Print " Available account names : "
    
    itemCount = 0
    For Each oAccount In oOutlook.Session.Accounts
        itemCount = itemCount + 1
        Debug.Print " Accounts(" & CStr(itemCount) & ") = " & oAccount.DisplayName
    Next
    
    Debug.Print "---------------------------------------------------"
End Sub

Public Sub ForAllRows(ByRef sourceWorksheet As Worksheet)
    Dim isAlreadySent As Boolean
    Dim isHeaderRow As Boolean
    Dim continueWithNextSend As Boolean
    
    If oOutlook Is Nothing Then
        Set oOutlook = CreateObject("Outlook.Application")
    End If
    
    Set ws = sourceWorksheet
    
    Dim rng As Range
    Set rng = ws.UsedRange
    Dim r As Range
    
    parseFieldsType
    validateExpectedMinimumFields
    
    For Each r In rng.Rows
        isAlreadySent = Trim(r.columns(getFieldTypeIndex(eStatus)).text) <> ""
        isHeaderRow = (r.row = 1)
        
        If isHeaderRow Or isAlreadySent Then
            GoTo ContinueLoop
        End If
        
        DebugCellsValue r
        continueWithNextSend = SendEmail(r)
        If continueWithNextSend = False Then
            Exit Sub
        End If
ContinueLoop:
    Next r
   
End Sub

Private Function ConfirmEmailSignature() As Boolean
    If oOutlook Is Nothing Then
        Set oOutlook = CreateObject("Outlook.Application")
    End If
    Set oMailItem = oOutlook.CreateItem(0)
    With oMailItem
        .Subject = "Make sure the selected email signature is the right one. All emails will be sent using this signature"
        .HTMLBody = signatureHtml
        .Display True
    End With
     
    Dim response
    response = MsgBox("Do you want to send all emails using the following signature ?", vbQuestion + vbYesNo, "Signature selection")
    ConfirmEmailSignature = (response = vbYes)
End Function

Private Sub DebugCellsValue(ByRef rng As Range)
    Debug.Print "------------------------------------------------------------------------------"
        
    Dim columnsIndexes() As Integer
    
    Debug.Print "Row" & Space(15 - Len("Row")) & " : " & CStr(rng.row)
    DebugFieldType rng, eRecipient_To
    DebugFieldType rng, eRecipient_CC
    DebugFieldType rng, eRecipient_BCC
    DebugFieldType rng, eSubject
    DebugFieldType rng, eBody
    DebugFieldType rng, eAttachments
    DebugFieldType rng, eStatus
End Sub

Private Sub DebugFieldType(ByRef rng As Range, ByVal ft As FieldType)
    Dim idx As Variant
    Dim columnsIndexes() As Integer
    Dim fieldName As String
    Dim text As String
    
    columnsIndexes = getFields(ft)
    fieldName = FieldTypeToString(ft)
    If ArraySize(columnsIndexes) > 0 Then
        For Each idx In columnsIndexes
            text = rng.columns(idx).value
            Select Case ft
            Case eRecipient_To, eRecipient_CC, eRecipient_BCC, eAttachments
                text = Join(SplitRe(rng.columns(idx).value, "[,;\n]"), ";")
            End Select
            
            Debug.Print fieldName & Space(15 - Len(fieldName)) & " : " & text
        Next idx
    End If
End Sub

Private Function SendEmail(ByRef currentRow As Range) As Boolean
    Set oMailItem = oOutlook.CreateItem(0)
    Dim fieldBody As String
    Dim lineBreak As String
    Dim fillTemplate As New clsReplaceMessageFields
    Dim wasSent As Boolean
    Dim columnIndex As Integer
    Dim ft As FieldType
    
    SendEmail = True
    
    lineBreak = "<br/>"
    
    fieldBody = Replace(Replace(currentRow.columns(getFieldTypeIndex(eBody)).value, vbCr, lineBreak), vbLf, lineBreak)
    fieldBody = fillTemplate.ParseTemplateMessage(ws, fieldBody, currentRow.row)
    
    For columnIndex = 1 To columnHeaderTypes.Count
        Select Case columnHeaderTypes(columnIndex)
        Case eRecipient_To, eRecipient_CC, eRecipient_BCC
            If addRecipients(currentRow, columnIndex) = False Then
                Exit Function
            End If
        Case eAttachments
            If addAttachments(currentRow, columnIndex) = False Then
                Exit Function
            End If
        End Select
    Next columnIndex
    
    Dim oRecip As Outlook.recipient
    For Each oRecip In oMailItem.recipients
        oRecip.Resolve
'        Debug.Print mailRecipientTypeToString(oRecip.Type) & " -> " & oRecip.Address
    Next
    
    currentMailUserAction = Unknown
    On Error Resume Next
    With oMailItem
'        If sendOnBehalf <> "" Then
'            ' Macros to send messages using a specific account
'            ' https://www.slipstick.com/developer/send-using-default-or-specific-account/
'            .SentOnBehalfOfName = sendOnBehalf
'        End If
        If Not selectedAccount Is Nothing Then
            .SendUsingAccount = selectedAccount
        End If
        .Subject = currentRow.columns(getFieldTypeIndex(eSubject)).value
        .HTMLBody = fieldBody & lineBreak & lineBreak & signatureHtml
        
        '.Send   'or use .Display
        .Display True
    End With
    
    While currentMailUserAction = Unknown
        DoEvents
    Wend
    
    Dim result As VbMsgBoxResult
    Dim msg As String
    If currentMailUserAction = Canceled Then
'        msg = msg & "Do you want to continue sending emails ?" + vbCrLf + vbCrLf
'        msg = msg & "Press YES to continue sending the remaining emails." + vbCrLf
'        msg = msg & "Press NO to cancel the process."
'
'        result = MsgBox(msg, vbYesNo + vbQuestion + vbSystemModal, "Continue or abort process")
'        If result = vbNo Then
'            SendEmail = False
'            Exit Function
'        End If
        SendEmail = False
        Exit Function
    End If
       
    If currentMailUserAction = Send Then
        currentRow.columns(getFieldTypeIndex(eStatus)).value = Now()
    End If
    On Error GoTo 0
    
    Set oMailItem = Nothing
End Function

Private Function addAttachments(ByRef currentRow As Range, ByVal columnIndex As Integer) As Boolean
    Dim attachments() As String
    Dim filename As Variant
    Dim allAttachmentExists As Boolean

    attachments = SplitRe(currentRow.columns(columnIndex).text, "[,;\n]")
    If LBound(attachments) = 0 Then
        addAttachments = True
        Exit Function
    End If
    
    allAttachmentExists = True
    For Each filename In attachments
        Debug.Print CStr(filename)
        allAttachmentExists = allAttachmentExists And Dir(filename, vbNormal) <> ""
        If Dir(filename, vbNormal) = "" Then
            MsgBox "Unable to find file" + vbCrLf + filename, vbCritical Or vbOKOnly, "File not found"
        End If
    Next filename

    If allAttachmentExists = False Then
        MsgBox "Unable to find one or more attachments." & vbCrLf & "Row " & Str(currentRow.row) + " will be skipped.", vbCritical Or vbOKOnly, "File not found"
        addAttachments = False
        Exit Function
    End If
    
    For Each filename In attachments
        oMailItem.attachments.Add filename
    Next filename

    addAttachments = True
End Function

Private Function addRecipients(ByRef currentRow As Range, ByVal columnIndex As Integer) As Boolean
    Dim recipients() As String
    Dim strRecipient As Variant
    Dim allValidEmails As Boolean
    
    recipients = SplitRe(currentRow.columns(columnIndex).text, "[,;\n]")
    If LBound(recipients) = 0 Then
        addRecipients = True
        Exit Function
    End If
    
    allValidEmails = True
    For Each strRecipient In recipients
        allValidEmails = allValidEmails And isValidEmail(strRecipient)
        If isValidEmail(strRecipient) = False Then
            MsgBox "The provided email address format is invalid : " + strRecipient, vbCritical Or vbOKOnly, "Invalid email address"
        End If
    Next strRecipient

    If allValidEmails = False Then
        MsgBox "Unable to resolve one or more recipients." & vbCrLf & "Row " & Str(currentRow.row) + " will be skipped.", vbCritical Or vbOKOnly, "Unable to resolve recipient"
        addRecipients = False
        Exit Function
    End If
    
    Dim oRecip As Outlook.recipient
    For Each strRecipient In recipients
        Set oRecip = oMailItem.recipients.Add(strRecipient)
        
        Select Case columnHeaderTypes(columnIndex)
        Case eRecipient_To
            oRecip.Type = olTo
        Case eRecipient_CC
            oRecip.Type = olCC
        Case eRecipient_BCC
            oRecip.Type = olBCC
        End Select
    Next strRecipient

    addRecipients = True
End Function

Private Function SplitRe(ByVal text As String, ByVal pattern As String, Optional IgnoreCase As Boolean) As String()
    Dim tmpSplit() As String
    Dim filteredSplit() As String
    Dim filteredCount As Integer
    Dim i As Integer
    
    Static re As Object
    If re Is Nothing Then
        Set re = New RegExp
        re.Global = True
        re.MultiLine = True
    End If

    re.IgnoreCase = IgnoreCase
    re.pattern = pattern
    tmpSplit = Strings.Split(re.Replace(text, ChrW(-1)), ChrW(-1))
    
    filteredCount = 0
    ReDim filteredSplit(1 To 255)
    For i = LBound(tmpSplit) To UBound(tmpSplit)
        If Trim(tmpSplit(i)) <> "" Then
            filteredCount = filteredCount + 1
            filteredSplit(filteredCount) = Trim(tmpSplit(i))
        End If
    Next i
    
    If filteredCount > 0 Then
        ReDim Preserve filteredSplit(1 To filteredCount)
        SplitRe = filteredSplit
    Else
        ReDim filteredSplit(0 To 0)
        SplitRe = filteredSplit
    End If
    
End Function

Private Function FieldTypeToString(ByVal rt As FieldType) As String
    Select Case rt
    Case FieldType.eRecipient_To
        FieldTypeToString = "To"
    Case FieldType.eRecipient_CC
        FieldTypeToString = "CC"
    Case FieldType.eRecipient_BCC
        FieldTypeToString = "BCC"
    Case FieldType.eSubject
        FieldTypeToString = "Subject"
    Case FieldType.eBody
        FieldTypeToString = "Body"
    Case FieldType.eAttachments
        FieldTypeToString = "Files"
    Case FieldType.eStatus
        FieldTypeToString = "Status"
    Case Else
        FieldTypeToString = "Custom"
    End Select
End Function

Private Function mailRecipientTypeToString(ByVal rt As OlMailRecipientType) As String
    mailRecipientTypeToString = FieldTypeToString(CInt(rt))
End Function

' Credits : https://officetricks.com/validate-email-address-vba-code-excel/
Private Function isValidEmail(ByVal sEmailAddress As String) As Boolean
    'Code from Officetricks
    'Define variables
    Dim sEmailPattern As String
    Dim oRegEx As RegExp
    Dim bReturn As Boolean
    
    'Use the below regular expressions
    sEmailPattern = "^\w+([\.-]?\w+)*@\w+([\.-]?\w+)*(\.\w{2,3})+$" 'or
    sEmailPattern = "^([a-zA-Z0-9_\-\.]+)@[a-z0-9-]+(\.[a-z0-9-]+)*(\.[a-z]{2,3})$"
    
    'Create Regular Expression Object
    Set oRegEx = New RegExp
    oRegEx.Global = True
    oRegEx.IgnoreCase = True
    oRegEx.pattern = sEmailPattern
    bReturn = False
    
    'Check if Email match regex pattern
    If oRegEx.test(sEmailAddress) Then
        'Debug.Print "Valid Email ('" & sEmailAddress & "')"
        bReturn = True
    Else
        'Debug.Print "Invalid Email('" & sEmailAddress & "')"
        bReturn = False
    End If

    'Return validation result
    isValidEmail = bReturn
End Function

Private Function LoadSignature(ByVal Signature As String) As Boolean

    Dim strbody As String
    Dim signatureFilename As String
    Dim signImageFolderName As String
    Dim completeFolderPath As String

    signatureFilename = Environ("appdata") & _
                "\Microsoft\Signatures\" & Signature & ".htm"

    signImageFolderName = Signature & "_files"
    completeFolderPath = Environ("appdata") & "\Microsoft\Signatures\" & signImageFolderName

    If Dir(signatureFilename) <> "" Then
        signatureHtml = ReadFileContent(signatureFilename)
        signatureHtml = VBA.Replace(signatureHtml, signImageFolderName, completeFolderPath)
        LoadSignature = True
    Else
        MsgBox "Signature '" & Signature & "' not found"
        LoadSignature = False
    End If
End Function

Private Function ReadFileContent(ByVal sFile As String) As String
'Dick Kusleika
    Dim fso As Object
    Dim ts As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.GetFile(sFile).OpenAsTextStream(1, -2)
    ReadFileContent = ts.readall
    ts.Close
End Function

Private Sub parseFieldsType()
    Set columnHeaderTypes = New Collection
    
    Dim re As New RegExp
    re.IgnoreCase = True

    Dim c As Range
    Dim headerName As String
    For Each c In ws.UsedRange.columns
        headerName = CStr(c.Value2(1, 1))
        
        re.pattern = "^To\s*"
        If re.test(headerName) Then
            columnHeaderTypes.Add FieldType.eRecipient_To
            GoTo ContinueLoop
        End If
        
        re.pattern = "^CC\s*"
        If re.test(headerName) Then
            columnHeaderTypes.Add FieldType.eRecipient_CC
            GoTo ContinueLoop
        End If
        
        re.pattern = "^BCC\s*"
        If re.test(headerName) Then
            columnHeaderTypes.Add FieldType.eRecipient_BCC
            GoTo ContinueLoop
        End If
        
        re.pattern = "^Subject\s*"
        If re.test(headerName) Then
            columnHeaderTypes.Add FieldType.eSubject
            GoTo ContinueLoop
        End If
        
        re.pattern = "^Body\s*"
        If re.test(headerName) Then
            columnHeaderTypes.Add FieldType.eBody
            GoTo ContinueLoop
        End If
        
        re.pattern = "^Attachments\s*"
        If re.test(headerName) Then
            columnHeaderTypes.Add FieldType.eAttachments
            GoTo ContinueLoop
        End If
        
        re.pattern = "^Status\s*"
        If re.test(headerName) Then
            columnHeaderTypes.Add FieldType.eStatus
            GoTo ContinueLoop
        End If
        
        columnHeaderTypes.Add FieldType.eCustomField
        
ContinueLoop:
    Next c
End Sub

Private Function validateExpectedMinimumFields() As Boolean
    Dim size As Integer
    validateExpectedMinimumFields = False
    
    size = ArraySize(getFields(eRecipient_To))
    If size = 0 Then
        MsgBox "Expected at least one 'To' column"
        Exit Function
    End If
    
    size = ArraySize(getFields(eSubject))
    If size = 0 Or size > 1 Then
        MsgBox "Expected exactly one 'Subject' column. But " & Str(size) & " were provided"
        Exit Function
    End If

    size = ArraySize(getFields(eBody))
    If size = 0 Or size > 1 Then
        MsgBox "Expected exactly one 'Body' column. But " & Str(size) & " were provided"
        Exit Function
    End If

    size = ArraySize(getFields(eAttachments))
    If size = 0 Or size > 1 Then
        MsgBox "Expected at most one 'Attachments' column. But " & Str(size) & " were provided"
        Exit Function
    End If
    
    size = ArraySize(getFields(eStatus))
    If size = 0 Or size > 1 Then
        MsgBox "Expected at most one 'Status' column. But " & Str(size) & " were provided"
        Exit Function
    End If
    
    validateExpectedMinimumFields = True
End Function

Private Function getFields(ByVal ft As FieldType) As Integer()
    Dim idx As Integer
    Dim foundIndexes() As Integer
    Dim foundCount As Integer
    ReDim foundIndexes(1 To 255)
    
    foundCount = 0
    For idx = 1 To columnHeaderTypes.Count
        If columnHeaderTypes(idx) = ft Then
            foundCount = foundCount + 1
            foundIndexes(foundCount) = idx
        End If
    Next idx
    
    If foundCount > 0 Then
        ReDim Preserve foundIndexes(1 To foundCount)
        getFields = foundIndexes
    End If
    
    
End Function

Private Function getFieldTypeIndex(ByVal ft As FieldType) As Integer
    Dim idx As Integer
    
    For idx = 1 To columnHeaderTypes.Count
        If columnHeaderTypes(idx) = ft Then
            getFieldTypeIndex = idx
            Exit Function
        End If
    Next idx
End Function

Private Function IsArrayAllocated(Arr As Variant) As Boolean
    Dim N As Long
    On Error Resume Next
    
    ' if Arr is not an array, return FALSE and get out.
    If IsArray(Arr) = False Then
        IsArrayAllocated = False
        Exit Function
    End If
    
    N = UBound(Arr, 1)
    If (Err.Number = 0) Then
        If LBound(Arr) <= UBound(Arr) Then
            ' no error. array has been allocated.
            IsArrayAllocated = True
        Else
            IsArrayAllocated = False
        End If
    Else
        ' error. unallocated array
        IsArrayAllocated = False
    End If

End Function

Private Function ArraySize(Arr As Variant) As Integer
    If IsArrayAllocated(Arr) = False Then
        ArraySize = 0
        Exit Function
    End If
    
    ArraySize = UBound(Arr) - LBound(Arr) + 1
    
End Function

Private Sub oMailItem_Close(Cancel As Boolean)
    ' User pressed Close button
    If currentMailUserAction = Unknown Then
        currentMailUserAction = Canceled
    End If
End Sub

Private Sub oMailItem_Send(Cancel As Boolean)
    ' User pressed Send button
    currentMailUserAction = Send
End Sub
