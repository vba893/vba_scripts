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
Public WithEvents OutMail As Outlook.MailItem
Attribute OutMail.VB_VarHelpID = -1

Enum MailUserAction
    Unknown = 0
    Canceled = 1
    Send = 2
End Enum

Enum RecipientType
    Field_To = OlMailRecipientType.olTo
    Field_CC = OlMailRecipientType.olCC
    Field_BCC = OlMailRecipientType.olBCC
End Enum

Private Const columnTo              As String = "A"
Private Const columnCC              As String = "B"
Private Const columnBCC             As String = "C"
Private Const columnSubject         As String = "D"
Private Const columnBody            As String = "E"
Private Const columnAttachments     As String = "F"
Private Const columnSendStatus      As String = "G"

Private selectedSignature As String
Private signatureHtml As String
Private ws As Worksheet
Private sendOnBehalf As String

Private columnNames As Collection
Private currentRowValues As Collection
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

Public Function SendUsingSpecificAccount(ByVal accountName As String) As Boolean
    Private oAccount As Outlook.Account
    
    For Each oAccount In Application.Session.Accounts
        If oAccount = accountName Then
            Set selectedAccount = oAccount
            SendUsingSpecificAccount = True
            Exit Function
        End If
    Next
    
    SendUsingSpecificAccount = False
End Function

Public Sub ForAllRows(ByRef sourceWorksheet As Worksheet)
    Dim OutApp As Object
    Dim continueWithNextSend As Boolean
    Set OutApp = CreateObject("Outlook.Application")
    Set ws = sourceWorksheet
    
    Dim rng As Range
    Set rng = ws.UsedRange
    Dim r As Range
       
    For Each r In rng.Rows
        DebugCellsValue r
        If r.row > 1 And Trim(r.Columns(columnSendStatus).text) = "" Then
            continueWithNextSend = SendEmail(OutApp, r)
            If continueWithNextSend = False Then
                Exit Sub
            End If
        End If
    Next r
   
End Sub

Private Function ConfirmEmailSignature() As Boolean
    Dim OutApp As Outlook.Application
    Set OutApp = New Outlook.Application
    Set OutMail = OutApp.CreateItem(0)
    With OutMail
        .Subject = "Make sure the selected email signature is the right one. All emails will be sent using this signature"
        .HTMLBody = signatureHtml
        .Display True
    End With
     
    Dim response
    response = MsgBox("Do you want to send all emails using the following signature ?", vbQuestion + vbYesNo, "Signature selection")
    If response <> vbYes Then
        ConfirmEmailSignature = False
    End If
    ConfirmEmailSignature = True
End Function

Private Sub DebugCellsValue(ByRef rng As Range)
    Debug.Print "------------------------------------------------------------------------------"
    Debug.Print "Row    : " & Str(rng.row)
    Debug.Print "To     : " & rng.Columns(columnTo).value
    Debug.Print "CC     : " & rng.Columns(columnCC).value
    Debug.Print "BCC    : " & rng.Columns(columnBCC).value
    Debug.Print "Subject: " & rng.Columns(columnSubject).value
    Debug.Print "Body   : " & rng.Columns(columnBody).value
    Debug.Print "Files  : " & rng.Columns(columnAttachments).value
    Debug.Print "Status : " & rng.Columns(columnSendStatus).value
End Sub

Private Function SendEmail(ByRef outlookApp As Object, ByRef currentRow As Range) As Boolean
    Set OutMail = outlookApp.CreateItem(0)
    Dim fieldBody As String
    Dim lineBreak As String
    Dim fillTemplate As New clsReplaceMessageFields
    Dim wasSent As Boolean
    
    SendEmail = True
    
    lineBreak = "<br/>"
    
    fieldBody = currentRow.Columns(columnBody).text
    fieldBody = Replace(Replace(currentRow.Columns(columnBody).value, vbCr, lineBreak), vbLf, lineBreak)
    fieldBody = fillTemplate.ParseTemplateMessage(ws, fieldBody, currentRow.row)
    
    If addAttachments(OutMail, currentRow) = False Then
        Exit Function
    End If
    
    If addRecipients(OutMail, currentRow, Field_To) = False Then
        Exit Function
    End If
    
    If addRecipients(OutMail, currentRow, Field_CC) = False Then
        Exit Function
    End If
    
    If addRecipients(OutMail, currentRow, Field_BCC) = False Then
        Exit Function
    End If
    
    Dim oRecip As Outlook.recipient
    For Each oRecip In OutMail.recipients
        oRecip.Resolve
        Debug.Print mailRecipientTypeToString(oRecip.Type) & " -> " & oRecip.Address
    Next
    
    currentMailUserAction = Unknown
    On Error Resume Next
    With OutMail
        If sendOnBehalf <> "" Then
            ' Macros to send messages using a specific account
            ' https://www.slipstick.com/developer/send-using-default-or-specific-account/
            .SentOnBehalfOfName = sendOnBehalf
        End If
        .Subject = currentRow.Columns(columnSubject).value
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
        msg = msg & "Do you want to continue sending emails ?" + vbCrLf + vbCrLf
        msg = msg & "Press YES to continue sending the remaining emails." + vbCrLf
        msg = msg & "Press NO to cancel the process."
        
        result = MsgBox(msg, vbYesNo + vbQuestion + vbSystemModal, "Continue or abort process")
        If result = vbNo Then
            SendEmail = False
            Exit Function
        End If
    End If
       
    If currentMailUserAction = Send Then
        currentRow.Columns(columnSendStatus).value = Now()
    End If
    On Error GoTo 0
    
    Set OutMail = Nothing
End Function

Private Function addAttachments(ByRef mail As Outlook.MailItem, ByRef currentRow As Range) As Boolean
    Dim attachments() As String
    Dim filename As Variant
    Dim allAttachmentExists As Boolean

    attachments = SplitRe(currentRow.Columns(columnAttachments).text, "[,;\n]")
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
        mail.attachments.Add filename
    Next filename

    addAttachments = True
End Function

Private Function addRecipients(ByRef mail As Outlook.MailItem, ByRef currentRow As Range, ByVal rt As RecipientType) As Boolean
    Dim recipients() As String
    Dim strRecipient As Variant
    Dim allValidEmails As Boolean
    Dim invalidEmails As New Collection
    
    Select Case rt
    Case RecipientType.Field_To
        recipients = SplitRe(currentRow.Columns(columnTo).text, "[,;\n]")
    Case RecipientType.Field_CC
        recipients = SplitRe(currentRow.Columns(columnCC).text, "[,;\n]")
    Case RecipientType.Field_BCC
        recipients = SplitRe(currentRow.Columns(columnBCC).text, "[,;\n]")
    End Select

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
    
    Select Case rt
    Case RecipientType.Field_To
        mail.To = Join(recipients, ";")
    Case RecipientType.Field_CC
        mail.CC = Join(recipients, ";")
    Case RecipientType.Field_BCC
        mail.BCC = Join(recipients, ";")
    End Select

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

Private Function recipientTypeToString(ByVal rt As RecipientType) As String
    Select Case rt
    Case RecipientType.Field_To
        recipientTypeToString = "To"
    Case RecipientType.Field_CC
        recipientTypeToString = "CC"
    Case RecipientType.Field_BCC
        recipientTypeToString = "BCC"
    End Select
End Function

Private Function mailRecipientTypeToString(ByVal rt As OlMailRecipientType) As String
    mailRecipientTypeToString = recipientTypeToString(CInt(rt))
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

Function ReadFileContent(ByVal sFile As String) As String
'Dick Kusleika
    Dim fso As Object
    Dim ts As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.GetFile(sFile).OpenAsTextStream(1, -2)
    ReadFileContent = ts.readall
    ts.Close
End Function

Private Sub OutMail_Close(Cancel As Boolean)
    ' User pressed Close button
    If currentMailUserAction = Unknown Then
        currentMailUserAction = Canceled
    End If
End Sub

Private Sub OutMail_Send(Cancel As Boolean)
    ' User pressed Send button
    currentMailUserAction = Send
End Sub
