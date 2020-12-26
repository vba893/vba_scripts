Attribute VB_Name = "Test_Email_Automation"
Option Explicit

Public Sub GetFirstOutlookItem()
    ' Set Outlook application object.
    Dim objOutlook As Object
    Set objOutlook = CreateObject("Outlook.Application")
    
    Debug.Print objOutlook.ProductCode
    Debug.Print objOutlook.Name
        
    Dim objNSpace As Object     ' Create and Set a NameSpace OBJECT.
    ' The GetNameSpace() method will represent a specified Namespace.
    Set objNSpace = objOutlook.GetNamespace("MAPI")
    
    Dim myFolder As MAPIFolder
    Set myFolder = objNSpace.GetDefaultFolder(olFolderInbox)
    
    Dim myItem As MailItem
    Set myItem = myFolder.Items(1)
    myItem.Display
End Sub

Public Sub TestSendBatchEmail()
    Dim result As Boolean
    Dim c As clsEmailAutomation
    Set c = New clsEmailAutomation
    
    ' Expected column names                         | Single or multiple        | Notes
    ' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------
    ' To    or To CustomName    or To (CustomName)  | Multiple columns allowed  |
    ' CC    or CC CustomName    or CC (CustomName)  | Multiple columns allowed  |
    ' BCC   or CC CustomName    or CC (CustomName)  | Multiple columns allowed  |
    ' Subject                                       | Single column             |
    ' Body                                          | Single column             | String matching <AnotherColumnName> will be replaced by the corresponding row/column value
    ' Attachments                                   | Single column             | One line per file
    ' Status                                        | Single column             | If cell if not empty, row will be skipped
    
    result = c.SetSignature(signatureName:="SignatureName", confirmSignature:=True)
    If Not result Then
        Exit Sub
    End If

    c.PrintAvailableAccounts
'    result = c.SendUsingSpecificAccount("AccountName")
'    If Not result Then
'        Exit Sub
'    End If

    c.ForAllRows Worksheets("Sheet2")
End Sub




