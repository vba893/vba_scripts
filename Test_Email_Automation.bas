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
    
    Dim myFolder As Object
    Set myFolder = objNSpace.GetDefaultFolder(olFolderInbox)
    
    Dim myItem As Object
    Set myItem = myFolder.Items(1)
    myItem.Display
End Sub

Public Sub TestSendBatchEmail()
    Dim c As clsEmailAutomation
    Set c = New clsEmailAutomation

    c.SetSignature _
        signatureName:="SignatureName", _
        confirmSignature:=False

    ' SendUsingSpecificAccount

    c.ForAllRows Worksheets("Sheet2")
End Sub




