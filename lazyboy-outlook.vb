Public WithEvents myOlItems As Outlook.Items

Public Sub Application_Startup()

   ' Reference the items in the Inbox
   ' "WithEvents" the ItemAdd event will fire below.
   Set myOlItems = Outlook.Session.GetDefaultFolder(olFolderInbox).Items

End Sub

Private Sub myOlItems_ItemAdd(ByVal Item As Object)
    
    ' Declare application object
    Set olApp = CreateObject("Outlook.Application")
    olApp.Session.Logon
    
    ' Declare email object
    Dim oAccount As Outlook.Account
    Dim myAttachment As Outlook.Attachment
    Dim objMail As Outlook.MailItem
    Set objMail = olApp.CreateItem(olMailItem)
    
    ' Declare email message components
    Dim URL As String
    Dim sName As String

    ' Declare filepaths for downloading attachments
    Dim filepath As String
    Dim subpath As String
    Dim lastpath As String
    Dim sendtime As String
        
    On Error Resume Next
    With objMail
            
        If IsEmpty(Item.To) Then
            'Do Nothing
        Else
        
        'Get the sender's first name
        sName = Left(Item.SenderName, InStr(1, Item.SenderName, " ") - 1)
            
            'Based on the email address the sender sent to, reply with corresponding custom message. Add additional cases for different addresses
            Select Case Item.To
            
                ' ------------------------------- Email Response #1 -------------------------------                
                Case ""   'Insert your email address that your sender is sending to (e.g. yourname@gmail.com)
                    URL = ""   'Insert an URL that you would like your sender to receive
                    filepath = "C:\attachments\"  'Specify local folder to store attachments
                    
					'Automatically creates subfolders for each sender based on their first names
                    MkDir filepath
                    subpath = filepath & Item.To & "\"
                    MkDir subpath
                    lastpath = subpath & Item.SenderName & "\"
                    MkDir lastpath
                        For Each myAttachment In Item.Attachments
                            myAttachment.SaveAsFile lastpath & myAttachment.FileName
                        Next

					'Specify which email account on Outlook to send your auto-reply FROM, the default account is "1"						
                    Set oAccount = olApp.Session.Accounts.Item(1)
                    
					'Specify the email gets sent out (e.g. today at 5pm)
                    sendtime = Date & " 5:00:00 PM"
                    
					'Creates email message and send
                    .SendUsingAccount = oAccount
                    .To = Item.SenderEmailAddress
                    .Subject = "Re: " & Item.Subject
                    .HTMLBody = "<span style=""font-family : arial; font-size : 12pt""><p>Hi " & sName & ",</p><p>Thank you for sending me an email! Visit my github profile here:<br /><br />" & URL & "<br /><br />Best regards,<br /><br />My Name</p></span>"
                    .DeferredDeliveryTime = sendtime
                    .Send
               
              End Select
            
        End If
        
    End With
    
	'Close out the objects when finished
    Set olApp = Nothing
    Set objMail = Nothing

End Sub