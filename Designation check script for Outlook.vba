Private Sub Application_ItemSend(ByVal item As Object, Cancel As Boolean)
 Dim isLeadershipAvailable As Boolean
 Dim nameofLeadership As String
 Dim objItem As Object
    Set objItem = GetCurrentItem()
    If Not objItem Is Nothing Then
          isLeadershipAvailable = False
          On Error GoTo HandleError
          'Cancel = True     'Comment this line out for testing. This line will prevent sending emails.
          For Each Recipient In objItem.Recipients
            If InStr(Recipient.AddressEntry, "@") = 0 Then       'This condition check external email address
                If InStr(Recipient.AddressEntry.GetExchangeUser.jobTitle, "Leadership") Then
                    isLeadershipAvailable = True
                    nameofLeadership = Recipient.AddressEntry.GetExchangeUser.Alias
                End If
            End If
HandleError:
        Resume ResumeProcess
ResumeProcess:
          Next Recipient
             If isLeadershipAvailable = True Then
                     Prompt$ = "We found Leadership " & nameofLeadership & " in recipients. Do you want to send?"
                     
                   If MsgBox(Prompt$, vbYesNo + vbQuestion + vbMsgBoxSetForeground, "Check for Leadership") = vbNo Then
                    Cancel = True
                    MsgBox ("Please reach out to your leads for the review")
                   Else
                        If MsgBox("Did your leads reviewed the email?", vbYesNo + vbQuestion + vbMsgBoxSetForeground, "Check for Leadership") = vbNo Then
                        Cancel = True
                        MsgBox ("Please reach out to your leads for the review")
                        End If
                   End If
              End If
    End If
    Set objItem = Nothing
    
  
End Sub

Function GetCurrentItem() As Object
    Dim objApp As Outlook.Application
    Set objApp = CreateObject("Outlook.Application")
    'MsgBox (TypeName(objApp.ActiveWindow))
    Select Case TypeName(objApp.ActiveWindow)
    Case "Explorer"
        Set GetCurrentItem = objApp.ActiveExplorer.ActiveInlineResponse
    Case "Inspector"
        Set GetCurrentItem = objApp.ActiveInspector.CurrentItem
    Case Else
    End Select
    Set objApp = Nothing
End Function