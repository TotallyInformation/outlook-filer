Sub fileMe()
'Select several emails, if one or more are NOT in the inbox or sentItems, file everything in the other folder.
' Checks for parent folder of selected items other than inbox and sent items. If one is found, move all of the other emails to that folder
' Works well in conversation view where you select entire conversation then run macro.
    Dim myItem As Outlook.mailItem
    Dim myDestFolder As Outlook.Folder
    For Each myItem In Application.ActiveExplorer.Selection
        If TypeName(myItem) = "MailItem" Then
            If myItem.Parent.Name <> "Inbox" And myItem.Parent.Name <> "sentItems" Then
                Set myDestFolder = myItem.Parent
                'Debug.Print myItem.Parent.Name
            End If
            
        End If
    Next
    
    If (Not myDestFolder Is Nothing) Then
        Call moveSelection(Application.ActiveExplorer.Selection, myDestFolder)
    End If
    
    Set myItem = Nothing
    Set myDestFolder = Nothing
End Sub

' Move a selection of emails to a given folder
Sub moveSelection(mySelection As Outlook.Selection, myDestFolder As Outlook.Folder)
    Dim myItem
    
    For Each myItem In mySelection
        If TypeName(myItem) = "MailItem" Then
            If myItem.Parent.Name <> myDestFolder.Name Then
                On Error Resume Next
                myItem.Move myDestFolder
                'Debug.Print "Moved to: ", myDestFolder.Name
                On Error GoTo 0
            End If
        End If
    Next
    
End Sub
