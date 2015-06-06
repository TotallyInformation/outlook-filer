VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FolderSelectBox 
   Caption         =   "Select Folder for Filing"
   ClientHeight    =   5380
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10695
   OleObjectBlob   =   "FolderSelectBox.frx":0000
   StartUpPosition =   1  'CenterOwner
   WhatsThisHelp   =   -1  'True
End
Attribute VB_Name = "FolderSelectBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Show a list of folders taken from the locations of a selected set of emails.
' Also shows a FULL list of all folders with a dynamic filter box, type some chars to filter the full list
' If a folder is selected (preference given to the select folder list), file all of the selected emails to that folder.
'
' Facilitates really easy filing once you have put 1 email in a conversation to a particular folder,
' you can easily file all of the other related emails without having to hunt through all of your
' folders. Just make sure you use a view with "Conversation View" turned on.
'
' Also has a View button to switch the current view in outlook to the selected folder instead of filing.
'
' NB: Double-click on on a folder in either list is the same as pressing the "File" Button.
'
' WARNINGS: Assumed English folder names for exclusion of Inbox, Sent, etc.
'           Max 999 folders are supported, change the DIM's below if more needed
'           It is possible to end up with >1 folder when filtering, this gives an error and doesn't move - change the filter
'
' TO DO:
'   2) Add ability to move to another mailbox (http://www.slipstick.com/developer/working-vba-nondefault-outlook-folders/)
'   3) Allow multiple filters for full folder list
'   4) Pre-populate multi filters from conversation subject
'
' Author: Julian Knight (Totally Information)
' Version: v1.0 20015-05-08
' History:
'   v1.2 20015-05-18 - Add double-click processing
'   v1.1 20015-05-12 - Various improvements - add filter to full folder list, add view button
'   v1.0 20015-05-08 - Initial Release

Option Explicit
    
Dim folderNames(0 To 99) As String
Dim maxNames As Long
Dim folderPaths(0 To 99) As String
Dim maxPaths As Long
Dim folderAllPaths(0 To 999) As Variant
Dim maxFAP As Long
Dim folderAllNames(0 To 999) As String
Dim maxFAN As Long
Dim mailbox As String

Private Sub btnCancel_Click()
    ' Do nothing other than cancel everything
    Unload Me
End Sub

Private Sub btnView_Click()
    Dim fldr As Outlook.MAPIFolder
    Dim objItem As Outlook.MailItem
    
    Set fldr = fldrDest
    ' If anywhere to move to, move each email now
    If IsObject(fldr) Then
        Set Application.ActiveExplorer.CurrentFolder = fldr
    End If
    
    ' End
    Set fldr = Nothing
    Set objItem = Nothing
    Unload Me
End Sub

Private Sub btnFileToFolder_Click()
    Dim fldr As Outlook.MAPIFolder
    Dim objItem As Outlook.MailItem
    Dim x As Long
    
    Set fldr = fldrDest
    On Error GoTo err
    ' If anywhere to move to, move each email now
    If IsObject(fldr) Then
        x = 0
        For Each objItem In ActiveExplorer.Selection
            ' Only move items not already in the dest folder
            If objItem.Parent.Name <> fldr.Name Then
                objItem.Move fldr
                x = x + 1
            End If
        Next objItem
    End If
    
    GoTo endit
err:
    MsgBox "Error processing selection, something odd selected?", vbCritical, "Folder Move Error"
    ' End
endit:
    On Error GoTo 0
    Set fldr = Nothing
    Set objItem = Nothing
    Unload Me
End Sub

Private Function fldrDest() As Outlook.MAPIFolder
    Dim obj As Object
    'Dim fldrDest As Outlook.MAPIFolder
    Dim destFldr As String
    Dim arr, e, i As Integer
    Dim objItem As Outlook.MailItem
    Dim x
    
    ' Index = -1 if nothing selected
    If lstFolders.ListIndex > -1 Then
        'Debug.Print "Selected Folder", lstFolders.Value, lstFolders.ListIndex
        'Debug.Print folderPaths(lstFolders.ListIndex)
        
        'NB: application.session.... gives the top level folder set for the current mailbox
        Set fldrDest = ReturnDestinationFolder(folderPaths(lstFolders.ListIndex), _
            Application.Session.GetDefaultFolder(olFolderInbox).Parent.Folders _
        )
        ' set i=1 as we can only ever select one entry on this side
        i = 1
        
    ElseIf lstAllFolders.ListIndex > -1 Then
        'Debug.Print "Selected from All Folders", lstAllFolders.Value; lstAllFolders.ListIndex
        
        arr = Filter(SourceArray:=folderAllPaths, match:=lstAllFolders.Value, Compare:=vbTextCompare)
        ' Annoyingly, filter has no way to do exact matches
        i = 0
        For Each e In arr
            If Len(e) = Len("\\" & mailbox & lstAllFolders.Value) Then
                ' exact match
                i = i + 1
                destFldr = e
            End If
        Next e
        ' If there is more than one matching folder, error, else move
        If i = 1 Then
            'Debug.Print "Filtered:", arr(0)
            
            'NB: application.session.... gives the top level folder set for the current mailbox
            Set fldrDest = ReturnDestinationFolder(destFldr, _
                Application.Session.GetDefaultFolder(olFolderInbox).Parent.Folders _
            )
        Else
            MsgBox "Zero or more than 1 folder was returned. Giving up", vbCritical, "Move to Folder Error"
        End If
    End If
    
    If i > 1 Or i = 0 Then
        fldrDest = Nothing
    End If
    
    Set obj = Nothing
    Set objItem = Nothing

End Function

Private Function ReturnDestinationFolder(findStr As Variant, fldrs As Outlook.Folders _
    ) As Outlook.MAPIFolder
    
    Dim fldr As Outlook.MAPIFolder
    Dim findArr As Variant
    Dim idx As Long
    
    ' Split the path into an array
    findStr = Replace(findStr, "\\", "")
    findArr = Split(findStr, "\")
    
    ' We are going to ignore the mailbox ID
    idx = LBound(findArr)
    If InStr(findArr(idx), "@") Then idx = idx + 1
    
    For Each fldr In fldrs
        If fldr.Name = findArr(idx) Then
            ' Any more to find?
            If UBound(findArr) > idx Then 'LBound(findArr) Then
                ' Yes, so recurse if there are any sub folders
                If fldr.Folders.Count Then
                    Set ReturnDestinationFolder = ReturnDestinationFolder( _
                        findArr(idx + 1), _
                        fldr.Folders _
                    )
                Else
                    ' No sub folders so we give up
                    Set ReturnDestinationFolder = Nothing
                End If
            Else
                ' No, so return the found folder
                Set ReturnDestinationFolder = fldr
            End If
            ' We either found it or failed to so no point in going further
            Exit For
        End If
    Next fldr
    
    Set fldr = Nothing

End Function

Private Sub lstAllFolders_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Call btnFileToFolder_Click
End Sub

Private Sub lstFolders_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Call btnFileToFolder_Click
End Sub

'Private Sub tbFilterAllFolders_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
'    With Me.tbFilterAllFolders
'        If .Value = vbNullString Then
'            Me.lstAllFolders.List = folderAllNames
'        Else
'            Me.lstAllFolders.List = Filter(SourceArray:=folderAllNames, match:=.Value, Compare:=vbTextCompare)
'        End If
'    End With
'
'End Sub

Private Sub tbFilterAllFolders_Change()
    With Me.tbFilterAllFolders
        If .Value = vbNullString Then
            Me.lstAllFolders.List = folderAllNames
        Else
            Me.lstAllFolders.List = Filter(SourceArray:=folderAllNames, match:=.Value, Compare:=vbTextCompare)
        End If
    End With
End Sub

Private Sub UserForm_Initialize()

    Dim objItem As Object
    Dim i As Long
    Dim numSelected As Long
    Dim numEmailsSelected As Long
    Dim mb
    
    ' Walk through all selected emails and compile a list of folders
    ' that they are in. Ignore the inbox
    i = 0
    maxNames = 0
    maxPaths = 0
    numSelected = 0
    numEmailsSelected = 0
    For Each objItem In ActiveExplorer.Selection
        numSelected = numSelected + 1
        ' Only interested in real mail items (not calendar entries, cancellation notices, etc.)
        If objItem.MessageClass = "IPM.Note" Then
            numEmailsSelected = numEmailsSelected + 1
            ' Check that parent item really is a folder
            If objItem.Parent.Class = olFolder Then
                ' Only want folders != Inbox/Send/Deleted
                If objItem.Parent.Name <> "Inbox" And _
                        objItem.Parent.Name <> "Sent Items" And _
                        objItem.Parent.Name <> "Deleted Items" And _
                        IsInArray(folderPaths, objItem.Parent.folderPath) = False Then
                        'Contains(folderPaths, objItem.Parent.folderPath) = False Then
                        
                    ' Save mailbox name
                    If maxPaths = 0 Then
                        mb = Split(objItem.Parent.folderPath, "\")
                        mailbox = mb(2)
                    End If
                    folderPaths(maxPaths) = objItem.Parent.folderPath
                    maxPaths = maxPaths + 1
                    folderNames(maxNames) = objItem.Parent.Name
                    maxNames = maxNames + 1
                    
                    'Debug.Print "New Folder:", folderPaths(i), folderNames(i), i
                    i = i + 1
                    
                End If
            End If
        End If
    Next objItem
    
    'Debug.Print "# Selected:", numSelected
    'Debug.Print "# Emails Sel:", numEmailsSelected
    'Debug.Print "# Folders:", i, "(Igoring Inbox)"
    
    lstFolders.List = folderNames
    
    'GetAllFolders
    maxFAP = 0
    maxFAN = 0
    ProcessFolder Application.Session.GetDefaultFolder(olFolderInbox).Parent
    
    Set objItem = Nothing
    
End Sub

Sub ProcessFolder(objStartFolder As Outlook.MAPIFolder, Optional blnRecurseSubFolders As Boolean = True, Optional strFolderPath As String = "", Optional strFolderName As String = "")

    Dim objFolder As Outlook.MAPIFolder

    Dim i As Long, mb

     ' Loop through the items in the current folder
    For i = 1 To objStartFolder.Folders.Count

        Set objFolder = objStartFolder.Folders(i)

        ' Populate the listbox & save actual folder paths
        ' But only for NOT sent, drafts, etc
        ' Don't block the inbox in case it has sub-folders
        If objFolder.Name <> "Sent Items" And _
            objFolder.Name <> "Deleted Items" And _
            objFolder.Name <> "Outbox" And _
            objFolder.Name <> "Calendar" And _
            objFolder.Name <> "Contacts" And _
            objFolder.Name <> "Notes" And _
            objFolder.Name <> "Journal" And _
            objFolder.Name <> "Junk E-mail" And _
            objFolder.Name <> "News Feed" And _
            objFolder.Name <> "RSS Feeds" And _
            objFolder.Name <> "Conversation History" And _
            objFolder.Name <> "Conversation Action Settings" And _
            objFolder.Name <> "Quick Step Settings" And _
            objFolder.Name <> "LinkedIn" And _
            objFolder.Name <> "Suggested Contacts" And _
            objFolder.Name <> "Sync Issues" And _
            objFolder.Name <> "Tasks" And _
            objFolder.Name <> "My Site" And _
            objFolder.Name <> "Drafts" _
        Then
            ' Save mailbox name
            If maxFAP = 0 And mailbox = "" Then
                mb = Split(objFolder.folderPath, "\")
                mailbox = mb(2)
            End If

            lstAllFolders.AddItem strFolderName + "\" + objFolder.Name
            folderAllPaths(maxFAP) = objFolder.folderPath
            maxFAP = maxFAP + 1
            folderAllNames(maxFAN) = strFolderName + "\" + objFolder.Name
            maxFAN = maxFAN + 1
            ' Recurse subfolders but not for subfolders of blocked folders
            If blnRecurseSubFolders Then
                ' Recurse through subfolders
                ProcessFolder objFolder, True, strFolderPath + "\" + objFolder.folderPath, _
                    strFolderName + "\" + objFolder.Name
            End If
        End If

    Next
    
    Set objFolder = Nothing

End Sub

' ---- Functions from other people ----

Private Sub AddToArray(ByRef arr As Variant, val As Variant)
On Error GoTo err
    ReDim Preserve arr(LBound(arr) To UBound(arr) + 1)
    On Error GoTo 0
    arr(UBound(arr)) = val
    Exit Sub
err:
    Debug.Print "poo"
End Sub

Function IsInArray(arr As Variant, valueToFind As Variant) As Boolean
' checks if valueToFind is found in arr, no loop!
On Error GoTo err
  IsInArray = (UBound(Filter(arr, valueToFind)) > -1)
  Exit Function
err:
    Debug.Print "poo"
End Function


