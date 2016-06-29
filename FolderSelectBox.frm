VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FolderSelectBox 
   Caption         =   "Select Folder for Filing"
   ClientHeight    =   5376
   ClientLeft      =   120
   ClientTop       =   472
   ClientWidth     =   10704
   OleObjectBlob   =   "FolderSelectBox.frx":0000
   StartUpPosition =   3  'Windows Default
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
'   5) List of recently selected folders
'
' Author: Julian Knight (Totally Information)
' Version: v1.3 20015-06-12
' History:
'   v1.4 20015-06-29 - Chg default location to Win default. Add auto-selects. Add recents list (not yet working)
'   v1.3 20015-06-12 - Add copy link to clipboard after moving
'   v1.2 20015-05-18 - Add double-click processing
'   v1.1 20015-05-12 - Various improvements - add filter to full folder list, add view button
'   v1.0 20015-05-08 - Initial Release

Option Explicit

' Define form global variables
Dim folderNames(0 To 99) As String
Dim maxNames As Long
Dim folderPaths(0 To 99) As String
Dim maxPaths As Long
Dim folderAllPaths(0 To 999) As Variant
Dim maxFAP As Long
Dim folderAllNames(0 To 999) As String
Dim maxFAN As Long
Dim mailbox As String
Dim exitDelay As Long ' seconds to delay closure of form to allow copy of link

Private Sub btnCancel_Click()
    ' Do nothing other than cancel everything
    Unload Me
End Sub

' Only change the current view to the selected folder
Private Sub btnView_Click()
    Dim fldr As Outlook.MAPIFolder
    Dim objItem As Outlook.MailItem
    
    Set fldr = fldrDest
    ' If anywhere selected, change the explorer view now
    If IsObject(fldr) Then
        Set Application.ActiveExplorer.CurrentFolder = fldr
    End If
    
    ' End
    Set fldr = Nothing
    Set objItem = Nothing
    Unload Me
End Sub

' Do the move
Private Sub btnFileToFolder_Click()
    Dim fldr As Outlook.MAPIFolder
    Dim objItem As Outlook.MailItem
    Dim x As Long
    
    Set fldr = fldrDest
    On Error GoTo err
    ' If anywhere to move to, move each email now
    If IsObject(fldr) Then
        ' Also add the selected destination to the top of the recents list
        lstRecent.AddItem fldr.Name
        x = 0
        For Each objItem In ActiveExplorer.Selection
            ' Only move items not already in the dest folder
            If objItem.Parent.Name <> fldr.Name Then
                objItem.Move fldr
                x = x + 1
                AddLinkToMessage objItem
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
    ' Delay exit to allow time to copy the new link
    WaitFor (5)
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

Private Sub lstAllFolders_Change()
    
    'Deselect the previously selected folder
    lstFolders.Selected(0) = False
    
End Sub

Private Sub lstAllFolders_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    
    'Deselect the previously selected folder
    lstFolders.Selected(0) = False
    
    Call btnFileToFolder_Click
    
End Sub

Private Sub lstFolders_Change()
    
    'deselect the first from the all folders list
    lstAllFolders.Selected(0) = False
    
End Sub

Private Sub lstFolders_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    
    'deselect the first from the all folders list
    lstAllFolders.Selected(0) = False
    
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

' If the text box contents change, begin filtering
Private Sub tbFilterAllFolders_Change()
    
    With Me.tbFilterAllFolders
        If .Value = vbNullString Then
            Me.lstAllFolders.List = folderAllNames
        Else
            Me.lstAllFolders.List = Filter(SourceArray:=folderAllNames, match:=.Value, Compare:=vbTextCompare)
        End If
    End With
    
    'When filtering, select the first from the all folders list
    lstAllFolders.Selected(0) = True
    'Deselect the previously selected folder
    lstFolders.Selected(0) = False

End Sub

' Set up the form
Private Sub UserForm_Initialize()

    Dim objItem As Object
    Dim i As Long
    Dim numSelected As Long
    Dim numEmailsSelected As Long
    Dim mb
    
    Dim x As Object
    Set x = Application
    
    'Start Userform Centered inside Excel Screen (for dual monitors)
    ' From http://www.thespreadsheetguru.com/the-code-vault/launch-vba-userforms-in-correct-window-with-dual-monitors
    'Me.StartUpPosition = 3
    'Me.Left = Application.ActiveWindow.Left + (0.5 * Application.ActiveWindow.Width) - (0.5 * Me.Width)
    'Me.Top = Application.ActiveWindow.Top + (0.5 * Application.ActiveWindow.Height) - (0.5 * Me.Height)
    'Debug.Print Me.Left, Me.Top
    
    'List of accounts: Application.Session.Accounts
    
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
            ' How many items?
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
    
    ' Show the list of folders where any of the selected items are already filed
    lstFolders.List = folderNames
    
    ' a selected email already filed so pre-select the first folder in that list
    If i > 0 Then
      lstFolders.Selected(0) = True
    End If
    
    'Create the AllFolders list
    maxFAP = 0
    maxFAN = 0
    ProcessFolder Application.Session.GetDefaultFolder(olFolderInbox).Parent
    
    'If no selected folder, select the first from the all folders list
    'Useful for filtering
    If lstFolders.Selected(0) = False Then
      lstAllFolders.Selected(0) = True
    End If
    
    Set objItem = Nothing
    
End Sub

' Create the all-folder list
Sub ProcessFolder(objStartFolder As Outlook.MAPIFolder, _
                  Optional blnRecurseSubFolders As Boolean = True, _
                  Optional strFolderPath As String = "", _
                  Optional strFolderName As String = "")

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

            ' Add to the All Folder List
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

' Create an outlook: link to a msg
Sub AddLinkToMessage(objMail As Outlook.MailItem)
    'Dim objMail As Object
    'was earlier Outlook.MailItem
    'Dim doClipboard As New DataObject
    Dim txt As String
      
    'One and ONLY one message muse be selected
    'Set objMail = Application.ActiveExplorer.Selection.Item(1)
    
    'If objMail.Class = olMail Then
    '    txt = "outlook:" + objMail.EntryID + "][MESSAGE: " + objMail.Subject + " (" + objMail.SenderName + ")"
    'ElseIf objMail.Class = olAppointment Then
    '    txt = "outlook:" + objMail.EntryID + "][MEETING: " + objMail.Subject + " (" + objMail.Organizer + ")"
    'ElseIf objMail.Class = olTask Then
    '    txt = "outlook:" + objMail.EntryID + "][TASK: " + objMail.Subject + " (" + objMail.Owner + ")>"
    'ElseIf objMail.Class = olContact Then
    '    txt = "outlook:" + objMail.EntryID + "][CONTACT: " + objMail.Subject + " (" + objMail.FullName + ")"
    'ElseIf objMail.Class = olJournal Then
    '    txt = "outlook:" + objMail.EntryID + "][JOURNAL: " + objMail.Subject + " (" + objMail.Type + ")"
    'ElseIf objMail.Class = olNote Then
    '    txt = "outlook:" + objMail.EntryID + "][NOTE: " + objMail.Subject + " (" + " " + ")"
    'Else
    '    txt = "outlook:" + objMail.EntryID + "][ITEM: " + objMail.Subject + " (" + objMail.MessageClass + ")"
    'End If
    
    txt = "outlook:" + objMail.EntryID
    
    ' Replace all spaces with %20
    CopyTextToClipboard (Replace(txt, " ", "%20"))
    
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

' @see: http://www.thespreadsheetguru.com/blog/2015/1/13/how-to-use-vba-code-to-copy-text-to-the-clipboard
Sub CopyTextToClipboard(txt As String)
    'PURPOSE: Copy a given text to the clipboard (using DataObject)
    'SOURCE: www.TheSpreadsheetGuru.com
    'NOTES: Must enable Forms Library: Checkmark Tools > References > Microsoft Forms 2.0 Object Library
    
    Dim obj As New DataObject
    
    'Make object's text equal above string variable
    obj.SetText txt
    
    'Place DataObject's text into the Clipboard
    ' >> WARNING: Not working in Windows 8.1! Just get "??" instead of content <<
    'obj.PutInClipboard
    tbLink.Value = txt
    
    'Notify User
    'MsgBox txt, vbInformation

End Sub

Sub WaitFor(NumOfSeconds As Long)
    Dim SngSec As Long
    SngSec = Timer + NumOfSeconds
    
    Do While Timer < SngSec
        DoEvents
    Loop

End Sub
