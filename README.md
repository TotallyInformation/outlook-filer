# outlook-filer
A Simple interface for filing emails when you have many folders.

This utility can save vast amounts of time if, like me, you have many (hundreds of) folders and receive large amounts of email that must be kept and filed.

# Installation
 - Download the [latest release](https://github.com/TotallyInformation/outlook-filer/releases/latest)
 - Unpack the *.frm and *.frx files to somewhere convenient.
 - Open the VBA project editor - you can use the alt-f11 keyboard shortcut
 - Import the *.frm file
 - Insert a new Module
 - Add the following code to the new module:

```VB
Option Explicit

Sub FileToFolder()
    FolderSelectBox.Show
End Sub
```

 - Make sure that you have changed the settings to allow Outlook to run your code. 
   You may wish to create a self-signed code signing certificate and sign your VBA code with that.
 - Customise the Ribbon, adding the ```FileToFolder``` macro as a button wherever convenient.

# Basic Use
 - Select one or more emails, click on the macro button.
 - Type some characters to filter the full folder list if you want to.
 - Select a folder and click on the File button. Or cancel if you don't want to go ahead.

# Advanced Use - Filing Conversations
Sometimes, you will have already recieved an email as part of an ongoing conversation and will have already filed that.

If you have your Outlook view in a conversation view mode, you can expand the conversation & choose all (or some) of the emails in the conversation and file them all to the same location. If one of the emails in the conversation is already filed, the folder(s) filed to will appear in the left-hand list and you can choose one of those entries instead of a folder in the right-hand list.

# Advanced Use - Quickly Switch to Another View
When you have many folders, it can be difficult to spot where they all are in the list - especially if you have sub-folders.

This utility will help. If you can remember part of the folder name, click on the macro button, type in some characters to folder the list, select the required folder and click on the View button instead of File.

# License
This utility including the code and the documentation is released under an [Apache 2.0 license](https://github.com/TotallyInformation/outlook-filer/blob/master/LICENSE).

Please note that no warranty is implied or given regarding the suitability of this code. Please review and test the code before use. I am happy to receive issues related to the code which I will try to fix as time permits but cannot guarantee that the code will work.
