'===================================================================
'Description: VBS script to import eml-files.
'
'Comment: Before executing the vbs-file, make sure that Outlook is
'         configured to open eml-files.
'         Depending on the performance of your computer, you may
'         need to increase the Wscript.Sleep value to give Outlook
'         more time to open the eml-file.
'
' author : Robert Sparnaaij
' version: 1.1
' website: http://www.howto-outlook.com/howto/import-eml-files.htm
'
' *******
' This version modified by Tim in Oct 2021 to add functionality
' that will allow importing from multiple folders that contain the
' eml files. Will create these folders in outlook and copy across
' the contents, maintaining the strucutre.
' Note: will only work with one sublevel of folders from the root
' folder specified. 
' *******
'===================================================================

Dim objShell : Set objShell = CreateObject("Shell.Application")
Dim objFolder : Set objFolder = objShell.BrowseForFolder(0, "Select the root folder containing folders of eml-files", 0)
Set objFSO = CreateObject("Scripting.FileSystemObject")
Dim Item
Dim i : i = 0
If (NOT objFolder is Nothing) Then
  Set rootPath = objFSO.GetFolder(objFolder.Self.Path)
  Set WShell = CreateObject("WScript.Shell")
  Set objOutlook = CreateObject("Outlook.Application")
  Set oNamespace = objOutlook.GetNamespace("MAPI")
  Set oRecipient = oNamespace.CreateRecipient("YOUR_ACCOUNT_HERE")
  Call oRecipient.Resolve()
  Set oInbox = oNamespace.GetDefaultFolder(6)

  For Each subFolder in rootPath.SubFolders
	Call oInbox.Folders.Add(subFolder.Name)
	Set curObjFolder = objFSO.GetFolder(Subfolder.Path)
	Set colFiles = curObjFolder.Files
	For Each objFile in colFiles
	  objShell.ShellExecute objFile.Path, "", "", "open", 1
	  WScript.Sleep 1000
	  Set MyInspector = objOutlook.ActiveInspector
	  Set MyItem = objOutlook.ActiveInspector.CurrentItem
	  MyItem.Move oInbox.folders(subFolder.Name)
    Next
  Next

Else
  MsgBox "cancelled", 64, "Import"
End If

Set objFolder = Nothing
Set objShell = Nothing