## Import .eml files into Outlook automatically, and retain folder structure

This is a VBScript that can import .eml files into Outlook maintaining folder names and structure.
Oringinally I was searching for a way to import a large amount of emails from Windows Live Mail (2011) that were in folders to move to Outlook 2019. Although there were lots of freeware programs online, most were limited to a few email transfers, and others didn't work on subfolders.

The original work was completed by Robert Sparnaaij aka Roady (https://www.howto-outlook.com/howto/import-eml-files.htm)
However, I have extended this version to allow the user to select a root folder containing folders of .eml files, and have these folders be created inside outlook, with the emails copied across to their respective folders.

# Usage:
1. Download 'import-eml-mod.vbs' to computer.
2. Modify line 34 of import-eml-mod.vbs so that "YOUR_ACCOUNT_HERE" is replaced by the email address in outlook you are wanting to add the emails to.
3. Open Outlook if it isn't already running.
4. Make sure Outlook is the default program for opening .eml files.
5. Run import-eml-mod.vbs and when the folder dialog opens, select the root folder containing the subfolders that each have the eml files inside.
6. Once you have selected this folder, do not move the mouse or use the keyboard as the program requires focus to be on the emails as they open in Outlook.

Each email will take approximately 1second to import, due to the sleep command in the script to allow enough time for the email to open. If your machine is slow, or really fast, adjust this sleep time as necessary.

# Example
If you have the following folder setup:
```
big_folder/
├─ some folder/
│  ├─ email1.eml
├─ my folder 2/
│  ├─ email1.eml
│  ├─ another_email.eml
│  ├─ something else.eml
├─ another folder/
```

Then you would select 'big_folder' as the root folder in the folder picker dialog. Then the program will create the folders 'some folder', 'my folder 2' and 'another folder', with the respective files in each one. Note that 'another folder' will be empty.
Also note that these folders will be placed underneath the 'Inbox' folder - this can be changed by changing line 36 according to [the docs](https://docs.microsoft.com/en-us/office/vba/api/outlook.oldefaultfolders).

The inbox will then look like this:

```
YOUR_EMAIL_ACCOUNT
Inbox/
├─ some folder/
│  ├─ email1.eml
├─ my folder 2/
│  ├─ email1.eml
│  ├─ another_email.eml
│  ├─ something else.eml
├─ another folder/
```

Tested on Windows 10 and Outlook 2019.
No responsibility or guarantees made with the use of this code. Use at your own risk.

