Const GeneratedItemFlag = &h4000

dim shellApp 
dim folderBrowseDialog
dim filePath
set shellApp = CreateObject("Shell.Application")

set folderBrowseDialog = shellApp.BrowseForFolder(0,"Select the file", GeneratedItemFlag, "c:\")


if folderBrowseDialog is nothing then
       msgbox "No file was selected.  This will now terminate."
       Wscript.Quit
else
       filePath= folderBrowseDialog.self.path
end if