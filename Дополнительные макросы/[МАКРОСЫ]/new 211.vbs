Const WORD_DOCS = "Word Documents (*.doc)|*.doc|"         ' Prep for file types
Const ALL_FILES = "All Files (*.*)|*.*|"                  ' Also include all files
Set objDialog = CreateObject("UserAccounts.CommonDialog") ' Use the UserAccounts Common Dialog

objDialog.Filter = WORD_DOCS & ALL_FILES                  ' Apply the filename filters created above
objDialog.FilterIndex = 1                                 ' Default is first item in filter list
objDialog.InitialDir = "%homepath%\My Documents"          ' Starting folder/directory

intResult = objDialog.ShowOpen                            ' Open Dialog and return selected filename to user

If intResult Then
   Wscript.Echo objDialog.FileName			   ' Tell user what you did.
Else
   Wscript.Echo ""
End If