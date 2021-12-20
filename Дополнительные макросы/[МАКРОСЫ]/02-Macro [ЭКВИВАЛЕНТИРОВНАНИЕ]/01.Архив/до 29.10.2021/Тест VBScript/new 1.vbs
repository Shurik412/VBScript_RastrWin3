r=Setlocale("en-us")
rrr=1

Function FolderAndMyFile() 
	set fso = CreateObject("Scripting.FileSystemObject")
	CurrentDirectory = fso.GetAbsolutePathName(".")
	sIniDir = CurrentDirectory &"\Myfile.rg2" 
	sFilter = "Regim files(*.rg2)|*.rg2| Dynamic files(*.rst)|*.rst| Excel files(*.xlsm)|*.xlsm|" 
	sTitle = "Open RastrWin3/Excel file" 
	FolderAndMyFile = GetFileDlgEx(Replace(sIniDir,"\","\\"),sFilter,sTitle) 
End Function

Function GetFileDlgEx(sIniDir,sFilter,sTitle) 
	Set oDlg = CreateObject("WScript.Shell").Exec("mshta.exe ""about:<object id=d classid=clsid:3050f4e1-98b5-11cf-bb82-00aa00bdce0b></object><script>moveTo(0,-9999);eval(new ActiveXObject('Scripting.FileSystemObject').GetStandardStream(0).Read("&Len(sIniDir)+Len(sFilter)+Len(sTitle)+41&"));function window.onload(){var p=/[^\0]*/;new ActiveXObject('Scripting.FileSystemObject').GetStandardStream(1).Write(p.exec(d.object.openfiledlg(iniDir,null,filter,title)));close();}</script><hta:application showintaskbar=no />""") 
	oDlg.StdIn.Write "var iniDir='" & sIniDir & "';var filter='" & sFilter & "';var title='" & sTitle & "';" 
	GetFileDlgEx = oDlg.StdOut.ReadAll 
End Function

FileRastr = FolderAndMyFile
set fso = CreateObject("Scripting.FileSystemObject")
Set fl = fso.GetFile(FileRastr)
fl.Copy FileRastr & ".rg2"
Set fl2 = fso.GetFile(FileRastr & ".rg2")
fl2.Name = "DRM2020.rg2"


