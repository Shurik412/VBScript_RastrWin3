
Set stdOut = WScript.StdOut
set robj = WScript.CreateObject("Astra.Rastr", "Rastr_")

Function Rastr_OnLog(code, level, id, name, index, description, formName)
    
	Select case code
		case 0
		   stdOut.WriteLine "[ERROR] " + description
		case 1
		   stdOut.WriteLine "[ERROR] " + description
		case 2
		   stdOut.WriteLine "[ERROR] " + description
		case 3
		   stdOut.WriteLine "[WARNING] " + description
		case 4
		   stdOut.WriteLine "[MESSAGE] " + description
		case 5
		   stdOut.WriteLine "[INFO] " + description
		case else
		   stdOut.WriteLine description
	End select
End Function

robj.Load 1, "D:\Documents\RastrWin3\test-rastr\cx195.rg2", ""

robj.rgm "p"
