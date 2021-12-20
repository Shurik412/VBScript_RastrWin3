'Set Pm = CreateObject("StringSplit")
Dim arr
Spt = "8136422"

MsgBox SplitNameFile(Spt)

Function SplitNameFile(NameFile)
    Dim arr
    arr = Split(NameFile, "6")
    NameFile = arr(0)
    NameExpansion = arr(1)
    SplitNameFile = NameFile
End Function
