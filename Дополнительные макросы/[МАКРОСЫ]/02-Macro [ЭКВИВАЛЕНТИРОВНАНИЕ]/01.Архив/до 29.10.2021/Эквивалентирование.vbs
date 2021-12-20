' 	 Макрос для формирования Динамической Расчётной Модели (ДРМ) - 2020
' 
' 1. Эквивалентирование - БРМ (корректировка БРМ)
' 2. Устранение ошибок после эквивалентирования (удаление узлов без связи, удаление УШР и Реакторов без узлов ... )
' 3. Заполнение актуального Динамического набора из Excel файла 
'
'**************************************************************************

r=Setlocale("en-us")
rrr=1

Set spShell = CreateObject("WScript.Shell")
Set FSO = CreateObject("Scripting.FileSystemObject")
Set t = Rastr

Shablon = spShell.SpecialFolders("MyDocuments") & "\RastrWin3\SHABLON\" ' ссылка на папку с Шаблонами

t.Printp("Запуск макроса " & "дата: " & date() & " | время: " & Hour(Now()) & " hour " & Minute(Now()) & " minut")

FileRastr = FolderAndMyFile  ' - Диалоговое окно выбора файла RastrWin3.
' FileExcelDynamicSet = FolderAndMyFile ' - Диалоговое окно выбора файла Excel.
Call DateInFile(FileRastr)
' Call DateInFile(FileExcelDynamicSet)
FileRastrName = FSO.GetFileName(FileRastr)
PathFileRastr = FSO.GetParentFolderName(FileRastr)

LinkCustomModels = "C:\CustomModels\"

SplitNameFile = Split(FileRastrName, ".")
NameFileRastr = SplitNameFile(0)
NameExpansion = SplitNameFile(1)

t.NewFile(Shablon & "режим.rg2") ' - создание нового файла RastrWin3.
t.Load RG_REPL, PathFileRastr & "\"& NameFileRastr & ".rg2", Shablon & "режим.rg2"  

VisibelExcelSet = True ' Настройка Excel: показывать Excel при заполнении (запуске).

'\\ 1.Запуск функции эквивалентирования:
Call CorrNA()' - корр. номеров районов АИП
Call Equivalence() ' - эквивалентирование БРМ.

'\\ 2.1.Сохраняем файл rg2 
t.Save(PathFileRastr & "\"& NameFileRastr & "_экв" & ".rg2"),(Shablon & "режим.rg2")

t.Printp("Завершение работы макроса " & "дата: " & date() & " | время: " & Hour(Now()) & " hour " & Minute(Now()) & " minut")
t.Printp "Заполнение модели ДРМ - завершено! (=_=)"


'\\*************************************************************************************************************************************************
Sub Ekvivalent_smart(vyborka_rayon)
    Set vet=t.tables("vetv")
    Set uzl=t.tables("node")
    Set ray=t.tables("area")
    Set gen=t.tables("Generator")
    Set pqd=t.Tables("graphik2")
    Set graphikIT=t.Tables("graphikIT")
    Set area=t.Tables("area")
    Set area2=t.Tables("area2")
    Set darea=t.Tables("darea")
    Set polin=t.Tables("polin")
    Set Reactors=t.Tables("Reactors")
	
	vet.SetSel("")
    vet.cols("sel").calc("0")
    uzl.SetSel("")
    uzl.cols("sel").calc("0")
	
    t.Tables("com_ekviv").Cols("zmax").z(0) = 1000
    t.Tables("com_ekviv").Cols("ek_sh").z(0) = 0
    t.Tables("com_ekviv").Cols("otm_n").z(0) = 0
    t.Tables("com_ekviv").Cols("smart").z(0) = 1
    t.Tables("com_ekviv").Cols("tip_ekv").z(0) = 0
    t.Tables("com_ekviv").Cols("ekvgen").z(0) = 0
    t.Tables("com_ekviv").Cols("tip_gen").z(0) = 1
	
    uzl.Setsel(vyborka_rayon)
    uzl.cols("sel").calc("1")
	
    t.Ekv ""
End Sub

Sub Ekvivalent_siln(vyborka_rayon2)
    Set vet=t.tables("vetv")
		Set uzl=t.tables("node")
		Set ray=t.tables("area")
		Set gen=t.tables("Generator")
		Set pqd=t.Tables("graphik2")
		Set graphikIT=t.Tables("graphikIT")
		Set area=t.Tables("area")
		Set area2=t.Tables("area2")
		Set darea=t.Tables("darea")
		Set polin=t.Tables("polin")
		Set Reactors=t.Tables("Reactors")
	
	vet.SetSel("")
    vet.cols("sel").calc("0")
    uzl.SetSel("")
    uzl.cols("sel").calc("0")
	
    t.Tables("com_ekviv").Cols("zmax").z(0) = 1000
    t.Tables("com_ekviv").Cols("ek_sh").z(0) = 0
    t.Tables("com_ekviv").Cols("otm_n").z(0) = 0
    t.Tables("com_ekviv").Cols("smart").z(0) = 0
    t.Tables("com_ekviv").Cols("tip_ekv").z(0) = 0
    t.Tables("com_ekviv").Cols("ekvgen").z(0) = 0
    t.Tables("com_ekviv").Cols("tip_gen").z(0) = 1
	
    uzl.Setsel vyborka_rayon2
    uzl.cols("sel").calc("1")
    vet.SetSel("iq.sel=1 & ip.sel=0 & !sta")
    k = vet.FindNextSel(-1)
    
	While k<>(-1)
		iq1 = vet.Cols("iq").z(k)
		uzl.Setsel "ny=" & iq1
		k2 = uzl.FindNextSel(-1)
		If k2<>-1 Then
			uzl.cols("sel").z(k2)=0
		End If
		k = vet.FindNextSel(k)
    Wend
    
	vet.SetSel("iq.sel=0 & ip.sel=1 & !sta")
    k = vet.FindNextSel(-1)
    
	While k<>(-1)
		ip1 = vet.Cols("ip").z(k)
		uzl.Setsel "ny=" & ip1
		k2 = uzl.FindNextSel(-1)
		
		If k2<>-1 Then
			uzl.cols("sel").z(k2) = 0
		End If
		
		k = vet.FindNextSel(k)
    Wend
    t.Ekv("")
End Sub

Sub Ekv_gen(vyborka_gen)
    Set vet=t.tables("vetv")
		Set uzl=t.tables("node")
		Set ray=t.tables("area")
		Set gen=t.tables("Generator")
		Set pqd=t.Tables("graphik2")
		Set graphikIT=t.Tables("graphikIT")
		Set area=t.Tables("area")
		Set area2=t.Tables("area2")
		Set darea=t.Tables("darea")
		Set polin=t.Tables("polin")
		Set Reactors=t.Tables("Reactors")
	
	uzl.Setsel vyborka_gen
    k = uzl.FindNextSel(-1)
	
    While k<>(-1)
		ny1 = uzl.Cols("ny").z(k)
		vet.SetSel("(ip.uhom<110 & iq=" & ny1 & ") | (iq.uhom<110 & ip=" & ny1 & ")") 
		k2 = vet.FindNextSel(-1)
		
		While k2<>(-1)
			ip1 = vet.Cols("ip").z(k2)
			iq1 = vet.Cols("iq").z(k2)
			If ip1 = ny1 Then
				ny2 = iq1
			else
				ny2 = ip1
			End If
			
			uzl.Setsel "ny=" & ny2
			k3 = uzl.FindNextSel(-1)
			If k3<>-1 Then
				uzl.cols("sel").z(k3) = 1
			End If
			k2 = vet.FindNextSel(k2)
		Wend
		uzl.Setsel vyborka_gen
		k = uzl.FindNextSel(k)
    Wend
    
	t.Tables("com_ekviv").Cols("zmax").z(0) = 1000
    t.Tables("com_ekviv").Cols("ek_sh").z(0) = 0
    t.Tables("com_ekviv").Cols("otm_n").z(0) = 0
    t.Tables("com_ekviv").Cols("smart").z(0) = 0
    t.Tables("com_ekviv").Cols("tip_ekv").z(0) = 0
    t.Tables("com_ekviv").Cols("ekvgen").z(0) = 0
    t.Tables("com_ekviv").Cols("tip_gen").z(0) = 1
	
    t.Ekv("")
    uzl.Setsel "uhom>50"
    uzl.cols("sel").calc("0")
    t.Ekv("")
    uzl.Setsel "uhom>50"
    uzl.cols("sel").calc("0")
    t.Ekv("")
    uzl.Setsel "uhom>50"
    uzl.cols("sel").calc("0")
    t.Ekv("")
    uzl.Setsel "uhom>50"
    uzl.cols("sel").calc("0")
    t.Ekv("")
    uzl.Setsel "uhom>50"
    uzl.cols("sel").calc("0")
    t.Ekv("")
    uzl.Setsel "uhom>50"
    uzl.cols("sel").calc("0")
    t.Ekv("")
End Sub

Sub ibnulenie(alpha)
	Set vet=t.tables("vetv")
	Set uzl=t.tables("node")
	vet.SetSel("")
	vet.cols("sel").calc("0")
	uzl.SetSel("")
	uzl.cols("sel").calc("0")
End Sub

Sub Vikluchatel(alpha)
    Set vet=t.tables("vetv")
		Set uzl=t.tables("node")
		Set ray=t.tables("area")
		Set gen=t.tables("Generator")
		Set pqd=t.Tables("graphik2")
		Set graphikIT=t.Tables("graphikIT")
		Set area=t.Tables("area")
		Set area2=t.Tables("area2")
		Set darea=t.Tables("darea")
		Set polin=t.Tables("polin")
		Set Reactors=t.Tables("Reactors")
		Set cvzd=uzl.Cols("vzd")
		Set csel=uzl.Cols("sel")
		Set cip=vet.cols("ip") 
		Set ciq=vet.cols("iq") 
	
	Dim nyplus(10000,8),vetmassiv(15000,3),nodes(15000)
	
	uzl.SetSel("na<500 | na>600")
    uzl.cols("sel").calc(1)
    vet.SetSel("iq.sel=1 &ip.sel=0 &!sta")
    k = vet.FindNextSel(-1)
    
	While k<>(-1)
		iq1 = vet.Cols("iq").z(k)
		uzl.Setsel "ny=" & iq1
		k2 = uzl.FindNextSel(-1)
		
		If k2<>(-1) Then
			uzl.cols("sel").z(k2) = 0
		End If
		
		k = vet.FindNextSel(k)
    Wend
	
    vet.SetSel("iq.sel=0 & ip.sel & !sta")
    k = vet.FindNextSel(-1)
	
    While k<>(-1)
		ip1 = vet.Cols("ip").z(k)
		uzl.Setsel "ny=" & ip1
		k2 = uzl.FindNextSel(-1)
		If k2<>(-1) Then
			uzl.cols("sel").z(k2) = 0
		End If
		k = vet.FindNextSel(k)
	Wend
    
	vet.SetSel("(iq.sel=1 &ip.sel=0) | (ip.sel=1 &iq.sel=0) & tip=2")
    k = vet.FindNextSel(-1)
    While k<>(-1)
		iq1 = vet.Cols("iq").z(k)
		uzl.Setsel "ny=" & iq1
		k2 = uzl.FindNextSel(-1)
		
		If k2<>(-1) Then
			uzl.cols("sel").z(k2) = 0
		End If
		
		ip1 = vet.Cols("ip").z(k)
		uzl.Setsel "ny=" & ip1
		k2 = uzl.FindNextSel(-1)
		
		If k2<>(-1) Then
			uzl.cols("sel").z(k2) = 0
		End If
		vet.SetSel("(iq.sel=1 &ip.sel=0) | (ip.sel=1 &iq.sel=0) & tip=2")
		k = vet.FindNextSel(-1)
    Wend
	    
    vetvyklvybexc = "(iq.bsh>0 & ip.bsh=0) | (ip.bsh>0 & iq.bsh=0) | (iq.bshr>0 & ip.bshr=0) | (ip.bshr>0 & iq.bshr=0)| ip.sel=0 | iq.sel=0)"
    flvykl = 0
	vet.SetSel "1"
	vet.cols("groupid").calc(0)
	'vet.SetSel "x=666"
	'vet.cols("x").calc(665)
	vet.SetSel vetvyklvybexc
	vet.cols("groupid").calc(1)
	nvet = 0
	
	for povet = 0 to 10000
		vet.SetSel("x<0.01 & x>-0.01 & r<0.005 & r>=0 & (ktr=0 | ktr=1) & !sta & groupid!=1 & b<0.000005")  'Выборка ветвей, которые считаем выключателями
		'vet.SetSel("tip=2 & x<0.01 & x>-0.01 & r<0.005 & r>=0 & (ktr=0 | ktr=1) & !sta &groupid!=1 & b<0.000005")
		ivet = vet.FindNextSel(-1)
		If ivet = -1 Then exit for
            ip = vet.Cols("ip").z(ivet)
            iq = vet.Cols("iq").z(ivet)
            If ip > iq Then
                ny = iq 
                ndel = ip
            else 
                ny = ip
                ndel = iq
            End If
            ndny = 0
            ndndel = 0
			'Проверка на наличие узла из списка неудаляемых
            for inodee = 0 to nnod
                If 	ndel = nodes(inodee) Then ndndel = 1
                If 	ny = nodes(inodee) Then ndny = 1
                If (ndndel = 1) and (ndny = 1) Then exit for
            next
			' Меняем местами, так как удаляемый нельзя удалять, а неудаляемый можно ))
            If (ndndel = 0) and (ndny = 1) Then
                buff = ny
                ny = ndel
                ndel = buff
            End If
			
            If (ndndel = 0) or (ndny = 0) Then 'Если хотя бы один можно удалить
                flvykl = flvykl + 1
				uzl.SetSel("ny=" & ny)
				iny = uzl.FindNextSel(-1)
				uzl.SetSel("ny=" & ndel)
				idel = uzl.FindNextSel(-1)
				pgdel = uzl.cols("pg").z(idel)
				qgdel = uzl.cols("qg").z(idel)
				pndel = uzl.cols("pn").z(idel)
				qndel = uzl.cols("qn").z(idel)
				bshdel = uzl.cols("bsh").z(idel)
				gshdel = uzl.cols("gsh").z(idel)
				pgny = uzl.cols("pg").z(iny)
				qgny = uzl.cols("qg").z(iny)
				pnny = uzl.cols("pn").z(iny)
				qnny = uzl.cols("qn").z(iny)
				bshny = uzl.cols("bsh").z(iny)
				gshny = uzl.cols("gsh").z(iny)
				uzl.cols("pg").z(iny) = pgdel + pgny
				uzl.cols("qg").z(iny) = qgdel + qgny
				uzl.cols("pn").z(iny) = pndel + pnny
				uzl.cols("qn").z(iny) = qndel + qnny
				uzl.cols("bsh").z(iny) = bshdel + bshny
				uzl.cols("gsh").z(iny) = gshdel + gshny
				v1 = uzl.cols("vzd").z(iny)
				v2 = uzl.cols("vzd").z(idel)
				qmax1 = uzl.cols("qmax").z(iny)
				qmax2 = uzl.cols("qmax").z(idel)
				'writelog "Выключатели. #"& flvykl &". Оставляем узел ny= "&ny&". Удаляем узел ndel= "& ndel                     
				gen.Setsel("Node=" & ndel)
				igen = gen.FindNextSel(-1) 'Меняем узлы подключения генераторов
				
				If igen<>(-1) Then
					While igen<>(-1) 
						gen.cols("Node").z(igen) = ny
						igen = gen.FindNextSel(igen)
					Wend
				End If
				
				If (v1<>v2) and (v1>0.3) and (v2>0.3) and (qmax1 + qmax2)<>0 Then
					uzl.cols("vzd").z(iny) = (v1*qmax1+v2*qmax2)/(qmax1+qmax2) 'Делаем средневзвешенное по qmax напряжение
				End If
				
				If (v1=0) and (v2<>0) Then
					uzl.cols("vzd").z(iny) = v2
				End If
				
				If (v1<>0) and (v2<>0) Then
					uzl.cols("qmin").z(iny) = (uzl.cols("qmin").z(iny)) + (uzl.cols("qmin").z(idel))
					uzl.cols("qmax").z(iny) = qmax1 + qmax2
				End If

				If (v1=0) and (v2<>0) Then
					uzl.cols("qmin").z(iny) = uzl.cols("qmin").z(idel)
					uzl.cols("qmax").z(iny) = uzl.cols("qmax").z(idel)
				End If
				
				vet.SetSel("(ip=" & ip & "& iq=" & iq & ")|(iq=" & ip & "& ip=" & iq & ")")
				vet.delrows 'Удаляем ветвь	
				vet.SetSel("iq=" & ndel) 'Меняем узлы ветвей с удаляемым узлом)))
				vet.cols("iq").calc(ny)	
				vet.SetSel("ip=" & ndel)
				vet.cols("ip").calc(ny)	
				uzl.delrows 		' Удаляем узел
            Else 'Если ни одного нельзя удалить
                vet.SetSel("(ip=" & ip & "& iq=" & iq & ")|(iq=" & ip & "& ip=" & iq & ")")
                vet.cols("groupid").calc(1)
            End If
    next
        'writelog "Выключатели. Обработано "& flvykl &" штук."
        kod = t.rgm ("p")
        If kod<>0 Then
            msgbox "Regim do not exist"
            'writelog "!!! After vykldel Regim do not exist!!!!!!"		
        End If
End Sub

Sub Ukraine(alpha)
	Set vet=t.tables("vetv")
    Set uzl=t.tables("node")
    Set ray=t.tables("area")
    Set gen=t.tables("Generator")
    Set pqd=t.Tables("graphik2")
    Set graphikIT=t.Tables("graphikIT")
    Set area=t.Tables("area")
    Set area2=t.Tables("area2")
    Set darea=t.Tables("darea")
    Set polin=t.Tables("polin")
    Set Reactors=t.Tables("Reactors")
    
    vet.SetSel("")
    vet.cols("sel").calc("0")
    uzl.SetSel("")
    uzl.cols("sel").calc("0")
    t.Tables("com_ekviv").Cols("zmax").z(0)=1000
    t.Tables("com_ekviv").Cols("ek_sh").z(0)=0
    t.Tables("com_ekviv").Cols("otm_n").z(0)=0
    t.Tables("com_ekviv").Cols("smart").z(0)=0
    t.Tables("com_ekviv").Cols("tip_ekv").z(0)=0
    t.Tables("com_ekviv").Cols("ekvgen").z(0)=0
    t.Tables("com_ekviv").Cols("tip_gen").z(0)=1
    uzl.SetSel("")
    uzl.cols("sel").calc(0)
    vet.SetSel("(iq.na=803 & ip.na>300 & ip.na<400) ")
    k=vet.FindNextSel(-1)
    
	While k<>(-1)
		iq1=vet.Cols("iq").z(k)
		uzl.Setsel "ny="&iq1
		k2=uzl.FindNextSel(-1)
		If k2<>-1 Then
		uzl.cols("sel").z(k2)=1
		End If
		k=vet.FindNextSel(k)
    Wend
    vet.SetSel("(ip.na=803 & iq.na>300 & iq.na<400) ")
    k=vet.FindNextSel(-1)
    While k<>(-1)
    ip1=vet.Cols("ip").z(k)
    uzl.Setsel "ny="&ip1
    k2=uzl.FindNextSel(-1)
    If k2<>-1 Then
    uzl.cols("sel").z(k2)=1
    End If
    k=vet.FindNextSel(k)
    Wend
    vet.SetSel("((iq.sel=1 &ip.sel=0) | (ip.sel=1 &iq.sel=0)) & ip.na=803 & iq.na=803 &!sta")
    k=vet.FindNextSel(-1)
    While k<>(-1)
    iq1=vet.Cols("iq").z(k)
    uzl.Setsel "ny="&iq1
    k2=uzl.FindNextSel(-1)
    If k2<>-1 Then
    uzl.cols("sel").z(k2)=1
    End If
    ip1=vet.Cols("ip").z(k)
    uzl.Setsel "ny="&ip1
    k2=uzl.FindNextSel(-1)
    If k2<>-1 Then
    uzl.cols("sel").z(k2)=1
    End If
    vet.SetSel("((iq.sel=1 &ip.sel=0) | (ip.sel=1 &iq.sel=0)) & ip.na=803 & iq.na=803 &!sta")
    k=vet.FindNextSel(-1)
    Wend
    t.Ekv""
End Sub

Sub Udalenie(alpha)
    Set vet=t.tables("vetv")
		Set uzl=t.tables("node")
		Set ray=t.tables("area")
		Set gen=t.tables("Generator")
		Set pqd=t.Tables("graphik2")
		Set graphikIT=t.Tables("graphikIT")
		Set area=t.Tables("area")
		Set area2=t.Tables("area2")
		Set darea=t.Tables("darea")
		Set polin=t.Tables("polin")
		Set Reactors=t.Tables("Reactors")
	
	uzl.Setsel("")
    k2=uzl.FindNextSel(-1)
	
    While k2<>(-1)
		ny1 = uzl.cols("ny").z(k2)
		vet.SetSel("((ip=" & ny1 & ") | (iq="&ny1 & "))" )
		If vet.count = 0 Then
			uzl.cols("sel").z(k2) = 1
		End If
		k2 = uzl.FindNextSel(k2)
    Wend
	
    uzl.Setsel("sel=1")
    uzl.delrows
    Reactors.Setsel("")
    k2 = Reactors.FindNextSel(-1)
    
	While k2<>(-1)
		ny1 = Reactors.cols("Id1").z(k2)
		uzl.SetSel("(ny=" & ny1 & ") " )
		If uzl.count = 0 Then
			Reactors.cols("sel").z(k2) = 1
		End If
		k2 = Reactors.FindNextSel(k2)
    Wend
	
    Reactors.Setsel("sel=1")
    Reactors.delrows
    gen.Setsel("Node.na=0")
    gen.delrows
    graphikIT.Setsel("")
    k = graphikIT.FindNextSel(-1)
    
	While k<>(-1)
		nzav = graphikIT.cols("Num").z(k)
		vet.Setsel("n_it=" & nzav)
		k2 = vet.FindNextSel(-1)
		If k2<>-1 Then
		else
			graphikIT.cols("Num").z(k) = 0
		End If
		k = graphikIT.FindNextSel(k)
    Wend
	
    graphikIT.Setsel("Num=0")
    graphikIT.delrows
    area.Setsel("")
    k = area.FindNextSel(-1)
    
	While k<>(-1)
		na1 = area.cols("na").z(k)
		uzl.Setsel("na=" & na1)
		k2 = uzl.FindNextSel(-1)
		If k2<>-1 Then
		else
			area.cols("na").z(k)=0
		End If
		k = area.FindNextSel(k)
    Wend
	
    area.Setsel("na=0")
    area.delrows
    area2.Setsel("")
    k = area2.FindNextSel(-1)
    
	While k<>(-1)
		na1 = area2.cols("npa").z(k)
		uzl.Setsel("npa="&na1)
		k2 = uzl.FindNextSel(-1)
		
		If k2<>-1 Then
		else
			area2.cols("npa").z(k) = 0
		End If
		
		k = area2.FindNextSel(k)
    Wend
    
	area2.Setsel("npa=0")
    area2.delrows
    darea.Setsel("")
    k = darea.FindNextSel(-1)
    
	While k<>(-1)
		na1 = darea.cols("no").z(k)
		area.Setsel("no=" & na1)
		k2 = area.FindNextSel(-1)
		If k2<>-1 Then
		else
			darea.cols("no").z(k) = 0
		End If
		k = darea.FindNextSel(k)
    Wend
	
    darea.Setsel("no=0")
    darea.delrows
    polin.Setsel("")
    k = polin.FindNextSel(-1)
    
	While k<>(-1)
		nsx1 = polin.cols("nsx").z(k)
		uzl.Setsel("nsx=" & nsx1)
		k2 = uzl.FindNextSel(-1)
		
		If k2<>-1 Then
		else
			polin.cols("nsx").z(k) = 0
		End If
		
		k = polin.FindNextSel(k)
    Wend
	
    polin.Setsel("nsx=0")
    polin.delrows
    t.rgm "p"
End Sub

Function Equivalence()
    t.Printp("Запуск функции Эквивалентирования - Equivalence()")
    Set vet=t.tables("vetv")
		Set uzl=t.tables("node")
		Set ray=t.tables("area")
		Set gen=t.tables("Generator")
		Set pqd=t.Tables("graphik2")
		Set graphikIT=t.Tables("graphikIT")
		Set area=t.Tables("area")
		Set area2=t.Tables("area2")
		Set darea=t.Tables("darea")
		Set polin=t.Tables("polin")
		Set Reactors=t.Tables("Reactors")
	
    t.rgm("p")
    Call ibnulenie("")
    Call Vikluchatel("")
    Call ibnulenie("")
    Call Ukraine("")
    Call ibnulenie("")
    vyborka_rayon2 = "(na=407)"
    Call Ekvivalent_siln(vyborka_rayon2)
    Call ibnulenie("")
    
	vyborka_gen = "(((na!=108 & (na>100 & na<200))|(na>201 & na<400 & na!=311)| na=202 | na=203 | na=204 | na=205 | na=206 | na=207 | na=208 | na=209 | na=301 | na=302 | na=309 | na=312 | na=401 | na=402 | na=404 | na=405 | na=407 | na=409 | na=801 | na=804 | na=805 | na=806 | na=813 | na=819 | na=820 | na=821 | na=822 | (na>822 & na<834)) & (uhom=110 | uhom=220)"
	' vyborka_gen = "((na>100 & na<200 & na!=108)|(na>300 & na<400 & na!=311) | na=201 | na=203 | na=205 | na=208 | na=206 | na=805 | na=806 | na=807) & (uhom=110 | uhom=220) "
    
	Call Ekv_gen(vyborka_gen)
    Call ibnulenie("")
    vyborka_rayon = "(((na!=108 & (na>100 & na<200))|(na>201 & na<400 & na!=311)| na=202 | na=203 | na=204 | na=205 | na=206 | na=207 | na=208 | na=209 | na=301 | na=302 | na=309 | na=312 | na=401 | na=402 | na=404 | na=405 | na=407 | na=409 | na=801 | na=804 | na=805 | na=806 | na=813 | na=819 | na=820 | na=821 | na=822 | (na>822 & na<834)) & (uhom=110 | uhom=220)"
	' vyborka_rayon = "(((na!=108 & (na>100 & na<200))|(na>300 & na<400 & na!=311)| na=201 | na=203 | na=205 | na=208 | na=206| na=805 | na=806 | na=807)& (uhom=110 | uhom=220)"
    Call Ekvivalent_smart(vyborka_rayon)
    Call ibnulenie("")
    'Udalenie ""
	t.rgm("p")  ' расчет режима плоским стартом
	
End Function

Function GetFileDlgEx(sIniDir,sFilter,sTitle) 
	Set oDlg = CreateObject("WScript.Shell").Exec("mshta.exe ""about:<object id=d classid=clsid:3050f4e1-98b5-11cf-bb82-00aa00bdce0b></object><script>moveTo(0,-9999);eval(new ActiveXObject('Scripting.FileSystemObject').GetStandardStream(0).Read("&Len(sIniDir)+Len(sFilter)+Len(sTitle)+41&"));function window.onload(){var p=/[^\0]*/;new ActiveXObject('Scripting.FileSystemObject').GetStandardStream(1).Write(p.exec(d.object.openfiledlg(iniDir,null,filter,title)));close();}</script><hta:application showintaskbar=no />""") 
	oDlg.StdIn.Write "var iniDir='" & sIniDir & "';var filter='" & sFilter & "';var title='" & sTitle & "';" 
	GetFileDlgEx = oDlg.StdOut.ReadAll 
End Function

Function FolderAndMyFile() 
	Set fso = CreateObject("Scripting.FileSystemObject")
	CurrentDirectory = fso.GetAbsolutePathName(".")
	sIniDir = CurrentDirectory &"\Myfile.rg2" 
	sFilter = "Regim files(*.rg2)|*.rg2| Dynamic files(*.rst)|*.rst| Excel files(*.xlsm)|*.xlsm|" 
	sTitle = "Open RastrWin3/Excel file" 
	FolderAndMyFile = GetFileDlgEx(Replace(sIniDir,"\","\\"),sFilter,sTitle) 
End Function

Sub TargetCustomModelsToDocuments()
	Set spCustomModelMap = t.Tables("CustomDeviceMap")
	Set spModule = spCustomModelMap.Cols("Module")
	for i = 0 To spCustomModelMap.Size - 1
		module = split(spModule.ZS(i),"\")
		spModule.ZS(i) = "<DOCUMENTS>\CustomModels\DLL\" & module(Ubound(module))
	next
End Sub

Function FillingUnspecIfiedGenerators()' Заполнение незадпнных генераторов 
    Set gen = t.Tables("Generator")
    Set pnom = gen.Cols("Pnom")
    Set Pgen = gen.Cols("P")
    Set Qgen = gen.Cols("Q")
    Set Qmin = gen.Cols("Qmin")
    Set Qmax = gen.Cols("Qmax")
    Set Pmax = gen.Cols("Pmax")
    Set uzl = t.Tables("node")
	Set unom = gen.Cols("Ugnom")
	Set cosfi = gen.Cols("cosFi")
	Set Demp = gen.Cols("Demp")
	Set mj = gen.Cols("Mj")
	Set xd1	= gen.Cols("xd1")
    Set nodeg = gen.Cols("Node")
    Set ModelType = gen.Cols("ModelType")
	ii = 0
	t.Printp("Запуск функции - заполнение незаданных генераторов'! ( FillingUnspecIfiedGenerators() )")
	gen.SetSel "ModelType=0"
	jj=gen.FindNextSel(-1)
	While jj<>-1
		uzl.SetSel "ny=" & nodeg.Z(jj)
		j1=uzl.FindNextSel(-1)
		If j1<>-1 Then
			ModelType.Z(jj)=3
            Pmax2 = t.Tables("Generator").Cols("Pmax").Z(jj)
            Qmax2 = t.Tables("Generator").Cols("Qmax").Z(jj)
            If pnom.Z(jj) > 0 Then
                unom.Z(jj)=uzl.Cols("uhom").z(j1)
                cosfi.Z(jj)=0.85
                Demp.Z(jj)=5
                mj.Z(j2)=5*ABS(pnom.Z(jj))/cosfi.Z(jj)
                xd1.Z(jj)=0.3*unom.Z(jj)*unom.Z(jj)*cosfi.Z(jj)/ABS(pnom.Z(jj))
                ii = ii + 1
                If Pgen.Z(jj) > pnom.Z(jj) Then
                    Pgen.Z(jj) = pnom.Z(jj)
                End If 
             End If 
             If pnom.Z(jj) < 0 Then
                unom.Z(jj)=uzl.Cols("uhom").z(j1)
                cosfi.Z(jj)=0.85
                Demp.Z(jj)=5
                mj.Z(j2)=5*ABS(pnom.Z(jj))/cosfi.Z(jj)
                xd1.Z(jj)=0.3*unom.Z(jj)*unom.Z(jj)*cosfi.Z(jj)/ABS(pnom.Z(jj))
                ii = ii + 1
             End If 
             If pnom.Z(jj) = 0 Then
                pnom.Z(jj) = 10
                unom.Z(jj)=uzl.Cols("uhom").z(j1)
                cosfi.Z(jj)=0.85
                Demp.Z(jj)=5
                mj.Z(j2)=5*ABS(pnom.Z(jj))/cosfi.Z(jj)
                xd1.Z(jj)=0.3*unom.Z(jj)*unom.Z(jj)*cosfi.Z(jj)/ABS(pnom.Z(jj))
                ii = ii + 1
             End If 
		End If
		gen.SetSel "ModelType=0"
		jj=gen.FindNextSel (jj)
	Wend
	t.Printp("Завершение работы функции - заполнение незаданных генераторов'! ( FillingUnspecIfiedGenerators() )")
End Function

Function DateInFile(FinleNameMsg)
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set File = FSO.GetFile(FinleNameMsg)
    Str = vbNullString
    Str = Str & "Был выбран файл: " & File.Type & vbCrLf
    Str = Str & "Дата создания - " & File.DateCreated & vbCrLf
    Str = Str & "Дата последнего доступа - " & File.DateLastAccessed & vbCrLf
    Str = Str & "Дата последней модификации - " & File.DateLastModIfied & vbCrLf
    Str = Str & "Диск - " & File.Drive.DriveLetter & vbCrLf
    Str = Str & "Имя - " & File.Name & vbCrLf
    Str = Str & "Родительский каталог - " & File.ParentFolder.Path & vbCrLf
    Str = Str & "Путь - " & File.Path & vbCrLf
    Str = Str & "Короткое имя - " & File.ShortName & vbCrLf
    Str = Str & "Путь в формате 8.3 - " & File.ShortPath & vbCrLf
    Str = Str & "Размер - " & File.Size & vbCrLf
    Str = Str & "Тип файла - " & File.Type
    t.Printp(Str)    
End Function 

Function CorrNA()
	Set vet=t.tables("vetv")
	Set uzl=t.tables("node")
	Set ray=t.tables("area")
	Set gen=t.tables("Generator")
	Set pqd=t.Tables("graphik2")
	Set graphikIT=t.Tables("graphikIT")
	Set area=t.Tables("area")
	Set area2=t.Tables("area2")
	Set darea=t.Tables("darea")
	Set polin=t.Tables("polin")

	uzl.SetSel ("na=832")
	uzl.cols("na").calc("510")

	uzl.SetSel ("na=834")
	uzl.cols("na").calc("803")

	uzl.SetSel ("na=833")
	uzl.cols("na").calc("106")

	uzl.SetSel ("na=831")
	uzl.cols("na").calc("110")

	uzl.SetSel ("na=829")
	uzl.cols("na").calc("803")

	uzl.SetSel ("na=826")
	uzl.cols("na").calc("803")

	uzl.SetSel ("na=825")
	uzl.cols("na").calc("803")

	uzl.SetSel ("na=827")
	uzl.cols("na").calc("803")

	uzl.SetSel ("na=828")
	uzl.cols("na").calc("803")

	uzl.SetSel ("na=824")
	uzl.cols("na").calc("805")

	uzl.SetSel ("na=823")
	uzl.cols("na").calc("403")


	uzl.SetSel ("na=821")
	uzl.cols("na").calc("403")

	uzl.SetSel ("na=820")
	uzl.cols("na").calc("403")

	uzl.SetSel ("na=819")
	uzl.cols("na").calc("807")

	uzl.SetSel ("na=813")
	uzl.cols("na").calc("201")

	uzl.SetSel ("na=822")
	uzl.cols("na").calc("813")

	uzl.SetSel ("na=0")
	uzl.cols("na").calc("ny*0.00001")

	gen.SetSel ("Node.sta=1")
	gen.cols("sta").calc(1)

	uzl.SetSel ("vzd=0 & qmax>0")
	uzl.cols("vzd").calc("uhom")
End Function