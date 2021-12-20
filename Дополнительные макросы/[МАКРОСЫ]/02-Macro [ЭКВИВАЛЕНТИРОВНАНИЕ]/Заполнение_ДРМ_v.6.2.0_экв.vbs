' 	 Макрос для формирования Динамической Расчётной Модели (ДРМ) - 2020
' 
' 1. Эквивалентирование - БРМ (корректировка БРМ)
' 2. Устранение ошибок после эквивалентирования (удаление узлов без связи, удаление УШР и Реакторов без узлов ... )
' 3. Заполнение актуального Динамического набора из Excel файла 
'
'**************************************************************************

r=Setlocale("en-us")
rrr=1

Time_1 = Timer()

Set spShell = CreateObject("WScript.Shell")
Set FSO = CreateObject("Scripting.FileSystemObject")
Set t = Rastr
t.Printp("Time_1 = " & Time_1)
Shablon = spShell.SpecialFolders("MyDocuments") & "\RastrWin3\SHABLON\" ' ссылка на папку с Шаблонами

t.Printp("Запуск макроса " & "дата: " & date() & " | время: " & Hour(Now()) & " hour " & Minute(Now()) & " minut")

FileRastr = FolderAndMyFile  ' - Диалоговое окно выбора файла RastrWin3.
'FileExcelDynamicSet = FolderAndMyFile ' - Диалоговое окно выбора файла Excel.
'Call DateInFile(FileRastr)
'Call DateInFile(FileExcelDynamicSet)
FileRastrName = FSO.GetFileName(FileRastr)
PathFileRastr = FSO.GetParentFolderName(FileRastr)

' LinkCustomModels = "C:\CustomModels\"

SplitNameFile = Split(FileRastrName, ".")
NameFileRastr = SplitNameFile(0)
NameExpansion = SplitNameFile(1)

'VisibelExcelSet = True ' Настройка Excel: показывать Excel при заполнении (запуске).

flag = 1
If flag = 1 Then
    t.NewFile(Shablon & "режим.rg2") ' - создание нового файла RastrWin3.
    t.Load RG_REPL, PathFileRastr & "\" & NameFileRastr & ".rg2", Shablon & "режим.rg2" 
    flag_eqv = 1
    if flag_eqv = 1 then
        '\\ 1.Запуск функции эквивалентирования:
        flag_CorrNA = 0
        if flag_CorrNA = 1 Then
            Call CorrNA()' - корр. номеров районов АИП
        End if
        Call Equivalence() ' - эквивалентирование БРМ.
    end if
    t.rgm("")
    '\\ 2.1.Сохраняем файл rg2 
    t.Save(PathFileRastr & "\" & NameFileRastr & "_экв1" & ".rg2"),(Shablon & "режим.rg2")

    '\\ 2.2.Запуск функций исправления ошибок и предупреждений:
    ' Call DelNode() ' - удаление Узлов без ветвей.
    ' Call OFF_LEP_one_STA() ' - отключение односторонне включенных ветвей.
    ' Call DelReactor() ' - удаление Реакторов без узлов.
    ' Call OffGenP_Q_Zero() ' - откл. ген. Pген=0 и Qген=0.
    ' Call OffGenIfNodeSta() ' - откл. ген. с откл. узлами.
 
    ' Call DelUSHR() ' - удаление УШР без узлов.
    t.Save(PathFileRastr & "\"& NameFileRastr & "_экв2" & ".rg2"),(Shablon & "режим.rg2")
End If

t.Printp("Завершение работы макроса " & "дата: " & date() & " | время: " & Hour(Now()) & " hour " & Minute(Now()) & " minut")
t.Printp "Заполнение модели ДРМ - завершено! (=_=)"
Time_2 = Timer()
t.Printp("Time_2=" & Time_2)
t.Printp("Время работы, в минутах = " & ((Time_2 - Time_1)/60))


'\\*************************************************************************************************************************************************
Sub Equivalence()
	t.Printp("Запуск эквивалентирования:")
    Set node = t.Tables("node")
    Set vetv = t.Tables("vetv")
    Set gen = t.Tables("Generator")
    
    t.rgm("p")
    
    Call Obnulenie()  ' обнуление всех sel (выделенных галочкой) УЗЛОВ и ВЕТВЕЙ
    Call Vikluchatel()
	t.Printp("  - Удалены выключатели")
	
    ' Call Obnulenie()  ' обнуление всех sel (выделенных галочкой) УЗЛОВ и ВЕТВЕЙ
    ' Call Ukraine()
	' t.Printp("  - эквивлент. Украина")
	
    'Call Obnulenie()  ' обнуление всех sel (выделенных галочкой) УЗЛОВ и ВЕТВЕЙ
    'vyborka_rayon2 = "na=407"
    'Call Ekvivalent_siln(vyborka_rayon2)
    	't.Printp("  - сильное эквивалентирование")
	
    Call Obnulenie()  ' обнуление всех sel (выделенных галочкой) УЗЛОВ и ВЕТВЕЙ
    ' vyborka_gen = "((na>100 & na<200 & na!=108)|(na>300 & na<400 & na!=311 & na!=403) | na=201 | na=203 | na=205 | na=208 | na=206 | na=805 | na=806 | na=807 | na=813 | na=830) & (uhom=110 | uhom=220) "
    ' vyborka_gen = "((na>100 & na<200)| na=205 | na=309 | na=312 | na=407 | na=409 | na=801 | na=805 | na=806 | na=807 | na=819 | na=821 | na=829 | na=830) & (uhom=110 | uhom=220) "
    ' vyborka_gen = "((na>100 & na<200)| na=205 | na=309 | na=312 | na=407 | na=409 | na=801 | na=805 | na=806 | na=807 | na=819 | na=821 | na=829 | na=830) & (uhom=110 | uhom=220) "
    vyborka_gen = "((na>100 & na<200)| na=202 | na=203 | na=204 | na=205 | na=206 | na=207 | na=208 | na=209 | na=301 | na=302 | na=309 | na=311 | na=312 | na=401 | na=402 | na=404 | na=405 | na=407 | na=408 | na=409 | na=801 | na=803 | na=804 | na=805 | na=806 | na=807 | na=813 | na=819 | na=820 | na=821 | na=822 | na=823 | na=825 | na=826 | na=827 | na=828 | na=830 | na=831 | na=832) & (uhom=35 | uhom=110 | uhom=220) "
     ' vyborka_gen = "(na=102 | na=103 | na=104 | na=105| na=106 | na=107 | na=108 | na=109 | na=301 | na=302 | na=309 | na=311 | na=312 | na=401 | na=402 | na=403 | na=404 | na=405 | na=405 | na=407 | na=801 | na=803 | na=805 | na=806 | na=807 | na=819 | na=821 | na=829 | na=830) & (uhom < 150) &  (ny != 20125101)"
    Call Ekv_gen(vyborka_gen)
	t.Printp("  - эквивалентирование генераторов")

    Call Obnulenie()  ' обнуление всех sel (выделенных галочкой) УЗЛОВ и ВЕТВЕЙ
    vyborka_rayon = "((na>100 & na<200)| na=202 | na=203 | na=204 | na=205 | na=206 | na=207 | na=208 | na=209 | na=301 | na=302 | na=309 | na=311 | na=312 | na=401 | na=402 | na=404 | na=405 | na=407 | na=408 | na=409 | na=801 | na=803 | na=804 | na=805 | na=806 | na=807 | na=813 | na=819 | na=820 | na=821 | na=822 | na=823 | na=825 | na=826 | na=827 | na=828 | na=830 | na=831 | na=832) & (uhom=35 | uhom=110 | uhom=220) "
    'vyborka_rayon = "(na=102 | na=103 | na=104 | na=105| na=106 | na=107 | na=108 | na=109 | na=202 | na=203 | na=204 | na=205 | na=206 | na=207 | na=208 | na=209 | na=301 | na=302 | na=309 | na=311 | na=312 | na=401 | na=402 | na=404 | na=405 | na=405 | na=407 | na=408 | na=409 | na=801 | na=803 | na=804 | na=805 | na=806 | na=807 | na=813 | na=819 | na=820 | na=821 | na=822 | na=823 | na=824 | na=825 | na=826 | na=826 | na=827 | na=828 | na=829 | na=830 | na=831 | na=832) & (uhom=35 | uhom=110 | uhom=220) "    
    'vyborka_rayon = "(na=102 | na=103 | na=104 | na=105| na=106 | na=107 | na=108 | na=109 | na=110 | na=301 | na=205 | na=206 | na=208 | na=302 | na=309 | na=311 | na=312 | na=401 | na=402 | na=404 | na=405 | na=405 | na=407 | na=801 | na=803 | na=804 | na=805 | na=806 | na=807 | na=819 | na=821 | na=829 | na=830) & (uhom < 150) &  (ny != 20125101)"    
    Call Ekvivalent_smart(vyborka_rayon)
    
    t.printp("Завершение эквиваленитрования.")
End Sub

Sub Obnulenie()  ' обнуление всех sel (выделенных галочкой) УЗЛОВ и ВЕТВЕЙ
    Set node = t.Tables("node")
    Set vetv = t.Tables("vetv")
    
    vetv.SetSel("")
	vetv.cols("sel").calc("0")
	node.SetSel("")
	node.cols("sel").calc("0")
End Sub

Sub Vikluchatel()
    Set vet=t.tables("vetv")
    Set uzl=t.tables("node")
    Set gen=t.tables("Generator")
    
    Dim nodes(15000)
	
	uzl.SetSel("na=102|	na=103|	na=104|	na=105|	na=106|	na=107|	na=108|	na=109|	na=202|	na=203|	na=204|	na=205|	na=206|	na=207|	na=208|	na=209|	na=301|	na=302|	na=309|	na=311|	na=312|	na=401|	na=402|	na=404|	na=405|	na=407|	na=408|	na=409|	na=801|	na=803|	na=804|	na=805|	na=806|	na=807|	na=813|	na=819|	na=820|	na=821|	na=822|	na=823|	na=824|	na=825|	na=827|	na=828|	na=829|	na=830|	na=832") ' выборка узлов всех районов кроме 500 (Центра)
    uzl.cols("sel").calc(1) ' выделение выбраных узлов
    vet.SetSel("iq.sel=1 & ip.sel=0 &!sta") ' выборка ветвей iq.sel = 1 ...
    k = vet.FindNextSel(-1)
	While k<>(-1) ' убирает sel-узла если на ВЛ с одной стороны выделен узел 
		iq1 = vet.Cols("iq").z(k)
		uzl.Setsel("ny=" & iq1)
		k2 = uzl.FindNextSel(-1)
		If k2<>(-1) Then
			uzl.cols("sel").z(k2) = 0
		End If
		k = vet.FindNextSel(k)
    Wend

    vet.SetSel("iq.sel=0 & ip.sel=1 & !sta")
    k = vet.FindNextSel(-1)
	
    While k<>(-1) ' убирает sel-узла если на ВЛ с одной стороны выделен узел 
		ip1 = vet.Cols("ip").z(k)
		uzl.Setsel "ny=" & ip1
		k2 = uzl.FindNextSel(-1)
		If k2<>(-1) Then
			uzl.cols("sel").z(k2) = 0
		End If
		k = vet.FindNextSel(k)
	Wend
  
	vet.SetSel("(iq.sel=1 & ip.sel=0)|(ip.sel=1 & iq.sel=0) & tip=2") ' tip=2 - выключатели (выборка всех выключателей если хотябы с одной стороны выделен узел sel)
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
	vet.SetSel("1")
	vet.cols("groupid").calc(0)
	vet.SetSel(vetvyklvybexc)
	vet.cols("groupid").calc(1)
	nvet = 0
	' удаление выключателей
	for povet = 0 to 10000
		vet.SetSel("x<0.01 & x>-0.01 & r<0.005 & r>=0 & (ktr=0 | ktr=1) & !sta & groupid!=1 & b<0.000005")  'Выборка ветвей, которые считаем выключателями
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
    kod = t.rgm ("p")
    If kod<>0 Then
        msgbox "Regim do not exist"		
    End If
    If_Vetv_Tr_otkl()
End Sub

Sub Ukraine()
	Set vet=t.tables("vetv")
    Set uzl=t.tables("node")
    
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
    
    vet.SetSel("(iq.na=803 & (ip.na>300 | ip.na<400))")
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

    vet.SetSel("(iq.na=803 & (ip.na>300 | ip.na<400))")
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

    vet.SetSel("((iq.sel=1 & ip.sel=0)|(ip.sel=1 & iq.sel=0)) & ip.na=803 & iq.na=803 & !sta")
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
        vet.SetSel("((iq.sel=1 & ip.sel=0)|(ip.sel=1 & iq.sel=0)) & ip.na=803 & iq.na=803 & !sta")
        k=vet.FindNextSel(-1)
    Wend
    
    If_Vetv_Tr_otkl()
    t.Ekv""
End Sub

Sub Ekvivalent_siln(vyborka_rayon2)
    Set vet=t.tables("vetv")
	Set uzl=t.tables("node")
    
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
	
    uzl.Setsel(vyborka_rayon2)
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
    
    ' If_Vetv_Tr_otkl()
    Call If_Vetv_Tr_otkl_new(220)
	t.Ekv("")
End Sub

Sub Ekv_gen(vyborka_gen)
    Set vet=t.tables("vetv")
    Set uzl=t.tables("node")
	uzl.Setsel(vyborka_gen)
    k = uzl.FindNextSel(-1)
    While k<>(-1)
		ny1 = uzl.Cols("ny").z(k)
		vet.SetSel("(ip.uhom<110 & iq=" & ny1 & ")|(iq.uhom<110 & ip=" & ny1 & ")") 
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
		uzl.Setsel(vyborka_gen)
		k = uzl.FindNextSel(k)
    Wend
    
	t.Tables("com_ekviv").Cols("zmax").z(0) = 1000
    t.Tables("com_ekviv").Cols("ek_sh").z(0) = 0
    t.Tables("com_ekviv").Cols("otm_n").z(0) = 0
    t.Tables("com_ekviv").Cols("smart").z(0) = 0
    t.Tables("com_ekviv").Cols("tip_ekv").z(0) = 0
    t.Tables("com_ekviv").Cols("ekvgen").z(0) = 0
    t.Tables("com_ekviv").Cols("tip_gen").z(0) = 1
    
	Call If_Vetv_Tr_otkl()
    t.Ekv("")
    uzl.Setsel "uhom>50"
    uzl.cols("sel").calc("0")
    Call If_Vetv_Tr_otkl()
    t.Ekv("")
    uzl.Setsel "uhom>50"
    uzl.cols("sel").calc("0")
    Call If_Vetv_Tr_otkl()
    t.Ekv("")
    uzl.Setsel "uhom>50"
    uzl.cols("sel").calc("0")
    Call If_Vetv_Tr_otkl()
    t.Ekv("")
    uzl.Setsel "uhom>50"
    uzl.cols("sel").calc("0")
    Call If_Vetv_Tr_otkl()
    t.Ekv("")
    uzl.Setsel "uhom>50"
    uzl.cols("sel").calc("0")
    Call If_Vetv_Tr_otkl()
    t.Ekv("")
    uzl.Setsel "uhom>50"
    uzl.cols("sel").calc("0")
    Call If_Vetv_Tr_otkl()
    t.Ekv("")
End Sub

Sub Ekvivalent_smart(vyborka_rayon)
    Set vet=t.tables("vetv")
    Set uzl=t.tables("node")
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
    
    Call If_Vetv_Tr_otkl()
    t.Ekv("")
End Sub

Sub If_Vetv_Tr_otkl()
    Set vetv = t.Tables("vetv")
    Set node = t.Tables("node")
    Set gen = t.Tables("Generator")
    MaxRowVetv = vetv.Count
    For i=0 to MaxRowVetv-1
        type_vetv = vetv.Cols("tip").Z(i)
        if type_vetv = 1 then
            ny_ip = vetv.Cols("ip").Z(i)
            ny_iq = vetv.Cols("iq").Z(i)
            v_ip = vetv.Cols("v_ip").Z(i)
            v_iq = vetv.Cols("v_iq").Z(i)
            node.SetSel("ny=" & ny_ip)
            j_ny_ip = node.FindNextSel(-1)
            if j_ny_ip <>(-1) Then
                tip_ny_ip = node.Cols("tip").Z(j_ny_ip) ' тип узла
                if tip_ny_ip > 1 Then
                    vetv.Cols("sel").Z(i) = 0
                end if
            end if 
            
            node.SetSel("ny=" & ny_iq)
            j_ny_iq = node.FindNextSel(-1)
            if j_ny_iq <>(-1) Then
                tip_ny_iq = node.Cols("tip").Z(j_ny_iq) ' тип узла
                if tip_ny_iq > 1 Then
                    vetv.Cols("sel").Z(i) = 0
                end if
            end if
            
            gen.SetSel("Node=" & ny_ip)
            j_gen_ny_ip = gen.FindNextSel(-1)
            if j_gen_ny_ip<>(-1) Then
                vetv.Cols("sel").Z(i) = 0
                node.SetSel("ny=" & ny_ip)
                j_node_ip = node.FindNextSel(-1)
                if j_node_ip <> (-1) then
                    node.Cols("sel").Z(j_node_ip) = 0
                end if
            end if
            
            gen.SetSel("Node=" & ny_iq)
            j_gen_ny_iq = gen.FindNextSel(-1)
            if j_gen_ny_iq <>(-1) Then
                vetv.Cols("sel").Z(i) = 0 
                node.SetSel("ny=" & ny_iq)
                j_node_iq = node.FindNextSel(-1)
                if j_node_iq <> (-1) then
                    node.Cols("sel").Z(j_node_iq) = 0
                end if
            end if 
        end if 
    next
End sub

Sub If_Vetv_Tr_otkl_new(U_LIMIT)
    Set vetv = t.Tables("vetv")
    Set node = t.Tables("node")
    Set gen = t.Tables("Generator")
    
    MaxRowVetv = vetv.Count
    If U_LIMIT = 220 then
        For i=0 to MaxRowVetv-1
            type_vetv = vetv.Cols("tip").Z(i)
            if type_vetv = 1 then
                ny_ip = vetv.Cols("ip").Z(i) ' номер начала
                ny_iq = vetv.Cols("iq").Z(i) ' номер конца
                
                node.SetSel("ny=" & ny_ip)
                j_ny_ip = node.FindNextSel(-1)
                uhom_ip = node.Cols("uhom").Z(j_ny_ip)
                node.SetSel("")
                
                node.SetSel("ny=" & ny_iq)
                j_ny_iq = node.FindNextSel(-1)
                uhom_iq = node.Cols("uhom").Z(j_ny_iq)
                node.SetSel("")
                'v_ip = vetv.Cols("v_ip").Z(i) ' напряжение начала 
                'v_iq = vetv.Cols("v_iq").Z(i) ' напряжение конца
                v_ip = uhom_ip
                v_iq = uhom_iq
                
                ' t.Printp("v_ip = " & v_ip & " - v_iq = " & v_iq)
                
                if (v_ip = 220) or (v_iq = 220) then 
                    vetv.Cols("sel").Z(i) = 0

                    node.SetSel("ny=" & ny_ip)
                    j_ny_ip = node.FindNextSel(-1)
                    node.Cols("sel").Z(j_ny_ip) = 0
                    node.SetSel("")
                    
                    node.SetSel("ny=" & ny_iq)
                    j_ny_iq = node.FindNextSel(-1)
                    node.Cols("sel").Z(j_ny_iq) = 0
                    node.SetSel("")
                end if
            end if 
        next
      end if
      
      If U_LIMIT = 110 then
        For i=0 to MaxRowVetv-1
            type_vetv = vetv.Cols("tip").Z(i)
            if type_vetv = 1 then
                ny_ip = vetv.Cols("ip").Z(i) ' номер начала
                ny_iq = vetv.Cols("iq").Z(i) ' номер конца
                
                node.SetSel("ny=" & ny_ip)
                j_ny_ip = node.FindNextSel(-1)
                uhom_ip = node.Cols("uhom").Z(j_ny_ip)
                node.SetSel("")
                
                node.SetSel("ny=" & ny_iq)
                j_ny_iq = node.FindNextSel(-1)
                uhom_iq = node.Cols("uhom").Z(j_ny_iq)
                node.SetSel("")
                'v_ip = vetv.Cols("v_ip").Z(i) ' напряжение начала 
                'v_iq = vetv.Cols("v_iq").Z(i) ' напряжение конца
                v_ip = uhom_ip
                v_iq = uhom_iq
                
                ' t.Printp("v_ip = " & v_ip & " - v_iq = " & v_iq)
                flag = 0
                if flag = 1 then
                    if (v_ip = 220) or (v_iq = 220) then 
                        vetv.Cols("sel").Z(i) = 0

                        node.SetSel("ny=" & ny_ip)
                        j_ny_ip = node.FindNextSel(-1)
                        node.Cols("sel").Z(j_ny_ip) = 0
                        node.SetSel("")
                        
                        node.SetSel("ny=" & ny_iq)
                        j_ny_iq = node.FindNextSel(-1)
                        node.Cols("sel").Z(j_ny_iq) = 0
                        node.SetSel("")
                    end if
                end if
                
                if (v_ip = 110) or (v_iq = 110) then 
                    vetv.Cols("sel").Z(i) = 0

                    node.SetSel("ny=" & ny_ip)
                    j_ny_ip = node.FindNextSel(-1)
                    node.Cols("sel").Z(j_ny_ip) = 0
                    node.SetSel("")
                    
                    node.SetSel("ny=" & ny_iq)
                    j_ny_iq = node.FindNextSel(-1)
                    node.Cols("sel").Z(j_ny_iq) = 0
                    node.SetSel("")
                end if
            end if 
        next
      end if
End sub

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

Function ModelIndexByType(strType)
    Set spIEEEExciters = t.Tables("DFWIEEE421")
    Set spType = spIEEEExciters.Cols("ModelType")
    ModelIndexByType = 0
	for each enumType in split(spType.Prop(FL_NAMEREF),"|")
		If enumType = strType Then Exit For
        ModelIndexByType = ModelIndexByType + 1
	next
End function

Function CorrNA()
	Set uzl=t.tables("node")
    Set gen = t.tables("Generator")
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