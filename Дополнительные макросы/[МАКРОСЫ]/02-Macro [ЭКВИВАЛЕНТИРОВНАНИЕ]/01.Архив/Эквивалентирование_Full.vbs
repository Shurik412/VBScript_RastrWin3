' 	 Макрос для формирования Динамической Расчётной Модели (ДРМ) - 2020
' 
' 1. Эквивалентирование - БРМ (корректировка БРМ)
' 2. Устранение ошибок после эквивалентирования (удаление узлов без связи, удаление УШР и Реакторов без узлов ... )
' 3. Заполнение актуального Динамического набора из Excel файла 
'
'**************************************************************************

r=Setlocale("en-us")
rrr=1
Set t=Rastr

print("Запуск макроса " & "дата: " & date() & " | время: " & Hour(Now()) & " hour " & Minute(Now()) & " minut")
Time_1 = Timer()
Call Equivalence() ' - эквивалентирование БРМ.

Time_2 = Timer()
print(" - Время работы МАКРОСА, в минутах = " & ((Time_2 - Time_1)/60))

'\\************************************************************************
Sub Equivalence()
	print("-= Запуск эквивалентирования =-")
    Set node = t.Tables("node")
    Set vetv = t.Tables("vetv")
    Set gen = t.Tables("Generator")
    
    t.rgm("p")
    
    print("1. Эквивалентирование генераторов:")
	vyborka_gen = "(na=202 | na=203 | na=205 | na=206 | na=207 | na=208 | na=301 | na=302 | na=309 | na=401 | na=402 | na=404 | na=405 | na=407 | na=408 | na=801 | na=803 | na=804 | na=805 | na=806 | na=807 | na=813 | na=819 | na=822 | na=823 | na=825 | na=826 | na=827 | na=828 | na=830 | na=831 | na=832) & (uhom < 160)"
    viborka_ray_vikl = "(na=202 | na=203 | na=205 | na=206 | na=207 | na=208 | na=301 | na=302 | na=309 | na=401 | na=402 | na=404 | na=405 | na=407 | na=408 | na=801 | na=803 | na=804 | na=805 | na=806 | na=807 | na=813 | na=819 | na=822 | na=823 | na=825 | na=826 | na=827 | na=828 | na=830 | na=831 | na=832)"
    print("1.1. Выбрка ген.: " & vyborka_gen)
    print(" - Удаление всех выключателей по выборке!")
	Call Obnulenie()
    Time_Vikl_1 = Timer()
	Call Off_Gen_if_off_node()
    Call Vikluchatel(viborka_ray_vikl)
	Call Obnulenie()
	Call Delete_Gen_Vikl()
    Time_Vikl_2 = Timer()
    print(" - Время работы ' Call Vikluchatel(viborka_ray)', в минутах = " & ((Time_Vikl_2 - Time_Vikl_1)/60))
    Call Obnulenie()
    
	for i=1 to 7 
		Time_3 = Timer()
		Call Ekvivalent_Node_Gen(vyborka_gen)
		Call Off_line_one_on()
		Time_3_1 = Timer()
		print(" - Время работы 'Call Ekvivalent_Node_Gen(vyborka_gen)', в минутах = " & ((Time_3_1 - Time_3)/60))
	next
    'print("2. Эквивалентирование СМАРТ:")
    'vyborka_rayon = "((na>100 & na<200) & (uhom<230))"
    'print("2.1. Выбрка SMART: " & vyborka_rayon)
    
    'Time_4 = Timer()
    ' Call Ekvivalent_smart(vyborka_rayon)
    'Time_4_1 = Timer()
    'print(" - Время работы 'Call Ekvivalent_smart(vyborka_rayon)', в минутах = " & ((Time_4_1 - Time_4)/60))
    print("-= Завершение эквиваленитрования =-")
End Sub

Sub Ekvivalent_Node_Gen(vyborka_gen)
    Set Vetv = t.Tables("vetv")
    Set Node = t.Tables("node")
    Set Generator = t.Tables("Generator")
	
    Time_5 = Timer()
	 
	Node.SetSel(vyborka_gen) ' выборка по узлам
    Node.Cols("sel").Calc("1")
    j = Node.FindNextSel(-1) 

    While j<>(-1)
        ny = Node.Cols("ny").Z(j)
        tip_node = Node.Cols("tip").Z(j)
        uhom = Node.Cols("uhom").Z(j)
        If tip_node > 1 Then ' все генераторные узла
            Generator.SetSel("Node.ny=" & ny)
            j_gen = Generator.FindNextSel(-1)
            if j_gen <> (-1) then
                Vetv.SetSel("(ip= " & ny & ")|(iq= " & ny & ")")
                j_vetv = Vetv.FindNextSel(-1)
                while j_vetv <>(-1)
                    tip_vetv = Vetv.Cols("tip").Z(j_vetv)
                    if tip_vetv = 1 then
                        v_ip = Vetv.Cols("v_ip").Z(j_vetv) 
                        v_iq = Vetv.Cols("v_iq").Z(j_vetv)
                        if (v_ip > 170 and v_iq < 250) or (v_ip < 170 and v_iq > 250) then
                            Node.Cols("sel").Z(j) = 0
                        end if
                    end if
                    j_vetv = Vetv.FindNextSel(j_vetv)
                wend
            end if
        Else
            Vetv.SetSel("(ip= " & ny & ")|(iq= " & ny & ")")
                j_vetv_2 = Vetv.FindNextSel(-1)
                while j_vetv_2 <>(-1)
                    tip_vetv_2 = Vetv.Cols("tip").Z(j_vetv_2)
                    if tip_vetv_2 = 1 then
                        v_ip_2 = Vetv.Cols("v_ip").Z(j_vetv_2) 
                        v_iq_2 = Vetv.Cols("v_iq").Z(j_vetv_2)
                        if (v_ip_2 > 170 and v_iq_2 < 250) or (v_ip_2 < 170 and v_iq_2 > 250) then
                            Node.Cols("sel").Z(j) = 0
                        end if
                    end if
                    j_vetv_2 = Vetv.FindNextSel(j_vetv_2)
                wend
        End If
        Node.SetSel(vyborka_gen)
		j = Node.FindNextSel(j)
    Wend
    
    Time_5_1 = Timer()
    print(" - Время работы 'Ekv_gen(vyborka_gen) цикл WHILE', в минутах = " & ((Time_5_1 - Time_5)/60))
    
    ' Выставляет настройки эквивалентирования
    print("1.2. Выставляет настройки ген. эквив;")
	t.Tables("com_ekviv").Cols("zmax").z(0) = 1000
		t.Tables("com_ekviv").Cols("ek_sh").z(0) = 0
		t.Tables("com_ekviv").Cols("otm_n").z(0) = 0
		t.Tables("com_ekviv").Cols("smart").z(0) = 0
		t.Tables("com_ekviv").Cols("tip_ekv").z(0) = 0
		t.Tables("com_ekviv").Cols("ekvgen").z(0) = 0
		t.Tables("com_ekviv").Cols("tip_gen").z(0) = 1

    print("1.3.2. Эквивалентирует!")
    t.Ekv("")
    'print("1.3.3. Делает выборку: (uhom>50)(все узлы с Uном больше 50 кВ) и снимает отметки с них;")
    'Node.Setsel("uhom>50")
    'Node.cols("sel").calc("0")
End Sub

Sub Ekvivalent_smart(vyborka_rayon)
    Set vet=t.tables("vetv")
    Set uzl=t.tables("node")
    
    print("2.2.1. Снимает все отметки с ветвей;")
	
    vet.SetSel("")
    vet.cols("sel").calc("0")
    
    uzl.SetSel("")
    uzl.cols("sel").calc("0")
    
	print("2.2.2. Выставляет настройки для СМАРТ экв.;")
    
    t.Tables("com_ekviv").Cols("zmax").z(0) = 1000
		t.Tables("com_ekviv").Cols("ek_sh").z(0) = 0
		t.Tables("com_ekviv").Cols("otm_n").z(0) = 0
		t.Tables("com_ekviv").Cols("smart").z(0) = 1
		t.Tables("com_ekviv").Cols("tip_ekv").z(0) = 0
		t.Tables("com_ekviv").Cols("ekvgen").z(0) = 0
		t.Tables("com_ekviv").Cols("tip_gen").z(0) = 1
    
    print("2.2.3. Делает выборку по району: " & vyborka_rayon &";")
    
    uzl.Setsel(vyborka_rayon)
    uzl.cols("sel").calc("1")
    
    U_LIMIT=220
 
    print("2.2.4. Удаляет отметки из ТР: " & U_LIMIT & " кВ;")
    Call If_Vetv_Tr_otkl_new(U_LIMIT)
    
    print("2.2.5. Эквивалентирует!")
    ' t.Ekv("")
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

Sub If_Vetv_Tr_otkl_new()
    Set vetv = t.Tables("vetv")
    Set node = t.Tables("node")
    Set gen = t.Tables("Generator")
    ' 1. Берет все ветики в цикл FOR и проверяет тип ветви, если тип 1 (т.е. трансформаторная ветвь)
    ' 2. Находит узел начала и конца ветви, берет ном. напряжение этих узлов.
    ' 3. Далее: если напряжение одного из узлов равно 220 (то же только для 110),
    ' то с этих ветвей  и узлов (этих ветвей) снимаются отметки (т.е. они не участвуют в эквив-и)
    
    MaxRowVetv = vetv.Count-1
    
    For i=0 to MaxRowVetv
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
                v_ip = uhom_ip
                v_iq = uhom_iq
                                               
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
End sub

Sub print(str)
    t.Printp(str)
End Sub

Sub Vikluchatel(viborka_ray_vikl)
    Set vet=t.tables("vetv")
    Set uzl=t.tables("node")
    Set gen=t.tables("Generator")

    Dim nodes(30000)
	
	uzl.SetSel(viborka_ray_vikl) ' выборка узлов всех районов кроме 500 (Центра)
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
				
				If (v1<>v2) and (v1>0.3) and (v2>0.3) and (qmax1 + qmax2) <> 0 Then
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
End Sub

Sub Obnulenie()  ' обнуление всех sel (выделенных галочкой) УЗЛОВ и ВЕТВЕЙ
    Set node = t.Tables("node")
    Set vetv = t.Tables("vetv")
    
    vetv.SetSel("")
	vetv.cols("sel").calc("0")
	node.SetSel("")
	node.cols("sel").calc("0")
End Sub

Sub Delete_Gen_Vikl()
	set vet=t.tables("vetv")
	set uzl=t.tables("node")
	set gen=t.tables("Generator")
	set ti=t.Tables("ti")
	
	vet.SetSel("")
	vet.cols("sel").calc("0")
	uzl.SetSel("")
	uzl.cols("sel").calc("0")
	
	uzl.SetSel("")
	k1=uzl.findnextsel(-1)

	While k1<>(-1)
		ny1=uzl.Cols("ny").z(k1)
		vet.SetSel("(ip=" & ny1 &") |(iq=" & ny1 &")" )
		if vet.Count=1 then
			vet.SetSel("x<1 & (tip=0 | tip=2) & ((ip=" & ny1 & ") |(iq=" & ny1 &"))")
			if vet.Count=1 then
				vet.SetSel("x<1 & (tip=0 | tip=2) & ((ip=" & ny1 & ") |(iq=" & ny1 &"))" )
				k3=vet.findnextsel(-1)
				if k3<>(-1) then
					if vet.Cols("ip").z(k3)=ny1 then
						ny2=vet.Cols("iq").z(k3)
					else
						ny2=vet.Cols("ip").z(k3)
					end if
					gen.SetSel("Node=" & ny1)
					k2=gen.findnextsel(-1)
					if k2<>(-1) then
						uzl.SetSel("ny=" & ny2)
						k4=uzl.findnextsel(-1)
						if k4<>(-1) then
							uzl.Cols("pn").z(k4) = uzl.Cols("pn").z(k1) + uzl.Cols("pn").z(k1)
							uzl.Cols("qn").z(k4) = uzl.Cols("qn").z(k1) + uzl.Cols("qn").z(k1)
							uzl.Cols("vzd").z(k4) = uzl.Cols("vzd").z(k1)
							uzl.Cols("exist_load").z(k4) = uzl.Cols("exist_load").z(k1)
							uzl.Cols("exist_gen").z(k4) = uzl.Cols("exist_gen").z(k1)
							uzl.Cols("pn_max").z(k4) =uzl.Cols("pn_max").z(k1) + uzl.Cols("pn_max").z(k4)
							if uzl.Cols("pn_min").z(k4) => uzl.Cols("pn_min").z(k1) then
								uzl.Cols("pn_min").z(k4) = uzl.Cols("pn_min").z(k1)
							end if
							uzl.Cols("pg_max").z(k4) = uzl.Cols("pg_max").z(k1) + uzl.Cols("pg_max").z(k4)
							if uzl.Cols("pg_min").z(k4) => uzl.Cols("pg_min").z(k1) then
								uzl.Cols("pg_min").z(k4) = uzl.Cols("pg_min").z(k1)
							end if
							uzl.Cols("sel").z(k1) = 1
							vet.Cols("sel").z(k3) = 1
							' ti.SetSel("(prv_num=20 | prv_num=7 | prv_num=6 | prv_num=5 | prv_num=4 | prv_num=3 | prv_num=2 | prv_num=1) & id1="&ny1)
							' ti.cols("id1").calc(ny2)
							gen.SetSel("Node=" & ny1)
							k2 = gen.findnextsel(-1)
							while k2 <> (-1)
								gen.Cols("Node").z(k2) = ny2
								k2=gen.findnextsel(k2)
							wend
						end if
					end if
				end if
			end if
		end if
		uzl.SetSel("")
		k1=uzl.findnextsel(k1)
	Wend
	vet.SetSel("sel=1")
	vet.delrows
	uzl.SetSel("sel=1")
	uzl.delrows
End Sub

Sub Off_Gen_if_off_node()
	set uzl = t.tables("node")
	set gen = t.tables("Generator")

	gen.setsel("")
	k = gen.findnextsel(-1)
	while k<>(-1)
		nygen = gen.Cols("Node").z(k)
		uzl.SetSel "ny=" & nygen
		kk = uzl.findnextsel(-1)
		if kk <> (-1) then
			if uzl.cols("sta").z(kk) = True then
				gen.cols("sta").z(k) = 1
			end if
		end if
		gen.setsel("")
		k = gen.findnextsel(k)
	wend
End Sub

Sub Off_line_one_on()
	set vet=t.tables("vetv")
	Set staVetv = vet.Cols("sta")
	
	ii = 0
	VetvMaxRow = vet.Count-1
	for i = 0 to VetvMaxRow
		sta = staVetv.Z(i)
		If sta = 2 or sta = 3 Then
			staVetv.Z(i) = 1
			ii = ii + 1
		end if
	next
	print("Количество ЛЭП с односторонним вкл., переведенных с состояние полного откючения: " & ii)
End Sub