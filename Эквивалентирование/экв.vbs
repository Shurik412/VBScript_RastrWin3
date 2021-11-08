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
print(" Время работы МАКРОСА, в минутах = " & ((Time_2 - Time_1)/60))
print("Работа макроса завершена.")

'\\************************************************************************
Sub Equivalence()
	print("-= Запуск эквивалентирования =-")
    Set node = t.Tables("node")
    Set vetv = t.Tables("vetv")
    Set gen = t.Tables("Generator")

	'###########################################
	' Урал
	viborka_ot_100_do_200_full = "(na>100 & na<200)" ' для выключателей без учета ограничений по напряжению
	viborka_ot_100_do_200 = "(na>100 & na<200) & (uhom<230)"  ' для эквивалентирования с учетом огран. по напряжению

	'###########################################
	' 201 => Самарская область (АИП => 813)
	viborka_201_Samarskay_obl_full = "na=813"
	viborka_201_Samarskay_obl = "(na=813) & (uhom < 160)"

	'###########################################
	' 205 => Республика Татарстан (Татарстан) (АИП => 205)
	viborka_205_Tatrskay_full = "na=205"
	viborka_205_Tatrskay = "(na=205) & (uhom < 160)"

	'###########################################
	' 206 => Чувашская Республика - Чувашия (АИП => 206)
	viborka_206_Chuvashy_full = "na=206"
	viborka_206_Chuvashy = "(na=206) & (uhom < 160)"

	'###########################################
	' 208 => Республика Марий Эл (АИП => 208)
	viborka_208_MariEl_full = "na=208"
	viborka_208_MariEl = "(na=208) & (uhom < 160)"

	'###########################################
	' 202 => Саратовская область (АИП => 202)
	viborka_202_Saratov_obl_full = "na=202"
	viborka_202_Saratov_obl = "(na=202) & (uhom < 160)"

	'###########################################
	' 301 => Ростовская область (АИП => 301)
	viborka_301_Rostov_obl_full = "na=301"
	viborka_301_Rostov_obl = "(na=301) & (uhom < 160)"

	'###########################################
	' 203 => Ульяновская область (АИП => 203)
	viborka_203_Ulynov_obl_full = "na=203"
	viborka_203_Ulynov_obl = "(na=203) & (uhom < 160)"

	'###########################################
	' 401 => Мурманская область (АИП => 401)
	viborka_401_Murmansk_obl_full = "na=401"
	viborka_401_Murmansk_obl = "(na=401) & (uhom < 160)"

	'###########################################
	' 402 => Республика Карелия (АИП => 402)
	viborka_402_Kareliy_full = "na=402"
	viborka_402_Kareliy = "(na=402) & (uhom < 160)"

	'###########################################
	' 405 => Псковская область (АИП => 405)
	viborka_405_Pskovskay_obl_full = "na=405"
	viborka_405_Pskovskay_obl = "(na=405) & (uhom < 160)"

	'###########################################
	' 407 => Калининградская область (АИП => 407)
	viborka_407_Kaliningrad_obl_full = "na=407"
	viborka_407_Kaliningrad_obl = "(na=407) & (uhom < 160)"

	'###########################################
	' 805 => Эстонская Республика (АИП => 805)
	viborka_805_Estony_full = "na=805"
	viborka_805_Estony = "(na=805) & (uhom < 160)"

	'###########################################
	' 806 => Латвийская Республика (АИП => 806)
	viborka_806_Latviy_full = "na=806"
	viborka_806_Latviy = "(na=806) & (uhom < 160)"

	'###########################################
	' 807 => Литовская Республика (АИП => 807)
	viborka_807_Litva_full = "na=807"
	viborka_807_Litva = "(na=807) & (uhom < 160)"

	'###########################################
	' 801 => Финляндская Республика (АИП => 801)
	viborka_801_Finskay_full = "na=801"
	viborka_801_Finskay = "(na=801) & (uhom < 160)"

	'###########################################
	' 823 => Донбасский регион (АИП => 823)
	viborka_823_Donbas_full = "na=823"
	viborka_823_Donbas = "(na=823) & (uhom < 160)"

	'###########################################
	' 825 => Оренбургская область (АИП => 825 (зима - 831))
	viborka_825_Orenburg_obl_full = "na=831"
	viborka_825_Orenburg_obl = "(na=831) & (uhom < 160)"


	'###########################################
	Call Control_Rgm()
	Call Settings_Ekv()
	Call Obnulenie()
	Call Off_line_one_on()
	Call Control_Rgm()
	Call Off_Gen_if_off_node()
	Call Control_Rgm()
	Call Obnulenie()
    kod = t.rgm("p")
	if kod<>(-1) then
		Call Vikluchatel(viborka_ot_100_do_200_full)
		Call Delete_Gen_Vikl()
		Call Obnulenie()
		Call Ekv_Urala_do_500kV(viborka_ot_100_do_200)
		Call Rastr_Ekv()
		Call Control_Rgm()
		'###################################################
		Call Obnulenie()
		Call Vikluchatel(viborka_201_Samarskay_obl_full)
		Call Obnulenie()
		Call Ekvivalent_Node_Gen(viborka_201_Samarskay_obl)
		Call Rastr_Ekv()
		Call Control_Rgm()
		'###################################################
		Call Obnulenie()
		Call Vikluchatel(viborka_205_Tatrskay_full)
		Call Obnulenie()
		Call Ekvivalent_Node_Gen(viborka_205_Tatrskay)
		Call Rastr_Ekv()
		Call Control_Rgm()
		'###################################################
		Call Obnulenie()
		Call Vikluchatel(viborka_206_Chuvashy_full)
		Call Obnulenie()
		Call Ekvivalent_Node_Gen(viborka_206_Chuvashy)
		Call Rastr_Ekv()
		Call Control_Rgm()
		'###################################################
		Call Obnulenie()
		Call Vikluchatel(viborka_208_MariEl_full)
		Call Obnulenie()
		Call Ekvivalent_Node_Gen(viborka_208_MariEl)
		Call Rastr_Ekv()
		Call Control_Rgm()
		'###################################################
		Call Obnulenie()
		Call Vikluchatel(viborka_202_Saratov_obl_full)
		Call Obnulenie()
		Call Ekvivalent_Node_Gen(viborka_202_Saratov_obl)
		Call Rastr_Ekv()
		Call Control_Rgm()
		'###################################################
		Call Obnulenie()
		Call Vikluchatel(viborka_301_Rostov_obl_full)
		Call Obnulenie()
		Call Ekvivalent_Node_Gen(viborka_301_Rostov_obl)
		Call Rastr_Ekv()
		Call Control_Rgm()
		'###################################################
		Call Obnulenie()
		Call Vikluchatel(viborka_203_Ulynov_obl_full)
		Call Obnulenie()
		Call Ekvivalent_Node_Gen(viborka_203_Ulynov_obl)
		Call Rastr_Ekv()
		Call Control_Rgm()
		'###################################################
		Call Obnulenie()
		Call Vikluchatel(viborka_401_Murmansk_obl_full)
		Call Obnulenie()
		Call Ekvivalent_Node_Gen(viborka_401_Murmansk_obl)
		Call Rastr_Ekv()
		Call Control_Rgm()
		'###################################################
		Call Obnulenie()
		Call Vikluchatel(viborka_402_Kareliy_full)
		Call Obnulenie()
		Call Ekvivalent_Node_Gen(viborka_402_Kareliy)
		Call Rastr_Ekv()
		Call Control_Rgm()
		'###################################################
		Call Obnulenie()
		Call Vikluchatel(viborka_405_Pskovskay_obl_full)
		Call Obnulenie()
		Call Ekvivalent_Node_Gen(viborka_405_Pskovskay_obl)
		Call Rastr_Ekv()
		Call Control_Rgm()
		'###################################################
		Call Obnulenie()
		Call Vikluchatel(viborka_407_Kaliningrad_obl_full)
		Call Obnulenie()
		Call Ekvivalent_Node_Gen(viborka_407_Kaliningrad_obl)
		Call Rastr_Ekv()
		Call Control_Rgm()
		'###################################################
		Call Obnulenie()
		Call Vikluchatel(viborka_805_Estony_full)
		Call Obnulenie()
		Call Ekvivalent_Node_Gen(viborka_805_Estony)
		Call Rastr_Ekv()
		Call Control_Rgm()
		'###################################################
		Call Obnulenie()
		Call Vikluchatel(viborka_806_Latviy_full)
		Call Obnulenie()
		Call Ekvivalent_Node_Gen(viborka_806_Latviy)
		Call Rastr_Ekv()
		Call Control_Rgm()
		'###################################################
		Call Obnulenie()
		Call Vikluchatel(viborka_807_Litva_full)
		Call Obnulenie()
		Call Ekvivalent_Node_Gen(viborka_807_Litva)
		Call Rastr_Ekv()
		Call Control_Rgm()
		'###################################################
		Call Obnulenie()
		Call Vikluchatel(viborka_801_Finskay_full)
		Call Obnulenie()
		Call Ekvivalent_Node_Gen(viborka_801_Finskay)
		Call Rastr_Ekv()
		Call Control_Rgm()
		'###################################################
		Call Obnulenie()
		Call Vikluchatel(viborka_823_Donbas_full)
		Call Obnulenie()
		Call Ekvivalent_Node_Gen(viborka_823_Donbas)
		Call Rastr_Ekv()
		Call Control_Rgm()
		'###################################################
		Call Obnulenie()
		Call Vikluchatel(viborka_825_Orenburg_obl_full)
		Call Obnulenie()
		Call Ekvivalent_Node_Gen(viborka_825_Orenburg_obl)
		Call Rastr_Ekv()
		Call Control_Rgm()

        '###################################################
        Call Obnulenie()
        Call Control_Rgm()
	else
		print("--- Ошибка: Режим расходится! ---")
		print("--- Работа макроса завершена АВАРИЙНО! ---")
	end if
End Sub

Sub Ekv_Urala_do_500kV(viborka_ot_100_do_200)
	Set Vetv = t.Tables("vetv")
    Set Node = t.Tables("node")
    Set Generator = t.Tables("Generator")

	Node.SetSel(viborka_ot_100_do_200) ' выборка по узлам
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
                        if (v_ip > 430 and v_iq < 580) or (v_ip < 430 and v_iq > 580) then
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
                    if (v_ip_2 > 430 and v_iq_2 < 580) or (v_ip_2 < 430 and v_iq_2 > 580) then
						Node.Cols("sel").Z(j) = 0
					end if
                end if
                j_vetv_2 = Vetv.FindNextSel(j_vetv_2)
			wend
        End If
        Node.SetSel(vyborka_gen)
		j = Node.FindNextSel(j)
    Wend
	print("-> Завершено выделение района(-ов): " & viborka_ot_100_do_200 )
End Sub

Sub Ekvivalent_Node_Gen(vyborka_gen)
    Set Vetv = t.Tables("vetv")
    Set Node = t.Tables("node")
    Set Generator = t.Tables("Generator")

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
	print("-> Завершено выделение района(-ов): " & vyborka_gen )
End Sub

Sub print(str)
    t.Printp(str)
End Sub

Sub Rastr_Ekv()
	t.Ekv("")
	print(" - Запуск: ЭКВИВАЛЕНТИРОВАНИЯ!")
End Sub

Sub Control_Rgm()
	kod = t.rgm("p")
	if kod<>(-1) then
		print(" - Режим сошелся!")
	else
		print(" - Режим расходится!")
	end if
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
	print(" - Сняты 'Отметки' с выделенных узлов и ветвей.")
End Sub

Sub Delete_Gen_Vikl()
	set vet=t.tables("vetv")
	set uzl=t.tables("node")
	set gen=t.tables("Generator")
	set ti=t.Tables("ti")

	Call Obnulenie()

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

Sub Settings_Ekv()
	' Выставляет настройки эквивалентирования
    print(" - Выставляет настройки ген. эквив;")
	t.Tables("com_ekviv").Cols("zmax").z(0) = 1000
	t.Tables("com_ekviv").Cols("ek_sh").z(0) = 0
	t.Tables("com_ekviv").Cols("otm_n").z(0) = 0
	t.Tables("com_ekviv").Cols("smart").z(0) = 0
	t.Tables("com_ekviv").Cols("tip_ekv").z(0) = 0
	t.Tables("com_ekviv").Cols("ekvgen").z(0) = 0
	t.Tables("com_ekviv").Cols("tip_gen").z(0) = 1
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
	print(" - Отключены генераторы в отключенных узлах.")
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
	print(" - Количество ЛЭП с односторонним вкл., переведенных с состояние полного откючения: " & ii)
End Sub