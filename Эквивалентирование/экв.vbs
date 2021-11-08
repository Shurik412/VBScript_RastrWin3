' 	 Макрос для формирования Динамической Расчётной Модели (ДРМ) - 2020
'
' 1. Эквивалентирование - БРМ (корректировка БРМ)
' 2. Устранение ошибок после эквивалентирования (удаление узлов без связи, удаление УШР и Реакторов без узлов ... )
' 3. Заполнение актуального Динамического набора из Excel файла
'
'**************************************************************************

r = Setlocale("en-us")
rrr = 1
Set t = Rastr
Set vetv = t.Tables("vetv")
Set node = t.Tables("node")
Set generator = t.Tables("Generator")

print("Запуск макроса " & "дата: " & date() & " | время: " & Hour(Now()) & " hour " & Minute(Now()) & " minut")
Time_1 = Timer()
Call Equivalence() ' - эквивалентирование БРМ.
Time_2 = Timer()
print(" Время работы МАКРОСА, в минутах = " & ((Time_2 - Time_1)/60))
print("Работа макроса завершена.")

'\\************************************************************************
Sub Equivalence()
	print("- =  Запуск эквивалентирования  = -")

	'###########################################
	' Урал
	viborka_ot_100_do_200_full = "(na>100 & na<200)" ' для выключателей без учета ограничений по напряжению
	viborka_ot_100_do_200 = "(na>100 & na<200) & (uhom<230)"  ' для эквивалентирования с учетом огран. по напряжению

	'###########################################
	' 201  = > Самарская область (АИП  = > 813)
	viborka_201_Samarskay_obl_full = "na = 813"
	viborka_201_Samarskay_obl = "(na = 813) & (uhom < 160)"

	'###########################################
	' 205  = > Республика Татарстан (Татарстан) (АИП  = > 205)
	viborka_205_Tatrskay_full = "na = 205"
	viborka_205_Tatrskay = "(na = 205) & (uhom < 160)"

	'###########################################
	' 206  = > Чувашская Республика - Чувашия (АИП  = > 206)
	viborka_206_Chuvashy_full = "na = 206"
	viborka_206_Chuvashy = "(na = 206) & (uhom < 160)"

	'###########################################
	' 208  = > Республика Марий Эл (АИП  = > 208)
	viborka_208_MariEl_full = "na = 208"
	viborka_208_MariEl = "(na = 208) & (uhom < 160)"

	'###########################################
	' 202  = > Саратовская область (АИП  = > 202)
	viborka_202_Saratov_obl_full = "na = 202"
	viborka_202_Saratov_obl = "(na = 202) & (uhom < 160)"

	'###########################################
	' 301  = > Ростовская область (АИП  = > 301)
	viborka_301_Rostov_obl_full = "na = 301"
	viborka_301_Rostov_obl = "(na = 301) & (uhom < 160)"

	'###########################################
	' 203  = > Ульяновская область (АИП  = > 203)
	viborka_203_Ulynov_obl_full = "na = 203"
	viborka_203_Ulynov_obl = "(na = 203) & (uhom < 160)"

	'###########################################
	' 401  = > Мурманская область (АИП  = > 401)
	viborka_401_Murmansk_obl_full = "na = 401"
	viborka_401_Murmansk_obl = "(na = 401) & (uhom < 160)"

	'###########################################
	' 402  = > Республика Карелия (АИП  = > 402)
	viborka_402_Kareliy_full = "na = 402"
	viborka_402_Kareliy = "(na = 402) & (uhom < 160)"

	'###########################################
	' 405  = > Псковская область (АИП  = > 405)
	viborka_405_Pskovskay_obl_full = "na = 405"
	viborka_405_Pskovskay_obl = "(na = 405) & (uhom < 160)"

	'###########################################
	' 407  = > Калининградская область (АИП  = > 407)
	viborka_407_Kaliningrad_obl_full = "na = 407"
	viborka_407_Kaliningrad_obl = "(na = 407) & (uhom < 160)"

	'###########################################
	' 805  = > Эстонская Республика (АИП  = > 805)
	viborka_805_Estony_full = "na = 805"
	viborka_805_Estony = "(na = 805) & (uhom < 160)"

	'###########################################
	' 806  = > Латвийская Республика (АИП  = > 806)
	viborka_806_Latviy_full = "na = 806"
	viborka_806_Latviy = "(na = 806) & (uhom < 160)"

	'###########################################
	' 807  = > Литовская Республика (АИП  = > 807)
	viborka_807_Litva_full = "na = 807"
	viborka_807_Litva = "(na = 807) & (uhom < 160)"

	'###########################################
	' 801  = > Финляндская Республика (АИП  = > 801)
	viborka_801_Finskay_full = "na = 801"
	viborka_801_Finskay = "(na = 801) & (uhom < 160)"

	'###########################################
	' 823  = > Донбасский регион (АИП  = > 823)
	viborka_823_Donbas_full = "na = 823"
	viborka_823_Donbas = "(na = 823) & (uhom < 160)"

	'###########################################
	' 825  = > Оренбургская область (АИП  = > 825 (зима - 831))
	viborka_825_Orenburg_obl_full = "na = 831"
	viborka_825_Orenburg_obl = "(na = 831) & (uhom < 160)"


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
	if kod<>(-1) Then
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
	End If
End Sub

Sub Ekv_Urala_do_500kV(viborka_ot_100_do_200)
	node.SetSel(viborka_ot_100_do_200) ' выборка по узлам
    node.Cols("sel").Calc("1")
    j = node.FindNextSel(-1)
	
	While j<>(-1)
        ny = node.Cols("ny").Z(j)
        tip_node = node.Cols("tip").Z(j)
        uhom = node.Cols("uhom").Z(j)
        If tip_node > 1 Then ' все генераторные узла
            generator.SetSel("Node.ny = " & ny)
            j_gen = generator.FindNextSel(-1)
            If j_gen <> (-1) Then
                vetv.SetSel("(ip =  " & ny & ")|(iq =  " & ny & ")")
                j_vetv = vetv.FindNextSel(-1)
                while j_vetv <>(-1)
                    tip_vetv = vetv.Cols("tip").Z(j_vetv)
                    If tip_vetv = 1 Then
                        v_ip = vetv.Cols("v_ip").Z(j_vetv)
                        v_iq = vetv.Cols("v_iq").Z(j_vetv)
                        If (v_ip > 430 and v_iq < 580) or (v_ip < 430 and v_iq > 580) Then
                            node.Cols("sel").Z(j) = 0
                        End If
                    End If
                    j_vetv = vetv.FindNextSel(j_vetv)
                wend
            End If
        Else
            vetv.SetSel("(ip =  " & ny & ")|(iq =  " & ny & ")")
            j_vetv_2 = vetv.FindNextSel(-1)
            while j_vetv_2 <>(-1)
				tip_vetv_2 = vetv.Cols("tip").Z(j_vetv_2)
                If tip_vetv_2 = 1 Then
					v_ip_2 = vetv.Cols("v_ip").Z(j_vetv_2)
                    v_iq_2 = vetv.Cols("v_iq").Z(j_vetv_2)
                    If (v_ip_2 > 430 and v_iq_2 < 580) or (v_ip_2 < 430 and v_iq_2 > 580) Then
						node.Cols("sel").Z(j) = 0
					End If
                End If
                j_vetv_2 = vetv.FindNextSel(j_vetv_2)
			wend
        End If
        node.SetSel(vyborka_gen)
		j = node.FindNextSel(j)
    Wend
	print("-> Завершено выделение района(-ов): " & viborka_ot_100_do_200 )
End Sub

Sub Ekvivalent_Node_Gen(vyborka_gen)
	node.SetSel(vyborka_gen) ' выборка по узлам
    node.Cols("sel").Calc("1")
    j = node.FindNextSel(-1)

    While j<>(-1)
        ny = node.Cols("ny").Z(j)
        tip_node = node.Cols("tip").Z(j)
        uhom = node.Cols("uhom").Z(j)
        If tip_node > 1 Then ' все генераторные узла
            generator.SetSel("Node.ny = " & ny)
            j_gen = generator.FindNextSel(-1)
            If j_gen <> (-1) Then
                vetv.SetSel("(ip =  " & ny & ")|(iq =  " & ny & ")")
                j_vetv = vetv.FindNextSel(-1)
                while j_vetv <>(-1)
                    tip_vetv = vetv.Cols("tip").Z(j_vetv)
                    If tip_vetv = 1 Then
                        v_ip = vetv.Cols("v_ip").Z(j_vetv)
                        v_iq = vetv.Cols("v_iq").Z(j_vetv)
                        If (v_ip > 170 and v_iq < 250) or (v_ip < 170 and v_iq > 250) Then
                            node.Cols("sel").Z(j) = 0
                        End If
                    End If
                    j_vetv = vetv.FindNextSel(j_vetv)
                wend
            End If
        Else
            vetv.SetSel("(ip =  " & ny & ")|(iq =  " & ny & ")")
            j_vetv_2 = vetv.FindNextSel(-1)
            while j_vetv_2 <>(-1)
				tip_vetv_2 = vetv.Cols("tip").Z(j_vetv_2)
                If tip_vetv_2 = 1 Then
					v_ip_2 = vetv.Cols("v_ip").Z(j_vetv_2)
					v_iq_2 = vetv.Cols("v_iq").Z(j_vetv_2)
                    If (v_ip_2 > 170 and v_iq_2 < 250) or (v_ip_2 < 170 and v_iq_2 > 250) Then
						node.Cols("sel").Z(j) = 0
                    End If
                End If
                j_vetv_2 = vetv.FindNextSel(j_vetv_2)
            wend
        End If
        node.SetSel(vyborka_gen)
		j = node.FindNextSel(j)
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
	If kod<>(-1) Then
		print(" - Режим сошелся!")
	else
		print(" - Режим расходится!")
	End If
End Sub

Sub Vikluchatel(viborka_ray_vikl)
    Dim nodes(30000)

	node.SetSel(viborka_ray_vikl) ' выборка узлов всех районов кроме 500 (Центра)
    node.Cols("sel").Calc(1) ' выделение выбраных узлов
    vetv.SetSel("iq.sel = 1 & ip.sel = 0 &!sta") ' выборка ветвей iq.sel = 1 ...
    k = vetv.FindNextSel(-1)
	While k<>(-1) ' убирает sel-узла если на ВЛ с одной стороны выделен узел
		iq1 = vetv.Cols("iq").Z(k)
		node.Setsel("ny = " & iq1)
		k2 = node.FindNextSel(-1)
		If k2<>(-1) Then
			node.Cols("sel").Z(k2) = 0
		End If
		k = vetv.FindNextSel(k)
    Wend
	
    vetv.SetSel("iq.sel = 0 & ip.sel = 1 & !sta")
    k = vetv.FindNextSel(-1)

    While k<>(-1) ' убирает sel-узла если на ВЛ с одной стороны выделен узел
		ip1 = vetv.Cols("ip").Z(k)
		node.Setsel "ny = " & ip1
		k2 = node.FindNextSel(-1)
		If k2<>(-1) Then
			node.Cols("sel").Z(k2) = 0
		End If
		k = vetv.FindNextSel(k)
	Wend

	vetv.SetSel("(iq.sel = 1 & ip.sel = 0)|(ip.sel = 1 & iq.sel = 0) & tip = 2") ' tip = 2 - выключатели (выборка всех выключателей если хотябы с одной стороны выделен узел sel)
    k = vetv.FindNextSel(-1)
    While k<>(-1)
		iq1 = vetv.Cols("iq").Z(k)
		node.Setsel "ny = " & iq1
		k2 = node.FindNextSel(-1)
		If k2<>(-1) Then
			node.Cols("sel").Z(k2) = 0
		End If
		ip1 = vetv.Cols("ip").Z(k)
		node.Setsel "ny = " & ip1
		k2 = node.FindNextSel(-1)
		If k2<>(-1) Then
			node.Cols("sel").Z(k2) = 0
		End If
		vetv.SetSel("(iq.sel = 1 &ip.sel = 0) | (ip.sel = 1 &iq.sel = 0) & tip = 2")
		k = vetv.FindNextSel(-1)
    Wend

    vetvyklvybexc = "(iq.bsh>0 & ip.bsh = 0) | (ip.bsh>0 & iq.bsh = 0) | (iq.bshr>0 & ip.bshr = 0) | (ip.bshr>0 & iq.bshr = 0)| ip.sel = 0 | iq.sel = 0)"
    flvykl = 0
	vetv.SetSel("1")
	vetv.Cols("groupid").Calc(0)
	vetv.SetSel(vetvyklvybexc)
	vetv.Cols("groupid").calc(1)
	nvet = 0
	' удаление выключателей
	For povet = 0 to 10000
		vetv.SetSel("x<0.01 & x>-0.01 & r<0.005 & r> = 0 & (ktr = 0 | ktr = 1) & !sta & groupid! = 1 & b<0.000005")  'Выборка ветвей, которые считаем выключателями
		ivet = vetv.FindNextSel(-1)
		If ivet = -1 Then exit For
            ip = vetv.Cols("ip").Z(ivet)
            iq = vetv.Cols("iq").Z(ivet)
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
            For inodee = 0 to nnod
                If 	ndel = nodes(inodee) Then ndndel = 1
                If 	ny = nodes(inodee) Then ndny = 1
                If (ndndel = 1) and (ndny = 1) Then exit For
            Next
			' Меняем местами, так как удаляемый нельзя удалять, а неудаляемый можно ))
            If (ndndel = 0) and (ndny = 1) Then
                buff = ny
                ny = ndel
                ndel = buff
            End If

            If (ndndel = 0) or (ndny = 0) Then 'Если хотя бы один можно удалить
                flvykl = flvykl + 1
				node.SetSel("ny = " & ny)
				iny = node.FindNextSel(-1)
				node.SetSel("ny = " & ndel)
				idel = node.FindNextSel(-1)
				pgdel = node.Cols("pg").Z(idel)
				qgdel = node.Cols("qg").Z(idel)
				pndel = node.Cols("pn").Z(idel)
				qndel = node.Cols("qn").Z(idel)
				bshdel = node.Cols("bsh").Z(idel)
				gshdel = node.Cols("gsh").Z(idel)
				pgny = node.Cols("pg").Z(iny)
				qgny = node.Cols("qg").Z(iny)
				pnny = node.Cols("pn").Z(iny)
				qnny = node.Cols("qn").Z(iny)
				bshny = node.Cols("bsh").Z(iny)
				gshny = node.Cols("gsh").Z(iny)

				node.Cols("pg").Z(iny) = pgdel + pgny
				node.Cols("qg").Z(iny) = qgdel + qgny
				node.Cols("pn").Z(iny) = pndel + pnny
				node.Cols("qn").Z(iny) = qndel + qnny
				node.Cols("bsh").Z(iny) = bshdel + bshny
				node.Cols("gsh").Z(iny) = gshdel + gshny
				v1 = node.Cols("vzd").Z(iny)
				v2 = node.Cols("vzd").Z(idel)
				qmax1 = node.Cols("qmax").Z(iny)
				qmax2 = node.Cols("qmax").Z(idel)

				generator.Setsel("Node = " & ndel)
				igen = generator.FindNextSel(-1) 'Меняем узлы подключения генераторов

				If igen<>(-1) Then
					While igen<>(-1)
						generator.Cols("Node").Z(igen) = ny
						igen = generator.FindNextSel(igen)
					Wend
				End If

				If (v1<>v2) and (v1>0.3) and (v2>0.3) and (qmax1 + qmax2) <> 0 Then
					node.Cols("vzd").Z(iny) = (v1*qmax1+v2*qmax2)/(qmax1+qmax2) 'Делаем средневзвешенное по qmax напряжение
				End If

				If (v1 = 0) and (v2<>0) Then
					node.Cols("vzd").Z(iny) = v2
				End If

				If (v1<>0) and (v2<>0) Then
					node.Cols("qmin").Z(iny) = (node.Cols("qmin").Z(iny)) + (node.Cols("qmin").Z(idel))
					node.Cols("qmax").Z(iny) = qmax1 + qmax2
				End If

				If (v1 = 0) and (v2<>0) Then
					node.Cols("qmin").Z(iny) = node.Cols("qmin").Z(idel)
					node.Cols("qmax").Z(iny) = node.Cols("qmax").Z(idel)
				End If

				vetv.SetSel("(ip = " & ip & "& iq = " & iq & ")|(iq = " & ip & "& ip = " & iq & ")")
				vetv.Delrows 'Удаляем ветвь
				vetv.SetSel("iq = " & ndel) 'Меняем узлы ветвей с удаляемым узлом)))
				vetv.Cols("iq").calc(ny)
				vetv.SetSel("ip = " & ndel)
				vetv.Cols("ip").calc(ny)
				node.Delrows 		' Удаляем узел
			Else					'Если ни одного нельзя удалить
                vetv.SetSel("(ip = " & ip & "& iq = " & iq & ")|(iq = " & ip & "& ip = " & iq & ")")
                vetv.Cols("groupid").calc(1)
        End If
    Next
    kod = t.rgm ("p")
    If kod<>0 Then
        msgbox "Regim do not exist"
    End If
End Sub

Sub zeroing()
	'************************************************************
	' Назначение: обнуление всех sel (выделенных галочкой) УЗЛОВ и ВЕТВЕЙ.
	'             
	' Входные
	' параметры:  
	' Возврат:    Nothing
	'************************************************************
	
    vetv.SetSel("")
	vetv.Cols("sel").Calc("0")
	node.SetSel("")
	node.Cols("sel").Calc("0")
	print(" - Сняты 'Отметки' с выделенных узлов и ветвей.")
End Sub

Sub deleting_generator_switches()
	' Назначение: Удаляет выключатели генераторов.
	'             
	' Входные
	' параметры:  
	' Возврат:    Nothing
	'**************************************************************
	Set ti = t.Tables("ti")
	Call zeroing()
	node.SetSel("")
	k1 = node.FindNextSel(-1)
	While k1<>(-1)
		ny1 = node.Cols("ny").Z(k1)
		vetv.SetSel("(ip = " & ny1 &") |(iq = " & ny1 &")" )
		If vetv.Count = 1 Then
			vetv.SetSel("x<1 & (tip = 0 | tip = 2) & ((ip = " & ny1 & ") |(iq = " & ny1 &"))")
			If vetv.Count = 1 Then
				vetv.SetSel("x<1 & (tip = 0 | tip = 2) & ((ip = " & ny1 & ") |(iq = " & ny1 &"))" )
				k3 = vetv.findNextsel(-1)
				If k3<>(-1) Then
					If vetv.Cols("ip").Z(k3) = ny1 Then
						ny2 = vetv.Cols("iq").Z(k3)
					else
						ny2 = vetv.Cols("ip").Z(k3)
					End If
					generator.SetSel("Node = " & ny1)
					k2 = generator.FindNextSel(-1)
					If k2<>(-1) Then
						node.SetSel("ny = " & ny2)
						k4 = node.FindNextSel(-1)
						If k4<>(-1) Then
							node.Cols("pn").Z(k4) = node.Cols("pn").Z(k1) + node.Cols("pn").Z(k1)
							node.Cols("qn").Z(k4) = node.Cols("qn").Z(k1) + node.Cols("qn").Z(k1)
							node.Cols("vzd").Z(k4) = node.Cols("vzd").Z(k1)
							node.Cols("exist_load").Z(k4) = node.Cols("exist_load").Z(k1)
							node.Cols("exist_gen").Z(k4) = node.Cols("exist_gen").Z(k1)
							node.Cols("pn_max").Z(k4)  = node.Cols("pn_max").Z(k1) + node.Cols("pn_max").Z(k4)
							If node.Cols("pn_min").Z(k4)  = > node.Cols("pn_min").Z(k1) Then
								node.Cols("pn_min").Z(k4) = node.Cols("pn_min").Z(k1)
							End If
							node.Cols("pg_max").Z(k4) = node.Cols("pg_max").Z(k1) + node.Cols("pg_max").Z(k4)
							If node.Cols("pg_min").Z(k4)  = > node.Cols("pg_min").Z(k1) Then
								node.Cols("pg_min").Z(k4) = node.Cols("pg_min").Z(k1)
							End If
							node.Cols("sel").Z(k1) = 1
							vetv.Cols("sel").Z(k3) = 1
							' ti.SetSel("(prv_num = 20 | prv_num = 7 | prv_num = 6 | prv_num = 5 | prv_num = 4 | prv_num = 3 | prv_num = 2 | prv_num = 1) & id1 = "&ny1)
							' ti.Cols("id1").calc(ny2)
							generator.SetSel("Node = " & ny1)
							k2 = generator.FindNextSel(-1)
							while k2 <> (-1)
								generator.Cols("Node").Z(k2) = ny2
								k2 = generator.FindNextSel(k2)
							wend
						End If
					End If
				End If
			End If
		End If
		node.SetSel("")
		k1 = node.FindNextSel(k1)
	Wend
	vetv.SetSel("sel = 1")
	vetv.Delrows
	node.SetSel("sel = 1")
	node.Delrows
End Sub

Sub equivalent_settings()
	' Назначение: Выставляет настройки эквивалентирования.
	'             
	' Входные
	' параметры:  
	' Возврат:    Nothing
	'**************************************************************
    print(" - Выставляет настройки ген. эквив;")
	t.Tables("com_ekviv").Cols("zmax").Z(0) = 1000
	t.Tables("com_ekviv").Cols("ek_sh").Z(0) = 0
	t.Tables("com_ekviv").Cols("otm_n").Z(0) = 0
	t.Tables("com_ekviv").Cols("smart").Z(0) = 0
	t.Tables("com_ekviv").Cols("tip_ekv").Z(0) = 0
	t.Tables("com_ekviv").Cols("ekvgen").Z(0) = 0
	t.Tables("com_ekviv").Cols("tip_gen").Z(0) = 1
End Sub

Sub off_the_generator_if_the_node_off()
	'**************************************************************
	' Назначение: Отключение генераора, если узел к которому подключен  
	'             генератор отключен.
	' Входные
	' параметры:  Nothing
	' Возврат:    Nothing
	'**************************************************************
	generator.SetSel("")
	k = generator.FindNextSel(-1)
	counter = 0
	while k<>(-1)
		Node_generator = generator.Cols("Node").Z(k)
		node.SetSel "ny = " & Node_generator
		kk = node.FindNextSel(-1)
		if kk<>(-1) Then
			if node.Cols("sta").Z(kk) = True Then
				generator.Cols("sta").Z(k) = 1
				counter = counter + 1
			End If
		End If
		generator.SetSel("")
		k = generator.FindNextSel(k)
	wend
	print(" - Отключено генераторов: " & counter)
End Sub

Sub off_the_line_from_two_side()
	'**************************************************************
	' Назначение: Отключение ЛЭП с двух сторон, если она включена с одной стороны.
	'             
	' Входные
	' параметры:  Nothing
	' Возврат:    Nothing
	'**************************************************************
	counter = 0
	For i = 0 to vetv.Count-1
		sta = vetv.Cols("sta").Z(i)
		If sta = 2 or sta = 3 Then
			vetv.Cols("sta").Z(i) = 1
			counter = counter + 1
		End If
	Next
	print(" - Количество ЛЭП с односторонним вкл., переведенных с состояние полного откючения: " & counter)
End Sub