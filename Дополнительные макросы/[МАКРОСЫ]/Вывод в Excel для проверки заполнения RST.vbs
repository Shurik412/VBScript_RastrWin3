r=setlocale("en-us")
rrr=1
Excel_off_on = 0
if Excel_off_on = 1 Then
    Set ExcelSet = CreateObject("Excel.Application")	
        ExcelSet.Workbooks.open SettingsFile
        ExcelSet.Visible = 1
end if 	
set t=RASTR
	set gen=t.Tables("Generator")
 
	set uzl=t.tables("node")
    set ny=uzl.Cols("ny")
    set name=uzl.Cols("name")
    
	set vozb=t.Tables("Exciter")
	set ARV_ID=t.Tables("ExcControl")
	set ARS=t.tables("Governor")
	set Turb=t.tables("ARS")
	set forc=t.Tables("Forcer")

CountGen = gen.Count-1
t.printp("CountGen = " & CountGen)

CountVozb = vozb.Count - 1
t.printp("CountVozb = " & CountVozb)

CountARV = ARV_ID.Count - 1
t.printp("CountARV = " & CountARV)

CountTurb = Turb.Count - 1
t.printp("CountTurb = " & CountTurb)

CountARS = ARS.Count - 1
t.printp("CountARS = " & CountARS)

CountFors = forc.Count - 1
t.printp("CountFors = " & CountFors)

Proverka_in_PrintP = 0
Proverka_Vozb = 0
For2_Exc = 1
For3_ARS = 1
For4_Turb = 1
For5_ARS = 1
For6_otkl_Gen_PQ_0 = 1
For7_Fors = 1


if Proverka_in_PrintP = 1 Then
	for i=0 to CountGen 
		 gen_num=gen.Cols("Num").Z(i)
		 gen_name=gen.Cols("Name").Z(i)
		 gen_node=gen.Cols("Node").Z(i)
		 gen_ExciterId=gen.Cols("ExciterId").Z(i)
		 gen_Turbina=gen.Cols("ARSId").Z(i)

		 If gen_ExciterId <> " " Then
			 vozb.SetSel "Id =" & gen_ExciterId
			 j_vozd_Id = vozb.FindNextSel(-1)		 
			 IF j_vozd_Id <> -1 Then
				vozb_name=vozb.Cols("Name").Z(j_vozd_Id)
				vozb_id	=vozb.Cols("Id").Z(j_vozd_Id)
				vozd_ARV_Id = vozb.Cols("ExcControlId").Z(j_vozd_Id)
				If gen_name <> vozb_name Then
					t.printp("========================================================================================")
					t.printp("Название Возбудителя_Ген и Возбудителя_по_Номеру - не совпадают! \\ " & gen_name & " \\ " & vozb_name)
					if gen_ExciterId <> vozb_id Then
						t.printp("Номер Возбудителя_Ген и Возбудителя_по_Номеру - не совпадают! \\"& gen_ExciterId & " \\ " & vozb_id)
						t.printp("========================================================================================")
					else
						t.printp("========================================================================================")
					end if
				end If
			 else
				 t.printp("========================================================================================")
				 t.printp("В генераторе с номером (названием) \\ " & gen_num & " \\ (" & gen_name &") - возбудитель не найдет,  либо задан несуществующий номер возбудителя!")
				 t.printp("========================================================================================")
			 End IF
		 else
			t.Printp("В генераторе не задан номер - возбудителя")
		 End If


		 If vozd_ARV_Id <> " " Then
			 ARV_ID.SetSel "Id =" & vozd_ARV_Id
			 j_ARV_ID = ARV_ID.FindNextSel(-1)		 
			 IF j_ARV_ID <> (-1) Then
				ARV_ID_id=ARV_ID.Cols("Id").Z(j_ARV_ID)
				ARV_ID_name=ARV_ID.Cols("Name").Z(j_ARV_ID)
				If vozb_name <> ARV_ID_name Then
					t.printp("========================================================================================")
					t.printp("Название Возбудителя_по_Номеру_Ген. и АРВ(ИД)_по_Номеру_Возб. - не совпадают! \\ " & vozb_name & " \\ " & ARV_ID_name)
					if vozb_name <> ARV_ID_name Then
						t.printp("Номер Возбудителя_по_Номеру_Ген и АРВ(ИД)_по_Номеру_Возб. - не совпадают! \\"& vozb_name & " \\ " & ARV_ID_name)
						t.printp("========================================================================================")
					else
						t.printp("========================================================================================")
					end if
				end If
			 else
				 t.printp("========================================================================================")
				 t.printp("В возбудителе с номером (названием) \\ " & vozb_id & " \\ (" & vozb_name &") - АРВ не найдет,  либо задан несуществующий номер возбудителя!")
				 t.printp("========================================================================================")
			 End IF
		  else
			t.Printp("В возбудителе не задан номер - АРВ(ИД)")	 
		  End if

		  If gen_Turbina <> "" Then
			 Turb.SetSel "Id =" & gen_Turbina
			 j_Turb = Turb.FindNextSel(-1)		 
			 IF j_Turb <> (-1) Then
				Turb_id=Turb.Cols("Id").Z(j_Turb)
				Turb_name=Turb.Cols("Name").Z(j_Turb)
				ARS_Turb_Id = Turb.Cols("GovernorId").Z(j_Turb)
				If gen_name <> Turb_name Then
					t.printp("========================================================================================")
					t.printp("Название Генератора и Турбины_по_Номеру_Ген. - не совпадают! \\ " & gen_name & " \\ " & Turb_name)
					if gen_Turbina <> Turb_id Then
						t.printp("Номер Турбины_по_Номеру_Ген и Турбины_по_Номеру_Ген. - не совпадают! \\"& gen_Turbina & " \\ " & Turb_id)
						t.printp("========================================================================================")
					else
						t.printp("========================================================================================")
					end if
				end If
			 else
				 t.printp("========================================================================================")
				 t.printp("В возбудителе с номером (названием) \\ " & gen_Turbina & " \\ (" & gen_Turbina &") - АРВ не найдет,  либо задан несуществующий номер возбудителя!")
				 t.printp("========================================================================================")
			 End IF
		 else
			t.Printp("В генераторе не задан номер - Турбины")
		 End If
		 
		  If ARS_Turb_Id <> "" Then
			 ARS.SetSel "Id =" & ARS_Turb_Id
			 j_ARS = ARS.FindNextSel(-1)		 
			 IF j_ARS <> (-1) Then
				ARS_Id = ARS.Cols("Id").Z(j_ARS)
				ARS_Name = ARS.Cols("Name").Z(j_ARS)
				If Turb_name <> ARS_Name Then
					t.printp("========================================================================================")
					t.printp("Название Турбины и АРС_по_Номеру_Турбины. - не совпадают! \\ " & Turb_name & " \\ " & ARS_Name)
					if ARS_Turb_Id <> ARS_Id Then
						t.printp("Номер Турбины_по_Номеру_Ген и АРС_по_Номеру_Турины. - не совпадают! \\"& ARS_Turb_Id & " \\ " & ARS_Id)
						t.printp("========================================================================================")
					else
						t.printp("========================================================================================")
					end if
				end If
			 else
				 t.printp("========================================================================================")
				 t.printp("В АРС(ИД) с номером (названием) \\ " & ARS_Id & " \\ (" & ARS_Name &") - АРВ не найдет,  либо задан несуществующий номер возбудителя!")
				 t.printp("========================================================================================")
			 End IF
		 else
			t.Printp("В генераторе не задан номер - АРС(ИД)")
		 End If
       next  
end if

If Proverka_Vozb = 1 Then
	for i=0 to CountGen 
		 gen_num=gen.Cols("Num").Z(i)
		 gen_name=gen.Cols("Name").Z(i)
		 gen_node=gen.Cols("Node").Z(i)
		 gen_ExciterId=gen.Cols("ExciterId").Z(i)
		 gen_Turbina=gen.Cols("ARSId").Z(i)	
		 t.PrintP("i= " & i)
		 If gen_name <> "" Then
			 vozb.SetSel "Name =" & gen_name
             j_vozd_Name = vozb.FindNextSel(-1)		 
             IF j_vozd_Name <> -1 Then
				vozb_name=vozb.Cols("Name").Z(j_vozd_Name)
				vozb_id	= vozb.Cols("Id").Z(j_vozd_Name)
				vozd_ARV_Id = vozb.Cols("ExcControlId").Z(j_vozd_Name) ' Номер АРВ(ИД) из таблицы возбудители (ИД)
                't.PrintP(gen_name & " ! " & vozb_name)
                if gen_ExciterId <> vozb_id or gen_name <> vozb_name Then
                    gen.Cols("ExciterId").Z(i) = vozb_id
                    t.Printp("Номер Возб в табл Ген не равен Номеру возбу по названию Ген: i = " & i & ". Номер возбудителя в табл. генераторов: " & gen_ExciterId & ", Название генератора: " & gen_name & ", Название возбудителя: " & vozb_name & ", Номер возбудителя по названию ген.: " & vozb_id)
                else
                    t.PrintP("Номер Возб в табл Ген равен Номеру возбу по названию Ген")
                end if
             End If   
             
            If j_vozd_Name = -1 Then
                t.Printp("j_vozd_Name = -1")
                t.Printp("----------------------------------------------------------")
            End If
        End if
     next
End if


If For2_Exc = 1 Then
    jj = 0
	for i=0 to CountGen
		gen_num=gen.Cols("Num").Z(i)
		gen_name=gen.Cols("Name").Z(i)
		gen_node=gen.Cols("Node").Z(i)
		gen_ExciterId=gen.Cols("ExciterId").Z(i)
		gen_Turbina=gen.Cols("ARSId").Z(i)	
		
		for j = 0 to CountVozb
            't.PrintP("j= " & j)
			Vozb_name = vozb.Cols("Name").Z(j)
			Vozb_id = vozb.Cols("Id").Z(j)
			If Vozb_name = gen_name Then
				gen.Cols("ExciterId").Z(i) = Vozb_id
                jj = jj + 1
				t.Printp("Номер " & jj & ", Генератор: " & gen_name & "(" & gen_ExciterId & ")" & " => Возбудитель: " & Vozb_id & ", " & "(" & Vozb_id & ")")
                Exit For
            End if
        next
     next 
End If

If For3_ARS = 1 Then
    jj = 0
	for i=0 to CountVozb
		vozb_id=vozb.Cols("Id").Z(i)
		vozb_name=vozb.Cols("Name").Z(i)
        
		for j = 0 to CountARV
            't.PrintP("j= " & j)
			ARV_name = ARV_ID.Cols("Name").Z(j)
			ARV_ids = ARV_ID.Cols("Id").Z(j)
			If Vozb_name = ARV_name Then
				vozb.Cols("ExcControlId").Z(i) = ARV_ids
                jj = jj + 1
				t.Printp("Номер " & jj & ", Возбудитель: " & vozb_name & "(" & vozb_id & ")" & " => АРВ(ИД): " & ARV_name & ", " & "(" & ARV_ids & ")")
                Exit For
            End if
        next
     next 
End If

If For4_Turb = 1 Then
    jj = 0
	for i=0 to CountGen
		gen_num=gen.Cols("Num").Z(i)
		gen_name=gen.Cols("Name").Z(i)
		gen_node=gen.Cols("Node").Z(i)
		gen_ExciterId=gen.Cols("ExciterId").Z(i)
		gen_Turbina=gen.Cols("ARSId").Z(i)	
        
		for j = 0 to CountTurb
            't.PrintP("j= " & j)
			Turb_name = Turb.Cols("Name").Z(j)
			Turb_id = Turb.Cols("Id").Z(j)
            
			If gen_name = Turb_name Then
				gen.Cols("ARSId").Z(i) = Turb_id
                jj = jj + 1
				t.Printp("Номер " & jj & ", Турбина: " & Turb_name & "(" & Turb_id & ")" & " => Генератор: " & gen_name & ", " & "(" & gen_Turbina & ")")
                Exit For
            End if
        next
      next 
End If

If For5_ARS = 1 Then
    jj = 0
	for i=0 to CountTurb
        Turb_name = Turb.Cols("Name").Z(i)
		Turb_id = Turb.Cols("Id").Z(i)	
        
		for j = 0 to CountARS
            't.PrintP("j= " & j)
			ARS_name = ARS.Cols("Name").Z(j)
			ARS_id = ARS.Cols("Id").Z(j)
            
			If Turb_name = ARS_name Then
				Turb.Cols("Id").Z(i) = ARS_id
                jj = jj + 1
                t.Printp("Номер " & jj & ", Турбина: " & Turb_name & "(" & Turb_id & ")" & " => АРС(ИД): " & ARS_name & ", " & "(" & ARS_id & ")")
                Exit For
            End if
        next
      next 
End If

if For6_otkl_Gen_PQ_0 = 1 Then
    r=setlocale("en-us")
    rrr=1
    Set t = Rastr
    Set spGen = t.Tables("Generator")
    Set spNode = t.Tables("node")

    spGenMax = spGen.Count-1
    t.printp(spGenMax)

    for i = 0 to spGenMax
        
        Pgen = spGen.Cols("P").Z(i)
        Qgen = spGen.Cols("Q").Z(i)
        If Pgen = 0 and Qgen = 0 Then
            spGen.Cols("sta").Z(i) = 1
        End if
    next

    t.Printp("Исследование завершено")

End if

If For7_Fors = 1 Then
    jj = 0
	for i = 0 to CountVozb
        Vozb_name = vozb.Cols("Name").Z(i)
		Vozb_id = vozb.Cols("ForcerId").Z(i)
		
        for j = 0 to CountFors
            't.PrintP("j= " & j)
            fors_id = forc.Cols("Id").Z(j)
            fors_name = forc.Cols("Name").Z(j)
            
			If  Vozb_name = fors_name Then
				vozb.Cols("ForcerId").Z(i) = fors_id
                jj = jj + 1
                t.Printp("Номер " & jj & ", Возбудитель: " & Vozb_name & "(" & Vozb_id & ")" & " => Форсировка(ИД): " & fors_name & ", " & "(" & fors_id & ")")
                Exit For
            End if
        next
      next 
End If