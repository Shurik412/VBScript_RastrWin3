r = Setlocale("en-us")
rrr = 1

Excel_off_on = 0
If Excel_off_on = 1 Then
    Set ExcElset = CreateObject("Excel.Application")	
        ExcElset.Workbooks.open SettingsFile
        ExcElset.Visible = 1
End If 	

Set t = RASTR
	Set gen=t.Tables("Generator")
 	Set uzl = t.tables("node")
	Set vozb = t.Tables("Exciter")
	Set ARV_ID = t.Tables("ExcControl")
	Set ARS = t.tables("Governor")
	Set Turb = t.tables("ARS")
	Set Forc = t.Tables("Forcer")
	Set Forc = t.Tables("Forcer")


CountGen = gen.Count-1
t.Printp("���-�� ����������� = " & CountGen)

CountVozb = vozb.Count - 1
t.Printp("���-�� ������������ = " & CountVozb)

CountARV = ARV_ID.Count - 1
t.Printp("���-�� ��� = " & CountARV)

CountTurb = Turb.Count - 1
t.Printp("���-�� ������ = " & CountTurb)

CountARS = ARS.Count - 1
t.Printp("���-�� ��� = " & CountARS)

CountFors = Forc.Count - 1
t.Printp("���-�� ���������� = " & CountFors)

Adjustment_Exciter = 1
Adjustment_ARV = 1
Adjustment_Turbine = 1
Adjustment_ARS = 1
Adjustment_Forsing = 1
OFF_Generators_with_P_and_Q_zero = 0


If Adjustment_Exciter = 1 Then
    jj = 0
	For i=0 to CountGen
		gen_num = gen.Cols("Num").Z(i)
		gen_name = gen.Cols("Name").Z(i)
		gen_node = gen.Cols("Node").Z(i)
		gen_ExciterId = gen.Cols("ExciterId").Z(i)
		gen_Turbina = gen.Cols("ARSId").Z(i)	
		
		For j = 0 to CountVozb
			Vozb_name = vozb.Cols("Name").Z(j)
			Vozb_id = vozb.Cols("Id").Z(j)
			
			If Vozb_name = gen_name Then
				gen.Cols("ExciterId").Z(i) = Vozb_id
                jj = jj + 1
				t.Printp("����� " & jj & ", ���������: " & gen_name & "(" & gen_ExciterId & ")" & " => �����������: " & Vozb_id & ", " & "(" & Vozb_id & ")")
                Exit For
            End If
        Next
    Next 
End If

If Adjustment_ARV = 1 Then
    jj = 0
	For i=0 to CountVozb
		vozb_id = vozb.Cols("Id").Z(i)
		vozb_name = vozb.Cols("Name").Z(i)
        
		For j = 0 to CountARV
            ARV_name = ARV_ID.Cols("Name").Z(j)
			ARV_ids = ARV_ID.Cols("Id").Z(j)
			
			If Vozb_name = ARV_name Then
				vozb.Cols("ExcControlId").Z(i) = ARV_ids
                jj = jj + 1
				t.Printp("����� " & jj & ", �����������: " & vozb_name & "(" & vozb_id & ")" & " => ���(��): " & ARV_name & ", " & "(" & ARV_ids & ")")
                Exit For
            End If
        Next
    Next
End If

If Adjustment_Turbine = 1 Then
    jj = 0
	For i = 0 to CountGen
		gen_num = gen.Cols("Num").Z(i)
		gen_name = gen.Cols("Name").Z(i)
		gen_node = gen.Cols("Node").Z(i)
		gen_ExciterId = gen.Cols("ExciterId").Z(i)
		gen_Turbina = gen.Cols("ARSId").Z(i)	
        
		For j = 0 to CountTurb
          	Turb_name = Turb.Cols("Name").Z(j)
			Turb_id = Turb.Cols("Id").Z(j)
            
			If gen_name = Turb_name Then
				gen.Cols("ARSId").Z(i) = Turb_id
                jj = jj + 1
				t.Printp("����� " & jj & ", �������: " & Turb_name & "(" & Turb_id & ")" & " => ���������: " & gen_name & ", " & "(" & gen_Turbina & ")")
                Exit For
            End If
        Next
    Next 
End If

If Adjustment_ARS = 1 Then
    jj = 0
	For i = 0 to CountTurb
        Turb_name = Turb.Cols("Name").Z(i)
		Turb_id = Turb.Cols("Id").Z(i)	
        
		For j = 0 to CountARS
			ARS_name = ARS.Cols("Name").Z(j)
			ARS_id = ARS.Cols("Id").Z(j)
            
			If Turb_name = ARS_name Then
				Turb.Cols("Id").Z(i) = ARS_id
                jj = jj + 1
                t.Printp("����� " & jj & ", �������: " & Turb_name & "(" & Turb_id & ")" & " => ���(��): " & ARS_name & ", " & "(" & ARS_id & ")")
                Exit For
            End If
        Next
    Next 
End If

If Adjustment_Forsing = 1 Then
    jj = 0
	For i = 0 to CountVozb
        Vozb_name = vozb.Cols("Name").Z(i)
		Vozb_id = vozb.Cols("ForcerId").Z(i)
		
        For j = 0 to CountFors
            Fors_id = Forc.Cols("Id").Z(j)
            Fors_name = Forc.Cols("Name").Z(j)
            
			If  Vozb_name = Fors_name Then
				vozb.Cols("ForcerId").Z(i) = Fors_id
                jj = jj + 1
                t.Printp("����� " & jj & ", �����������: " & Vozb_name & "(" & Vozb_id & ")" & " => ����������(��): " & Fors_name & ", " & "(" & Fors_id & ")")
                Exit For
            End If
        Next
      Next 
End If

If OFF_Generators_with_P_and_Q_zero = 1 Then
	jj = 0
    For i = 0 to CountGen
        Pgen = gen.Cols("P").Z(i)
        Qgen = gen.Cols("Q").Z(i)
        If Pgen = 0 and Qgen = 0 Then
            gen.Cols("sta").Z(i) = 1
			jj = jj + 1
		End If
    Next
	t.Printp("���-�� ����������� ������� ���� ����-�: " & jj)
End If

t.Printp("������������ ���������")
