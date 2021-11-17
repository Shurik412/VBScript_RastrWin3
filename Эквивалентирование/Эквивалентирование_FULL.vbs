' 	 ������ ��� ������������ ������������ ��������� ������ (���) - 2020
' 
' 1. ������������������ - ��� (������������� ���)
' 2. ���������� ������ ����� ������������������ (�������� ����� ��� �����, �������� ��� � ��������� ��� ����� ... )
' 3. ���������� ����������� ������������� ������ �� Excel ����� 
'
'**************************************************************************

r=Setlocale("en-us")
rrr=1

Set t = Rastr
Set node = t.Tables("node")
Set vetv = t.Tables("vetv")
Set Generator = t.Tables("Generator")
Set ti = t.Tables("ti")
Set Reactors = t.Tables("Reactors")
Set area = t.tables("area")

print("������ ������� " & "����: " & date() & " | �����: " & Hour(Now()) & " hour " & Minute(Now()) & " minut")
Time_1 = Timer()
Call main() ' - ������������������ ���.
Time_2 = Timer()
print(" ����� ������ �������, � ������� = " & ((Time_2 - Time_1)/60))
print("������ ������� ���������.")

'\\************************************************************************
Sub main()
	'************************************************************
	' ����������: �������� ��������� ��� ������������������.
	' ������� ���������:  Nothing
	' �������:    Nothing
	'************************************************************
	
	'###########################################
	' ���� 
	viborka_ot_100_do_200_full = "(na>100 & na<200)" ' ��� ������������ ��� ����� ����������� �� ����������
	viborka_ot_100_do_200 = "(na>100 & na<200) & (uhom<230)"  ' ��� ������������������ � ������ �����. �� ����������
	
	'###########################################
	' 201 => ��������� ������� (��� => 813)
	na_Samarskay_obl = na_of_the_area_by_name("��������� �������")
	viborka_201_Samarskay_obl_full = "na=" & na_Samarskay_obl
	viborka_201_Samarskay_obl = "(na="& na_Samarskay_obl &") & (uhom < 160)"
	
	'###########################################
	' 205 => ���������� ��������� (���������) (��� => 205)
	na_Tatrskay = na_of_the_area_by_name("���������� ��������� (���������)")
	viborka_205_Tatrskay_full = "na=" & na_Tatrskay
	viborka_205_Tatrskay = "(na="& na_Tatrskay &") & (uhom < 160)"
	
	'###########################################
	' 206 => ��������� ���������� - ������� (��� => 206)
	na_Chuvashy = na_of_the_area_by_name("��������� ���������� - �������")
	viborka_206_Chuvashy_full = "na=" & na_Chuvashy
	viborka_206_Chuvashy = "(na="& na_Chuvashy &") & (uhom < 160)"
	
	'###########################################
	' 208 => ���������� ����� �� (��� => 208)
	na_MariEl = na_of_the_area_by_name("���������� ����� ��")
	viborka_208_MariEl_full = "na=" & na_MariEl
	viborka_208_MariEl = "(na="& na_MariEl &") & (uhom < 160)"
	
	'###########################################
	' 202 => ����������� ������� (��� => 202)
	na_Saratov_obl = na_of_the_area_by_name("����������� �������")
	viborka_202_Saratov_obl_full = "na=" & na_Saratov_obl
	viborka_202_Saratov_obl = "(na="& na_Saratov_obl &") & (uhom < 160)"
	
	'###########################################
	' 301 => ���������� ������� (��� => 301)
	na_Rostov_obl = na_of_the_area_by_name("���������� �������")
	viborka_301_Rostov_obl_full = "na=" & na_Rostov_obl
	viborka_301_Rostov_obl = "(na="& na_Rostov_obl &") & (uhom < 160)"
	
	'###########################################
	' 203 => ����������� ������� (��� => 203)
	na_Ulynov_obl = na_of_the_area_by_name("����������� �������")
	viborka_203_Ulynov_obl_full = "na=" & na_Ulynov_obl
	viborka_203_Ulynov_obl = "(na="& na_Ulynov_obl &") & (uhom < 160)"
	
	'###########################################
	' 401 => ���������� ������� (��� => 401)
	na_Murmansk_obl = na_of_the_area_by_name("���������� �������")
	viborka_401_Murmansk_obl_full = "na=" & na_Murmansk_obl
	viborka_401_Murmansk_obl = "(na="& na_Murmansk_obl &") & (uhom < 160)"
	
	'###########################################
	' 402 => ���������� ������� (��� => 402)
	na_Kareliy = na_of_the_area_by_name("���������� �������")
	viborka_402_Kareliy_full = "na=" & na_Kareliy
	viborka_402_Kareliy = "(na="& na_Kareliy &") & (uhom < 160)"
	
	'###########################################
	' 405 => ��������� ������� (��� => 405)
	na_Pskovskay_obl = na_of_the_area_by_name("��������� �������")
	viborka_405_Pskovskay_obl_full = "na=" & na_Pskovskay_obl
	viborka_405_Pskovskay_obl = "(na="& na_Pskovskay_obl &") & (uhom < 160)"
	
	'###########################################
	' 407 => ��������������� ������� (��� => 407)
	na_Kaliningrad_obl = na_of_the_area_by_name("��������������� �������")
	viborka_407_Kaliningrad_obl_full = "na=" & na_Kaliningrad_obl
	viborka_407_Kaliningrad_obl = "(na="& na_Kaliningrad_obl &") & (uhom < 160)"
	
	'###########################################
	' 805 => ��������� ���������� (��� => 805)
	na_Estony = na_of_the_area_by_name("��������� ����������")
	viborka_805_Estony_full = "na=" & na_Estony
	viborka_805_Estony = "(na="& na_Estony &") & (uhom < 160)"
	
	'###########################################
	' 806 => ���������� ���������� (��� => 806)
	na_Latviy = na_of_the_area_by_name("���������� ����������")
	viborka_806_Latviy_full = "na=" & na_Latviy
	viborka_806_Latviy = "(na="& na_Latviy &") & (uhom < 160)"
	
	'###########################################
	' 807 => ��������� ���������� (��� => 807)
	na_Litva = na_of_the_area_by_name("��������� ����������")
	viborka_807_Litva_full = "na=" & na_Litva
	viborka_807_Litva = "(na="& na_Litva &") & (uhom < 160)"
	
	'###########################################
	' 801 => ����������� ���������� (��� => 801)
	na_Finskay = na_of_the_area_by_name("����������� ����������")
	viborka_801_Finskay_full = "na=" & na_Finskay
	viborka_801_Finskay = "(na="& na_Finskay &") & (uhom < 160)"
	
	'###########################################
	' 823 => ���������� ������ (��� => 823)
	na_Donbas = na_of_the_area_by_name("���������� ������")
	viborka_823_Donbas_full = "na=" & na_Donbas
	viborka_823_Donbas = "(na="& na_Donbas &") & (uhom < 160)"
	 
	'###########################################
	' 825 => ������������ ������� (��� => 825 (����_max - 831))
	na_Orenburg_obl = na_of_the_area_by_name("������������ �������")
	viborka_825_Orenburg_obl_full = "na=" & na_Orenburg_obl
	viborka_825_Orenburg_obl = "(na="& na_Orenburg_obl &") & (uhom < 160)"
	
	'###########################################
	Call control_rgm()
	Call equivalent_settings()
	Call zeroing()
	Call off_the_line_from_two_side()
	Call control_rgm()
	Call off_the_generator_if_the_node_off()
	Call control_rgm()
	Call zeroing()
    kod = t.rgm("p")
	if kod<>(-1) then
		Call deleting_switches_by_selection(viborka_ot_100_do_200_full)
		Call deleting_generator_switches()
		Call zeroing()
		Call equalization_of_the_Urals_energy_system(viborka_ot_100_do_200)
		Call rastr_ekv()
		Call control_rgm()
		'###################################################
		Call zeroing()
		Call deleting_switches_by_selection(viborka_201_Samarskay_obl_full)
		Call zeroing()
		Call equivalent_to_generator_nodes(viborka_201_Samarskay_obl)
		Call rastr_ekv()
		Call control_rgm()
		'###################################################
		Call zeroing()
		Call deleting_switches_by_selection(viborka_205_Tatrskay_full)
		Call zeroing()
		Call equivalent_to_generator_nodes(viborka_205_Tatrskay)
		Call rastr_ekv()
		Call control_rgm()
		'###################################################
		Call zeroing()
		Call deleting_switches_by_selection(viborka_206_Chuvashy_full)
		Call zeroing()
		Call equivalent_to_generator_nodes(viborka_206_Chuvashy)
		Call rastr_ekv()
		Call control_rgm()
		'###################################################
		Call zeroing()
		Call deleting_switches_by_selection(viborka_208_MariEl_full)
		Call zeroing()
		Call equivalent_to_generator_nodes(viborka_208_MariEl)
		Call rastr_ekv()
		Call control_rgm()
		'###################################################
		Call zeroing()
		Call deleting_switches_by_selection(viborka_202_Saratov_obl_full)
		Call zeroing()
		Call equivalent_to_generator_nodes(viborka_202_Saratov_obl)
		Call rastr_ekv()
		Call control_rgm()
		'###################################################
		Call zeroing()
		Call deleting_switches_by_selection(viborka_301_Rostov_obl_full)
		Call zeroing()
		Call equivalent_to_generator_nodes(viborka_301_Rostov_obl)
		Call rastr_ekv()
		Call control_rgm()
		'###################################################
		Call zeroing()
		Call deleting_switches_by_selection(viborka_203_Ulynov_obl_full)
		Call zeroing()
		Call equivalent_to_generator_nodes(viborka_203_Ulynov_obl)
		Call rastr_ekv()
		Call control_rgm()
		'###################################################
		Call zeroing()
		Call deleting_switches_by_selection(viborka_401_Murmansk_obl_full)
		Call zeroing()
		Call equivalent_to_generator_nodes(viborka_401_Murmansk_obl)
		Call rastr_ekv()
		Call control_rgm()
		'###################################################
		Call zeroing()
		Call deleting_switches_by_selection(viborka_402_Kareliy_full)
		Call zeroing()
		Call equivalent_to_generator_nodes(viborka_402_Kareliy)
		Call rastr_ekv()
		Call control_rgm()
		'###################################################
		Call zeroing()
		Call deleting_switches_by_selection(viborka_405_Pskovskay_obl_full)
		Call zeroing()
		Call equivalent_to_generator_nodes(viborka_405_Pskovskay_obl)
		Call rastr_ekv()
		Call control_rgm()
		'###################################################
		Call zeroing()
		Call deleting_switches_by_selection(viborka_407_Kaliningrad_obl_full)
		Call zeroing()
		Call equivalent_to_generator_nodes(viborka_407_Kaliningrad_obl)
		Call rastr_ekv()
		Call control_rgm()	
		'###################################################
		Call zeroing()
		Call deleting_switches_by_selection(viborka_805_Estony_full)
		Call zeroing()
		Call equivalent_to_generator_nodes(viborka_805_Estony)
		Call rastr_ekv()
		Call control_rgm()	
		'###################################################
		Call zeroing()
		Call deleting_switches_by_selection(viborka_806_Latviy_full)
		Call zeroing()
		Call equivalent_to_generator_nodes(viborka_806_Latviy)
		Call rastr_ekv()
		Call control_rgm()	
		'###################################################
		Call zeroing()
		Call deleting_switches_by_selection(viborka_807_Litva_full)
		Call zeroing()
		Call equivalent_to_generator_nodes(viborka_807_Litva)
		Call rastr_ekv()
		Call control_rgm()	
		'###################################################
		Call zeroing()
		Call deleting_switches_by_selection(viborka_801_Finskay_full)
		Call zeroing()
		Call equivalent_to_generator_nodes(viborka_801_Finskay)
		Call rastr_ekv()
		Call control_rgm()	
		'###################################################	
		Call zeroing()
		Call deleting_switches_by_selection(viborka_823_Donbas_full)
		Call zeroing()
		Call equivalent_to_generator_nodes(viborka_823_Donbas)
		Call rastr_ekv()
		Call control_rgm()
		'###################################################	
		Call zeroing()
		Call deleting_switches_by_selection(viborka_825_Orenburg_obl_full)
		Call zeroing()
		Call control_rgm()
		Call equivalent_to_generator_nodes(viborka_825_Orenburg_obl)
		Call rastr_ekv()
		Call control_rgm()
        '###################################################
        Call zeroing()
        Call control_rgm()
		Call removing_nodes_without_branches()
		Call control_rgm()
		Call Delete_Generator_without_nodes()
		Call control_rgm()
		Call reactors_change()
		Call control_rgm()
	else
		print("--- ������: ����� ����������! ---")
		print("--- ������ ������� ��������� ��������! ---")
	end if
End Sub

Sub equalization_of_the_Urals_energy_system(selection_of_the_area)
	'************************************************************
	' ����������: ������������������ ��� �����
	' ������� ���������: selection_of_the_area - ������� ������� ��� �����
	' �������:    Nothing
	'************************************************************
	node.SetSel(selection_of_the_area) ' ������� �� �����
    node.Cols("sel").Calc("1")
    j = node.FindNextSel(-1)
	while j<>(-1)
        ny = node.Cols("ny").Z(j)
        tip_node = node.Cols("tip").Z(j)
        uhom = node.Cols("uhom").Z(j)
        if tip_node > 1 Then ' ��� ������������ ����
            Generator.SetSel("Node.ny=" & ny)
            j_Generator = Generator.FindNextSel(-1)
            if j_Generator <> (-1) then
                vetv.SetSel("(ip= " & ny & ")|(iq= " & ny & ")")
                j_vetv = vetv.FindNextSel(-1)
                while j_vetv <>(-1)
                    tip_vetv = vetv.Cols("tip").Z(j_vetv)
                    if tip_vetv = 1 then
                        v_ip = vetv.Cols("v_ip").Z(j_vetv) 
                        v_iq = vetv.Cols("v_iq").Z(j_vetv)
                        if (v_ip > 430 and v_iq < 580) or (v_ip < 430 and v_iq > 580) then
                            node.Cols("sel").Z(j) = 0
                        end if
                    end if
                    j_vetv = vetv.FindNextSel(j_vetv)
                wend
            end if
        else
            vetv.SetSel("(ip= " & ny & ")|(iq= " & ny & ")")
            j_vetv_2 = vetv.FindNextSel(-1)
            while j_vetv_2<>(-1)
				tip_vetv_2 = vetv.Cols("tip").Z(j_vetv_2)
                if tip_vetv_2 = 1 then
					v_ip_2 = vetv.Cols("v_ip").Z(j_vetv_2) 
                    v_iq_2 = vetv.Cols("v_iq").Z(j_vetv_2)
                    if (v_ip_2 > 430 and v_iq_2 < 580) or (v_ip_2 < 430 and v_iq_2 > 580) then
						node.Cols("sel").Z(j) = 0
					end if
                end if
                j_vetv_2 = vetv.FindNextSel(j_vetv_2)
			wend
        end If
        node.SetSel(selection_of_the_area)
		j = node.FindNextSel(j)
    wend
	print(" -> ��������� ��������� ������(-��): " & selection_of_the_area)
End Sub

Sub equivalent_to_generator_nodes(vyborka_Generator)
	'************************************************************
	' ����������: ������������������ ������������ �����.
	' �������
	' ���������:  
	' �������:    Nothing
	'************************************************************
	node.SetSel(vyborka_Generator) ' ������� �� �����
    node.Cols("sel").Calc("1")
    j = node.FindNextSel(-1) 
    While j<>(-1)
        ny = node.Cols("ny").Z(j)
        tip_node = node.Cols("tip").Z(j)
        uhom = node.Cols("uhom").Z(j)
        If tip_node > 1 Then ' ��� ������������ ����
            Generator.SetSel("Node.ny=" & ny)
            j_Generator = Generator.FindNextSel(-1)
            if j_Generator <> (-1) then
                vetv.SetSel("(ip= " & ny & ")|(iq= " & ny & ")")
                j_vetv = vetv.FindNextSel(-1)
                while j_vetv <>(-1)
                    tip_vetv = vetv.Cols("tip").Z(j_vetv)
                    if tip_vetv = 1 then
                        v_ip = vetv.Cols("v_ip").Z(j_vetv) 
                        v_iq = vetv.Cols("v_iq").Z(j_vetv)
                        if (v_ip > 170 and v_iq < 250) or (v_ip < 170 and v_iq > 250) then
                            node.Cols("sel").Z(j) = 0
                        end if
                    end if
                    j_vetv = vetv.FindNextSel(j_vetv)
                wend
            end if
        Else
            vetv.SetSel("(ip="& ny &")|(iq="& ny &")")
            j_vetv_2 = vetv.FindNextSel(-1)
            while j_vetv_2 <>(-1)
				tip_vetv_2 = vetv.Cols("tip").Z(j_vetv_2)
                if tip_vetv_2 = 1 then
					v_ip_2 = vetv.Cols("v_ip").Z(j_vetv_2) 
					v_iq_2 = vetv.Cols("v_iq").Z(j_vetv_2)
                    if (v_ip_2 > 170 and v_iq_2 < 250) or (v_ip_2 < 170 and v_iq_2 > 250) then
						node.Cols("sel").Z(j) = 0
                    end if
                end if
                j_vetv_2 = vetv.FindNextSel(j_vetv_2)
            wend
        End If
        node.SetSel(vyborka_Generator)
		j = node.FindNextSel(j)
    Wend
	print(" -> ��������� ��������� ������(-��): " & vyborka_Generator )
End Sub

Sub deleting_switches_by_selection(viborka_ray_vikl)
	'************************************************************
	' ����������: �������� ����������� �� ���������� ������
	' ������� ���������: pra: viborka_ray_vikl - �������
	' �������:    Nothing
	'************************************************************
    Dim nodes(30000)
	
	node.SetSel(viborka_ray_vikl) ' ������� ����� ���� ������� ����� 500 (������)
    node.Cols("sel").calc(1) ' ��������� �������� �����
    vetv.SetSel("iq.sel=1 & ip.sel=0 &!sta") ' ������� ������ iq.sel = 1 ...
    k = vetv.FindNextSel(-1)
	While k<>(-1) ' ������� sel-���� ���� �� �� � ����� ������� ������� ���� 
		iq1 = vetv.Cols("iq").Z(k)
		node.Setsel("ny=" & iq1)
		k2 = node.FindNextSel(-1)
		If k2<>(-1) Then
			node.Cols("sel").Z(k2) = 0
		End If
		k = vetv.FindNextSel(k)
    Wend
	
    vetv.SetSel("iq.sel=0 & ip.sel=1 & !sta")
    k = vetv.FindNextSel(-1)
	
    While k<>(-1) ' ������� sel-���� ���� �� �� � ����� ������� ������� ���� 
		ip1 = vetv.Cols("ip").Z(k)
		node.Setsel("ny=" & ip1)
		k2 = node.FindNextSel(-1)
		If k2<>(-1) Then
			node.Cols("sel").Z(k2) = 0
		End If
		k = vetv.FindNextSel(k)
	Wend
	 
	vetv.SetSel("(iq.sel=1 & ip.sel=0)|(ip.sel=1 & iq.sel=0) & tip=2") ' tip=2 - ����������� (������� ���� ������������ ���� ������ � ����� ������� ������� ���� sel)
    k = vetv.FindNextSel(-1)
    While k<>(-1)
		iq1 = vetv.Cols("iq").Z(k)
		node.Setsel("ny=" & iq1)
		k2 = node.FindNextSel(-1)
		If k2<>(-1) Then
			node.Cols("sel").Z(k2) = 0
		End If
		ip1 = vetv.Cols("ip").Z(k)
		node.Setsel("ny=" & ip1)
		k2 = node.FindNextSel(-1)
		If k2<>(-1) Then
			node.Cols("sel").Z(k2) = 0
		End If
		vetv.SetSel("(iq.sel=1 & ip.sel=0)|(ip.sel=1 & iq.sel=0) & tip=2")
		k = vetv.FindNextSel(-1)
    Wend
	
    vetvyklvybexc = "(iq.bsh>0 & ip.bsh=0)|(ip.bsh>0 & iq.bsh=0)|(iq.bshr>0 & ip.bshr=0)|(ip.bshr>0 & iq.bshr=0)|ip.sel=0|iq.sel=0)"
    flvykl = 0
	vetv.SetSel("1")
	vetv.Cols("groupid").calc(0)
	vetv.SetSel(vetvyklvybexc)
	vetv.Cols("groupid").calc(1)
	nvetv = 0
	' �������� ������������
	for povetv = 0 to 10000
		'������� ������, ������� ������� �������������
		vetv.SetSel("x<0.01 & x>-0.01 & r<0.005 & r>=0 & (ktr=0 | ktr=1) & !sta & groupid!=1 & b<0.000005") 
		ivetv = vetv.FindNextSel(-1)
		If ivetv = -1 Then exit for
            ip = vetv.Cols("ip").Z(ivetv)
            iq = vetv.Cols("iq").Z(ivetv)
            If ip > iq Then
                ny = iq 
                ndel = ip
            Else 
                ny = ip
                ndel = iq
            End If
			
            ndny = 0
            ndndel = 0
			'�������� �� ������� ���� �� ������ �����������
            For inodee = 0 to nnod
                If 	ndel = nodes(inodee) Then ndndel = 1
                If 	ny = nodes(inodee) Then ndny = 1
                If (ndndel = 1) and (ndny = 1) Then exit for
            Next
			' ������ �������, ��� ��� ��������� ������ �������, � ����������� ����� ))
            If (ndndel = 0) and (ndny = 1) Then
                buff = ny
                ny = ndel
                ndel = buff
            End If
			
            If (ndndel = 0) or (ndny = 0) Then '���� ���� �� ���� ����� �������
                flvykl = flvykl + 1
				node.SetSel("ny=" & ny)
				iny = node.FindNextSel(-1)
				node.SetSel("ny=" & ndel)
				idel = node.FindNextSel(-1)
				pgdel = node.cols("pg").Z(idel)
				qgdel = node.cols("qg").Z(idel)
				pndel = node.cols("pn").Z(idel)
				qndel = node.cols("qn").Z(idel)
				bshdel = node.cols("bsh").Z(idel)
				gshdel = node.cols("gsh").Z(idel)
				pgny = node.cols("pg").Z(iny)
				qgny = node.cols("qg").Z(iny)
				pnny = node.cols("pn").Z(iny)
				qnny = node.cols("qn").Z(iny)
				bshny = node.cols("bsh").Z(iny)
				gshny = node.cols("gsh").Z(iny)
                 
				node.cols("pg").Z(iny) = pgdel + pgny
				node.cols("qg").Z(iny) = qgdel + qgny
				node.cols("pn").Z(iny) = pndel + pnny
				node.cols("qn").Z(iny) = qndel + qnny
				node.cols("bsh").Z(iny) = bshdel + bshny
				node.cols("gsh").Z(iny) = gshdel + gshny
				v1 = node.cols("vzd").Z(iny)
				v2 = node.cols("vzd").Z(idel)
				qmax1 = node.cols("qmax").Z(iny)
				qmax2 = node.cols("qmax").Z(idel)
				 
				Generator.Setsel("Node=" & ndel)
				iGenerator = Generator.FindNextSel(-1) '������ ���� ����������� �����������
				 
				If iGenerator<>(-1) Then
					While iGenerator<>(-1) 
						Generator.cols("Node").Z(iGenerator) = ny
						iGenerator = Generator.FindNextSel(iGenerator)
					Wend
				End If
					
				If (v1<>v2) and (v1>0.3) and (v2>0.3) and (qmax1 + qmax2) <> 0 Then
					node.cols("vzd").Z(iny) = (v1*qmax1+v2*qmax2)/(qmax1+qmax2) 
					'������ ���������������� �� qmax ����������
				End If
					
				If (v1=0) and (v2<>0) Then
					node.Cols("vzd").Z(iny) = v2
				End If
					
				If (v1<>0) and (v2<>0) Then
					node.Cols("qmin").Z(iny) = (node.Cols("qmin").Z(iny)) + (node.Cols("qmin").Z(idel))
					node.Cols("qmax").Z(iny) = qmax1 + qmax2
				End If
					
				If (v1=0) and (v2<>0) Then
					node.cols("qmin").Z(iny) = node.Cols("qmin").Z(idel)
					node.cols("qmax").Z(iny) = node.Cols("qmax").Z(idel)
				End If
					
				vetv.SetSel("(ip=" & ip & "& iq=" & iq & ")|(iq=" & ip & "& ip=" & iq & ")")
				vetv.delrows '������� �����	
				vetv.SetSel("iq=" & ndel) '������ ���� ������ � ��������� �����)))
				vetv.cols("iq").Calc(ny)	
				vetv.SetSel("ip=" & ndel)
				vetv.cols("ip").Calc(ny)	
				node.delrows 		' ������� ����
			Else '���� �� ������ ������ �������
                vetv.SetSel("(ip=" & ip & "& iq=" & iq & ")|(iq=" & ip & "& ip=" & iq & ")")
                vetv.cols("groupid").Calc(1)
			End If
    next
	Call control_rgm()
End Sub

Sub zeroing()
    '************************************************************
	' ����������:  ��������� ���� sel (���������� ��������) ����� � ������.
	' ������� ���������: 
	' �������:    Nothing
	'************************************************************   
    vetv.SetSel("")
	vetv.Cols("sel").Calc("0")
	node.SetSel("")
	node.Cols("sel").Calc("0")
	print(" - ����� '�������' � ���������� ����� � ������.")
End Sub

Sub deleting_generator_switches()
	'************************************************************
	' ����������:  ������� ����������� �����������.
	' ������� ���������: 
	' �������:    Nothing
	'************************************************************  
	Call zeroing()
	node.SetSel("")
	k1=node.findnextsel(-1)
	While k1<>(-1)
		ny1=node.Cols("ny").Z(k1)
		vetv.SetSel("(ip=" & ny1 &")|(iq=" & ny1 &")")
		if vetv.Count=1 then
			vetv.SetSel("x<1 & (tip=0|tip=2)&((ip=" & ny1 & ")|(iq=" & ny1 &"))")
			if vetv.Count=1 then
				vetv.SetSel("x<1&(tip=0|tip=2)&((ip=" & ny1 & ")|(iq=" & ny1 &"))")
				k3=vetv.findnextsel(-1)
				if k3<>(-1) then
					if vetv.Cols("ip").Z(k3)=ny1 then
						ny2=vetv.Cols("iq").Z(k3)
					else
						ny2=vetv.Cols("ip").Z(k3)
					end if
					Generator.SetSel("Node=" & ny1)
					k2=Generator.findnextsel(-1)
					if k2<>(-1) then
						node.SetSel("ny=" & ny2)
						k4=node.findnextsel(-1)
						if k4<>(-1) then
							node.Cols("pn").Z(k4) = node.Cols("pn").Z(k1) + node.Cols("pn").Z(k1)
							node.Cols("qn").Z(k4) = node.Cols("qn").Z(k1) + node.Cols("qn").Z(k1)
							node.Cols("vzd").Z(k4) = node.Cols("vzd").Z(k1)
							node.Cols("exist_load").Z(k4) = node.Cols("exist_load").Z(k1)
							node.Cols("exist_gen").Z(k4) = node.Cols("exist_gen").Z(k1)
							node.Cols("pn_max").Z(k4) =node.Cols("pn_max").Z(k1) + node.Cols("pn_max").Z(k4)
							if node.Cols("pn_min").Z(k4) => node.Cols("pn_min").Z(k1) then
								node.Cols("pn_min").Z(k4) = node.Cols("pn_min").Z(k1)
							end if
							node.Cols("pg_max").Z(k4) = node.Cols("pg_max").Z(k1) + node.Cols("pg_max").Z(k4)
							if node.Cols("pg_min").Z(k4) => node.Cols("pg_min").Z(k1) then
								node.Cols("pg_min").Z(k4) = node.Cols("pg_min").Z(k1)
							end if
							node.Cols("sel").Z(k1) = 1
							vetv.Cols("sel").Z(k3) = 1
							' ti.SetSel("(prv_num=20 | prv_num=7 | prv_num=6 | prv_num=5 | prv_num=4 | prv_num=3 | prv_num=2 | prv_num=1) & id1="&ny1)
							' ti.cols("id1").calc(ny2)
							Generator.SetSel("Node=" & ny1)
							k2 = Generator.findnextsel(-1)
							while k2 <> (-1)
								Generator.cols("Node").Z(k2) = ny2
								k2=Generator.findnextsel(k2)
							wend
						end if
					end if
				end if
			end if
		end if
		node.SetSel("")
		k1=node.findnextsel(k1)
	Wend
	vetv.SetSel("sel=1")
	vetv.delrows
	node.SetSel("sel=1")
	node.delrows
End Sub

Sub equivalent_settings()
	'************************************************************
	' ����������: ���������� ��������� ������������������
	' ������� ���������: Nothing
	' �������:    Nothing
	'************************************************************  
    print(" - ����������� ��������� ������������������;")
	t.Tables("com_ekviv").Cols("zmax").Z(0) = 1000
	t.Tables("com_ekviv").Cols("ek_sh").Z(0) = 0
	t.Tables("com_ekviv").Cols("otm_n").Z(0) = 0
	t.Tables("com_ekviv").Cols("smart").Z(0) = 0
	t.Tables("com_ekviv").Cols("tip_ekv").Z(0) = 0
	t.Tables("com_ekviv").Cols("ekvgen").Z(0) = 0
	t.Tables("com_ekviv").Cols("tip_gen").Z(0) = 1
End Sub

Sub off_the_generator_if_the_node_off()
	'************************************************************
	' ����������: ���������� ���������, ���� ���� � �������� ���������  
	'             ��������� ��������.
	' ������� ���������: Nothing
	' �������:    Nothing
	'************************************************************  
	Generator.setsel("")
	k = Generator.FindNextSel(-1)
	while k<>(-1)
		nyGenerator = Generator.cols("Node").Z(k)
		node.SetSel("ny=" & nyGenerator)
		kk = node.findnextsel(-1)
		if kk <> (-1) then
			if node.cols("sta").Z(kk) = True then
				Generator.cols("sta").Z(k) = 1
			end if
		end if
		Generator.setsel("")
		k = Generator.FindNextSel(k)
	wend
	print(" - ��������� ���������� � ����������� �����.")
End Sub

Sub off_the_line_from_two_side()
	'************************************************************
	' ����������: ���������� ��� � ���� ������, ���� ��� �������� � ����� �������.
	' ������� ���������: Nothing
	' �������:    Nothing
	'************************************************************  
	ii = 0
	vetvMaxRow = vetv.Count-1
	for i = 0 to vetvMaxRow
		sta = vetv.Cols("sta").Z(i)
		If sta = 2 or sta = 3 Then
			vetv.Cols("sta").Z(i) = 1
			ii = ii + 1
		end if
	next
	print(" - ���������� ��� � ������������� ���., ������������ � ��������� ������� ���������: " & ii)
End Sub

Function na_of_the_area_by_name(name_area)
	'************************************************************
	' ����������: 
	' ������� ���������: Nothing
	' �������:    Nothing
	'************************************************************  
    max_count_area = area.Count-1
    for i=0 to max_count_area 
        name_ = area.Cols("name").Z(i)
        if name_ = name_area then
            na_of_the_area_by_name = area.Cols("na").Z(i)
			print(" - �������� ������: "& name_ &"; ����� ������: "& na_of_the_area_by_name)
		end if 
    next
End function

Function removing_nodes_without_branches()
	'************************************************************
	' ����������: �������� ����� ��� ����� � �������
	' ������� ���������: Nothing
	' �������:    Nothing
	'************************************************************  
	nodeColMax = node.Count-1
	vetvColMax = vetv.Count-1
	ii = 0
	for i=0 to nodeColMax
		Bsh = node.Cols("bsh").Z(i)
		id_ny = node.Cols("ny").Z(i)
		vetv.SetSel("ip.ny=" & id_ny & "|iq.ny=" & id_ny)
		colvetv = vetv.FindNextSel(-1)
		key_1 = 1
         
		If key_1=1 Then
			node.Cols("sel").Z(i) = 0
			If colvetv=(-1) Then 
				node.Cols("sel").Z(i) = 1
				ii = ii + 1
			End If
		End If
         
		If key_1=0 Then
			vetv.Cols("sel").Z(i) = 0
			If colvetv<>(-1) Then
				type_id = vetv.Cols("tip").Z(colvetv)    
				If type_id=2 Then
				   If Bsh=0 Then
						vetv.Cols("sel").Z(colvetv) = 1
				   End If
				End If
			 End If
		End If
	next
    node.SetSel("sel=1")
	ii = node.Count-1
	node.DelRows
	print(" - ������� ����� ��� ������ � �������: " & ii)
	Call control_rgm()
End Function

Sub Delete_Generator_without_nodes()
	'************************************************************
	' ����������: 
	' ������� ���������: Nothing
	' �������:    Nothing
	'************************************************************  
	Generator.SetSel("node.ny=0")
	Generator.DelRows
End Sub

Sub reactors_change()
	'************************************************************
	' ����������: 
	' ������� ���������: Nothing
	' �������:    Nothing
	'************************************************************  
	Reactors.SetSel("")
	Reactors.Cols("sel").Calc(0)
	Reactors.SetSel("")
	k=Reactors.FindNextSel(-1)
	while k<>(-1)
		ip1=Reactors.Cols("Id1").Z(k)
		B1=Reactors.Cols("B").Z(k)
		reac_sta=Reactors.Cols("sta").Z(k)
		node.SetSel("ny=" & ip1  )
		if node.count > 0 then
			k2=node.FindNextSel(-1)
			while k2<>(-1)
				node.Cols("bsh").Z(k2) = node.Cols("bsh").Z(k2) + B1
				if reac_sta = 1 then
					node.Cols("sel").Z(k2) = 1
				end if
				k2=node.FindNextSel(k2)
			wend
		end if
		k=Reactors.FindNextSel(k)
	wend
	Reactors.SetSel("")
	Reactors.Delrows
End Sub 

Sub rastr_ekv()
	'************************************************************
	' ����������: ������ ������������������
	' ������� ���������:  Nothing
	' �������:    Nothing
	'************************************************************
	t.Ekv("")
	print(" - ������: ������������������!")
End Sub

Sub control_rgm()
	'************************************************************
	' ����������: ������ ������� �� ������� �������
	' ������� ���������:  
	' �������:    Nothing
	'************************************************************
	kod = t.rgm("p")
	if kod<>(-1) then
		print(" - ����� �������!")
	else
		print(" - ����� ����������!")
	end if
End Sub

Sub print(msg)
	'************************************************************
	' ����������: ������� ��������� (msg) � ��������
	' ������� ���������:  par: msg; type: string
	' �������:    Nothing
	'************************************************************
    t.Printp(msg)
End Sub