' 	 ������ ��� ������������ ������������ ��������� ������ (���) - 2020
' 
' 1. ������������������ - ��� (������������� ���)
' 2. ���������� ������ ����� ������������������ (�������� ����� ��� �����, �������� ��� � ��������� ��� ����� ... )
' 3. ���������� ����������� ������������� ������ �� Excel ����� 
'
'**************************************************************************

r=Setlocale("en-us")
rrr=1
Set t=Rastr

print("������ ������� " & "����: " & date() & " | �����: " & Hour(Now()) & " hour " & Minute(Now()) & " minut")
Time_1 = Timer()
Call Equivalence() ' - ������������������ ���.
Time_2 = Timer()
print("����� ������ �������, � ������� = " & ((Time_2 - Time_1)/60))
print("������ ������� ���������.")

'\\************************************************************************
Sub Equivalence()
	print("-= ������ ������������������ =-")
    Set node = t.Tables("node")
    Set vetv = t.Tables("vetv")
    Set gen = t.Tables("Generator")
	
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
	viborka_202_Saratov_obl = "(na=" & na_Saratov_obl & ") & (uhom < 160)"
	
	'###########################################
	' 301 => ���������� ������� (��� => 301)
	na_Rostov_obl = na_of_the_area_by_name("���������� �������")
	viborka_301_Rostov_obl_full = "na=" & na_Rostov_obl
	viborka_301_Rostov_obl = "(na=" & na_Rostov_obl & ") & (uhom < 160)"
	
	'###########################################
	' 203 => ����������� ������� (��� => 203)
	na_Ulynov_obl = na_of_the_area_by_name("����������� �������")
	viborka_203_Ulynov_obl_full = "na=" & na_Ulynov_obl
	viborka_203_Ulynov_obl = "(na=" & na_Ulynov_obl & ") & (uhom < 160)"
	
	'###########################################
	' 401 => ���������� ������� (��� => 401)
	na_Murmansk_obl = na_of_the_area_by_name("���������� �������")
	viborka_401_Murmansk_obl_full = "na=" & na_Murmansk_obl
	viborka_401_Murmansk_obl = "(na=" & na_Murmansk_obl & ") & (uhom < 160)"
	
	'###########################################
	' 402 => ���������� ������� (��� => 402)
	na_Kareliy = na_of_the_area_by_name("���������� �������")
	viborka_402_Kareliy_full = "na=" & na_Kareliy
	viborka_402_Kareliy = "(na="& na_Kareliy & ") & (uhom < 160)"
	
	'###########################################
	' 405 => ��������� ������� (��� => 405)
	na_Pskovskay_obl = na_of_the_area_by_name("��������� �������")
	viborka_405_Pskovskay_obl_full = "na=" & na_Pskovskay_obl
	viborka_405_Pskovskay_obl = "(na=" & na_Pskovskay_obl & ") & (uhom < 160)"
	
	'###########################################
	' 407 => ��������������� ������� (��� => 407)
	na_Kaliningrad_obl = na_of_the_area_by_name("��������������� �������")
	viborka_407_Kaliningrad_obl_full = "na=" & na_Kaliningrad_obl
	viborka_407_Kaliningrad_obl = "(na="& na_Kaliningrad_obl &") & (uhom < 160)"
	
	'###########################################
	' 805 => ��������� ���������� (��� => 805)
	na_Estony = na_of_the_area_by_name("��������� ����������")
	viborka_805_Estony_full = "na=" & na_Estony
	viborka_805_Estony = "(na="& na_Estony & ") & (uhom < 160)"
	
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
	' 204 => ���������� �������
	na_Penza_obl = na_of_the_area_by_name("���������� �������")
	viborka_204_Penza_obl_full = "na=" & na_Penza_obl
	viborka_204_Penza_obl = "(na="& na_Penza_obl &") & (uhom < 160)"
	
	'###########################################
	' 207 => ���������� ��������
	na_Republic_Mordova = na_of_the_area_by_name("���������� ��������")
	viborka_207_Republic_Mordova_full = "na=" & na_Republic_Mordova
	viborka_207_Republic_Mordova = "(na="& na_Republic_Mordova &") & (uhom < 160)"
	
	'###########################################
	' 209 => ������������� �������
	na_Nijegor_obl = na_of_the_area_by_name("������������� �������")
	viborka_209_Nijegor_obl_full = "na=" & na_Nijegor_obl
	viborka_209_Nijegor_obl = "(na="& na_Nijegor_obl &") & (uhom < 160)"
	
	'###########################################
	' 311 => ������������� �������
	na_Vologda_obl = na_of_the_area_by_name("������������� �������")
	viborka_311_Vologda_obl_full = "na=" & na_Vologda_obl
	viborka_311_Vologda_obl = "(na="& na_Vologda_obl &") & (uhom < 160)"
	
	'###########################################
	' 404 => ������������ �������
	na_Nigegor_obl = na_of_the_area_by_name("������������ �������")
	viborka_404_Nigegor_obl_full = "na=" & na_Nigegor_obl
	viborka_404_Nigegor_obl = "(na="& na_Nigegor_obl &") & (uhom < 160)"
	
	'###########################################
	' 803 => �������� ������
	na_Zapad_reg = na_of_the_area_by_name("�������� ������")
	viborka_803_Zapad_reg_full = "na=" & na_Zapad_reg
	viborka_803_Zapad_reg = "(na="& na_Zapad_reg &") & (uhom < 160)"
	
	'###########################################
	' 804 => �������� ������
	na_Belorus = na_of_the_area_by_name("�������� ������")
	viborka_804_Belorus_full = "na=" & na_Belorus
	viborka_804_Belorus = "(na="& na_Belorus &") & (uhom < 160)"
	
	'###########################################
	' 819 => ������
	na_Shvec = na_of_the_area_by_name("������")
	viborka_819_Shvec_full = "na=" & na_Shvec
	viborka_819_Shvec = "(na="& na_Shvec &") & (uhom < 160)"
	
	'###########################################
	' 820 => �����-���������
	na_SPB = na_of_the_area_by_name("�����-���������")
	viborka_820_SPB_full = "na=" & na_SPB
	viborka_820_SPB = "(na="& na_SPB &") & (uhom < 160)"
	
	'###########################################
	' 822 => ������������� �������
	na_Leningral_obl = na_of_the_area_by_name("������������� �������")
	viborka_822_Leningral_obl_full = "na=" & na_Leningral_obl
	viborka_822_Leningral_obl = "(na="& na_Leningral_obl &") & (uhom < 160)"
	
	'###########################################
	' 823 => ���-�������� ������
	na_Ugo_Zapad_reg = na_of_the_area_by_name("���-�������� ������")
	viborka_823_Ugo_Zapad_reg_full = "na=" & na_Ugo_Zapad_reg
	viborka_823_Ugo_Zapad_reg = "(na="& na_Ugo_Zapad_reg &") & (uhom < 160)"

	'###########################################
	' 825 => ����� ������
	na_Ugny_reg = na_of_the_area_by_name("����� ������")
	viborka_825_Ugny_reg_full = "na=" & na_Ugny_reg
	viborka_825_Ugny_reg = "(na="& na_Ugny_reg &") & (uhom < 160)"	
	
	'###########################################
	' 826 => ����������� ������
	na_Dnepov_reg = na_of_the_area_by_name("����������� ������")
	viborka_826_Dnepov_reg_full = "na=" & na_Dnepov_reg
	viborka_826_Dnepov_reg = "(na="& na_Dnepov_reg &") & (uhom < 160)"	
	
	'###########################################
	' 827 => �������� ������
	na_Sever_reg = na_of_the_area_by_name("�������� ������")
	viborka_827_Sever_reg_full = "na=" & na_Sever_reg
	viborka_827_Sever_reg = "(na="& na_Sever_reg &") & (uhom < 160)"	
	'###########################################
	Call Control_Rgm()
	Call Settings_Ekv()
	Call Obnulenie()
	Call Off_line_one_on()
	Call Control_Rgm()
	Call Off_Gen_if_off_node()
	Call Control_Rgm()
	Call Obnulenie()
    Call print_parm_shem()
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
		print("1")
        print(viborka_201_Samarskay_obl_full)
		Call Vikluchatel(viborka_201_Samarskay_obl_full)
		print("2")
		Call Obnulenie()
		print("3")
        print(viborka_201_Samarskay_obl)
		Call Ekvivalent_Node_Gen(viborka_201_Samarskay_obl)
		print("4")
		Call Rastr_Ekv()
		print("5")
		Call Control_Rgm()
		print("6")
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
		Call Obnulenie()
        '###################################################
        Call Control_Rgm()
		Call ReactorsChange()
		Call Obnulenie()
		Call Delete_Node_not_connect()
		Call Delete_Generators_without_nodes()
		Call Control_Rgm()
		Call print_parm_shem()
	
		' ############################################################################
		'viborka_ot_100_do_200_full = "(na>100 & na<200)" ' ��� ������������ ��� ����� ����������� �� ����������
		'na_Samarskay_obl -> 201 => ��������� �������
		'na_Tatrskay -> 205 => ���������� ��������� (���������)
		'na_Chuvashy -> 206 => ��������� ���������� - �������
		'na_MariEl -> 208 => ���������� ����� ��
		'na_Saratov_obl -> 202 => ����������� �������
		'na_Rostov_obl -> 301 => ���������� �������
		'na_Ulynov_obl -> 203 => ����������� �������
		'na_Murmansk_obl -> 401 => ���������� �������
		'na_Kareliy -> 402 => ���������� �������
		'na_Pskovskay_obl -> 405 => ��������� �������
		'na_Kaliningrad_obl -> 407 => ��������������� �������
		'na_Estony -> 805 => ��������� ����������
		'na_Latviy -> 806 => ���������� ����������
		'na_Litva -> 807 => ��������� ����������
		'na_Finskay -> 801 => ����������� ����������
		'na_Donbas -> 823 => ���������� ������
		'na_Orenburg_obl -> 825 => ������������ �������
		'na_Penza_obl -> 204 => ���������� �������
		'na_Republic_Mordova -> 207 => ���������� ��������
		'na_Nijegor_obl -> 209 => ������������� �������
		'na_Vologda_obl -> 311 => ������������� �������
		'na_Nigegor_obl -> 404 => ������������ �������
		'na_Zapad_reg -> 803 => �������� ������
		'na_Belorus -> 804 => �������� ������
		'na_Shvec -> 819 => ������
		'na_SPB -> 820 => �����-���������
		'na_Leningral_obl -> 822 => ������������� �������
		'na_Ugo_Zapad_reg -> 823 => ���-�������� ������
		'na_Ugny_reg -> 825 => ����� ������
		'na_Dnepov_reg -> 826 => ����������� ������
		'na_Sever_reg -> 827 => �������� ������
	
		'##############################################################################
		'viborka_ot_100_do_200_full -> ����
		'viborka_201_Samarskay_obl_full -> 201 => ��������� ������� (��� => 813)	
		'viborka_205_Tatrskay_full -> 205 => ���������� ��������� (���������) (��� => 205)
		'viborka_206_Chuvashy_full -> 206 => ��������� ���������� - ������� (��� => 206)
		'viborka_208_MariEl_full -> 208 => ���������� ����� �� (��� => 208)
		'viborka_202_Saratov_obl_full -> 202 => ����������� ������� (��� => 202)
		'viborka_301_Rostov_obl_full -> 301 => ���������� ������� (��� => 301)
		'viborka_203_Ulynov_obl_full -> 203 => ����������� ������� (��� => 203)
		'viborka_401_Murmansk_obl_full -> 401 => ���������� ������� (��� => 401)
		'viborka_402_Kareliy_full -> 402 => ���������� ������� (��� => 402)
		'viborka_405_Pskovskay_obl_full -> 405 => ��������� ������� (��� => 405)
		'viborka_407_Kaliningrad_obl_full -> 407 => ��������������� ������� (��� => 407)
		'viborka_805_Estony_full -> 805 => ��������� ���������� (��� => 805)
		'viborka_806_Latviy_full -> 806 => ���������� ���������� (��� => 806)
		'viborka_807_Litva_full -> 807 => ��������� ���������� (��� => 807)
		'viborka_801_Finskay_full -> 801 => ����������� ���������� (��� => 801)
		'viborka_823_Donbas_full -> 823 => ���������� ������ (��� => 823)
		'viborka_825_Orenburg_obl_full -> 825 => ������������ ������� (��� => 825 (����_max - 831))
		'viborka_204_Penza_obl_full -> 204 => ���������� �������
		'viborka_207_Republic_Mordova_full -> 207 => ���������� ��������
		'viborka_209_Nijegor_obl_full -> 209 => ������������� �������
		'viborka_311_Vologda_obl_full -> 311 => ������������� �������
		'viborka_404_Nigegor_obl_full -> 404 => ������������ �������
		'viborka_803_Zapad_reg_full -> 803 => �������� ������
		'viborka_804_Belorus_full -> 804 => �������� ������
		'viborka_819_Shvec_full -> 819 => ������
		'viborka_820_SPB_full -> 820 => �����-���������
		'viborka_822_Leningral_obl_full -> 822 => ������������� �������
		'viborka_823_Ugo_Zapad_reg_full -> 823 => ���-�������� ������
		'viborka_825_Ugny_reg_full -> 825 => ����� ������ 
		'viborka_826_Dnepov_reg_full -> 826 => ����������� ������
		'viborka_827_Sever_reg_full ->  827 => �������� ������
		' ############################################################################
		
		viborka = viborka_ot_100_do_200_full & Test_Area(na_Samarskay_obl) & Test_Area(na_Tatrskay) & Test_Area(na_Chuvashy) & Test_Area(na_MariEl) & Test_Area(na_Saratov_obl) & Test_Area(na_Rostov_obl) & Test_Area(na_Ulynov_obl) & Test_Area(na_Murmansk_obl) & Test_Area(na_Kareliy) & Test_Area(na_Pskovskay_obl) & Test_Area(na_Kaliningrad_obl) & Test_Area(na_Estony) & Test_Area(na_Latviy) & Test_Area(na_Litva) & Test_Area(na_Finskay) & Test_Area(na_Donbas) & Test_Area(na_Orenburg_obl) & Test_Area(na_Penza_obl) & Test_Area(na_Republic_Mordova) & Test_Area(na_Nijegor_obl) & Test_Area(na_Vologda_obl) & Test_Area(na_Nigegor_obl) & Test_Area(na_Zapad_reg) & Test_Area(na_Belorus) & Test_Area(na_Shvec) & Test_Area(na_SPB) & Test_Area(na_Leningral_obl) & Test_Area(na_Ugo_Zapad_reg) & Test_Area(na_Ugny_reg) & Test_Area(na_Dnepov_reg) & Test_Area(na_Sever_reg)
		
		print(viborka)
		
		'Call Vikluchatel(viborka)
		Call Obnulenie()
		Call Control_Rgm()
		Call print_parm_shem()
	else
		print("--- ������: ����� ����������! ---")
		print("--- ������ ������� ���������  !!! ������ !!! ---")
	end if
End Sub

Sub Ekv_Urala_do_500kV(viborka_ot_100_do_200)
	Set Vetv = t.Tables("vetv")
    Set Node = t.Tables("node")
    Set Generator = t.Tables("Generator")
	
	Node.SetSel(viborka_ot_100_do_200) ' ������� �� �����
    Node.Cols("sel").Calc("1")
    j = Node.FindNextSel(-1)
	
	While j<>(-1)
        ny = Node.Cols("ny").Z(j)
        tip_node = Node.Cols("tip").Z(j)
        uhom = Node.Cols("uhom").Z(j)
        If tip_node > 1 Then ' ��� ������������ ����
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
	print("-> ��������� ��������� ������(-��): " & viborka_ot_100_do_200 )
End Sub

Sub Ekvivalent_Node_Gen(vyborka_gen)
    Set Vetv = t.Tables("vetv")
    Set Node = t.Tables("node")
    Set Generator = t.Tables("Generator")
 
	Node.SetSel(vyborka_gen) ' ������� �� �����
    Node.Cols("sel").Calc("1")
    j = Node.FindNextSel(-1) 
	
    While j <> (-1)
        ny = Node.Cols("ny").Z(j)
        tip_node = Node.Cols("tip").Z(j)
        uhom = Node.Cols("uhom").Z(j)
        If tip_node > 1 Then ' ��� ������������ ����
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
	print("-> ��������� ��������� ������(-��): " & vyborka_gen )
End Sub

Sub print(str)
    t.Printp(str)
End Sub

Sub Rastr_Ekv()
    Time_Rastr_Ekv_1 = Timer()
        t.Ekv("")
    Time_Rastr_Ekv_2 = Timer()
	print(" - ������: ������������������! TIMER: " & (Time_Rastr_Ekv_2 - Time_Rastr_Ekv_1) & " [�e����] (" & (Time_Rastr_Ekv_2 - Time_Rastr_Ekv_1)/60 & " [�����])")
End Sub

Sub Control_Rgm()
	kod = t.rgm("p")
	if kod<>(-1) then
		print(" - ����� �������������!")
	else
		print(" - ����� �� �������������!")
	end if
End Sub

Sub Vikluchatel(viborka_ray_vikl)
    Time_Vikluchatel_1 = Timer()
    Set vet=t.tables("vetv")
    Set uzl=t.tables("node")
    Set gen=t.tables("Generator")

    Dim nodes(30000)
	
	uzl.SetSel(viborka_ray_vikl) ' ������� ����� ���� ������� ����� 500 (������)
    uzl.cols("sel").calc(1) ' ��������� �������� �����
    vet.SetSel("iq.sel=1 & ip.sel=0 &!sta") ' ������� ������ iq.sel = 1 ...
    k = vet.FindNextSel(-1)
	While k<>(-1) ' ������� sel-���� ���� �� �� � ����� ������� ������� ���� 
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
	
    While k<>(-1) ' ������� sel-���� ���� �� �� � ����� ������� ������� ���� 
		ip1 = vet.Cols("ip").z(k)
		uzl.Setsel "ny=" & ip1
		k2 = uzl.FindNextSel(-1)
		If k2<>(-1) Then
			uzl.cols("sel").z(k2) = 0
		End If
		k = vet.FindNextSel(k)
	Wend
  
	vet.SetSel("(iq.sel=1 & ip.sel=0)|(ip.sel=1 & iq.sel=0) & tip=2") ' tip=2 - ����������� (������� ���� ������������ ���� ������ � ����� ������� ������� ���� sel)
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
	' �������� ������������
	for povet = 0 to 10000
		vet.SetSel("x<0.01 & x>-0.01 & r<0.005 & r>=0 & (ktr=0 | ktr=1) & !sta & groupid!=1 & b<0.000005")  '������� ������, ������� ������� �������������
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
			'�������� �� ������� ���� �� ������ �����������
            for inodee = 0 to nnod
                If 	ndel = nodes(inodee) Then ndndel = 1
                If 	ny = nodes(inodee) Then ndny = 1
                If (ndndel = 1) and (ndny = 1) Then exit for
            next
			' ������ �������, ��� ��� ��������� ������ �������, � ����������� ����� ))
            If (ndndel = 0) and (ndny = 1) Then
                buff = ny
                ny = ndel
                ndel = buff
            End If
			
            If (ndndel = 0) or (ndny = 0) Then '���� ���� �� ���� ����� �������
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
				igen = gen.FindNextSel(-1) '������ ���� ����������� �����������
				 
				If igen<>(-1) Then
					While igen<>(-1) 
						gen.cols("Node").z(igen) = ny
						igen = gen.FindNextSel(igen)
					Wend
				End If
				
				If (v1<>v2) and (v1>0.3) and (v2>0.3) and (qmax1 + qmax2) <> 0 Then
					uzl.cols("vzd").z(iny) = (v1*qmax1+v2*qmax2)/(qmax1+qmax2) '������ ���������������� �� qmax ����������
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
				vet.delrows '������� �����	
				vet.SetSel("iq=" & ndel) '������ ���� ������ � ��������� �����)))
				vet.cols("iq").calc(ny)	
				vet.SetSel("ip=" & ndel)
				vet.cols("ip").calc(ny)	
				uzl.delrows 		' ������� ����
          Else '���� �� ������ ������ �������
                vet.SetSel("(ip=" & ip & "& iq=" & iq & ")|(iq=" & ip & "& ip=" & iq & ")")
                vet.cols("groupid").calc(1)
        End If
    next
    kod = t.rgm ("p")
    If kod<>0 Then
        msgbox "Regim do not exist"		
    End If
    Time_Vikluchatel_2 = Timer()
    print(" @TIMER - ����� ������ ������� �������� ������������, � ������� = " & ((Time_Vikluchatel_2 - Time_Vikluchatel_1)/60))
End Sub

Sub Obnulenie()  ' ��������� ���� sel (���������� ��������) ����� � ������
    Set node = t.Tables("node")
    Set vetv = t.Tables("vetv")
    
    vetv.SetSel("")
	vetv.cols("sel").calc("0")
	node.SetSel("")
	node.cols("sel").calc("0")
	print(" - ����� '�������' � ���������� ����� � ������.")
End Sub

Sub Delete_Gen_Vikl()
    Time_VikluchatelGEN_1 = Timer()
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
    Time_VikluchatelGEN_2 = Timer()
    print(" @TIMER - ����� ������ ������� �������� ������������ �����������, � ������� = " & ((Time_VikluchatelGEN_2 - Time_VikluchatelGEN_1)/60))
End Sub

Sub Settings_Ekv()
	' ���������� ��������� ������������������
    print(" - ���������� ��������� ���. �����;")
	t.Tables("com_ekviv").Cols("zmax").z(0) = 1000
	t.Tables("com_ekviv").Cols("ek_sh").z(0) = 0
	t.Tables("com_ekviv").Cols("otm_n").z(0) = 0
	t.Tables("com_ekviv").Cols("smart").z(0) = 0
	t.Tables("com_ekviv").Cols("tip_ekv").z(0) = 0
	t.Tables("com_ekviv").Cols("ekvgen").z(0) = 0
	t.Tables("com_ekviv").Cols("tip_gen").z(0) = 1
End Sub

Sub Off_Gen_if_off_node()
	Time_Off_Gen_if_off_node_1 = Timer()
    
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
	print(" - ��������� ���������� � ����������� �����.")
    Time_Off_Gen_if_off_node_2 = Timer()
    print(" @TIMER - ����� ������ �������: ���������� ����������� ���� ����. ������������ ����, � ������� = " & ((Time_Off_Gen_if_off_node_2 - Time_Off_Gen_if_off_node_1)/60))
End Sub

Sub Off_line_one_on()
    Time_Off_line_one_on_1 = Timer()
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
	print(" - ���������� ��� � ������������� ���., ������������ � ��������� ������� ���������: " & ii)
    Time_Off_line_one_on_2 = Timer()
    print(" @TIMER - ����� ������ �������: ���������� ��� � ���� ������ ���� ��� ����. � ����� �������, � ������� = " & ((Time_Off_line_one_on_2 - Time_Off_line_one_on_1)/60))
End Sub

Function na_of_the_area_by_name(name_area)
    set area=t.tables("area")
    max_count_area = area.Count-1
    for i=0 to max_count_area 
        name_ = area.Cols("name").Z(i)
        if name_ = name_area then
            na_of_the_area_by_name = area.Cols("na").Z(i)
			print(" - �������� ������: "& name_ &"; ����� ������ => "& na_of_the_area_by_name)
		else
			na_of_the_area_by_name = ""
		end if 
    next
End function

Sub print_parm_shem()
    Set com_cxema = t.Tables("com_cxema")
    print("---------------------------------------------")
    print("          ����� ���������� � �����          ")
    print(" - �����: " & com_cxema.Cols("ny").Z(0))
    print(" - ������: " & com_cxema.Cols("nv").Z(0))
    print(" - �������: " & com_cxema.Cols("na").Z(0))
	print(" - ����� ����������� ����: " & com_cxema.Cols("ny_o").Z(0))
	print(" - ����� ����������� ������: " & com_cxema.Cols("nv_o").Z(0))
	print(" - ����� ������������� �����: " & com_cxema.Cols("nby").Z(0))
	print(" - ����� ����� � �������� V: " & com_cxema.Cols("ngen").Z(0))
	print(" - ����� ���������������: " & com_cxema.Cols("ntran").Z(0))
	print(" - ����� ���: " & com_cxema.Cols("nlep").Z(0))
	print(" - ����� ������������: " & com_cxema.Cols("nvikl").Z(0))
	print(" - �_���: " & com_cxema.Cols("pg").Z(0))
	print(" - �_���: " & com_cxema.Cols("pn").Z(0))
	print(" - ������ � (����������): " & com_cxema.Cols("dp").Z(0))
	print(" - �_������. �����: " & com_cxema.Cols("pby").Z(0))
	print(" - ���������� ������: " & com_cxema.Cols("dpsh").Z(0))
	print(" - ����������� ���������� V(%): " & com_cxema.Cols("dv_min").Z(0))
	print(" - ������������ ���������� V(%): " & com_cxema.Cols("dv_max").Z(0))
    print("---------------------------------------------")
End Sub

Sub Delete_Node_not_connect()
    Time_Delete_Node_not_connect_1 = Timer()
    
	Set gen = t.Tables("Generator")
    Set nodeg = gen.Cols("Node")
	Set uzl = t.Tables("node")
	Set vet = t.Tables("vetv")
	
	NodeColMax = uzl.Count-1
	VetvColMax = vet.Count-1
	ii = 0
	uzl.SetSel("sta=1")
    i = uzl.FindNextSel(-1)
    while i<>(-1)
        Bsh = uzl.Cols("bsh").Z(i)
		id_ny = uzl.Cols("ny").Z(i)
		vet.SetSel("ip.ny=" & id_ny & "| iq.ny=" & id_ny)
		ColVetv = vet.FindNextSel(-1)
		key_1 = 1
        
		If key_1=1 Then
			uzl.Cols("sel").Z(i) = 0
			If ColVetv=(-1) Then 
				uzl.Cols("sel").Z(i) = 1
				ii = ii + 1
			End If
		End If
        
		If key_1=0 Then
			vet.Cols("sel").Z(i) = 0
			If ColVetv<>(-1) Then
				TypeId = vet.Cols("tip").Z(ColVetv)    
				If TypeId=2 Then
				   If Bsh=0 Then
						vet.Cols("sel").Z(ColVetv) = 1
				   End If
				End If
			 End If
        End If
        i = uzl.FindNextSel(i)
    wend
    uzl.SetSel("sel=1")
	ii = uzl.Count-1
	uzl.DelRows
	print(" - ������� ����� ��� ������ � �������: " & ii+1)
	Call Control_Rgm()
	Call Obnulenie()
    Time_Delete_Node_not_connect_2 = Timer()
    print(" @TIMER - ����� ������ ������� �������� ����� ��� ����� � �������, � ������� = " & ((Time_Delete_Node_not_connect_2 - Time_Delete_Node_not_connect_1)/60))
End Sub

Sub Delete_Generators_without_nodes()
	Time_Delete_Generators_without_nodes_1 = Timer()
    Set gen = t.Tables("Generator")
	Set node = t.Tables("node")
	
	gen.SetSel("Node.ny=0")
	gen.DelRows
	gen.SetSel("")
	Call Obnulenie()
	Call Control_Rgm()
    Time_Delete_Generators_without_nodes_2 = Timer()
    print(" @TIMER - ����� ������ ������� �������� ����������� ��� �����, � ������� = " & ((Time_Delete_Generators_without_nodes_2 - Time_Delete_Generators_without_nodes_1)/60))
End Sub

Sub ReactorsChange()
    Time_ReactorsChange_1 = Timer()
	
    Set uzl=t.Tables("node")
	Set Reactors=t.Tables("Reactors")

	Reactors.SetSel("")
	Reactors.Cols("sel").Calc(0)
	Reactors.SetSel("")
	
	k=Reactors.FindNextSel(-1)
	while k<>(-1)
		ip1=Reactors.Cols("Id1").z(k)
		B1=Reactors.Cols("B").z(k)
		reac_sta=Reactors.Cols("sta").z(k)
		uzl.SetSel("ny=" & ip1  )
		if uzl.count > 0 then
			k2=uzl.FindNextSel(-1)
			while k2<>(-1)
				uzl.Cols("bsh").z(k2) = uzl.Cols("bsh").z(k2) + B1
				if reac_sta = 1 then
					uzl.Cols("sel").z(k2) = 1
				end if
				k2=uzl.FindNextSel(k2)
			wend
		end if
		k=Reactors.FindNextSel(k)
	wend

	Reactors.SetSel("")
	Reactors.Delrows
    
    Time_ReactorsChange_2 = Timer()
    print(" @TIMER - ����� ������ ������� ������� ��������� � ����, � ������� = " & ((Time_ReactorsChange_2 - Time_ReactorsChange_1)/60))
End Sub

Function Test_Area(str)
	if str <> "" then
		Test_Area = "| na=" & str
	else
		Test_Area = ""
	end if
End Function