r=Setlocale("en-us")
rrr=1

Set t = Rastr

Call Populating_Dynamic_Set()






Function Populating_Dynamic_Set()
	t.Printp("������ ������� ���������� � ���_������ - Populating_Dynamic_Set")
    Set vet=t.tables("vetv")
		Set ray=t.tables("area")
		Set area=t.Tables("area")
		Set area2=t.Tables("area2")
		Set darea=t.Tables("darea")
		Set polin=t.Tables("polin")
		Set Reactors=t.Tables("Reactors")
		Set pqd=t.Tables("graphik2")
		Set graphikIT=t.Tables("graphikIT")
		Set uzl=t.tables("node")
		Set gen=t.Tables("Generator")
		Set vozb=t.Tables("Exciter")
		Set arv=t.Tables("ExcControl")
		Set vieee=t.Tables("DFWIEEE421")
		Set stieee=t.Tables("DFWIEEE421PSS13")
		Set pss4=t.Tables("DFWIEEE421PSS4B")
		Set ars=Rastr.Tables("ARS")
		Set forc=Rastr.Tables("Forcer")
		Set omv=t.Tables("DFW421UEL")
		Set bor=t.Tables("DFWOELUNITROL")
		Set FuncPQ=t.Tables("FuncPQ")
		Set Governor=t.Tables("Governor")
		Set decs400=t.Tables("DFWDECS400")
		Set Thyne=t.Tables("DFWTHYNE14")

		Set nsta=gen.Cols("sta")
		Set numg=gen.Cols("Num")
		Set nameg=gen.Cols("Name")
		Set gNumBrand=gen.Cols("NumBrand")
		Set nodeg=gen.Cols("Node")
		Set ExciterId=gen.Cols("ExciterId")
		Set ARSId=gen.Cols("ARSId")
		Set pgen=gen.Cols("P")
		Set ModelType=gen.Cols("ModelType")
		Set pnom=gen.Cols("Pnom")
		Set unom=gen.Cols("Ugnom")
		Set cosfi=gen.Cols("cosFi")
		Set Demp=gen.Cols("Demp")
		Set mj=gen.Cols("Mj")
		Set xd1=gen.Cols("xd1")
		Set qmin=gen.Cols("Qmin")
		Set qmax=gen.Cols("Qmax")
		Set xd1	=gen.Cols("xd1")
		Set xd=gen.Cols("xd")
		Set xq	=gen.Cols("xq")
		Set xd2	=gen.Cols("xd2")
		Set xq2	=gen.Cols("xq2")
		Set td01=gen.Cols("td01")
		Set td02=gen.Cols("td02")
		Set tq02=gen.Cols("tq02")
		Set xq1	=gen.Cols("xq1")
		Set xl	=gen.Cols("xl")
		Set x2	=gen.Cols("x2")
		Set x2	=gen.Cols("x0")
		Set tq01=gen.Cols("tq01")
		Set gIVActuatorId=gen.Cols("IVActuatorId")

		Set ny=uzl.Cols("ny")
		Set name=uzl.Cols("name")
		Set pg		=uzl.Cols("pg") 
		Set tip		=uzl.Cols("tip") 
		Set nqmin	=uzl.Cols("qmin")
		Set nqmax	=uzl.Cols("qmax")
		Set uhom	=uzl.Cols("uhom")

		Set idv	=vozb.Cols("Id")
		Set namev	=vozb.Cols("Name")
		Set ModelTypev	=vozb.Cols("ModelType")
		Set Brandv	=vozb.Cols("Brand")
		Set ExcControlIdv	=vozb.Cols("ExcControlId")
		Set ForcerIdv	=vozb.Cols("ForcerId")
		Set Texcv	=vozb.Cols("Texc")
		Set Kig=vozb.Cols("Kig")
		Set KIf=vozb.Cols("Kif")
		Set Uf_min=vozb.Cols("Uf_min")
		Set Uf_max=vozb.Cols("Uf_max")
		Set If_min=vozb.Cols("If_min")
		Set If_max=vozb.Cols("If_max")
		Set Type_rgv=vozb.Cols("Type_rg")
		Set vozbCustomModel=vozb.Cols("CustomModel")
		Set vozbKarv=vozb.Cols("Karv")
		Set vozbT3exc=vozb.Cols("T3exc")

		Set ida	=arv.Cols("Id")
		Set namea	=arv.Cols("Name")
		Set ModelTypea	=arv.Cols("ModelType")
		Set Branda	=arv.Cols("Brand")
		Set Trv	=arv.Cols("Trv")
		Set Ku	=arv.Cols("Ku")
		Set Ku1	=arv.Cols("Ku1")
		Set KIf1	=arv.Cols("Kif1")
		Set Kf	=arv.Cols("Kf")
		Set Kf1	=arv.Cols("Kf1")
		Set Tf	=arv.Cols("Tf")
		Set Urv_min	=arv.Cols("Urv_min")
		Set Urv_max	=arv.Cols("Urv_max")
		Set Alpha	=arv.Cols("Alpha")
		Set arvCustomModel	=arv.Cols("CustomModel")
		Set arvTINT	=arv.Cols("TINT")

		Set arsids	=ars.Cols("Id")
		Set arsname	=ars.Cols("Name")
		Set arsModelTypes	=ars.Cols("ModelType")
		Set arsCustomModel	=ars.Cols("ModelType")
		Set arsBrands	=ars.Cols("Brand")
		Set arsGovernorId	=ars.Cols("GovernorId")
		Set arsao	=ars.Cols("ao")
		Set arsaz	=ars.Cols("az")
		Set arsotmin	=ars.Cols("otmin")
		Set arsotmax	=ars.Cols("otmax")
		Set arsstrs	=ars.Cols("strs")
		Set arszn	=ars.Cols("zn")
		Set arsdpo	=ars.Cols("dpo")
		Set arsThp	=ars.Cols("Thp")
		Set arsTrlp	=ars.Cols("Trlp")
		Set arsTw	=ars.Cols("Tw")
		Set arspt	=ars.Cols("pt")
		Set arsMu	=ars.Cols("Mu")
		Set arsPsteam	=ars.Cols("Psteam")
		Set arsCustomModel	=ars.Cols("CustomModel")

		Set idf	=forc.Cols("Id")
		Set namef	=forc.Cols("Name")
		Set ModelTypef	=forc.Cols("ModelType")
		Set Ubf	=forc.Cols("Ubf")
		Set Uef	=forc.Cols("Uef")
		Set Rf	=forc.Cols("Rf")
		Set Texc_f	=forc.Cols("Texc_f")
		Set Tz_in	=forc.Cols("Tz_in")
		Set Tz_out	=forc.Cols("Tz_out")
		Set Ubrf	=forc.Cols("Ubrf")
		Set Uerf	=forc.Cols("Uerf")
		Set Rrf	=forc.Cols("Rrf")
		Set Texc_rf	=forc.Cols("Texc_rf")

		Set ststa=stieee.Cols("sta")
		Set stId=stieee.Cols("Id")
		Set stName=stieee.Cols("Name")
		Set stModel=stieee.Cols("ModelType")
		Set stBrand=stieee.Cols("Brand")
		Set stCustomModel=stieee.Cols("CustomModel")
		Set stInput1Type=stieee.Cols("Input1Type")
		Set stInput2Type=stieee.Cols("Input2Type")
		Set stVstmin=stieee.Cols("Vstmin")
		Set stVstmax=stieee.Cols("Vstmax")
		Set stKs1=stieee.Cols("Ks1")
		Set stT1=stieee.Cols("T1")
		Set stT2=stieee.Cols("T2")
		Set stT3=stieee.Cols("T3")
		Set stT4=stieee.Cols("T4")
		Set stT5=stieee.Cols("T5")
		Set stT6=stieee.Cols("T6")
		Set stT7=stieee.Cols("T7")
		Set stT8=stieee.Cols("T8")
		Set stT9=stieee.Cols("T9")
		Set stT10=stieee.Cols("T10")
		Set stT11=stieee.Cols("T11")
		Set stA1=stieee.Cols("A1")
		Set stA2=stieee.Cols("A2")
		Set stA3=stieee.Cols("A3")
		Set stA4=stieee.Cols("A4")
		Set stA5=stieee.Cols("A5")
		Set stA6=stieee.Cols("A6")
		Set stA7=stieee.Cols("A7")
		Set stA8=stieee.Cols("A8")
		Set stKs2=stieee.Cols("Ks2")
		Set stKs3=stieee.Cols("Ks3")
		Set stTw1=stieee.Cols("Tw1")
		Set stTw2=stieee.Cols("Tw2")
		Set stTw3=stieee.Cols("Tw3")
		Set stTw4=stieee.Cols("Tw4")
		Set stM=stieee.Cols("M")
		Set stN=stieee.Cols("N")
		Set stVsi1min=stieee.Cols("Vsi1min")
		Set stVsi1max=stieee.Cols("Vsi1max")
		Set stVsi2min=stieee.Cols("Vsi2min")
		Set stVsi2max=stieee.Cols("Vsi2max")

		Set vista=vieee.Cols("sta")
		Set viId=vieee.Cols("Id")
		Set viName=vieee.Cols("Name")
		Set viModel=vieee.Cols("ModelType")
		Set viBrand=vieee.Cols("Brand")
		Set viCustomModel=vieee.Cols("CustomModel")
		Set viUELId=vieee.Cols("UELId")
		Set viUELPos=vieee.Cols("UELPos")
		Set viOELId=vieee.Cols("OELId")
		Set viOELPos=vieee.Cols("OELPos")
		Set viPSSId=vieee.Cols("PSSId")
		Set viPSSPos=vieee.Cols("PSSPos")
		Set viTe=vieee.Cols("Te")
		Set viKe=vieee.Cols("Ke")
		Set viSe1=vieee.Cols("Se1")
		Set viEfd1=vieee.Cols("Efd1")
		Set viVe1=vieee.Cols("Ve1")
		Set viSe2=vieee.Cols("Se2")
		Set viEfd2=vieee.Cols("Efd2")
		Set viVe2=vieee.Cols("Ve2")
		Set viVemin=vieee.Cols("Vemin")
		Set viVrmin=vieee.Cols("Vrmin")
		Set viVrmax=vieee.Cols("Vrmax")
		Set viKa=vieee.Cols("Ka")
		Set viTa=vieee.Cols("Ta")
		Set viTf=vieee.Cols("Tf")
		Set viKf=vieee.Cols("Kf")
		Set viTc=vieee.Cols("Tc")
		Set viTb=vieee.Cols("Tb")
		Set viKv=vieee.Cols("Kv")
		Set viTrh=vieee.Cols("Trh")
		Set viKpr=vieee.Cols("Kpr")
		Set viKir=vieee.Cols("Kir")
		Set viKdr=vieee.Cols("Kdr")
		Set viTdr=vieee.Cols("Tdr")
		Set viKc=vieee.Cols("Kc")
		Set viKd=vieee.Cols("Kd")
		Set viVfemax=vieee.Cols("Vfemax")
		Set viVamin=vieee.Cols("Vamin")
		Set viVamax=vieee.Cols("Vamax")
		Set viKb=vieee.Cols("Kb")
		Set viKh=vieee.Cols("Kh")
		Set viKr=vieee.Cols("Kr")
		Set viKn=vieee.Cols("Kn")
		Set viEfdn=vieee.Cols("Efdn")
		Set viKlv=vieee.Cols("Klv")
		Set viVlv=vieee.Cols("Vlv")
		Set viVimin=vieee.Cols("Vimin")
		Set viVimax=vieee.Cols("Vimax")
		Set viTf2=vieee.Cols("Tf2")
		Set viTf3=vieee.Cols("Tf3")
		Set viTk=vieee.Cols("Tk")
		Set viTj=vieee.Cols("Tj")
		Set viTh=vieee.Cols("Th")
		Set viVhmax=vieee.Cols("Vhmax")
		Set viVfelim=vieee.Cols("Vfelim")
		Set viKp=vieee.Cols("Kp")
		Set viKpa=vieee.Cols("Kpa")
		Set viKia=vieee.Cols("Kia")
		Set viKf1=vieee.Cols("Kf1")
		Set viKf2=vieee.Cols("Kf2")
		Set viKl=vieee.Cols("Kl")
		Set viTb1=vieee.Cols("Tb1")
		Set viTc1=vieee.Cols("Tc1")
		Set viKlr=vieee.Cols("Klr")
		Set viIlr=vieee.Cols("Ilr")
		Set viKi=vieee.Cols("Ki")
		Set viTheta=vieee.Cols("Theta")
		Set viVmmin=vieee.Cols("Vmmin")
		Set viVmmax=vieee.Cols("Vmmax")
		Set viKg=vieee.Cols("Kg")
		Set viVBmax=vieee.Cols("VBmax")
		Set viVGmax=vieee.Cols("VGmax")
		Set viXl=vieee.Cols("Xl")
		Set viKm=vieee.Cols("Km")
		Set viTm=vieee.Cols("Tm")
		Set viTb2=vieee.Cols("Tb2")
		Set viTc2=vieee.Cols("Tc2")
		Set viTub1=vieee.Cols("Tub1")
		Set viTuc1=vieee.Cols("Tuc1")
		Set viTub2=vieee.Cols("Tub2")
		Set viTuc2=vieee.Cols("Tuc2")
		Set viTob1=vieee.Cols("Tob1")
		Set viToc1=vieee.Cols("Toc1")
		Set viTob2=vieee.Cols("Tob2")
		Set viToc2=vieee.Cols("Toc2")

		Set viAex=vieee.Cols("Aex")
		Set viBex=vieee.Cols("Bex")
		Set viKcf=vieee.Cols("Kcf")
		Set viKhf=vieee.Cols("Khf")
		Set viKIf=vieee.Cols("Kif")
		Set viSamovozb=vieee.Cols("Samovozb")
		Set viTr=vieee.Cols("Tr")

		Set pss4sta=pss4.Cols("sta")
		Set pss4Id=pss4.Cols("Id")
		Set pss4Name=pss4.Cols("Name")
		Set pss4ModelType=pss4.Cols("ModelType")
		Set pss4Brand=pss4.Cols("Brand")
		Set pss4CustomModel=pss4.Cols("CustomModel")
		Set pss4Input1Type=pss4.Cols("Input1Type")
		Set pss4Input2Type=pss4.Cols("Input2Type")
		Set pss4MBPSS1=pss4.Cols("MBPSS1")
		Set pss4MBPSS2=pss4.Cols("MBPSS2")
		Set pss4Vstmin=pss4.Cols("Vstmin")
		Set pss4Vstmax=pss4.Cols("Vstmax")
		Set pss4KL1=pss4.Cols("KL1")
		Set pss4KL2=pss4.Cols("KL2")
		Set pss4KL11=pss4.Cols("KL11")
		Set pss4KL17=pss4.Cols("KL17")
		Set pss4TL1=pss4.Cols("TL1")
		Set pss4TL2=pss4.Cols("TL2")
		Set pss4TL3=pss4.Cols("TL3")
		Set pss4TL4=pss4.Cols("TL4")
		Set pss4TL5=pss4.Cols("TL5")
		Set pss4TL6=pss4.Cols("TL6")
		Set pss4TL7=pss4.Cols("TL7")
		Set pss4TL8=pss4.Cols("TL8")
		Set pss4TL9=pss4.Cols("TL9")
		Set pss4TL10=pss4.Cols("TL10")
		Set pss4TL11=pss4.Cols("TL11")
		Set pss4TL12=pss4.Cols("TL12")
		Set pss4KL=pss4.Cols("KL")
		Set pss4Vlmin=pss4.Cols("Vlmin")
		Set pss4Vlmax=pss4.Cols("Vlmax")
		Set pss4KI1=pss4.Cols("KI1")
		Set pss4KI2=pss4.Cols("KI2")
		Set pss4KI11=pss4.Cols("KI11")
		Set pss4KI17=pss4.Cols("KI17")
		Set pss4TI1=pss4.Cols("TI1")
		Set pss4TI2=pss4.Cols("TI2")
		Set pss4TI3=pss4.Cols("TI3")
		Set pss4TI4=pss4.Cols("TI4")
		Set pss4TI5=pss4.Cols("TI5")
		Set pss4TI6=pss4.Cols("TI6")
		Set pss4TI7=pss4.Cols("TI7")
		Set pss4TI8=pss4.Cols("TI8")
		Set pss4TI9=pss4.Cols("TI9")
		Set pss4TI10=pss4.Cols("TI10")
		Set pss4TI11=pss4.Cols("TI11")
		Set pss4TI12=pss4.Cols("TI12")
		Set pss4KI=pss4.Cols("KI")
		Set pss4Vimin=pss4.Cols("Vimin")
		Set pss4Vimax=pss4.Cols("Vimax")
		Set pss4KH1=pss4.Cols("KH1")
		Set pss4KH2=pss4.Cols("KH2")
		Set pss4KH11=pss4.Cols("KH11")
		Set pss4KH17=pss4.Cols("KH17")
		Set pss4TH1=pss4.Cols("TH1")
		Set pss4TH2=pss4.Cols("TH2")
		Set pss4TH3=pss4.Cols("TH3")
		Set pss4TH4=pss4.Cols("TH4")
		Set pss4TH5=pss4.Cols("TH5")
		Set pss4TH6=pss4.Cols("TH6")
		Set pss4TH7=pss4.Cols("TH7")
		Set pss4TH8=pss4.Cols("TH8")
		Set pss4TH9=pss4.Cols("TH9")
		Set pss4TH10=pss4.Cols("TH10")
		Set pss4TH11=pss4.Cols("TH11")
		Set pss4TH12=pss4.Cols("TH12")
		Set pss4KH=pss4.Cols("KH")
		Set pss4Vhmin=pss4.Cols("Vhmin")
		Set pss4Vhmax=pss4.Cols("Vhmax")
		Set pss4sta=pss4.Cols("sta")

		Set omvsta=omv.Cols("sta")
		Set omvId=omv.Cols("Id")
		Set omvName=omv.Cols("Name")
		Set omvModelType=omv.Cols("ModelType")
		Set omvBrand=omv.Cols("Brand")
		Set omvCustomModel=omv.Cols("CustomModel")
		Set omvTu1=omv.Cols("Tu1")
		Set omvTu2=omv.Cols("Tu2")
		Set omvTu3=omv.Cols("Tu3")
		Set omvTu4=omv.Cols("Tu4")
		Set omvVulmin=omv.Cols("Vulmin")
		Set omvVulmax=omv.Cols("Vulmax")
		Set omvKul=omv.Cols("Kul")
		Set omvKui=omv.Cols("Kui")
		Set omvVuimin=omv.Cols("Vuimin")
		Set omvVuimax=omv.Cols("Vuimax")
		Set omvKuf=omv.Cols("Kuf")
		Set omvTuf=omv.Cols("Tuf")
		Set omvKur=omv.Cols("Kur")
		Set omvKuc=omv.Cols("Kuc")
		Set omvVurmax=omv.Cols("Vurmax")
		Set omvVucmax=omv.Cols("Vucmax")
		Set omvTuV=omv.Cols("TuV")
		Set omvTuP=omv.Cols("TuP")
		Set omvTuQ=omv.Cols("TuQ")
		Set omvK1=omv.Cols("K1")
		Set omvK2=omv.Cols("K2")
		'Set omvDepEndency_F1=omv.Cols("DepEndency_F1")
		'Set omvOutput=omv.Cols("Output")
		'Set omvKl=omv.Cols("Kl")

		Set borsta=bor.Cols("sta")
		Set borId=bor.Cols("Id")
		Set borName=bor.Cols("Name")
		Set borModelType=bor.Cols("ModelType")
		Set borBrand=bor.Cols("Brand")
		Set borCustomModel=bor.Cols("CustomModel")
		Set borIfMax=bor.Cols("IfMax")
		Set borIfth=bor.Cols("Ifth")
		Set borKexpIf=bor.Cols("KexpIf")
		Set borKR3=bor.Cols("KR3")
		Set borKR3i=bor.Cols("KR3i")
		Set borTc23=bor.Cols("Tc23")
		Set borTb23=bor.Cols("Tb23")
		Set borTc13=bor.Cols("Tc13")
		Set borTb13=bor.Cols("Tb13")
		Set borVamin=bor.Cols("Vamin")
		Set borVamax=bor.Cols("Vamax")
		Set borTdel=bor.Cols("Tdel")
		Set borKth=bor.Cols("Kth")
		Set borKToF=bor.Cols("KToF")
		Set borKcF=bor.Cols("KcF")
		Set borKhF=bor.Cols("KhF")
		Set borTRFout=bor.Cols("TRFout")
		Set borTr=bor.Cols("Tr")
		Set borOutput=bor.Cols("Output")
		Set borKl=bor.Cols("Kl")

		Set FuncPQId=FuncPQ.Cols("Id")
		Set FuncPQP=FuncPQ.Cols("P")
		Set FuncPQQ=FuncPQ.Cols("Q")

		Set Governorsta=Governor.Cols("sta")
		Set GovernorId=Governor.Cols("Id")
		Set GovernorName=Governor.Cols("Name")
		Set GovernorModelType=Governor.Cols("ModelType")
		Set GovernorBrand=Governor.Cols("Brand")
		Set Governorstrs=Governor.Cols("strs")
		Set Governorzn=Governor.Cols("zn")
		Set GovernorTr=Governor.Cols("Tr")
		Set Governorotmin=Governor.Cols("otmin")
		Set Governorotmax=Governor.Cols("otmax")
		Set GovernorCVmin=Governor.Cols("CVmin")
		Set GovernorCVmax=Governor.Cols("CVmax")
		Set GovernorT1=Governor.Cols("T1")
		Set GovernorK1=Governor.Cols("K1")
		Set GovernorK2=Governor.Cols("K2")
		Set GovernorBoilerId=Governor.Cols("BoilerId")

		Set decs400sta=decs400.Cols("sta")
		Set decs400Id=decs400.Cols("Id")
		Set decs400Name=decs400.Cols("Name")
		Set decs400ModelType=decs400.Cols("ModelType")
		Set decs400Brand=decs400.Cols("Brand")
		Set decs400CustomModel=decs400.Cols("CustomModel")
		Set decs400PSSId=decs400.Cols("PSSId")
		Set decs400UELId=decs400.Cols("UELId")
		Set decs400OELId=decs400.Cols("OELId")
		Set decs400Xl=decs400.Cols("Xl")
		Set decs400DRP=decs400.Cols("DRP")
		Set decs400VrMin=decs400.Cols("VrMin")
		Set decs400VrMax=decs400.Cols("VrMax")
		Set decs400VmMin=decs400.Cols("VmMin")
		Set decs400VmMax=decs400.Cols("VmMax")
		Set decs400VbMax=decs400.Cols("VbMax")
		Set decs400Kc=decs400.Cols("Kc")
		Set decs400Kp=decs400.Cols("Kp")
		Set decs400Kpm=decs400.Cols("Kpm")
		Set decs400Kpr=decs400.Cols("Kpr")
		Set decs400Kir=decs400.Cols("Kir")
		Set decs400Kpd=decs400.Cols("Kpd")
		Set decs400Ta=decs400.Cols("Ta")
		Set decs400Td=decs400.Cols("Td")
		Set decs400Tr=decs400.Cols("Tr")
		Set decs400SelfExc=decs400.Cols("SelfExc")
		Set decs400Del=decs400.Cols("Del")

		Set Thynesta=Thyne.Cols("sta")
		Set ThyneId=Thyne.Cols("Id")
		Set ThyneName=Thyne.Cols("Name")
		Set ThyneModelType=Thyne.Cols("ModelType")
		Set ThyneBrand=Thyne.Cols("Brand")
		Set ThyneCustomModel=Thyne.Cols("CustomModel")
		Set ThyneUELId=Thyne.Cols("UELId")
		Set ThynePSSId=Thyne.Cols("PSSId")
		Set ThyneAex=Thyne.Cols("Aex")
		Set ThyneBex=Thyne.Cols("Bex")
		Set ThyneAlpha=Thyne.Cols("Alpha")
		Set ThyneBeta=Thyne.Cols("Beta")
		Set ThyneIfdMin=Thyne.Cols("IfdMin")
		Set ThyneKc=Thyne.Cols("Kc")
		Set ThyneKd1=Thyne.Cols("Kd1")
		Set ThyneKd2=Thyne.Cols("Kd2")
		Set ThyneKe=Thyne.Cols("Ke")
		Set ThyneKetb=Thyne.Cols("Ketb")
		Set ThyneKh=Thyne.Cols("Kh")
		Set ThyneKp1=Thyne.Cols("Kp1")
		Set ThyneKp2=Thyne.Cols("Kp2")
		Set ThyneKp3=Thyne.Cols("Kp3")
		Set ThyneTd1=Thyne.Cols("Td1")
		Set ThyneTe1=Thyne.Cols("Te1")
		Set ThyneTe2=Thyne.Cols("Te2")
		Set ThyneTi1=Thyne.Cols("Ti1")
		Set ThyneTi2=Thyne.Cols("Ti2")
		Set ThyneTi3=Thyne.Cols("Ti3")
		Set ThyneTr1=Thyne.Cols("Tr1")
		Set ThyneTr2=Thyne.Cols("Tr2")
		Set ThyneTr3=Thyne.Cols("Tr3")
		Set ThyneTr4=Thyne.Cols("Tr4")
		Set ThyneVO1Min=Thyne.Cols("VO1Min")
		Set ThyneVO1Max=Thyne.Cols("VO1Max")
		Set ThyneVO2Min=Thyne.Cols("VO2Min")
		Set ThyneVO2Max=Thyne.Cols("VO2Max")
		Set ThyneVO3Min=Thyne.Cols("VO3Min")
		Set ThyneVO3Max=Thyne.Cols("VO3Max")
		Set ThyneVD1Min=Thyne.Cols("VD1Min")
		Set ThyneVD1Max=Thyne.Cols("VD1Max")
		Set ThyneVI1Min=Thyne.Cols("VI1Min")
		Set ThyneVI1Max=Thyne.Cols("VI1Max")
		Set ThyneVI2Min=Thyne.Cols("VI2Min")
		Set ThyneVI2Max=Thyne.Cols("VI2Max")
		Set ThyneVI3Min=Thyne.Cols("VI3Min")
		Set ThyneVI3Max=Thyne.Cols("VI3Max")
		Set ThyneVP1Min=Thyne.Cols("VP1Min")
		Set ThyneVP1Max=Thyne.Cols("VP1Max")
		Set ThyneVP2Min=Thyne.Cols("VP2Min")
		Set ThyneVP2Max=Thyne.Cols("VP2Max")
		Set ThyneVP3Min=Thyne.Cols("VP3Min")
		Set ThyneVP3Max=Thyne.Cols("VP3Max")
		Set ThyneVrMin=Thyne.Cols("VrMin")
		Set ThyneVrMax=Thyne.Cols("VrMax")
		Set ThyneXp=Thyne.Cols("Xp")

	'---------------------------------------------------------------------------------------------------------
	'������ ������ ��� ���������������� ���������

    'LinkCustomModels = LinkCustomModels
	LinkCustomModels = "C:\CustomModels\"
	SettingsFile ="L:\SecretDisk\SER\������������ ������\2021 ��� ����\���_����� (����������) ���.xlsm"
    'SettingsFile = FileExcelDynamicSet
	
	Set ExcElSet = CreateObject("Excel.Application")	
		ExcElSet.Workbooks.open SettingsFile
		ExcElSet.Visible = True
		ExcElSet.EnableEvents = True
		ExcElSet.ScreenUpdating = True
		ExcElSet.DisplayAlerts = True

	Set Ex1=ExcElSet.Worksheets("1")
		Set Ex2=ExcElSet.Worksheets("2")
		Set Ex3=ExcElSet.Worksheets("3")
		Set Ex4=ExcElSet.Worksheets("4")
		Set Ex5=ExcElSet.Worksheets("5")
		Set Ex6=ExcElSet.Worksheets("6")
		Set Ex7=ExcElSet.Worksheets("7")
		Set Ex8=ExcElSet.Worksheets("8")
		Set Ex9=ExcElSet.Worksheets("9")
		Set Ex10=ExcElSet.Worksheets("10")
		Set Ex11=ExcElSet.Worksheets("11")
		Set Ex12=ExcElSet.Worksheets("12")
		Set Ex13=ExcElSet.Worksheets("13")
		Set Ex14=ExcElSet.Worksheets("14")


    
    arv.delrows
	vozb.delrows
	vieee.delrows
	ars.delrows
	forc.delrows
	pss4.delrows
	stieee.delrows
	omv.delrows
	bor.delrows
	FuncPQ.delrows
	Governor.delrows
	decs400.delrows
	Thyne.delrows

	ffff = 8999
	If ffff = 8999 Then
		gener = 1
		If gener = 1 Then
			i = 3 ' � 3-� ������
			While Ex1.cells(i,1).value > 0 ' �� ��� ��� ���� � 1-� ������� (������������ �������� �� ������) Ex1 ����� 1
                If Ex1.cells(i,3).value > 0 Then ' ���� N��� ������ 0 �� ...
					Eny = Ex1.cells(i,3) ' N���
					Eny2 = Ex1.cells(i,5) ' N����
					Name_gen = Ex1.cells(i,4) ' �������� ���������� 
					'gen.SetSel("Num = " & Eny & " & Node=" & Eny2)
					'gen.SetSel("Name = " & Name_gen)
					Id_generator = 0 
                    
					gen.SetSel("")
					j = gen.FindNextSel(-1)
					While j<>-1 ' �� ��� ��� ���� j �� ����� -1
						'If InStr(gen.cols("Name").z(j),Name_gen) Then
						If gen.cols("Name").z(j) = Name_gen Then ' ���� �������� ����������  �� Rastr ����� ���� ��� �� Excel						
							Id_generator = gen.cols("Num").z(j) ' �  ���������� Id_generator ������������� Num ����������� �� RASTR
						End If
						j = gen.FindNextSel(j)
					Wend
                    gen.SetSel("Num=" & Id_generator) ' ������� �� ������������ �������� � Id_generator
					j = gen.FindNextSel(-1) ' ������� ����� ������ � RASTR
					If j<>(-1) Then '  ���� j �� ����� -1 (�.�. ����� ��� ���� � RASTR)
						'nameg.Z(j)=Ex1.cells(i,4)
						'�� Excel ����������� �������� ���������� ��� ����������
						ModelType.Z(j) = Ex1.cells(i,6) 
						gNumBrand.Z(j) = Ex1.cells(i,8)
						ExciterId.Z(j) = Ex1.cells(i,9)
						ARSId.Z(j) = Ex1.cells(i,10)
						gIVActuatorId.Z(j) = Ex1.cells(i,11)								
						'napgen=Rastr.Calc("val","node","na","ny="&nodeg.Z(j))
						korrPGgen = 0
						If korrPGgen = 1 Then
							If (pgen.Z(j) > 1.05 * Ex1.cells(i,14)) and (pgen.Z(j) > 0) and ((napgen > 550) or (napgen < 500)) Then
								Ex1.cells(i,14) = pgen.Z(j)
							End If
							If (pgen.Z(j) < 0.9 * Ex1.cells(i,14)) and (pgen.Z(j)>0.5*Ex1.cells(i,14)) and (pgen.Z(j)>0) and (napgen>550 or napgen<500) Then
								Ex1.cells(i,14) = pgen.Z(j)
							End If
						End If
						'nameg.Z(j)=Ex1.cells(i,4)
						'nodeg.Z(j)=Ex1.cells(i,5)
						'pgen.Z(j)=Ex1.cells(i,12)
						pnom.Z(j) = Ex1.cells(i,14)
						unom.Z(j) = Ex1.cells(i,15)
						cosfi.Z(j) = Ex1.cells(i,16)
						Demp.Z(j) = Ex1.cells(i,17)
						mj.Z(j) = Ex1.cells(i,18)
						xd1.Z(j) = Ex1.cells(i,19)
						xd.Z(j) = Ex1.cells(i,20)
						xq.Z(j) = Ex1.cells(i,21)
						xd2.Z(j) = Ex1.cells(i,22)
						xq2.Z(j) = Ex1.cells(i,23)
						td01.Z(j) = Ex1.cells(i,24)
						td02.Z(j) = Ex1.cells(i,25)
						tq02.Z(j) = Ex1.cells(i,26)
						xq1.Z(j) = Ex1.cells(i,27)
						xl.Z(j) = Ex1.cells(i,28)
						x2.Z(j) = Ex1.cells(i,29)
						x2.Z(j) = Ex1.cells(i,30)
						tq01.Z(j) = Ex1.cells(i,31)
						'Ex1.cells(i,45) = gen.Cols("Pmin")
						'Ex1.cells(i,46) = gen.Cols("Pmax")
						Ex1.cells(i,47) = nameg.Z(j)
					Else
                    
					End If
				End If
				i = i+1
			Wend
			t.Printp("������� 1 '����������' - ���������!")
		End If
		
        i = 3
		vozbuzd = 1
		If vozbuzd = 1 Then
			While Ex2.cells(i,1).value>0
				'Eny=Ex2.cells(i,3)
				Name_gen = Ex2.cells(i,4)
				Id_generator = 0
				gen.SetSel("")
				j = gen.FindNextSel(-1)
				While j<>-1
					If gen.cols("Name").z(j) = Name_gen Then
						Id_generator = gen.cols("Num").z(j)
					End If
					j = gen.FindNextSel(j)
				Wend
				If (Ex2.cells(i,5).value>0) and (Id_generator>0) Then
					vozb.SetSel("")
					vozb.AddRow
					vozb.SetSel("Id = 0")
					j2 = vozb.FindNextSel(-1)
					
					If j2<>(-1) Then
						idv.Z(j2) = Id_generator
						namev.Z(j2) = Ex2.cells(i,4)
						ModelTypev.Z(j2) = Ex2.cells(i,5)
						'Brandv.Z(j2)=Ex2.cells(i,6)
						ExcControlIdv.Z(j2) = Id_generator
						
						If Ex2.cells(i,8)>0 Then
							ForcerIdv.Z(j2) = Id_generator
						End If
						
						Texcv.Z(j2) = Ex2.cells(i,9)
						Kig.Z(j2) = Ex2.cells(i,10)
						KIf.Z(j2) = Ex2.cells(i,11)
						Uf_min.Z(j2) = Ex2.cells(i,12)
						Uf_max.Z(j2) = Ex2.cells(i,13)
						If_min.Z(j2) = Ex2.cells(i,14)
						If_max.Z(j2) = Ex2.cells(i,15)
						Type_rgv.Z(j2) = Ex2.cells(i,16)
						vozbCustomModel.Z(j2) = Ex2.cells(i,17)
						vozbKarv.Z(j2) = Ex2.cells(i,18)
						vozbT3exc.Z(j2) = Ex2.cells(i,19)
					
					End If
				
				End If
				i = i + 1
			Wend
			t.Printp("������� 2 '����������� (��)' - ���������!")
		End If
		
		arv2 = 1
		If arv2 = 1 Then
			i = 3
			While Ex3.cells(i,1).value > 0
				'Eny=Ex3.cells(i,3)
				Name_gen = Ex3.cells(i,4)
				Id_generator = 0
				gen.SetSel("")
				j = gen.FindNextSel(-1)
				While j<>-1
					If gen.cols("Name").z(j) = Name_gen Then
						Id_generator = gen.cols("Num").z(j)
					End If
					j=gen.FindNextSel(j)
				Wend
				If Ex3.cells(i,5).value > 0 and Id_generator > 0 Then
					arv.SetSel("")
					arv.AddRow
					arv.SetSel("Id = 0")
					j2 = arv.FindNextSel(-1)
					If j2<>-1 Then
						ida.Z(j2) = Id_generator
						namea.Z(j2) = Ex3.cells(i,4)
						ModelTypea.Z(j2) = Ex3.cells(i,5)
						Branda.Z(j2) = Ex3.cells(i,6)
						Trv.Z(j2) = Ex3.cells(i,7)
						Ku.Z(j2) = Ex3.cells(i,8)
						Ku1.Z(j2) = Ex3.cells(i,9)
						KIf1.Z(j2) = Ex3.cells(i,10)
						Kf.Z(j2) = Ex3.cells(i,11)
						Kf1.Z(j2) = Ex3.cells(i,12)
						Tf.Z(j2) = Ex3.cells(i,13)
						Urv_min.Z(j2) = Ex3.cells(i,14)
						Urv_max.Z(j2) = Ex3.cells(i,15)
						Alpha.Z(j2) = Ex3.cells(i,16)
						arvCustomModel.Z(j2) = Ex3.cells(i,17)
						arvTINT.Z(j2) = Ex3.cells(i,18)
					End If
				End If
				i = i + 1
			Wend
			 t.Printp("������� 3 '��� (��)' - ���������!")
		End If
	End If 	' �������� ������ End If
	'---------------------------������������ IEEE-----------------
	vozbuzdIEEE = 1
	If vozbuzdIEEE = 1 Then
		i = 3
		
		While Ex5.cells(i,1).value>0
			'Eny=Ex5.cells(i,3)
			Name_gen = Ex5.cells(i,4)
			Id_generator = 0
			gen.SetSel("")
			j = gen.FindNextSel(-1)
			
			While j<>(-1)
				If gen.cols("Name").z(j) = Name_gen Then
					Id_generator = gen.cols("Num").z(j)
				End If
				j = gen.FindNextSel(j)
			Wend
			
			If (Ex5.cells(i,5).value>0) and (Id_generator>0) Then
				vieee.SetSel("")
				vieee.AddRow
				vieee.SetSel("Id = 0")
				j2 = vieee.FindNextSel (-1)
				
				If j2<>(-1) Then
					vista.Z(j2) = Ex5.cells(i,2)
					viId.Z(j2) = Id_generator
					viName.Z(j2) = Ex5.cells(i,4)
					viBrand.Z(j2) = Ex5.cells(i,6)
					viCustomModel.Z(j2) = Ex5.cells(i,7)
					viUELId.Z(j2) = Ex5.cells(i,8)
					viUELPos.Z(j2) = Ex5.cells(i,9)
					viOELId.Z(j2) = Ex5.cells(i,10)
					viOELPos.Z(j2) = Ex5.cells(i,11)
					
					If Ex5.cells(i,12)>0 Then
						viPSSId.Z(j2) = Id_generator
					End If
					
					viPSSPos.Z(j2) = Ex5.cells(i,13)
					viTe.Z(j2) = Ex5.cells(i,14)
					viKe.Z(j2) = Ex5.cells(i,15)
					viSe1.Z(j2) = Ex5.cells(i,16)
					viEfd1.Z(j2) = Ex5.cells(i,17)
					viVe1.Z(j2) = Ex5.cells(i,18)
					viSe2.Z(j2) = Ex5.cells(i,19)
					viEfd2.Z(j2) = Ex5.cells(i,20)
					viVe2.Z(j2) = Ex5.cells(i,21)
					viVemin.Z(j2) = Ex5.cells(i,22)
					viVrmin.Z(j2) = Ex5.cells(i,23)
					viVrmax.Z(j2) = Ex5.cells(i,24)
					viKa.Z(j2) = Ex5.cells(i,25)
					viTa.Z(j2) = Ex5.cells(i,26)
					viTf.Z(j2) = Ex5.cells(i,27)
					viKf.Z(j2) = Ex5.cells(i,28)
					viTc.Z(j2) = Ex5.cells(i,29)
					viTb.Z(j2) = Ex5.cells(i,30)
					viKv.Z(j2) = Ex5.cells(i,31)
					viTrh.Z(j2) = Ex5.cells(i,32)
					viKpr.Z(j2) = Ex5.cells(i,33)
					viKir.Z(j2) = Ex5.cells(i,34)
					viKdr.Z(j2) = Ex5.cells(i,35)
					viTdr.Z(j2) = Ex5.cells(i,36)
					viKc.Z(j2) = Ex5.cells(i,37)
					viKd.Z(j2) = Ex5.cells(i,38)
					viVfemax.Z(j2) = Ex5.cells(i,39)
					viVamin.Z(j2) = Ex5.cells(i,40)
					viVamax.Z(j2) = Ex5.cells(i,41)
					viKb.Z(j2) = Ex5.cells(i,42)
					viKh.Z(j2) = Ex5.cells(i,43)
					viKr.Z(j2) = Ex5.cells(i,44)
					viKn.Z(j2) = Ex5.cells(i,45)
					viEfdn.Z(j2) = Ex5.cells(i,46)
					viKlv.Z(j2) = Ex5.cells(i,47)
					viVlv.Z(j2) = Ex5.cells(i,48)
					viVimin.Z(j2) = Ex5.cells(i,49)
					viVimax.Z(j2) = Ex5.cells(i,50)
					viTf2.Z(j2) = Ex5.cells(i,51)
					viTf3.Z(j2) = Ex5.cells(i,52)
					viTk.Z(j2) = Ex5.cells(i,53)
					viTj.Z(j2) = Ex5.cells(i,54)
					viTh.Z(j2) = Ex5.cells(i,55)
					viVhmax.Z(j2) = Ex5.cells(i,56)
					viVfelim.Z(j2) = Ex5.cells(i,57)
					viKp.Z(j2) = Ex5.cells(i,58)
					viKpa.Z(j2) = Ex5.cells(i,59)
					viKia.Z(j2) = Ex5.cells(i,60)
					viKf1.Z(j2) = Ex5.cells(i,61)
					viKf2.Z(j2) = Ex5.cells(i,62)
					viKl.Z(j2) = Ex5.cells(i,63)
					viTb1.Z(j2) = Ex5.cells(i,64)
					viTc1.Z(j2) = Ex5.cells(i,65)
					viKlr.Z(j2) = Ex5.cells(i,66)
					viIlr.Z(j2) = Ex5.cells(i,67)
					viKi.Z(j2) = Ex5.cells(i,68)
					viTheta.Z(j2) = Ex5.cells(i,69)
					viVmmin.Z(j2) = Ex5.cells(i,70)
					viVmmax.Z(j2) = Ex5.cells(i,71)
					viKg.Z(j2) = Ex5.cells(i,72)
					viVBmax.Z(j2) = Ex5.cells(i,73)
					viVGmax.Z(j2) = Ex5.cells(i,74)
					viXl.Z(j2) = Ex5.cells(i,75)
					viKm.Z(j2) = Ex5.cells(i,76)
					viTm.Z(j2) = Ex5.cells(i,77)
					viTb2.Z(j2) = Ex5.cells(i,78)
					viTc2.Z(j2) = Ex5.cells(i,79)
					viTub1.Z(j2) = Ex5.cells(i,80)
					viTuc1.Z(j2) = Ex5.cells(i,81)
					viTub2.Z(j2) = Ex5.cells(i,82)
					viTuc2.Z(j2) = Ex5.cells(i,83)
					viTob1.Z(j2) = Ex5.cells(i,84)
					viToc1.Z(j2) = Ex5.cells(i,85)
					viTob2.Z(j2) = Ex5.cells(i,86)
					viToc2.Z(j2) = Ex5.cells(i,87)
					viAex.Z(j2) = Ex5.cells(i,88)
					viBex.Z(j2) = Ex5.cells(i,89)
					viKcf.Z(j2) = Ex5.cells(i,90)
					viKhf.Z(j2) = Ex5.cells(i,91)
					viKIf.Z(j2) = Ex5.cells(i,92)
					viSamovozb.Z(j2) = Ex5.cells(i,93)
					viTr.Z(j2) = Ex5.cells(i,94)
					viModel1 = Ex5.cells(i,5)
					viModel.Z(j2) = viModel1
				End If
			End If
			i = i + 1
		Wend
		t.Printp("������� 5 '����������� IEEE' - ���������!")
	End If
	'---------------------------������������� IEEE-----------------
	PSSE2 = 1
	If PSSE2 = 1 Then
		i = 3
		While Ex6.cells(i,1).value>0
			'Eny=Ex6.cells(i,3)
			Name_gen = Ex6.cells(i,4)
			Id_generator = 0
			gen.SetSel("")
			j = gen.FindNextSel(-1)
			
			While j<>(-1)
				If gen.cols("Name").z(j)=Name_gen Then
					Id_generator = gen.cols("Num").z(j)
				End If
				j = gen.FindNextSel(j)
			Wend
			
			If (Ex6.cells(i,5).value>0) and (Id_generator>0) Then
				stieee.SetSel("")
				stieee.AddRow
				stieee.SetSel("Id = 0")
				j2 = stieee.FindNextSel(-1)
				If j2<>(-1) Then
					ststa.Z(j2) = Ex6.cells(i,2)
					stId.Z(j2) = Id_generator
					stName.Z(j2) = Ex6.cells(i,4)
					stModel1 = Ex6.cells(i,5)
					stModel.Z(j2) = stModel1
					stBrand.Z(j2) = Ex6.cells(i,6)
					stCustomModel.Z(j2) = Ex6.cells(i,7)
					stInput1Type.Z(j2) = Ex6.cells(i,8)
					stInput2Type.Z(j2) = Ex6.cells(i,9)
					stVstmin.Z(j2) = Ex6.cells(i,10)
					stVstmax.Z(j2) = Ex6.cells(i,11)
					stKs1.Z(j2) = Ex6.cells(i,12)
					stT1.Z(j2) = Ex6.cells(i,13)
					stT2.Z(j2) = Ex6.cells(i,14)
					stT3.Z(j2) = Ex6.cells(i,15)
					stT4.Z(j2) = Ex6.cells(i,16)
					stT5.Z(j2) = Ex6.cells(i,17)
					stT6.Z(j2) = Ex6.cells(i,18)
					stT7.Z(j2) = Ex6.cells(i,19)
					stT8.Z(j2) = Ex6.cells(i,20)
					stT9.Z(j2) = Ex6.cells(i,21)
					stT10.Z(j2) = Ex6.cells(i,22)
					stT11.Z(j2) = Ex6.cells(i,23)
					stA1.Z(j2) = Ex6.cells(i,24)
					stA2.Z(j2) = Ex6.cells(i,25)
					stA3.Z(j2) = Ex6.cells(i,26)
					stA4.Z(j2) = Ex6.cells(i,27)
					stA5.Z(j2) = Ex6.cells(i,28)
					stA6.Z(j2) = Ex6.cells(i,29)
					stA7.Z(j2) = Ex6.cells(i,30)
					stA8.Z(j2) = Ex6.cells(i,31)
					stKs2.Z(j2) = Ex6.cells(i,32)
					stKs3.Z(j2) = Ex6.cells(i,33)
					stTw1.Z(j2) = Ex6.cells(i,34)
					stTw2.Z(j2) = Ex6.cells(i,35)
					stTw3.Z(j2) = Ex6.cells(i,36)
					stTw4.Z(j2) = Ex6.cells(i,37)
					stM.Z(j2) = Ex6.cells(i,38)
					stN.Z(j2) = Ex6.cells(i,39)
					stVsi1min.Z(j2) = Ex6.cells(i,40)
					stVsi1max.Z(j2) = Ex6.cells(i,41)
					stVsi2min.Z(j2) = Ex6.cells(i,42)
					stVsi2max.Z(j2) = Ex6.cells(i,43)
				End If
			End If
			i=i+1
		Wend
		t.Printp("������� 6 '������������ 1-3 PSS2' - ���������!")
	End If
	'---------------------------������������ PSS4B IEEE-----------------
	PSS4_2 = 1
	If PSS4_2 = 1 Then 
		i = 3
		While Ex7.cells(i,1).value > 0
			'Eny=Ex7.cells(i,3)
			Name_gen = Ex7.cells(i,4)
			Id_generator = 0
			gen.SetSel("")
			j = gen.FindNextSel(-1)
			While j<>-1
				If gen.cols("Name").z(j) = Name_gen Then
					Id_generator = gen.cols("Num").z(j)
				End If
				j = gen.FindNextSel (j)
			Wend
			If Ex7.cells(i,5).value>0 and Id_generator>0 Then
				pss4.SetSel("")
				pss4.AddRow
				pss4.SetSel("Id = 0")
				j2=pss4.FindNextSel (-1)
				If j2<>-1 Then
					pss4sta.Z(j2)=Ex7.cells(i,2)
					pss4Id.Z(j2)=Id_generator
					pss4Name.Z(j2)=Ex7.cells(i,4)
					pss4model1=Ex7.cells(i,5)
					pss4ModelType.Z(j2)=pss4model1
					pss4Brand.Z(j2)=Ex7.cells(i,6)
					pss4CustomModel.Z(j2)=Ex7.cells(i,7)
					pss4Input1Type.Z(j2)=Ex7.cells(i,8)
					pss4Input2Type.Z(j2)=Ex7.cells(i,9)
					pss4MBPSS1.Z(j2)=Ex7.cells(i,10)
					pss4MBPSS2.Z(j2)=Ex7.cells(i,11)
					pss4Vstmin.Z(j2)=Ex7.cells(i,12)
					pss4Vstmax.Z(j2)=Ex7.cells(i,13)
					pss4KL1.Z(j2)=Ex7.cells(i,14)
					pss4KL2.Z(j2)=Ex7.cells(i,15)
					pss4KL11.Z(j2)=Ex7.cells(i,16)
					pss4KL17.Z(j2)=Ex7.cells(i,17)
					pss4TL1.Z(j2)=Ex7.cells(i,18)
					pss4TL2.Z(j2)=Ex7.cells(i,19)
					pss4TL3.Z(j2)=Ex7.cells(i,20)
					pss4TL4.Z(j2)=Ex7.cells(i,21)
					pss4TL5.Z(j2)=Ex7.cells(i,22)
					pss4TL6.Z(j2)=Ex7.cells(i,23)
					pss4TL7.Z(j2)=Ex7.cells(i,24)
					pss4TL8.Z(j2)=Ex7.cells(i,25)
					pss4TL9.Z(j2)=Ex7.cells(i,26)
					pss4TL10.Z(j2)=Ex7.cells(i,27)
					pss4TL11.Z(j2)=Ex7.cells(i,28)
					pss4TL12.Z(j2)=Ex7.cells(i,29)
					pss4KL.Z(j2)=Ex7.cells(i,30)
					pss4Vlmin.Z(j2)=Ex7.cells(i,31)
					pss4Vlmax.Z(j2)=Ex7.cells(i,32)
					pss4KI1.Z(j2)=Ex7.cells(i,33)
					pss4KI2.Z(j2)=Ex7.cells(i,34)
					pss4KI11.Z(j2)=Ex7.cells(i,35)
					pss4KI17.Z(j2)=Ex7.cells(i,36)
					pss4TI1.Z(j2)=Ex7.cells(i,37)
					pss4TI2.Z(j2)=Ex7.cells(i,38)
					pss4TI3.Z(j2)=Ex7.cells(i,39)
					pss4TI4.Z(j2)=Ex7.cells(i,40)
					pss4TI5.Z(j2)=Ex7.cells(i,41)
					pss4TI6.Z(j2)=Ex7.cells(i,42)
					pss4TI7.Z(j2)=Ex7.cells(i,43)
					pss4TI8.Z(j2)=Ex7.cells(i,44)
					pss4TI9.Z(j2)=Ex7.cells(i,45)
					pss4TI10.Z(j2)=Ex7.cells(i,46)
					pss4TI11.Z(j2)=Ex7.cells(i,47)
					pss4TI12.Z(j2)=Ex7.cells(i,48)
					pss4KI.Z(j2)=Ex7.cells(i,49)
					pss4Vimin.Z(j2)=Ex7.cells(i,50)
					pss4Vimax.Z(j2)=Ex7.cells(i,51)
					pss4KH1.Z(j2)=Ex7.cells(i,52)
					pss4KH2.Z(j2)=Ex7.cells(i,53)
					pss4KH11.Z(j2)=Ex7.cells(i,54)
					pss4KH17.Z(j2)=Ex7.cells(i,55)
					pss4TH1.Z(j2)=Ex7.cells(i,56)
					pss4TH2.Z(j2)=Ex7.cells(i,57)
					pss4TH3.Z(j2)=Ex7.cells(i,58)
					pss4TH4.Z(j2)=Ex7.cells(i,59)
					pss4TH5.Z(j2)=Ex7.cells(i,60)
					pss4TH6.Z(j2)=Ex7.cells(i,61)
					pss4TH7.Z(j2)=Ex7.cells(i,62)
					pss4TH8.Z(j2)=Ex7.cells(i,63)
					pss4TH9.Z(j2)=Ex7.cells(i,64)
					pss4TH10.Z(j2)=Ex7.cells(i,65)
					pss4TH11.Z(j2)=Ex7.cells(i,66)
					pss4TH12.Z(j2)=Ex7.cells(i,67)
					pss4KH.Z(j2)=Ex7.cells(i,68)
					pss4Vhmin.Z(j2)=Ex7.cells(i,69)
					pss4Vhmax.Z(j2)=Ex7.cells(i,70)
				End If
			End If
			i=i+1
		Wend
		t.Printp("������� 7 '������������ PSS 4' - ���������!")
	End If
	'----------------�������----------------
	turbina = 1
	If turbina = 1 Then
		i = 3
		While Ex4.cells(i,1).value > 0
			'Eny=Ex4.cells(i,3)
			Name_gen = Ex4.cells(i,4)
			Id_generator = 0
			gen.SetSel("")
			j = gen.FindNextSel (-1)
			While j<>-1
				If gen.cols("Name").z(j) = Name_gen Then
					Id_generator = gen.cols("Num").z(j)
				End If
				j = gen.FindNextSel (j)
			Wend
			If Ex4.cells(i,5).value > 0 and Id_generator > 0 Then
				ars.SetSel("")
				ars.AddRow
				ars.SetSel("Id = 0")
				j2 = ars.FindNextSel (-1)
				If j2<>-1 Then
					arsids.Z(j2)=Ex4.cells(i,3)
					ideg1=Id_generator
					arsname.Z(j2)=Ex4.cells(i,4)
					ModelTypes111=Ex4.cells(i,5)
					arsModelTypes.Z(j2)=ModelTypes111
					arsBrands.Z(j2)=Ex4.cells(i,6)
					If Ex4.cells(i,7)>0 Then
						arsGovernorId.Z(j2)=Id_generator
					End If
					arsao.Z(j2)=Ex4.cells(i,8)
					arsaz.Z(j2)=Ex4.cells(i,9)
					arsotmin.Z(j2)=Ex4.cells(i,10)
					arsotmax.Z(j2)=Ex4.cells(i,11)
					arsstrs.Z(j2)=Ex4.cells(i,12)
					arszn.Z(j2)=Ex4.cells(i,13)
					arsdpo.Z(j2)=Ex4.cells(i,14)
					arsThp.Z(j2)=Ex4.cells(i,15)
					arsTrlp.Z(j2)=Ex4.cells(i,16)
					arsTw.Z(j2)=Ex4.cells(i,17)
					arspt.Z(j2)=Ex4.cells(i,18)
					arsMu.Z(j2)=Ex4.cells(i,19)
					arsPsteam.Z(j2)=Ex4.cells(i,20)
					arsCustomModel.Z(j2)=Ex4.cells(i,21)
				End If
			End If
			i=i+1
		Wend
	' ���
		i = 3
		While Ex12.cells(i,1).value > 0
			'Eny=Ex12.cells(i,3)
			Name_gen = Ex12.cells(i,4)
			Id_generator = 0
			gen.SetSel("")
			j = gen.FindNextSel (-1)
			While j<>-1
				If gen.cols("Name").z(j) = Name_gen Then
					Id_generator = gen.cols("Num").z(j)
				End If
				j = gen.FindNextSel (j)
			Wend
			If Ex12.cells(i,5).value > 0 and Id_generator > 0 Then
				Governor.SetSel("")
				Governor.AddRow
				Governor.SetSel("Id = 0")
				j2 = Governor.FindNextSel(-1)
				If j2<>-1 Then
					Governorsta.Z(j2)=Ex12.cells(i,2)
					GovernorId.Z(j2)=Id_generator
					GovernorName.Z(j2)=Ex12.cells(i,4)
					GovernorModelType.Z(j2)=Ex12.cells(i,5)
					GovernorBrand.Z(j2)=Ex12.cells(i,6)
					Governorstrs.Z(j2)=Ex12.cells(i,7)
					Governorzn.Z(j2)=Ex12.cells(i,8)
					GovernorTr.Z(j2)=Ex12.cells(i,9)
					Governorotmin.Z(j2)=Ex12.cells(i,10)
					Governorotmax.Z(j2)=Ex12.cells(i,11)
					GovernorCVmin.Z(j2)=Ex12.cells(i,12)
					GovernorCVmax.Z(j2)=Ex12.cells(i,13)
					GovernorT1.Z(j2)=Ex12.cells(i,14)
					GovernorK1.Z(j2)=Ex12.cells(i,15)
					GovernorK2.Z(j2)=Ex12.cells(i,16)
					GovernorBoilerId.Z(j2)=Ex12.cells(i,17)
				End If
			End If
			i=i+1
		Wend
		t.Printp("������� 4 � 12 '�������  � ���' - ���������!")
	End If
	'-----------����������
	forc2 = 1
	If forc2 = 1 Then
		i = 3
		While Ex8.cells(i,1).value > 0
			'Eny=Ex8.cells(i,3)
			Name_gen = Ex8.cells(i,4)
			Id_generator = 0
			gen.SetSel("")
			j = gen.FindNextSel (-1)
			While j<>-1
				If gen.cols("Name").z(j) = Name_gen Then
					Id_generator = gen.cols("Num").z(j)
				End If
				j = gen.FindNextSel (j)
			Wend
			If Ex8.cells(i,5).value>0 and Id_generator>0 Then
				forc.SetSel("")
				forc.AddRow
				forc.SetSel("Id = 0")
				j2=forc.FindNextSel (-1)
				If j2<>-1 Then
					idf.Z(j2)=Id_generator
					namef.Z(j2)=Ex8.cells(i,4)
					ModelTypef.Z(j2)=Ex8.cells(i,5)
					Ubf.Z(j2)=Ex8.cells(i,7)
					Uef.Z(j2)=Ex8.cells(i,8)
					Rf.Z(j2)=Ex8.cells(i,9)
					Texc_f.Z(j2)=Ex8.cells(i,10)
					Tz_in.Z(j2)=Ex8.cells(i,11)
					Tz_out.Z(j2)=Ex8.cells(i,12)
					Ubrf.Z(j2)=Ex8.cells(i,13)
					Uerf.Z(j2)=Ex8.cells(i,14)
					Rrf.Z(j2)=Ex8.cells(i,15)
					Texc_rf.Z(j2)=Ex8.cells(i,16)
				End If
			End If
			i=i+1
		Wend
		t.Printp("������� 5 '������������ 1-3 PSS2' - ���������!")
	End If
	'-----------���
	OMV_2 = 0 ' ����
	If OMV_2 = 1 Then
		i = 3
		While Ex9.cells(i,1).value > 0
			Eny=Ex9.cells(i,3)
			If Ex9.cells(i,5).value > 0 and Ex9.cells(i,3).value > 0 Then
				omv.SetSel("")
				omv.AddRow
				omv.SetSel("Id = 0")
				j2 = omv.FindNextSel (-1)
				If j2<>-1 Then
					omvsta.Z(j2)=Ex9.cells(i,2)
					omvId.Z(j2)=Ex9.cells(i,3)
					omvName.Z(j2)=Ex9.cells(i,4)
					omvModelType.Z(j2)=3
					omvBrand.Z(j2)=Ex9.cells(i,6)
					omvCustomModel.Z(j2)=Ex9.cells(i,7)
					omvTu1.Z(j2)=Ex9.cells(i,8)
					omvTu2.Z(j2)=Ex9.cells(i,9)
					omvTu3.Z(j2)=Ex9.cells(i,10)
					omvTu4.Z(j2)=Ex9.cells(i,11)
					omvVulmin.Z(j2)=Ex9.cells(i,12)
					omvVulmax.Z(j2)=Ex9.cells(i,13)
					omvKul.Z(j2)=Ex9.cells(i,14)
					omvKui.Z(j2)=Ex9.cells(i,15)
					omvVuimin.Z(j2)=Ex9.cells(i,16)
					omvVuimax.Z(j2)=Ex9.cells(i,17)
					omvKuf.Z(j2)=Ex9.cells(i,18)
					omvTuf.Z(j2)=Ex9.cells(i,19)
					omvKuc.Z(j2)=Ex9.cells(i,20)
					omvKuc.Z(j2)=Ex9.cells(i,21)
					omvVurmax.Z(j2)=Ex9.cells(i,22)
					omvVucmax.Z(j2)=Ex9.cells(i,23)
					omvTuV.Z(j2)=Ex9.cells(i,24)
					omvTuP.Z(j2)=Ex9.cells(i,25)
					omvTuQ.Z(j2)=Ex9.cells(i,26)
					omvK1.Z(j2)=Ex9.cells(i,27)
					omvK2.Z(j2)=Ex9.cells(i,28)
					omvDepEndency_F1.Z(j2)=Ex9.cells(i,29)
					omvOutput.Z(j2)=Ex9.cells(i,30)
					omvKl.Z(j2)=Ex9.cells(i,31)
				End If
			End If
			i=i+1
		Wend
		t.Printp("������� 9 '���' - ���������!")
	End If  
	'----------���
	BOR_2 = 0 ' ����
	If BOR_2 = 1 Then
		i = 3
		While Ex10.cells(i,1).value > 0
			Eny = Ex10.cells(i,3)
			If Ex10.cells(i,5).value > 0 and Ex10.cells(i,3).value > 0 Then
				bor.SetSel("")
				bor.AddRow
				bor.SetSel("Id = 0")
				j2=bor.FindNextSel (-1)
				If j2<>-1 Then
					borsta.Z(j2)=Ex10.cells(i,2)
					borId.Z(j2)=Ex10.cells(i,3)
					borName.Z(j2)=Ex10.cells(i,4)
					borModelType.Z(j2)=Ex10.cells(i,5)
					borBrand.Z(j2)=Ex10.cells(i,6)
					borCustomModel.Z(j2)=Ex10.cells(i,7)
					borIfMax.Z(j2)=Ex10.cells(i,8)
					borIfth.Z(j2)=Ex10.cells(i,9)
					borKexpIf.Z(j2)=Ex10.cells(i,10)
					borKR3.Z(j2)=Ex10.cells(i,11)
					borKR3i.Z(j2)=Ex10.cells(i,12)
					borTc23.Z(j2)=Ex10.cells(i,13)
					borTb23.Z(j2)=Ex10.cells(i,14)
					borTc13.Z(j2)=Ex10.cells(i,15)
					borTb13.Z(j2)=Ex10.cells(i,16)
					borVamin.Z(j2)=Ex10.cells(i,17)
					borVamax.Z(j2)=Ex10.cells(i,18)
					borTdel.Z(j2)=Ex10.cells(i,19)
					borKth.Z(j2)=Ex10.cells(i,20)
					borKToF.Z(j2)=Ex10.cells(i,21)
					borKcF.Z(j2)=Ex10.cells(i,22)
					borKhF.Z(j2)=Ex10.cells(i,23)
					borTRFout.Z(j2)=Ex10.cells(i,24)
					borTr.Z(j2)=Ex10.cells(i,25)
					borOutput.Z(j2)=Ex10.cells(i,26)
					borKl.Z(j2)=Ex10.cells(i,27)
				End If
			End If
			i=i+1
		Wend
		t.Printp("������� 10 '���' - ���������!")
	End If
	'----------��� PQ
	OMVPQ_2 = 0 ' ����
	If OMVPQ_2 = 1 Then
		i = 3
		While Ex11.cells(i,1).value > 0
			'Eny=Ex11.cells(i,3)
			If Ex11.cells(i,2).value > 0  Then
				i0 = 3
				While Ex11.cells(i,i0).value <> 0 or Ex11.cells(i,i0+1).value <> 0
					FuncPQ.SetSel("")
					FuncPQ.AddRow
					FuncPQ.SetSel("Id = 0")
					j2 = FuncPQ.FindNextSel(-1)
					If j2<>-1 Then
						FuncPQId.Z(j2)=Ex11.cells(i,2)
						FuncPQP.Z(j2)=Ex11.cells(i,i0)
						FuncPQQ.Z(j2)=Ex11.cells(i,i0+1)
						i0=i0+2
					End If
				Wend
			End If
			i=i+1
		Wend
		t.Printp("������� 11 '����������� Q(P)' - ���������!")
	End If
	'------------- DECS-400 ----------------------
	DECS_400 = 1
	If DECS_400 = 1 Then 
		i = 3
		While Ex13.cells(i,1).value > 0
			Eny = Ex13.cells(i,3)
			Name_gen = Ex13.cells(i,4)
			Id_generator = 0
			gen.SetSel("")
			j = gen.FindNextSel(-1)
			While j<>-1
				If gen.cols("Name").z(j) = Name_gen Then
					Id_generator = gen.cols("Num").z(j)
				End If
				j = gen.FindNextSel (j)
			Wend
			If Ex13.cells(i,5).value>0 and Id_generator>0 Then
				decs400.SetSel("")
				decs400.AddRow
				decs400.SetSel("Id = 0")
				j2 = decs400.FindNextSel(-1)
				If j2<>-1 Then
					decs400sta.Z(j2)=Ex13.cells(i,2)
					decs400Id.Z(j2)=Id_generator
					decs400Name.Z(j2)=Ex13.cells(i,4)
					decs400ModelType.Z(j2)=Ex13.cells(i,5)
					decs400Brand.Z(j2)=Ex13.cells(i,6)
					decs400CustomModel.Z(j2)=Ex13.cells(i,7)
					decs400PSSId.Z(j2)=Ex13.cells(i,8)
					decs400UELId.Z(j2)=Ex13.cells(i,9)
					decs400OELId.Z(j2)=Ex13.cells(i,10)
					decs400Xl.Z(j2)=Ex13.cells(i,11)
					decs400DRP.Z(j2)=Ex13.cells(i,12)
					decs400VrMin.Z(j2)=Ex13.cells(i,13)
					decs400VrMax.Z(j2)=Ex13.cells(i,14)
					decs400VmMin.Z(j2)=Ex13.cells(i,15)
					decs400VmMax.Z(j2)=Ex13.cells(i,16)
					decs400VbMax.Z(j2)=Ex13.cells(i,17)
					decs400Kc.Z(j2)=Ex13.cells(i,18)
					decs400Kp.Z(j2)=Ex13.cells(i,19)
					decs400Kpm.Z(j2)=Ex13.cells(i,20)
					decs400Kpr.Z(j2)=Ex13.cells(i,21)
					decs400Kir.Z(j2)=Ex13.cells(i,22)
					decs400Kpd.Z(j2)=Ex13.cells(i,23)
					decs400Ta.Z(j2)=Ex13.cells(i,24)
					decs400Td.Z(j2)=Ex13.cells(i,25)
					decs400Tr.Z(j2)=Ex13.cells(i,26)
					decs400SelfExc.Z(j2)=Ex13.cells(i,27)
					decs400Del.Z(j2)=Ex13.cells(i,28)
				End If
			End If
			i=i+1
		Wend
		t.Printp("������� 13 'DECS - 400' - ���������!")
	End If
	'------------- Thyne-4 ----------------------
	Thyne4 = 1
	If Thyne4 = 1 Then 
		i = 3
		While Ex14.cells(i,1).value > 0
			Eny = Ex14.cells(i,3)
			Name_gen = Ex14.cells(i,4)
			Id_generator = 0
			gen.SetSel("")
			j = gen.FindNextSel(-1)
			While j<>-1
				If gen.cols("Name").z(j) = Name_gen Then
					Id_generator = gen.cols("Num").z(j)
				End If
				j = gen.FindNextSel (j)
			Wend
			If Ex14.cells(i,5).value > 0 and Id_generator > 0 Then
				Thyne.SetSel("")
				Thyne.AddRow
				Thyne.SetSel("Id = 0")
				j2 = Thyne.FindNextSel (-1)
				If j2<>-1 Then
					Thynesta.Z(j2)=Ex14.cells(i,2)
					ThyneId.Z(j2)=Id_generator
					ThyneName.Z(j2)=Ex14.cells(i,4)
					ThyneModelType.Z(j2)=Ex14.cells(i,5)
					ThyneBrand.Z(j2)=Ex14.cells(i,6)
					ThyneCustomModel.Z(j2)=Ex14.cells(i,7)
					ThyneUELId.Z(j2)=Ex14.cells(i,8)
					ThynePSSId.Z(j2)=Ex14.cells(i,9)
					ThyneAex.Z(j2)=Ex14.cells(i,10)
					ThyneBex.Z(j2)=Ex14.cells(i,11)
					ThyneAlpha.Z(j2)=Ex14.cells(i,12)
					ThyneBeta.Z(j2)=Ex14.cells(i,13)
					ThyneIfdMin.Z(j2)=Ex14.cells(i,14)
					ThyneKc.Z(j2)=Ex14.cells(i,15)
					ThyneKd1.Z(j2)=Ex14.cells(i,16)
					ThyneKd2.Z(j2)=Ex14.cells(i,17)
					ThyneKe.Z(j2)=Ex14.cells(i,18)
					ThyneKetb.Z(j2)=Ex14.cells(i,19)
					ThyneKh.Z(j2)=Ex14.cells(i,20)
					ThyneKp1.Z(j2)=Ex14.cells(i,21)
					ThyneKp2.Z(j2)=Ex14.cells(i,22)
					ThyneKp3.Z(j2)=Ex14.cells(i,23)
					ThyneTd1.Z(j2)=Ex14.cells(i,24)
					ThyneTe1.Z(j2)=Ex14.cells(i,25)
					ThyneTe2.Z(j2)=Ex14.cells(i,26)
					ThyneTi1.Z(j2)=Ex14.cells(i,27)
					ThyneTi2.Z(j2)=Ex14.cells(i,28)
					ThyneTi3.Z(j2)=Ex14.cells(i,29)
					ThyneTr1.Z(j2)=Ex14.cells(i,30)
					ThyneTr2.Z(j2)=Ex14.cells(i,31)
					ThyneTr3.Z(j2)=Ex14.cells(i,32)
					ThyneTr4.Z(j2)=Ex14.cells(i,33)
					ThyneVO1Min.Z(j2)=Ex14.cells(i,34)
					ThyneVO1Max.Z(j2)=Ex14.cells(i,35)
					ThyneVO2Min.Z(j2)=Ex14.cells(i,36)
					ThyneVO2Max.Z(j2)=Ex14.cells(i,37)
					ThyneVO3Min.Z(j2)=Ex14.cells(i,38)
					ThyneVO3Max.Z(j2)=Ex14.cells(i,39)
					ThyneVO3Max.Z(j2)=Ex14.cells(i,40)
					ThyneVD1Min.Z(j2)=Ex14.cells(i,41)
					ThyneVD1Max.Z(j2)=Ex14.cells(i,42)
					ThyneVI1Min.Z(j2)=Ex14.cells(i,43)
					ThyneVI1Max.Z(j2)=Ex14.cells(i,44)
					ThyneVI2Min.Z(j2)=Ex14.cells(i,45)
					ThyneVI2Max.Z(j2)=Ex14.cells(i,46)
					ThyneVI3Min.Z(j2)=Ex14.cells(i,47)
					ThyneVI3Max.Z(j2)=Ex14.cells(i,48)
					ThyneVP1Min.Z(j2)=Ex14.cells(i,49)
					ThyneVP1Max.Z(j2)=Ex14.cells(i,50)
					ThyneVP2Min.Z(j2)=Ex14.cells(i,51)
					ThyneVP2Max.Z(j2)=Ex14.cells(i,52)
					ThyneVP3Min.Z(j2)=Ex14.cells(i,53)
					ThyneVP3Max.Z(j2)=Ex14.cells(i,54)
					ThyneVrMin.Z(j2)=Ex14.cells(i,55)
					ThyneVrMax.Z(j2)=Ex14.cells(i,56)
					ThyneXp.Z(j2)=Ex14.cells(i,57)
				End If
			End If
			i=i+1
		Wend
		t.Printp("������� 14 'Thyne-4' - ���������!")
	End If
	'---------------------------------------------------
	
	'-----------------------
	otklstab = 0 ' ����
	If otklstab = 1 Then
		gen.SetSel("")
		j = gen.FindNextSel(-1)
		While j<>-1
			stagen = nsta.Z(j)
			'rastr.printp stagen
			If nsta.Z(j) = True Then
				stagen = 1
			Else
				stagen = 0
			End If
			'rastr.printp stagen
			agrgen=numg.Z(j)
			'rastr.printp agrgen
			'vozb.SetSel("Id="&agrgen)
			'vozb.cols("sta").calc(stagen)
			'arv.SetSel("Id="&agrgen)
			'arv.cols("sta").calc(stagen)
			vieee.SetSel("Id="&agrgen)
			vieee.cols("sta").calc(stagen)
			'stieee.SetSel("Id="&agrgen)
			'stieee.cols("sta").calc(stagen)
			gen.SetSel("")
			j=gen.FindNextSel (j)
		Wend
	End If
	'---------------------------------------------------
	comDynamics = 1
	If comDynamics = 1 Then
		t.Tables("com_dynamics").Cols("Tras").z(0)=5
		t.Tables("com_dynamics").Cols("Hint").z(0)=0.001
		t.Tables("com_dynamics").Cols("Hmin").z(0)=0.001
		t.Tables("com_dynamics").Cols("Mint").z(0)=0
		t.Tables("com_dynamics").Cols("SMint").z(0)=2
		t.Tables("com_dynamics").Cols("IntEpsilon").z(0)=0.001
		t.Tables("com_dynamics").Cols("Hmax").z(0)=1
		t.Tables("com_dynamics").Cols("Tf").z(0)=0.04
		t.Tables("com_dynamics").Cols("dEf").z(0)=0.01
		t.Tables("com_dynamics").Cols("Hout").z(0)=0.01
		t.Tables("com_dynamics").Cols("Npf").z(0)=999
		t.Tables("com_dynamics").Cols("frSXNtoY").z(0)=0.3
		t.Tables("com_dynamics").Cols("SXNTolerance").z(0)=0.01
		t.Tables("com_dynamics").Cols("SnapPath").z(0)="C:\tmp\"
		t.Tables("com_dynamics").Cols("MaxResultFiles").z(0)=3
		t.Tables("com_dynamics").Cols("SnapTemplate").z(0)="<count>.sna"
		t.Tables("com_dynamics").Cols("SnapAutoLoad").z(0)=1
		t.Tables("com_dynamics").Cols("SnapMaxCount").z(0)=3

		t.Tables("AltUnit").delrows
		t.Tables("AltUnit").AddRow
		t.Tables("AltUnit").AddRow
		t.Tables("AltUnit").Cols("Activ").z(0)=1
		t.Tables("AltUnit").Cols("Unit").z(0)="���*�"
		t.Tables("AltUnit").Cols("Alt").z(0)="�"
		t.Tables("AltUnit").Cols("Formula").z(0)="nonz(cosFi)/(Pnom)"
		t.Tables("AltUnit").Cols("Prec").z(0)=3
		t.Tables("AltUnit").Cols("Tabl").z(0)="Generator"

		t.Tables("AltUnit").Cols("Activ").z(1)=1
		t.Tables("AltUnit").Cols("Unit").z(1)="��"
		t.Tables("AltUnit").Cols("Alt").z(1)="�.�."
		t.Tables("AltUnit").Cols("Formula").z(1)="Pnom/(Ugnom*Ugnom*nonz(cosFi))"
		t.Tables("AltUnit").Cols("Prec").z(1)=4
		t.Tables("AltUnit").Cols("Tabl").z(1)="Generator"
	End If
	'---------------------------------------------------
	Custom_Models = 1
	If Custom_Models = 1 Then 
		'������ ������ ��� ���������������� ���������
		Link = LinkCustomModels
		'1.----------AC8B--------------------
		t.Tables("CustomDeviceMap").delrows
		t.Tables("CustomDeviceMap").AddRow
		t.Tables("CustomDeviceMap").Cols("Id").z(0)="1"
		t.Tables("CustomDeviceMap").Cols("Module").z(0)= Link +"dll\AC8B"
		t.Tables("CustomDeviceMap").Cols("Name").z(0)="AC8B"
		t.Tables("CustomDeviceMap").Cols("InstanceTable").z(0)="DFWIEEE421"
		t.Tables("CustomDeviceMap").Cols("Tag").z(0)="Exciter"
		t.Tables("CustomDeviceMap").Cols("LinkList").z(0)="Generator"
		t.Tables("CustomDeviceMap").Cols("SiblingLinkList").z(0)=" "

		'2.----------ARV_REM--------------------
		t.Tables("CustomDeviceMap").AddRow
		t.Tables("CustomDeviceMap").Cols("Id").z(1)="2"
		t.Tables("CustomDeviceMap").Cols("Module").z(1)=Link +"dll\ARV_REM"
		t.Tables("CustomDeviceMap").Cols("Name").z(1)="ARV_REM"
		t.Tables("CustomDeviceMap").Cols("InstanceTable").z(1)="ExcControl"
		t.Tables("CustomDeviceMap").Cols("Tag").z(1)="ExcControl"
		t.Tables("CustomDeviceMap").Cols("LinkList").z(1)="Exciter"
		t.Tables("CustomDeviceMap").Cols("SiblingLinkList").z(1)=" "

		'3.----------ARV2M--------------------
		t.Tables("CustomDeviceMap").AddRow
		t.Tables("CustomDeviceMap").Cols("Id").z(2)="3"
		t.Tables("CustomDeviceMap").Cols("Module").z(2)=Link +"dll\ARV2M"
		t.Tables("CustomDeviceMap").Cols("Name").z(2)="ARV2M"
		t.Tables("CustomDeviceMap").Cols("InstanceTable").z(2)="ExcControl"
		t.Tables("CustomDeviceMap").Cols("Tag").z(2)="ExcControl"
		t.Tables("CustomDeviceMap").Cols("LinkList").z(2)="Exciter"
		t.Tables("CustomDeviceMap").Cols("SiblingLinkList").z(2)=" "

		'4.----------ARV-3MTK--------------------
		t.Tables("CustomDeviceMap").AddRow
		t.Tables("CustomDeviceMap").Cols("Id").z(3)="4"
		t.Tables("CustomDeviceMap").Cols("Module").z(3)=Link +"dll\ARV-3MTK"
		t.Tables("CustomDeviceMap").Cols("Name").z(3)="ARV-3MTK"
		t.Tables("CustomDeviceMap").Cols("InstanceTable").z(3)="ExcControl"
		t.Tables("CustomDeviceMap").Cols("Tag").z(3)="ExcControl"
		t.Tables("CustomDeviceMap").Cols("LinkList").z(3)="Exciter"
		t.Tables("CustomDeviceMap").Cols("SiblingLinkList").z(3)=" "

		'5.----------ARV-4M--------------------
		t.Tables("CustomDeviceMap").AddRow
		t.Tables("CustomDeviceMap").Cols("Id").z(4)="5"
		t.Tables("CustomDeviceMap").Cols("Module").z(4)=Link +"dll\ARV-4M"
		t.Tables("CustomDeviceMap").Cols("Name").z(4)="ARV-4M"
		t.Tables("CustomDeviceMap").Cols("InstanceTable").z(4)="ExcControl"
		t.Tables("CustomDeviceMap").Cols("Tag").z(4)="ExcControl"
		t.Tables("CustomDeviceMap").Cols("LinkList").z(4)="Exciter"
		t.Tables("CustomDeviceMap").Cols("SiblingLinkList").z(4)=" "

		'6.----------ARVNL--------------------
		t.Tables("CustomDeviceMap").AddRow
		t.Tables("CustomDeviceMap").Cols("Id").z(5)="6"
		t.Tables("CustomDeviceMap").Cols("Module").z(5)=Link +"dll\ARVNL"
		t.Tables("CustomDeviceMap").Cols("Name").z(5)="ARVNL"
		t.Tables("CustomDeviceMap").Cols("InstanceTable").z(5)="ExcControl"
		t.Tables("CustomDeviceMap").Cols("Tag").z(5)="ExcControl"
		t.Tables("CustomDeviceMap").Cols("LinkList").z(5)="Exciter"
		t.Tables("CustomDeviceMap").Cols("SiblingLinkList").z(5)=" "

		'7.----------ARVSDE--------------------
		t.Tables("CustomDeviceMap").AddRow
		t.Tables("CustomDeviceMap").Cols("Id").z(6)="7"
		t.Tables("CustomDeviceMap").Cols("Module").z(6)=Link +"dll\ARVSDE"
		t.Tables("CustomDeviceMap").Cols("Name").z(6)="ARVSDE"
		t.Tables("CustomDeviceMap").Cols("InstanceTable").z(6)="ExcControl"
		t.Tables("CustomDeviceMap").Cols("Tag").z(6)="ExcControl"
		t.Tables("CustomDeviceMap").Cols("LinkList").z(6)="Exciter"
		t.Tables("CustomDeviceMap").Cols("SiblingLinkList").z(6)=" "

		'8.----------ARVSDS--------------------
		t.Tables("CustomDeviceMap").AddRow
		t.Tables("CustomDeviceMap").Cols("Id").z(7)="8"
		t.Tables("CustomDeviceMap").Cols("Module").z(7)=Link +"dll\ARVSDS"
		t.Tables("CustomDeviceMap").Cols("Name").z(7)="ARVSDS"
		t.Tables("CustomDeviceMap").Cols("InstanceTable").z(7)="ExcControl"
		t.Tables("CustomDeviceMap").Cols("Tag").z(7)="ExcControl"
		t.Tables("CustomDeviceMap").Cols("LinkList").z(7)="Exciter"
		t.Tables("CustomDeviceMap").Cols("SiblingLinkList").z(7)=" "

		'9.----------ARVSG--------------------
		t.Tables("CustomDeviceMap").AddRow
		t.Tables("CustomDeviceMap").Cols("Id").z(8)="9"
		t.Tables("CustomDeviceMap").Cols("Module").z(8)=Link +"dll\ARVSG"
		t.Tables("CustomDeviceMap").Cols("Name").z(8)="ARVSG"
		t.Tables("CustomDeviceMap").Cols("InstanceTable").z(8)="ExcControl"
		t.Tables("CustomDeviceMap").Cols("Tag").z(8)="ExcControl"
		t.Tables("CustomDeviceMap").Cols("LinkList").z(8)="Exciter"
		t.Tables("CustomDeviceMap").Cols("SiblingLinkList").z(8)=" "

		'10.----------AVR2--------------------
		t.Tables("CustomDeviceMap").AddRow
		t.Tables("CustomDeviceMap").Cols("Id").z(9)="10"
		t.Tables("CustomDeviceMap").Cols("Module").z(9)=Link +"dll\AVR2"
		t.Tables("CustomDeviceMap").Cols("Name").z(9)="AVR2"
		t.Tables("CustomDeviceMap").Cols("InstanceTable").z(9)="ExcControl"
		t.Tables("CustomDeviceMap").Cols("Tag").z(9)="ExcControl"
		t.Tables("CustomDeviceMap").Cols("LinkList").z(9)="Exciter"
		t.Tables("CustomDeviceMap").Cols("SiblingLinkList").z(9)=" "

		'11.----------AVR-2_br--------------------
		t.Tables("CustomDeviceMap").AddRow
		t.Tables("CustomDeviceMap").Cols("Id").z(10)="11"
		t.Tables("CustomDeviceMap").Cols("Module").z(10)=Link +"dll\AVR-2_br"
		t.Tables("CustomDeviceMap").Cols("Name").z(10)="AVR-2_br"
		t.Tables("CustomDeviceMap").Cols("InstanceTable").z(10)="ExcControl"
		t.Tables("CustomDeviceMap").Cols("Tag").z(10)="ExcControl"
		t.Tables("CustomDeviceMap").Cols("LinkList").z(10)="Exciter"
		t.Tables("CustomDeviceMap").Cols("SiblingLinkList").z(10)=" "

		'12.----------DECS--------------------
		t.Tables("CustomDeviceMap").AddRow
		t.Tables("CustomDeviceMap").Cols("Id").z(11)="12"
		t.Tables("CustomDeviceMap").Cols("Module").z(11)=Link +"dll\DECS"
		t.Tables("CustomDeviceMap").Cols("Name").z(11)="DECS"
		t.Tables("CustomDeviceMap").Cols("InstanceTable").z(11)="DFWIEEE421"
		t.Tables("CustomDeviceMap").Cols("Tag").z(11)="Exciter"
		t.Tables("CustomDeviceMap").Cols("LinkList").z(11)="Generator"
		t.Tables("CustomDeviceMap").Cols("SiblingLinkList").z(11)=" "

		'13.----------EAA--------------------
		t.Tables("CustomDeviceMap").AddRow
		t.Tables("CustomDeviceMap").Cols("Id").z(12)="13"
		t.Tables("CustomDeviceMap").Cols("Module").z(12)=Link +"dll\EAA"
		t.Tables("CustomDeviceMap").Cols("Name").z(12)="EAA"
		t.Tables("CustomDeviceMap").Cols("InstanceTable").z(12)="ExcControl"
		t.Tables("CustomDeviceMap").Cols("Tag").z(12)="ExcControl"
		t.Tables("CustomDeviceMap").Cols("LinkList").z(12)="Exciter"
		t.Tables("CustomDeviceMap").Cols("SiblingLinkList").z(12)=" "

		'14.----------EX2100--------------------
		t.Tables("CustomDeviceMap").AddRow
		t.Tables("CustomDeviceMap").Cols("Id").z(13)="14"
		t.Tables("CustomDeviceMap").Cols("Module").z(13)=Link +"dll\EX2100"
		t.Tables("CustomDeviceMap").Cols("Name").z(13)="EX2100"
		t.Tables("CustomDeviceMap").Cols("InstanceTable").z(13)="DFWIEEE421"
		t.Tables("CustomDeviceMap").Cols("Tag").z(13)="Exciter"
		t.Tables("CustomDeviceMap").Cols("LinkList").z(13)="Generator"
		t.Tables("CustomDeviceMap").Cols("SiblingLinkList").z(13)="DFWIEEE421PSS13"

		'15.----------EX2100br--------------------
		t.Tables("CustomDeviceMap").AddRow
		t.Tables("CustomDeviceMap").Cols("Id").z(14)="15"
		t.Tables("CustomDeviceMap").Cols("Module").z(14)=Link +"dll\EX2100br"
		t.Tables("CustomDeviceMap").Cols("Name").z(14)="EX2100br"
		t.Tables("CustomDeviceMap").Cols("InstanceTable").z(14)="DFWIEEE421"
		t.Tables("CustomDeviceMap").Cols("Tag").z(14)="Exciter"
		t.Tables("CustomDeviceMap").Cols("LinkList").z(14)="Generator"
		t.Tables("CustomDeviceMap").Cols("SiblingLinkList").z(14)=" "

		'16.----------K0SUR2--------------------
		t.Tables("CustomDeviceMap").AddRow
		t.Tables("CustomDeviceMap").Cols("Id").z(15)="16"
		t.Tables("CustomDeviceMap").Cols("Module").z(15)=Link +"dll\K0SUR2"
		t.Tables("CustomDeviceMap").Cols("Name").z(15)="K0SUR2"
		t.Tables("CustomDeviceMap").Cols("InstanceTable").z(15)="ExcControl"
		t.Tables("CustomDeviceMap").Cols("Tag").z(15)="ExcControl"
		t.Tables("CustomDeviceMap").Cols("LinkList").z(15)="Exciter"
		t.Tables("CustomDeviceMap").Cols("SiblingLinkList").z(15)=" "

		'17.----------Prismic--------------------
		t.Tables("CustomDeviceMap").AddRow
		t.Tables("CustomDeviceMap").Cols("Id").z(16)="17"
		t.Tables("CustomDeviceMap").Cols("Module").z(16)=Link +"dll\Prismic"
		t.Tables("CustomDeviceMap").Cols("Name").z(16)="Prismic"
		t.Tables("CustomDeviceMap").Cols("InstanceTable").z(16)="DFWIEEE421"
		t.Tables("CustomDeviceMap").Cols("Tag").z(16)="Exciter"
		t.Tables("CustomDeviceMap").Cols("LinkList").z(16)="Generator"
		t.Tables("CustomDeviceMap").Cols("SiblingLinkList").z(16)=" "


		'18.----------Semipol--------------------
		t.Tables("CustomDeviceMap").AddRow
		t.Tables("CustomDeviceMap").Cols("Id").z(17)="18"
		t.Tables("CustomDeviceMap").Cols("Module").z(17)=Link +"dll\Semipol"
		t.Tables("CustomDeviceMap").Cols("Name").z(17)="Semipol"
		t.Tables("CustomDeviceMap").Cols("InstanceTable").z(17)="DFWIEEE421"
		t.Tables("CustomDeviceMap").Cols("Tag").z(17)="Exciter"
		t.Tables("CustomDeviceMap").Cols("LinkList").z(17)="Generator"
		t.Tables("CustomDeviceMap").Cols("SiblingLinkList").z(17)="DFWIEEE421PSS13"

		'19.----------Semipol_PSS--------------------
		t.Tables("CustomDeviceMap").AddRow
		t.Tables("CustomDeviceMap").Cols("Id").z(18)="19"
		t.Tables("CustomDeviceMap").Cols("Module").z(18)=Link +"dll\Semipol"
		t.Tables("CustomDeviceMap").Cols("Name").z(18)="Semipol_PSS"
		t.Tables("CustomDeviceMap").Cols("InstanceTable").z(18)="DFWIEEE421PSS13"
		t.Tables("CustomDeviceMap").Cols("Tag").z(18)="PSS"
		t.Tables("CustomDeviceMap").Cols("LinkList").z(18)="DFWIEEE421"
		t.Tables("CustomDeviceMap").Cols("SiblingLinkList").z(18)=" "


		'20.----------THYNE_4--------------------
		t.Tables("CustomDeviceMap").AddRow
		t.Tables("CustomDeviceMap").Cols("Id").z(19)="20"
		t.Tables("CustomDeviceMap").Cols("Module").z(19)=Link +"dll\THYNE_4"
		t.Tables("CustomDeviceMap").Cols("Name").z(19)="THYNE_4"
		t.Tables("CustomDeviceMap").Cols("InstanceTable").z(19)="DFWIEEE421"
		t.Tables("CustomDeviceMap").Cols("Tag").z(19)="Exciter"
		t.Tables("CustomDeviceMap").Cols("LinkList").z(19)="Generator"
		t.Tables("CustomDeviceMap").Cols("SiblingLinkList").z(19)=" "

		'21.----------U5001--------------------
		t.Tables("CustomDeviceMap").AddRow
		t.Tables("CustomDeviceMap").Cols("Id").z(20)="21"
		t.Tables("CustomDeviceMap").Cols("Module").z(20)=Link +"dll\U5001"
		t.Tables("CustomDeviceMap").Cols("Name").z(20)="U5001"
		t.Tables("CustomDeviceMap").Cols("InstanceTable").z(20)="DFWIEEE421"
		t.Tables("CustomDeviceMap").Cols("Tag").z(20)="Exciter"
		t.Tables("CustomDeviceMap").Cols("LinkList").z(20)="Generator"
		t.Tables("CustomDeviceMap").Cols("SiblingLinkList").z(20)=" "

		'22.----------u6800--------------------
		t.Tables("CustomDeviceMap").AddRow
		t.Tables("CustomDeviceMap").Cols("Id").z(21)="22"
		t.Tables("CustomDeviceMap").Cols("Module").z(21)=Link +"dll\u6800"
		t.Tables("CustomDeviceMap").Cols("Name").z(21)="u6800"
		t.Tables("CustomDeviceMap").Cols("InstanceTable").z(21)="DFWIEEE421"
		t.Tables("CustomDeviceMap").Cols("Tag").z(21)="Exciter"
		t.Tables("CustomDeviceMap").Cols("LinkList").z(21)="Generator"
		t.Tables("CustomDeviceMap").Cols("SiblingLinkList").z(21)="DFWIEEE421PSS4B"

		'23.----------Hydrot--------------------
		t.Tables("CustomDeviceMap").AddRow
		t.Tables("CustomDeviceMap").Cols("Id").z(22)="23"
		t.Tables("CustomDeviceMap").Cols("Module").z(22)=Link +"dll\Hydrot"
		t.Tables("CustomDeviceMap").Cols("Name").z(22)="Hydrot"
		t.Tables("CustomDeviceMap").Cols("InstanceTable").z(22)="ARS"
		t.Tables("CustomDeviceMap").Cols("Tag").z(22)="ARS"
		t.Tables("CustomDeviceMap").Cols("LinkList").z(22)="Generator"
		t.Tables("CustomDeviceMap").Cols("SiblingLinkList").z(22)=" "

		'24.----------BESSCH--------------------
		t.Tables("CustomDeviceMap").AddRow
		t.Tables("CustomDeviceMap").Cols("Id").z(23)="24"
		t.Tables("CustomDeviceMap").Cols("Module").z(23)=Link +"dll\BESSCH"
		t.Tables("CustomDeviceMap").Cols("Name").z(23)="BESSCH"
		t.Tables("CustomDeviceMap").Cols("InstanceTable").z(23)="Exciter"
		t.Tables("CustomDeviceMap").Cols("Tag").z(23)="Exciter"
		t.Tables("CustomDeviceMap").Cols("LinkList").z(23)="Generator"
		t.Tables("CustomDeviceMap").Cols("SiblingLinkList").z(23)=" "

		'25.----------K0SUR2_br--------------------
		t.Tables("CustomDeviceMap").AddRow
		t.Tables("CustomDeviceMap").Cols("Id").z(24)="25"
		t.Tables("CustomDeviceMap").Cols("Module").z(24)=Link +"dll\K0SUR2_br"
		t.Tables("CustomDeviceMap").Cols("Name").z(24)="K0SUR2_br"
		t.Tables("CustomDeviceMap").Cols("InstanceTable").z(24)="ExcControl"
		t.Tables("CustomDeviceMap").Cols("Tag").z(24)="ExcControl"
		t.Tables("CustomDeviceMap").Cols("LinkList").z(24)="Exciter"
		t.Tables("CustomDeviceMap").Cols("SiblingLinkList").z(24)=" "

		'26.----------gglite--------------------
		t.Tables("CustomDeviceMap").AddRow
		t.Tables("CustomDeviceMap").Cols("Id").z(25)="26"
		t.Tables("CustomDeviceMap").Cols("Module").z(25)=Link +"dll\gglite"
		t.Tables("CustomDeviceMap").Cols("Name").z(25)="gglite"
		t.Tables("CustomDeviceMap").Cols("InstanceTable").z(25)="ARS"
		t.Tables("CustomDeviceMap").Cols("Tag").z(25)="ARS"
		t.Tables("CustomDeviceMap").Cols("LinkList").z(25)="Generator"
		t.Tables("CustomDeviceMap").Cols("SiblingLinkList").z(25)=" "

		'27.----------Alstom2--------------------
		t.Tables("CustomDeviceMap").AddRow
		t.Tables("CustomDeviceMap").Cols("Id").z(26)="27"
		t.Tables("CustomDeviceMap").Cols("Module").z(26)=Link +"dll\Alstom2"
		t.Tables("CustomDeviceMap").Cols("Name").z(26)="Alstom2"
		t.Tables("CustomDeviceMap").Cols("InstanceTable").z(26)="DFWIEEE421"
		t.Tables("CustomDeviceMap").Cols("Tag").z(26)="Exciter"
		t.Tables("CustomDeviceMap").Cols("LinkList").z(26)="Generator"
		t.Tables("CustomDeviceMap").Cols("SiblingLinkList").z(26)="DFWIEEE421PSS13"

		'28.----------Alstom2_PSS--------------------
		t.Tables("CustomDeviceMap").AddRow
		t.Tables("CustomDeviceMap").Cols("Id").z(27)="28"
		t.Tables("CustomDeviceMap").Cols("Module").z(27)=Link +"dll\Alstom2_PSS"
		t.Tables("CustomDeviceMap").Cols("Name").z(27)="Alstom2_PSS"
		t.Tables("CustomDeviceMap").Cols("InstanceTable").z(27)="DFWIEEE421PSS13"
		t.Tables("CustomDeviceMap").Cols("Tag").z(27)="PSS"
		t.Tables("CustomDeviceMap").Cols("LinkList").z(27)="DFWIEEE421"
		t.Tables("CustomDeviceMap").Cols("SiblingLinkList").z(27)=" "

		'29.----------Thyripol--------------------
		t.Tables("CustomDeviceMap").AddRow
		t.Tables("CustomDeviceMap").Cols("Id").z(28)="29"
		t.Tables("CustomDeviceMap").Cols("Module").z(28)=Link +"dll\Thyripol"
		t.Tables("CustomDeviceMap").Cols("Name").z(28)="Thyripol"
		t.Tables("CustomDeviceMap").Cols("InstanceTable").z(28)="DFWIEEE421"
		t.Tables("CustomDeviceMap").Cols("Tag").z(28)="Exciter"
		t.Tables("CustomDeviceMap").Cols("LinkList").z(28)="Generator"
		t.Tables("CustomDeviceMap").Cols("SiblingLinkList").z(28)="DFWIEEE421PSS13"

		'30.----------Thyripol PSS--------------------
		t.Tables("CustomDeviceMap").AddRow
		t.Tables("CustomDeviceMap").Cols("Id").z(29)="30"
		t.Tables("CustomDeviceMap").Cols("Module").z(29)=Link +"dll\ThyrPSS"
		t.Tables("CustomDeviceMap").Cols("Name").z(29)="ThyrPSS"
		t.Tables("CustomDeviceMap").Cols("InstanceTable").z(29)="DFWIEEE421PSS13"
		t.Tables("CustomDeviceMap").Cols("Tag").z(29)="PSS"
		t.Tables("CustomDeviceMap").Cols("LinkList").z(29)="DFWIEEE421"
		t.Tables("CustomDeviceMap").Cols("SiblingLinkList").z(29)=" "
	End If
	'-----------------------------------------------------------------------------------------
	ExcControlParam = 1
	If ExcControlParam = 1 Then
		'--------------���� �������
		t.Tables("ExcControlParam").delrows
		t.Tables("ExcControlParam").AddRow
		t.Tables("ExcControlParam").Cols("Id").z(0)=2
		t.Tables("ExcControlParam").Cols("Param").z(0)=0
		t.Tables("ExcControlParam").Cols("Value").z(0)=0.72

		t.Tables("ExcControlParam").AddRow
		t.Tables("ExcControlParam").Cols("Id").z(1)=2
		t.Tables("ExcControlParam").Cols("Param").z(1)=1
		t.Tables("ExcControlParam").Cols("Value").z(1)=0.2

		t.Tables("ExcControlParam").AddRow
		t.Tables("ExcControlParam").Cols("Id").z(2)=2
		t.Tables("ExcControlParam").Cols("Param").z(2)=2
		t.Tables("ExcControlParam").Cols("Value").z(2)=1.3

		t.Tables("ExcControlParam").AddRow
		t.Tables("ExcControlParam").Cols("Id").z(3)=2
		t.Tables("ExcControlParam").Cols("Param").z(3)=3
		t.Tables("ExcControlParam").Cols("Value").z(3)=0.5
		'---------------------------
		t.Tables("ExcControlParam").AddRow
		t.Tables("ExcControlParam").Cols("Id").z(4)=3
		t.Tables("ExcControlParam").Cols("Param").z(4)=0
		t.Tables("ExcControlParam").Cols("Value").z(4)=0.72

		t.Tables("ExcControlParam").AddRow
		t.Tables("ExcControlParam").Cols("Id").z(5)=3
		t.Tables("ExcControlParam").Cols("Param").z(5)=1
		t.Tables("ExcControlParam").Cols("Value").z(5)=0.2

		t.Tables("ExcControlParam").AddRow
		t.Tables("ExcControlParam").Cols("Id").z(6)=3
		t.Tables("ExcControlParam").Cols("Param").z(6)=2
		t.Tables("ExcControlParam").Cols("Value").z(6)=1.3

		t.Tables("ExcControlParam").AddRow
		t.Tables("ExcControlParam").Cols("Id").z(7)=3
		t.Tables("ExcControlParam").Cols("Param").z(7)=3
		t.Tables("ExcControlParam").Cols("Value").z(7)=0.5
		'---------------------------
		t.Tables("ExcControlParam").AddRow
		t.Tables("ExcControlParam").Cols("Id").z(8)=4
		t.Tables("ExcControlParam").Cols("Param").z(8)=0
		t.Tables("ExcControlParam").Cols("Value").z(8)=0.1

		t.Tables("ExcControlParam").AddRow
		t.Tables("ExcControlParam").Cols("Id").z(9)=4
		t.Tables("ExcControlParam").Cols("Param").z(9)=1
		t.Tables("ExcControlParam").Cols("Value").z(9)=0.1

		t.Tables("ExcControlParam").AddRow
		t.Tables("ExcControlParam").Cols("Id").z(10)=4
		t.Tables("ExcControlParam").Cols("Param").z(10)=2
		t.Tables("ExcControlParam").Cols("Value").z(10)=0.1

		t.Tables("ExcControlParam").AddRow
		t.Tables("ExcControlParam").Cols("Id").z(11)=4
		t.Tables("ExcControlParam").Cols("Param").z(11)=3
		t.Tables("ExcControlParam").Cols("Value").z(11)=0.1
		'---------------------------
		t.Tables("ExcControlParam").AddRow
		t.Tables("ExcControlParam").Cols("Id").z(12)=5
		t.Tables("ExcControlParam").Cols("Param").z(12)=0
		t.Tables("ExcControlParam").Cols("Value").z(12)=0.1

		t.Tables("ExcControlParam").AddRow
		t.Tables("ExcControlParam").Cols("Id").z(13)=5
		t.Tables("ExcControlParam").Cols("Param").z(13)=1
		t.Tables("ExcControlParam").Cols("Value").z(13)=0.1

		t.Tables("ExcControlParam").AddRow
		t.Tables("ExcControlParam").Cols("Id").z(14)=5
		t.Tables("ExcControlParam").Cols("Param").z(14)=2
		t.Tables("ExcControlParam").Cols("Value").z(14)=0.1

		t.Tables("ExcControlParam").AddRow
		t.Tables("ExcControlParam").Cols("Id").z(15)=5
		t.Tables("ExcControlParam").Cols("Param").z(15)=3
		t.Tables("ExcControlParam").Cols("Value").z(15)=0.1

		'---------------------------

		t.Tables("ExcControlParam").AddRow
		t.Tables("ExcControlParam").Cols("Id").z(16)=6
		t.Tables("ExcControlParam").Cols("Param").z(16)=0
		t.Tables("ExcControlParam").Cols("Value").z(16)=1

		t.Tables("ExcControlParam").AddRow
		t.Tables("ExcControlParam").Cols("Id").z(17)=6
		t.Tables("ExcControlParam").Cols("Param").z(17)=1
		t.Tables("ExcControlParam").Cols("Value").z(17)=1

		t.Tables("ExcControlParam").AddRow
		t.Tables("ExcControlParam").Cols("Id").z(18)=6
		t.Tables("ExcControlParam").Cols("Param").z(18)=2
		t.Tables("ExcControlParam").Cols("Value").z(18)=1

		t.Tables("ExcControlParam").AddRow
		t.Tables("ExcControlParam").Cols("Id").z(19)=6
		t.Tables("ExcControlParam").Cols("Param").z(19)=3
		t.Tables("ExcControlParam").Cols("Value").z(19)=1
		'----------------------------------------------------
	End If
End Function