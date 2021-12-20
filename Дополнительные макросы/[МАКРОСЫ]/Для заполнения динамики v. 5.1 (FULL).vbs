r=setlocale("en-us")
rrr=1

set t=RASTR

set vet=t.tables("vetv")
	set uzl=t.tables("node")
	set gen=t.Tables("Generator")
	set vozb=t.Tables("Exciter")
	set arv=t.Tables("ExcControl")
	set vieee=t.Tables("DFWIEEE421")
	set stieee=t.Tables("DFWIEEE421PSS13")
	set pss4	=t.Tables("DFWIEEE421PSS4B")
	set ars=Rastr.Tables("ARS")
	set forc=Rastr.Tables("Forcer")
	set omv=t.Tables("DFW421UEL")
	set bor=t.Tables("DFWOELUNITROL")
	set FuncPQ=t.Tables("FuncPQ")
	set Governor=t.Tables("Governor")
	set decs400=t.Tables("DFWDECS400")
	set Thyne=t.Tables("DFWTHYNE14")

	set nsta	=gen.Cols("sta")
	set numg	=gen.Cols("Num")
	set nameg	=gen.Cols("Name")
	set gNumBrand = gen.Cols("NumBrand")
	set nodeg	=gen.Cols("Node")
	set ExciterId=gen.Cols("ExciterId")
	set ARSId=gen.Cols("ARSId")
	set pgen	=gen.Cols("P")
	set ModelType	=gen.Cols("ModelType")
	set pnom	=gen.Cols("Pnom")
	set unom	=gen.Cols("Ugnom")
	set cosfi	=gen.Cols("cosFi")
	set Demp	=gen.Cols("Demp")
	set mj		=gen.Cols("Mj")
	set xd1		=gen.Cols("xd1")
	set qmin	=gen.Cols("Qmin")
	set qmax	=gen.Cols("Qmax")
	't nodeld	=gen.Cols("nodeld")
	set xd1	=gen.Cols("xd1")
	set xd	=gen.Cols("xd")
	set xq	=gen.Cols("xq")
	set xd2	=gen.Cols("xd2")
	set xq2	=gen.Cols("xq2")
	set td01	=gen.Cols("td01")
	set td02	=gen.Cols("td02")
	set tq02	=gen.Cols("tq02")
	set xq1	=gen.Cols("xq1")
	set xl	=gen.Cols("xl")
	set x2	=gen.Cols("x2")
	set x2	=gen.Cols("x0")
	set tq01	=gen.Cols("tq01")
	set gIVActuatorId	=gen.Cols("IVActuatorId")

	set ny		=uzl.Cols("ny")
	set name	=uzl.Cols("name")
	set pg		=uzl.Cols("pg") 
	set tip		=uzl.Cols("tip") 
	set nqmin	=uzl.Cols("qmin")
	set nqmax	=uzl.Cols("qmax")
	set uhom	=uzl.Cols("uhom")

	set idv	=vozb.Cols("Id")
	set namev	=vozb.Cols("Name")
	set ModelTypev	=vozb.Cols("ModelType")
	set Brandv	=vozb.Cols("Brand")
	set ExcControlIdv	=vozb.Cols("ExcControlId")
	set ForcerIdv	=vozb.Cols("ForcerId")
	set Texcv	=vozb.Cols("Texc")
	set Kig=vozb.Cols("Kig")
	set Kif=vozb.Cols("Kif")
	set Uf_min=vozb.Cols("Uf_min")
	set Uf_max=vozb.Cols("Uf_max")
	set If_min=vozb.Cols("If_min")
	set If_max=vozb.Cols("If_max")
	set Type_rgv=vozb.Cols("Type_rg")
	set vozbCustomModel=vozb.Cols("CustomModel")
	set vozbKarv=vozb.Cols("Karv")
	set vozbT3exc=vozb.Cols("T3exc")

	set ida	=arv.Cols("Id")
	set namea	=arv.Cols("Name")
	set ModelTypea	=arv.Cols("ModelType")
	set Branda	=arv.Cols("Brand")
	set Trv	=arv.Cols("Trv")
	set Ku	=arv.Cols("Ku")
	set Ku1	=arv.Cols("Ku1")
	set Kif1	=arv.Cols("Kif1")
	set Kf	=arv.Cols("Kf")
	set Kf1	=arv.Cols("Kf1")
	set Tf	=arv.Cols("Tf")
	set Urv_min	=arv.Cols("Urv_min")
	set Urv_max	=arv.Cols("Urv_max")
	set Alpha	=arv.Cols("Alpha")
	set arvCustomModel	=arv.Cols("CustomModel")
	set arvTINT	=arv.Cols("TINT")

	set arsids	=ars.Cols("Id")
	set arsname	=ars.Cols("Name")
	set arsModelTypes	=ars.Cols("ModelType")
	set arsCustomModel	=ars.Cols("ModelType")
	set arsBrands	=ars.Cols("Brand")
	set arsGovernorId	=ars.Cols("GovernorId")
	set arsao	=ars.Cols("ao")
	set arsaz	=ars.Cols("az")
	set arsotmin	=ars.Cols("otmin")
	set arsotmax	=ars.Cols("otmax")
	set arsstrs	=ars.Cols("strs")
	set arszn	=ars.Cols("zn")
	set arsdpo	=ars.Cols("dpo")
	set arsThp	=ars.Cols("Thp")
	set arsTrlp	=ars.Cols("Trlp")
	set arsTw	=ars.Cols("Tw")
	set arspt	=ars.Cols("pt")
	set arsMu	=ars.Cols("Mu")
	set arsPsteam	=ars.Cols("Psteam")
	set arsCustomModel	=ars.Cols("CustomModel")

	set idf	=forc.Cols("Id")
	set namef	=forc.Cols("Name")
	set ModelTypef	=forc.Cols("ModelType")
	set Ubf	=forc.Cols("Ubf")
	set Uef	=forc.Cols("Uef")
	set Rf	=forc.Cols("Rf")
	set Texc_f	=forc.Cols("Texc_f")
	set Tz_in	=forc.Cols("Tz_in")
	set Tz_out	=forc.Cols("Tz_out")
	set Ubrf	=forc.Cols("Ubrf")
	set Uerf	=forc.Cols("Uerf")
	set Rrf	=forc.Cols("Rrf")
	set Texc_rf	=forc.Cols("Texc_rf")

	set ststa=stieee.Cols("sta")
	set stId=stieee.Cols("Id")
	set stName=stieee.Cols("Name")
	set stModel=stieee.Cols("ModelType")
	set stBrand=stieee.Cols("Brand")
	set stCustomModel=stieee.Cols("CustomModel")
	set stInput1Type=stieee.Cols("Input1Type")
	set stInput2Type=stieee.Cols("Input2Type")
	set stVstmin=stieee.Cols("Vstmin")
	set stVstmax=stieee.Cols("Vstmax")
	set stKs1=stieee.Cols("Ks1")
	set stT1=stieee.Cols("T1")
	set stT2=stieee.Cols("T2")
	set stT3=stieee.Cols("T3")
	set stT4=stieee.Cols("T4")
	set stT5=stieee.Cols("T5")
	set stT6=stieee.Cols("T6")
	set stT7=stieee.Cols("T7")
	set stT8=stieee.Cols("T8")
	set stT9=stieee.Cols("T9")
	set stT10=stieee.Cols("T10")
	set stT11=stieee.Cols("T11")
	set stA1=stieee.Cols("A1")
	set stA2=stieee.Cols("A2")
	set stA3=stieee.Cols("A3")
	set stA4=stieee.Cols("A4")
	set stA5=stieee.Cols("A5")
	set stA6=stieee.Cols("A6")
	set stA7=stieee.Cols("A7")
	set stA8=stieee.Cols("A8")
	set stKs2=stieee.Cols("Ks2")
	set stKs3=stieee.Cols("Ks3")
	set stTw1=stieee.Cols("Tw1")
	set stTw2=stieee.Cols("Tw2")
	set stTw3=stieee.Cols("Tw3")
	set stTw4=stieee.Cols("Tw4")
	set stM=stieee.Cols("M")
	set stN=stieee.Cols("N")
	set stVsi1min=stieee.Cols("Vsi1min")
	set stVsi1max=stieee.Cols("Vsi1max")
	set stVsi2min=stieee.Cols("Vsi2min")
	set stVsi2max=stieee.Cols("Vsi2max")

	set vista=vieee.Cols("sta")
	set viId=vieee.Cols("Id")
	set viName=vieee.Cols("Name")
	set viModel=vieee.Cols("ModelType")
	set viBrand=vieee.Cols("Brand")
	set viCustomModel=vieee.Cols("CustomModel")
	set viUELId=vieee.Cols("UELId")
	set viUELPos=vieee.Cols("UELPos")
	set viOELId=vieee.Cols("OELId")
	set viOELPos=vieee.Cols("OELPos")
	set viPSSId=vieee.Cols("PSSId")
	set viPSSPos=vieee.Cols("PSSPos")
	set viTe=vieee.Cols("Te")
	set viKe=vieee.Cols("Ke")
	set viSe1=vieee.Cols("Se1")
	set viEfd1=vieee.Cols("Efd1")
	set viVe1=vieee.Cols("Ve1")
	set viSe2=vieee.Cols("Se2")
	set viEfd2=vieee.Cols("Efd2")
	set viVe2=vieee.Cols("Ve2")
	set viVemin=vieee.Cols("Vemin")
	set viVrmin=vieee.Cols("Vrmin")
	set viVrmax=vieee.Cols("Vrmax")
	set viKa=vieee.Cols("Ka")
	set viTa=vieee.Cols("Ta")
	set viTf=vieee.Cols("Tf")
	set viKf=vieee.Cols("Kf")
	set viTc=vieee.Cols("Tc")
	set viTb=vieee.Cols("Tb")
	set viKv=vieee.Cols("Kv")
	set viTrh=vieee.Cols("Trh")
	set viKpr=vieee.Cols("Kpr")
	set viKir=vieee.Cols("Kir")
	set viKdr=vieee.Cols("Kdr")
	set viTdr=vieee.Cols("Tdr")
	set viKc=vieee.Cols("Kc")
	set viKd=vieee.Cols("Kd")
	set viVfemax=vieee.Cols("Vfemax")
	set viVamin=vieee.Cols("Vamin")
	set viVamax=vieee.Cols("Vamax")
	set viKb=vieee.Cols("Kb")
	set viKh=vieee.Cols("Kh")
	set viKr=vieee.Cols("Kr")
	set viKn=vieee.Cols("Kn")
	set viEfdn=vieee.Cols("Efdn")
	set viKlv=vieee.Cols("Klv")
	set viVlv=vieee.Cols("Vlv")
	set viVimin=vieee.Cols("Vimin")
	set viVimax=vieee.Cols("Vimax")
	set viTf2=vieee.Cols("Tf2")
	set viTf3=vieee.Cols("Tf3")
	set viTk=vieee.Cols("Tk")
	set viTj=vieee.Cols("Tj")
	set viTh=vieee.Cols("Th")
	set viVhmax=vieee.Cols("Vhmax")
	set viVfelim=vieee.Cols("Vfelim")
	set viKp=vieee.Cols("Kp")
	set viKpa=vieee.Cols("Kpa")
	set viKia=vieee.Cols("Kia")
	set viKf1=vieee.Cols("Kf1")
	set viKf2=vieee.Cols("Kf2")
	set viKl=vieee.Cols("Kl")
	set viTb1=vieee.Cols("Tb1")
	set viTc1=vieee.Cols("Tc1")
	set viKlr=vieee.Cols("Klr")
	set viIlr=vieee.Cols("Ilr")
	set viKi=vieee.Cols("Ki")
	set viTheta=vieee.Cols("Theta")
	set viVmmin=vieee.Cols("Vmmin")
	set viVmmax=vieee.Cols("Vmmax")
	set viKg=vieee.Cols("Kg")
	set viVBmax=vieee.Cols("VBmax")
	set viVGmax=vieee.Cols("VGmax")
	set viXl=vieee.Cols("Xl")
	set viKm=vieee.Cols("Km")
	set viTm=vieee.Cols("Tm")
	set viTb2=vieee.Cols("Tb2")
	set viTc2=vieee.Cols("Tc2")
	set viTub1=vieee.Cols("Tub1")
	set viTuc1=vieee.Cols("Tuc1")
	set viTub2=vieee.Cols("Tub2")
	set viTuc2=vieee.Cols("Tuc2")
	set viTob1=vieee.Cols("Tob1")
	set viToc1=vieee.Cols("Toc1")
	set viTob2=vieee.Cols("Tob2")
	set viToc2=vieee.Cols("Toc2")

	set viAex=vieee.Cols("Aex")
	set viBex=vieee.Cols("Bex")
	set viKcf=vieee.Cols("Kcf")
	set viKhf=vieee.Cols("Khf")
	set viKif=vieee.Cols("Kif")
	set viSamovozb=vieee.Cols("Samovozb")
	set viTr=vieee.Cols("Tr")

	set pss4sta=pss4.Cols("sta")
	set pss4Id=pss4.Cols("Id")
	set pss4Name=pss4.Cols("Name")
	set pss4ModelType=pss4.Cols("ModelType")
	set pss4Brand=pss4.Cols("Brand")
	set pss4CustomModel=pss4.Cols("CustomModel")
	set pss4Input1Type=pss4.Cols("Input1Type")
	set pss4Input2Type=pss4.Cols("Input2Type")
	set pss4MBPSS1=pss4.Cols("MBPSS1")
	set pss4MBPSS2=pss4.Cols("MBPSS2")
	set pss4Vstmin=pss4.Cols("Vstmin")
	set pss4Vstmax=pss4.Cols("Vstmax")
	set pss4KL1=pss4.Cols("KL1")
	set pss4KL2=pss4.Cols("KL2")
	set pss4KL11=pss4.Cols("KL11")
	set pss4KL17=pss4.Cols("KL17")
	set pss4TL1=pss4.Cols("TL1")
	set pss4TL2=pss4.Cols("TL2")
	set pss4TL3=pss4.Cols("TL3")
	set pss4TL4=pss4.Cols("TL4")
	set pss4TL5=pss4.Cols("TL5")
	set pss4TL6=pss4.Cols("TL6")
	set pss4TL7=pss4.Cols("TL7")
	set pss4TL8=pss4.Cols("TL8")
	set pss4TL9=pss4.Cols("TL9")
	set pss4TL10=pss4.Cols("TL10")
	set pss4TL11=pss4.Cols("TL11")
	set pss4TL12=pss4.Cols("TL12")
	set pss4KL=pss4.Cols("KL")
	set pss4Vlmin=pss4.Cols("Vlmin")
	set pss4Vlmax=pss4.Cols("Vlmax")
	set pss4KI1=pss4.Cols("KI1")
	set pss4KI2=pss4.Cols("KI2")
	set pss4KI11=pss4.Cols("KI11")
	set pss4KI17=pss4.Cols("KI17")
	set pss4TI1=pss4.Cols("TI1")
	set pss4TI2=pss4.Cols("TI2")
	set pss4TI3=pss4.Cols("TI3")
	set pss4TI4=pss4.Cols("TI4")
	set pss4TI5=pss4.Cols("TI5")
	set pss4TI6=pss4.Cols("TI6")
	set pss4TI7=pss4.Cols("TI7")
	set pss4TI8=pss4.Cols("TI8")
	set pss4TI9=pss4.Cols("TI9")
	set pss4TI10=pss4.Cols("TI10")
	set pss4TI11=pss4.Cols("TI11")
	set pss4TI12=pss4.Cols("TI12")
	set pss4KI=pss4.Cols("KI")
	set pss4Vimin=pss4.Cols("Vimin")
	set pss4Vimax=pss4.Cols("Vimax")
	set pss4KH1=pss4.Cols("KH1")
	set pss4KH2=pss4.Cols("KH2")
	set pss4KH11=pss4.Cols("KH11")
	set pss4KH17=pss4.Cols("KH17")
	set pss4TH1=pss4.Cols("TH1")
	set pss4TH2=pss4.Cols("TH2")
	set pss4TH3=pss4.Cols("TH3")
	set pss4TH4=pss4.Cols("TH4")
	set pss4TH5=pss4.Cols("TH5")
	set pss4TH6=pss4.Cols("TH6")
	set pss4TH7=pss4.Cols("TH7")
	set pss4TH8=pss4.Cols("TH8")
	set pss4TH9=pss4.Cols("TH9")
	set pss4TH10=pss4.Cols("TH10")
	set pss4TH11=pss4.Cols("TH11")
	set pss4TH12=pss4.Cols("TH12")
	set pss4KH=pss4.Cols("KH")
	set pss4Vhmin=pss4.Cols("Vhmin")
	set pss4Vhmax=pss4.Cols("Vhmax")
	set pss4sta=pss4.Cols("sta")

	set omvsta=omv.Cols("sta")
	set omvId=omv.Cols("Id")
	set omvName=omv.Cols("Name")
	set omvModelType=omv.Cols("ModelType")
	set omvBrand=omv.Cols("Brand")
	set omvCustomModel=omv.Cols("CustomModel")
	set omvTu1=omv.Cols("Tu1")
	set omvTu2=omv.Cols("Tu2")
	set omvTu3=omv.Cols("Tu3")
	set omvTu4=omv.Cols("Tu4")
	set omvVulmin=omv.Cols("Vulmin")
	set omvVulmax=omv.Cols("Vulmax")
	set omvKul=omv.Cols("Kul")
	set omvKui=omv.Cols("Kui")
	set omvVuimin=omv.Cols("Vuimin")
	set omvVuimax=omv.Cols("Vuimax")
	set omvKuf=omv.Cols("Kuf")
	set omvTuf=omv.Cols("Tuf")
	set omvKur=omv.Cols("Kur")
	set omvKuc=omv.Cols("Kuc")
	set omvVurmax=omv.Cols("Vurmax")
	set omvVucmax=omv.Cols("Vucmax")
	set omvTuV=omv.Cols("TuV")
	set omvTuP=omv.Cols("TuP")
	set omvTuQ=omv.Cols("TuQ")
	set omvK1=omv.Cols("K1")
	set omvK2=omv.Cols("K2")
	set omvDependency_F1=omv.Cols("Dependency_F1")
	set omvOutput=omv.Cols("Output")
	set omvKl=omv.Cols("Kl")

	set borsta=bor.Cols("sta")
	set borId=bor.Cols("Id")
	set borName=bor.Cols("Name")
	set borModelType=bor.Cols("ModelType")
	set borBrand=bor.Cols("Brand")
	set borCustomModel=bor.Cols("CustomModel")
	set borIfMax=bor.Cols("IfMax")
	set borIfth=bor.Cols("Ifth")
	set borKexpIf=bor.Cols("KexpIf")
	set borKR3=bor.Cols("KR3")
	set borKR3i=bor.Cols("KR3i")
	set borTc23=bor.Cols("Tc23")
	set borTb23=bor.Cols("Tb23")
	set borTc13=bor.Cols("Tc13")
	set borTb13=bor.Cols("Tb13")
	set borVamin=bor.Cols("Vamin")
	set borVamax=bor.Cols("Vamax")
	set borTdel=bor.Cols("Tdel")
	set borKth=bor.Cols("Kth")
	set borKToF=bor.Cols("KToF")
	set borKcF=bor.Cols("KcF")
	set borKhF=bor.Cols("KhF")
	set borTRFout=bor.Cols("TRFout")
	set borTr=bor.Cols("Tr")
	set borOutput=bor.Cols("Output")
	set borKl=bor.Cols("Kl")

	set FuncPQId=FuncPQ.Cols("Id")
	set FuncPQP=FuncPQ.Cols("P")
	set FuncPQQ=FuncPQ.Cols("Q")

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

	set decs400sta=decs400.Cols("sta")
	set decs400Id=decs400.Cols("Id")
	set decs400Name=decs400.Cols("Name")
	set decs400ModelType=decs400.Cols("ModelType")
	set decs400Brand=decs400.Cols("Brand")
	set decs400CustomModel=decs400.Cols("CustomModel")
	set decs400PSSId=decs400.Cols("PSSId")
	set decs400UELId=decs400.Cols("UELId")
	set decs400OELId=decs400.Cols("OELId")
	set decs400Xl=decs400.Cols("Xl")
	set decs400DRP=decs400.Cols("DRP")
	set decs400VrMin=decs400.Cols("VrMin")
	set decs400VrMax=decs400.Cols("VrMax")
	set decs400VmMin=decs400.Cols("VmMin")
	set decs400VmMax=decs400.Cols("VmMax")
	set decs400VbMax=decs400.Cols("VbMax")
	set decs400Kc=decs400.Cols("Kc")
	set decs400Kp=decs400.Cols("Kp")
	set decs400Kpm=decs400.Cols("Kpm")
	set decs400Kpr=decs400.Cols("Kpr")
	set decs400Kir=decs400.Cols("Kir")
	set decs400Kpd=decs400.Cols("Kpd")
	set decs400Ta=decs400.Cols("Ta")
	set decs400Td=decs400.Cols("Td")
	set decs400Tr=decs400.Cols("Tr")
	set decs400SelfExc=decs400.Cols("SelfExc")
	set decs400Del=decs400.Cols("Del")

	set Thynesta=Thyne.Cols("sta")
	set ThyneId=Thyne.Cols("Id")
	set ThyneName=Thyne.Cols("Name")
	set ThyneModelType=Thyne.Cols("ModelType")
	set ThyneBrand=Thyne.Cols("Brand")
	set ThyneCustomModel=Thyne.Cols("CustomModel")
	set ThyneUELId=Thyne.Cols("UELId")
	set ThynePSSId=Thyne.Cols("PSSId")
	set ThyneAex=Thyne.Cols("Aex")
	set ThyneBex=Thyne.Cols("Bex")
	set ThyneAlpha=Thyne.Cols("Alpha")
	set ThyneBeta=Thyne.Cols("Beta")
	set ThyneIfdMin=Thyne.Cols("IfdMin")
	set ThyneKc=Thyne.Cols("Kc")
	set ThyneKd1=Thyne.Cols("Kd1")
	set ThyneKd2=Thyne.Cols("Kd2")
	set ThyneKe=Thyne.Cols("Ke")
	set ThyneKetb=Thyne.Cols("Ketb")
	set ThyneKh=Thyne.Cols("Kh")
	set ThyneKp1=Thyne.Cols("Kp1")
	set ThyneKp2=Thyne.Cols("Kp2")
	set ThyneKp3=Thyne.Cols("Kp3")
	set ThyneTd1=Thyne.Cols("Td1")
	set ThyneTe1=Thyne.Cols("Te1")
	set ThyneTe2=Thyne.Cols("Te2")
	set ThyneTi1=Thyne.Cols("Ti1")
	set ThyneTi2=Thyne.Cols("Ti2")
	set ThyneTi3=Thyne.Cols("Ti3")
	set ThyneTr1=Thyne.Cols("Tr1")
	set ThyneTr2=Thyne.Cols("Tr2")
	set ThyneTr3=Thyne.Cols("Tr3")
	set ThyneTr4=Thyne.Cols("Tr4")
	set ThyneVO1Min=Thyne.Cols("VO1Min")
	set ThyneVO1Max=Thyne.Cols("VO1Max")
	set ThyneVO2Min=Thyne.Cols("VO2Min")
	set ThyneVO2Max=Thyne.Cols("VO2Max")
	set ThyneVO3Min=Thyne.Cols("VO3Min")
	set ThyneVO3Max=Thyne.Cols("VO3Max")
	set ThyneVD1Min=Thyne.Cols("VD1Min")
	set ThyneVD1Max=Thyne.Cols("VD1Max")
	set ThyneVI1Min=Thyne.Cols("VI1Min")
	set ThyneVI1Max=Thyne.Cols("VI1Max")
	set ThyneVI2Min=Thyne.Cols("VI2Min")
	set ThyneVI2Max=Thyne.Cols("VI2Max")
	set ThyneVI3Min=Thyne.Cols("VI3Min")
	set ThyneVI3Max=Thyne.Cols("VI3Max")
	set ThyneVP1Min=Thyne.Cols("VP1Min")
	set ThyneVP1Max=Thyne.Cols("VP1Max")
	set ThyneVP2Min=Thyne.Cols("VP2Min")
	set ThyneVP2Max=Thyne.Cols("VP2Max")
	set ThyneVP3Min=Thyne.Cols("VP3Min")
	set ThyneVP3Max=Thyne.Cols("VP3Max")
	set ThyneVrMin=Thyne.Cols("VrMin")
	set ThyneVrMax=Thyne.Cols("VrMax")
	set ThyneXp=Thyne.Cols("Xp")

'---------------------------------------------------------------------------------------------------------
'Задать ссылку для пользовательских устройств

Link_Custom_Models = "C:\CustomModels\"

SettingsFile="P:\RASTRWIN\16-РЗА\02- Расчеты\ВЛ 750 кВ Смоленская АЭС - Новобрянская\Режимы\дин_набор21.xlsx"

Set ExcelSet = CreateObject("Excel.Application")	
	ExcelSet.Workbooks.open SettingsFile
	ExcelSet.Visible = 1

Set Ex1=ExcelSet.Worksheets("1")
Set Ex2=ExcelSet.Worksheets("2")
Set Ex3=ExcelSet.Worksheets("3")
Set Ex4=ExcelSet.Worksheets("4")
Set Ex5=ExcelSet.Worksheets("5")
Set Ex6=ExcelSet.Worksheets("6")
Set Ex7=ExcelSet.Worksheets("7")
Set Ex8=ExcelSet.Worksheets("8")
Set Ex9=ExcelSet.Worksheets("9")
Set Ex10=ExcelSet.Worksheets("10")
Set Ex11=ExcelSet.Worksheets("11")
Set Ex12=ExcelSet.Worksheets("12")
'Set Ex13=ExcelSet.Worksheets("13")
'Set Ex14=ExcelSet.Worksheets("14")

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
if ffff = 8999 then
	gener = 1
	if gener = 1 then
		i = 3 ' начало с 3-й строки
		while Ex1.cells(i,1).value > 0 ' до тех пор пока в 1-м столбце (опеределение заполнен ли регион) Ex1 стоит 1
			if Ex1.cells(i,3).value > 0 then ' если Nагр больше 0 то ...
				Eny = Ex1.cells(i,3) ' Nагр
				Eny2 = Ex1.cells(i,5) ' Nузла
				Name_gen = Ex1.cells(i,4) ' Название генератора 
				'gen.SetSel("Num="& Eny & "& Node=" &Eny2)
				'gen.SetSel("Name="& Name_gen )
				Id_generator = 0 
				gen.SetSel("") 
				j = gen.FindNextSel(-1)
					while j<>-1 ' до тех пор пока j не равно -1
						'If InStr(gen.cols("Name").z(j),Name_gen) then
						If gen.cols("Name").z(j) = Name_gen then ' если название генератора  из Rastr равно назв ген из Excel						
							Id_generator = gen.cols("Num").z(j) ' к  переменной Id_generator присваевается Num генераторва из RASTR
						end if
						j = gen.FindNextSel(j)
					wend
				gen.SetSel("Num=" & Id_generator) ' выборка по присвоенному значению к Id_generator
				j = gen.FindNextSel(-1) ' находит номер строки в RASTR
				if j<>-1 then '  если j не равно -1 (т.е. такой ген есть в RASTR)
					'nameg.Z(j)=Ex1.cells(i,4)
					'  из Excel переносятся значения параметров для генератора
					ModelType.Z(j) = Ex1.cells(i,6) 
					gNumBrand.Z(j)=Ex1.cells(i,8)
					ExciterId.Z(j)=Ex1.cells(i,9)
					ARSId.Z(j)=Ex1.cells(i,10)
					gIVActuatorId.Z(j)=Ex1.cells(i,11)								
					'napgen=Rastr.Calc("val","node","na","ny="&nodeg.Z(j))
					korrPGgen = 0
					if korrPGgen = 1 then
						If pgen.Z(j) > 1.05*Ex1.cells(i,14) and pgen.Z(j) > 0 and (napgen > 550 or napgen < 500) then
							Ex1.cells(i,14)=pgen.Z(j)
						end if
						If pgen.Z(j) < 0.9*Ex1.cells(i,14) and pgen.Z(j) > 0.5*Ex1.cells(i,14) and pgen.Z(j) > 0 and (napgen > 550 or napgen < 500) then
							Ex1.cells(i,14) = pgen.Z(j)
						end if
					end if
					'nameg.Z(j)=Ex1.cells(i,4)
					'nodeg.Z(j)=Ex1.cells(i,5)
					'pgen.Z(j)=Ex1.cells(i,12)
					pnom.Z(j)=Ex1.cells(i,14)
					unom.Z(j)=Ex1.cells(i,15)
					cosfi.Z(j)=Ex1.cells(i,16)
					Demp.Z(j)=Ex1.cells(i,17)
					mj.Z(j)=Ex1.cells(i,18)
					xd1.Z(j)=Ex1.cells(i,19)
					xd.Z(j)=Ex1.cells(i,20)
					xq.Z(j)=Ex1.cells(i,21)
					xd2.Z(j)=Ex1.cells(i,22)
					xq2.Z(j) = Ex1.cells(i,23)
					td01.Z(j) = Ex1.cells(i,24)
					td02.Z(j) = Ex1.cells(i,25)
					tq02.Z(j) = Ex1.cells(i,26)
					xq1.Z(j) = Ex1.cells(i,27)
					xl.Z(j) = Ex1.cells(i,28)
					x2.Z(j) = Ex1.cells(i,29)
					x2.Z(j) = Ex1.cells(i,30)
					tq01.Z(j) = Ex1.cells(i,31)
					'Ex1.cells(i,45)=gen.Cols("Pmin")
					'Ex1.cells(i,46)=gen.Cols("Pmax")
					Ex1.cells(i,47) = nameg.Z(j)
				else
				end if
			end if
			i = i+1
		wend
         t.Printp("Таблица 1 'Генераторы' - загружена!")
	end if
    
	i = 3
	vozbuzd = 1
	if vozbuzd = 1 then
		while Ex2.cells(i,1).value > 0
			'Eny=Ex2.cells(i,3)
			Name_gen = Ex2.cells(i,4)
			Id_generator = 0
			gen.SetSel("")
			j = gen.FindNextSel(-1)
			while j<>-1
				If gen.cols("Name").z(j) = Name_gen then
					Id_generator = gen.cols("Num").z(j)
				end if
				j = gen.FindNextSel(j)
			wend
			if Ex2.cells(i,5).value > 0 and Id_generator > 0 then
				vozb.SetSel("")
				vozb.AddRow
				vozb.SetSel(" Id = 0")
				j2 = vozb.FindNextSel(-1)
				if j2<>-1 then
					idv.Z(j2) = Id_generator
					namev.Z(j2) = Ex2.cells(i,4)
					ModelTypev.Z(j2) = Ex2.cells(i,5)
					'Brandv.Z(j2)=Ex2.cells(i,6)
					ExcControlIdv.Z(j2) = Id_generator
					if Ex2.cells(i,8) > 0 then
						ForcerIdv.Z(j2) = Id_generator
					end if
					Texcv.Z(j2)=Ex2.cells(i,9)
					Kig.Z(j2)=Ex2.cells(i,10)
					Kif.Z(j2)=Ex2.cells(i,11)
					Uf_min.Z(j2)=Ex2.cells(i,12)
					Uf_max.Z(j2)=Ex2.cells(i,13)
					If_min.Z(j2)=Ex2.cells(i,14)
					If_max.Z(j2)=Ex2.cells(i,15)
					Type_rgv.Z(j2)=Ex2.cells(i,16)
					vozbCustomModel.Z(j2)=Ex2.cells(i,17)
					vozbKarv.Z(j2)=Ex2.cells(i,18)
					vozbT3exc.Z(j2)=Ex2.cells(i,19)
				end if
			end if
			i=i+1
		wend
         t.Printp("Таблица 2 'Возбудители (ИД)' - загружена!")
	end if
    
	arv2 = 1
	if arv2 = 1 then
		i = 3
		while Ex3.cells(i,1).value > 0
			'Eny=Ex3.cells(i,3)
			Name_gen = Ex3.cells(i,4)
			Id_generator = 0
			gen.SetSel("")
			j = gen.FindNextSel(-1)
			while j<>-1
				If gen.cols("Name").z(j) = Name_gen then
					Id_generator = gen.cols("Num").z(j)
				end if
				j=gen.FindNextSel(j)
			wend
			if Ex3.cells(i,5).value > 0 and Id_generator > 0 then
				arv.SetSel("")
				arv.AddRow
				arv.SetSel("Id = 0")
				j2 = arv.FindNextSel(-1)
				if j2<>-1 then
					ida.Z(j2)=Id_generator
					namea.Z(j2)=Ex3.cells(i,4)
					ModelTypea.Z(j2)=Ex3.cells(i,5)
					Branda.Z(j2)=Ex3.cells(i,6)
					Trv.Z(j2)=Ex3.cells(i,7)
					Ku.Z(j2)=Ex3.cells(i,8)
					Ku1.Z(j2)=Ex3.cells(i,9)
					Kif1.Z(j2)=Ex3.cells(i,10)
					Kf.Z(j2)=Ex3.cells(i,11)
					Kf1.Z(j2)=Ex3.cells(i,12)
					Tf.Z(j2)=Ex3.cells(i,13)
					Urv_min.Z(j2)=Ex3.cells(i,14)
					Urv_max.Z(j2)=Ex3.cells(i,15)
					Alpha.Z(j2)=Ex3.cells(i,16)
					arvCustomModel.Z(j2)=Ex3.cells(i,17)
					arvTINT.Z(j2)=Ex3.cells(i,18)
				end if
			end if
			i=i+1
		wend
         t.Printp("Таблица 3 'АРВ (ИД)' - загружена!")
	end if
    
end if ' возможно лишний End If
'---------------------------Возбудителли IEEE-----------------
vozbuzdIEEE = 1
if vozbuzdIEEE = 1 then
	i = 3
	while Ex5.cells(i,1).value > 0
		'Eny=Ex5.cells(i,3)
		Name_gen = Ex5.cells(i,4)
		Id_generator = 0
		gen.SetSel("")
		j = gen.FindNextSel(-1)
		while j<>-1
			If gen.cols("Name").z(j) = Name_gen then
				Id_generator = gen.cols("Num").z(j)
			end if
			j = gen.FindNextSel(j)
		wend
		if Ex5.cells(i,5).value > 0 and Id_generator > 0 then
			vieee.SetSel("")
			vieee.AddRow
			vieee.SetSel("Id = 0")
			j2 = vieee.FindNextSel (-1)
			if j2<>-1 then
				vista.Z(j2)=Ex5.cells(i,2)
				viId.Z(j2)=Id_generator
				viName.Z(j2)=Ex5.cells(i,4)
				viBrand.Z(j2)=Ex5.cells(i,6)
				viCustomModel.Z(j2)=Ex5.cells(i,7)
				viUELId.Z(j2)=Ex5.cells(i,8)
				viUELPos.Z(j2)=Ex5.cells(i,9)
				viOELId.Z(j2)=Ex5.cells(i,10)
				viOELPos.Z(j2)=Ex5.cells(i,11)
				if Ex5.cells(i,12) > 0 then
					viPSSId.Z(j2) = Id_generator
				end if
				viPSSPos.Z(j2)=Ex5.cells(i,13)
				viTe.Z(j2)=Ex5.cells(i,14)
				viKe.Z(j2)=Ex5.cells(i,15)
				viSe1.Z(j2)=Ex5.cells(i,16)
				viEfd1.Z(j2)=Ex5.cells(i,17)
				viVe1.Z(j2)=Ex5.cells(i,18)
				viSe2.Z(j2)=Ex5.cells(i,19)
				viEfd2.Z(j2)=Ex5.cells(i,20)
				viVe2.Z(j2)=Ex5.cells(i,21)
				viVemin.Z(j2)=Ex5.cells(i,22)
				viVrmin.Z(j2)=Ex5.cells(i,23)
				viVrmax.Z(j2)=Ex5.cells(i,24)
				viKa.Z(j2)=Ex5.cells(i,25)
				viTa.Z(j2)=Ex5.cells(i,26)
				viTf.Z(j2)=Ex5.cells(i,27)
				viKf.Z(j2)=Ex5.cells(i,28)
				viTc.Z(j2)=Ex5.cells(i,29)
				viTb.Z(j2)=Ex5.cells(i,30)
				viKv.Z(j2)=Ex5.cells(i,31)
				viTrh.Z(j2)=Ex5.cells(i,32)
				viKpr.Z(j2)=Ex5.cells(i,33)
				viKir.Z(j2)=Ex5.cells(i,34)
				viKdr.Z(j2)=Ex5.cells(i,35)
				viTdr.Z(j2)=Ex5.cells(i,36)
				viKc.Z(j2)=Ex5.cells(i,37)
				viKd.Z(j2)=Ex5.cells(i,38)
				viVfemax.Z(j2)=Ex5.cells(i,39)
				viVamin.Z(j2)=Ex5.cells(i,40)
				viVamax.Z(j2)=Ex5.cells(i,41)
				viKb.Z(j2)=Ex5.cells(i,42)
				viKh.Z(j2)=Ex5.cells(i,43)
				viKr.Z(j2)=Ex5.cells(i,44)
				viKn.Z(j2)=Ex5.cells(i,45)
				viEfdn.Z(j2)=Ex5.cells(i,46)
				viKlv.Z(j2)=Ex5.cells(i,47)
				viVlv.Z(j2)=Ex5.cells(i,48)
				viVimin.Z(j2)=Ex5.cells(i,49)
				viVimax.Z(j2)=Ex5.cells(i,50)
				viTf2.Z(j2)=Ex5.cells(i,51)
				viTf3.Z(j2)=Ex5.cells(i,52)
				viTk.Z(j2)=Ex5.cells(i,53)
				viTj.Z(j2)=Ex5.cells(i,54)
				viTh.Z(j2)=Ex5.cells(i,55)
				viVhmax.Z(j2)=Ex5.cells(i,56)
				viVfelim.Z(j2)=Ex5.cells(i,57)
				viKp.Z(j2)=Ex5.cells(i,58)
				viKpa.Z(j2)=Ex5.cells(i,59)
				viKia.Z(j2)=Ex5.cells(i,60)
				viKf1.Z(j2)=Ex5.cells(i,61)
				viKf2.Z(j2)=Ex5.cells(i,62)
				viKl.Z(j2)=Ex5.cells(i,63)
				viTb1.Z(j2)=Ex5.cells(i,64)
				viTc1.Z(j2)=Ex5.cells(i,65)
				viKlr.Z(j2)=Ex5.cells(i,66)
				viIlr.Z(j2)=Ex5.cells(i,67)
				viKi.Z(j2)=Ex5.cells(i,68)
				viTheta.Z(j2)=Ex5.cells(i,69)
				viVmmin.Z(j2)=Ex5.cells(i,70)
				viVmmax.Z(j2)=Ex5.cells(i,71)
				viKg.Z(j2)=Ex5.cells(i,72)
				viVBmax.Z(j2)=Ex5.cells(i,73)
				viVGmax.Z(j2)=Ex5.cells(i,74)
				viXl.Z(j2)=Ex5.cells(i,75)
				viKm.Z(j2)=Ex5.cells(i,76)
				viTm.Z(j2)=Ex5.cells(i,77)
				viTb2.Z(j2)=Ex5.cells(i,78)
				viTc2.Z(j2)=Ex5.cells(i,79)
				viTub1.Z(j2)=Ex5.cells(i,80)
				viTuc1.Z(j2)=Ex5.cells(i,81)
				viTub2.Z(j2)=Ex5.cells(i,82)
				viTuc2.Z(j2)=Ex5.cells(i,83)
				viTob1.Z(j2)=Ex5.cells(i,84)
				viToc1.Z(j2)=Ex5.cells(i,85)
				viTob2.Z(j2)=Ex5.cells(i,86)
				viToc2.Z(j2)=Ex5.cells(i,87)
				viAex.Z(j2)=Ex5.cells(i,88)
				viBex.Z(j2)=Ex5.cells(i,89)
				viKcf.Z(j2)=Ex5.cells(i,90)
				viKhf.Z(j2)=Ex5.cells(i,91)
				viKif.Z(j2)=Ex5.cells(i,92)
				viSamovozb.Z(j2)=Ex5.cells(i,93)
				viTr.Z(j2)=Ex5.cells(i,94)
				viModel1=Ex5.cells(i,5)
				viModel.Z(j2)=viModel1
			end if
		end if
		i=i+1
	wend
    t.Printp("Таблица 5 'Возбудители IEEE' - загружена!")
end if
'---------------------------стабилизаторы IEEE-----------------
PSSE2 = 1
if PSSE2 = 1 then
	i = 3
	while Ex6.cells(i,1).value > 0
		'Eny=Ex6.cells(i,3)
		Name_gen = Ex6.cells(i,4)
		Id_generator=0
		gen.SetSel("")
		j = gen.FindNextSel (-1)
		while j<>-1
			If gen.cols("Name").z(j) = Name_gen then
				Id_generator = gen.cols("Num").z(j)
			end if
			j = gen.FindNextSel (j)
		wend
		if Ex6.cells(i,5).value > 0 and Id_generator > 0 then
			stieee.SetSel("")
			stieee.AddRow
			stieee.SetSel("Id = 0")
			j2 = stieee.FindNextSel (-1)
			if j2<>-1 then
				ststa.Z(j2)=Ex6.cells(i,2)
				stId.Z(j2)=Id_generator
				stName.Z(j2)=Ex6.cells(i,4)
				stModel1=Ex6.cells(i,5)
				stModel.Z(j2)=stModel1
				stBrand.Z(j2)=Ex6.cells(i,6)
				stCustomModel.Z(j2)=Ex6.cells(i,7)
				stInput1Type.Z(j2)=Ex6.cells(i,8)
				stInput2Type.Z(j2)=Ex6.cells(i,9)
				stVstmin.Z(j2)=Ex6.cells(i,10)
				stVstmax.Z(j2)=Ex6.cells(i,11)
				stKs1.Z(j2)=Ex6.cells(i,12)
				stT1.Z(j2)=Ex6.cells(i,13)
				stT2.Z(j2)=Ex6.cells(i,14)
				stT3.Z(j2)=Ex6.cells(i,15)
				stT4.Z(j2)=Ex6.cells(i,16)
				stT5.Z(j2)=Ex6.cells(i,17)
				stT6.Z(j2)=Ex6.cells(i,18)
				stT7.Z(j2)=Ex6.cells(i,19)
				stT8.Z(j2)=Ex6.cells(i,20)
				stT9.Z(j2)=Ex6.cells(i,21)
				stT10.Z(j2)=Ex6.cells(i,22)
				stT11.Z(j2)=Ex6.cells(i,23)
				stA1.Z(j2)=Ex6.cells(i,24)
				stA2.Z(j2)=Ex6.cells(i,25)
				stA3.Z(j2)=Ex6.cells(i,26)
				stA4.Z(j2)=Ex6.cells(i,27)
				stA5.Z(j2)=Ex6.cells(i,28)
				stA6.Z(j2)=Ex6.cells(i,29)
				stA7.Z(j2)=Ex6.cells(i,30)
				stA8.Z(j2)=Ex6.cells(i,31)
				stKs2.Z(j2)=Ex6.cells(i,32)
				stKs3.Z(j2)=Ex6.cells(i,33)
				stTw1.Z(j2)=Ex6.cells(i,34)
				stTw2.Z(j2)=Ex6.cells(i,35)
				stTw3.Z(j2)=Ex6.cells(i,36)
				stTw4.Z(j2)=Ex6.cells(i,37)
				stM.Z(j2)=Ex6.cells(i,38)
				stN.Z(j2)=Ex6.cells(i,39)
				stVsi1min.Z(j2)=Ex6.cells(i,40)
				stVsi1max.Z(j2)=Ex6.cells(i,41)
				stVsi2min.Z(j2)=Ex6.cells(i,42)
				stVsi2max.Z(j2)=Ex6.cells(i,43)
			end if
		end if
		i=i+1
	wend
    t.Printp("Таблица 6 'Стабилизатор 1-3 PSS2' - загружена!")
end if
'---------------------------стабилизатор PSS4B IEEE-----------------
PSS4_2 = 1
if PSS4_2 = 1 then 
	i = 3
	while Ex7.cells(i,1).value > 0
		'Eny=Ex7.cells(i,3)
		Name_gen = Ex7.cells(i,4)
		Id_generator = 0
		gen.SetSel("")
		j = gen.FindNextSel(-1)
		while j<>-1
			If gen.cols("Name").z(j) = Name_gen then
				Id_generator = gen.cols("Num").z(j)
			end if
			j = gen.FindNextSel (j)
		wend
		if Ex7.cells(i,5).value>0 and Id_generator>0 then
			pss4.SetSel("")
			pss4.AddRow
			pss4.SetSel("Id = 0")
			j2=pss4.FindNextSel (-1)
			if j2<>-1 then
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
			end if
		end if
		i=i+1
	wend
    t.Printp("Таблица 7 'Стабилизатор PSS 4' - загружена!")
end if
'----------------турбина----------------
turbina = 1
if turbina = 1 then
	i = 3
	while Ex4.cells(i,1).value > 0
		'Eny=Ex4.cells(i,3)
		Name_gen = Ex4.cells(i,4)
		Id_generator = 0
		gen.SetSel("")
		j = gen.FindNextSel (-1)
		while j<>-1
			If gen.cols("Name").z(j) = Name_gen then
				Id_generator = gen.cols("Num").z(j)
			end if
			j = gen.FindNextSel (j)
		wend
		if Ex4.cells(i,5).value > 0 and Id_generator > 0 then
			ars.SetSel("")
			ars.AddRow
			ars.SetSel("Id = 0")
			j2 = ars.FindNextSel (-1)
			if j2<>-1 then
				arsids.Z(j2)=Ex4.cells(i,3)
				ideg1=Id_generator
				arsname.Z(j2)=Ex4.cells(i,4)
				ModelTypes111=Ex4.cells(i,5)
				arsModelTypes.Z(j2)=ModelTypes111
				arsBrands.Z(j2)=Ex4.cells(i,6)
				if Ex4.cells(i,7)>0 then
					arsGovernorId.Z(j2)=Id_generator
				end if
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
			end if
		end if
		i=i+1
	wend
' АРС
	i = 3
	while Ex12.cells(i,1).value > 0
		'Eny=Ex12.cells(i,3)
		Name_gen = Ex12.cells(i,4)
		Id_generator = 0
		gen.SetSel("")
		j = gen.FindNextSel (-1)
		while j<>-1
			If gen.cols("Name").z(j) = Name_gen then
				Id_generator = gen.cols("Num").z(j)
			end if
			j = gen.FindNextSel (j)
		wend
		if Ex12.cells(i,5).value > 0 and Id_generator > 0 then
			Governor.SetSel("")
			Governor.AddRow
			Governor.SetSel("Id = 0")
			j2 = Governor.FindNextSel(-1)
			if j2<>-1 then
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
			end if
		end if
		i=i+1
	wend
    t.Printp("Таблица 4 и 12 'Турбина  и АРС' - загружена!")
end if
'-----------Форсировка
forc2 = 1
if forc2 = 1 then
	i = 3
	while Ex8.cells(i,1).value > 0
		'Eny=Ex8.cells(i,3)
		Name_gen = Ex8.cells(i,4)
		Id_generator = 0
		gen.SetSel("")
		j = gen.FindNextSel (-1)
		while j<>-1
			If gen.cols("Name").z(j) = Name_gen then
				Id_generator = gen.cols("Num").z(j)
			end if
			j = gen.FindNextSel (j)
		wend
		if Ex8.cells(i,5).value>0 and Id_generator>0 then
			forc.SetSel("")
			forc.AddRow
			forc.SetSel("Id = 0")
			j2=forc.FindNextSel (-1)
			if j2<>-1 then
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
			end if
		end if
		i=i+1
	wend
    t.Printp("Таблица 5 'Стабилизатор 1-3 PSS2' - загружена!")
end if
'-----------ОМВ
OMV_2 = 0
if OMV_2 = 1 then
	i = 3
	while Ex9.cells(i,1).value > 0
		Eny=Ex9.cells(i,3)
		if Ex9.cells(i,5).value > 0 and Ex9.cells(i,3).value > 0 then
			omv.SetSel("")
			omv.AddRow
			omv.SetSel("Id = 0")
			j2 = omv.FindNextSel (-1)
			if j2<>-1 then
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
				omvDependency_F1.Z(j2)=Ex9.cells(i,29)
				omvOutput.Z(j2)=Ex9.cells(i,30)
				omvKl.Z(j2)=Ex9.cells(i,31)
			end if
		end if
		i=i+1
	wend
    t.Printp("Таблица 9 'ОМВ' - загружена!")
end if  
'----------БОР
BOR_2 = 0
if BOR_2 = 1 then
	i = 3
	while Ex10.cells(i,1).value > 0
		Eny = Ex10.cells(i,3)
		if Ex10.cells(i,5).value > 0 and Ex10.cells(i,3).value > 0 then
			bor.SetSel("")
			bor.AddRow
			bor.SetSel("Id = 0")
			j2=bor.FindNextSel (-1)
			if j2<>-1 then
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
			end if
		end if
		i=i+1
	wend
    t.Printp("Таблица 10 'БОР' - загружена!")
end if
'----------ОМВ PQ
OMVPQ_2 = 0
if OMVPQ_2 = 1 then
	i = 3
	while Ex11.cells(i,1).value > 0
		'Eny=Ex11.cells(i,3)
		if Ex11.cells(i,2).value > 0  then
			i0 = 3
			while Ex11.cells(i,i0).value <> 0 or Ex11.cells(i,i0+1).value <> 0
				FuncPQ.SetSel("")
				FuncPQ.AddRow
				FuncPQ.SetSel("Id = 0")
				j2 = FuncPQ.FindNextSel(-1)
				if j2<>-1 then
					FuncPQId.Z(j2)=Ex11.cells(i,2)
					FuncPQP.Z(j2)=Ex11.cells(i,i0)
					FuncPQQ.Z(j2)=Ex11.cells(i,i0+1)
					i0=i0+2
				end if
			wend
		end if
		i=i+1
	wend
    t.Printp("Таблица 11 'Зависимость Q(P)' - загружена!")
end if
'------------- DECS-400 ----------------------
DECS_400 = 0
if DECS_400 = 1 then 
	i = 3
	while Ex13.cells(i,1).value > 0
		Eny = Ex13.cells(i,3)
		Name_gen = Ex13.cells(i,4)
		Id_generator = 0
		gen.SetSel("")
		j = gen.FindNextSel(-1)
		while j<>-1
			If gen.cols("Name").z(j) = Name_gen then
				Id_generator = gen.cols("Num").z(j)
			end if
			j = gen.FindNextSel (j)
		wend
		if Ex13.cells(i,5).value>0 and Id_generator>0 then
			decs400.SetSel("")
			decs400.AddRow
			decs400.SetSel("Id = 0")
			j2 = decs400.FindNextSel(-1)
			if j2<>-1 then
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
			end if
		end if
		i=i+1
	wend
    t.Printp("Таблица 13 'DECS - 400' - загружена!")
end if
'------------- Thyne-4 ----------------------
Thyne4 = 0
if Thyne4 = 1 then 
	i = 3
	while Ex14.cells(i,1).value > 0
		Eny = Ex14.cells(i,3)
		Name_gen = Ex14.cells(i,4)
		Id_generator = 0
		gen.SetSel("")
		j = gen.FindNextSel(-1)
		while j<>-1
			If gen.cols("Name").z(j) = Name_gen then
				Id_generator = gen.cols("Num").z(j)
			end if
			j = gen.FindNextSel (j)
		wend
		if Ex14.cells(i,5).value > 0 and Id_generator > 0 then
			Thyne.SetSel("")
			Thyne.AddRow
			Thyne.SetSel("Id = 0")
			j2 = Thyne.FindNextSel (-1)
			if j2<>-1 then
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
			end if
		end if
		i=i+1
	wend
    t.Printp("Таблица 14 'Thyne-4' - загружена!")
end if
'---------------------------------------------------
allGEN = 1
if allGEN = 1 then
	uzl.SetSel("(pg!=0 | qg!=0 | qmax!=0 | qmin!=0) & !sta")
	j = uzl.FindNextSel(-1)
	while j<>-1
		nygen1=uzl.Cols("ny").z(j)
		gen.SetSel "Node="&nygen1
		jj = gen.findnextsel(-1)
		if jj<>-1 then
		else
			gen.AddRow
			gen.SetSel(" Node = 0")
			j2 = gen.FindNextSel(-1)
			if j2<>-1 then
				numg.Z(j2)=nygen1
				nameg.Z(j2)=uzl.Cols("name").z(j)
				nodeg.Z(j2)=nygen1
				ModelType.Z(j2)=3
				pgen1=uzl.Cols("pg").z(j)
				pgen.Z(j2)=pgen1
				if pgen1 > 10 or pgen1 < (-10) then
					pnom.Z(j2)=abs(pgen1)
				else
					pnom.Z(j2)=10
				end if
				unom.Z(j2)=uzl.Cols("uhom").z(j)
				cosfi.Z(j2)=0.85
				Demp.Z(j2)=20
				mj.Z(j2)=5*pnom.Z(j2)/cosfi.Z(j2)
				xd1.Z(j2)=0.2*unom.Z(j2)*unom.Z(j2)*cosfi.Z(j2)/pnom.Z(j2)
			end if
		end if
		j = uzl.FindNextSel(j)
	wend
    t.Printp("Заполнение всех генераторов - завершено!")
end if
'-----------------------
'-----------------------
otklstab = 0
if otklstab = 1 then
	gen.SetSel("")
	j = gen.FindNextSel(-1)
	while j<>-1
		stagen = nsta.Z(j)
		'rastr.printp stagen
		if nsta.Z(j) = True then
			stagen = 1
		else
			stagen = 0
		end if
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
	wend
end if

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
t.Tables("AltUnit").Cols("Unit").z(0)="МВт*с"
t.Tables("AltUnit").Cols("Alt").z(0)="с"
t.Tables("AltUnit").Cols("Formula").z(0)="nonz(cosFi)/(Pnom)"
t.Tables("AltUnit").Cols("Prec").z(0)=3
t.Tables("AltUnit").Cols("Tabl").z(0)="Generator"

t.Tables("AltUnit").Cols("Activ").z(1)=1
t.Tables("AltUnit").Cols("Unit").z(1)="Ом"
t.Tables("AltUnit").Cols("Alt").z(1)="о.е."
t.Tables("AltUnit").Cols("Formula").z(1)="Pnom/(Ugnom*Ugnom*nonz(cosFi))"
t.Tables("AltUnit").Cols("Prec").z(1)=4
t.Tables("AltUnit").Cols("Tabl").z(1)="Generator"


Custom_Models = 1
if Custom_Models = 1 then 
	'Задать ссылку для пользовательских устройств
	Link = Link_Custom_Models
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
end if
'-----------------------------------------------------------------------------------------

'--------------цены делений
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
Rastr.printp "Исследование завершено (=_=)"
