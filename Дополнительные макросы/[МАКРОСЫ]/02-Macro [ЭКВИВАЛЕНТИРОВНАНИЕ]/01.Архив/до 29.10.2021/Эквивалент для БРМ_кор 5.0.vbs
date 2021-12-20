
r=setlocale("en-us")
rrr=1


set t=RASTR

set vet=t.tables("vetv")
set uzl=t.tables("node")
set ray=t.tables("area")
set gen=t.tables("Generator")
set pqd=t.Tables("graphik2")
set graphikIT=t.Tables("graphikIT")
set area=t.Tables("area")
set area2=t.Tables("area2")
set darea=t.Tables("darea")
set polin=t.Tables("polin")
set Reactors=t.Tables("Reactors")



t.rgm "p"
ibnulenie ""

Vikluchatel ""

ibnulenie ""

Ukraine ""
ibnulenie ""

vyborka_rayon2="(na=407 )"
Ekvivalent_siln vyborka_rayon2
ibnulenie ""

vyborka_gen="((na>100 & na<200 & na!=108)| (na>300 & na<400 & na!=311) | na=201 | na=203 | na=205 | na=208 | na=206 | na=805 | na=806 | na=807) & (uhom=110 | uhom=220) "
Ekv_gen vyborka_gen
ibnulenie ""
'Ekv_gen vyborka_gen
'ibnulenie ""
'Ekv_gen vyborka_gen
'ibnulenie ""
'Ekv_gen vyborka_gen
'ibnulenie ""
'Ekv_gen vyborka_gen

vyborka_rayon="(((na!=108 & (na>100 & na<200))| (na>300 & na<400 & na!=311)| na=201 | na=203 | na=205 | na=208 | na=206| na=805 | na=806 | na=807)& (uhom=110 | uhom=220)"
Ekvivalent_smart vyborka_rayon

ibnulenie ""
'Udalenie ""
'--- эквивалентирование----


Sub Ekvivalent_smart (vyborka_rayon)
vet.SetSel("")
vet.cols("sel").calc("0")
uzl.SetSel("")
uzl.cols("sel").calc("0")


t.Tables("com_ekviv").Cols("zmax").z(0)=1000
t.Tables("com_ekviv").Cols("ek_sh").z(0)=0
t.Tables("com_ekviv").Cols("otm_n").z(0)=0
t.Tables("com_ekviv").Cols("smart").z(0)=1
t.Tables("com_ekviv").Cols("tip_ekv").z(0)=0
t.Tables("com_ekviv").Cols("ekvgen").z(0)=0
t.Tables("com_ekviv").Cols("tip_gen").z(0)=1


uzl.setsel vyborka_rayon
uzl.cols("sel").calc("1")



'vet.SetSel("iq.sel=1 &ip.sel=0 &!sta")
'k=vet.findnextsel(-1)
'while k<>(-1)
'iq1=vet.Cols("iq").z(k)
'uzl.setsel "ny="&iq1
'k2=uzl.findnextsel(-1)
'if k2<>-1 then
'uzl.cols("sel").z(k2)=0
'end if
'k=vet.findnextsel(k)
'wend

'vet.SetSel("iq.sel=0 &ip.sel=1 &!sta")

'k=vet.findnextsel(-1)
'while k<>(-1)
'ip1=vet.Cols("ip").z(k)
'uzl.setsel "ny="&ip1
'k2=uzl.findnextsel(-1)
'if k2<>-1 then
'uzl.cols("sel").z(k2)=0
'end if
'k=vet.findnextsel(k)
'wend

t.Ekv""

end sub

Sub Ekvivalent_siln (vyborka_rayon2)
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


uzl.setsel vyborka_rayon2
uzl.cols("sel").calc("1")



vet.SetSel("iq.sel=1 &ip.sel=0 &!sta")
k=vet.findnextsel(-1)
while k<>(-1)
iq1=vet.Cols("iq").z(k)
uzl.setsel "ny="&iq1
k2=uzl.findnextsel(-1)
if k2<>-1 then
uzl.cols("sel").z(k2)=0
end if
k=vet.findnextsel(k)
wend

vet.SetSel("iq.sel=0 &ip.sel=1 &!sta")
k=vet.findnextsel(-1)
while k<>(-1)
ip1=vet.Cols("ip").z(k)
uzl.setsel "ny="&ip1
k2=uzl.findnextsel(-1)
if k2<>-1 then
uzl.cols("sel").z(k2)=0
end if
k=vet.findnextsel(k)
wend

t.Ekv""

end sub


Sub Ekv_gen (vyborka_gen)
uzl.setsel vyborka_gen
k=uzl.findnextsel(-1)
while k<>(-1)
ny1=uzl.Cols("ny").z(k)
vet.SetSel("(ip.uhom<110 & iq=" & ny1 & ") | (iq.uhom<110 & ip="&ny1) 
k2=vet.findnextsel(-1)
while k2<>(-1)
ip1=vet.Cols("ip").z(k2)
iq1=vet.Cols("iq").z(k2)

if ip1=ny1 then
ny2=iq1
else
ny2=ip1
end if

uzl.setsel "ny="&ny2
k3=uzl.findnextsel(-1)
if k3<>-1 then
uzl.cols("sel").z(k3)=1
end if

k2=vet.findnextsel(k2)
wend

uzl.setsel vyborka_gen
k=uzl.findnextsel(k)
wend

t.Tables("com_ekviv").Cols("zmax").z(0)=1000
t.Tables("com_ekviv").Cols("ek_sh").z(0)=0
t.Tables("com_ekviv").Cols("otm_n").z(0)=0
t.Tables("com_ekviv").Cols("smart").z(0)=0
t.Tables("com_ekviv").Cols("tip_ekv").z(0)=0
t.Tables("com_ekviv").Cols("ekvgen").z(0)=0
t.Tables("com_ekviv").Cols("tip_gen").z(0)=1


t.Ekv""

uzl.setsel "uhom>50"
uzl.cols("sel").calc("0")
t.Ekv""

uzl.setsel "uhom>50"
uzl.cols("sel").calc("0")
t.Ekv""
uzl.setsel "uhom>50"
uzl.cols("sel").calc("0")
t.Ekv""
uzl.setsel "uhom>50"
uzl.cols("sel").calc("0")
t.Ekv""
uzl.setsel "uhom>50"
uzl.cols("sel").calc("0")
t.Ekv""
uzl.setsel "uhom>50"
uzl.cols("sel").calc("0")
t.Ekv""
end sub


Sub ibnulenie (alpha)
vet.SetSel("")
vet.cols("sel").calc("0")
uzl.SetSel("")
uzl.cols("sel").calc("0")
end sub


Sub Vikluchatel (alpha)

uzl.SetSel("na<500 | na>600")
uzl.cols("sel").calc(1)


vet.SetSel("iq.sel=1 &ip.sel=0 &!sta")
k=vet.findnextsel(-1)
while k<>(-1)
iq1=vet.Cols("iq").z(k)
uzl.setsel "ny="&iq1
k2=uzl.findnextsel(-1)
if k2<>-1 then
uzl.cols("sel").z(k2)=0
end if
k=vet.findnextsel(k)
wend

vet.SetSel("iq.sel=0 &ip.sel &!sta")

k=vet.findnextsel(-1)
while k<>(-1)
ip1=vet.Cols("ip").z(k)
uzl.setsel "ny="&ip1
k2=uzl.findnextsel(-1)
if k2<>-1 then
uzl.cols("sel").z(k2)=0
end if
k=vet.findnextsel(k)
wend


vet.SetSel("(iq.sel=1 &ip.sel=0) | (ip.sel=1 &iq.sel=0) & tip=2")
k=vet.findnextsel(-1)
while k<>(-1)
iq1=vet.Cols("iq").z(k)
uzl.setsel "ny="&iq1
k2=uzl.findnextsel(-1)
if k2<>-1 then
uzl.cols("sel").z(k2)=0
end if
ip1=vet.Cols("ip").z(k)
uzl.setsel "ny="&ip1
k2=uzl.findnextsel(-1)
if k2<>-1 then
uzl.cols("sel").z(k2)=0
end if
vet.SetSel("(iq.sel=1 &ip.sel=0) | (ip.sel=1 &iq.sel=0) & tip=2")
k=vet.findnextsel(-1)
wend

Set cvzd=uzl.Cols("vzd")
set csel=uzl.Cols("sel")
set cip=vet.cols("ip") 
set ciq=vet.cols("iq") 
Dim nyplus(10000,8),vetmassiv(15000,3),nodes(15000)

vetvyklvybexc = "(iq.bsh>0 & ip.bsh=0) | (ip.bsh>0 & iq.bsh=0) | (iq.bshr>0 & ip.bshr=0) | (ip.bshr>0 & iq.bshr=0)| ip.sel=0 | iq.sel=0)"


	flvykl=0

	vet.SetSel "1"
	vet.cols("groupid").calc(0)

	'vet.SetSel "x=666"
	'vet.cols("x").calc(665)

	vet.SetSel vetvyklvybexc
	vet.cols("groupid").calc(1)


	nvet=0
	for povet=0 to 10000
		vet.SetSel("x<0.01 & x>-0.01 & r<0.005 & r>=0 & (ktr=0 | ktr=1) & !sta &groupid!=1 & b<0.000005")'Выборка ветвей, которые считаем выключателями
		'vet.SetSel("tip=2 & x<0.01 & x>-0.01 & r<0.005 & r>=0 & (ktr=0 | ktr=1) & !sta &groupid!=1 & b<0.000005")
		ivet=vet.FindNextSel(-1)

		if ivet=-1 then exit for

		ip=vet.Cols("ip").z(ivet)
		iq=vet.Cols("iq").z(ivet)
		if ip>iq then
			ny=iq 
			ndel=ip
		else 
			ny=ip
			ndel=iq
		end if

		ndny=0
		ndndel=0
'Проверка на наличие узла из списка неудаляемых
		for inodee=0 to nnod
			if 	ndel=nodes(inodee) then ndndel=1
			if 	ny=nodes(inodee) then ndny=1
			if ndndel=1 and ndny=1 then exit for
		next

' Меняем местами, так как удаляемый нельзя удалять, а неудаляемый можно ))
		if ndndel=0 and ndny=1 then
			buff=ny
			ny=ndel
			ndel=buff
		end if

		if ndndel=0 or ndny=0 then 'Если хотя бы один можно удалить


			flvykl=flvykl+1

				uzl.SetSel("ny="&ny)
				iny=uzl.findnextsel(-1)

				uzl.SetSel("ny="&ndel)
				idel=uzl.findnextsel(-1)

				pgdel=uzl.cols("pg").z(idel)
				qgdel=uzl.cols("qg").z(idel)
				pndel=uzl.cols("pn").z(idel)
				qndel=uzl.cols("qn").z(idel)
				bshdel=uzl.cols("bsh").z(idel)
				gshdel=uzl.cols("gsh").z(idel)

				pgny=uzl.cols("pg").z(iny)
				qgny=uzl.cols("qg").z(iny)
				pnny=uzl.cols("pn").z(iny)
				qnny=uzl.cols("qn").z(iny)
				bshny=uzl.cols("bsh").z(iny)
				gshny=uzl.cols("gsh").z(iny)

				uzl.cols("pg").z(iny)=pgdel+pgny
				uzl.cols("qg").z(iny)=qgdel+qgny
				uzl.cols("pn").z(iny)=pndel+pnny
				uzl.cols("qn").z(iny)=qndel+qnny
				uzl.cols("bsh").z(iny)=bshdel+bshny
				uzl.cols("gsh").z(iny)=gshdel+gshny


				v1=uzl.cols("vzd").z(iny)
				v2=uzl.cols("vzd").z(idel)
				qmax1=uzl.cols("qmax").z(iny)
				qmax2=uzl.cols("qmax").z(idel)



				'writelog "Выключатели. #"& flvykl &". Оставляем узел ny= "&ny&". Удаляем узел ndel= "& ndel 
				
				gen.setsel("Node="&ndel)
				igen=gen.findnextsel(-1) 'Меняем узлы подключения генераторов
				if igen<>-1 then
					while igen<>-1 
						gen.cols("Node").z(igen)=ny
						igen=gen.findnextsel(igen)
					wend
				end if

				if v1<>v2 and v1>0.3 and v2>0.3 and (qmax1+qmax2)<>0 then
					uzl.cols("vzd").z(iny)=(v1*qmax1+v2*qmax2)/(qmax1+qmax2) 'Делаем средневзвешенное по qmax напряжение
				end if

				if v1=0 and v2<>0 then
					uzl.cols("vzd").z(iny)=v2
				end if


				if v1<>0 and v2<>0 then
					uzl.cols("qmin").z(iny)=uzl.cols("qmin").z(iny)+uzl.cols("qmin").z(idel)
					uzl.cols("qmax").z(iny)=qmax1+qmax2
				end if

				if v1=0 and v2<>0 then
					uzl.cols("qmin").z(iny)=uzl.cols("qmin").z(idel)
					uzl.cols("qmax").z(iny)=uzl.cols("qmax").z(idel)
				end if


				vet.SetSel("(ip="&ip &"& iq="&iq& ")|(iq="&ip &"& ip="&iq& ")")
				vet.delrows 'Удаляем ветвь	


				vet.SetSel("iq="&ndel) 'Меняем узлы ветвей с удаляемым узлом)))
				vet.cols("iq").calc(ny)	

				vet.SetSel("ip="&ndel)
				vet.cols("ip").calc(ny)	


				uzl.delrows' Удаляем узел



		else 'Если ни одного нельзя удалить
			vet.SetSel("(ip="&ip &"& iq="&iq& ")|(iq="&ip &"& ip="&iq& ")")
			vet.cols("groupid").calc(1)
		end if


	next

	'writelog "Выключатели. Обработано "& flvykl &" штук."

	kod = t.rgm ("p")
	if kod<>0 then
		msgbox "Regim do not exist"
		'writelog "!!! After vykldel Regim do not exist!!!!!!"		
	end if

end sub

Sub Ukraine (alpha)
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


vet.SetSel("(iq.na=803 & ip.na>300 & ip.na<400) ")
k=vet.findnextsel(-1)
while k<>(-1)
iq1=vet.Cols("iq").z(k)
uzl.setsel "ny="&iq1
k2=uzl.findnextsel(-1)
if k2<>-1 then
uzl.cols("sel").z(k2)=1
end if

k=vet.findnextsel(k)
wend

vet.SetSel("(ip.na=803 & iq.na>300 & iq.na<400) ")
k=vet.findnextsel(-1)
while k<>(-1)
ip1=vet.Cols("ip").z(k)
uzl.setsel "ny="&ip1
k2=uzl.findnextsel(-1)
if k2<>-1 then
uzl.cols("sel").z(k2)=1
end if
k=vet.findnextsel(k)
wend


vet.SetSel("((iq.sel=1 &ip.sel=0) | (ip.sel=1 &iq.sel=0)) & ip.na=803 & iq.na=803 &!sta")
k=vet.findnextsel(-1)
while k<>(-1)
iq1=vet.Cols("iq").z(k)
uzl.setsel "ny="&iq1
k2=uzl.findnextsel(-1)
if k2<>-1 then
uzl.cols("sel").z(k2)=1
end if
ip1=vet.Cols("ip").z(k)
uzl.setsel "ny="&ip1
k2=uzl.findnextsel(-1)
if k2<>-1 then
uzl.cols("sel").z(k2)=1
end if
vet.SetSel("((iq.sel=1 &ip.sel=0) | (ip.sel=1 &iq.sel=0)) & ip.na=803 & iq.na=803 &!sta")
k=vet.findnextsel(-1)
wend


t.Ekv""

end sub


Sub Udalenie (alpha)


uzl.setsel("")
k2=uzl.findnextsel(-1)
while k2<>(-1)
ny1=uzl.cols("ny").z(k2)
vet.SetSel("((ip=" & ny1 & ") | (iq="&ny1 & "))" )
if vet.count=0 then
uzl.cols("sel").z(k2)=1
end if
k2=uzl.findnextsel(k2)
wend
uzl.setsel("sel=1")
uzl.delrows

Reactors.setsel("")
k2=Reactors.findnextsel(-1)
while k2<>(-1)
ny1=Reactors.cols("Id1").z(k2)
uzl.SetSel("(ny=" & ny1 & ") " )
if uzl.count=0 then
Reactors.cols("sel").z(k2)=1
end if
k2=Reactors.findnextsel(k2)
wend
Reactors.setsel("sel=1")
Reactors.delrows


gen.setsel("Node.na=0")
gen.delrows


graphikIT.setsel("")
k=graphikIT.findnextsel(-1)
while k<>(-1)
nzav=graphikIT.cols("Num").z(k)


vet.setsel("n_it="&nzav)
k2=vet.findnextsel(-1)
if k2<>-1 then
else
graphikIT.cols("Num").z(k)=0
end if



k=graphikIT.findnextsel(k)
wend

graphikIT.setsel("Num=0")
graphikIT.delrows



area.setsel("")
k=area.findnextsel(-1)
while k<>(-1)
na1=area.cols("na").z(k)


uzl.setsel("na="&na1)
k2=uzl.findnextsel(-1)
if k2<>-1 then
else
area.cols("na").z(k)=0
end if



k=area.findnextsel(k)
wend

area.setsel("na=0")
area.delrows



area2.setsel("")
k=area2.findnextsel(-1)
while k<>(-1)
na1=area2.cols("npa").z(k)


uzl.setsel("npa="&na1)
k2=uzl.findnextsel(-1)
if k2<>-1 then
else
area2.cols("npa").z(k)=0
end if



k=area2.findnextsel(k)
wend

area2.setsel("npa=0")
area2.delrows




darea.setsel("")
k=darea.findnextsel(-1)
while k<>(-1)
na1=darea.cols("no").z(k)


area.setsel("no="&na1)
k2=area.findnextsel(-1)
if k2<>-1 then
else
darea.cols("no").z(k)=0
end if



k=darea.findnextsel(k)
wend

darea.setsel("no=0")
darea.delrows



polin.setsel("")
k=polin.findnextsel(-1)
while k<>(-1)
nsx1=polin.cols("nsx").z(k)


uzl.setsel("nsx="&nsx1)
k2=uzl.findnextsel(-1)
if k2<>-1 then
else
polin.cols("nsx").z(k)=0
end if



k=polin.findnextsel(k)
wend

polin.setsel("nsx=0")
polin.delrows




t.rgm "p"
end sub