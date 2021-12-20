r=setlocale("en-us")
rrr=1
set t=RASTR
    set vet=t.tables("vetv")
    set uzl=t.tables("node")
    set gen=t.Tables("Generator")

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
    set P	    =gen.Cols("P")
    set unom	=gen.Cols("Ugnom")
    set cosfi	=gen.Cols("cosFi")
    set Demp	=gen.Cols("Demp")
    set mj		=gen.Cols("Mj")
    set xd1		=gen.Cols("xd1")
    set qmin	=gen.Cols("Qmin")
    set qmax	=gen.Cols("Qmax")
    'set nodeld	=gen.Cols("nodeld")
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


CountGen = gen.Count - 1
t.printp("CountGen = " & CountGen)
key1 = 1 ' јвто выбор
key2 = 0 ' ”р движ
for i=0 to CountGen 
     ModelType1 = ModelType.z(i)
     't.Printp("Model Type =" & ModelType1)
     if ModelType1 = "" Then 
        't.Printp("Model Type =" & ModelType1)
        ModelType1.Z(i) = 0
     End If
	 if key1 = 1 Then  
        Call Viborka_po_P(i)
     End If
     If key2 = 1 Then
        Call Viborka_ALL(i,ModelType1)
     End If
next

allGEN2=0
if allGEN2=1 then
    uzl.SetSel("(pg!=0 | qg!=0 | qmax!=0 | qmin!=0) & !sta")
    j=uzl.FindNextSel (-1)
    while j<>-1
        nygen1=uzl.Cols("ny").z(j)
        gen.SetSel "Node="&nygen1
        jj=gen.findnextsel(-1)
        if jj<>-1 then
        
        else
            gen.AddRow
            gen.SetSel(" Node = 0")
            j2=gen.FindNextSel(-1)
            t.printp(j2)
            if j2<>-1 then
                numg.Z(j2)=nygen1
                nameg.Z(j2)=uzl.Cols("name").z(j)
                nodeg.Z(j2)=nygen1
                ModelType.Z(j2)=3
                pgen1=uzl.Cols("pg").z(j)
                pgen.Z(j2)=pgen1
                qgen1=uzl.Cols("qg").z(j)
                
                if abs(pgen1)>5  then
                    pnom.Z(j2)=abs(pgen1)
                else
                    pnom.Z(j2)=5
                end if
                
                if abs(qgen1) > abs(pgen1) and abs(qgen1) > 10  then
                    pnom.Z(j2) = abs(qgen1)/0.85
                end if
                
                'if abs(uzl.Cols("qmax").z(j))>abs(pgen1) then
                    'pnom.Z(j2)=abs(uzl.Cols("qmax").z(j))
                'end if
                
                unom.Z(j2)=uzl.Cols("uhom").z(j)
                cosfi.Z(j2)=0.85
                Demp.Z(j2)=5
                mj.Z(j2)=5*pnom.Z(j2)/cosfi.Z(j2)
                xd1.Z(j2)=0.3*unom.Z(j2)*unom.Z(j2)*cosfi.Z(j2)/pnom.Z(j2)
                qmax.Z(j2)=uzl.Cols("qmax").z(j)
                qmin.Z(j2)=uzl.Cols("qmin").z(j)
            end if
        end if
        j = uzl.FindNextSel (j)
    wend
end if
'-----------------------


Sub Viborka_ALL(i,ModelType1)
	if ModelType1 = 3 Then
        nygen1 = gen.Cols("Node").Z(i)
        gen.SetSel "Node =" & nygen1
        j = gen.findnextsel(-1)
        ModelType.Z(i) = 3
        pnom_per = pnom.Z(i)
        
		if pnom_per = 0 then
            pnom.z(i) = P.z(i)
        end if
		
		unom.Z(i)= uzl.Cols("uhom").z(j)
		cosfi.Z(i)=0.85
		Demp.Z(i)=5
		mj.Z(i)= 5
		xd1.Z(i)=0.3
		qmax.Z(i)=uzl.Cols("qmax").z(j)
		qmin.Z(i)=uzl.Cols("qmin").z(j)       
    end if	
End Sub

Sub Viborka_po_P(i)
	if ModelType1 = 0 Then
        nygen1 = gen.Cols("Node").Z(i)
        gen.SetSel "Node =" & nygen1
        j = gen.findnextsel(-1)
        ModelType.Z(i) = 3
        pnom_per = pnom.Z(i)
        if pnom_per = 0 then
            pnom.z(i) = P.z(i)
        end if
        if pnom.Z(i)=< 0 Then 
            pnom.z(i) = (P.z(i)) * (-1)
            unom.Z(i)= uzl.Cols("uhom").z(j)
            cosfi.Z(i)=0.85
            Demp.Z(i)=5
            if P.Z(i) = 0 Then 
                mj.Z(i)= 5
                xd1.Z(i)=0.3
                qmax.Z(i)=uzl.Cols("qmax").z(j)
                qmin.Z(i)=uzl.Cols("qmin").z(j)
            else 
				mj.Z(i)= 5
                xd1.Z(i)=0.3
                qmax.Z(i)=uzl.Cols("qmax").z(j)
                qmin.Z(i)=uzl.Cols("qmin").z(j)
				
                'mj.Z(i)=5 * P.Z(i)/cosfi.Z(i)
                'xd1.Z(i)=0.3*unom.Z(i)*unom.Z(i)*cosfi.Z(i)/P.Z(i)
                'qmax.Z(i)=uzl.Cols("qmax").z(j)
                'qmin.Z(i)=uzl.Cols("qmin").z(j)
            end if
        else 
            unom.Z(i)= uzl.Cols("uhom").z(j)
            cosfi.Z(i)=0.85
            Demp.Z(i)=5
            mj.Z(i)=5*pnom.Z(i)/cosfi.Z(i)
            xd1.Z(i)=0.3*unom.Z(i)*unom.Z(i)*cosfi.Z(i)/pnom.Z(i)
            qmax.Z(i)=uzl.Cols("qmax").z(j)
            qmin.Z(i)=uzl.Cols("qmin").z(j)
        end if
     end if
End Sub