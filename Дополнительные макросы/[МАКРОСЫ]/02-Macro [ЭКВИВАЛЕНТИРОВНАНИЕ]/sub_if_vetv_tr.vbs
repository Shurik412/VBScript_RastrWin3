Set t=Rastr
Set vetv = t.Tables("vetv")

vetv.Calc("sel=1")

Call If_Vetv_Tr_otkl()



Sub If_Vetv_Tr_otkl()
    Set vetv = t.Tables("vetv")
    Set node = t.Tables("node")
    Set gen = t.Tables("Generator")
    MaxRowVetv = vetv.Count
    For i=0 to MaxRowVetv-1
        type_vetv = vetv.Cols("tip").Z(i)
        if type_vetv = 1 then
			 
            ny_ip = vetv.Cols("ip").Z(i)
            ny_iq = vetv.Cols("iq").Z(i)
            v_ip = vetv.Cols("v_ip").Z(i)
            v_iq = vetv.Cols("v_iq").Z(i)
			
            node.SetSel("ny=" & ny_ip)
            j_ny_ip = node.FindNextSel(-1)
			
            if j_ny_ip <>(-1) Then
                tip_ny_ip = node.Cols("tip").Z(j_ny_ip) ' тип узла
                if tip_ny_ip > 1 Then
                    vetv.Cols("sel").Z(i) = 0
                end if
            end if 
             
            node.SetSel("ny=" & ny_iq)
            j_ny_iq = node.FindNextSel(-1)
			
            if j_ny_iq <>(-1) Then
                tip_ny_iq = node.Cols("tip").Z(j_ny_iq) ' тип узла
                if tip_ny_iq > 1 Then
                    vetv.Cols("sel").Z(i) = 0
                end if
            end if
             
            gen.SetSel("Node=" & ny_ip)
            j_gen_ny_ip = gen.FindNextSel(-1)
			
            if j_gen_ny_ip<>(-1) Then
                vetv.Cols("sel").Z(i) = 0
                node.SetSel("ny=" & ny_ip)
                j_node_ip = node.FindNextSel(-1)
				
                if j_node_ip <> (-1) then
                    node.Cols("sel").Z(j_node_ip) = 0
                end if
            end if
             
            gen.SetSel("Node=" & ny_iq)
            j_gen_ny_iq = gen.FindNextSel(-1)
			
            if j_gen_ny_iq <>(-1) Then
                vetv.Cols("sel").Z(i) = 0 
                node.SetSel("ny=" & ny_iq)
                j_node_iq = node.FindNextSel(-1)
                if j_node_iq <> (-1) then
                    node.Cols("sel").Z(j_node_iq) = 0
                end if
            end if
        end if 
    next
End sub