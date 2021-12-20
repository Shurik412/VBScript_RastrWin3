Set t=Rastr
Dim arr(200)
'Call com_cxema("ny", 0)
'Set com_cxema_ = t.Tables("com_cxema_")

t.LockEvent=True

print(review_name_table("com_cxema"))


Function AllTablesRastr(index)
    Call create_array(arr)
    AllTablesRastr = arr(index)
End Function

Sub create_array(arr)
    Set Tabs = Rastr.Tables 
    For i = 0 to Tabs.Count-1
        arr(i) = Tabs(i).Name
    Next  
End Sub

Function review_name_table(name_table)
    For index=0 to UBound(AllTablesRastr)
        name_tables_rastr = AllTablesRastr(index)
        if name_table = name_tables_rastr then
            review_name_table = name_tables_rastr
        end if
    next
    if review_name_table = "" then
        print("Таблица <" & name_table & "> отсутствует в RastrWin3")
    end if
End Function

Function com_cxema(cols_tb, z)
    Set com_cxema = t.Tables("com_cxema")
    Set ny_com_cxema = com_cxema.Cols("ny")
    print(com_cxema.Cols(cols_tb).Z(z))
End Function 

Function print_param(tables, cols, z)
    if tables<>"" then 
        if cols <> "" then
            if z <> "" then
                Set table = t.Tables(table)
				t.Printp("Таблица: " & table.Name)
				t.Printp("Параметр: " & cols)
				t.Printp("Значение: " & table.Cols(cols).z(z))
            end if
        end if
    end if
End Function 

Sub print(str)
    t.Printp(str)
End Sub