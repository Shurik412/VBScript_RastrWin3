'********************************************************
'Урок VBScript №15:
'Объект-Коллекция Dictionary или словарь
'file_1.vbs
'********************************************************
Dim Dict
Set Dict = CreateObject("Scripting.Dictionary")
Dict.Add 0, "Ноль"
Dict.Add 1, "Один"
Dict.Add "Строка", "Строка"
Dict.Add "Строка2", 4