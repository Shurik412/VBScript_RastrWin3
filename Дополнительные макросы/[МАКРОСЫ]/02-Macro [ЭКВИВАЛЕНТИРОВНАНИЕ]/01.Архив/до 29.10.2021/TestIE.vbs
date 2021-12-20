'22.07.2014 исправлен под ie11

'заполнение comboBox...
Dim TextSS, TextSS2

set TablArea = Rastr.Tables("node")
Set ColArea = TablArea.Cols("na")
Set NameArea = TablArea.Cols("name")



TextSS = "<SELECT NAME = SelectOptionDrop SIZE = 1 ID = SelectOptionDrop ONCHANGE = OnBtnProv()> "

For i=0 to Rastr.Tables("area").Count-1 
	TextSS = TextSS + "<OPTION NAME  = Option_"+ColArea.ZS(i)+" VALUE = "+ColArea.ZS(i)+" > № "+ColArea.ZS(i)+" ( "+NameArea.ZS(i)+" )</OPTION>"
Next


TextSS = TextSS + "</SELECT>"

set TablArea2 = Rastr.Tables("area2")
Set ColArea2 = TablArea2.Cols("npa")
Set NameArea2 = TablArea2.Cols("name")
TextSS2 = "<SELECT NAME = SelectOptionDrop SIZE = 1 ID = SelectOptionDrop ONCHANGE = OnBtnProv()> "
For i=0 to Rastr.Tables("area2").Count-1 
	TextSS2 = TextSS2 + "<OPTION NAME  = ""Option_"+ColArea2.ZS(i)+""" VALUE = "+ColArea2.ZS(i)+" > № "+ColArea2.ZS(i)+" ( "+NameArea2.ZS(i)+" )</OPTION>"
Next
TextSS2 = TextSS2 + "</SELECT>"




'динамическая страница HTML
htmlDialog = _
"<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.01 Transitional//EN"">"+vbCrLf+_
"<html>"+vbCrLf+_
"<head>"+vbCrLf+_
"<meta http-equiv=""Content-Type"" content=""text/html"" charset=""windows-1251"">"+vbCrLf+_
"<title>Задать потребление района / территории</title>"+vbCrLf+_
"<STYLE type=""text/css"">"+vbCrLf+_
"	body { margin: 3%;}"+vbCrLf+_
"	DIV {text-align: left}"+vbCrLf+_
"</STYLE>"+vbCrLf+_
"</head>"+vbCrLf+_
" <SCRIPT type=""text/javascript"">"+vbCrLf+_
"   var g_Rastr = null"+vbCrLf+_
"	var TextVar"+vbCrLf+_
"	var TextVar2"+vbCrLf+_
""+vbCrLf+_
"	function InputRadioOnClick(Text){"+vbCrLf+_
"		document.getElementById(""SelectOptionDrop"").outerHTML = Text"+vbCrLf+_
"		OnBtnProv()"+vbCrLf+_
"	}"+vbCrLf+_
""+vbCrLf+_
"	function SetRastr(obj){"+vbCrLf+_
"		g_Rastr = obj"+vbCrLf+_
"	}"+vbCrLf+_
"	function OnBtnProv(){"+vbCrLf+_
"		var objSel = document.formHTML.SelectOptionDrop"+vbCrLf+_
"		raion = objSel.options[objSel.selectedIndex].value"+vbCrLf+_
"		if(document.getElementById(""ID_1"").checked){"+vbCrLf+_
"			TextLabel = Math.round(g_Rastr.Calc(""val"",""area"",""pop"",""na=""+raion)*10)/10}"+vbCrLf+_
"		else{"+vbCrLf+_
"			TextLabel = Math.round(g_Rastr.Calc(""val"",""area2"",""pop"",""npa=""+raion)*10)/10}"+vbCrLf+_
"		document.getElementById(""rastrP"").innerHTML = TextLabel"+vbCrLf+_
"	}"+vbCrLf+_
" </SCRIPT>"+vbCrLf+_
""+vbCrLf+_
"<body SCROLL=""NO"" >"+vbCrLf+_
"<DIV>"+vbCrLf+_
"<form name=""formHTML"" onsubmit=""return false;"">"+vbCrLf+_
" <H2> Задать потребление района / территории</H2>"+vbCrLf+_
"<HR>"+vbCrLf+_
" <INPUT NAME=""InputRadio"" ID=ID_1 TYPE=""radio"" CHECKED Onclick = ""InputRadioOnClick(TextVar)""><label FOR = ID_1> Район </label><BR>"+vbCrLf+_
" <INPUT NAME=""InputRadio"" ID=ID_2 TYPE=""radio"" Onclick = ""InputRadioOnClick(TextVar2)""><label FOR = ID_2> Территория </label><BR>"+vbCrLf+_
"<p><label>Введите номер района (территории):</label>"+vbCrLf+_
"  &nbsp;&nbsp;&nbsp"+vbCrLf+_	
" <SELECT NAME = ""SelectOptionDrop"" SIZE = 1 ID = ""SelectOptionDrop"" ONCHANGE=""OnBtnProv()"">"+vbCrLf+_
"   </SELECT>"+vbCrLf+_
"</p>"+vbCrLf+_
"Текущее потребление: <label id=""rastrP"">н/д</label>"+vbCrLf+_
"<p><label>Введите новое потребление:</label>&nbsp;"+vbCrLf+_
"  <input type=""text"" name=""InputPop"">"+vbCrLf+_
"</p>"+vbCrLf+_
"<INPUT NAME=""CheckBoxReact""  TYPE = ""checkbox"" ID = ID_INPUT1 CHECKED > <LABEL FOR = ID_INPUT1 > менять реактивное потребление пропорционально </LABEL>"+vbCrLf+_
"<p><label>Дополнительная выборка: </label>&nbsp;&nbsp;&nbsp;&nbsp;"+vbCrLf+_
"  <input type=""text"" name=""InputSel"">"+vbCrLf+_
"</p>"+vbCrLf+_
"  <BUTTON name=""BtnOK""> Расчет </BUTTON>"+vbCrLf+_
"&nbsp"+vbCrLf+_
"  <BUTTON name=""BtnCancel""> Закрыть </BUTTON>"+vbCrLf+_
"</form>"+vbCrLf+_
"<DIV>"+vbCrLf+_
""+vbCrLf+_
"</body>"+vbCrLf+_
"</html>"





Label = TRUE	
SET g_oIE = CreateObjectEx("InternetExplorer.Application","g_IE_")

r=setlocale("en-us")



Sub CorPotr(raion,potr,reac_p,RadioCheck,Sel)
	max_it=10   ' максимальное число итераций
	eps=0.0001   ' точность расчета
	Set pnode=Rastr.Tables("node")
	IF RadioCheck Then
		IF Sel <> "" Then
			pnode.SetSel("na="&raion & "&" &sel)
		Else
			pnode.SetSel("na="&raion)
		End If
	Else
		IF Sel <> "" Then
			pnode.SetSel("npa="&raion & "&" &sel)
		Else
			pnode.SetSel("npa="&raion)
		End If
	End If
	Set pn=pnode.Cols("pn")
	Set qn=pnode.Cols("qn")
	set st=Rastr.stage ("Коррекция потребления в районе=" & raion)
	st.Log LOG_INFO,"Задано потребление=" & potr
	for i=1 to max_it 
		IF RadioCheck Then
			pop=Rastr.Calc("val","area","pop","na="&raion)
		Else
			pop=Rastr.Calc("val","area2","pop","npa="&raion)
		End If
		koef=potr/pop
		st.Log LOG_INFO,"Текущее потребление =" & pop
		st.Log LOG_INFO,"Отношение заданное/текущее =" & koef
		if( abs(koef -1) > eps) then
			pn.Calc("pn*"&koef)
			if(reac_p)	then qn.Calc("qn*"&koef)
			kod=Rastr.rgm("")
			if(kod <> 0) then
				st.Log LOG_ERROR,"---------Аварийное завершение расчета режима----------- "
				exit sub
			end if
		else exit sub
		end if
	next
end Sub




Sub CloseHtml
	g_oIE.Quit()
	ExitDo = True
End Sub





Sub Html_OnUnload
    ExitDo = True
End Sub





Sub Calculate
	n = g_oIE.Document.formHTML.SelectOptionDrop.Value
	pop = g_oIE.Document.formHTML.InputPop.Value
	reac_p = g_oIE.Document.formHTML.CheckBoxReact.Checked
	RadioCheck = g_oIE.Document.getElementById("ID_1").Checked
	Sel = CStr(g_oIE.Document.formHTML.InputSel.Value)

	if (n <> "") and (pop <> "") then			'расчет
		CorPotr n+0,pop+0,reac_p, RadioCheck, Sel
		g_oIE.Document.Script.OnBtnProv()
	end if
end sub




SUB g_IE_Quit(a)
	ExitDo = True
END SUB





ExitDo = FALSE

g_oIE.TheaterMode = FALSE
g_oIE.Left      = 250   'коррдината верхнего левого угла окна IEx
g_oIE.Top       = 250   'координата верха окна IE
g_oIE.Height    = 430   'высота окна IE
g_oIE.Width     = 550   'ширина окна IE
g_oIE.MenuBar   = FALSE 'без меню IE
g_oIE.ToolBar   = FALSE 'без тулбара IE
g_oIE.StatusBar = FALSE 'без строки состояния IE

g_oIE.Navigate  "about:blank"	'пустая страница

'выжидаем пока IE не освободится
DO WHILE ( g_oIE.Busy )
	SLEEP 100
LOOP
g_oIE.Document.write ( htmlDialog )	'загрузка динамич. страницы
g_oIE.document.body.onunload = GetRef("Html_OnUnload")
g_oIE.document.formHTML.BtnCancel.onclick = GetRef("CloseHtml")
g_oIE.document.formHTML.BtnOk.onclick = GetRef("Calculate")
g_oIE.Document.Script.SetRastr(Rastr) 	
g_oIE.Document.Script.TextVar = TextSS
g_oIE.Document.Script.TextVar2 = TextSS2
g_oIE.document.getElementById("SelectOptionDrop").outerHTML = TextSS
g_oIE.Document.Script.OnBtnProv()
g_oIE.Visible = True	'показать IE

'завершаем работу с IE
Do ' ожидание закрытия окна IE
    Sleep 1000
Loop Until ExitDo

SET g_oIE=NOTHING
