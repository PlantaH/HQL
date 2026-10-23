<%
Function CargarDiccionario()
	Set XMLdoc = Server.CreateObject("Microsoft.XMLDOM")
	XMLdoc.async = false	
	XMLdoc.load(Application("MapPath")&"\includes\diccionario.xml")	
	Application("XMLDictionario") = XMLdoc.xml	
End function

Function CargarDiccionarioDesdeAdmin()
	Set XMLdoc = Server.CreateObject("Microsoft.XMLDOM")
	XMLdoc.async = false	
	XMLdoc.load("D:\ASP\hotelypago\includes\diccionario.xml")	
	Application("XMLDictionario") = XMLdoc.xml	
End function

if request("adm")="1" then
	RecargarDiccionario = 1
	if RecargarDiccionario or isEmpty(Application("XMLDictionario")) then
		CargarDiccionarioDesdeAdmin()
	end if
else
	RecargarDiccionario = 1
	if RecargarDiccionario or isEmpty(Application("XMLDictionario")) then
		CargarDiccionario()
	end if
end if

Set XMLdiccionario = Server.CreateObject("Microsoft.XMLDOM")
XMLdiccionario.async = false
XMLdiccionario.loadXML(Application("XMLDictionario"))

function TratarStr(str)
	Dim texto
	texto = replace(str,"[","<")	
	texto = replace(texto,"]",">")
	TratarStr = texto
end function

function Traduccion(str)
	dim s, l, Nodo
	s = replace(lcase(str)," ","")
	l = setLang
	
	if l="" then l= "es"'application("IdiomaPrincipal")
	Set Nodo = XMLdiccionario.documentElement.selectSingleNode(s & "/" & l)	
	if not Nodo is nothing then
		'if isRss then
		'	Traduccion = "<acronym style=color:#009;>" & toXHTML(unescape(TratarStr(Nodo.Text))) & "</acronym>"
		'else
			Traduccion = (unescape(TratarStr(Nodo.Text)))
		'end if
		
	else
		'if not isRss then
		'	Traduccion = "<acronym style=color:#f00>" & str & "</acronym>"
		'else
			Traduccion = str
		'end if
	end if
end function

function TraduccionJS(str)
	dim s, l, Nodo
	s = replace(lcase(str)," ","")
	l = Idioma
	'if l="" then l="es"
	Set Nodo = XMLdiccionario.documentElement.selectSingleNode(s & "/" & l)
	if not Nodo is nothing then
		TraduccionJS = replace(unescape(TratarStr(Nodo.Text)),"'","\'")
	else
		TraduccionJS = str
	end if
end function
%>