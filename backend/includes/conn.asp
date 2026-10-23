
<!-- #include file="../_nohack.asp" -->
<% 
Server.ScriptTimeout=360

Dim conn, idioma

Set conn = Server.CreateObject("ADODB.Connection")
conn.ConnectionString = Application("String")
conn.open
conn.CommandTimeout = 960


''''FUNCIONES'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Function Traduccion(str)
'	Traduccion = str
'End Function 

''Language seting
if request("lang")<>"" then
	session("idioma") = "es"
	session("set_ogLocale") = "es_LA"
	if request("lang") = "en" then	
		session("idioma") = "en"
		session("set_ogLocale") = "en_US"
	end if
end if

if session("idioma")<>"" then
	setLang = session("idioma")
	set_ogLocale = session("set_ogLocale")
else
	setLang = "es"
	set_ogLocale = "es_LA"
end if

PER_Cons = false
PER_Alta = false
PER_Baja = false
PER_Modi = false

public Sub Permisos(PaginaID)
	 
	Set PER = conn.execute( "exec Permisos_Check " & session("UsuarioID") & "," & PaginaID )
	if not PER.eof then	
		PER_Cons = PER("Consulta")
		PER_Alta = PER("Alta")
		PER_Baja = PER("Baja")
		PER_Modi = PER("Modi")
	else
		PER_Cons = false
		PER_Alta = false
		PER_Baja = false
		PER_Modi = false
	end if
	
end sub	


Function crearMarkUp(co)	
	crearMarkUp = formatNumber(1-(co/100),2)
End Function


function ConvertText(str)
	ConvertText = replace(str , "'", "")	
end function
Function Convert(valor)
	Convert = replace(valor,".",",")
end function

Function ReConvert(valor)
	ReConvert = replace(valor,",",".")
end function

function VerificarNull(valor)
	dim cnull 
	cnull=null	
	if trim(valor)<>"" then 
		if isNumeric(valor) then 
			cnull=replace(valor,".",",")
		else
			cnull=valor
		end if
	end if
	VerificarNull = cnull
end function

Function FormatPrecio(pre)
	FormatPrecio = formatNumber(pre,0)
End Function

Function soloLogueado()
	Dim rQuery
	if not logueado then
		if request("nl")<>"" then
			rQuery = "?" & Request.ServerVariables("QUERY_STRING")
			if rQuery="?nl=1" then
				rQuery = ""
			elseif instr(rQuery, "&nl=1")>0 then
				rQuery = replace(rQuery, "&nl=1", "")
			end if
				
			session("redirectOff") = NombreDelAsp() & rQuery
			
		end if
		session("expiro") = true
		response.redirect(Application("Dominio"))
	end if
End Function

Function GetHTML(url)
	Dim xml, strData
	Set xml = Server.CreateObject("MSXML2.ServerXMLHTTP")
	xml.Open "GET", url, true
	Call xml.Send()
	
	On Error Resume Next
	
	If xml.readyState <> 4 then
		xml.waitForResponse 3
	End If
	
	If Err.Number <> 0 then
		strData = "."
	Else
		If (xml.readyState <> 4) Or (xml.Status <> 200) Then
		  
			xml.Abort
				
			strData = "..."
		Else
			strData = xml.ResponseText
		End If
	End If
	
	GetHTML = strData
End Function

Function ArmarFecha(Fecha)
	ArmarFecha = right(Fecha,2) & "/" & mid(Fecha,5,2) & "/" & left(Fecha,4)
End Function

Function ArmarFechaG(Fecha)
	ArmarFechaG = left(Fecha,4) & "-" & mid(Fecha,5,2) & "-" & right(Fecha,2)
End Function

Function Fecha2US(f)
	Dim fe, v
	fe = f
	if instr(fe , "/")>0 then v = split(fe,"/")
	if instr(fe , "-")>0 then v = split(fe,"-")	
	Fecha2US = v(1) & "/" & v(0) & "/" & v(2)
End Function

Function Fecha2Int(f)
	Dim fe, v
	fe = f
	if instr(fe , "/")>0 then v = split(fe,"/")
	if instr(fe , "-")>0 then v = split(fe,"-")
	Fecha2Int = v(2) * 10000 + v(1)*100 + v(0)
End Function

Function FechaHoy()
	arrayMeses = array("Enero","Febrero","Marzo","Abril","Mayo","Junio","Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre")
	arraySemana = array("Domingo","Lunes","Martes","Miercoles","Jueves","Viernes","Sabado")
	FechaHoy = arraySemana(weekday(date)-1)& ", " & day(date) & " de " & arrayMeses(round(month(date)-1)) & " " & year(date)
End Function
 
Function FormatearFecha(fecha,idioma)
	if instr(fecha, "/")=0 then
		fecha = right(Fecha,2) & "/" & mid(Fecha,5,2) & "/" & left(Fecha,4)
	end if
	if idioma="" then
		arrayMeses = array("Enero","Febrero","Marzo","Abril","Mayo","Junio","Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre")	
		fechaFormateada = Round(left(fecha,2)) & " de " & arrayMeses(mid(fecha,4,2)-1) & " " & right(fecha,4)
	else
		arrayMeses = array("January","February","March","April","May","June","July","August","September","October","November","December")	
		dia=int(left(Fecha,2))
		if dia=1 or dia=31 then 
			dia = dia & "st "
		elseif dia=2 then 
			dia = " 2nd "
		elseif dia=3 then 
			dia = " 3rd "				
		elseif dia=21 then 
			dia = " 21st "
		elseif dia=22 then 
			dia = " 22nd "
		elseif dia=23 then 
			dia = " 23rd "
		else 
			dia = dia & "th " 
		end if
		fechaFormateada = arrayMeses(mid(fecha,4,2)-1) & " " & dia & right(fecha,4)
	end if		
	FormatearFecha = fechaFormateada
End Function

Function Capitalizar(myString)
   sArray = Split(myString, " ")
   For Each sWord in sArray
      Capitalizar = Capitalizar & UCase(Left(sWord, 1)) & LCase(Mid(sWord, 2)) & " "
   Next
   Capitalizar = trim(Capitalizar)
End Function

Function noErrorSQL(str)
    noErrorSQL = Replace(Trim(str), "'", "''")
End Function

Function noErrorJS(str)
	Dim st
	st = str
	st = Trim(str)
	st = Replace(st, chr(10), "")
	st = Replace(st, chr(13), "")
	st = Replace(st, "'", "\'")
	st = Replace(st, """", "\'")
    noErrorJS = st
End Function

Function BuscarConAcentos(st)
	Dim str
	str = st
	str = Trim(str)
	str = Lcase(str)
	str = Replace(str, "[", "")
	str = Replace(str, "]", "")
	str = Replace(str, "'", "''")

	str = Replace(str, "ã", "a")
	str = Replace(str, "â", "a")
	str = Replace(str, "à", "a")
	
	str = Replace(str, "ç", "c")
	
	str = Replace(str, "ê", "e")
	
	str = Replace(str, "õ", "o")
	str = Replace(str, "ô", "o")
	
	str = Replace(str, "ü", "u")
	
	str = Replace(str, "a", "á")
    str = Replace(str, "e", "é")
    str = Replace(str, "i", "í")
    str = Replace(str, "o", "ó")
    str = Replace(str, "u", "ú")	
		
	str = Replace(str, "á", "[ãâàáa]")
	str = Replace(str, "c", "[çc]")
    str = Replace(str, "é", "[êée]")
    str = Replace(str, "í", "[íi]")
    str = Replace(str, "ó", "[óo]")
    str = Replace(str, "ú", "[üúu]")	
	
    BuscarConAcentos = str
End Function


Function toXHTML(str)
	dim st, i, tempV, vX
	st = str
	if instr(st,"<")>0 then
		st = replace(replace(st,chr(10),""),chr(13),"")
	else		
		st = replace(st,chr(10),"<br />")
	end if
	
	st = Server.HTMLEncode(st)
	vX = Array("&lt;br&gt;I<br />","&lt;BR&gt;I<br />","&lt;br /&gt;I<br />")	
	for each i in vX
		tempV = split(i,"I")
		st = Replace(st, tempV(0), tempV(1))
	next
	toXHTML = st
End Function


Function FixToXHTML(str)
	dim st
	st = stripHTMLtags(str)
	st = toXHTML(st)
	FixToXHTML = st
End Function

Function selectedHTM(str,str2)
	if str=str2 then
		selectedHTM=" selected=""selected"""
	end if
End Function

Function RestarFecha(fecha, fecha2)
	fv = split(fecha, "/")
	ffv = split(fecha2, "/")
	RestarFecha = dateDiff("d", fv(1) & "/" & fv(0) & "/" & fv(2), ffv(1) & "/" & ffv(0) & "/" & ffv(2))
End function

Function SumarFecha(fecha, dia)
	Dim fv, ffv, d, m, a
	fv = split(fecha, "/")	
	ffv = dateAdd("d", dia, fv(1) & "/" & fv(0) & "/" & fv(2))
	ffv = split(ffv, "/")
	d = ffv(1)
	if d<10 then d = "0" & ffv(1)
	m = ffv(0)
	if m<10 then m = "0" & ffv(0)
	a = ffv(2)	
	SumarFecha = d & "/" & m & "/" & a
End function

Function Resumir(str, length)
	If Len(str) < length Then
		Resumir = str
	Else
		If Mid(str, length, 1) = " " Then
			strTemp = Left(str, length - 1)
		Else
			strTemp = Left(str, length)
			For i = length + 1 To Len(str)
				If Mid(str, i, 1) = " " Then Exit For
				strTemp = strTemp & Mid(str, i, 1)
			Next
		End If
		Resumir = strTemp & "..."
	End If
End Function

function nombreDelASP()
	Dim nds
	nds = Request.ServerVariables("URL")
	for i = len(nds) -1 to 0 step -1
		if mid(nds, i, 1) = "/" then exit for
	next
	nombreDelASP = lcase(right(nds, len(nds) - i))
end function

Function stripHTMLtags(HTMLstring)
	dim str
	str = HTMLstring
	if len(str) > 0 then
		Set RegularExpressionObject = New RegExp	
		With RegularExpressionObject
			.Pattern = "<[^>]+>"
			.IgnoreCase = True
			.Global = True
		End With	
	
		stripHTMLtags = sacarP(RegularExpressionObject.Replace(str, ""))
	
		Set RegularExpressionObject = nothing
	else
		stripHTMLtags = sacarP(str)
	end if			
End Function

Function sacarP(t)	
	dim texto
	
	texto = trim(t)
	
	if len(texto) = 6 and texto = "&nbsp;" then 
		sacarP = ""
	elseif len(texto) = 13 and texto = "<p>&nbsp;</p>" then 
		sacarP = ""
	elseif len(texto) = 8 and texto = "<p> </p>" then 
		sacarP = ""
	else
		sacarP = texto
	end if	
	
end function

Dim pagePrev
pagePrev = Request.ServerVariables("HTTP_REFERER")
if pagePrev = "" then pagePrev = "#"

Function ceil(n)
	'Dim numR, numDec
    'numR = clng(n)
    'numDec = n - numR
    'if numDec>0 then numR = numR + 1
 
	if n<>"" then 
	    ceil = round(n)
	else
		ceil = n
	end if
End Function

Function Encriptar(str)    
    Dim vLongitud, vCaracter, vCodificado
    vLongitud = Len(str)
    vCodificado = ""
    For vCaracter = 1 To vLongitud
        vCodificado = vCodificado & Chr(Asc(Mid(str, vCaracter, 1)) + 2)
    Next
    Encriptar = vCodificado
End Function

Function Desencriptar(str)
    Dim vLongitud, vCaracter, vCodificado
    vLongitud = Len(str)
    vCodificado = ""
    For vCaracter = 1 To vLongitud
       vCodificado = vCodificado & Chr(Asc(Mid(str, vCaracter, 1)) - 2)
    Next
    Desencriptar = vCodificado
End Function

Function isRSS()
    if instr(request.ServerVariables("REMOTE_ADDR"),"192.168")>0 then
		isRSS = true
	else
		isRSS = false
	end if
End Function

Function RPM(q,s)
	Dim sEmail, envioemail
	set sEmail = server.createobject("RSS_Email.RSS_Email")		
	sEmail.DataSMTP Application("ServerAddress"),Application("Sender"), Application("SenderName") , Application("RSSemailsender"), Application("RSSemailclave"), Application("RSSemailpuerto"), Application("RSS_SSL")
	sEmail.Asunto = "Reporte"
	sEmail.Cuerpo  = s
	
	sEmail.AddDestinatario cstr(q & "@rss.com.ar")
	
	envioemail=sEmail.Enviar()	
	Set sEmail = Nothing
End Function

''Imagenes
Function FotoPathPrevio(f)
	dim fo, vfo, ffo, fi
	fo = f
	if isNull(fo) then fo = ""
	vfo = split(fo, "/")
	
	for fi=0 to ubound(vfo)
		sfo = "/"
		if fi = ubound(vfo)-1 then sfo = "/previo/"
		if fi = ubound(vfo) then sfo = ""
		ffo = ffo & vfo(fi) & sfo
	next
	
	FotoPathPrevio = ffo
End Function 

Function FotoPathDestacada(f)
	dim fo, vfo, ffo, fi
	fo = f
	if isNull(fo) then fo = ""
	vfo = split(fo, "/")
	
	for fi=0 to ubound(vfo)
		sfo = "/"
		if fi = ubound(vfo)-1 then sfo = "/destacada/"
		if fi = ubound(vfo) then sfo = ""
		ffo = ffo & vfo(fi) & sfo
	next
	
	FotoPathDestacada = ffo
End Function 

Function FotoPathMediana(f)
	dim fo, vfo, ffo, fi
	fo = f
	if isNull(fo) then fo = ""
	vfo = split(fo, "/")
	
	for fi=0 to ubound(vfo)
		sfo = "/"
		if fi = ubound(vfo)-1 then sfo = "/mediana/"
		if fi = ubound(vfo) then sfo = ""
		ffo = ffo & vfo(fi) & sfo
	next
	
	FotoPathMediana = ffo
End Function 

Function FotoPathGrande(f)
	dim fo, vfo, ffo, fi
	fo = f
	if isNull(fo) then fo = ""
	vfo = split(fo, "/")
	
	for fi=0 to ubound(vfo)
		sfo = "/"
		if fi = ubound(vfo)-1 then sfo = "/grande/"
		if fi = ubound(vfo) then sfo = ""
		ffo = ffo & vfo(fi) & sfo
	next
	
	FotoPathGrande = ffo
End Function 


Function Traducir(Tabla,Campo,id,Texto)
	'texto=Texto & ""
	'if session("Idioma")<>"" then
	'	sSql = "exec Sistema_Traducir '" & Tabla & "','" & Campo & "','IN'," & id
	'	Set tra = conn.Execute (sSql)
	'
	'	If tra.eof Then
	'		Traducir = ""
	'	Else
	'		Traducir = toXHTML(stripHTMLtags(tra("texto")))
	'	End If		
	'else
	'	Traducir = toXHTML(stripHTMLtags(Texto))
	'end 
	Traducir = toXHTML(stripHTMLtags(Texto))
End Function


Function TraducirH(Tabla,Campo,id,Texto)
	texto=Texto & ""
	'if session("Idioma")<>"" then
	'	sSql = "exec Sistema_Traducir '" & Tabla & "','" & Campo & "','IN'," & id
	'	Set tra = conn.Execute (sSql)
	'
	'	If tra.eof Then
	'		TraducirH = ""
	'	Else			
	'		TraducirH = tra("texto")		
	'	End If		
	'else		
	'	TraducirH = Texto	
	'end if
	TraducirH = Texto
End Function


Function AntiSpamEncode(str)
    Dim strOut, i
    strOut = ""
    For i = 1 To Len(str)
        strOut = strOut & "&#" & ASC(Mid(str, i, 1)) & ";"
    Next
    AntiSpamEncode = strOut
End function


Function FixURL(str)
	url = trim(str)
	if instr(url,"@")>0 and instr(url,"mailto:")<1 then
		FixURL = "mailto:" & url
	elseif instr(url,"http://")<1 and instr(url,"https://")<1 then
		FixURL = "http://" & url
	else 
		FixURL = url
	end if
End Function

Function FixThumbHost(i,p)
	Dim s
	select case p
		case "GTA"
			s = replace(i,"th/","/")
			
		case "TOURICO"
			s = replace(i,"_Thumbnailed.jpg",".jpg")
			
		case "HOTELBEDS"
			s = replace(i,"/small/","/")
			
		case "MISHOTELES"
			s = FotoPathMediana(i)
			
		case else
			s = i
		
	end select
	
	FixThumbHost = s
End Function

Function QRencoder(t,s)
	If s>0 and Len(t)>1 Then
		QRencoder = "http://chart.apis.google.com/chart?cht=qr&chs=" & s & "x" & s & "&chl=" & t
	Else
		QRencoder = "error"
	End If
End Function


Function isBrowser(b)
	if Instr(Lcase(Request.ServerVariables("HTTP_USER_AGENT")), b)>0 then
		isBrowser = true
	else
		isBrowser = false
	end if
End Function

Function getVersionIE()
	Dim str, strPos
	str = Lcase(Request.ServerVariables("HTTP_USER_AGENT"))
	strPos = Instr(str, "msie")
	if strPos>0 then
		getVersionIE = Mid(str, strPos+5, 1)
	else
		getVersionIE = 1000
	end if
End Function

''''''''''''''''

Dim versionIE
versionIE = getVersionIE()

Dim bodyCSSfix
if versionIE<1000 then bodyCSSfix = " class=""fckie ie" & versionIE & """"

Dim FechaServer
FechaServer = right(day(date)+100,2) & "/" & right(month(date)+100,2) & "/" & year(date)

PathIMG=""
if idioma <> "es" then PathIMG = idioma & "/"

Function RGB2Hex(c)
	Dim s
	s = Replace(Trim(c),"#", "")
	RGB2Hex = CLng("&H" & Left(s, 2)) & "," & CLng("&H" & Mid(s, 3, 2)) & "," & CLng("&H" & Right(s, 2))
End Function

function Crear_PDF(link,Servicio)

	rotate = false
	
	pdfNombre = Servicio & "-" & year(date) & month(date) & day(date()) & hour(time) & minute(time) & second(time)
		
	max=9
	min=2
	
	Randomize
	nocache = int((max-min+1)*rnd+min)
					
	pdfPath = Server.MapPath("pdf_temp/" & pdfNombre & ".pdf")
	
	'Set Pdf = Server.CreateObject("Persits.Pdf")
	'Set Doc = Pdf.CreateDocument
	'Doc.Title = Unescape(pdfNombre)
	'Doc.Creator = "HotelesB2B"
	'Doc.ImportFromUrl link, "LeftMargin=10, RightMargin=10, TopMargin=30, BottomMargin=30"
	
	'Doc.Save pdfPath
	
	
	Set pdfObj = Server.CreateObject("ABCpdf6.Doc")
	pdfObj.Rect.Inset 10, 10
	pdfID = pdfObj.AddImageUrl(link)

	Do
		If Not pdfObj.Chainable(pdfID) Then Exit Do
		pdfObj.Page = pdfObj.AddPage()
		pdfID = pdfObj.AddImageToChain(pdfID)
	Loop

	For i = 1 To pdfObj.PageCount
		pdfObj.PageNumber = i
		pdfObj.Flatten
	Next

	''''
	pdfObj.Save(pdfPath)
	
	
	Crear_PDF = pdfNombre
	
end function	

function EnviarXML (ReservaID)

	Set connXML = Server.CreateObject("ADODB.Connection")
	connXML.ConnectionString = Application("String")
	connXML.open

	connXML.execute("exec RES_GenerarXML " & ReservaID)
	connXML.close
	Set connXML = nothing
	
	EnviarXML = true
end function	



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function Control_HOST(sistemaID,TripPlannerID)
	set sis = conn.Execute("select * from sistemas_thread where sistemaID=" & sistemaID )
	
	if TripPlannerID = "" then TripPlannerID  = 0
	
	Set Cmd = Server.CreateObject("ADODB.Command")
	Cmd.Activeconnection = conn
	Cmd.CommandText = "Control_HOST_Alta"
	Cmd.CommandType = 4
								
	Cmd.parameters.refresh				
		
	Cmd(1).Direction = 2
	Cmd(1).Type = 3	
	Cmd(2) = session("CLienteID")
	Cmd(3) = sistemaID
	Cmd(4) = "Sistema:" & sis("sistema") & "<br>Hora:" & now()
	Cmd(5) = TripPlannerID
					
	Cmd.Execute , , adExecuteNoRecords
	ID = Cmd(1).Value
	
	Control_HOST = ID		
end function

function Control_HOST_Fin(id,sistemaID,BookingCode)
	set sis = conn.Execute("select * from sistemas_thread where sistemaID=" & sistemaID )	
	
	Set Cmd = Server.CreateObject("ADODB.Command")
	Cmd.Activeconnection = conn
	Cmd.CommandText = "Control_HOST_Fin"
	Cmd.CommandType = 4
								
	Cmd.parameters.refresh				
		
	Cmd(1) = id
	Cmd(2) = "Sistema:" & sis("sistema") & "<br>Hora:" & now()	
	Cmd(3)	= BookingCode
	Cmd.Execute , , adExecuteNoRecords		
	
	Control_HOST_Fin = id
end function

Function GetJsonFromRs(RS,Name)
    Dim JSON
    if Name<>"" then
		JSON = "{" & """" & Name & """:["  
	else
		JSON = "["  
	end if
	
    Do While NOT RS.EOF
        JSON = JSON & "{"
        For x = 0 To RS.Fields.Count - 1
            IF x < RS.Fields.Count - 1 THEN
                JSON = JSON & """" & RS(x).Name & """:""" & toxhtml(RS(x)) & ""","
            ELSE
                JSON = JSON & """" & RS(x).Name & """:""" & toxhtml(RS(x)) & """"
            END IF          
        Next      
    RS.MoveNext
        IF NOT RS.EOF THEN
            JSON = JSON & "},"
        ELSE
            JSON = JSON & "}"
        END IF
    Loop  
    JSON = JSON & "]"
    GetJsonFromRs = JSON
End Function

Function CalcularComision(t,valor)
	
	if t = "Requerimiento" then CalcularComision = FormatearNumero(valor * session("Cliente_Comision_Requerimiento")/100)
	if t = "Deposito" then CalcularComision = FormatearNumero(valor * session("Cliente_Comision_Deposito")/100)

end function

function CalcularSaldo(m,t)
	dim Monto, MontoTotal, Cargo, Comision
	Monto = m
	
	Cargo=0
	Comision = 0
	if t = "R" then
		Cargo = Monto * session("Cliente_Comision_Requerimiento") /100
		MontoTotal = Monto + Cargo
	end if
	
	if t = "D" then
		Comision = Monto * session("Cliente_Comision_Deposito") /100
		MontoTotal = Monto + Comision
	end if
	
	
	if cdbl(MontoTotal)<cdbl(session("Cliente_Saldo")) then
		CalcularSaldo = True
		session("Cliente_Saldo") = cdbl(session("Cliente_Saldo")) - cdbl(MontoTotal)
	else 
		CalcularSaldo = False
	end if
	
	
end function

function FormatearNumero(n)
	FormatearNumero = FormatNumber(n,2)
end function

Function TraerFotos(carpeta,idservicio,dimension)
	Dim Fotos
	Fotos = ""
	if idservicio="" then 
		nombre_carpeta = Server.MapPath(".") & "\upload\"& carpeta &"\"&dimension&"\"
	else
		nombre_carpeta = Server.MapPath(".") & "\upload\"& carpeta &"\"&idservicio&"\"&dimension&"\"		
	end if
	set FSO = server.createObject("Scripting.FileSystemObject")
	
	'creo el objeto carpeta
	if FSO.FolderExists(nombre_carpeta) then
		Set carpeta = FSO.GetFolder(nombre_carpeta)
		
		'traigo los archivos de la carpeta
		Set archivos = carpeta.Files
		
		for each file in archivos
			Fotos = file.name &","& Fotos
		next
		Fotos = left(Fotos,len(fotos)-1)
		TraerFotos = Fotos
	else
		TraerFotos = Fotos
	end if	
end function

Function BorrarFotos(carpeta,idservicio,archivoborrar)
	Dim Fotos
	Fotos = ""
	if idservicio="" then 
		nombre_carpeta = Server.MapPath(".") & "\upload\"& carpeta &"\"
	else
		nombre_carpeta = Server.MapPath(".") & "\upload\"& carpeta &"\"&idservicio&"\"		
	end if
	
	set FSO = server.createObject("Scripting.FileSystemObject")
	set FSO_sub = server.createObject("Scripting.FileSystemObject")
	
	'creo el objeto carpeta
	if FSO.FolderExists(nombre_carpeta) then
		'creo el objeto carpeta
		
		Set carpeta = FSO.GetFolder(nombre_carpeta)
		
		Set archivosRaiz = carpeta.Files
		for each file in archivosRaiz			
			if lcase(file.name)=lcase(archivoborrar) then
				fso.DeleteFile file,true
			end if			
		next
	
		'extraigo la coleccion de carpetas
		set coleccionCarpetas = carpeta.subFolders
		
		'para cada carpeta coloco un bot�n para ir a ella.
		for each carpeta in coleccionCarpetas
			'response.write "<br>" &carpeta.path
			
			Carpeta_Sub = carpeta.path&"\"
			'creo el objeto carpeta
			Set carpeta = FSO_sub.GetFolder(Carpeta_Sub)
			
			'traigo los archivos de la carpeta
			Set archivos = carpeta.Files
		
			'para cada archivo, muestro su nombre.
			for each file in archivos
				if lcase(file.name)=lcase(archivoborrar) then
					FSO_sub.DeleteFile file,true
				end if				
			next			
		next 
		BorrarFotos = "Ok"
	else
		BorrarFotos = "No existe el archivo"
	end if	
end function

function VencimientoMuestra(FechaFin,v,vp)
	
	ColorMuestra = "bg-green"
    Barra = "success"
		
	if datediff("h",now(),FechaFin)>v+vp then 
        ColorMuestra = "bg-green"
        Barra = "success"
    elseif datediff("h",now(),FechaFin)=>v and datediff("h",now(),FechaFin)<vp+v then 
        ColorMuestra = "bg-green"
        Barra = "success"
		'ColorMuestra = "bg-yellow"
        'Barra = "warning"
    elseif datediff("h",now(),FechaFin)<v then 
        ColorMuestra = "bg-yellow"
        Barra = "warning"
		'ColorMuestra = "bg-red"
		'Barra = "danger"
    end if
    
    VencimientoMuestra = ColorMuestra & "|" & Barra

end function

function VencimientoMuestra_Pago(FechaFin,v,vp)

		'if datediff("h",now(),FechaFin)=>12 then 
		'	ColorMuestra = "bg-green"
		'	Barra = "success"
		'end if
		'if datediff("h",now(),FechaFin)=>4 and datediff("h",now(),FechaFin)<12 then 
		'	ColorMuestra = "bg-yellow"
		'	Barra = "warning"
		'end if
		'if datediff("h",now(),FechaFin)<4 then 
		'	ColorMuestra = "bg-red"
		'	Barra = "danger"
		'end if
		ColorMuestra = "bg-green"
		Barra = "success"
		
		if datediff("h",now(),FechaFin)>v+vp then 
			ColorMuestra = "bg-green"
			Barra = "success"
		elseif datediff("h",now(),FechaFin)=>v and datediff("h",now(),FechaFin)<v+vp then 
			ColorMuestra = "bg-green"
			Barra = "success"
			'ColorMuestra = "bg-yellow"
			'Barra = "warning"
		elseif datediff("h",now(),FechaFin)<v+1 then 
			ColorMuestra = "bg-yellow"
			Barra = "warning"
			'ColorMuestra = "bg-red"
			'Barra = "danger"
		end if
		
		VencimientoMuestra_Pago = ColorMuestra & "|" & Barra

end function

function FormatearHora(f)

	MinutosTotales = f

	H = MinutosTotales / 60
	Min = MinutosTotales - (Fix(H) * 60)
	
	if FormatNumber(H,0)>0 then Hora = Fix(H)&"hs "
	Minutos = Min & "min"
	
	FormatearHora = Hora & Minutos
end function



Function CambiarMoneda(Precio,monedaID,tipo)
	
	if isNumeric(Precio) then
		if ucase(monedaID)<>2 then
			vM = session("v_Cambio_Monedas")
			vV = session("v_Cambio_Venta")
			vC = session("v_Cambio_Compra")
	
			PosC = -1
			for iM=0 to ubound(vM)
				if cint(vM(iM)) = cint(monedaID) then  PosC = iM
			next							
						
			if tipo = "C" then	
				CambiarMoneda = FormatNumber(cdbl(Precio) / cdbl( vC(PosC) ),4)	'Comprador
			else 
				CambiarMoneda = FormatNumber(cdbl(Precio) / cdbl( vV(PosC) ),4)	'Vendedor
			end if

		else
			CambiarMoneda = FormatNumber(Precio,4)
		end if
	else
		CambiarMoneda =  FormatNumber(Precio,4)
	end if	
	
	CambiarMoneda = replace(CambiarMoneda,",","")
	
end Function

Function TraerCambioDia(monedaID,tipo)

	vM = session("v_Cambio_Monedas")
	vV = session("v_Cambio_Venta")
	vC = session("v_Cambio_Compra")
	
	if cint(monedaID)<>2 then
		PosC = -1
		for iM=0 to ubound(vM)
			if cint(vM(iM)) = cint(monedaID) then  PosC = iM
		next							
					
		if tipo = "C" then	
			TraerCambioDia = FormatNumber(cdbl( vC(PosC) ),4)	'Comprador
		else 
			TraerCambioDia = FormatNumber(cdbl( vV(PosC) ),4)	'Vendedor
		end if
	else
		TraerCambioDia = 1
	end if
	
	TraerCambioDia = replace(TraerCambioDia,",","")
	
end function


Function RPT_STR(t,s)	
	if IsNull(s) then s = ""
	
	if t="s" and s<>"" then s = replace(s,",","")
	if t="d" and s<>"" then s = ArmarFecha(s)
	if t="n" and s<>"" then s = replace( s , ".", ",")
	if t="p" and s<>"" then s = replace( s , ";", ",")
	 
	RPT_STR = s
	
end function

Cliente_Saldo = FormatearNumero(session("Cliente_Saldo"))
%>
<!-- #include file="diccionario.asp" -->