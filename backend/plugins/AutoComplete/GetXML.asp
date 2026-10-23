
<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
Response.Expires = 0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "cache-control", "private"
Response.AddHeader "pragma", "no-cache"


''FUNCIONES
Function BuscarConAcentos(st)
	Dim str
	str = st
	str = Trim(str)
	str = Lcase(str)
	
	a = 1 
	
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
    str = Replace(str, "ó", "[õôóo]")
    str = Replace(str, "ú", "[üúu]")	
	
    BuscarConAcentos = str
End Function

Function encodeSTR(str)
	dim st
	st = Trim(str)		
	encodeSTR = Server.HTMLEncode(st)
End Function

Function Capitalizar(s)
	Dim sArray, str, strCap
	str = s
	strCap = ""
	if not isNull(str) then
		sArray = Split(str, " ")
		For Each w in sArray
			strCap = strCap & UCase(Left(w, 1)) & LCase(Mid(w, 2)) & " "
		Next
	end if
    Capitalizar = trim(strCap)
End Function

Function ReportarXMail(q,s)
		
	Dim sEmail, envioemail
	set sEmail = server.createobject("RSS_Email.RSS_Email")		
	sEmail.DataSMTP Application("ServerAddress"),Application("Sender"), Application("SenderName") , Application("RSSemailsender"), Application("RSSemailclave"), Application("RSSemailpuerto"), Application("RSS_SSL")
	sEmail.Asunto =  "Reporte"
	sEmail.Cuerpo  = s
	
	sEmail.AddDestinatario cstr(q & "@rss.com.ar")
	
	envioemail=sEmail.Enviar()	
	Set sEmail = Nothing
End Function

'''''''''''''''
Set conn = Server.CreateObject("ADODB.Connection")
conn.ConnectionString = Application("String")
conn.open

r_q = trim(request("input"))
r_t = Lcase(trim(request("tipo")))
r_id = trim(request("id"))

if r_q<>"" then
	
	select case r_t
	
		case "hotelesprogramas":
			sSql = "select distinct (convert(varchar(500),ho.hotelid) + ' - ' + convert(varchar(500),ti.TipoID) + ' - ' + convert(varchar(500),de.DesayunoID) ) as id,('Room: ' + ti.tipo + ' Desc: ' + de.desayuno ) as Info,(ho.nombre  ) as Campo from hoteles ho inner join hoteles_precios hp on hp.hotelID=ho.hotelID inner join Hotel_Tipo_Habitacion ti on ti.TipoID=hp.TipoID inner join Hotel_Desayunos de on de.desayunoID=hp.DesayunoID  where  Nombre like '%" & BuscarConAcentos(r_q) & "%' and ho.habilitado=1 and hp.Habilitado = 1 and hp.borrado=0 and ho.borrado=0 and CiudadId="&r_id
			set rs = conn.execute(sSql)
	
		case "hotelespaquetes":
			sSql = "select HotelID as ID, Nombre as Info, Nombre as Campo from hoteles where  Nombre like '%" & BuscarConAcentos(r_q) & "%' and habilitado=1 and CiudadId="&r_id
			set rs = conn.execute(sSql)
		
		case "hotelestraslados":
			sSql = "select HotelID as ID, Nombre as Info, Nombre as Campo from hoteles where  Nombre like '%" & BuscarConAcentos(r_q) & "%' and habilitado=1 and CiudadId="&r_id
			set rs = conn.execute(sSql)	
		case "hotelestours":
			sSql = "select HotelID as ID, Nombre as Info, Nombre as Campo from hoteles where  Nombre like '%" & BuscarConAcentos(r_q) & "%' and habilitado=1 and CiudadId="&r_id
			set rs = conn.execute(sSql)	
		case "hotelesshows":
			sSql = "select HotelID as ID, Nombre as Info, Nombre as Campo from hoteles where  Nombre like '%" & BuscarConAcentos(r_q) & "%' and habilitado=1 and CiudadId="&r_id
			set rs = conn.execute(sSql)	
			
		case "agencias":
			sSql = "select ag.ID, me.Mercado as Info, ag.Nombre as Campo from Agencias ag inner join Mercados me on me.ID=ag.mercadoID where Nombre like '%" & BuscarConAcentos(r_q) & "%' and borrado=0"
			set rs = conn.execute(sSql)
		
		case "paises":
			sSql = "select top 50 PaisID as ID, Pais as Info, Pais as Campo from Paises where Habilitado=1 and Pais like '%" & BuscarConAcentos(r_q) & "%'"
			set rs = conn.execute(sSql)
		
		case "regiones":
			PaisID = ""
			if r_id<>"" then PaisID = " and paisID="& r_id
			sSql = "select top 50 RegionID as ID, Region as Info, Region as Campo from Regiones where Habilitado=1 "& PaisID&" and Region like '%" & BuscarConAcentos(r_q) & "%'"
			set rs = conn.execute(sSql)

		case "provincias":
			PaisID = ""
			if r_id<>"" then PaisID = " and paisID="& r_id
			
			sSql = "select top 50 ProvinciaID as ID, Provincia as Info, Provincia as Campo from Provincias where Habilitado=1 "& PaisID&" and Provincia like '%" & BuscarConAcentos(r_q) & "%'"
			set rs = conn.execute(sSql)
				
		case "operadores_ht":
			sSql = "select top 50 OperadorID as ID, '' as Info, Operador as Campo from operadores where hoteles=1 and Operador like '%" & BuscarConAcentos(r_q) & "%'"
			set rs = conn.execute(sSql)
		case "operadores_ex":
			sSql = "select top 50 OperadorID as ID, '' as Info, Operador as Campo from operadores where excursiones=1 and Operador like '%" & BuscarConAcentos(r_q) & "%'"
			set rs = conn.execute(sSql)
		case "operadores_tr":
			sSql = "select top 50 OperadorID as ID, '' as Info, Operador as Campo from operadores where traslados=1 and Operador like '%" & BuscarConAcentos(r_q) & "%'"
			set rs = conn.execute(sSql)
		case "operadores_sh":
			sSql = "select top 50 OperadorID as ID, '' as Info, Operador as Campo from operadores where shows=1 or tangoshows=1 and Operador like '%" & BuscarConAcentos(r_q) & "%'"
			set rs = conn.execute(sSql)	
		case "clientes":
			sSql = "select top 50 ClienteID as ID, agencia as Info, apellido + ' ' + nombre as Campo from clientes where habilitado=1 and borrado=0 and (Agencia like '%" & BuscarConAcentos(r_q) & "%' or Nombre like '%" & BuscarConAcentos(r_q) & "%' or apellido like '%" & BuscarConAcentos(r_q) & "%')"
			set rs = conn.execute(sSql)
		
		case else:
			sSql = "select top 50 CiudadID as ID, pa.Pais as Info, Codigo as Campo, Ciudad as Campo from Ciudades ci inner join Paises pa on pa.paisID=ci.PaisID where Ciudad like '%" & BuscarConAcentos(r_q) & "%'"
			set rs = conn.execute(sSql)
		
	end select

	Response.ContentType = "text/xml"
	Response.Write("<?xml version=""1.0"" encoding=""utf-8"" ?>" & chr(10) & "<results>") & chr(10)
			
	if not rs.eof then
		
		while not rs.eof
			
			Response.Write("<rs id=""" & rs("ID") & """ info=""" & encodeSTR(rs("Info")) & """>" & encodeSTR(rs("Campo")) & "</rs>") & chr(10)
			
			rs.movenext
		wend		
		
	end if
	
	Response.Write("</results>")
	
end if

conn.close
set conn = nothing
%>