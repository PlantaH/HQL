<%
dim socket
if Request.ServerVariables("HTTPS") = "on" then 
	socket = "https"
	UsarHttps = true
else
	socket = "http"
	UsarHttps = False
end if

if socket = "http" then response.redirect "https://www.hql.com.ar/"

server.scripttimeout = 960

Function noErrorSQL(str)
	str = Replace(Trim(str), "ñ", "n")
	str = Replace(Trim(str), "Ñ", "N")
    noErrorSQL = Replace(Trim(str), "'", "''")
End Function

Function ArmarFecha(Fecha)
	ArmarFecha = right(Fecha,2) & "/" & mid(Fecha,5,2) & "/" & left(Fecha,4)
End Function

Response.ContentType = "text/html"
Response.AddHeader "Content-Type", "text/html;charset=UTF-8"
Response.CodePage = 65001
Response.CharSet = "UTF-8"

Set conn = Server.CreateObject("ADODB.Connection")
conn.ConnectionString = Application("String")
conn.open

%>