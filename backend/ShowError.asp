<%
set objasperror = server.getlasterror()

''''''''''''
str = "<div style=""padding:10px; font:normal 12px 'Courier New', Courier, monospace;color:#993300; background:#F9F8E3;width:500px;top:20px;left:20px;border:5px solid #E2DABA;"">"
str = str & "<p style=""color:#930""><strong>Detalle de Error</strong></p>"
str = str & "N&uacute;mero: " & objasperror.number & " (0x" & hex(objasperror.number) & ")<br>"
str = str & "Categor&iacute;a: " & objasperror.category & "<br>"
str = str & "Archivo = " & objasperror.file & "<br>"
str = str & "Linea: <strong>" & objasperror.line & "</strong><br>"
str = str & "Columna: <strong>" & objasperror.column & "</strong><br><br>"
str = str & "Descripci&oacute;n: <strong>" & objasperror.description & "</strong><br>"
str = str & "Pagina: <strong>" & Request.ServerVariables("URL") & "?" & Request.ServerVariables("QUERY_STRING") & "</strong><br><br>"
str = str & "Method: <strong>" & Request.ServerVariables("REQUEST_METHOD") & "</strong><br>"
if Request.ServerVariables("REQUEST_METHOD")="POST" then
	str = str & "Request:<ul>"
	for each x in Request.Form
		str = str & chr(10) & "<li>" & x & " = " & Request.Form(x) & "</li>"
	next
	str = str & "</ul>"
end if

str = str & "<hr>"

str = str & "IP: <strong>" & Request.ServerVariables("REMOTE_ADDR") & "</strong><br><br>"
if session("log") then
	str = str & "Logueado: <strong>SI</strong>"
end if
str = str & "</div>"

session("SHOWERROR_num") = objasperror.number

response.write(str)

%>