<%

Set conn = Server.CreateObject("ADODB.Connection")
conn.ConnectionString = Application("String")
conn.open 

 
set rs = conn.execute("select ID from Servicios where  Manifiesto='" & request("Manifiesto") & "'")
 
if  rs.eof then
	response.write "OK"
else
	response.write "NOK"
end if

conn.close
set conn = nothing

%>