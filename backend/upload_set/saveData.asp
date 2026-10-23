<%
'Dim conn

'Set conn = Server.CreateObject("ADODB.Connection")
'conn.ConnectionString = Application("String")
'conn.open


'select case request("tabla")
'	case "pagos"
		'conn.execute("insert into Pagos_Comprobantes (PagoID,Imagen) values ("&request("ID")&",'"&request("imagen")&"')")
'end select
response.write "ok"
%>