<%
response.ContentType = "application/json; charset=iso-8859-1"
str_Data = ""
 
Set conn = Server.CreateObject("ADODB.Connection")
conn.ConnectionString = Application("String")
conn.open 
 
''''''''''''''''''''''''''''''''''''''
session("draw") = session("draw") + 1 
draw     = session("draw")
start    = request("start")
cantReg  = request("length")
ordenCol = request("order[0][column]")
ordenDir = request("order[0][dir]")
filtro   = request("search[value]")

if ordenCol = "0" or ordenCol = "" then
	sql_order = " ORDER BY p.marca " ' & UCase(ordenDir)
elseif  ordenCol = "1" then
	sql_order = " ORDER BY p.marca " & UCase(ordenDir)
elseif  ordenCol = "2" then
	sql_order = " ORDER BY p.codigoVenta " & UCase(ordenDir)
elseif  ordenCol = "3" then
	sql_order = " ORDER BY p.codigo " & UCase(ordenDir)
elseif  ordenCol = "4" then
	sql_order = " ORDER BY p.nombre " & UCase(ordenDir)	
elseif  ordenCol = "5" then
	sql_order = " ORDER BY p.cantidad " & UCase(ordenDir)		
end if

''''''CANTIDAD DE REGISTROS NECESARIOS PARA LA PAGINACION
sSql = "SELECT COUNT(*) as cant FROM  productos p "
sSqlFiltro =  " where p.borrado = 0  "
if filtro <> "" then
	sSqlFiltro = sSqlFiltro & " and ( (p.nombre like '%" & filtro & "%') or (p.marca like '%" & filtro & "%') or (p.codigoVenta like '%" & filtro & "%') or (p.codigo like '%" & filtro & "%')) "
end if
set rC = conn.execute(sSql & sSqlFiltro )
recordsTotal = rC("cant")
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 
pageSize = cantReg

if start = 0 then	
	pageNum = 1
else
	pageNum = (start/cantReg) + 1
end if
 
sSql = "SELECT p.* FROM productos p "
sSql = sSql & sSqlFiltro
sSql = sSql & sql_order & " "
sSql = sSql & "OFFSET " & (pageNum-1)*pageSize & " ROWS FETCH NEXT " & pageSize & " ROWS ONLY "  
'''''''''''''''''''''''''''''''''''''''
 
Set RS = Server.CreateObject("ADODB.Recordset")
RS.Open sSql, Conn,3,1



Do While Not RS.EOF

	cant_Reg = cant_Reg + 1
 
	
	str_Data = str_Data & "[" & rs("ID") & ",'" & rs("marca") & "','" & rs("codigoVenta") & "','" & rs("codigo") & "','" &  rs("nombre") & "','" & rs("cantidad") & "','Borrar']"
		
	RS.MoveNext
	if Not RS.EOF then str_Data = str_Data & ", "			
Loop	

set rs = nothing

str_response = "{'draw':" & draw & ",'recordsTotal':" & recordsTotal & ",'recordsFiltered':" & recordsTotal  & ",'data':[" & str_Data &"]}"


conn.close
set conn = nothing

response.write replace(str_response,"'","""")


%>