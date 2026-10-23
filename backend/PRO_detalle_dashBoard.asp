<%
Function ArmarFecha(Fecha)
	ArmarFecha = right(Fecha,2) & "/" & mid(Fecha,5,2) & "/" & left(Fecha,4)
End Function
response.ContentType = "charset=iso-8859-1"

Set conn = Server.CreateObject("ADODB.Connection")
conn.ConnectionString = Application("String")
conn.open

if request("ID")<>"" then
 
	producto = replace(request("ID"), "-","")
	
	 
	entradas = 0
	set rs = conn.execute("select count(*) as cantidad from Productos_Log 	where codigo='" & producto & "' and Modo='E'  ")
	if not rs.eof then entradas = rs("cantidad")
	salidas = 0
	set rs = conn.execute("select   count(*) as cantidad from Productos_Log 	where codigo='" & producto & "' and Modo='S'  ")
	if not rs.eof then salidas = rs("cantidad")
	

 %>

<!DOCTYPE html>
<html>
  <head>
    <style>
    body {
        height: 842px;
         
        /* to centre page on screen*/
        margin-left: 0;
        margin-right: 0;
    }
    
	.td 	{
		border-left:1px solid black;
		border-top:1px solid black;
	}
	.table{
		border-right:1px solid black;
		border-bottom:1px solid black;
	}
	</style>
  </head>
  <body>
	 
	<table   width="100%" cellspacing="0" cellpadding="0">
        <tr>
			<td width="100%">
				<table   width="100%" cellspacing="2" cellpadding="2" class="table">
					<tr>
						<td class="td" align="center">PRODUCTO</td>
						<td class="td" align="center">CANTIDAD</td> 
					</tr>
					
				 
					
					<%if not rs.eof then%>            
						<tr>
							<td class="td"  align="center"><strong><%=producto%></strong></td>
							<td class="td"  align="center"><strong><%=entradas-salidas%></strong></td>
						</tr>
						<%
						sSql = "select  lote,FechaVencimiento,sum(cantidad) as total   from("
						sSql = sSql & "select codigo,lote,FechaVencimiento,count(codigo) as cantidad  "
						sSql = sSql & "from Productos_Log where Codigo = '" & producto & "' and modo='E' "
						sSql = sSql & "group by codigo,lote,FechaVencimiento, modo "
						sSql = sSql & "union all "
						sSql = sSql & "select codigo,lote,FechaVencimiento,-1*count(codigo)as cantidad " 
						sSql = sSql & "from Productos_Log where Codigo = '" & producto & "' and modo='S' "
						sSql = sSql & "group by codigo,lote,FechaVencimiento, modo ) as t "
						sSql = sSql & "group by codigo,lote,FechaVencimiento "
						set rs = conn.execute(sSql)
						while not rs.eof
							if rs("total")>0 then%>
							<tr>
								<td class="td"  align="center">Lote <%=rs("lote")%> - <%=rs("FechaVencimiento")%> </td>
								<td class="td"  align="center"><%=rs("total")%></td>
							</tr>
							<%end if
							rs.movenext
						wend
					end if
					set rs = nothing%>
					 
					
				</table>
			</td>
			 
			 
		</tr>				
    </table>
	 
  </body>
</html>

<%end if


conn.close
set conn = nothing%>

 