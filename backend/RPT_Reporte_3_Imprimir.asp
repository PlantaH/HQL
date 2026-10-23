<%
Function ArmarFecha(Fecha)
	ArmarFecha = right(Fecha,2) & "/" & mid(Fecha,5,2) & "/" & left(Fecha,4)
End Function
response.ContentType = "charset=iso-8859-1"

Set conn = Server.CreateObject("ADODB.Connection")
conn.ConnectionString = Application("String")
conn.open

if request.ServerVariables("REQUEST_METHOD")="POST" and request.form("boton")="ACEPTAR" then

	fechadesde = replace(request("fechadesde"), "-","")
	fechahasta = replace(request("fechahasta"), "-","")
	usuarioID = replace(request("usuarioID"), "-","") 
	
	set rs = conn.execute("select * from dbo.RPT_Reporte_3(" & usuarioID & "," & fechadesde & "," & fechahasta & ") order by nombre,fechaCarga")
	
		

end if


if fechadesde<>"" and fechahasta<>"" then%>

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
	<h2>Reporte desde <%=request("fechadesde")%> hasta el <%=request("fechahasta")%></h2>
	<table   width="100%" cellspacing="0" cellpadding="0">
        <tr>
			<td width="100%">
				<table   width="100%" cellspacing="2" cellpadding="2" class="table">
					<tr>
						<td class="td" align="center">USUARIO</td> 
						<td class="td" align="center">CODIGOVENTA</td> 
						<td class="td" align="center">CODIGO</td>
						<td class="td" align="center">PRODUCTO</td> 
						<td class="td" align="center">REMITO</td> 
						<td class="td" align="center">LOTE</td> 
						<td class="td" align="center">VENCIMIENTO</td> 						
						<td class="td" align="center">MODO</td> 
						<td class="td" align="center">FECHA</td>
						<td class="td" align="center">CANTIDAD</td> 
					</tr>
					
				 
					
					<%
					TotalE=0
					TotalS=0
					cantidad = 0 
					while not rs.eof
						
						if rs("modoC") = "E"	then 
							cantidad = cantidad + rs("cantidad")
						else
							cantidad = cantidad - rs("cantidad")
						end if	
						
						if rs("modoC") ="E" then TotalE = TotalE + rs("cantidad")
						if rs("modoC") ="S" then TotalS = TotalS + rs("cantidad")						
						%>            
						<tr>
							<td class="td"  align="center"><%=rs("nombre")%></td>  	
							<td class="td"  align="center"><%=rs("codigoVenta")%></td>							
							<td class="td"  align="center"><%=rs("codigo")%></td>							
							
							<td class="td"  align="center"><%=rs("producto")%></td>
							<td class="td"  align="center"><%=rs("remito")%></td>
							
							<td class="td"  align="center"><%=rs("lote")%></td>
							<td class="td"  align="center"><%=rs("fechaVencimiento")%></td>
							
							<td class="td"  align="center"><%=rs("modo")%></td>
							<td class="td"  align="center"><%=ArmarFecha(rs("fechaCarga"))%></td>
							<td class="td"  align="center"><%=rs("cantidad")%></td>
						 	 	
						</tr>
						<%rs.movenext
					wend%>
					
					<tr>
						<td class="td" align="center"></td> 
						<td class="td" align="center"></td> 
						<td class="td" align="center"></td>
						<td class="td" align="center"></td> 
						<td class="td" align="center"></td> 
						<td class="td" align="center"></td> 
						<td class="td" align="center"></td> 						
						<td class="td" align="center"></td> 
						<td class="td" align="center"></td>
						<td class="td" align="center"><%=cantidad%></td> 
					</tr> 
					
				</table>
			</td>
			 
			 
		</tr>				
    </table>
	<br />
	<center><input type="button" id="imprimir" value="IMPRIMIR" onclick="imprimir()" /></center>
  </body>
</html>

<%end if

conn.close
set conn = nothing%>

<script>
function imprimir(){
	document.getElementById('imprimir').style.display = 'none';
	window.print()
	document.getElementById('imprimir').style.display = 'block';
}
</script>
 