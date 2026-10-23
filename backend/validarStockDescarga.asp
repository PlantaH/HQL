<!-- #include file="includes/conn.asp" -->

<%
 

backURL = Application("Dominio")  & "salidastock/i/"

 
 

Set Cmd = Server.CreateObject("ADODB.Command")
Cmd.Activeconnection = conn
Cmd.CommandText = "PRO_Lectura_Descarga_Valida"
Cmd.CommandType = 4
						
Cmd.parameters.refresh				
Cmd(1) = session("usuarioID")
												
Cmd.execute,,adexecutenorecords	

session("MSJ_Fin_carga_pedido") = "Descargado con exito!"
 
	
 
		
response.redirect backURL
%>