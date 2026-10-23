<!-- #include file="includes/conn.asp" -->

<%
 

backURL = Application("Dominio")  & "entradastock/i/"

 
if session("nro_remito")<>"" then

	Set Cmd = Server.CreateObject("ADODB.Command")
	Cmd.Activeconnection = conn
	Cmd.CommandText = "PRO_Lectura_Remito_Valida"
	Cmd.CommandType = 4
							
	Cmd.parameters.refresh				
	Cmd(1) = session("usuarioID")
													
	Cmd.execute,,adexecutenorecords	
	
	session("MSJ_Fin_carga_remito") = "Remito Nro " & session("nro_remito") & " cargado con exito!"
	session("nro_remito") = ""
	
end if
		
response.redirect backURL
%>