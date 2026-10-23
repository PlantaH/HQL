<!-- #include file="includes/conn.asp" -->

<%
ID = request("ID")
pN = request("pN") 

backURL = Application("Dominio")  & "usuarios/"


Set Cmd = Server.CreateObject("ADODB.Command")
Cmd.Activeconnection = conn
Cmd.CommandText = "USR_Usuarios_Baja"
Cmd.CommandType = 4
						
Cmd.parameters.refresh				
Cmd(1) = ID
												
Cmd.execute,,adexecutenorecords	
		
response.redirect backURL
%>