<!-- #include file="includes/conn.asp" -->

<%
ID = request("ID")
pN = request("pN") 

backURL = Application("Dominio")  & "descargastock/"


Set Cmd = Server.CreateObject("ADODB.Command")
Cmd.Activeconnection = conn
Cmd.CommandText = "PRO_Descarga_Stock_Baja"
Cmd.CommandType = 4
					
Cmd.parameters.refresh				
Cmd(1) = session("usuarioID")
Cmd(2) = request("codigo")
Cmd(3) = request("lote")
Cmd(4) = request("FechaVencimiento")
											
Cmd.execute,,adexecutenorecords	
 
		
response.redirect backURL
%>