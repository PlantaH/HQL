<!-- #include file="includes/conn.asp" -->

<%
ID = request("ID")
pN = request("pN") 

backURL = Application("Dominio")  & "productos/"

vID = split(ID,"-")

for i=0 to ubound(vID)
	Set Cmd = Server.CreateObject("ADODB.Command")
	Cmd.Activeconnection = conn
	Cmd.CommandText = "PRO_Productos_Baja"
	Cmd.CommandType = 4
							
	Cmd.parameters.refresh				
	Cmd(1) = vID(i)
													
	Cmd.execute,,adexecutenorecords	
next
		
response.redirect backURL
%>