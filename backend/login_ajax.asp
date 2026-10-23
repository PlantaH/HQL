<!-- #include file="includes/conn.asp" -->
<%
rw =  "error"
		
set rs = conn.execute("select *,NEWID() as userID from usuarios where email = '" & noErrorSQL(request("email")) & "' and password = '" &  noErrorSQL(request("password")) & "' and estado=1")


if not rs.eof then
	
		session("log")=true
		
		session("UsuarioID") = rs("UsuarioID")		
		session("Cliente_Master") = False
		if rs("master") then session("Cliente_Master") = True
		
		session("EsOperador") = true
		
		session("SessionID") =  session.SessionID
		
		session("Cliente_Nombre") = rs("Nombre") & " <br> " &rs("apellido")
		session("Cliente_Empresa") = "CONECTTA"'rs("Empresa")
		session("Cliente_Email") = rs("email")
		session("Cliente_RazonSocial") = "CONECTTA"'rs("Empresa")
		session("Cliente_Tipo") = rs("tipo")
		session("Modo_DashBoard") = "1"
		
		
		conn.execute("update usuarios set userID = '" & rs("userID") & "' where usuarioID=" & rs("UsuarioID")	)
		
		rw =  rs("userID")
	
end if

Response.Write rw
%>