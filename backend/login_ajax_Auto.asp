<!-- #include file="includes/conn.asp" -->
<%
rw =  "error" 

set rs = conn.execute("select * from usuarios where userID = '" & noErrorSQL(request("id")) & "'")


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
		session("Modo_DashBoard") = "1"
		
		session("UsuarioID_Encrypt") = rs("UsuarioID")
		session("Cliente_Tipo") = rs("tipo")
		
		
		rw =  "ok"	
		
		
	
end if

Response.Write rw
%>