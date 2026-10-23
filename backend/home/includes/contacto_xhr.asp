<!-- #include file="../../includes/conn.asp" -->
<%
msj = "error"
if request.ServerVariables("REQUEST_METHOD")="POST" and request.form("sendType")="contacto"  then

	If request("email") <> "" Then
		
		'''''''''''''''''''''''''''''''''''''''
		bodyHTML = bodyHTML & chr(10) & "<div>"
		bodyHTML = bodyHTML & chr(10) & "<style type=""text/css"">"
		bodyHTML = bodyHTML & chr(10) & ".rss_mail_tabla{background:#fff;border:1px solid #e6e8ea;width:600px;}"
		bodyHTML = bodyHTML & chr(10) & ".rss_mail_titulo{font:bold 20px Arial, Helvetica, sans-serif;color:#4c5355;text-transform:uppercase;}"
		bodyHTML = bodyHTML & chr(10) & ".rss_mail_desc{font:normal 11px Arial, Helvetica, sans-serif;color:#828282;line-height:15px;padding-bottom:10px;list-style-type:square;}"
		bodyHTML = bodyHTML & chr(10) & ".rss_mail_tipo{font:bold 14px Arial, Helvetica, sans-serif;color:#4c5355;padding:0;margin:10px 0 5px 0;}"
		bodyHTML = bodyHTML & chr(10) & ".rss_mail_link{font:bold 11px Arial, Helvetica, sans-serif;color:#FF0C4B;}"
		bodyHTML = bodyHTML & chr(10) & ".rss_mail_fecha{font:normal 10px Arial, Helvetica, sans-serif;color:#4c5355;text-align:right;}"
		bodyHTML = bodyHTML & chr(10) & "a:hover.rss_mail_link{color:#FF0C4B}"
		bodyHTML = bodyHTML & chr(10) & ".rss_mail_data{border-top:1px solid #e6e8ea;}"
		bodyHTML = bodyHTML & chr(10) & ".rss_mail_ul{padding:5px 20px;margin:0;}"
		bodyHTML = bodyHTML & chr(10) & "</style>"
		bodyHTML = bodyHTML & chr(10) & "<table cellpadding=""6"" cellspacing=""10"" align=""center"" class=""rss_mail_tabla"">"
		bodyHTML = bodyHTML & chr(10) & "<tr>"
		bodyHTML = bodyHTML & chr(10) & "<td class=""rss_mail_titulo"">Contacto Web</td>"
		bodyHTML = bodyHTML & chr(10) & "<td class=""rss_mail_fecha"">" & FechaHoy() & "</td>"
		bodyHTML = bodyHTML & chr(10) & "</tr>"
		bodyHTML = bodyHTML & chr(10) & "<tr>"
		bodyHTML = bodyHTML & chr(10) & "<td class=""rss_mail_data"" colspan=""2"">"
		bodyHTML = bodyHTML & chr(10) & "<p class=""rss_mail_tipo"">Nuevo contacto desde la web</p>"
		bodyHTML = bodyHTML & chr(10) & "<ul class=""rss_mail_ul"">"
		bodyHTML = bodyHTML & chr(10) & "<li class=""rss_mail_desc"">Nombre:&nbsp;<strong>" & FixToXHTML(request("name")) & "</strong></li>"
		bodyHTML = bodyHTML & chr(10) & "<li class=""rss_mail_desc"">Tel&eacute;fono:&nbsp;<strong>" & FixToXHTML(request("telefono")) & "</strong></li>"
		bodyHTML = bodyHTML & chr(10) & "<li class=""rss_mail_desc"">Email:&nbsp;<strong><a href=""mailto:" & FixToXHTML(request("email")) & """ class=""rss_mail_link"">" & FixToXHTML(request("email")) & "</a></strong></li>"
		bodyHTML = bodyHTML & chr(10) & "<li class=""rss_mail_desc"">Comentarios:&nbsp;<strong>" & FixToXHTML(request("message")) & "</strong></li>"
		bodyHTML = bodyHTML & chr(10) & "</ul>"
		bodyHTML = bodyHTML & chr(10) & "</td>"
		bodyHTML = bodyHTML & chr(10) & "</tr>"
		bodyHTML = bodyHTML & chr(10) & "</table>"
		bodyHTML = bodyHTML & chr(10) & "</div>"
		'''''''''''''''''''''''''''''''''''''''
			
		Dim sEmail, envioemail
		set sEmail = server.createobject("RSS_Email.RSS_Email")		
		sEmail.DataSMTP Application("RSS_ServerAddres"),Application("Sender"), Application("SenderName") , Application("RSS_Usuario"), Application("RSS_Password"), Application("RSS_Puerto"), Application("RSS_SSL")
		sEmail.Asunto =  "Nuevo contacto web"
		sEmail.Cuerpo  = bodyHTML
		sEmail.AddDestinatario cstr(Application("AddRecipientContact"))
       '' sEmail.AddDestinatario cstr("raul@rss.com.ar")
		
		envioemail=sEmail.Enviar()		
		Set sEmail = Nothing	
		
		msj = "ok"

	end if

end if

response.write msj
%>
