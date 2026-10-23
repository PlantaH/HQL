<!-- #include file="conn.asp" -->
<%set rsTer = conn.execute("select * from terminoscondiciones")
if not rsTer.eof then

	response.write rsTer("Terminos")
	
end if%>