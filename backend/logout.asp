<%
Set conn = Server.CreateObject("ADODB.Connection")
conn.ConnectionString = Application("String")
conn.open
conn.execute("update usuarios set userID = '' where usuarioID=" &session("ClienteID")	)
conn.close
set conn = nothing

session.Contents.RemoveAll()

session("ClienteID") = session.SessionID*-1	
session("HotelID") = session.SessionID*-1	
session.timeout = 60*2

session("SessionID") =  session.SessionID
session("TripPlannerID") = 0
	
Response.Redirect "login/"
%>