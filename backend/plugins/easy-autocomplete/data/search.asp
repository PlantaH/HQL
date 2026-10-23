
<%
Response.CodePage = 65001
Response.CharSet = "utf-8"
response.ContentType = "application/json; charset=utf-8"

Function GetJsonFromRs(RS,Name)
    Dim JSON
    if Name<>"" then
		JSON = "{" & """" & Name & """:["  
	else
		JSON = "["  
	end if
	
    Do While NOT RS.EOF
        JSON = JSON & "{"
        For x = 0 To RS.Fields.Count - 1
            IF x < RS.Fields.Count - 1 THEN
                JSON = JSON & """" & RS(x).Name & """:""" & toxhtml(RS(x)) & ""","
            ELSE
                JSON = JSON & """" & RS(x).Name & """:""" & toxhtml(RS(x)) & """"
            END IF          
        Next      
    RS.MoveNext
        IF NOT RS.EOF THEN
            JSON = JSON & "},"
        ELSE
            JSON = JSON & "}"
        END IF
    Loop  
    JSON = JSON & "]"
    GetJsonFromRs = JSON
End Function



Function toXHTML(str)
	dim st, i, tempV, vX
	st = str
	if len(st)>0 then
		if instr(st,"<")>0 then
			st = replace(replace(st,chr(10),""),chr(13),"")
		else		
			st = replace(st,chr(10),"<br />")
		end if
		
		st = Server.HTMLEncode(st)
		vX = Array("&lt;br&gt;I<br />","&lt;BR&gt;I<br />","&lt;br /&gt;I<br />")	
		for each i in vX
			tempV = split(i,"I")
			st = Replace(st, tempV(0), tempV(1))
		next
	end if
	
	toXHTML = st
End Function

if request("phrase")<>"" then
	Set conn = Server.CreateObject("ADODB.Connection")
	conn.ConnectionString = Application("String")
	conn.open

	if request("m") = "1" then 
		set ht=conn.execute("select top 20 ID, codigo + ' - ' + marca   as Info, nombre as campo  from productos where borrado=0 and (nombre like '%" & request("phrase") & "%' or  codigo like '%" & request("phrase") & "%' or  marca like '%" & request("phrase") & "%') order by nombre")
	 
	end if
	
	 
	a = GetJsonFromRs(ht,"")

	conn.close
	set conn = nothing 
else
	a = ""
end if
response.write a
%>
