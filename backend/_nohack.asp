<%
''CRISTIAN 15/12/2010 <- Mejorado
''Evitar Hackeos v2.0
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Response.ContentType = "text/html"
Response.Charset = "utf-8"

Response.Expires = 0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "cache-control", "private"
Response.AddHeader "pragma", "no-cache"

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''Detectar si ya han hackeado antes
If Not Session("detectarIPhack") Then
	Session("detectarIPhack") = True
	
	Set ConnH = Server.CreateObject("ADODB.Connection")
	ConnH.ConnectionString = Application("String")
	ConnH.Open
	
	Set RsH = ConnH.Execute("SELECT ip FROM LogHack WHERE ip='" & Request.ServerVariables("REMOTE_ADDR") & "'")
	
	If Not RsH.Eof Then Session("BanearHack") = True
	
	ConnH.Close
	Set ConnH = Nothing
End If

''Si lo hizo, no ejecuta el resto de la pagina
If Session("BanearHack") Then Response.End()
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''Verifica cada Request
Dim r_Rhack
If Request.ServerVariables("REQUEST_METHOD")="POST" Then
	For Each r_Rhack In Request.Form
		If isHacking(Request.Form(r_Rhack)) Then
			Call RegistrarHack(Request.Form(r_Rhack), r_Rhack)
			Exit For
		End If
	Next	
Else
	For Each r_Rhack In Request.QueryString	
		If isHacking(Request.QueryString(r_Rhack)) Then
			Call RegistrarHack(Request.QueryString(r_Rhack), r_Rhack)
			Exit For
		End If
	Next
End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''FUNCIONES : para Detectarlo, Registrarlo o Eliminarlo

Function isHacking(s)
	
	Dim VectorPosiblesHacks
	
	if IsEmpty(Session("VectorPosiblesHacks")) then
		Dim fsoHack, fsoHacksFile, fsoHacksLista
		
		Set fsoHack = CreateObject("Scripting.FileSystemObject")
		Set fsoHacksFile = fsoHack.GetFile("c:\ASP\_NOHACK.ini")
		Set fsoHacksLista = fsoHacksFile.OpenAsTextStream(1)
		
		Session("VectorPosiblesHacks") = Replace(Replace(fsoHacksLista.ReadAll,chr(13),""),chr(10),"")
		
		fsoHacksLista.Close
		Set fsoHacksLista = nothing
		Set fsoHacksFile = nothing
		Set fsoHack = nothing
	end if
	
	VectorPosiblesHacks = Split(Session("VectorPosiblesHacks"), "##")
	
	Dim v_H, s_H
	s_H = False
	If Not IsEmpty(s) and Ubound(VectorPosiblesHacks)>0 Then
		For v_H=0 to Ubound(VectorPosiblesHacks)
			If Instr(Ucase(Unescape(s)),Ucase(Unescape(VectorPosiblesHacks(v_H))))>0 Then
				s_H = True
				Exit For
			End If		
		Next
	End If
	isHacking = s_H
	
End Function

Sub RegistrarHack(cod,req)
		
	Set ConnH = Server.CreateObject("ADODB.Connection")
	ConnH.ConnectionString = Application("String")
	ConnH.Open
	
	Set Cmd = Server.CreateObject("ADODB.Command")
	Cmd.Activeconnection = ConnH
	Cmd.CommandText = "LogHack_Alta"
	Cmd.CommandType = 4
						
	Cmd.parameters.refresh
			
	Cmd(1) = Request.ServerVariables("REMOTE_ADDR") ' IP
	Cmd(2) = Replace(req, "'","''") ' Request
	Cmd(3) = Replace(cod, "'","''") ' C&oacute;digo
												
	Cmd.execute,,adexecutenorecords	
	
	ConnH.Close
	Set ConnH = Nothing
	
	 
	Response.End()
	
End Sub

Sub LimpiarIPHack(ip)

	Set ConnH = Server.CreateObject("ADODB.Connection")
	ConnH.ConnectionString = Application("String")
	ConnH.Open
	
	ConnH.Execute("EXEC LogHack_Baja '" & ip & "'")
	
	ConnH.Close
	Set ConnH = Nothing

End Sub

'LimpiarIPHack("192.168.1.94")
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'--TABLA Y PROCEDIMIENTOS
'
'CREATE TABLE [LogHack] (
'	[id] [int] IDENTITY (1, 1) NOT NULL ,
'	[ip] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
'	[request] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
'	[hack] [text] COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
'	[fecha] [datetime] NOT NULL CONSTRAINT [DF_LogHack_fecha] DEFAULT (getdate()),
'	CONSTRAINT [PK_LogHack] PRIMARY KEY  CLUSTERED 
'	(
'		[id]
'	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
') ON [PRIMARY]
'GO
'
'SET QUOTED_IDENTIFIER ON 
'GO
'SET ANSI_NULLS ON 
'GO
'
'CREATE PROCEDURE [LogHack_Alta]
'	(@ip_1 	[varchar](50),
'	 @request_2 	[varchar](50),
'	 @hack_3 	[text])
'
'AS INSERT INTO [LogHack] 
'	 ( [ip],
'    [request],
'	 [hack]) 
' 
'VALUES 
'	(@ip_1,
'	 @request_2,
'	 @hack_3)
'GO
'SET QUOTED_IDENTIFIER OFF 
'GO
'SET ANSI_NULLS ON 
'GO
'
'SET QUOTED_IDENTIFIER OFF 
'GO
'SET ANSI_NULLS ON 
'GO
'
'CREATE PROCEDURE [LogHack_Baja]
'	(@ip 	varchar(50))
'
'AS DELETE [LogHack] 
'
'WHERE 
'	( [ip]	 = @IP)
'GO
'SET QUOTED_IDENTIFIER OFF 
'GO
'SET ANSI_NULLS ON 
'GO
%>