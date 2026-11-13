<%@ Language=VBScript CodePage=65001 %>
<%
Response.CharSet = "utf-8"
Response.ContentType = "application/json"

' 1. Obtener la firma enviada por POST
Dim firmaBase64
firmaBase64 = Request.Form("firma")

If firmaBase64 = "" Then
    Response.Write "{""success"": false, ""message"": ""No se recibió ninguna firma.""}"
    Response.End
End If

' 2. Configurar la cadena de conexión a SQL Server
Dim connStr, conn, cmd
connStr ="Provider=SQLOLEDB; Data Source=WIN-IFK26VFLCQ9; Initial Catalog=CoordInt; User ID=CoordInt; Password=CoordInt$1905?"  

On Error Resume Next

Set conn = Server.CreateObject("ADODB.Connection")
conn.Open connStr

If Err.Number <> 0 Then
    Response.Write "{""success"": false, ""message"": ""Error de conexión a la base de datos.""}"
    Response.End
End If

' 3. Insertar usando ADO Command y parámetros (para soportar texto largo NVARCHAR(MAX))
Set cmd = Server.CreateObject("ADODB.Command")
Set cmd.ActiveConnection = conn
cmd.CommandType = 1 ' adCmdText
cmd.CommandText = "INSERT INTO Firmas (FirmaBase64, FechaCreacion) VALUES (?, GETDATE())"

' Parámetro adLongVarWChar (203) para soportar textos muy grandes tipo NVARCHAR(MAX)
cmd.Parameters.Append cmd.CreateParameter("Firma", 203, 1, Len(firmaBase64), firmaBase64)

cmd.Execute

If Err.Number <> 0 Then
    Dim errMsg
    errMsg = Replace(Err.Description, """", "\""")
    Response.Write "{""success"": false, ""message"": """ & errMsg & """}"
    conn.Close
    Set conn = Nothing
    Response.End
End If

' 4. Limpiar y responder éxito
conn.Close
Set conn = Nothing
Set cmd = Nothing

Response.Write "{""success"": true, ""message"": ""Firma guardada correctamente.""}"
%>