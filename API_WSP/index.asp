<%
' 1. Configurar respuesta HTTP básica para UltraMsg
Response.ContentType = "application/json"
Response.CharSet = "UTF-8"

' 2. Leer el cuerpo completo del JSON que envía UltraMsg
Dim TotalBytes, BinaryData, JsonString
TotalBytes = Request.TotalBytes

If TotalBytes > 0 Then
    BinaryData = Request.BinaryRead(TotalBytes)
    
    ' Convertir los datos binarios recibidos a un String entendible por ASP
    JsonString = BinaryToString(BinaryData)
    
    ' 3. Conectar a SQL Server y guardar el JSON en bruto
    Dim Conn, ConnectionString, SQL
    ConnectionString = "Provider=SQLOLEDB;Data Source=WIN-IFK26VFLCQ9; Initial Catalog=Conectta; User ID=Conectta; Password=Conectta1905"
    
    Set Conn = Server.CreateObject("ADODB.Connection")
    Conn.Open ConnectionString
	
	conn.execute("INSERT INTO HistorialWebhooksRaw (FechaRecibido, JsonData) VALUES (GETDATE(), '" & JsonString &"')")
    
    Conn.Close
    Set Conn = Nothing
    
    ' 4. Responder con un HTTP 200 OK para que UltraMsg sepa que llegó con éxito
    Response.Status = "200 OK"
    Response.Write "{""status"": ""success""}"
Else
    ' Si no llegó información, responder un error genérico
    Response.Status = "400 Bad Request"
    Response.Write "{""status"": ""error"", ""message"": ""No data received""}"
End If

' Función auxiliar obligatoria en ASP Classic para convertir la petición binaria a Texto (String)
Function BinaryToString(Binary)
    Dim Stream
    Set Stream = Server.CreateObject("ADODB.Stream")
    Stream.Type = 1 ' adTypeBinary
    Stream.Open
    Stream.Write Binary
    Stream.Position = 0
    Stream.Type = 2 ' adTypeText
    Stream.Charset = "utf-8"
    BinaryToString = Stream.ReadText
    Set Stream = Nothing
End Function
%>