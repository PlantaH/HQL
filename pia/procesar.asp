 
<%@ Language="VBScript" CodePage=65001 %>
<%
Response.CharSet = "UTF-8"
Response.Buffer = True

' =========================================================================
' CONFIGURACIÓN: Ingresa tu API Key de Gemini
' =========================================================================

' =========================================================================

Dim textoCliente
textoCliente = Trim(Request.Form("texto"))

If textoCliente = "" Then
    Response.Write "<div style='color:red;'>El texto ingresado está vacío.</div>"
    Response.End
End If

' 1. Llamar a la API de Gemini
Dim respuestaIA, listaProductos, i
respuestaIA = ConsultarIA(textoCliente, API_KEY_GROQ)

' Limpiar comillas, saltos de línea y etiquetas Markdown si la IA los devolvió
respuestaIA = Replace(respuestaIA, VB32, "") ' Quitar comillas si las hay
respuestaIA = Replace(respuestaIA, "```", "")
respuestaIA = Replace(respuestaIA, vbCr, "")
respuestaIA = Replace(respuestaIA, vbLf, "")
respuestaIA = Trim(respuestaIA)

If respuestaIA = "" Or InStr(respuestaIA, "ERROR_API:") > 0 Then
    Response.Write "<div style='color:red;'>Error al comunicarse con la IA: " & Server.HTMLEncode(respuestaIA) & "</div>"
    Response.End
End If

' 2. Convertir la respuesta de comas en un Array de VBScript
listaProductos = Split(respuestaIA, ",")

If UBound(listaProductos) < 0 Then
    Response.Write "<div style='color:orange;'>{" & Chr(34) & "status" & Chr(34) & ":" & Chr(34) & "error" & Chr(34) & "," & Chr(34) & "message" & Chr(34) & ":" & Chr(34) & "No se pudieron interpretar productos en la frase." & Chr(34) & "}</div>"
    Response.End
End If

' 3. Renderizar resultado e iterar para buscar en SQL Server
Response.Write "<h3>Productos interpretados por la IA:</h3>"
Response.Write "<ul>"
For i = 0 To UBound(listaProductos)
    Dim prodLimpio
    prodLimpio = Trim(listaProductos(i))
    If prodLimpio <> "" Then
        Response.Write "<li><b>" & Server.HTMLEncode(prodLimpio) & "</b></li>"
        
        ' AQUÍ EJECUTAS TU CONSULTA A SQL SERVER
        ' Dim sql
        ' sql = "SELECT * FROM Productos WHERE Nombre LIKE '%" & Replace(prodLimpio, "'", "''") & "%'"
        ' Set rs = conn.Execute(sql)
        ' ...
		dim conn, rs  , primerProducto
		primerProducto = true
		 
		
		Set conn = Server.CreateObject("ADODB.Connection")
		On Error Resume Next
		conn.Open CADENA_CONEXION_SQL
		
		sql = "SELECT id, nombre, precio, stock FROM Productos WHERE nombre LIKE '%" &  Replace(prodLimpio, "'", "''") & "%' AND activo = 1"
		Set rs = conn.Execute(sql)
		
		Dim jsonOutput
		jsonOutput = "{""status"":""ok"",""productos"":["
		Do While Not rs.EOF
			If Not primerProducto Then jsonOutput = jsonOutput & ","
			
			jsonOutput = jsonOutput & "{"
			jsonOutput = jsonOutput & """id"":" & rs("id") & ","
			jsonOutput = jsonOutput & """nombre"":""" & EscapeJSON(rs("nombre")) & ""","
			jsonOutput = jsonOutput & """precio"":" & Replace(FormatNumber(rs("precio"), 2), ",", ".") & ","
			jsonOutput = jsonOutput & """stock"":" & rs("stock")
			jsonOutput = jsonOutput & "}"
			
			primerProducto = False
			rs.MoveNext
		Loop
		rs.Close

				
		conn.Close
		Set conn = Nothing
		
		jsonOutput = jsonOutput & "]}"
		Response.Write jsonOutput
    End If
Next
Response.Write "</ul>"

' =========================================================================
' FUNCIÓN PARA COMUNICARSE CON GEMINI VIA HTTP
' =========================================================================

 Function ConsultarGemini(frase, apiKey)
    On Error Resume Next
    Dim http, url, jsonPayload, responseText
    
    ' URL actualizada con el modelo gemini-3.8-flash
    url = "https://generativelanguage.googleapis.com/v1beta/models/gemini-3.8-flash:generateContent?key=" & apiKey
    
    ' Escapar comillas dobles en el texto de entrada
    Dim fraseEscapada
    fraseEscapada = Replace(frase, """", "\""")
    
    jsonPayload = "{" & _
      """contents"": [{" & _
        """parts"": [{" & _
          """text"": ""Extrae los productos solicitados de esta frase: '" & fraseEscapada & "'. Devuelve UNICAMENTE los nombres de los productos separados por comas, en texto plano, sin comillas ni listas.""}]" & _
        "}]" & _
      "}"

    Set http = Server.CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.open "POST", url, False
    http.setRequestHeader "Content-Type", "application/json"
    http.send jsonPayload

    If Err.Number <> 0 Then
        ConsultarGemini = "ERROR_API: " & Err.Description
        Err.Clear
        Exit Function
    End If

    If http.Status = 200 Then
        responseText = http.responseText
        
        Dim posStart, posEnd
        posStart = InStr(responseText, """text"": """)
        If posStart > 0 Then
            posStart = posStart + 9
            posEnd = InStr(posStart, responseText, """")
            If posEnd > posStart Then
                ConsultarGemini = Mid(responseText, posStart, posEnd - posStart)
            Else
                ConsultarGemini = ""
            End If
        Else
            ConsultarGemini = ""
        End If
    Else
        ConsultarGemini = "ERROR_API: HTTP " & http.Status & " - " & http.responseText
    End If

    Set http = Nothing
End Function

Function ConsultarIA(frase, apiKey)
    On Error Resume Next
    Dim http, url, jsonPayload, responseText
    
    url = "https://api.groq.com/openai/v1/chat/completions"
    
    Dim fraseEscapada
    fraseEscapada = Replace(frase, """", "\""")
    
    ' Usamos el modelo activo llama-3.3-70b-versatile
    jsonPayload = "{" & _
      """model"": ""llama-3.3-70b-versatile""," & _
      """messages"": [{" & _
        """role"": ""user""," & _
        """content"": ""Extrae los productos solicitados de esta frase: '" & fraseEscapada & "'. Devuelve UNICAMENTE los nombres de los productos separados por comas, en texto plano, sin comillas ni listas.""" & _
      "}]" & _
    "}"

    Set http = Server.CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.open "POST", url, False
    http.setRequestHeader "Content-Type", "application/json"
    http.setRequestHeader "Authorization", "Bearer " & apiKey
    http.send jsonPayload

    If Err.Number <> 0 Then
        ConsultarIA = "ERROR_API: " & Err.Description
        Err.Clear
        Exit Function
    End If

    If http.Status = 200 Then
        responseText = http.responseText
        
        ' Extraer la respuesta del JSON de respuesta de Groq
        Dim posStart, posEnd
        posStart = InStr(responseText, """content"": """)
        If posStart > 0 Then
            posStart = posStart + 12
            posEnd = InStr(posStart, responseText, """")
            If posEnd > posStart Then
                ConsultarIA = Mid(responseText, posStart, posEnd - posStart)
                ConsultarIA = Replace(ConsultarIA, "\n", "")
            End If
        End If
    Else
        ConsultarIA = "ERROR_API: HTTP " & http.Status & " - " & http.responseText
    End If

    Set http = Nothing
End Function
%>