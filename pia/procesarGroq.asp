<%@ Language="VBScript" CodePage=65001 %>
<%
Response.CharSet = "UTF-8"
Response.Buffer = True

' =========================================================================
' CONFIGURACIÓN DE CREDENCIALES
' =========================================================================
' =========================================================================

Dim textoCliente
textoCliente = Trim(Request.Form("texto"))

If textoCliente = "" Then
    Response.Write "<div style='color:red;'>El texto ingresado está vacío.</div>"
    Response.End
End If

' 1. Consultar a la IA
Dim respuestaIA, listaProductos, i
respuestaIA = ConsultarGroq(textoCliente, API_KEY_GROQ)

If respuestaIA = "" Or InStr(respuestaIA, "ERROR_API:") > 0 Then
    Response.Write "<div style='color:red;'>Error al procesar con IA: " & Server.HTMLEncode(respuestaIA) & "</div>"
    Response.End
End If

' Limpieza básica de respuesta
respuestaIA = Replace(respuestaIA, "```", "")
respuestaIA = Replace(respuestaIA, vbCr, "")
respuestaIA = Replace(respuestaIA, vbLf, "")
respuestaIA = Trim(respuestaIA)

listaProductos = Split(respuestaIA, ",")

' 2. Conectar a la BD
Dim conn, rs, sql, prodBuscado, totalEncontrados
totalEncontrados = 0

'On Error Resume Next
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open CADENA_CONEXION_SQL

If Err.Number <> 0 Then
    Response.Write "<div style='color:red;'>Error al conectar con la BD: " & Server.HTMLEncode(Err.Description) & "</div>"
    Err.Clear
    Response.End
End If
On Error GoTo 0

Response.Write "<h2>Resultado de la búsqueda</h2>"
Response.Write "<p><i>Términos detectados por la IA: <b>" & Server.HTMLEncode(respuestaIA) & "</b></i></p>"

Response.Write "<table border='1' cellpadding='8' cellspacing='0' style='border-collapse: collapse; width: 100%; font-family: sans-serif;'>"
Response.Write "  <tr style='background-color: #f2f2f2; text-align: left;'>"
 
Response.Write "    <th>Producto en Inventario</th>"
Response.Write "    <th>Precio</th>"
Response.Write "    <th>Stock</th>"
Response.Write "  </tr>"

For i = 0 To UBound(listaProductos)
    prodBuscado = Trim(listaProductos(i))
    
    If prodBuscado <> "" Then
        ' Construir consulta SQL inteligente por tokens de palabras
        sql = ConstruirSQLBusqueda(prodBuscado,respuestaIA)
        
        Set rs = conn.Execute(sql)
        
        If Not rs.EOF Then
            Do While Not rs.EOF
                totalEncontrados = totalEncontrados + 1
                Response.Write "<tr>" 
                Response.Write "  <td>" & Server.HTMLEncode(rs("nombre")) & "</td>"
                Response.Write "  <td>$" & FormatNumber(rs("precio"), 2) & "</td>"
                
                If rs("stock") > 0 Then
                    Response.Write "  <td style='color: green;'>Disponible (" & rs("stock") & ")</td>"
                Else
                    Response.Write "  <td style='color: red;'>Sin Stock</td>"
                End If
                
                Response.Write "</tr>"
                rs.MoveNext
            Loop
        Else
            Response.Write "<tr>"
            Response.Write "  <td><b>" & Server.HTMLEncode(prodBuscado) & "</b></td>"
            Response.Write "  <td colspan='3' style='color: gray;'><i>No se encontraron coincidencias en el inventario.</i></td>"
            Response.Write "</tr>"
        End If
        
        rs.Close
        Set rs = Nothing
    End If
Next

Response.Write "</table>"

conn.Close
Set conn = Nothing


' =========================================================================
' FUNCIÓN PARA CONSTRUIR SQL DINÁMICO MULTI-PALABRA
' =========================================================================
Function ConstruirSQLBusqueda(termino,respuestaIA)
    Dim palabras, p, palabraLimpia, sqlWhere
    
    ' Reemplazar caracteres especiales y separar por espacios
    termino = Replace(termino, "'", "''")
    palabras = Split(termino, " ")
    
    sqlWhere = " WHERE activo = 1 "
    
    For p = 0 To UBound(palabras)
        palabraLimpia = Trim(palabras(p))
        ' Ignorar palabras cortas o conectores (de, el, la, y, con, etc.)
        If Len(palabraLimpia) > 2 Then
            sqlWhere = sqlWhere & " AND nombre LIKE '%" & palabraLimpia & "%' "
        End If
    Next
    
    ' Si eran solo palabras de 1 o 2 letras, usar la frase completa
    If sqlWhere = " WHERE activo = 1 " Then
        sqlWhere = " WHERE activo = 1 AND nombre LIKE '%" & termino & "%' "
    End If
    
	'ConstruirSQLBusqueda = "select p.* from ProductosExtraidos pe inner join productos p on p.id=pe.productoid where pe.pedidoid="  & respuestaIA & " order by nombre"
	ConstruirSQLBusqueda = "select 0 as id,'necesitas ' + producto as nombre,0 as precio,0 as stock,1 from ProductosNecesarios where  pedidoid="  & respuestaIA
	ConstruirSQLBusqueda = ConstruirSQLBusqueda & " union all "
	ConstruirSQLBusqueda = ConstruirSQLBusqueda & "select p.* from ProductosExtraidos pe inner join productos p on p.id=pe.productoid where pe.pedidoid="  & respuestaIA
	' "SELECT TOP 5 id, nombre, precio, stock FROM Productos " & sqlWhere & " ORDER BY nombre ASC"
End Function


' =========================================================================
' FUNCIÓN DE CONEXIÓN A GROQ (CON PROMPT MEJORADO)
' =========================================================================
Function ConsultarGroq(frase, apiKey)
   ' On Error Resume Next
    Dim http, url, jsonPayload, responseText
    
     url = "https://api.groq.com/openai/v1/chat/completions"
    
    Dim fraseEscapada
    fraseEscapada = Replace(frase, """", "\""")
    fraseEscapada = Replace(fraseEscapada, vbCr, " ")
    fraseEscapada = Replace(fraseEscapada, vbLf, " ")

    ' Prompt optimizado para búsqueda en bases de datos: palabras en singular y limpias
   ' Dim promptEspecial
    'promptEspecial = "Extrae los conceptos principales de los productos solicitados en esta frase: '" & fraseEscapada & "'. " & _
   '                  "REGLAS: " & _
    '                 "1. Convierte las palabras al singular (ejemplo: 'fideos' -> 'fideo', 'quesos' -> 'queso'). " & _
    '                 "2. Omita marcas o adjetivos innecesarios. " & _
    '                 "3. Devuelve UNICAMENTE los nombres de los productos separados por comas en texto plano sin comillas."
	
	Dim promptEspecial
    'promptEspecial = "Analiza la solicitud del usuario en ESPAÑOL: '" & fraseEscapada & "'. " & _
    '                 "Si pide una receta o plato (ej: 'torta de chocolate'), deduce los ingredientes de almacén necesarios para cocinarlo. " & _
    '                 "Si pide artículos directos, extráelos. " & _
    '                 "REGLAS OBLIGATORIAS: " & _
    '                 "1. Responde SIEMPRE Y EXCLUSIVAMENTE EN ESPAÑOL. " & _
    '                 "2. Devuelve UNICAMENTE los nombres de los productos o ingredientes en singular (ej: 'huevo', 'harina', 'leche'). " & _
    '                 "3. Sepáralos ÚNICAMENTE por comas en texto plano. " & _
    '                 "4. Sin explicaciones, sin traducciones al inglés, sin cantidades, sin viñetas y sin comillas."
	
	promptEspecial = "Analiza la solicitud del usuario en ESPAÑOL: '" & fraseEscapada & "'. " & _
                     "Si pide una receta o plato (ej: 'torta de chocolate'), deduce los ingredientes de almacén necesarios para cocinarlo. " & _
                     "Si pide artículos directos, extráelos. " & _
                     "REGLAS OBLIGATORIAS: " & _
                     "1. Responde SIEMPRE Y EXCLUSIVAMENTE EN ESPAÑOL. " & _
                     "2. Sepáralos ÚNICAMENTE por comas en texto plano. "  
					 
    jsonPayload = "{" & _
      """model"": ""openai/gpt-oss-20b""," & _
      """messages"": [{" & _
        """role"": ""user""," & _
        """content"": """ & promptEspecial & """" & _
      "}]" & _
    "}"
	
    jsonPayload = "{" & _
      """model"": ""openai/gpt-oss-20b""," & _
      """messages"": [{" & _
        """role"": ""user""," & _
        """content"": """ & promptEspecial & """" & _
      "}]" & _
    "}"

    Set http = Server.CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.open "POST", url, False
    http.setRequestHeader "Content-Type", "application/json"
    http.setRequestHeader "Authorization", "Bearer " & Trim(apiKey)
    http.send jsonPayload

    If Err.Number <> 0 Then
        ConsultarGroq = "ERROR_API: " & Err.Description
        Err.Clear
        Exit Function
    End If

    If http.Status = 200 Then
        responseText = http.responseText
        
		
        Dim html, window
        Set html = Server.CreateObject("htmlfile")
        Set window = html.parentWindow
        window.execScript "var json = " & responseText & ";", "JScript"
        
		Set conn = Server.CreateObject("ADODB.Connection")
		conn.Open CADENA_CONEXION_SQL
		
		Set Cmd = Server.CreateObject("ADODB.Command")
		Cmd.Activeconnection = conn
		Cmd.CommandText = "IA_Pedidos_Alta"
		Cmd.CommandType = 4
								
		Cmd.parameters.refresh				
		
		Cmd(1).Direction = 2
		Cmd(1).Type = 3	
		Cmd(2) = responseText
		 
	 
		Cmd.Execute , , adExecuteNoRecords
		ID = Cmd(1).Value
		
		conn.close
		set conn = nothing
		
        ConsultarGroq = ID ' window.json.choices(0).message.content
        
        Set http = Nothing
        Set html = Nothing
    Else
        ConsultarGroq = "ERROR_API: HTTP " & http.Status & " - " & http.responseText
    End If
End Function
%>