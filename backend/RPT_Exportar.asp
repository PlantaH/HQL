<%
Set conn = Server.CreateObject("ADODB.Connection")
conn.ConnectionString = Application("String")
conn.open
conn.CommandTimeout = 960

rsT = conn.execute("select 'Marca;Codigo;Producto;Lote;Vencimiento;Cantidad' as titulos ")

Dim fso, f

nFile = "productos_"

fFile = year(date) & month(date) & day(date()) & hour(time) & minute(time) & second(time)

Set fso = CreateObject("Scripting.FileSystemObject")
Set f = fso.OpenTextFile("C:\ASP\Conectta\rpt\" & "rpt_" & nfile & "_" & fFile & ".csv", 2, True)
  
url_file = "http://148.113.197.232/Conectta/rpt/rpt_"  & nfile & "_" & fFile & ".csv"

f.Write rsT("titulos")
f.Write chr(13) & chr(10) 

sSql = "select *,"
sSql = sSql & "(select count(*) as cantidad from Productos_Log 	where codigo=p.codigo and Modo='E' ) as entradas,"
sSql = sSql & "(select count(*) as cantidad from Productos_Log 	where codigo=p.codigo and Modo='S' ) as salidas "
sSql = sSql & "from  Productos p "
sSql = sSql & "where  p.borrado=0 and p.marca='" & session("Marca_DashBoard") & "'   "
sSql = sSql & "order by p.nombre"
set rs = conn.execute(sSql)
Cant = 0
while not rs.eof	
	entradas = rs("entradas")
	salidas = rs("salidas")
	cantidad = entradas-salidas		
	
	
	sSql = "select  lote,FechaVencimiento,sum(cantidad) as total   from("
	sSql = sSql & "select codigo,lote,FechaVencimiento,count(codigo) as cantidad  "
	sSql = sSql & "from Productos_Log where Codigo = '" & rs("codigo") & "' and modo='E' "
	sSql = sSql & "group by codigo,lote,FechaVencimiento, modo "
	sSql = sSql & "union all "
	sSql = sSql & "select codigo,lote,FechaVencimiento,-1*count(codigo)as cantidad " 
	sSql = sSql & "from Productos_Log where Codigo = '" & rs("codigo") & "' and modo='S' "
	sSql = sSql & "group by codigo,lote,FechaVencimiento, modo ) as t "
	sSql = sSql & "group by codigo,lote,FechaVencimiento "
	set rt = conn.execute(sSql)
	
	while not rt.eof	
		if rt("total")>0 then	
			f.Write rs("marca") & ";"
			f.Write rs("codigo") & ";"
			f.Write rs("nombre") & ";"
			
			f.Write rt("lote") & ";"
			f.Write rt("FechaVencimiento") & ";"
			f.Write rt("total") & ";"
			
				
			f.Write chr(13) & chr(10) 
			
			Cant = Cant + 1 
		end if
		rt.movenext
	wend
	set rt = nothing
	
	rs.movenext
wend
set rs = nothing
f.Close
MsjDatos = "Se exportaron " & Cant & " productos"
 
conn.close
set conn = nothing 
%>
<!-- Right side column. Contains the navbar and content of the page -->


<h4><i class="icon fa fa-check"></i> Archivo generado con exito - <%=MsjDatos%></h4>
<%if Cant>0 then%>
<a href="<%=url_file%>">Por favor haga click aqui para guardarlo.</a>
<%end if%>
<br><br>
<a onclick="window.close()" style="cursor:pointer">Cerrar ventana</a>
