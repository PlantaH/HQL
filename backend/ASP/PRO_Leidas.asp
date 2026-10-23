<%
response.ContentType = "charset=iso-8859-1"
usuarioID = 0


if request.ServerVariables("REQUEST_METHOD")="POST" and request.form("boton")="ACEPTAR" then
	session("MensajeError") = ""  
	session("MensajeColor") = ""
	if request("sol_Paquete")<>"" then
	 
		 
		set rs = conn.execute("select * from  Productos where codigo='" & request("sol_Paquete") & "' and borrado=0")
		if not rs.eof then	 
				 
			
			
			if rs("cantidad")=0 and session("PRO_Modo") = "S" then
				session("MensajeError") = "El producto " & rs("nombre") & "tiene cantidad 0 no se puede sacar "
			 
				session("MensajeColor") = "danger"
			else
			
				Set Cmd = Server.CreateObject("ADODB.Command")
				Cmd.Activeconnection = conn
				Cmd.CommandText = "PRO_Lectura"
				Cmd.CommandType = 4
									
				Cmd.parameters.refresh
						
				Cmd(1) = request("sol_Paquete")
				Cmd(2) = session("usuarioID")
				Cmd(3) = session("PRO_Modo")
				 
										
				Cmd.execute,,adexecutenorecords	
				
				session("MensajeError") = "Producto " & rs("nombre")   & " cargado"
			
				session("MensajeColor") = "warning"
			
			end if
		
		else		
			 session("MensajeError") = "No se encontro el producto: " & request("sol_Paquete")
			 
			 session("MensajeColor") = "danger"
		end if	
			  
			
		 
		 
	end if
	
	response.redirect linkRedir
	 
end if
%>
 
   
<div class="content-wrapper">
    <!-- Content Header (Page header) -->
    <section class="content-header">
      <h1><%=titPagina%></h1>
    </section>
	
	<section class="content">
	    
		<div class="row">
			<div class="col-md-12">
			
				<div class="box-body body-form" style="background-color:<%=colorCaja%>">
				  <form role="form"  method="post" id="formABM" onsubmit="document.getElementById('sol_Paquete').readOnly = true; "> 
					<input type="hidden" name="boton" value="ACEPTAR">  
					<div class="form-group" >
						
						<div class="row" >									
							
							  
							
							<div class="col-xs-12 col-sm-2">
								<label>Codigo Caja</label>
								<input type="text" class="form-control" name="sol_Paquete" id="sol_Paquete" tabindex="1" required />
							</div>
							<div class="col-xs-12 col-sm-6">
								<br>
								<strong ><%=msj%><strong>
							</div>
							 
							<div class="box-footer" style="background-color:<%=colorCaja%>">
								 <button type="submit" class="btn btn-primary">Guardar</button>
								 
							</div>
						</div>
					</div>
				   </form>
				</div>
								
			</div>
		</div>
		
		<%if session("MensajeError")<>"" then%>
		<div class="alert alert-<%=session("MensajeColor")%> alert-dismissible"> 
			<h5 style="color: black; font-size:22px; font-weight:bold;"><i class="fa fa-exclamation-triangle"></i> Atencion!</h5>
			<h5 style="color: black; font-size:18px; font-weight:bold;"><%=session("MensajeError")%></h5> 
		</div>
		<%end if%>
		
		
		 
		<div class="row">
			<div class="col-md-6">
			
				<div class="box-body body-form">
					<table width="100%" class="table" border="1">
						<tr style="background-color:#CCE6C3">
							<td>Producto</td>
							<td>Usuario</td>
							<td>Modo</td>
							<td>Hora</td>
						</tr>
						<%
						set rAs = conn.execute("select  us.nombre + ' ' + us.apellido as usuario, pl.*, convert(varchar,pl.Fecha,108)  as hora from Productos_Log pl inner join usuarios us on us.usuarioID=pl.usuarioID  where   convert(varchar,pl.Fecha,112) = convert(varchar,getdate(),112)order by pl.fecha desc ")
						while not rAs.eof%>
							<tr>
								<td><%=rAs("Codigo")%></td>
								<td><%=rAs("usuario")%></td>
								<td><%=rAs("modo")%></td>
								<td><%=rAs("hora")%></td>
							</tr>
							<%rAs.movenext
						wend
						set rAs = nothing%>
					</table>	
				</div>
			</div>
		</div>
		 
	</section>
	

</div>

    
  