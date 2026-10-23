<%
response.ContentType = "charset=iso-8859-1"
usuarioID = 0


if request.ServerVariables("REQUEST_METHOD")="POST" and request.form("boton")="ACEPTAR" then
	
	session("MensajeError") = ""  
	session("MensajeColor") = ""
	sol_paquete = trim(request("sol_Paquete"))
	if sol_paquete<>"" then
	 
		 
		set rs = conn.execute("select * from  Productos where codigo='" & sol_paquete & "' and borrado=0")
		if not rs.eof then	 
				 			
			Lote = ""
			if request("sol_Lote")<>"" then session("sol_Lote") = request("sol_Lote")
			FechaVencimiento = ""
			if request("sol_Vencimiento")<>"" then session("sol_Vencimiento") = request("sol_Vencimiento")
		
			Set Cmd = Server.CreateObject("ADODB.Command")
			Cmd.Activeconnection = conn
			Cmd.CommandText = "PRO_Lectura_Remito"
			Cmd.CommandType = 4
								
			Cmd.parameters.refresh
					
			Cmd(1) = sol_paquete
			Cmd(2) = session("usuarioID")
			Cmd(3) = session("nro_remito")
			Cmd(4) = session("sol_Lote")
			Cmd(5) = session("sol_Vencimiento")
			Cmd(6) = "E"
			 
									
			Cmd.execute,,adexecutenorecords	
			
			session("TituloError") = ""
			session("MensajeError") = "Producto " & rs("nombre")   & " cargado"
			session("MensajeColor") = "success"		
			 
		
		else		
			 session("TituloError") = "ERROR"
			 session("MensajeError") = "No se encontro el producto: " & request("sol_Paquete")
			 
			 session("MensajeColor") = "danger"
		end if	
			  
			
		 
		 
	end if
	
	response.redirect  Application("Dominio") & "cargastock/"
	 
end if
%>
 
   
<div class="content-wrapper">
    <!-- Content Header (Page header) -->
    <section class="content-header">
      <h1>Cargar Productos del Remito Nro: <strong><%=session("nro_remito")%></strong></h1>
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
								<label>Lote</label>
								<input type="text" class="form-control" name="sol_Lote" id="sol_Lote" value="<%=session("sol_Lote")%>" tabindex="2" maxlength="50" />
							</div>
							<div class="col-xs-12 col-sm-2">
								<label>Vencimiento</label>
								<input type="text" class="form-control" name="sol_Vencimiento" id="sol_Vencimiento" value="<%=session("sol_Vencimiento")%>" tabindex="3" maxlength="12"  />
							</div>							
							<div class="col-xs-12 col-sm-2">
								<label>Codigo Caja</label>
								<input type="text" class="form-control" name="sol_Paquete" id="sol_Paquete" tabindex="1" required />
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
			<h5 style="color: black; font-size:22px; font-weight:bold;"><i class="fa fa-exclamation-triangle"></i> <%=session("TituloError")%></h5>
			<h5 style="color: black; font-size:18px; font-weight:bold;"><%=session("MensajeError")%></h5> 
		</div>
		<%end if%>
		
		
		 
		<div class="row">
			<div class="col-md-6">
			
				<div class="box-body body-form">
					<table width="100%" class="table" border="1">
						<tr style="background-color:#CCE6C3">
							<td>CodigoVenta</td>
							<td>Codigo</td>
							<td>Producto</td>
							<td>Lote</td>
							<td>Vencimiento</td>
							<td>Cantidad</td>	
							<td></td>							
						</tr>
						<%mostrarTErminar = false
						set rAs = conn.execute("select p.CodigoVenta,pl.codigo,pl.lote,pl.fechaVencimiento,p.nombre, count(pl.codigo) as cantidad from Productos_Log_Temporal pl inner join productos p on p.codigo=pl.codigo where pl.usuarioID=" & session("usuarioID") & " and p.borrado=0 group by p.CodigoVenta,pl.codigo,pl.lote,pl.fechaVencimiento,p.nombre")
						while not rAs.eof
							mostrarTErminar = true%>
							<tr>
								<td><%=rAs("CodigoVenta")%></td>
								<td><%=rAs("Codigo")%></td>
								<td><%=rAs("nombre")%></td>
								<td><%=rAs("lote")%></td>
								<td><%=rAs("FechaVencimiento")%></td>
								<td><%=rAs("cantidad")%></td>	

								<td><a href="PRO_Carga_Eliminar.asp?codigo=<%=rAs("codigo")%>&lote=<%=rAs("lote")%>&FechaVencimiento=<%=rAs("FechaVencimiento")%>">Eliminar uno</a></td>									
							</tr>
							<%rAs.movenext
						wend
						set rAs = nothing%>
					</table>	
				</div>
			</div>
		</div>
		
		<%if mostrarTErminar then%>
		<div class="alert alert-info alert-dismissible"> 
			<h5 style="color: black; font-size:22px; font-weight:bold;"><i class="fa fa-exclamation-triangle"></i> Atencion!</h5>
			<h5 style="color: black; font-size:18px; font-weight:bold;"><a onclick="ValidarCarga()" style="cursor:pointer">Para confimar la carga del remito cliquee aqui</a></h5> 
		</div>		
		<%end if%>
	</section>
	

</div>

    
  