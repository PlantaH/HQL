<%
response.ContentType = "charset=iso-8859-1"
usuarioID = 0

CartelProductos = false

if request.ServerVariables("REQUEST_METHOD")="POST" and request.form("boton")="ACEPTAR" then
	
	
	if request("sol_Paquete_Revisar")<>"" then
		set rs = conn.execute("select * from  Productos where codigo='" & request("sol_Paquete_Revisar") & "' and borrado=0")
		if not rs.eof then	
			CartelProductos = true
			
			sSql = "select codigo,lote,FechaVencimiento, sum(cantidad) as cantidad from( "
			sSql = sSql & "select codigo,lote,FechaVencimiento, count(*) as cantidad from Productos_Log where codigo='" & request("sol_Paquete_Revisar") & "' and Modo='E' "
			sSql = sSql & "group by codigo,lote,FechaVencimiento "
			sSql = sSql & "union all "
			sSql = sSql & "select codigo,lote,FechaVencimiento,-1* count(*) as cantidad from Productos_Log where codigo='" & request("sol_Paquete_Revisar") & "' and Modo='S' "
			sSql = sSql & "group by codigo,lote,FechaVencimiento "
			
			sSql = sSql & "union all "
			sSql = sSql & "select codigo,lote,FechaVencimiento,-1* count(*) as cantidad from Productos_Log_Temporal_descarga "
			sSql = sSql & "where codigo='" & request("sol_Paquete_Revisar") & "' and Modo='S' group by codigo,lote,FechaVencimiento "
			sSql = sSql & ") as t   group by codigo,lote,FechaVencimiento"
		 
			set rL = conn.execute(sSql)
			
			cuantosPr = 0
			sol_Lote = ""
			sol_FechaVencimiento = ""
			while not rL.eof 
				if rL("cantidad")>0 then 
					cuantosPr = cuantosPr + 1
					sol_Lote = rL("lote")
					sol_FechaVencimiento = rL("FechaVencimiento")
				end if
				rL.movenext
			wend
			if cuantosPr>1 then rl.movefirst
			
			if cuantosPr = 1 then
				Set Cmd = Server.CreateObject("ADODB.Command")
				Cmd.Activeconnection = conn
				Cmd.CommandText = "PRO_Lectura_Descarga"
				Cmd.CommandType = 4
									
				Cmd.parameters.refresh
						
				Cmd(1) = request("sol_Paquete_Revisar")
				Cmd(2) = session("usuarioID")
				Cmd(3) = sol_Lote
				Cmd(4) = sol_FechaVencimiento
				Cmd(5) = "S"
				 
										
				Cmd.execute,,adexecutenorecords
				
				session("MensajeError") = "Producto " & rs("nombre")   & " descargado"
			
				session("MensajeColor") = "warning"
				
				response.redirect Application("Dominio") & "descargastock/"	
			
			end if
			 
		else		
			 session("MensajeError") = "No se encontro el producto: " & request("sol_Paquete_Revisar")
			 
			 session("MensajeColor") = "danger"
		end if	
	end if
	
	if request("sol_Paquete")<>"" then
	 
		 
		set rs = conn.execute("select * from  Productos where codigo='" & request("sol_Paquete") & "' and borrado=0")
		if not rs.eof then	 
				 
			
			
			if rs("cantidad")=0 and session("PRO_Modo") = "S" then
				session("MensajeError") = "El producto " & rs("nombre") & "tiene cantidad 0 no se puede sacar "
			 
				session("MensajeColor") = "danger"
			else
			
				Set Cmd = Server.CreateObject("ADODB.Command")
				Cmd.Activeconnection = conn
				Cmd.CommandText = "PRO_Lectura_Descarga"
				Cmd.CommandType = 4
									
				Cmd.parameters.refresh
						
				Cmd(1) = request("sol_Paquete")
				Cmd(2) = session("usuarioID")
				Cmd(3) = request("sol_Lote")
				Cmd(4) = request("sol_FechaVencimiento")
				Cmd(5) = "S"
				 
										
				Cmd.execute,,adexecutenorecords
				
				session("MensajeError") = "Producto " & rs("nombre")   & " descargado"
			
				session("MensajeColor") = "warning"
			
			end if
		
		else		
			 session("MensajeError") = "No se encontro el producto: " & request("sol_Paquete")
			 
			 session("MensajeColor") = "danger"
		end if	
			  
			
		response.redirect Application("Dominio") & "descargastock/"	 
	 
	end if
	
	
	 
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
				  <form role="form"  method="post" id="formABM" onsubmit="document.getElementById('sol_Paquete_Revisar').readOnly = true; "> 
					<input type="hidden" name="boton" value="ACEPTAR">  
					<div class="form-group" >
						
						<div class="row" >									
							
							<div class="col-xs-12 col-sm-1">
								<label>Codigo Caja</label> 
							</div>  
							<div class="col-xs-12 col-sm-4">
								<input type="text" class="form-control" name="sol_Paquete_Revisar" id="sol_Paquete_Revisar" tabindex="1" required />
							</div>
							<div class="col-xs-12 col-sm-4">								 
								 <button type="submit" class="btn btn-primary">Aceptar</button>
							</div>
							<div class="col-xs-12 col-sm-3"></div>
						</div>
						
						 
					</div>
				   </form>
				</div>
								
			</div>
		</div>
		
		<%if session("MensajeError")<>"" then%>
		<!--<div class="alert alert-<%=session("MensajeColor")%> alert-dismissible"> 
			<h5 style="color: black; font-size:22px; font-weight:bold;"><i class="fa fa-exclamation-triangle"></i> Atencion!</h5>
			<h5 style="color: black; font-size:18px; font-weight:bold;"><%=session("MensajeError")%></h5> 
		</div>-->
		<%end if%>
		
		<%if CartelProductos then%>
		<div class="alert alert-info alert-dismissible"> 
			<h5 style="color: black; font-size:22px; font-weight:bold;"><i class="fa fa-exclamation-triangle"></i> Atencion!</h5>
			<h5 style="color: black; font-size:18px; font-weight:bold;">Seleccione el Lote y fecha de vencimiento del producto <%=request("sol_Paquete_Revisar")%></h5> 
		</div>	
		
		<div class="row">
			<div class="col-md-6">
			
				<div class="box-body body-form">
					<table width="100%" class="table" border="1">
						<tr style="background-color:#279FF5">
							<td>Producto</td>
							<td>Lote</td>
							<td>Fecha Vencimiento</td>
							<td>Stock</td>
							<td></td>
						</tr>
						<%
						sinStock = true
						while not rL.eof
							if rL("cantidad")>0 then
								sinStock = false%>
								<tr>
									<td><%=rL("Codigo")%></td>
									<td><%=rL("lote")%></td>
									<td><%=rL("FechaVencimiento")%></td>
									<td><%=rL("cantidad")%></td>
									<td>
										<form name="" method="post">
											<input type="hidden" name="sol_Paquete" value="<%=request("sol_Paquete_Revisar")%>">
											<input type="hidden" name="sol_lote" value="<%=rL("lote")%>">
											<input type="hidden" name="sol_FechaVencimiento" value="<%=rL("FechaVencimiento")%>">
											<input type="hidden" name="boton" value="ACEPTAR">  
											
											<input type="submit" name="ACEPTAR" value="Seleccionar este producto" />
										</form>
									</td>
								</tr>
							<%end if
							rL.movenext
						wend
						set Rl = nothing%>
					</table>	
				</div>
			</div>
		</div>
		<%if sinStock then%>
		<div class="alert alert-danger alert-dismissible"> 
			<h5 style="color: black; font-size:22px; font-weight:bold;"><i class="fa fa-exclamation-triangle"></i> Atencion!</h5>
			<h5 style="color: black; font-size:18px; font-weight:bold;">No se encontro stock</h5> 
		</div>
		<%end if%>		
		
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
						set rAs = conn.execute("select p.CodigoVenta,pl.codigo,pl.lote,pl.fechaVencimiento,p.nombre, count(pl.codigo) as cantidad from Productos_Log_Temporal_descarga pl inner join productos p on p.codigo=pl.codigo where pl.usuarioID=" & session("usuarioID") & " and p.borrado=0 group by p.CodigoVenta,pl.codigo,pl.lote,pl.fechaVencimiento,p.nombre")
						while not rAs.eof
							mostrarTErminar = true%>
							<tr>
								<td><%=rAs("CodigoVenta")%></td>
								<td><%=rAs("Codigo")%></td>
								<td><%=rAs("nombre")%></td>
								<td><%=rAs("lote")%></td>
								<td><%=rAs("FechaVencimiento")%></td>
								<td><%=rAs("cantidad")%></td>	

								<td><a href="PRO_descarga_Eliminar.asp?codigo=<%=rAs("codigo")%>&lote=<%=rAs("lote")%>&FechaVencimiento=<%=rAs("FechaVencimiento")%>">Eliminar uno</a></td>									
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
			<h5 style="color: black; font-size:18px; font-weight:bold;"><a onclick="ValidarDescarga()" style="cursor:pointer">Para confimar la descarga cliquee aqui</a></h5> 
		</div>		
		<%end if%>
	</section>
	

</div>

    
  <%
  session("MensajeError") = ""  
	session("MensajeColor") = ""
  %>