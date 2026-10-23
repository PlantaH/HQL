<%
response.ContentType = "charset=iso-8859-1"
usuarioID = 0

conn.execute("delete from Productos_Log_Temporal_Descarga where usuarioID=" & session("usuarioID"))

if request.ServerVariables("REQUEST_METHOD")="POST" and request.form("boton")="ACEPTAR" then
	 
	 
	 
		 
	
	session("TituloError") = ""
	session("MensajeError") = ""
	session("MensajeColor") = ""	
	
	
	
	response.redirect Application("Dominio") & "descargastock/"	
		 
	 
	
	 
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
				  <form role="form"  method="post" id="formABM" > 
					<input type="hidden" name="boton" value="ACEPTAR">  
					<div class="form-group" >
						
						<div class="row" >									
							
							    
							<div class="col-xs-12 col-sm-4">								 
								 <button type="submit" class="btn btn-primary">Empezar nuevo proceso de descarga</button>
							</div>
							<div class="col-xs-12 col-sm-4"></div>
						</div>
					</div>
				   </form>
				</div>
								
			</div>
		</div>
		
		 
		<%if session("MSJ_Fin_carga_pedido")<>"" then%>
			<div class="alert alert-info alert-dismissible"> 
				<h5 style="color: black; font-size:22px; font-weight:bold;"><i class="fa fa-exclamation-triangle"></i> Atencion!</h5>
				<h5 style="color: black; font-size:18px; font-weight:bold;"><%=session("MSJ_Fin_carga_pedido")%></h5> 
			</div>		
			<%session("MSJ_Fin_carga_pedido") = ""
		end if%>  
		 
		 
	</section>
	

</div>

    
  