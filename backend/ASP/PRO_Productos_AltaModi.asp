<%	
backURL = Application("Dominio")  & "productos/"
ID = request("ID")
pN = request("pN") 

session("ClienteID_Work") = 0
session("backURL_Detalle") = "productos-modi/" & ID & "/" & pN & "/" 

if request.ServerVariables("REQUEST_METHOD")="POST" and request.form("boton")="ACEPTAR" then
	
		
	'***************************************************************************************************
	 
	
	if request("ID") = "" then
		Set Cmd = Server.CreateObject("ADODB.Command")
		Cmd.Activeconnection = conn
		Cmd.CommandText = "PRO_Productos_Alta"
		Cmd.CommandType = 4
							
		Cmd.parameters.refresh
				
		Cmd(1) = request("marca")	
		Cmd(2) = request("nombre")
		Cmd(3) = request("codigo")	
		Cmd(4) = 0'request("cantidad")
		Cmd(5) = request("CodigoVenta")
								
		Cmd.execute,,adexecutenorecords	
	else
		Set Cmd = Server.CreateObject("ADODB.Command")
		Cmd.Activeconnection = conn
		Cmd.CommandText = "PRO_Productos_Modi"
		Cmd.CommandType = 4
							
		Cmd.parameters.refresh
				
		Cmd(1) = request("ID")
		Cmd(2) = request("marca")	
		Cmd(3) = request("nombre")
		Cmd(4) = request("codigo")	
		Cmd(5) = 0'request("cantidad")
		Cmd(6) = request("CodigoVenta")
														
		Cmd.execute,,adexecutenorecords	
	end if
		
	'***************************************************************************************************
		
	response.redirect backURL 
end if 
		

if ID<>"" then
	
	sSql = "select * from productos where ID = " & ID
	set rs=conn.execute(sSql)
		
	ID		= rs("ID")	
	Marca = rs("Marca")
	Nombre = rs("Nombre") 
	codigo = rs("codigo")	
	cantidad = rs("cantidad")
	CodigoVenta = rs("CodigoVenta")
	
	 
end if
%>

      <!-- Right side column. Contains the navbar and content of the page -->
      <div class="content-wrapper">
        <!-- Content Header (Page header) -->
        <section class="content-header">
          <h1>
            Productos
            <!--<small>advanced tables</small>-->
          </h1>
          <!-- include file="botones.asp" -->
        </section>

        <!-- Main content -->
		
		
			<section class="content">
			  
			  <div class="row">
				<div class="col-xs-12 col-lg-12">
				  <!-- general form elements disabled -->
				  <div class="box box-warning">
					<%if MsjValidado<>"" then%>
						<div class="box-header">
						  <h3 class="box-title"><%=MsjValidado%></h3>			  
						</div><!-- /.box-header -->
					<%else%>
					 
						<div class="box-body body-form">
						  <form role="form"  method="post" id="formABM"> 
							<input type="hidden" name="boton" value="ACEPTAR">
							<input type="hidden" name="ID" value="<%=ID%>">
							<input type="hidden" name="pN" value="<%=pN%>">
							
							<div class="">
								
								<div class="row">
 
							
									<div class="col-xs-12 col-sm-6">
										<label>Marca</label>
										<input type="text" class="form-control" name="Marca" id="Marca" maxlength="50" placeholder="ingrese el campo ..."    value="<%=Marca%>" />
									</div>
									<div class="col-xs-12 col-sm-6">
										<label>Nombre</label>
										<input type="text" class="form-control" name="Nombre" id="Nombre" maxlength="Nombre50" placeholder="ingrese el campo ..." required data-bv-notempty-message="El codigo es requerido"   value="<%=Nombre%>" />
									</div>
								</div>
								<div class="row">								
									
									<div class="col-xs-6 col-sm-6">
										<label>Codigo Venta</label>
										<input type="text" class="form-control" name="CodigoVenta" id="CodigoVenta" maxlength="50" placeholder="ingrese el campo ..."  value="<%=CodigoVenta%>" />
									</div>
									<div class="col-xs-6 col-sm-6">
										<label>Codigo Barras</label>
										<input type="text" class="form-control" name="Codigo" id="Codigo" maxlength="50" placeholder="ingrese el campo ..." required data-bv-notempty-message="El codigo es requerido" value="<%=Codigo%>" />
									</div>
									
									<!--<div class="col-xs-6 col-sm-2">
										<label>Cantidad</label>
										<input type="number"  min="0" max="10000" step="1" class="form-control" name="Cantidad" id="Cantidad" placeholder="ingrese el campo ..." required data-bv-notempty-message="El campo es requerido" value="<%=Cantidad%>"/>
									</div>-->
								 			 
								</div>
								 
							</div>
							 
							
							<div class="box-footer">
								<button type="submit" class="btn btn-primary">Guardar</button>
								
								<button type="button" class="btn btn-primary" onclick="window.location = '<%=backURL%>'">Cancelar</button>
							</div>
						  </form>
						</div><!-- /.box-body -->		
					<%end if%>
				  </div><!-- /.box -->
				</div><!--/.col (right) -->
			  </div>   <!-- /.row -->
			<!--</section>--><!-- /.content -->
		 
		 
      </div><!-- /.content-wrapper -->

