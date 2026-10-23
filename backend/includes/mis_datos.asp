<!-- Right side column. Contains the navbar and content of the page -->
<div class="content-wrapper">
	<!-- Content Header (Page header) -->
	<section class="content-header">
	  <h1>
		Mis datos
		<!--<small>advanced tables</small>-->
	  </h1>
	 
	</section>

	<!-- Main content -->
	<%
	MsjDatos = ""
	
	Set rs = conn.execute("exec  USR_Usuarios_Cons  -1," & session("usuarioID")& ",null")

	ID		  = rs("UsuarioID")	
	nombre    = rs("nombre")
	apellido  = rs("apellido")
	clave     = rs("clave")
	pass  = rs("password")
	email     = rs("email")
	telefono  = rs("telefono")
	empresa   = rs("empresa")
	if rs("estado") then
		Habilitado = " checked "
	else
		Habilitado = ""
	end if	
	
	if rs("master") then
		master = " checked "
	else
		master = ""
	end if
	
	if request.ServerVariables("REQUEST_METHOD")="POST" and request.form("QueCarga")="Datos"  then	
		
		set es = conn.execute("select * from usuarios where usuarioID<>" & session("usuarioID") & " and email='" & request("Email") & "' and password='" & request("password") & "'")
		
		if es.eof then
		
			Set Cmd = Server.CreateObject("ADODB.Command")
			Cmd.Activeconnection = conn
			Cmd.CommandText = "USR_Usuarios_Modi"
			Cmd.CommandType = 4
									
			Cmd.parameters.refresh				
			Cmd(1) = session("usuarioID")
			Cmd(2) = request("nombre")
			Cmd(3) = request("apellido")
			Cmd(4) = ""'clave
			Cmd(5) = request("password")
			Cmd(6) = request("email")
			Cmd(7) = request("telefono")
			Cmd(8) = "" 'empresa
			Cmd(9) = 1
			Cmd(10) = 0'master
															
			Cmd.execute,,adexecutenorecords
		
			session("Cliente_Nombre") = request("Nombre")
			session("Cliente_Email") = request("email")
			session("Cliente_RazonSocial") = "QUALITA" 'request("RazonSocial")
			'session("Cliente_PaisID") = request("PaisID")
			
			MsjDatos = "Los datos fueron actualizados con éxito!"
		else
			response.write "Error. Ya existe el usuario"
			response.End()
		end if
		
	end if%>
	
	<section class="content">
	  <div class="row">
		<div class="col-xs-12 col-lg-12">
		  <!-- general form elements disabled -->
		  <div class="box box-warning">
			<%if MsjDatos<>"" then%>
				<div class="box-header">
				  <h3 class="box-title"><%=MsjDatos%></h3>
					<div class="box-footer">
						<button type="button" class="btn btn-primary pull-right" onclick="window.location='./principal/'">continuar</button>
					</div>				  
				</div><!-- /.box-header -->
			<%else%>
				<div class="box-body body-form">
				  <form role="form"  method="post" id="formDatos">
					<input type="hidden" name="QueCarga" value="Datos">
					<div class="form-group form-register">
						<label for="">Nombre de la empresa:</label>
						<span class="fa fa-user"></span>
						<input type="text" class="form-control" name="nombre" id="nombre" value="<%=Nombre%>" placeholder="Nombre" required data-bv-notempty-message="El nombre es requerido"/>
					</div>
					<div class="form-group form-register">
						<label for="">Apellido:</label>
						<span class="fa fa-user"></span>
						<input type="text" class="form-control" name="Apellido" id="Apellido" value="<%=Apellido%>" placeholder="Apellido" />
					</div>
					
					<div class="form-group form-register">
						<label for="Contacto">Email:</label>
						<span class="fa fa-user"></span>
						<input type="text" class="form-control" name="Email" id="Email" placeholder="el nombre de email" value="<%=Email%>" required data-bv-notempty-message="El email de contacto es requerido"/>
					</div>
					<div class="form-group form-register">
						<label for="Telefono">Teléfono:</label>
						<span class="fa fa-mobile"></span>
						<input type="text" class="form-control" name="Telefono" id="Telefono" placeholder="el teléfono" value="<%=Telefono%>" />
					</div>
				 

					<div class="form-group form-register show-password">
						<label for="">Contraseña:</label>
						<span class="fa fa-unlock-alt"></span>
						<input type="password" class="form-control" name="password" id="password" value="<%=pass%>" placeholder="define tu contraseña"
						required data-bv-notempty-message="La contraseña es requerida y no puede estar vacía"
						data-bv-identical="true" data-bv-identical-field="password2" data-bv-identical-message="La contraseña no coincide con la confirmación"
						data-bv-different="true" data-bv-different-field="nombre" data-bv-different-message="La contraseña no puede ser igual al nombre"/>
						<span class="show-password-btn fa fa-eye" title="Mostrar contraseña" data-title-on="Mostrar contraseña" data-title-off="Ocultar contraseña"></span>
					</div>
					 
					<div class="form-group form-register show-password">
						<label for="">Confirmar contraseña:</label>
						<span class="fa fa-unlock-alt"></span>
						<input type="password" class="form-control" name="password2" id="password2" value="<%=pass%>" placeholder="repite tu contraseña"
						required data-bv-notempty-message="La confirmación de la contraseña es requerida y no puede estar vacía"
						data-bv-identical="true" data-bv-identical-field="password" data-bv-identical-message="La contraseña que confirmó no coincide"
						data-bv-different="true" data-bv-different-field="nombre" data-bv-different-message="La contraseña no puede ser igual al nombre"/>
						<span class="show-password-btn fa fa-eye" title="Mostrar contraseña" data-title-on="Mostrar contraseña" data-title-off="Ocultar contraseña"></span>
					</div>
					

					<div class="register-footer">
						<div class="row">
							<div class="wrap-footer-group">
								<div class="footer-group">
									<div class="form-group col-xs-12 register-footer-text">
										 
									</div>
									<!-- Modal -->
									<div class="modal fade" id="myModal2" role="dialog">
										<div class="modal-dialog">
											<!-- Modal content-->
											<div class="modal-content">
												<div class="modal-header">
													 
												</div>
												<div class="modal-body body-modal-2">

												</div>
												<div class="modal-footer">
													<button type="button" class="btn btn-default" data-dismiss="modal">Cerrar</button>
												</div>
											</div>
										</div>
									</div>							
									<div class="col-xs-12 col-md-4">
										<button type="submit" class="btn btn-primary btn-block">Guardar</button>
									</div>
								</div>
							</div>
						</div>
					</div>
				  </form>
				</div><!-- /.box-body -->		
			<%end if%>
		  </div><!-- /.box -->
		</div><!--/.col (right) -->        
	  </div>   <!-- /.row -->
	</section><!-- /.content -->
</div><!-- /.content-wrapper -->