<%	
if session("USR_Tipo") = "USR" then 
	titulo = "Usuarios"
	backURL = Application("Dominio")  & "usuarios/"
end if

if session("USR_Tipo") = "ENC" then 
	titulo = "Encargados"
	backURL = Application("Dominio")  & "encargados/"
end if

if session("USR_Tipo") = "OPE" then 
	titulo = "Operadores"
	backURL = Application("Dominio")  & "operadores/"
end if

ID = request("ID")
pN = request("pN") 
 

if request.ServerVariables("REQUEST_METHOD")="POST" and request.form("boton")="ACEPTAR" then
	
		
	'***************************************************************************************************
	ID		  = request("ID")
	nombre    = request("nombre")
	apellido  = request("apellido")
	clave     = request("clave")
	password  = request("password")
	email     = request("email")
	telefono  = request("telefono")
	empresa   = request("empresa")
	if request("habilitado")="on" then
		Habilitado = 1
	else
		Habilitado = 0
	end if
	
	if request("master")="on" then
		master = 1
	else
		master = 0
	end if	
	
	if request("ID") = "" then
		Set Cmd = Server.CreateObject("ADODB.Command")
		Cmd.Activeconnection = conn
		Cmd.CommandText = "USR_Usuarios_Alta"
		Cmd.CommandType = 4
								
		Cmd.parameters.refresh			
		Cmd(1) = nombre
		Cmd(2) = apellido
		Cmd(3) = clave
		Cmd(4) = password
		Cmd(5) = email
		Cmd(6) = telefono
		Cmd(7) = empresa
		Cmd(8) = Habilitado
		Cmd(9) = 0'master
		Cmd(10)= session("USR_Tipo")
														
		Cmd.execute,,adexecutenorecords	
	else
				
		Set Cmd = Server.CreateObject("ADODB.Command")
		Cmd.Activeconnection = conn
		Cmd.CommandText = "USR_Usuarios_Modi"
		Cmd.CommandType = 4
								
		Cmd.parameters.refresh				
		Cmd(1) = ID
		Cmd(2) = nombre
		Cmd(3) = apellido
		Cmd(4) = clave
		Cmd(5) = password
		Cmd(6) = email
		Cmd(7) = telefono
		Cmd(8) = empresa
		Cmd(9) = Habilitado
		Cmd(10) = 0'master
														
		Cmd.execute,,adexecutenorecords
	end if
		
	'***************************************************************************************************
		
	response.redirect backURL 
end if 
		

if ID<>"" then
	
		
	Set rs = conn.execute("exec  USR_Usuarios_Cons  -1," & ID & ",null")

	ID		  = rs("UsuarioID")	
	nombre    = rs("nombre")
	apellido  = rs("apellido")
	clave     = rs("clave")
	password  = rs("password")
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
else	
	empresa = "CONECTTA S.A."
end if



%>

      <!-- Right side column. Contains the navbar and content of the page -->
      <div class="content-wrapper">
        <!-- Content Header (Page header) -->
        <section class="content-header">
          <h1>
            <%=titulo%>
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
									<div class="form-group col-xs-12 col-sm-6">
										<label>Nombre</label>
										<input type="text" class="form-control" name="Nombre" id="Nombre" maxlength="50" placeholder="ingrese el nombre ..." required data-bv-notempty-message="El nombre es requerido" value="<%=Nombre%>" />
									</div>
								 								
									<div class="form-group col-xs-12 col-sm-6">
										<label>Apellido</label>
										<input type="text" class="form-control" name="Apellido" id="Apellido" placeholder="Ingrese el Apellido" required data-bv-notempty-message="El Apellido es requerido" value="<%=Apellido%>"/>
									</div>
								</div>
								
								
								<div class="row">									
									<div class="form-group col-xs-12 col-sm-6">
										<label>Usuario</label>
										<input type="text" class="form-control" name="clave" id="clave" maxlength="50" placeholder="ingrese el usuario ..." required data-bv-notempty-message="El usuario es requerido" value="<%=clave%>" />
									</div>
								 								
									<div class="form-group col-xs-12 col-sm-6 show-password">
										<label>Clave</label>
										<input type="password" class="form-control" name="password" id="password" required data-bv-notempty-message="La clave es requerida"  placeholder="Ingrese la clave" value="<%=password%>"/>
										<span class="show-password-btn fa fa-eye" title="Mostrar contraseña" data-title-on="Mostrar contraseña" data-title-off="Ocultar contraseña"></span>
									</div>
								</div>
								
								<div class="row">									
									<div class="form-group col-xs-12 col-sm-6">
										<label>Email</label>
										<!--<input type="email" class="form-control" name="email" id="email" maxlength="50" pattern="[a-zA-Z0-9_]+([.][a-zA-Z0-9_]+)*@[a-zA-Z0-9_]+([.][a-zA-Z0-9_]+)*[.][a-zA-Z]{1,5}"  required data-bv-notempty-message="El email es requerido"  placeholder="ingrese el email ..."  value="<%=email%>" />-->
										<input type="email" class="form-control" name="email" id="email" placeholder="e-mail@ejemplo.com.ar" required data-bv-notempty-message="El emil es requerido" data-bv-emailaddress-message="El formato de e-mail no es válido" value="<%=email%>" />
									</div>
								 								
									<div class="form-group col-xs-12 col-sm-6">
										<label>Telefono</label>
										<input type="number" class="form-control" name="telefono" id="telefono" placeholder="Ingrese el telefono" value="<%=telefono%>"/>
									</div>
								</div>
								
								<div class="row">									
									<div class="form-group col-xs-12 col-sm-6">
										<label>Empresa</label>
										<input type="text" class="form-control" name="empresa" id="empresa" required data-bv-notempty-message="La empresa es requerida" maxlength="50" placeholder="ingrese la empresa ..." value="<%=empresa%>" />
									</div>
								</div>
							</div>
							
							
							
							<div class="form-group">
							  <label>Habilitado</label>
							  <input type="checkbox" class="flat-blue <%=Habilitado%>" id="Habilitado" name="Habilitado" <%=Habilitado%>>
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

