<div class="wrap-login-box">

    <div class="login-box register-box">
        
      <div class="login-box-body">

		<div id="responseAjax"></div>
        
        <form id="registroForm" method="post">

            <div class="form-group form-register">
            	<label for=""><%=traduccion("Nombredelaempresa")%>:</label>
            	<span class="fa fa-user"></span>
            	<input type="text" class="form-control" name="nombre" id="nombre" placeholder="<%=traduccion("elnombredelaempresa")%>" required data-bv-notempty-message="<%=traduccion("Elnombredelaempresaesrequerido")%>"/>
            </div>
            <div class="form-group form-register">
            	<label for=""><%=traduccion("RazonSocial")%>:</label>
            	<span class="fa fa-user"></span>
            	<input type="text" class="form-control" name="RazonSocial" id="RazonSocial" placeholder="<%=traduccion("larazonsocial")%>" required data-bv-notempty-message="<%=traduccion("Larazonsocialesrequerida")%>"/>
            </div>
            <div class="form-group form-register">
            	<label for=""><%=traduccion("Pais")%>:</label>
            	<span class="fa fa-map-marker"></span>
                <select name="PaisID" id="PaisID" class="form-control">
					<%set pa = conn.execute("select * from paises order by pais")
					while not pa.eof%>
						<option value="<%=pa("PaisID")%>"><%=pa("Pais")%></option>
					<%pa.movenext
					wend%>
				</select>
            </div>
			<div class="form-group form-register">
            	<label for="Contacto"><%=traduccion("Nombredecontacto")%>:</label>
            	<span class="fa fa-user"></span>
            	<input type="text" class="form-control" name="Contacto" id="Contacto" placeholder="<%=traduccion("elnombredecontacto")%>" required data-bv-notempty-message="<%=traduccion("Elnombredecontactoesrequerido")%>"/>
            </div>
			<div class="form-group form-register">
            	<label for="Telefono"><%=traduccion("Telefono")%>:</label>
            	<span class="fa fa-mobile"></span>
            	<input type="text" class="form-control" name="Telefono" id="Telefono" placeholder="<%=traduccion("eltelefono")%>" required data-bv-notempty-message="<%=traduccion("Eltelefonoesrequerido")%>"/>
            </div>
			<div class="form-group form-register">
            	<label for="Mensajeria"><%=traduccion("Serviciodemensajeria")%>:</label>
            	<span class="fa fa-comments"></span>
            	<input type="text" class="form-control" name="Mensajeria" id="Mensajeria" placeholder="<%=traduccion("elserviciodemensajeria")%>" required data-bv-notempty-message="<%=traduccion("Elserviciodemensajeriaesrequerido")%>"/>
            </div>
			<div class="form-group form-register">
            	<label for="DatosConexion"><%=traduccion("Datosdeconexion")%>:</label>
            	<span class="fa fa-group"></span>
            	<input type="text" class="form-control" name="DatosConexion" id="DatosConexion" placeholder="<%=traduccion("losdatosparaconectarpormensajeria")%>" required data-bv-notempty-message="<%=traduccion("Losdatosdeconexionsonrequeridos")%>"/>
            </div>
            <div class="form-group form-register">
            	<label for="">e-mail:</label>
            	<span class="fa fa-envelope"></span>
            	<input type="email" class="form-control" name="email" id="email" placeholder="<%=traduccion("mailejemplo")%>" required data-bv-emailaddress-message="<%=traduccion("formerror_emailvalido")%>"
                data-bv-identical="true" data-bv-identical-field="email2" data-bv-identical-message="<%=traduccion("Elemailnocoincideconlaconfirmacion")%>"/>
            </div>
            <div class="form-group form-register">
            	<label for="">Confirmar e-mail:</label>
            	<span class="fa fa-envelope"></span>
            	<input type="text" class="form-control" name="email2" id="email2" placeholder="repite tu e-mail" required data-bv-notempty-message="Repetir la contraseña es requerido"
                data-bv-identical="true" data-bv-identical-field="email" data-bv-identical-message="El e-mail no coincide"/>
            </div>
            
            <div class="form-group form-register show-password">
            	<label for="">Contraseña:</label>
            	<span class="fa fa-unlock-alt"></span>
            	<input type="password" class="form-control" name="password" id="password" placeholder="define tu contraseña"
                required data-bv-notempty-message="La contraseña es requerida y no puede estar vacía"
                data-bv-identical="true" data-bv-identical-field="password2" data-bv-identical-message="La contraseña no coincide con la confirmación"
                data-bv-different="true" data-bv-different-field="nombre" data-bv-different-message="La contraseña no puede ser igual al nombre"/>
                <span class="show-password-btn fa fa-eye" title="Mostrar contraseña" data-title-on="Mostrar contraseña" data-title-off="Ocultar contraseña"></span>
            </div>
			<div class="form-group form-register">
            	<label for="">Website:</label>
            	<span class="fa fa-globe"></span>
            	<input type="text" class="form-control" name="PaginaTurismo" id="PaginaTurismo" placeholder="website de la empresa" required data-bv-message="El website de la empresa es requerido" />
            </div>
            <div class="form-group form-register show-password">
            	<label for="">Confirmar contraseña:</label>
            	<span class="fa fa-unlock-alt"></span>
            	<input type="password" class="form-control" name="password2" id="password2" placeholder="repite tu contraseña"
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
                                <input type="checkbox" name="terminos" required data-bv-message="Por favor debe aceptar los términos y condiciones">
                                Acepto los <a href="#" class="openBtn2" >Términos y condiciones</a>
                            </div>
							<!-- Modal -->
							<div class="modal fade" id="myModal2" role="dialog">
								<div class="modal-dialog">
									<!-- Modal content-->
									<div class="modal-content">
										<div class="modal-header">
											<button type="button" class="close" data-dismiss="modal">×</button>
											<h4 class="modal-title">Términos y condiciones</h4>
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
                                <button type="submit" class="btn btn-primary btn-block">Enviar registro</button>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
			
        </form>

      </div><!-- /.login-box-body -->
    </div><!-- /.login-box -->
    
</div>