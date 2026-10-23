<%	
backURL = Application("Dominio")  & "usuarios/"
ID = request("ID")
pN = request("pN") 
 

if request.ServerVariables("REQUEST_METHOD")="POST" and request.form("boton")="ACEPTAR" then
	
		
	'***************************************************************************************************
	sSql="exec Permisos_Alta " & noErrorSQL(ID) & ",'" & noErrorSQL(request("cons")) & "','" & noErrorSQL(request("alta")) & "','" & noErrorSQL(request("baja")) & "','" & noErrorSQL(request("modi")) & "'"
	conn.execute(sSql)	
		
	'***************************************************************************************************
		
	response.redirect backURL 
end if 
		

sSql="exec Permisos_Cons " & ID
set rs=conn.execute(sSQL)

%>

		<!-- Right side column. Contains the navbar and content of the page -->
		<div class="content-wrapper">
			<!-- Content Header (Page header) -->
			<section class="content-header">
				<h1>
					Usuarios - Permisos
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
									
									<div class="form-group">
										
										<div class="row">									
											<div class="col-xs-12 col-sm-12">
											
											<table cellpadding="0" cellspacing="0" class="form-grilla-tabla">							
					
												<thead>
													<tr>												
														<th width="100%">Acceso</th>
														<th class="text-center"><input type="checkbox" name="cons_all" id="cons_all" class="flat-blue js--icheck-all" /><label for="cons_all"><small>Consulta</small></label></th>
														<th class="text-center"><input type="checkbox" name="modi_all" id="modi_all" class="flat-blue js--icheck-all" /><label for="modi_all"><small>Modificacion</small></label></th>
														<th class="text-center" ><input type="checkbox" name="alta_all" id="alta_all" class="flat-blue js--icheck-all" /><label for="alta_all"><small>Alta</small></label></th>
														<th class="text-center"><input type="checkbox" name="baja_all" id="baja_all" class="flat-blue js--icheck-all" /><label for="baja_all"><small>Baja</small></label></th>
													</tr>
												</thead>
												
												<tbody>
												<%
												While not rs.eof
												
													if grid = "grid2" then
														grid = "grid1"
													else
														grid = "grid2"
													end if
													%>
													<tr class="<%=grid%>">

														<td><strong><%=(rs("Menu"))%></strong>&nbsp;<img src="base/f.gif" align="absbottom">&nbsp;<%=(trim(replace(rs("Pagina")," ","")))%></td>
														<td class="text-center"><input type="checkbox" name="cons" id="c_<%=rs("PaginaID")%>" class="flat-blue<%if rs("consulta") then Response.Write " checked"%>" value="<%=rs("PaginaID")%>"<%if rs("consulta") then Response.Write " checked=""checked"""%> /></td>
														<td class="text-center"><input type="checkbox" name="modi" id="m_<%=rs("PaginaID")%>" class="flat-blue<%if rs("modi") then Response.Write " checked"%>" value="<%=rs("PaginaID")%>"<%if rs("modi") then Response.Write " checked=""checked"""%> /></td>
														<td class="text-center"><input type="checkbox" name="alta" id="a_<%=rs("PaginaID")%>" class="flat-blue<%if rs("alta") then Response.Write " checked"%>" value="<%=rs("PaginaID")%>"<%if rs("alta") then Response.Write " checked=""checked"""%> /></td>
														<td class="text-center"><input type="checkbox" name="baja" id="b_<%=rs("PaginaID")%>" class="flat-blue<%if rs("baja") then Response.Write " checked"%>" value="<%=rs("PaginaID")%>"<%if rs("baja") then Response.Write " checked=""checked"""%> /></td>
														
													</tr>
													<%
													registro_mostrado = registro_mostrado + 1
													
													ant_Menu = rs("MenuID")
																			
													rs.MoveNext
													
													if not rs.eof then
														if rs("MenuID")<>ant_Menu then%>
													<tr>
														<td class="form-grilla-separate" colspan="5"></td>
													</tr>
													<%
														end if
													end if
												
												Wend		
												%>
												</tbody>
												
											</table>
				
											</div>
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

			</section><!-- /.content -->

		</div>

