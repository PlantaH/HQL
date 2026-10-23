<div class="user-saldo">
	<%if url_nuevo<>"" then%>
		<button class="btn btn-block btn-primary btn-lg" onclick="window.location='<%=url_nuevo%>'">Nuevo</button>
	<%end if%>	

	<%if url_duplicar<>"" then%>
		<button class="btn btn-block btn-primary btn-lg" onclick="window.location='<%=url_duplicar%>'">Duplicar Producto</button>
	<%end if%>	
 
	 
	
	<%if url_exportar<>"" then%>
	 
		
		<a class="btn btn-primary btn-sm" href="<%=url_exportar%>" download="<%=url_exportar%>" data-toggle="tooltip" data-placement="bottom" title="Descargue el archivo *.CSV para verlo en Excel"><i class="fa fa-download"></i> Descargar *.CSV</a>
			

	<%end if%>	
	
	<%if url_exportar_detalle<>"" then%>
	 
		
		<a class="btn btn-danger btn-sm"   href="<%=url_exportar_detalle%>" download="<%=url_exportar_detalle%>" data-toggle="tooltip" data-placement="bottom" title="Descargue el archivo *.CSV para verlo en Excel"><i class="fa fa-download"></i> Detalle *.CSV</a>
			

	<%end if%>	
	
	
	
	
	<%if rpt_exportar_pendientes<>"" then%>
	 
		
		<a class="btn btn-info btn-sm" href="<%=rpt_exportar_pendientes%>"   data-toggle="tooltip" data-placement="bottom" title="Descargue el reporte *.CSV para verlo en Excel"><i class="fa fa-download"></i> Descargar Pendientes *.CSV</a>
			

	<%end if%>	
	
	<%if rpt_importar_pendientes<>"" then%>
	 
		
		<a class="btn btn-warning btn-sm" href="<%=rpt_importar_pendientes%>"   data-toggle="tooltip" data-placement="bottom" title="Cargar el reporte *.CSV"><i class="fa fa-upload"></i> Importar Pendientes *.CSV</a>
			

	<%end if%>

    <%if rpt_exportar_remito<>"" then%>
	 
		
		<a class="btn btn-warning btn-sm" href="<%=rpt_exportar_remito%>"   data-toggle="tooltip" data-placement="bottom" title="Descargue el reporte *.CSV para verlo en Excel"><i class="fa fa-download"></i> Descargar Reporte Remitos *.CSV</a>
			

	<%end if%>	

	<%if rpt_exportar<>"" then%>
	 
		
		<a class="btn btn-primary btn-sm" href="<%=rpt_exportar%>"   data-toggle="tooltip" data-placement="bottom" title="Descargue el reporte *.CSV para verlo en Excel"><i class="fa fa-download"></i> Descargar *.CSV</a>
			

	<%end if%>		
		
	
	<%if url_actualizar_precios<>"" then%>
		<button class="btn btn-block btn-primary btn-lg" onclick="window.location='<%=url_actualizar_precios%>'">Actualizar los servicios listados</button>
	<%end if%>
	
	<%if rpt_importar_observaciones<>"" then%>
	 
		
		<a class="btn btn-warning btn-sm" href="<%=rpt_importar_observaciones%>"   data-toggle="tooltip" data-placement="bottom" title="Cargar el reporte *.CSV"><i class="fa fa-upload"></i> Importar Observaciones *.CSV</a>
			

	<%end if%>
	
	<%if rpt_importar_retirados<>"" then%>
	 
		
		<a class="btn btn-warning btn-sm" href="<%=rpt_importar_retirados%>"   data-toggle="tooltip" data-placement="bottom" title="Cargar el reporte *.CSV"><i class="fa fa-upload"></i> Importar Retirados *.CSV</a>
			

	<%end if%>
	
	<%if rpt_importar_dispuestos<>"" then%>
	 
		
		<a class="btn btn-warning btn-sm" href="<%=rpt_importar_dispuestos%>"   data-toggle="tooltip" data-placement="bottom" title="Cargar el reporte *.CSV"><i class="fa fa-upload"></i> Importar Dispuestos *.CSV</a>
			

	<%end if%>
	
	<%if rpt_importar_dispuestos_historico<>"" then%>
	 
		
		<a class="btn btn-warning btn-sm" href="<%=rpt_importar_dispuestos_historico%>"   data-toggle="tooltip" data-placement="bottom" title="Cargar el reporte *.CSV"><i class="fa fa-upload"></i> Importar Dispuestos *.CSV</a>
			

	<%end if%>
	
	<%if rpt_importar_derivaciones<>"" then%>
	 
		
		<a class="btn btn-warning btn-sm" href="<%=rpt_importar_derivaciones%>"   data-toggle="tooltip" data-placement="bottom" title="Cargar el reporte *.CSV"><i class="fa fa-upload"></i> Importar Derivaciones *.CSV</a>
			

	<%end if%>
	
	
	<%if rpt_importar_certificadosderivaciones<>"" then%>
	 
		
		<a class="btn btn-warning btn-sm" href="<%=rpt_importar_certificadosderivaciones%>"   data-toggle="tooltip" data-placement="bottom" title="Cargar el reporte *.CSV"><i class="fa fa-upload"></i> Importar Certificados *.CSV</a>
			

	<%end if%>
	
	
</div>