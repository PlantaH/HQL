<%
if request("tipo")<>"" then session("USR_Tipo") = request("tipo")

if session("USR_Tipo") = "USR" then
	pageGroup = "Usuarios"
	pageName = "ABMUsuarios"
end if

if session("USR_Tipo") = "ENC" then
	pageGroup = "Usuarios"
	pageName = "ABMEncargados"
end if

if session("USR_Tipo") = "OPE" then
	pageGroup = "Usuarios"
	pageName = "ABMOperadores"
end if

session("draw") = 0


URL_XHR  = "USR_Usuarios_XHR.asp"
URL_ALTA = "usuarios-modi"
URL_BAJA = "usuarios-baja"
URL_PERM = "usuarios-permisos"


%>
<!DOCTYPE html>
<html lang="es">

    <head>
        
		<%''Común al sitema%>
        <!-- #include file="includes/head.asp" -->

        <%''Inicia personalizado CSS de la página%>
        <!-- DATA TABLES -->
        <link href="plugins/datatables/dataTables.bootstrap.css" rel="stylesheet" type="text/css" />
        <link href="plugins/datatables/jquery.datatables.checkboxes/1.2.11/css/dataTables.checkboxes.css" rel="stylesheet" type="text/css" />
		<%''Fin personalizado CSS de la página%>
        
		<%''Files (CSS y js IE<9) que personalizan el sitema%>
        <!-- #include file="includes/head_template_custome.asp" -->
		
    </head>
<!--skin-black-light-->
    <body class="<%=Application("SiteSkin")%>">

        <!-- header -->
        <!-- #include file="includes/header.asp" -->
        
        <!-- sidebar -->
        <!-- #include file="includes/sidebar.asp" -->

        <!-- home -->
        <!-- #include file="ASP/USR_Usuarios_Principal.asp" -->

        <!-- footer -->
        <!-- #include file="includes/footer.asp" -->
        
		<!-- scripts -->
        <!-- #include file="includes/scripts.asp" -->
        
        <%''Inicia personalizado JS de la página%>
		<!--<script src="plugins/datatables/jquery.dataTables.js" type="text/javascript"></script>-->

        <script src="plugins/datatables/datatables.min.js" type="text/javascript"></script>
		
		<script src="plugins/datatables/jquery.dataTables.min.js" type="text/javascript"></script>
		
        <!--<script src="plugins/datatables/dataTables.bootstrap.js" type="text/javascript"></script>-->
     
        <script src="plugins/datatables/dataTables.bootstrap.min.js" type="text/javascript"></script>

        <script src="plugins/datatables/jquery.datatables.checkboxes/1.2.11/js/dataTables.checkboxes.min.js" type="text/javascript" ></script>

        <!-- SlimScroll -->
        <script src="plugins/slimScroll/jquery.slimscroll.min.js" type="text/javascript"></script>
        <!-- FastClick -->
        <script src='plugins/fastclick/fastclick.min.js'></script>
        <!-- AdminLTE App -->
        <script src="dist/js/app.min.js" type="text/javascript"></script>

		
        <!-- page script -->
        <script type="text/javascript">		
			
            $(function () {			
                $('#tablaDatos').dataTable({
										
                    //Opcional, se usa para ocultar columna/s
                    //En caso de usarse se deben pasar la cantidad total de columnas
                    "columns": [
                    { "visible": true},
                    { "visible": true },
                    { "visible": true }, 
                    { "visible": true },
                    { 
                        "visible": true, "orderable": false, "targets": 0,
                        render: function (o) { return '<i class="fa fa-key table-item-permission" data-original-title="Permission" data-toggle="datatable-tooltip" data-placement="top" title="Permisos"></i>'; }
                    },
                    { 
                        "visible": true, "orderable": false, "targets": 0,
                        render: function (o) { return '<i class="fa fa-pencil table-item-edit" data-original-title="Editar" data-toggle="datatable-tooltip" data-placement="top" title="Editar"></i>'; }
                    },					
                    { 
                        "visible": true, "orderable": false, "targets": 0,
                        render: function (o) { return '<i class="fa fa-trash-o table-item-delete" data-original-title="Delete" data-toggle="datatable-tooltip" data-placement="top" title="Eliminar"></i>'; }
                    }
                    ],
                    //
					
					"columnDefs": [
                    {"targets": 0,   "checkboxes": { "selectRow": true }   },
					{ "width": "5%", "targets": 3 },
					{ "width": "5%", "targets": 4 },
                    { "width": "5%", "targets": 5 },
                    { "width": "5%", "targets": 6 }
					],
					
                    "displayLength": 10,
                    "lengthMenu": [[10, 20, 25, 50], [10, 20, 25, 50]],
                    "responsive": true,
                    "processing": true,
                    "serverSide": true,
                    "ajax": {
                        url: "ASP/<%=URL_XHR%>", // json datasource
                        //data: {action: 'getEMP2'},
                        type: 'post'  // method  , by default get
                    },
                    "language": {
                        "url": "<%=Application("SiteLang")%>"
                    }

                });
				 
            });
			 
			$(document).ready(function() {

                var table = $('#tablaDatos').DataTable();	
                
                $('#frm-example').on('submit', function(e){

                    e.preventDefault();
                    
                    var form = this,
                        rows_selected = table.column(0).checkboxes.selected(),
                        info = table.page.info(),
                        pageNmr = info.page;

                    if(rows_selected.length > 0){

                        // Iterate over all selected checkboxes
                        $.each(rows_selected, function(index, rowId){
                            // Create a hidden element 
                            $(form).append(
                                $('<input>')
                                    .attr('type', 'hidden')
                                    .attr('name', 'id[]')
                                    .val(rowId)
                            );
                        });
                                    
                        if(window.confirm("Desea eliminar todos los registros seleccionados?")) { 
                            window.location = "<%=URL_BAJA%>/" + rows_selected.join("-") + "/" + pageNmr + "/"
                        };
                        
                    }else{

                        alert('Debe seleccionar a menos un registro!');

                    }
				   
			    });
							 
		
				$('#tablaDatos tbody').on('click', 'tr', function() {				
					var data = table.row(this).data();					 					
				} );
				
				$('#tablaDatos tbody').on('click', 'td', function () {					
					var rowIdx = table.cell(this).index().row;				 
					var data = table.row(rowIdx).data();		

					var info = table.page.info();
                    var pageNmr = info.page;
					
					if((table.cell(this).index().columnVisible != 0) && (table.cell(this).index().columnVisible != 4) && (table.cell(this).index().columnVisible != 6)) {
						window.location = "<%=URL_ALTA%>/" + data[0] + "/" + pageNmr + "/"
					}else{
						if(table.cell(this).index().columnVisible == 6) {
							if (window.confirm("Desea eliminar?")) { 
								window.location = "<%=URL_BAJA%>/" + data[0] + "/" + pageNmr + "/"
							};
						}else if(table.cell(this).index().columnVisible == 4){
							<%if session("USR_Tipo") = "USR" then %>
							window.location = "<%=URL_PERM%>/" + data[0] + "/" + pageNmr + "/"
							<%end if%>
						};
					};
                });

                //Tooltip en tablas
                setTimeout(() => {
                    $('[data-toggle="datatable-tooltip"]').tooltip();
                }, 800);
                
			});
			
			
			
        </script>
        <%''Fin personalizado JS de la página%>

    </body>

</html>
