<%
pageName = "principal"
%>
<!DOCTYPE html>
<html lang="es">

    <head>
        
		<%''Común al sitema%>
        <!-- #include file="includes/head.asp" -->

        <%''Inicia personalizado CSS de la página%>
        <!-- DATA TABLES -->
        <link href="plugins/datatables/dataTables.bootstrap.css" rel="stylesheet" type="text/css" />
		<%''Fin personalizado CSS de la página%>
        
		<%''Files (CSS y js IE<9) que personalizan el sitema%>
        <!-- #include file="includes/head_template_custome.asp" -->
        
    </head>

    <body class="<%=Application("SiteSkin")%>">
	 
        <!-- header -->
        <!-- #include file="includes/header.asp" -->
		
        <!-- sidebar -->
        <!-- #include file="includes/sidebar.asp" -->

        <!-- home -->
        <!-- include file="includes/principal.asp" -->
		
        <!-- footer -->
        <!-- #include file="includes/footer.asp" -->
        
		<!-- scripts -->
        <!-- #include file="includes/scripts.asp" -->
        
        <%''Inicia personalizado JS de la página%>
		<script src="plugins/datatables/jquery.dataTables.js" type="text/javascript"></script>
        <script src="plugins/datatables/dataTables.bootstrap.js" type="text/javascript"></script>
        <!-- SlimScroll -->
        <script src="plugins/slimScroll/jquery.slimscroll.min.js" type="text/javascript"></script>
        <!-- FastClick -->
        <script src='plugins/fastclick/fastclick.min.js'></script>
        <!-- AdminLTE App -->
        <script src="dist/js/app.min.js" type="text/javascript"></script>
		<script src="https://cdn.jsdelivr.net/npm/sweetalert2@11"></script>
		
        <!-- page script -->
        <script type="text/javascript">
			function AbrirProd(id){
				$('.body-modal-2').load('PRO_detalle_dashBoard.asp?id='+id,function(){
					$('#myModal2').modal({show:true});
				});
			}
			var Call_Ajax= function(obj){			 
				$.ajax({
				
					url:   obj.url,
					type:  'post',
					dataType : 'html',			
					beforeSend: function () {
				
						$('#' + obj.nodeReload).fadeTo("slow", 0.30);
				
					},
					success: function(response) {
						console.log(response)
						if (response == 'NOK'){
							Swal.fire({
							  title: 'ERROR!',
							  text: 'El manifiesto ya se encuentra cargado',
							  icon: 'error',
							  confirmButtonText: 'Cerrar'
							})
						}	
						if (response == 'OK'){
							Swal.fire({
							  title: 'CORRECTO!',
							  text: 'El manifiesto no se encuentra cargado',
							  icon: 'success',
							  confirmButtonText: 'Cerrar'
							})
						}		
							
						$('#' + obj.nodeReload).fadeTo("slow", 1);
									
				
					} 
				
				});	

			}

			var Call_Ajax_Operador= function(obj){			 
				$.ajax({
				
					url:   obj.url,
					type:  'post',
					dataType : 'html',			
					beforeSend: function () {
				
						$('#' + obj.nodeReload).fadeTo("slow", 0.30);
				
					},
					success: function(response) {
						 
						if (response != ''){
							Swal.fire({
							  title: response,
							  text: '',
							  icon: 'info',
							  confirmButtonText: 'Cerrar'
							})
						} 	
							
						$('#' + obj.nodeReload).fadeTo("slow", 1);
									
				
					} 
				
				});	

			}
			
		   function BuscarManifiesto(){
				 
				 Call_Ajax({url:"ASP/Principal_XHR.asp?Manifiesto="+ $('#nroManifiesto').val(), nodeReload:"nroManifiesto"});			
			};
			
			function BuscarInfoOperador(){
				 
				 Call_Ajax_Operador({url:"ASP/Principal_Operador_XHR.asp?Residuo=" + $('#Residuo_Tipo1').val() + "&Categoría=" + $('#Residuo_Categoria1').val() + "&Generador="+ $('#Cliente_Generador').val(), nodeReload:"Generador"});			
			};
			
			var options1 = {
                
                adjustWidth: false,//responsive (no ajusta el ancho)

                //"phrase" es la variable con la que se envía el value del input
                url: function(phrase) {
                    return "plugins/easy-autocomplete/data/search.asp?m=1";
                },

                //Este método define con que nodo del JSON de retorno se completa la lista del autocomplete
                getValue: function(element) {
                    return element.nombre;
                },

                ajaxSettings: {
                    dataType: "json",
                    method: "POST",
                    //variables opcionales
                    data: {
                        varSend: "valor1",
                        varSend2: "valor2"
                    }
                },

                preparePostData: function(data) {
                    data.phrase = $("#Cliente_Generador").val();
                    return data;
                },
    

                list: {
                    maxNumberOfElements: 50, //(opcional) default 6

                    onChooseEvent: function(){
                     
                        $("#Cliente_GeneradorID").val($("#Cliente_Generador").getSelectedItemData().id);
						
						var GeneraID= $("#Cliente_Generador").getSelectedItemData().id;
						
						 
					}

                },
				 

                requestDelay: 1
            };
            $("#Cliente_Generador").easyAutocomplete(options1);		 
            //
			
			var options4_1 = {					
				adjustWidth: false, 					 
				url: function(phrase) {	return "plugins/easy-autocomplete/data/search.asp?m=4000";	}, 
				getValue: function(element) {return element.nombre;	},
				ajaxSettings: {	dataType: "json",method: "POST",data: {	varSend: "valor1",varSend2: "valor2"}},
				preparePostData: function(data) {data.phrase = $("#Residuo_Categoria1").val();return data;},
				list: {	maxNumberOfElements: 50, },	requestDelay: 1
			};
			$("#Residuo_Categoria1").easyAutocomplete(options4_1);
			
			var options5_1 = {					
				adjustWidth: false, 					 
				url: function(phrase) {	return "plugins/easy-autocomplete/data/search.asp?m=5000";	}, 
				getValue: function(element) {return element.nombre;	},
				ajaxSettings: {	dataType: "json",method: "POST",data: {	varSend: "valor1",varSend2: "valor2"}},
				preparePostData: function(data) {data.phrase = $("#Residuo_Tipo1").val();return data;},
				list: {	maxNumberOfElements: 50, },	requestDelay: 1
			};
			$("#Residuo_Tipo1").easyAutocomplete(options5_1);
			
        </script>
        <%''Fin personalizado JS de la página%>


		<!-- Modal -->
		<div class="modal fade" id="myModal2" role="dialog">
			<div class="modal-dialog" style="width:90%;">
				<!-- Modal content-->
				<div class="modal-content" >
					<div class="modal-header">
						<button type="button" class="close" data-dismiss="modal">X</button>
						<h4 class="modal-title">Detalle producto</h4>
					</div>
					<div class="modal-body body-modal-2">

					</div>
					
				</div>
			</div>
		</div>
		
		
    </body>

</html>
