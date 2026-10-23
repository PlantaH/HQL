<%
if request("t")<>"" then session("RPT_Tipo") = request("t")

pageGroup = "Reportes"
if session("RPT_Tipo") = 1 then
	pageName = "Reporteporfecha"
	titlePage = "Reporte por fecha"
end if
if session("RPT_Tipo") = 2 then
	pageName = "Reporteporproducto"
	titlePage = "Reporte por producto"
end if
if session("RPT_Tipo") = 3 then
	pageName = "Reporteporusuario"
	titlePage = "Reporte por usuario"
end if
session("draw") = 0
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
		
		  <!-- fullCalendar 2.2.5-->
		<link href="plugins/fullcalendar/fullcalendar.min.css" rel="stylesheet" type="text/css" />
		<link href="plugins/fullcalendar/fullcalendar.print.css" rel="stylesheet" type="text/css" media='print' />
  
		<!-- Ionicons -->
		<link href="plugins/Ionicons/css/ionicons.min.css"  rel="stylesheet" >
        
		<%''Files (CSS y js IE<9) que personalizan el sitema%>
        <!-- #include file="includes/head_template_custome.asp" -->
        
    </head>

    <body class="<%=Application("SiteSkin")%>">

        <!-- header -->
        <!-- #include file="includes/header.asp" -->
        
        <!-- sidebar -->
        <!-- #include file="includes/sidebar.asp" -->

        <!-- home -->
        <!-- #include file="reportes/RPT_Reporte_Estadistica.asp" -->

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
		
		 
        <%''Fin personalizado JS de la página%>
		<script>
		 
			 
			function abrirReporte(t){
				 
				document.getElementById("formABM").action = 'RPT_Reporte_' + t + '_Imprimir.asp' 
				//if (t == '1') document.getElementById("formABM").action = 'RPT_Reporte_1_Imprimir.asp' 
				
			
			}
			
			//			
			var options5 = {
                
                adjustWidth: false,//responsive (no ajusta el ancho)

                //"phrase" es la variable con la que se envía el value del input
                url: function(phrase) {
                    return "plugins/easy-autocomplete/data/search.asp?m=1&r_t=productos";
                },

                //Este método define con que nodo del JSON de retorno se completa la lista del autocomplete
                getValue: function(element) {
                    return element.campo;
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
                    data.phrase = $("#Producto").val();					
                    return data;
                },
				
				
				template: {
					type: "custom",
					method: function(value, item) {
					  let classSpan = '';
					 
					  let nodePais = (item.Info) ? "<div class='pais'><small>(" + item.Info + ")</small></div>" : "";

					  return "<span"+classSpan+"></span> " +  value + nodePais;
					}
				},
				
    

                list: {
                    maxNumberOfElements: 8, //(opcional) default 6
					
					onChooseEvent: function(){
                        $("#ProductoID").val($("#Producto").getSelectedItemData().ID);	
                    },
					onShowListEvent : function(){
					  $("#Producto + .easy-autocomplete-container").addClass('active');
					},
					onHideListEvent : function(){
					  $("#Producto + .easy-autocomplete-container").removeClass('active');
					}
                },

                requestDelay: 10
            };

            $("#Producto").easyAutocomplete(options5);
            //
			 
    </script>
     
    </body>

</html>
