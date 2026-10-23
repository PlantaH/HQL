<%
pageGroup = "micuenta"
pageName = "Misdatos"
%>
<!DOCTYPE html>
<html lang="es">

    <head>
        
		<%''Común al sitema%>
        <!-- #include file="includes/head.asp" -->

        <%''Inicia personalizado CSS de la página%>
        <!-- DATA TABLES -->
        <link href="plugins/datatables/dataTables.bootstrap.css" rel="stylesheet" type="text/css" />
		
		<link rel="stylesheet" href="plugins/bootstrapValidator/css/bootstrapValidator.css"/>
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
        <!-- #include file="includes/mis_datos.asp" -->

        <!-- footer -->
        <!-- #include file="includes/footer.asp" -->
        
        </div><!-- ./wrapper -->
        
		<!-- scripts -->
        <!-- #include file="includes/scripts.asp" -->
        
        <%''Inicia personalizado JS de la página%>
		
		<script src="plugins/datatables/jquery.dataTables.js" type="text/javascript"></script>
        <script src="plugins/datatables/dataTables.bootstrap.js" type="text/javascript"></script>
        <!-- SlimScroll -->
        <script src="plugins/slimScroll/jquery.slimscroll.min.js" type="text/javascript"></script>
        <!-- FastClick -->
        <script src='plugins/fastclick/fastclick.min.js'></script>
        
		<script src="plugins/main.js"></script>
		<!-- Validator -->
        <script src="plugins/bootstrapValidator/js/bootstrapValidator.js" type="text/javascript"></script>
		
		<!-- AdminLTE App -->
        <script src="dist/js/app.min.js" type="text/javascript"></script>

        <!-- page script -->
        <script type="text/javascript">
            $(function () {
                //$("#example1").dataTable();
                $('#example2').dataTable({
                    "bPaginate": true,
                    "bLengthChange": false,
                    "bFilter": false,
                    "bSort": true,
                    "bInfo": true,
                    "bAutoWidth": false,
                    "language": {
                        "url": "plugins/datatables/dataTables.spanish.lang.json"
                    }

                });	
                
                showPasword();

            });
			 		
			
        </script>
        <%''Fin personalizado JS de la página%>

    </body>

</html>
