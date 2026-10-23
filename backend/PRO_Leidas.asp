<%
pageGroup = "Control"

if request("t")="i" then 
	session("PRO_Modo") = request("modo")
	session("MensajeColor") = ""
	session("MensajeError") = ""
end if

if session("PRO_Modo") = "E" then
	titPagina = "Carga Directa"
	pageName = "CargaDirecta"
	colorCaja = "#D4F2C9"
	
	linkRedir = Application("Dominio") & "entrada/n/"
end if
if session("PRO_Modo") = "S" then
	titPagina = "Desarga Directa"
	pageName = "DescargaDirecta"
	colorCaja = "#F29696"
	linkRedir = Application("Dominio") & "salida/n/"
end if


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
		  
		 
        <!-- #include file="ASP/PRO_Leidas.asp" -->

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
		 
			 
			
			document.getElementById("sol_Paquete").focus() 
		 
       
       
    </script>
     
    </body>

</html>
