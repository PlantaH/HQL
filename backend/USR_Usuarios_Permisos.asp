<%
pageGroup = "Usuarios"
pageName = "ABMUsuarios"
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

        <!-- iCheck for checkboxes and radio inputs -->
		<link href="plugins/iCheck/all.css" rel="stylesheet" type="text/css" />
        
		<%''Files (CSS y js IE<9) que personalizan el sitema%>
        <!-- #include file="includes/head_template_custome.asp" -->
        
    </head>

    <body class="<%=Application("SiteSkin")%>">

        <!-- header -->
        <!-- #include file="includes/header.asp" -->
        
        <!-- sidebar -->
        <!-- #include file="includes/sidebar.asp" -->

        <!-- home -->
        <!-- #include file="ASP/USR_Usuarios_Permisos.asp" -->
 
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
		
        <!-- iCheck 1.0.1 -->
		<script src="plugins/iCheck/icheck.min.js" type="text/javascript"></script>

        <%''Fin personalizado JS de la página%>

        <script>
		
            //Flat red color scheme for iCheck
            $('input[type="checkbox"].flat-blue, input[type="radio"].flat-blue').iCheck({
                checkboxClass: 'icheckbox_flat-blue',
                radioClass: 'iradio_flat-blue'
            });

            $('.js--icheck-all').each(function(i, e){

                var thisName = $(this).attr('name').replace('_all', '');
                //$(e).on('ifChanged', function(){
                $(e).on('ifChecked', function(){

                    $('input[name='+thisName+']').iCheck('check');

                });

                $(e).on('ifUnchecked', function(){

                    $('input[name='+thisName+']').iCheck('uncheck');

                });
                
            });

		</script>

    </body>

</html>
