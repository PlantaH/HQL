<%
pageName = "modal"
%>
<!DOCTYPE html>
<html lang="es">

    <head>
        
		<%''Común al sitema%>
        <!-- #include file="../includes/head.asp" -->

        <%''Inicia personalizado CSS de la página%>
        <!-- DATA TABLES -->
        <link href="plugins/datatables/dataTables.bootstrap.css" rel="stylesheet" type="text/css" />
		<%''Fin personalizado CSS de la página%>
        
		<%''Files (CSS y js IE<9) que personalizan el sitema%>
        <!-- #include file="../includes/head_template_custome.asp" -->
        
    </head>

    <body class="skin-blue">

        <!-- header -->
        <!-- #include file="../includes/header.asp" -->
        
        <!-- sidebar -->
        <!-- #include file="../includes/sidebar.asp" -->

        <!-- modal -->
        <div class="content-wrapper"> 
            <!-- Content Header (Page header) -->
            <section class="content-header">
                <h1> Ventanas modales</h1>
            </section>

            <!-- Main content -->
            <section class="content table-custome1">
                <h1> Ventana modal estática</h1>
                <div class="row">

                    <div class="col-xs-12">
                        <div class="box">
                            <div class="box-body">
                        
                                <!-- Trigger the modal with a button -->
                                <button type="button" class="btn btn-info btn-lg" data-toggle="modal" data-target="#myModal1">Open Modal</button>

                                <!-- Modal -->
                                <div id="myModal1" class="modal fade" role="dialog">
                                    <div class="modal-dialog">

                                        <!-- Modal content-->
                                        <div class="modal-content">
                                            <div class="modal-header">
                                                <button type="button" class="close" data-dismiss="modal">&times;</button>
                                                <h4 class="modal-title">Modal Header</h4>
                                            </div>
                                        <div class="modal-body">
                                            <p>Some text in the modal.</p>
                                        </div>
                                            <div class="modal-footer">
                                                <button type="button" class="btn btn-default" data-dismiss="modal">Close</button>
                                            </div>
                                        </div>

                                    </div>
                                </div>

                            </div>
                            <!-- /.box-body --> 
                        </div>
                        <!-- /.box --> 
                    </div>
                    <!-- /.col --> 

                </div>
                <!-- /.row --> 
            </section>
            <!-- /.content -->




           <!-- Main content -->
            <section class="content table-custome1">
                <h1> Ventana modal dinámica 1 (Carga contenido externo, HTML, ASP...)</h1>
                <div class="row">

                    <div class="col-xs-12">
                        <div class="box">
                            <div class="box-body">
                        
                                <!-- Trigger the modal with a button -->
                                <button type="button" class="btn btn-success openBtn2">Open Modal</button>

                                <!-- Modal -->
                                <div class="modal fade" id="myModal2" role="dialog">
                                    <div class="modal-dialog">
                                        <!-- Modal content-->
                                        <div class="modal-content">
                                            <div class="modal-header">
                                                <button type="button" class="close" data-dismiss="modal">×</button>
                                                <h4 class="modal-title">Modal with Dynamic Content</h4>
                                            </div>
                                            <div class="modal-body body-modal-2">

                                            </div>
                                            <div class="modal-footer">
                                                <button type="button" class="btn btn-default" data-dismiss="modal">Close</button>
                                            </div>
                                        </div>
                                    </div>
                                </div>

                            </div>
                            <!-- /.box-body --> 
                        </div>
                        <!-- /.box --> 
                    </div>
                    <!-- /.col --> 

                </div>
                <!-- /.row --> 
            </section>
            <!-- /.content -->


        </div>
        <!-- /.content-wrapper --> 

        <!-- footer -->
        <!-- #include file="../includes/footer.asp" -->
        
        </div><!-- ./wrapper -->
        
		<!-- scripts -->
        <!-- #include file="../includes/scripts.asp" -->
        
        <%''Inicia personalizado JS de la página%>

        <!-- AdminLTE App -->
        <script src="dist/js/app.min.js" type="text/javascript"></script>

        <!-- page script -->
        <script type="text/javascript">

                $('.openBtn2').on('click',function(){
                    $('.body-modal-2').load('tools/contenido.asp',function(){
                        $('#myModal2').modal({show:true});
                    });
                });

        </script>
        <%''Fin personalizado JS de la página%>

    </body>

</html>