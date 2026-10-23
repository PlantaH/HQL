<%
pageName = "modal"
%>
<!DOCTYPE html>
<html lang="es">

    <head>
        
		<%''Común al sitema%>
        <!-- #include file="../includes/head.asp" -->

        <%''Inicia personalizado CSS de la página%>

        <!-- blueimp Gallery styles -->
        <link rel="stylesheet" href="https://blueimp.github.io/Gallery/css/blueimp-gallery.min.css">

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
                <h1> Galerías de Imágenes</h1>
            </section>

            <!-- Main content -->
            <section class="content table-custome1">
                <h1>Link tipo lightbox</h1>
                <p>Fuente: <a href="https://github.com/blueimp/Gallery/blob/master/README.md" target="_blank">https://github.com/blueimp/Gallery/blob/master/README.md</a></p>
                <div class="row">

                    <div class="col-xs-12">
                        <div class="box">
                            <div class="box-body">
                        
                                <!-- links a las imágenes -->
                                <div id="links">
                                    <a href="tools/img/gal1.jpg" title="Título - Muestra de imagén 1">
                                        <img src="tools/img/thumbnails/gal1.jpg" alt="Gal 1">
                                    </a>
                                    <a href="tools/img/gal2.jpg" title="Título - Muestra de imagén 2">
                                        <img src="tools/img/thumbnails/gal2.jpg" alt="Gal 2">
                                    </a>
                                    <a href="tools/img/gal3.jpg" title="Título - Muestra de imagén 3">
                                        <img src="tools/img/thumbnails/gal3.jpg" alt="Gal 3">
                                    </a>
                                </div>



                                <!-- The Gallery as lightbox dialog, should be a child element of the document body -->
                                <!--<div id="blueimp-gallery" class="blueimp-gallery">-->

                                <!-- Si se le agraga la clase "blueimp-gallery-controls" los controles se mostrarán por default-->
                                <div id="blueimp-gallery" class="blueimp-gallery blueimp-gallery-controls">
                                    <div class="slides"></div>
                                    <h3 class="title"></h3>
                                    <a class="prev">‹</a>
                                    <a class="next">›</a>
                                    <a class="close">×</a>
                                    <a class="play-pause"></a>
                                    <ol class="indicator"></ol>
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

        <!-- blueimp Gallery script -->
        <script src="https://blueimp.github.io/Gallery/js/jquery.blueimp-gallery.min.js"></script>


        <!-- page script -->
        <script>
            /*document.getElementById('links').onclick = function (event) {
                event = event || window.event;
                var target = event.target || event.srcElement,
                    link = target.src ? target.parentNode : target,
                    options = {index: link, event: event},
                    links = this.getElementsByTagName('a');
                blueimp.Gallery(links, options);
            };*/
            $('#links').on('click', function(event){
                event = event || window.event;
                var target = event.target || event.srcElement,
                    link = target.src ? target.parentNode : target,
                    options = {index: link, event: event},
                    links = this.getElementsByTagName('a');
                blueimp.Gallery(links, options);
            });
        </script>
        <%''Fin personalizado JS de la página%>

    </body>

</html>