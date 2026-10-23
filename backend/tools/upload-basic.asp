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

        <link href="plugins/jQuery-File-Upload-9.21.0/css/jquery.fileupload.css" rel="stylesheet">
        
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
                <h1> jQuery File Upload Demo</h1>
            </section>

            <!-- Main content -->
            <section class="content table-custome1">
                <div class="row">

                    <div class="col-xs-12">
                        <div class="box">
                            <div class="box-body body-form">


                                <h2 class="lead">Basic version</h2>
                                <br>
                                <!-- The fileinput-button span is used to style the file input field as button -->
                                <span class="btn btn-success fileinput-button">
                                    <i class="glyphicon glyphicon-plus"></i>
                                    <span>Seleccionar archivos...</span>
                                    <!-- The file input field used as target for the file upload widget -->
                                    <input id="fileupload" type="file" name="files[]" multiple>
                                </span>
                                <br>
                                <br>
                                <!-- The global progress bar -->
                                <div id="progress" class="progress">
                                    <div class="progress-bar progress-bar-success"></div>
                                </div>
                                <!-- The container for the uploaded files -->
                                <div id="files" class="files"></div>
                                <br>

                                <div class="panel panel-default">
                                    <div class="panel-heading">
                                        <h3 class="panel-title">Notas de esta versión</h3>
                                    </div>
                                    <div class="panel-body">
                                        <ul>
                                            <li>Esta versión carga automaticamente el archivo en el servidor y devuelve el nombre del archivo.</li>
                                        </ul>
                                    </div>
                                    
                                    <div class="panel-heading">
                                        <h3 class="panel-title">Notas de versión Jquery File Upload</h3>
                                    </div>
                                    <div class="panel-body">
                                        <ul>
                                            <li>The maximum file size for uploads in this demo is <strong>999 KB</strong> (default file size is unlimited).</li>
                                            <li>Only image files (<strong>JPG, GIF, PNG</strong>) are allowed in this demo (by default there is no file type restriction).</li>
                                            <li>Uploaded files will be deleted automatically after <strong>5 minutes or less</strong> (demo files are stored in memory).</li>
                                            <li>You can <strong>drag &amp; drop</strong> files from your desktop on this webpage (see <a href="https://github.com/blueimp/jQuery-File-Upload/wiki/Browser-support">Browser support</a>).</li>
                                            <li>Please refer to the <a href="https://github.com/blueimp/jQuery-File-Upload">project website</a> and <a href="https://github.com/blueimp/jQuery-File-Upload/wiki">documentation</a> for more information.</li>
                                            <li>Built with the <a href="http://getbootstrap.com/">Bootstrap</a> CSS framework and Icons from <a href="http://glyphicons.com/">Glyphicons</a>.</li>
                                        </ul>
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

        <!-- The jQuery UI widget factory, can be omitted if jQuery UI is already included -->
        <script src="plugins/jQuery-File-Upload-9.21.0/js/vendor/jquery.ui.widget.js"></script>
        <!-- The Iframe Transport is required for browsers without support for XHR file uploads -->
        <script src="plugins/jQuery-File-Upload-9.21.0/js/jquery.iframe-transport.js"></script>
        <!-- The basic File Upload plugin -->
        <script src="plugins/jQuery-File-Upload-9.21.0/js/jquery.fileupload.js"></script>

        <script src="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/js/bootstrap.min.js"></script>

        <!-- AdminLTE App -->
        <script src="dist/js/app.min.js" type="text/javascript"></script>

        <!-- page script -->
        <script type="text/javascript">

            $(function () {
                'use strict';
                // Change this to the location of your server-side upload handler:
                var url = 'plugins/jQuery-File-Upload-9.21.0/server/php/';
                $('#fileupload').fileupload({
                    url: url,
                    dataType: 'json',
                    done: function (e, data) {
                        $.each(data.result.files, function (index, file) {
                            $('<p/>').text(file.name).appendTo('#files');
                        });
                    },
                    progressall: function (e, data) {
                        var progress = parseInt(data.loaded / data.total * 100, 10);
                        $('#progress .progress-bar').css(
                            'width',
                            progress + '%'
                        );
                    }
                }).prop('disabled', !$.support.fileInput)
                    .parent().addClass($.support.fileInput ? undefined : 'disabled');
            });

        </script>
        <%''Fin personalizado JS de la página%>

    </body>

</html>