<%
pageName = "modal"
%>
<!DOCTYPE html>
<html lang="es">

    <head>
        
		<%''Común al sitema%>
        <!-- #include file="../includes/head.asp" -->

        <%''Inicia personalizado CSS de la página%>
        <!-- CSS to style the file input field as button and adjust the Bootstrap progress bars -->
        <link rel="stylesheet" href="plugins/jQuery-File-Upload-9.21.0/css/jquery.fileupload.css">
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
                <h1> jQuery File Upload Demo</h1>
            </section>

            <!-- Main content -->
            <section class="content table-custome1">
                <div class="row">

                    <div class="col-xs-12">
                        <div class="box">
                            <div class="box-body body-form">


                                <h2 class="lead">Basic PLUS version</h2>
                                <br>
                                <!-- The fileinput-button span is used to style the file input field as button -->
                                <span class="btn btn-success fileinput-button">
                                    <i class="glyphicon glyphicon-plus"></i>
                                    <span>Add files...</span>
                                    <!-- The file input field used as target for the file upload widget -->
                                    <input id="fileupload" type="file" name="files[]" multiple>
                                </span>
                                <br>
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
                            
                                <!-- The blueimp Gallery widget -->
                                <div id="blueimp-gallery" class="blueimp-gallery blueimp-gallery-controls" data-filter=":even">
                                    <div class="slides"></div>
                                    <h3 class="title"></h3>
                                    <a class="prev">‹</a>
                                    <a class="next">›</a>
                                    <a class="close">×</a>
                                    <a class="play-pause"></a>
                                    <ol class="indicator"></ol>
                                </div>
                                <!-- /The blueimp Gallery widget -->

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

        <!-- The jQuery UI widget factory, can be omitted if jQuery UI is already included -->
        <script src="plugins/jQuery-File-Upload-9.21.0/js/vendor/jquery.ui.widget.js"></script>
        <!-- The Load Image plugin is included for the preview images and image resizing functionality -->
        <script src="https://blueimp.github.io/JavaScript-Load-Image/js/load-image.all.min.js"></script>
        <!-- The Canvas to Blob plugin is included for image resizing functionality -->
        <script src="https://blueimp.github.io/JavaScript-Canvas-to-Blob/js/canvas-to-blob.min.js"></script>
        <!-- blueimp Gallery script -->
        <script src="https://blueimp.github.io/Gallery/js/jquery.blueimp-gallery.min.js"></script>
        <!-- The Iframe Transport is required for browsers without support for XHR file uploads -->
        <script src="plugins/jQuery-File-Upload-9.21.0/js/jquery.iframe-transport.js"></script>
        <!-- The basic File Upload plugin -->
        <script src="plugins/jQuery-File-Upload-9.21.0/js/jquery.fileupload.js"></script>
        <!-- The File Upload processing plugin -->
        <script src="plugins/jQuery-File-Upload-9.21.0/js/jquery.fileupload-process.js"></script>
        <!-- The File Upload image preview & resize plugin -->
        <script src="plugins/jQuery-File-Upload-9.21.0/js/jquery.fileupload-image.js"></script>
        <!-- The File Upload audio preview plugin -->
        <script src="plugins/jQuery-File-Upload-9.21.0/js/jquery.fileupload-audio.js"></script>
        <!-- The File Upload video preview plugin -->
        <script src="plugins/jQuery-File-Upload-9.21.0/js/jquery.fileupload-video.js"></script>
        <!-- The File Upload validation plugin -->
        <script src="plugins/jQuery-File-Upload-9.21.0/js/jquery.fileupload-validate.js"></script>
        <!-- The File Upload user interface plugin -->
        <script src="plugins/jQuery-File-Upload-9.21.0/js/jquery.fileupload-ui.js"></script>

        <!-- The XDomainRequest Transport is included for cross-domain file deletion for IE 8 and IE 9 -->
        <!--[if (gte IE 8)&(lt IE 10)]>
        <script src="plugins/jQuery-File-Upload-9.21.0/js/cors/jquery.xdr-transport.js"></script>
        <![endif]-->

        <script>
        /*jslint unparam: true, regexp: true */
        /*global window, $ */
        $(function () {
            'use strict';
            // Change this to the location of your server-side upload handler:
            var url = 'plugins/jQuery-File-Upload-9.21.0/server/php/',
                uploadButton = $('<button/>')
                    .addClass('btn btn-primary')
                    .prop('disabled', true)
                    .text('Processing...')
                    .on('click', function () {
                        var $this = $(this),
                            data = $this.data();
                        $this
                            .off('click')
                            .text('Abort')
                            .on('click', function () {
                                $this.remove();
                                data.abort();
                            });
                        data.submit().always(function () {
                            $this.remove();
                        });
                    });
            $('#fileupload').fileupload({
                url: url,
                dataType: 'json',
                autoUpload: false,
                acceptFileTypes: /(\.|\/)(gif|jpe?g|png)$/i,
                maxFileSize: 999000,
                // Enable image resizing, except for Android and Opera,
                // which actually support image resizing, but fail to
                // send Blob objects via XHR requests:
                disableImageResize: /Android(?!.*Chrome)|Opera/
                    .test(window.navigator.userAgent),
                previewMaxWidth: 100,
                previewMaxHeight: 100,
                previewCrop: true
            }).on('fileuploadadd', function (e, data) {
                data.context = $('<div/>').appendTo('#files');
                $.each(data.files, function (index, file) {
                    var node = $('<p/>')
                            .append($('<span/>').text(file.name));
                    if (!index) {
                        node
                            .append('<br>')
                            .append(uploadButton.clone(true).data(data));
                    }
                    node.appendTo(data.context);
                });
            }).on('fileuploadprocessalways', function (e, data) {
                var index = data.index,
                    file = data.files[index],
                    node = $(data.context.children()[index]);
                if (file.preview) {
                    node
                        .prepend('<br>')
                        .prepend(file.preview);
                }
                if (file.error) {
                    node
                        .append('<br>')
                        .append($('<span class="text-danger"/>').text(file.error));
                }
                if (index + 1 === data.files.length) {
                    data.context.find('button')
                        .text('Upload')
                        .prop('disabled', !!data.files.error);
                }
            }).on('fileuploadprogressall', function (e, data) {
                var progress = parseInt(data.loaded / data.total * 100, 10);
                $('#progress .progress-bar').css(
                    'width',
                    progress + '%'
                );
            }).on('fileuploaddone', function (e, data) {
                $.each(data.result.files, function (index, file) {
                    if (file.url) {
                        var link = $('<a>')
                            .attr('target', '_blank')
                            .prop('href', file.url);
                        $(data.context.children()[index])
                            .wrap(link);
                    } else if (file.error) {
                        var error = $('<span class="text-danger"/>').text(file.error);
                        $(data.context.children()[index])
                            .append('<br>')
                            .append(error);
                    }
                });
            }).on('fileuploadfail', function (e, data) {
                $.each(data.files, function (index) {
                    var error = $('<span class="text-danger"/>').text('File upload failed.');
                    $(data.context.children()[index])
                        .append('<br>')
                        .append(error);
                });
            }).prop('disabled', !$.support.fileInput)
                .parent().addClass($.support.fileInput ? undefined : 'disabled');
        });
        </script>

        <%''Fin personalizado JS de la página%>

    </body>

</html>