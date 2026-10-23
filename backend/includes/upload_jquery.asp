        <!-- The template to display files available for upload -->
        <script id="template-upload" type="text/x-tmpl">
        {% for (var i=0, file; file=o.files[i]; i++) { %}
            <tr class="template-upload fade">
                <td class="file-preview">
                    <span class="preview"></span>
                </td>
                <td class="file-filename">
                    <p class="name">{%=file.name%}</p>
                    <strong class="error text-danger"></strong>
                </td>
                <td class="file-filesize">
                    <p class="size">Subiendo archivo...</p>
                    <div class="progress progress-striped active" role="progressbar" aria-valuemin="0" aria-valuemax="100" aria-valuenow="0"><div class="progress-bar progress-bar-success" style="width:0%;"></div></div>
                </td>
                <td class="file-filebutton">
                    {% if (!i && !o.options.autoUpload) { %}
                        <button class="btn btn-primary start" disabled>
                            <i class="glyphicon glyphicon-upload"></i>
                            <span>Subir</span>
                        </button>
                    {% } %}
                    {% if (!i) { %}
                        <button class="btn btn-warning cancel">
                            <i class="glyphicon glyphicon-ban-circle"></i>
                            <span>Cancelar</span>
                        </button>
                    {% } %}
                </td>
            </tr>
        {% } %}
        </script>
        <!-- The template to display files available for download -->
        <script id="template-download" type="text/x-tmpl">
        {% for (var i=0, file; file=o.files[i]; i++) { %}
            <tr class="template-download fade">
                <td class="file-preview">
                    <span class="preview">
                        {% if (file.thumbnailUrl) { %}
                            <a href="{%=file.url%}" title="{%=file.name%}" download="{%=file.name%}" data-gallery><img src="{%=file.thumbnailUrl%}"></a>
                        {% } %}
                    </span>
                </td>
                <td class="file-filename">
                    <p class="name">
                        {% if (file.url) { %}
                            <a href="{%=file.url%}" title="{%=file.name%}" download="{%=file.name%}" {%=file.thumbnailUrl?'data-gallery':''%}>{%=file.name%}</a>
                        {% } else { %}
                            <span>{%=file.name%}</span>
                        {% } %}
                    </p>
                    {% if (file.error) { %}
                        <div><span class="label label-danger">Error</span> {%=file.error%}</div>
                    {% } %}
                </td>
                <td class="file-filesize">
                    <span class="size">{%=o.formatFileSize(file.size)%}</span>
                </td>
                <td class="file-filebutton">
                    {% if (file.deleteUrl) { %}
                        <button class="btn btn-danger delete temp-display-none" data-type="{%=file.deleteType%}" data-url="{%=file.deleteUrl%}"{% if (file.deleteWithCredentials) { %} data-xhr-fields='{"withCredentials":true}'{% } %}>
                            <i class="glyphicon glyphicon-trash"></i>
                            <span>Borrar</span>
                        </button>
                        <input type="checkbox" name="delete" value="1" class="toggle temp-display-none">
                    {% } else { %}
                        <button class="btn btn-warning cancel">
                            <i class="glyphicon glyphicon-ban-circle"></i>
                            <span>Cancelar</span>
                        </button>
                    {% } %}
                </td>
            </tr>
        {% } %}
        </script>

        <!-- The jQuery UI widget factory, can be omitted if jQuery UI is already included -->
        <script src="plugins/jQuery-File-Upload-9.21.0/js/vendor/jquery.ui.widget.js"></script>
        <!-- The Templates plugin is included to render the upload/download listings -->
        <script src="plugins/JavaScript-Templates/js/tmpl.min.js"></script>
        <!-- The Load Image plugin is included for the preview images and image resizing functionality -->
        <script src="plugins/JavaScript-Load-Image/js/load-image.all.min.js"></script>
        <!-- The Canvas to Blob plugin is included for image resizing functionality -->
        <script src="plugins/JavaScript-Canvas-to-Blob/js/canvas-to-blob.min.js"></script>
        <!-- blueimp Gallery script -->
        <script src="plugins/gallery/js/jquery.blueimp-gallery.min.js"></script>
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
        <%''Fin personalizado JS de la página%>

        <script>
            /*  inicializacion */
            $(function () {
                'use strict';
                
                var sorageData = [];

                function loadIconNoImageFile($thisNode, $data){
                    console.log($thisNode);
                    $thisNode.each(function(i, e){
                        console.log(e.text());
                        console.log(data.files[0].name);
                        if(e.text() == data.files[0].name){
                            e.parent('.template-upload').children('.file-preview .preview').html('es pdf');   
                        }
                    });
                }

                // Initialize the jQuery File Upload widget:
                $('.js--fileupload').fileupload({
                    // Uncomment the following to send cross-domain cookies:
                    //xhrFields: {withCredentials: true},
                    url: 'upload_set/uploadFile.php?type=' + $('.js--fileupload').attr('data-file-set-type'),
                    // Enable image resizing, except for Android and Opera,
                    // which actually support image resizing, but fail to
                    // send Blob objects via XHR requests:
                    disableImageResize: /Android(?!.*Chrome)|Opera/
                        .test(window.navigator.userAgent),
                    maxFileSize: 999000,
                    acceptFileTypes: /(\.|\/)(gif|jpe?g|png|gif|pdf|pptx|pptm|ppt|docx|docm|dotx|dotm|doc|dot|xls|xlt|xlsx|xlsm)$/i
                })
                //Preview de archivo
                .bind('fileuploadadd', function (e, data) {
                    //console.log(data);
                    //console.log(data.files[0].name);
                    //console.log($(this).children('.file-filename'));

                    console.log(e);
                    console.log(data);
                    console.log(data.files[0].size);

                    if(data.files[0].name.indexOf('.pdf') !== -1){
                        console.log('Es PDF');
                        //loadIconNoImageFile($(this).find('.file-filename p.name'), data);
                    }

                })
                //Si es exitosa la subida
                .bind('fileuploaddone', function (e, data) {
                    if(data._response.result.files[0].name.length > 0){
                        //console.log(data._response.result.files[0].name);
                        //console.log($(this));
                        //console.log(data._response.result.files[0]);

                        //Save data
                        //Guarda en DDBB los nombres de los archivos subidos
                        //Se guarda el nombre del archivo en la DDBB
                        //Los parametros de tabla y ID se setean en el formulario (data-file-db-table, data-file-db-id)
                        var vars = {
                            id : $(this).attr('data-file-db-id'),
                            tabla : $(this).attr('data-file-db-table'),  
                            imagen : data._response.result.files[0].name
                        };

                        if($(this).attr('data-file-db-table') != undefined || $(this).attr('data-file-db-table') != ''){
                            $.ajax({
                                data: vars,
                                url:   'upload_set/saveData.asp',
                                type:  'post',
                                dataType : 'html',			
                                beforeSend: function () {
                                    console.log('Enviando a DDBB...');
                                },
                                success:  function(response) {
                                    //console.log(response);
                                    if(response.indexOf('ok') !== -1){
                                        console.log('Guardado en DDBB!');
                                        $(document).trigger('fileLoadSuccessful');
                                    }else{
                                        console.log('ERROR al guardar!');
                                    }
                                            
                                }
                        
                            });
                        }
                        //

                        //Storage data
                        //Guarda en un input hidden con id #storageDataUpload el array con los nombres de los archivos subidos
                        if($('#storageDataUpload').length > 0){
                            sorageData.push(data._response.result.files[0].name);
                            $('#storageDataUpload').val(sorageData);
                        }
                        //
                    }
                })
                .bind('fileuploadfail', function (e, data) {
                    console.log('Error al cargar');
                })
                .bind('fileuploaddestroy', function(e, data) {
                    console.log('borrando...');
                    // If you edit the default template, you can acquire some other
                    // information about the uploaded file (for example the file size)
                    /*var fileName = data.context.find('a[download]').attr('download');

                    $.post(
                        'upload_set/deleteFile.php',
                        {
                            fileName: fileName
                        },
                        function(data, textStatus) {
                            // Process result
                        },
                        'json'
                    );*/

                });

                // Enable iframe cross-domain access via redirect option:
                $('.js--fileupload').fileupload(
                    'option',
                    'redirect',
                    window.location.href.replace(
                        /\/[^\/]*$/,
                        '/cors/result.html?%s'
                    )
                );

            });
        </script>