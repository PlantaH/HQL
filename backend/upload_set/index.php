<?php
/*
 * jQuery File Upload Plugin PHP Example
 * https://github.com/blueimp/jQuery-File-Upload
 *
 * Copyright 2010, Sebastian Tschan
 * https://blueimp.net
 *
 * Licensed under the MIT license:
 * https://opensource.org/licenses/MIT
 */

error_reporting(E_ALL | E_STRICT);
require('UploadHandler.class.php');

//Extiende la clase original
//Customiza el nombre del archivo
class CustomUploadHandler extends UploadHandler {
    protected function trim_file_name($file_path, $name, $size, $type, $error, $index, $content_range) {
        $name = 'comprobante_' . microtime(true);
        $name = str_replace('.', '', $name);
        return $name;
    }
}

//Instancia origianl
//$upload_handler = new UploadHandler();

//Intancia extendida y customizada
$upload_handler = new CustomUploadHandler(array(
    //Tamaño máximo del archivo por default es libre
    //'max_file_size' => 1048576, //1MB file size
    //Path donde se guarda el alchivo (relativo a este alchivo)
    'upload_dir' => '../upload/comprobantes/', 
    //Path del preview (relativo a la pagina de upload)
    //Si nose define 'imagen_versions' la miniatura se guarda en el directorio "thumbnail" (si no existe el directorio, es creado en el momento)
    'upload_url' => 'upload/comprobantes/',
    //'script_url' => '/url/of/uploadhandler/script',
    'accept_file_types' => '/\.(jpe?g|png|pdf)$/i',

    //Versionados del archivo (Se usa cuando subimos imagenes y queremos guardar diferentes tamaños)
    //Si se define se debe definir además el 'thumbnail' pasarle lo parametros de lo contrario el preview no funcionara 
    'image_versions' => array(
        'size_800x600' => array(
            'max_width' => 800,
            'max_height' => 600
        ),
        'size_200x150' => array(
            'max_width' => 200,
            'max_height' => 150
        ),
        //Tamaño del thumbnail (Es la imagen que se usa para el previwe del upload)
        'thumbnail' => array(
            'max_width' => 80,
            'max_height' => 80,
            'upload_dir' => '../upload/comprobantes/thumb/',
            'upload_url' => 'upload/comprobantes/thumb/',
            // Uncomment the following to force the max
            // dimensions and e.g. create square thumbnails:
            'jpeg_quality' => 70
            //'no_cache' => true, (there's a caching option, but this remembers thumbnail sizes from a previous action!)
        )
    )
));
