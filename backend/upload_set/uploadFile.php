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

$GLOBALS['set_upload_prefix'] = $_GET['type'] . '_';

//Extiende la clase original
//Customiza el nombre del archivo
class CustomUploadHandler extends UploadHandler {
    protected function trim_file_name($file_path, $name, $size, $type, $error, $index, $content_range) {
        $name = $GLOBALS['set_upload_prefix'] . microtime(true);
        $name = str_replace('.', '', $name);
        return $name;
    }
}

//Extiende la clase original
//Customiza el nombre del archivo y permite también subir PDF y otro tipos de archivos
class CustomFilesUploadHandler extends UploadHandler {
    protected function trim_file_name($file_path, $name, $size, $type, $error, $index, $content_range) {
        
        $extension = explode(".", strtolower($name));
        $name = $GLOBALS['set_upload_prefix'] . microtime(true);
        $name = str_replace('.', '', $name);
        $name = $name . '.' . end($extension);

        return $name;
    }
}
//

//Seteo por parametros
//Switch que setea las caracteristicas de la imagen segÚn el valor de la variable enviada por get $_GET['type']
switch ($_GET['type']) {
    
   
    //Comprobantes de pago
    //Pueden ser imagenes o PDF, por esta razón no se manimpula el tamaño de las imatenes.
    //Cuando es imagen se genera una miniatura de forma automatica para el preview)
    case 'marcas2':
        $set_upload_dir               = '../upload/' . $_GET['type'] . '/';
        $set_upload_url               = 'upload/' . $_GET['type'] . '/';
        $set_accept_file_types        = '/\.(jpe?g|png|gif|pdf|pptx|pptm|ppt|docx|docm|dotx|dotm|doc|dot|xls|xlt|xlsx|xlsm)$/i';

        /*INICIA EL PROCESO DE UPLOAD */
        $upload_handler = new CustomFilesUploadHandler(array(
            //Path donde se guarda el alchivo (relativo a este alchivo)
            'upload_dir' => $set_upload_dir, 
            //Path del preview (relativo a la pagina de upload)
            'upload_url' => $set_upload_url
        ));

        break;
    //
	
	//Ejemplo de upload COMPLETO
    case 'categorias':
        $set_upload_dir               = '../upload/' . $_GET['type'] . '/';
        $set_upload_url               = 'upload/' . $_GET['type'] . '/';
        $set_accept_file_types        = '/\.(jpe?g|png|gif)$/i';

        //Versionados del archivo (Se usa cuando subimos imagenes y queremos guardar diferentes tamaños)
        //Si se define se debe definir además el 'thumbnail' pasarle lo parametros de lo contrario el preview no funcionara 
        $set_image_versions = array(
            'size_300x300' => array(
                'max_width' => 300,
                'max_height' => 300,
                'jpeg_quality' => 100
            ),
			'size_100x100' => array(
                'max_width' => 100,
                'max_height' => 100,
                'jpeg_quality' => 100
            ),
            //Tamaño del thumbnail (Es la imagen que se usa para el previwe del upload)
            'thumbnail' => array(
                'max_width' => 80,
                'max_height' => 80,
                'upload_dir' => '../upload/' . $_GET['type'] . '/thumb/',
                'upload_url' => 'upload/' . $_GET['type'] . '/thumb/',
                // Uncomment the following to force the max
                // dimensions and e.g. create square thumbnails:
                'jpeg_quality' => 70
                //'no_cache' => true, (there's a caching option, but this remembers thumbnail sizes from a previous action!)
            )
        );

        /*INICIA EL PROCESO DE UPLOAD */
        //Instancia origianl
        //$upload_handler = new UploadHandler();

        //Intancia extendida y customizada
        $upload_handler = new CustomUploadHandler(array(
            //Tamaño máximo del archivo por default es libre
            //'max_file_size' => 1048576, //1MB file size
            //Path donde se guarda el alchivo (relativo a este alchivo)
            'upload_dir' => $set_upload_dir, 
            //Path del preview (relativo a la pagina de upload)
            //Si nose define 'imagen_versions' la miniatura se guarda en el directorio "thumbnail" (si no existe el directorio, es creado en el momento)
            'upload_url' => $set_upload_url,
            //'script_url' => '/url/of/uploadhandler/script',
            'accept_file_types' => $set_accept_file_types,

            //Versionados del archivo (Se usa cuando subimos imagenes y queremos guardar diferentes tamaños)
            //Si se define se debe definir además el 'thumbnail' pasarle lo parametros de lo contrario el preview no funcionara 
            'image_versions' => $set_image_versions
        ));
        break;
    //
	
	
	 //Ejemplo de upload COMPLETO
    case 'bannershome':
        $set_upload_dir               = '../../upload/' . $_GET['type'] . '/';
        $set_upload_url               = '../upload/' . $_GET['type'] . '/';
        $set_accept_file_types        = '/\.(jpe?g|png|gif)$/i';

        //Versionados del archivo (Se usa cuando subimos imagenes y queremos guardar diferentes tamaños)
        //Si se define se debe definir además el 'thumbnail' pasarle lo parametros de lo contrario el preview no funcionara 
        $set_image_versions = array(
            'size_1920x336' => array(
                'max_width' => 1920,
                'max_height' => 336,
                'jpeg_quality' => 100
            ),
            'size_50x50' => array(
                'max_width' => 50,
                'max_height' => 50,
                'jpeg_quality' => 100
            ),
            //Tamaño del thumbnail (Es la imagen que se usa para el previwe del upload)
            'thumbnail' => array(
                'max_width' => 80,
                'max_height' => 80,
                'upload_dir' => '../../upload/' . $_GET['type'] . '/thumb/',
                'upload_url' => '../upload/' . $_GET['type'] . '/thumb/',
                // Uncomment the following to force the max
                // dimensions and e.g. create square thumbnails:
                'jpeg_quality' => 70
                //'no_cache' => true, (there's a caching option, but this remembers thumbnail sizes from a previous action!)
            )
        );

        /*INICIA EL PROCESO DE UPLOAD */
        //Instancia origianl
        //$upload_handler = new UploadHandler();

        //Intancia extendida y customizada
        $upload_handler = new CustomUploadHandler(array(
            //Tamaño máximo del archivo por default es libre
            //'max_file_size' => 1048576, //1MB file size
            //Path donde se guarda el alchivo (relativo a este alchivo)
            'upload_dir' => $set_upload_dir, 
            //Path del preview (relativo a la pagina de upload)
            //Si nose define 'imagen_versions' la miniatura se guarda en el directorio "thumbnail" (si no existe el directorio, es creado en el momento)
            'upload_url' => $set_upload_url,
            //'script_url' => '/url/of/uploadhandler/script',
            'accept_file_types' => $set_accept_file_types,

            //Restricciones de tamaño para las fotos, tomar en cuenta el tamaño de la imagen más grande
            'max_width' => 1920,
            'max_height' => 336,
            'min_width' => 1920,
            'min_height' => 336,
            //

            //Versionados del archivo (Se usa cuando subimos imagenes y queremos guardar diferentes tamaños)
            //Si se define se debe definir además el 'thumbnail' pasarle lo parametros de lo contrario el preview no funcionara 
            'image_versions' => $set_image_versions
        ));
        break;
    //
	
	

	//Ejemplo de upload COMPLETO
    case 'destinosdestacados':
        $set_upload_dir               = '../../upload/' . $_GET['type'] . '/';
        $set_upload_url               = '../upload/' . $_GET['type'] . '/';
        $set_accept_file_types        = '/\.(jpe?g|png|gif)$/i';

        //Versionados del archivo (Se usa cuando subimos imagenes y queremos guardar diferentes tamaños)
        //Si se define se debe definir además el 'thumbnail' pasarle lo parametros de lo contrario el preview no funcionara 
        $set_image_versions = array(
            'size_1100x1100' => array(
                'max_width' => 1100,
                'max_height' => 1100,
                'jpeg_quality' => 100
            ),
            'size_50x50' => array(
                'max_width' => 50,
                'max_height' => 50,
                'jpeg_quality' => 100
            ),
            //Tamaño del thumbnail (Es la imagen que se usa para el previwe del upload)
            'thumbnail' => array(
                'max_width' => 80,
                'max_height' => 80,
                'upload_dir' => '../../upload/' . $_GET['type'] . '/thumb/',
                'upload_url' => '../upload/' . $_GET['type'] . '/thumb/',
                // Uncomment the following to force the max
                // dimensions and e.g. create square thumbnails:
                'jpeg_quality' => 70
                //'no_cache' => true, (there's a caching option, but this remembers thumbnail sizes from a previous action!)
            )
        );

        /*INICIA EL PROCESO DE UPLOAD */
        //Instancia origianl
        //$upload_handler = new UploadHandler();

        //Intancia extendida y customizada
        $upload_handler = new CustomUploadHandler(array(
            //Tamaño máximo del archivo por default es libre
            //'max_file_size' => 1048576, //1MB file size
            //Path donde se guarda el alchivo (relativo a este alchivo)
            'upload_dir' => $set_upload_dir, 
            //Path del preview (relativo a la pagina de upload)
            //Si nose define 'imagen_versions' la miniatura se guarda en el directorio "thumbnail" (si no existe el directorio, es creado en el momento)
            'upload_url' => $set_upload_url,
            //'script_url' => '/url/of/uploadhandler/script',
            'accept_file_types' => $set_accept_file_types,

            //Restricciones de tamaño para las fotos, tomar en cuenta el tamaño de la imagen más grande
            'max_width' => 1102,
            'max_height' => 1102,
            'min_width' => 1000,
            'min_height' => 1000,
            //

            //Versionados del archivo (Se usa cuando subimos imagenes y queremos guardar diferentes tamaños)
            //Si se define se debe definir además el 'thumbnail' pasarle lo parametros de lo contrario el preview no funcionara 
            'image_versions' => $set_image_versions
        ));
        break;
    //
}
/**/
?>