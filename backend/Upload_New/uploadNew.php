<?php
//////////////////////////////////////////////////
// Upload_New (PHP 7+)                          //
// Server side v 1.0.1                          //
// Dete: 11/09/2020                             //
// Author: Raúl Ghioldi                         //
//////////////////////////////////////////////////

ini_set("display_errors", 1);
//error_reporting(E_ALL & ~E_NOTICE);

require_once 'includes/imageTools.class.php';

require_once './includes/imageResize/ImageResize.php';
require_once './includes/imageResize/ImageResizeException.php';

use \Gumlet\ImageResize; //Espacio en memoria para la clase ImageResize.php 


function is_valid_type($file){
  
  $valid_types = array(
                      "application/vnd.openxmlformats-officedocument.wordprocessingml.document", 
                      "application/vnd.ms-excel", 
                      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", 
                      "application/vnd.ms-powerpoint",
                      "application/vnd.openxmlformats-officedocument.presentationml.presentation",
                      "application/pdf",
                      "application/x-rar-compressed", 
                      "application/octet-stream",
                      "application/zip", 
                      "application/octet-stream", 
                      "application/x-zip-compressed", 
                      "multipart/x-zip",
					  "text/csv"
					  
                    );
  
  return (in_array($file, $valid_types)) ? true : false;

}

//INIT
// @TODO - Handle the $_POST vars if you have any

$isFile = true;
$fileType = false;
$error = "";

//FILE CHECK
if (!isset($_FILES['upfile']['tmp_name'])) {
  $isFile = false;
  $error = "No se subió ningún archivo.";
}

//Si existe archivo
if($isFile) {

  // Detecto la extención de la imagen
  $extImgObj = new ImageTools();
  $extImgObj->ImagenTemp = $_FILES['upfile'];
  $extImgObj->ImagenType = array("image/jpeg", "image/jpg", "image/gif", "image/png");//Estos son los valores por defaul que ya tiene la case 
  $extFileInfo = $extImgObj->infoTipoImagen();

  if($extFileInfo[0] == 'image'){
    $fileType = 'image';
  }else{
    $fileType = 'doc';
    if(is_valid_type($_FILES['upfile']['type'])){

      $extFileInfo[1] = $_FILES['upfile']['type'];
      $extFileInfo[2] = pathinfo($_FILES['upfile']['name'], PATHINFO_EXTENSION);

    }else{

      $error = "El tipo de archivo (" . $_FILES['upfile']['type'] . ") no es soportado!";

    }
  }

}








//print_r($extFileInfo);

$path = isset($_POST['file_path']) ? $_POST['file_path'] : '';

//Seteos
include_once 'config.php';
//


//SERVER RESPONSE
header('Content-Type: application/json');

$responseJson = json_encode([
  "is_file" => $isFile,
  "file_type" => $fileType, 
  "mime_type" => $isFile ? $extFileInfo[1] : "",
  "extention" => $isFile ? $extFileInfo[2] : "",
  "temp_name" => $isFile ? $_FILES["upfile"]["tmp_name"] : "",
  "name" => isset($nombreFile) ? $nombreFile : "",
  "path" => isset($file_path) ? str_replace("../", "", $file_path) : "",
  "record_saved" => isset($option_folder) ? $option_folder . $nombreFile : $nombreFile,
  "preview" => isset($preview) ? $file_path . $preview . $nombreFile : "",
  "preview_large" => isset($previewLarge) ? $file_path . $previewLarge . $nombreFile : "",
  "preview_link" => isset($option_folder) ? $relativePath . $option_folder . $nombreFile : $nombreFile,
  "error" => $error,
  "message" => ($error == "") ? "Upload OK" : $error
]);

echo $responseJson;

?>