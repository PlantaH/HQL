<?php
ini_set("display_errors", 1);
//error_reporting(E_ALL & ~E_NOTICE);

$relativePath = '';
$relativeAdminPath = '../';

switch($path) {
	
	

	case 1:

		//ARCHIVO DE IMAGEN $fileType == 'image'
		//Si es Imágen
		if($fileType == 'image'){

		  $folder = "upload/Paquetes/Fotos/";
		  $file_path = $relativePath . $folder;
		  $file_path_admin = $relativeAdminPath . $folder;
		  
		  $preview = "Previo/";
		  $previewLarge = "Grande/";//Es el subdirectorio para el modal de visualización , es relaltivo a $folder
		  $prefix = "foto_";
		  $subfolder = "Grande/";
		  
		  $option_folder = $folder ;
		  
		  // Guardo la imagen en el directorio especificado
		  //$nombreFile = $prefix . md5(time()).".".$extFileInfo[2];
			$nombreFile = $prefix . date("Y-m-d-H-i-s") .".".$extFileInfo[2];
		  chmod($_FILES['upfile']['tmp_name'], 0777); 
		  move_uploaded_file($_FILES['upfile']['tmp_name'], $file_path_admin . $subfolder . $nombreFile);
		
			
		  //Redimensiona la imágen original;
		  $image = new \Gumlet\ImageResize($file_path_admin . $subfolder . $nombreFile);
		  $image->quality_jpg = 100;
		  $image->resizeToWidth(480);
		  $image->save($file_path_admin . "Destacada/" .  $nombreFile);
		  //
		  
		  //Redimensiona la imágen original;
		  $image = new \Gumlet\ImageResize($file_path_admin . $subfolder . $nombreFile);
		  $image->quality_jpg = 100;
		  $image->resize(120, 80);
		  $image->save($file_path_admin . "Mediana/" .  $nombreFile);
		  //
		 

		  //Resize obligatorio para preview;
		  $image = new \Gumlet\ImageResize($file_path_admin . $subfolder . $nombreFile);
		  $image->quality_jpg = 80;
		  $image->resizeToWidth(200);
		  $image->save($file_path_admin . $preview . $nombreFile);
		  //

		}else{

		  $error = $extFileInfo[2] . " es un tipo de archivo no permitido - " . $_FILES["upfile"]["name"];

		}

    break;
	
	

  /*
  case 234:

    //ARCHIVO DE IMAGEN $fileType == 'image'
    //Si es Imágen
    if($fileType == 'image'){
      
      $folder = "Uploads/temp-files2/";
      $file_path = $relativePath . $folder;
      $file_path_admin = $relativeAdminPath . $folder;
      $option_folder = $folder;
      $preview = "previo/";
      $prefix = "banner_";
      
      // Guardo la imagen en el directorio especificado
      //$nombreFile = $prefix . date("Y-m-d-H-i-s") .".".$extFileInfo[2]; //ej: banner_2020-03-10-17-16-18.jpeg
      $nombreFile = $prefix . md5(time()).".".$extFileInfo[2];
      chmod($_FILES['upfile']['tmp_name'], 0777); 
      move_uploaded_file($_FILES['upfile']['tmp_name'], $file_path_admin . $nombreFile);

      //Resize obligatorio para preview;
      $image = new \Gumlet\ImageResize($file_path_admin . $nombreFile);
      $image->quality_jpg = 80;
      $image->resizeToWidth(200);
      $image->save($file_path_admin . $preview . $nombreFile);
      //

    }else{

      $error = $extFileInfo[2] . " es un tipo de archivo no permitido - " . $_FILES["upfile"]["name"];

    }

    break;

*/
  case 235:

    //ARCHIVO DE NO-IMAGEN $fileType == 'doc'
    //
    if($fileType == 'doc'){

      $folder = "Uploads/rpt/Archivos/";
      $file_path = $relativePath . $folder;
      $file_path_admin = $relativeAdminPath . $folder;
      $option_folder = $folder;
      $preview = "";
      $prefix = "";

      $nombreFile = $prefix . md5(time()).".".$extFileInfo[2];
      chmod($_FILES['upfile']['tmp_name'], 0777); 
      move_uploaded_file($_FILES['upfile']['tmp_name'], $file_path_admin . $nombreFile);

    }else{
      
      $error = $extFileInfo[2] . " es un tipo de archivo no permitido - " . $_FILES["upfile"]["name"];

    }

    break;
	
  //En caso que el PATH no exista
	default;

	$nombreFile = "";  
	$error = "El PATH de destino no es válido!";

break;     

}
?>