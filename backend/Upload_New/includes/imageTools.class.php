<?php
//Tamaño de imagenes
//Los archivos jpg deben ser declarados como "image/jpeg" pues ese es el tipo de extención que se reconoserá
class ImageTools {

// Imagen a examinar
// Array del tipo $imagen['temp_name']	
public $ImagenTemp;

// Tamaño Máximo de la imagen	
public $ImagenPeso = 1024000; //Bits = 1Mb

// Tipos de archivos de imagenes permitidos
public $ImagenType = array("image/jpeg", "image/jpg", "image/gif", "image/png");

public $FileInfo;

// String con con la extención del archivo de imagen
private $ImagenExtencion;

private $ImagenFullInfo;

// Error devuelto
private $Error;



	// Método que devuelve, en caso de exito, la extención de una imagen
	function infoTipoImagen(){

		if(!empty($this->ImagenTemp)){

			if(in_array($this->ImagenTemp['type'], $this->ImagenType)){
				
				$fileValid = 'image';
				$this->ImagenExtencion = explode("image/", $this->ImagenTemp['type']);	
			   
		    }else{

				$fileValid = 'no-image';
				$this->ImagenExtencion = explode("/", $this->ImagenTemp['type']);
				//print_r($this->ImagenExtencion);
			}
			
			$this->FileInfo = array($fileValid, $this->ImagenTemp['type'], $this->ImagenExtencion[1]);	

			return $this->FileInfo;

		}
		
	}
	
	// Metodo para determinar que el peso de una imagen no sobrepase el establecido
	function infoPesoImagen(){
		
		if(!empty($this->ImagenTemp)){
			
		   if($this->ImagenTemp['size'] > $this->ImagenPeso){
			   
			   $this->Error = "La imagen debe pesar como máximo ".$this->ImagenPeso." Kbs <br>";
			   return $this->Error;
			   
		   }
		   
		}
		
		return;
	}

	// Método que devuelve, en caso de exito, toda la información com las medidas de la imagen
	function infoFull(){
	
		if(!empty($this->ImagenTemp)){

			$this->ImagenFullInfo = getimagesize($this->ImagenTemp['tmp_name']);
					
			return $this->ImagenFullInfo;

		}
		
	}
	
}
?>