// Login
var loginAjax = {
	
	options : {
		
		url_login : 'login_ajax.asp',	
		tipo_login: 'C'
				
	},
	
	init : function(tipoLogin){
        
        if(tipoLogin == 'H'){
			
			loginAjax.options.url_login = 'login_ajaxH.asp';	
			loginAjax.options.tipo_login = 'H';
		}
		
		loginAjax.validator();

    },
    
    validator : function(){

        $('#loginForm').bootstrapValidator({
            message: 'Este valor no es válido',
            feedbackIcons: {
                valid: 'glyphicon glyphicon-ok',
                invalid: 'glyphicon glyphicon-remove',
                validating: 'glyphicon glyphicon-refresh'
            },
            fields: {
                email: {
                    validators: {
                        notEmpty: {
                            message: 'Debe ingresar una dirección de e-mail'
                        },
                        emailAddress: {
                            message: 'El formato del e-mail no es válido'
                        }
                    }
                },
                password: {
                    validators: {
                        notEmpty: {
                            message: 'Debe ingresar la contraseña'
                        }
                    }
                }
            }
        }).on('success.form.bv', function(e) {
            // Prevent form submission
            e.preventDefault();
        
            // Get the form instance
            var $form = $(e.target);
        
            // Get the BootstrapValidator instance
            var bv = $form.data('bootstrapValidator');

            loginAjax.login($form.serialize());

        });
        //


    },
	
	login : function(data){
	
		$.ajax({
	
			data:  data,
			url:   loginAjax.options.url_login,
			type:  'post',
			dataType : 'html',
			beforeSend: function () {
	
				$("#responseAjax").fadeIn(400).html('<div id="preloader"><div id="loader"></div></div>');
	
			},
	
			success:  function(response) {
		
					if(response == 'error'){
						
						$("#responseAjax").fadeIn(400).html("<p class='errorAjax'>Datos de acceso incorrectos!</p><a href='#' class='js--close-response'><span>vuelva a intentarlo</span></a>");
						
					}else{
	
						localStorage.setItem('UsuarioID_Encrypt', response)
						
						$("#responseAjax").html("<h3>Validando usuario...</h3>");
						 
						setTimeout(function(){$(location).attr('href','principal/');}, 1500);
	
					}
	
			}
	
		});	
	
	},
	
	loginAuto : function(data){
	
		$.ajax({
	
			data:  'id=' + data,
			url:   'login_ajax_Auto.asp',
			type:  'post',
			dataType : 'html',
			beforeSend: function () {
	
				$("#responseAjax").fadeIn(0).html('<div id="preloader"><div id="loader"></div></div>');
	
			},
	
			success:  function(response) {
		
					if(response == 'error'){
						
						$("#responseAjax").fadeIn(10).html("<p class='errorAjax'>Datos de acceso incorrectos!</p><a href='#' class='js--close-response'><span>vuelva a intentarlo</span></a>");
						
					}else{
	
						$(location).attr('href','principal/')
	
					}
	
			}
	
		});	
	
	}

}
//
let callLogOut = () =>{
	localStorage.removeItem('UsuarioID_Encrypt');
	$(location).attr('href','logout.asp')
	
} 

//Registro
var registroAjax = {
	
	options : {
		
		url_registro : 'registro_ajax.asp'	
				
	},

	init : function(tipoLogin){
		
		if(tipoLogin == 'H'){
			
			registroAjax.options.url_registro = 'registro_ajaxH.asp';			
		}
        
        registroAjax.validator();

    },
    
    validator : function(){
		
		$('#registroForm').bootstrapValidator().on('success.form.bv', function(e) {
            // Prevent form submission
            e.preventDefault();
        
            // Get the form instance
            var $form = $(e.target);
        
            // Get the BootstrapValidator instance
            var bv = $form.data('bootstrapValidator');

            registroAjax.registro($form.serialize());

        });
		
	},
			
	registro : function(data){
	
		$.ajax({
	
			data:  data,
			url:   registroAjax.options.url_registro,
			type:  'post',
			dataType : 'html',
			beforeSend: function () {
	
				$("#responseAjax").fadeIn(400).html('<div id="preloader"><div id="loader"></div></div>');
	
			},
	
			success:  function(response) {
	
				//console.log(response);
				if(response == 'ok'){

					$('#registroForm').html("");
					$("#responseAjax").html("<h1>¡Solicitud de registro enviada con éxito!</h1><h3>Le enviaremos un mail para confirmarle la habilitación de su cuenta <span>(Esto puede tardar hasta 24hs.)</span></h3>");

					setTimeout(function(){
						$("#responseAjax").append('<i class="close-registro fa fa-times"></i>').click(function(){
							$(location).attr('href','./');
						});
					}, 1000);

				}else{

					$("#responseAjax").fadeIn(400).html("<p class='errorAjax'>Error en el registro!</p><a href='#' class='js--close-response'><span>vuelva a intentarlo</span></a>");

				}
	
			}
	
		});	
	
	}	

};
//

//Olvido
var olvidoAjax = {
	
	options : {
		
		url_olvido : 'olvido_ajax.asp'	
				
	},

	init : function(tipoLogin){
		
		if(tipoLogin == 'H'){
			
			olvidoAjax.options.url_olvido = 'olvido_ajaxH.asp';			
		}
        
        olvidoAjax.validator();

    },
    
    validator : function(){
		
		$('#olvidoForm').bootstrapValidator().on('success.form.bv', function(e) {
            // Prevent form submission
            e.preventDefault();
        
            // Get the form instance
            var $form = $(e.target);
        
            // Get the BootstrapValidator instance
            var bv = $form.data('bootstrapValidator');

            olvidoAjax.olvido($form.serialize());

        });
		
	},
			
	olvido : function(data){
	
		$.ajax({
	
			data:  data,
			url:   olvidoAjax.options.url_olvido,
			type:  'post',
			dataType : 'html',
			beforeSend: function () {
	
				$("#responseAjax").fadeIn(400).html('<div id="preloader"><div id="loader"></div></div>');
	
			},
	
			success:  function(response) {
	
				//console.log(response);
				if(response == 'ok'){

					$('#olvidoForm, .login-footer').html("");
					$("#titleOlvido").html("¡Datos enviados!");
					$("#titleOlvido").append('<span>(Si no recibe el mail en unos minutos revise la carpeta de "spam", o vuelva a intentarlo)</span>');
					
					setTimeout(function(){
						$(".login-box-body").append('<i class="close-registro fa fa-times"></i>').click(function(){
							$(location).attr('href','./');
						});
					}, 1000);

				}else{

					$("#responseAjax").fadeIn(400).html("<p class='errorAjax'>Error recuperando los datos!</p><a href='#' class='js--close-response'><span>vuelva a intentarlo</span></a>");

				}
	
			}
	
		});	
	
	}	

};

//
var showPasword = function(){
	
	$('.show-password .show-password-btn').click(function(ev){

		var titleOn  = $(this).attr('data-title-on'),
			titleOff = $(this).attr('data-title-off');
			
		if($(this).parent('.show-password').children('input').attr('type') == 'password'){

			$(this).parent('.show-password').children('input').attr('type', 'text')
			$(this).removeClass('fa-eye').addClass('fa-eye-slash').attr('title', titleOff);
			
		}else{
			
			$(this).parent('.show-password').children('input').attr('type', 'password')
			$(this).removeClass('fa-eye-slash').addClass('fa-eye').attr('title', titleOn);
			
		}
		
	});
	
}
//

//Saldo
var saldoAjax = {

	saldo : function(data){
	
		$.ajax({
	
			data:  data,
			url:   'saldo_ajax.asp',
			type:  'post',
			dataType : 'html',			
			beforeSend: function () {
	
				$("#saldo-ajax").fadeIn(400).html('<i class="loading-circle"></i>');
	
			},
			success:  function(response) {
	
				//console.log(response);
				$('#saldo-ajax').html(response);			
	
			}
	
		});	
	
	}	

};

//

//Saldo
var saldoAjaxH = {

	saldo : function(data){
	
		$.ajax({
	
			data:  data,
			url:   'saldo_ajaxH.asp',
			type:  'post',
			dataType : 'html',			
			beforeSend: function () {
	
				$("#saldo-ajax").fadeIn(400).html('<i class="loading-circle"></i>');
	
			},
			success:  function(response) {
	
				/*console.log(response);*/
				$('#saldo-ajax').html(response);				
	
			}
	
		});	
	
	}	

};

//
//Espera el objeto por parametro {url:'archivo.asp', nodeReload:'nombreDelNodo'}
//Opcional select a borrar (reset: Array)
//Ej: {url:'ASP/PRO_Categorias_XHR_new.asp?id='+this.value, nodeReload:'CategoriaID', reset:['SubCategoriaID']}
var selectReloadEstablecimiento= function(obj){

	$.ajax({
	
		url:   obj.url,
		type:  'post',
		dataType : 'html',			
		beforeSend: function () {
	
			$('#' + obj.nodeReload).fadeTo("slow", 0.30);
	
		},
		success: function(response) {

			$('#' + obj.nodeReload).fadeTo("slow", 1).html(response);
			//Option: Reset other selects if request array obj.reset is seted
			if(obj.reset){
				for(index in obj.reset) {
					$('#' + obj.reset[index]).html('');
				}
			}
			//				
	
		}
	
	});	

}
var textReloadEstablecimiento= function(obj){

	$.ajax({
	
		url:   obj.url,
		type:  'post',
		dataType : 'html',			
		beforeSend: function () {
	  
			$('#' + obj.nodeReload).fadeTo("slow", 0.30);
			 
		},
		success: function(response) {
			  
			$('#' + obj.nodeReload).fadeTo("slow", 1).val(response);
						
	
		}
	
	});	

}

var BuscarManifiesto= function(obj){
	$.ajax({
	
		url:   obj.url,
		type:  'post',
		dataType : 'html',			
		beforeSend: function () {
	
			$('#' + obj.nodeReload).fadeTo("slow", 0.30);
	
		},
		success: function(response) {

			$('#' + obj.nodeReload).val(response);
			 			
	
		}
	
	});	

}
//

var selectReload = function(obj){

	$.ajax({
	
		url:   obj.url,
		type:  'post',
		dataType : 'html',			
		beforeSend: function () {
	
			$('#' + obj.nodeReload).fadeTo("slow", 0.30);
	
		},
		success: function(response) {

			$('#' + obj.nodeReload).fadeTo("slow", 1).html(response);
			//Option: Reset other selects if request array obj.reset is seted
			if(obj.reset){
				for(index in obj.reset) {
					$('#' + obj.reset[index]).html('');
				}
			}
			//				
	
		}
	
	});	

}
//


//Only numbers
function isNumberKey(evt){
	var charCode = (evt.which) ? evt.which : evt.keyCode;
	if (charCode != 46 && charCode > 32 
	&& (charCode < 48 || charCode > 57))
	return false;
	return true;
} 
//

//Only numbers and only one dot (.)
function isNumberAndDotKey(node, evt){
	var charCode = (evt.which) ? evt.which : evt.keyCode,
		thisNode = node,
		thisNodeVal = thisNode.value,
		dotExist = thisNode.value.indexOf('.') !== -1 ? 'yes' : 'not';
		//console.log(charCode);
	if ((charCode != 46 || (charCode == 46 && dotExist == 'yes')
	&& charCode > 32)
	&& (charCode < 48 || charCode > 57)
	&& (charCode != 8))
	return false;
	return true;
} 
//

//Only CUIT
function isNumberCuit(evt, ele){

	if(evt.type == 'paste'){
		let clipboardData = evt.clipboardData || window.clipboardData,
		pastedData = clipboardData.getData('Text');
		//console.log(pastedData);
		if(!validateCuit(pastedData)){
			vanillaTriggerEvent(ele, 'input');
			return false;
		} 
	}

	var charCode = (evt.which) ? evt.which : evt.keyCode;

	if(ele.value.length == 2 || ele.value.length == 11){

		ele.value = ele.value + '-'

	}

	if (charCode > 32 && (charCode < 48 || charCode > 57) || ele.value.length > 12)
	return false;
	return true;

} 
//

//Validate CUIT format number
function validateCuit(val){
  	var format = /^\(?([0-9]{2})\)?[-]?([0-9]{8})[-]?([0-9]{1})$/;
	if(val.match(format)){
		return true;
	}else{
		return false;
	}
}
//

//Vanilla Trigger an event
function vanillaTriggerEvent(el, type){
	if ('createEvent' in document) {
		 // modern browsers, IE9+
		 var e = document.createEvent('HTMLEvents');
		 e.initEvent(type, false, true);
		 el.dispatchEvent(e);
	 } else {
		 // IE 8
		 var e = document.createEventObject();
		 e.eventType = type;
		 el.fireEvent('on'+e.eventType, e);
	 }
 }
 //

$(function(){
	$('[data-toggle="tooltip"]').tooltip();
});

$(document).ready(function() {
	
	/*loginAjax.init();
	showPasword();*/
	

});


//Pantalla: Cambios de orientación
window.addEventListener('orientationchange', function() {
	
	
});


//Pantalla: Cambio de tamaño
$(window).resize(function() {


});


//Pantalla: Scroll
$(window).scroll(function() {

	
});
