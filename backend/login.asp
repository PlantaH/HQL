<%
pageName = "login"
%>
<!DOCTYPE html>
<html lang="es">

    <head>
        
		<!-- head -->
        <!-- #include file="includes/head_login.asp" -->

        <link rel="stylesheet" href="plugins/bootstrapValidator/css/bootstrapValidator.css"/>
        
    </head>
	
	 
		
    <body class="login-page">
        
        <!-- header -->
        <!-- #include file="includes/header_login.asp" -->

        <!-- home -->
        <!-- #include file="includes/login.asp" -->

        <!-- footer -->
        <!-- #include file="includes/footer_login.asp" -->
		 
        <!-- jQuery 2.1.3 -->
        <script src="plugins/jQuery/jQuery-2.1.3.min.js"></script>
        <!-- Bootstrap 3.3.2 JS -->
        <script src="bootstrap/js/bootstrap.min.js" type="text/javascript"></script>
        <!-- iCheck -->
        <script src="plugins/iCheck/icheck.min.js" type="text/javascript"></script>
        <!-- Validator -->
        <script src="plugins/bootstrapValidator/js/bootstrapValidator.js" type="text/javascript"></script>
        
        <script src="plugins/main.js"></script>
        
		<script>
			
			 
			if (localStorage.getItem('UsuarioID_Encrypt')) {
				let u = localStorage.getItem('UsuarioID_Encrypt')
				console.log(u)
				//loginAjax.loginAuto(u)
			} 
			 
			  
			$(document).ready(function() {
				
				loginAjax.init('C');
                showPasword();

                //ADD REMOVE PADDING CLASS ON SCROLL
                $(window).scroll(function () {
                    if ($(".navbar").offset().top >40) {
                        $(".navbar-fixed-top").addClass("navbar-pad-original");
                    } else {
                        $(".navbar-fixed-top").removeClass("navbar-pad-original");
                    }
                });
			
				
				
			});
			 
        </script>
		 
    </body>

</html>
