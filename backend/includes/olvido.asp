<div class="wrap-login-box">
 
  <div class="container">
    <div class="row text-center pad-top">
      <div class="col-md-12 col-sm-12">

        <div class="login-box olvido-box">
            
          <div class="login-box-body">
            <p class="login-box-msg" id="titleOlvido">recuperar datos
            <span>Ingrese la dirección de e-mail con la que se registró en nuestro sistema y le reenviaremos sus datos de acceso</span>
            </p>

            <form id="olvidoForm" method="post">
              <div class="form-group form-login">
                <span class="fa fa-envelope"></span>
                <input type="email" class="form-control" name="email" id="email" placeholder="e-mail@ejemplo.com.ar" required data-bv-notempty-message="El e-mail es requerido" data-bv-emailaddress-message="El formato de e-mail no es válido"/>
              </div>
              <div class="row">
                <div class="col-xs-12">
                  <button type="submit" class="btn btn-primary btn-block js--send-login">enviar</button>
                </div><!-- /.col -->
              </div>
              <div id="responseAjax"></div>
            </form>

            <div class="login-footer">
                <div class="row">   
                    <div class="col-xs-12 col-md-6">
                      <a href="registro/" id="loginLinkReg" class="btn-registro">Registrate<strong>como operador</strong></a>
                        <!--<a href="registro/" id="loginLinkReg2" class="btn-registro">Registrate<strong>como hotel</strong></a>-->
                    </div> 
                    <div class="col-xs-12 col-md-6"><a href="login/" class="btn-olvido" id="loginLinkOlv">Login</a></div> 
                </div>
            </div>

          </div><!-- /.login-box-body -->
        </div><!-- /.login-box -->

      </div>
    </div>
  </div>
    
</div>