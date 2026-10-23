<div class="wrap-login-box">

  <div class="container">
    <div class="row text-center pad-top">
      <!--
      <div class="col-md-12 col-sm-12">
        <a class="navbar-brand-login" href="#">
          <img src="<%=Application("Dominio")%>dist/img/logo.gif" alt="<%=Application("Dominio")%>">
        </a>
      </div>
		-->
      <div class="col-md-12 col-sm-12">

        <div class="login-box">

          <div class="login-box-body">
            <p class="login-box-msg" id="userName">BACKEND</p>
            
            <form id="loginForm" method="post">
              <div class="form-group form-login">
                <span class="fa fa-envelope"></span>
                <input type="text" class="form-control" name="email" id="email" placeholder="mail@mail.com"/>
              </div>
              <div class="form-group form-login show-password">
                <span class="fa fa-unlock-alt"></span>
                <input type="password" class="form-control" name="password" id="password" placeholder="contrasena"/>
                <span class="show-password-btn fa fa-eye" title="Mostrar contrasena" data-title-on="Mostrar contrasena" data-title-off="Ocultar contrasena"></span>
              </div>
              <div class="row">
                <div class="col-xs-12">
                  <button type="submit" class="btn btn-primary btn-block js--send-login">entrar</button>
                </div><!-- /.col -->
              </div>
              <div id="responseAjax"></div>
            </form>

            <div class="login-footer">
                <div class="row">   
                    <div class="col-xs-12 col-md-12"> </div> 
                </div>
            </div>
        
          <div>
            <div class="row">   
                                  
                </div>
            </div>

          </div><!-- /.login-box-body -->
        </div><!-- /.login-box -->
    
      </div>

      
    </div>
  </div>

</div>