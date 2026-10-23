
        <!--NAVBAR SECTION-->
        <div class="navbar navbar-inverse navbar-fixed-top scrollclass">
            <div class="container">
                <div class="navbar-header">
                    <button type="button" class="navbar-toggle" data-toggle="collapse" data-target=".navbar-collapse">
                        <span class="icon-bar"></span>
                        <span class="icon-bar"></span>
                        <span class="icon-bar"></span>
                    </button>
                    <a class="navbar-brand" href="#"><img src="<%Application("Dominio")%>dist/img/logo-hotelypago.png" alt="hotelypago.com"></a>
                </div>
                <div class="navbar-collapse collapse">
                    <ul class="nav navbar-nav navbar-right">
                        <li><a href="<%if pagename <> "home" then%>default.asp<%end if%>#home">HOME</a></li>
                        <li><a href="<%if pagename <> "home" then%>default.asp<%end if%>#service">SERVICIOS</a></li>
                        <li><a href="<%if pagename <> "home" then%>default.asp<%end if%>#costs">COSTOSHEADER</a></li>
                        <li><a href="<%if pagename <> "home" then%>default.asp<%end if%>#about">NOSOTROS</a></li>
                        <li><a href="<%if pagename <> "home" then%>default.asp<%end if%>#contact">CONTACTO</a></li>
                        <li><!--<a href="http://52.4.15.66/hotelypago_desarrollo/login/">LOGIN</a>-->
                            <div class="btn-group rss-select-language">
                                <button type="button" class="btn btn-default dropdown-toggle login-btn" data-toggle="dropdown" aria-haspopup="true" aria-expanded="false">
                                    <i class="svg-language">iniciar sesion</i>
                                    <span class="sr-only">Toggle Dropdown</span>
                                </button>
                                <ul class="dropdown-menu">
                                    <li><a href="<%=Application("Dominio")%>loginH/"><i class="fa fa-user" aria-hidden="true"></i> LoginHoteles</a></li>
                                    <li role="separator" class="divider"></li>
                                    <li><a href="<%=Application("Dominio")%>login/"><i class="fa fa-users" aria-hidden="true"></i> LoginOperadores</a></li>
                                    <li role="separator" class="divider"></li>
                                    <li><a href="<%=Application("Dominio")%>registro/"><i class="fa fa-id-card-o" aria-hidden="true"></i>registrate!</a></li>
                                </ul>
                            </div>
                        </li>
                        <!--<li>
                            <div class="btn-group rss-select-language">
                                <button type="button" class="btn btn-default dropdown-toggle" data-toggle="dropdown" aria-haspopup="true" aria-expanded="false">
                                    <i class="svg-language"><img src="home/svg/flags/es.svg" alt=""></i>
                                    <span class="sr-only">Toggle Dropdown</span>
                                </button>
                                <ul class="dropdown-menu">
                                    <li><a href="http://52.4.15.66/hotelypago_desarrollo/es/" class="current"><i><img src="home/svg/flags/es.svg" alt="Español"></i> Español</a></li>
                                    <li role="separator" class="divider"></li>
                                    <li><a href="https://www.hotelypago.com/en/" class=""><i><img src="home/svg/flags/en.svg" alt="English"></i> English</a></li>
                                    <li role="separator" class="divider"></li>
                                    <li><a href="http://52.4.15.66/hotelypago_desarrollo/po/" class=""><i><img src="home/svg/flags/po.svg" alt="Español"></i> Portugues</a></li>
                                </ul>
                            </div>
                        </li>-->

                    </ul>
                </div>

            </div>
        </div>
        <!--END NAVBAR SECTION-->