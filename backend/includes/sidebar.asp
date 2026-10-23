      <!-- Left side column. contains the logo and sidebar -->
      <aside class="main-sidebar">
        <!-- sidebar: style can be found in sidebar.less -->
        <section class="sidebar">

          <!-- Sidebar user panel -->
          <div class="user-panel">
            <div class="image">
               <i class="fa fa-user"></i>
            </div>
            <div class="info">
              <p><%=session("Cliente_Nombre")%></p>
            </div>
          </div>

          <ul class="sidebar-menu user-menu">
            <li class="treeview<%if pageGroup = "micuenta"  then%> active<%end if%>">
              <a href="#" class="user-data">
                <span>Mi cuenta</span> <i class="fa fa-sort-desc pull-right"></i>
              </a>
              
                <ul class="treeview-menu user-data-ul">
                  <li<%if pageGroup = "micuenta" and pageName = "Misdatos"  then%> class="active"<%end if%>><a href="mis-datos/"><i class="fa <%if pageGroup = "micuenta" and pageName = "Misdatos"  then%>fa-circle<%else%>fa-circle-o<%end if%>"></i> Mis datos</a></li>                
                </ul>
              
            </li>
          </ul>
          <!-- search form -->
          <!--<form action="#" method="get" class="sidebar-form">
            <div class="input-group">
              <input type="text" name="q" class="form-control" placeholder="Search..."/>
              <span class="input-group-btn">
                <button type='submit' name='search' id='search-btn' class="btn btn-flat"><i class="fa fa-search"></i></button>
              </span>
            </div>
          </form>-->
          <!-- /.search form -->
          <!-- sidebar menu: : style can be found in sidebar.less -->
          <ul class="sidebar-menu">	
			<li class="treeview<%if pageGroup = "principal" and pageName = "principal" then%> active<%end if%>">
                    <a href="principal/">
                      <i class="fa fa-bar-chart"></i>
                      <span>Panel de Control</span>
                    </a>
                  </li>
            
              <%
              set Menu=conn.execute("select distinct me.menuID,me.nombre,me.orden,me.iconoMenu from Permisos pe inner join paginas pa on pa.paginaID=pe.PaginaID inner join menu me on me.menuID=pa.menuID where  pe.UsuarioID=" & session("UsuarioID") & " and me.borrado=0 and pa.mostrar=1 and (pe.Alta=1 or pe.Baja=1 or pe.Modi=1 or pe.Consulta=1) order by me.orden")
              while not Menu.eof%>
                <li class="treeview<%if Menu("Nombre") = pageGroup then%> active<%end if%>">
                  <a href="#">
                  <i class="<%=Menu("iconoMenu")%>"></i>
                  <span><%=(Menu("Nombre"))%></span>
                  <i class="fa fa-angle-left pull-right"></i>
                  </a>
                  <ul class="treeview-menu">
                  <%set Items=conn.Execute("select * from paginas pa inner join Permisos Pe on Pe.PaginaID=Pa.PaginaID where pa.MenuID=" & Menu("MenuID") & " and pa.mostrar=1 and UsuarioID=" & session("UsuarioID") & " and Consulta=1   order by Orden")
                  while not Items.eof%>
                    <li <%if Menu("Nombre") = pageGroup and pageName = replace(Items("Nombre"), " ", "")  then%> class="active" <%end if%>><a href="<%=Items("LinkAmigable")%>/"><i class="fa <%if Menu("Nombre") = pageGroup and pageName = replace(Items("Nombre"), " ", "")  then%> fa-circle<%else%>fa-circle-o<%end if%>"></i><%=(Items("Nombre"))%></a></li>
                    <%Items.movenext
                  wend%>
                  </ul>
                </li>
              <%Menu.movenext
              wend%>	
              
           
			 
          </ul>	
		
        </section>
        <!-- /.sidebar -->
      </aside>