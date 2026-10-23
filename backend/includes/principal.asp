
<!-- Right side column. Contains the navbar and content of the page -->

<div class="content-wrapper"> 
  <!-- Content Header (Page header) -->
  
   
  <section class="content-header">
		<%
		 set rt = conn.execute("select distinct Marca from  Productos where  borrado=0  order by marca  ")
		
		while not rt.eof
			if sqlAnd ="" then sqlAnd = " and Marca='" & rt("marca") & "' "%>
			<button type="button" class="btn btn-primary" onclick="window.location ='CambiarDashboard.asp?marca=<%=rt("marca")%>'">Ver <%=ucase(rt("marca"))%></button>
		 <%rt.movenext
		wend
		 
		%>
		 
	  
  </section>
  
  <%
  if session("Marca_DashBoard") <> "" then sqlAnd =  " and Marca='" &  session("Marca_DashBoard") & "' "
   
  set rt = conn.execute("select distinct Marca from  Productos where  borrado=0 " & sqlAnd & " order by marca")
  %>
   
  
  <!-- Main content -->
  <section class="content table-custome1">
		<%while not rt.eof%>
		<div class="row">
	
		  <div class="col-xs-12">
			<div class="box">
			  <div class="box-body">
				<strong><h3><%=rt("marca")%>&nbsp;<%if session("usuarioid") =1 or session("usuarioid") =2 or session("usuarioid") =3 then%> <a href="RPT_Exportar.asp" target="_blank">Exporta a .csv</a><%end if%></h3> </strong>
			  </div>
			  <!-- /.box-body --> 
			</div>
			<!-- /.box --> 
			
		  </div>
		  <!-- /.col --> 
		</div>
		<!-- /.row --> 
	 
		<div class="row">

			<%'set rs = conn.execute("select * from  Productos where  borrado=0 and marca='" & rt("marca") & "'   order by nombre")
			
			sSql = "select *,"
			sSql = sSql & "(select count(*) as cantidad from Productos_Log 	where codigo=p.codigo and Modo='E' ) as entradas,"
			sSql = sSql & "(select count(*) as cantidad from Productos_Log 	where codigo=p.codigo and Modo='S' ) as salidas "
			sSql = sSql & "from  Productos p "
			sSql = sSql & "where  p.borrado=0 and p.marca='" & rt("marca") & "'   "
			sSql = sSql & "order by p.nombre"
			set rs = conn.execute(sSql)
			while not rs.eof
			
				entradas = rs("entradas")
				salidas = rs("salidas")
			
			'entradas = 0
			'set rCa = conn.execute("select count(*) as cantidad from Productos_Log 	where codigo='" & rs("codigo") & "' and Modo='E'  ")
			'if not rCa.eof then entradas = rCa("cantidad")
			'salidas = 0
			'set rCa = conn.execute("select   count(*) as cantidad from Productos_Log 	where codigo='" & rs("codigo") & "' and Modo='S'  ")
			'if not rCa.eof then salidas = rCa("cantidad")
			'set rCa = nothing
			cantidad = entradas-salidas
			'rs("cantidad")
			if cantidad<5  then esteColor="#F7E052"
			if cantidad>=5 then esteColor="#8EF265"
			if cantidad<=0 then esteColor="#F27465"
			
			if cantidad<>0 then%>
			<div class="col-md-3 col-sm-6 col-xs-12" >
				<div class="info-box" <%if cantidad>0 then%> onclick="AbrirProd('<%=rs("codigo")%>')" style="cursor:pointer" <%end if%>>
					<span class="info-box-icon" style="background-color:<%=esteColor%>"><%=cantidad%></span>
					<div class="info-box-content">
						<h5><strong><%=rs("nombre")%>&nbsp;</strong></h5>
						 <div class="info-box-footer">
							<span class="info-box-text"><%=rs("codigo")%>&nbsp;</span> 
						</div>
					</div>
				</div>
			</div>
			<%end if
			rs.movenext
			wend%>
			  
		</div>
	<%rt.movenext
	wend%>
	 

	
  </section>
  <!-- /.content -->
    
   
</div>
<!-- /.content-wrapper --> 