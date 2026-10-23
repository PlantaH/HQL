<%
Permisos(2)

if PER_Alta then url_nuevo = "productos-nuevo/"
%>
<!-- Right side column. Contains the navbar and content of the page -->


<div class="content-wrapper"> 
  <!-- Content Header (Page header) -->
  <section class="content-header">
    <h1> Productos
      <!--<small>advanced tables</small>--> 
    </h1>	
    <!-- #include file="Botones.asp" -->
  </section>
  
  <!-- Main content -->
  <section class="content table-custome1">
	<div class="row">
	  <div class="col-xs-12">
		<div class="box">
		  <div class="box-body">
			<form name="frm-example" id="frm-example">

				<table id="tablaDatos"  class="table table-striped table-bordered" style="width:100%; cursor:pointer"  >
					<thead>
						<tr>
						<th>ID</th>
						<th>Marca</th>
						<th>Codigo Venta</th>
						<th>Codigo Barras</th>						
						<th>Nombre</th>
						<th>Cantidad</th>
						<th></th> 
						<th></th> 						
						</tr>
					</thead>                  
				</table>

				<div class="content-btns form-group">
					<button type="submit" class="btn btn-primary"><i class="fa fa-trash-o" aria-hidden="true"></i> Eliminar seleccionados</button>
				</div>

			</form>
		  </div>
		  <!-- /.box-body --> 
		</div>
		<!-- /.box --> 
		
	  </div>
	  <!-- /.col --> 
	</div>
	<!-- /.row --> 
  </section>
  <!-- /.content -->
   
</div>
<!-- /.content-wrapper --> 