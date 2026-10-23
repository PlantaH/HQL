<%
pageName = "autocomplete"
%>
<!DOCTYPE html>
<html lang="es">

    <head>
        
		<%''Común al sitema%>
        <!-- #include file="../includes/head.asp" -->

        <%''Inicia personalizado CSS de la página%>
        <!-- DATA TABLES -->
        <link href="plugins/datatables/dataTables.bootstrap.css" rel="stylesheet" type="text/css" />
        <link href="plugins/autocomplete/bootstrap-typeahead/bootstrap-typeahead.css" rel="stylesheet" type="text/css" />
		<%''Fin personalizado CSS de la página%>
        
		<%''Files (CSS y js IE<9) que personalizan el sitema%>
        <!-- #include file="../includes/head_template_custome.asp" -->
        
    </head>

    <body class="skin-blue">

        <!-- header -->
        <!-- #include file="../includes/header.asp" -->
        
        <!-- sidebar -->
        <!-- #include file="../includes/sidebar.asp" -->

        <!-- modal -->
        <div class="content-wrapper"> 
            <!-- Content Header (Page header) -->
            <section class="content-header">
                <h1>Autocomplete 2 </h1>
            </section>

            <!-- Main content -->
            <section class="content table-custome1">
                <h1>Autocompletado JSON estático (Básico)</h1>
                <p>Fuente: <a href="http://plugins.upbootstrap.com/bootstrap-ajax-typeahead/" target="_blank">http://plugins.upbootstrap.com/bootstrap-ajax-typeahead/</a></p>
                <div class="row">

                    <div class="col-xs-12">
                        <div class="box">
                            <div class="box-body body-form">
                        
                                <div class="form-group">
					                <label>Ciudad</label>
                                    <input id="ciudad" name="ciudad" autocomplete="off" class="typeahead form-control" type="text" required data-bv-notempty-message="El país es requerido" placeholder="Ingrese el nombre una ciudad: Toronto, Montreal, New York...">
                                </div>

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


            <!-- Main content -->
            <section class="content table-custome1">
                <h1>Autocompletado JSON estático + ID</h1>
                <p>Fuente: <a href="http://plugins.upbootstrap.com/bootstrap-ajax-typeahead/" target="_blank">http://plugins.upbootstrap.com/bootstrap-ajax-typeahead/</a></p>
                <div class="row">

                    <div class="col-xs-12">
                        <div class="box">
                            <div class="box-body body-form">
                        
                                <div class="form-group">
					                <label>Ciudad</label>
                                    <input id="ciudad2" name="ciudad2" autocomplete="off" class="typeahead form-control" type="text" required data-bv-notempty-message="El país es requerido" placeholder="Ingrese el nombre una ciudad: Toronto, Montreal, New York...">
                                </div>

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


            <!-- Main content -->
            <section class="content table-custome1">
                <h1>Autocompletado JSON dinámico (AJAX / ASP)</h1>
                <p>Fuente: <a href="http://plugins.upbootstrap.com/bootstrap-ajax-typeahead/" target="_blank">http://plugins.upbootstrap.com/bootstrap-ajax-typeahead/</a></p>
                <div class="row">

                    <div class="col-xs-12">
                        <div class="box">
                            <div class="box-body body-form">
                        
                                <div class="form-group">
					                <label>País</label>
                                    <input id="pais" name="pais" autocomplete="off" class="typeahead form-control" type="text" required data-bv-notempty-message="El país es requerido" placeholder="Ingrese el nombre de su país">
                                </div>

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


            <!-- Main content -->
            <section class="content table-custome1">
                <h1>Autocompletado JSON dinámico (AJAX / ASP) + ID</h1>
                <p>Fuente: <a href="http://plugins.upbootstrap.com/bootstrap-ajax-typeahead/" target="_blank">http://plugins.upbootstrap.com/bootstrap-ajax-typeahead/</a></p>
                <div class="row">

                    <div class="col-xs-12">
                        <div class="box">
                            <div class="box-body body-form">
                        
                                <div class="form-group">
					                <label>País</label>
                                    <input id="pais2" name="pais2" autocomplete="off" class="typeahead form-control" type="text" required data-bv-notempty-message="El país es requerido" placeholder="Ingrese el nombre de su país">
                                </div>

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


            <!-- Main content -->
            <section class="content table-custome1">
                <h1>Autocompletado JSON dinámico (AJAX / ASP) FULL</h1>
                <p>Fuente: <a href="http://plugins.upbootstrap.com/bootstrap-ajax-typeahead/" target="_blank">http://plugins.upbootstrap.com/bootstrap-ajax-typeahead/</a></p>
                <div class="row">

                    <div class="col-xs-12">
                        <div class="box">
                            <div class="box-body body-form">
                        
                                <div class="form-group">
					                <label>País</label>
                                    <input id="pais3" name="pais3" autocomplete="off" class="typeahead form-control" type="text" required data-bv-notempty-message="El país es requerido" placeholder="Ingrese el nombre de su país">
                                </div>

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

        <!-- footer -->
        <!-- #include file="../includes/footer.asp" -->
        
        </div><!-- ./wrapper -->
        
		<!-- scripts -->
        <!-- #include file="../includes/scripts.asp" -->
        
        <%''Inicia personalizado JS de la página%>

        <!-- AdminLTE App -->
        <script src="dist/js/app.min.js" type="text/javascript"></script>

        <script src="plugins/autocomplete/bootstrap-typeahead/bootstrap-typeahead.js" type="text/javascript"></script>
        <script type="text/javascript">
            
            /*Ciudad*/
            /*JSON estático USO BÁSICO*/
            //"$('#ciudad')" es el ID del input text, también puede referenciar un nodo o una clase
            $('#ciudad').typeahead({
                source: [
                    'Toronto',
                    'Montreal',
                    'New York',
                    'Buffalo',
                    'Boston',
                    'Columbus',
                    'Dallas',
                    'Vancouver',
                    'Seattle',
                    'Los Angeles'
                ]
            });

            /*Ciudad 2*/
            /*JSON estático*/
            //"$('#ciudad2')" es el ID del input text, también puede referenciar un nodo o una clase
            $('#ciudad2').typeahead({
                //callback (opcional)
                onSelect: function(item) {
                    //"item" contiene el objeto devuelto por typeahead
                    //estructura "{value: "2", text: "Montreal"}"
                    //"value" es el valor seteado al input y "text" es el texto seteado en el autocomletado
                    console.log(item);
                    //Accedemos al value, para por ejemplo enviarlo a otra función que la procese como en un select anidado que dependa de este
                    console.log(item.value);
                },
                source: [
                    { id: 1, name: 'Toronto' },
                    { id: 2, name: 'Montreal' },
                    { id: 3, name: 'New York' },
                    { id: 4, name: 'Buffalo' },
                    { id: 5, name: 'Boston' },
                    { id: 6, name: 'Columbus' },
                    { id: 7, name: 'Dallas' },
                    { id: 8, name: 'Vancouver' },
                    { id: 9, name: 'Seattle' },
                    { id: 10, name: 'Los Angeles' }
                ]
            });

            /*Paises*/
            /*JSON dinámico*/
            //"$('#pais')" es el ID del input text, también puede referenciar un nodo o una clase
            $('#pais').typeahead({
                ajax: { 
                    //Url del ASP o PHP que devuelve el JSON
                    url: 'tools/autocomplete-data.asp',
                    //cantidad de caracteres minimos para lanzar la consulta
                    triggerLength: 2
                }
            });

            /*Paises*/
            /*JSON dinámico*/
            //"$('#pais2')" es el ID del input text, también puede referenciar un nodo o una clase
            $('#pais2').typeahead({
                ajax: { 
                    //Url del ASP o PHP que devuelve el JSON
                    url: 'tools/autocomplete-data-id.php',
                    
                    //"displayField"
                    //Se debe definir OBLIGATORIAMENTE si el JSON es un conjunto de objetos: [{"id":"AR","name":"Andorra"},{"id":"AR","name":"United Arab Emirates"}] 
                    displayField: "name",
                    
                    //"valueField"
                    //Se debe definir OBLIGATORIAMENTE si el JSON es un conjunto de objetos: [{"id":"AR","name":"Andorra"},{"id":"AR","name":"United Arab Emirates"}]
                    valueField: 'id',
                    
                    //cantidad de caracteres minimos para lanzar la consulta
                    triggerLength: 2

                }
            });

            /*Paises*/
            /*JSON dinámico*/
            //"$('#pais3')" es el ID del input text, también puede referenciar un nodo o una clase
            $("#pais3").typeahead({
                //callback (opcional)
                onSelect: function(item) {
                    //"item" contiene el objeto devuelto por typeahead
                    //estructura "{value: "2", text: "Montreal"}"
                    //"value" es el valor seteado al input y "text" es el texto seteado en el autocomletado
                    console.log(item);
                    //Accedemos al value, para por ejemplo enviarlo a otra función que la procese como en un select anidado que dependa de este
                    console.log(item.value);
                },
                ajax: {
                    url: 'tools/autocomplete-data-id.php',
                    timeout: 500,

                    //"displayField"
                    //Se debe definir OBLIGATORIAMENTE si el JSON es un conjunto de objetos: [{"id":"AR","name":"Andorra"},{"id":"AR","name":"United Arab Emirates"}] 
                    displayField: "name",
                    
                    //"valueField"
                    //Se debe definir OBLIGATORIAMENTE si el JSON es un conjunto de objetos: [{"id":"AR","name":"Andorra"},{"id":"AR","name":"United Arab Emirates"}]
                    valueField: 'id',

                    //cantidad de caracteres minimos para lanzar la consulta
                    triggerLength: 2,

                    method: "post",
                    
                    //clase (OPCIONAL) que se setea en el input mintras se realiza la consulta
                    loadingClass: "loading-circle",

                    preDispatch: function (query) {
                        //console.log(1);
                        //showLoadingMask(true);
                        return {
                            search: query
                        }
                    },
                    preProcess: function (data) {
                        //console.log(2);
                        //showLoadingMask(false);
                        if (data.success === false) {
                            // Hide the list, there was some error
                            return false;
                        }
                        // We good!
                        //console.log(data);
                        return data;
                    }
                }
            });

        </script>
        <%''Fin personalizado JS de la página%>

    </body>

</html>