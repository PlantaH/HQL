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
        <link href="plugins/autocomplete/typeahead.bundle/typeahead.css" rel="stylesheet" type="text/css" />
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
                <h1>Autocomplete</h1>
            </section>

            <!-- Main content -->
            <section class="content table-custome1">
                <h1>Autocomepletado JSON estático</h1>
                <p>Fuente: <a href="https://twitter.github.io/typeahead.js/examples/" target="_blank">https://twitter.github.io/typeahead.js/examples/</a></p>
                <div class="row">

                    <div class="col-xs-12">
                        <div class="box">
                            <div class="box-body body-form">
                        
                                <div id="prefetch" class="form-group">
					                <label>País</label>
                                    <input class="typeahead form-control" type="text" required data-bv-notempty-message="El país es requerido" placeholder="Ingrese el nombre de su país">
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
                <h1>Autocomepletado JSON estático + callback</h1>
                <p>Fuente: <a href="https://stackoverflow.com/questions/10877394/how-to-attach-a-callback-to-bootstrap-typeahead-plugin" target="_blank">https://stackoverflow.com/questions/10877394/how-to-attach-a-callback-to-bootstrap-typeahead-plugin</a></p>
                <div class="row">

                    <div class="col-xs-12">
                        <div class="box">
                            <div class="box-body body-form">
                        
                                <div id="prefetch2" class="form-group">
					                <label>País</label>
                                    <input class="typeahead form-control" type="text" required data-bv-notempty-message="El país es requerido" placeholder="Ingrese el nombre de su país">
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
                <h1>Autocomepletado JSON dinámico</h1>
                <p>Fuente: <a href="https://twitter.github.io/typeahead.js/examples/" target="_blank">https://twitter.github.io/typeahead.js/examples/</a></p>
                <div class="row">

                    <div class="col-xs-12">
                        <div class="box">
                            <div class="box-body body-form">
                        
                                <div id="remote" class="form-group">
					                <label>Peliculas</label>
                                    <input class="typeahead form-control" type="text" required data-bv-notempty-message="Pelicula es requerída" placeholder="Peliculas ganadoras del premio Oscar">
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

        <script src="plugins/autocomplete/typeahead.bundle/typeahead.bundle.js" type="text/javascript"></script>
        <script type="text/javascript">
            
            //Paises
            //JSON estático
            var countries = new Bloodhound({
                datumTokenizer: Bloodhound.tokenizers.whitespace,
                queryTokenizer: Bloodhound.tokenizers.whitespace,
                // url points to a json file that contains an array of country names, see
                // https://github.com/twitter/typeahead.js/blob/gh-pages/data/countries.json
                prefetch: 'plugins/autocomplete/typeahead.bundle/data/countries.json'
            });

            // passing in `null` for the `options` arguments will result in the default
            // options being used
            $('#prefetch .typeahead').typeahead({
                hint: true,
                highlight: true,
                minLength: 1
                },
                {
                name: 'countries',
                source: countries
            });
            //Reset to validate with jqBootstrapValidation.js
            $("input.tt-hint").removeAttr("required data-validation-required-message data-val data-val-required data-val-maxlength data-val-maxlength-max");
            ////



            // passing in `null` for the `options` arguments will result in the default
            // options being used
            $('#prefetch2 .typeahead').typeahead({
                hint: true,
                highlight: true,
                minLength: 1
                },
                {
                name: 'countries',
                source: countries
            }).on('typeahead:selected', function(obj){
                // on selected
                console.log('Se cambió');
                console.log(obj);
                console.log($(this).val());
            });
            //Reset to validate with jqBootstrapValidation.js
            $("input.tt-hint").removeAttr("required data-validation-required-message data-val data-val-required data-val-maxlength data-val-maxlength-max");
            ////




            //Peliculas
            //JSON dinámico
            var bestPictures = new Bloodhound({
                datumTokenizer: Bloodhound.tokenizers.obj.whitespace('value'),
                queryTokenizer: Bloodhound.tokenizers.whitespace,
                prefetch: 'plugins/autocomplete/typeahead.bundle/data/post_1960.json',
                remote: {
                    url: '../data/films/queries/%QUERY.json',
                    wildcard: '%QUERY'
                }
            });

            $('#remote .typeahead').typeahead(null, {
                name: 'best-pictures',
                display: 'value',
                source: bestPictures
            });
            ////


        </script>
        <%''Fin personalizado JS de la página%>

    </body>

</html>