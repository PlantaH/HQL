// Localizador con Google maps(v3) : v1.0
// 18/10/2019

var gMaps_ADM = {};

//Colores del mapa
gMaps_ADM.styles = [];
gMaps_ADM.colorMap = "#003853";//Custom
gMaps_ADM.colorRoad = "#FFB80F";//Custom
gMaps_ADM.customStyle = [];//Custom styles
gMaps_ADM.objRoadBasic = {};//Custom roads
gMaps_ADM.objRoadHighligh = {};//Custom roads resaltadas
gMaps_ADM.latDefault = '40.6976701';
gMaps_ADM.lngDefault = '-74.2598609';
gMaps_ADM.keyMap = $('#map_ad_imglink img').attr('data-key');

var LocalizadorGmap = {

    SearchInputOption : function(){

        /*$('.node-id-puntos-interes').on('click', function(){

            if($(this).hasClass('disabled')){

            }

        });*/

    },
    GmapsPointsInterest : {
        
        Init : function(){

            //$('#mapPunInt').fadeOut();

            //Default lat long seting
            gMaps_ADM.objLatLoad = $("#map_ad_lat").val() ? $("#map_ad_lat").val() : gMaps_ADM.latDefault;
            gMaps_ADM.objLngLoad = $("#map_ad_lng").val() ? $("#map_ad_lng").val() : gMaps_ADM.lngDefault;
            //
            
            //Centrado inicial (se redefine al final)
            var latlng_ADM = new google.maps.LatLng(gMaps_ADM.objLatLoad, gMaps_ADM.objLngLoad);
            //
            
            gMaps_ADM.geocoder = new google.maps.Geocoder();
            
            //Estilo del mapa!
            //Se cargan los estilos custom
            gMaps_ADM.styles = gMaps_ADM.customStyle;
            
            
            //Opciones del mapa (seteo) 
            var options = {
                'zoom': 10, //Zomm inicial (se redefine al final)
                'center': latlng_ADM,
                'styles': gMaps_ADM.styles,
                mapTypeControl: true,
                mapTypeControlOptions: {
                    style: google.maps.MapTypeControlStyle.TOP_LEFT,
                    position: google.maps.ControlPosition.RIGHT_BOTTOM
                }
            
            };
            
            //Mapa
            gMaps_ADM.map = new google.maps.Map(document.getElementById('mapPunInt'), options);
            //
                
            //Buscador puntos de interes
            var input_ADM = /** @type {!HTMLInputElement} */(document.getElementById('map_ad_search_bar'));
            
            //var types = document.getElementById('type-selector');
            //gMaps_ADM.map.controls[google.maps.ControlPosition.TOP_LEFT].push(input_ADM);
            //gMaps_ADM.map.controls[google.maps.ControlPosition.TOP_LEFT].push(types);
            
            var autocomplete_ADM = new google.maps.places.Autocomplete(input_ADM);
            autocomplete_ADM.bindTo('bounds', gMaps_ADM.map);

            var initPosition = new google.maps.LatLng(gMaps_ADM.latDefault, gMaps_ADM.lngDefault);
            
            var infowindow_ADM = new google.maps.InfoWindow();
            var marker_ADM = new google.maps.Marker({
                map: gMaps_ADM.map,
                position: latlng_ADM,
                draggable: true,
                animation: google.maps.Animation.DROP,
                anchorPoint: new google.maps.Point(0, -29)
            });

            //Change marker
            var imageIcon_ADM = 'https://maps.gstatic.com/mapfiles/api-3/images/spotlight-poi2.png';
            marker_ADM.addListener('dragend', function() {
                //infowindow_ADM.setContent('test');
                marker_ADM.setIcon(imageIcon_ADM);
                
                objData_ADM = {
                    place : null,
                    lat : this.getPosition().lat(),
                    lng : this.getPosition().lng()
                }

                reloadData(objData_ADM);

            });
            //

            autocomplete_ADM.addListener('place_changed', function() {

                infowindow_ADM.close();

                $('#mapPunInt').fadeIn();//muestra el mapa

                marker_ADM.setVisible(false);
                var place = autocomplete_ADM.getPlace();
                if (!place.geometry) {
                    window.alert("Autocompetar requiere que ingrese una búsqueda");
                    return;
                }
        
                // If the place has a geometry, then present it on a map.
                if(place.geometry.viewport){
                    gMaps_ADM.map.fitBounds(place.geometry.viewport);
                }else{
                    gMaps_ADM.map.setCenter(place.geometry.location);
                    gMaps_ADM.map.setZoom(14);// Why 17? Because it looks good.
                }

                //customiza el marcador de acurdo a tipo (si se saca muestra el marcador por default)
                marker_ADM.setIcon(/** @type {google.maps.Icon} */({
                    url: place.icon,
                    size: new google.maps.Size(71, 71),
                    origin: new google.maps.Point(0, 0),
                    anchor: new google.maps.Point(17, 34),
                    scaledSize: new google.maps.Size(35, 35)
                }));
                //

                marker_ADM.setPosition(place.geometry.location);
                marker_ADM.setVisible(true);
        
                var address = '',
                    cityName = '',
                    countryName;
                if (place.address_components) {
                    address = [
                        (place.address_components[0] && place.address_components[0].short_name || ''),
                        (place.address_components[1] && place.address_components[1].short_name || ''),
                        (place.address_components[2] && place.address_components[2].short_name || '')
                    ].join(' ');

                    //find country name
                    for (var i=0; i<place.address_components.length; i++) {
                        for (var ii=0; ii<place.address_components[i].types.length; ii++) {
                                    
                            if(place.address_components[i].types[ii] == "locality") {
                                //this is the object you are looking for
                                cityName = place.address_components[i];
                                //break;
                            }
                            if(place.address_components[i].types[ii] == "country") {
                                //this is the object you are looking for
                                countryName = place.address_components[i];
                                //break;
                            }

                        }
                    }


                    if(cityName == '') {					

                        var geocoder = new google.maps.Geocoder;
                        geocodeLatLng(geocoder, place, gMaps_ADM.map, false);

                        //cityName = {long_name : 'long_name'};

                    }

                    let latNew = String(place.geometry.location.lat()).replace('.',','),
                        lngNew = String(place.geometry.location.lng()).replace('.',',');

                    //set values
                    LocalizadorGmap.GmapsPointsInterest.Set({
                        long_name : cityName.long_name,
                        short_name : countryName.short_name,
                        long_name : countryName.long_name,
                        address : $('#map_ad_search_bar').val(),
                        lat : latNew,
                        lng : lngNew
                    });
         
                }
                //console.log(place.address_components);
                var thisCenter = String(place.geometry.location);
        
                thisCenter = thisCenter.replace("(", '');
                thisCenter = thisCenter.replace(")", '');
                thisCenter = thisCenter.split(',');
        
                var htmlToolTipAutocomplete_ADM = '<div class="map--ad-search-mapwindow-tootip">' +
                                            '<strong>' + place.name + '</strong>' +
                                            '<span>' + address + '</span>' +
                                            '<div>' +
                                            '<small>(Puede arrastrar este marcador para ajustar las coordenadas)</small>' +
                                            '</div>';
        
                infowindow_ADM.setContent(htmlToolTipAutocomplete_ADM);
                infowindow_ADM.open(gMaps_ADM.map, marker_ADM);

            });

            function geocodeLatLng(geocoder, place, map, reload) {

                geocoder.geocode({'location': place.geometry.location}, function(results, status) {
                    console.log(place.geometry.location);
                    if(status === 'OK'){
                        if(results[0]){

                            var cityName = null;

                            for(var i=0; i<results[0].address_components.length; i++){
                                for (var ii=0; ii<results[0].address_components[i].types.length; ii++) {
                                            
                                    /*if(results[0].address_components[i].types[ii] == "administrative_area_level_1") {
                                        //this is the object you are looking for
                                        cityName = results[0].address_components[i];
                                        //break;
                                    }*/
                                    if(results[0].address_components[i].types[ii] == "locality") {
                                        //this is the object you are looking for
                                        cityName = results[0].address_components[i];
                                        //break;
                                    }
                                    if(results[0].address_components[i].types[ii] == "country") {
                                        //this is the object you are looking for
                                        countryName = results[0].address_components[i];
                                        //break;
                                    }

                                    //console.log(results[0].address_components[i].types[ii]);
            
                                }

                            }
                            console.log('------------------');
                            console.log(cityName);
                            console.log('------------------');

                            if(cityName === null){

                                for(var i=0; i<results[0].address_components.length; i++){
                                    for (var ii=0; ii<results[0].address_components[i].types.length; ii++) {
                                                
                                        if(results[0].address_components[i].types[ii] == "administrative_area_level_1") {
                                            //this is the object you are looking for
                                            cityName = results[0].address_components[i];
                                            //break;
                                        }
                
                                    }

                                }

                            }


                            if(reload){

                                console.log(results[0].formatted_address);
                                console.log(results[0]);
                                $("#map_ad_search_bar").val(results[0].formatted_address);

                                LocalizadorGmap.GmapsPointsInterest.Set({
                                    long_name : cityName.long_name,
                                    short_name : countryName.short_name,
                                    long_name : countryName.long_name,
                                    address : results[0].formatted_address
                                });

                                var htmlToolTipAutocomplete_reload_ADM = '<div class="map--ad-search-mapwindow-tootip">' +
                                '<strong>' + cityName.long_name + ', ' + countryName.long_name + '</strong>' +
                                '<span>' + results[0].formatted_address + '</span>' +
                                '<div>' +
                                '<small>(Puede arrastrar este marcador para ajustar las coordenadas)</small>' +
                                '</div>';

                                infowindow_ADM.setContent(htmlToolTipAutocomplete_reload_ADM);

                            }

                            /*$("#map_ad_cityname").val(cityName.long_name);
                            $("#map_ad_countryNameShort").val(countryName.short_name);
                            $("#map_ad_countryNameLong").val(countryName.long_name);*/

                            LocalizadorGmap.GmapsPointsInterest.Set({
                                long_name : cityName.long_name,
                                short_name : countryName.short_name,
                                long_name : countryName.long_name
                            });

                        }else{
                            //window.alert('No results found');
                            var htmlToolTipAutocomplete_reload_ADM = '<div class="map--ad-search-mapwindow-tootip">' +
                            '<strong>No se encontró información!</strong>' +
                            '<span>Vuelve a arrastrar el marcador sobre el mapa<br>o ingresa la búsqueda en el campo para "puntos de interés"</span>' +
                            '<div>' +
                            '</div>';

                            infowindow_ADM.setContent(htmlToolTipAutocomplete_reload_ADM);
                        }

                    }else{
                        //window.alert('Geocoder failed due to: ' + status);
                        var htmlToolTipAutocomplete_reload_ADM = '<div class="map--ad-search-mapwindow-tootip">' +
                        '<strong>Puede que estés situado sobre el mar, en un punto ciego o no registrado!</strong>' +
                        '<span>Vuelve a arrastrar el marcador sobre el mapa<br>o ingresa la búsqueda en el campo para "puntos de interés"</span>' +
                        '<div>' +
                        '</div>';

                        infowindow_ADM.setContent(htmlToolTipAutocomplete_reload_ADM);
                    }
                });

            }

            function reloadData(obj){

                var geocoder = new google.maps.Geocoder,
                    customPlace = {geometry: { location: {lat:obj.lat, lng:obj.lng } }};

                geocodeLatLng(geocoder, customPlace, gMaps_ADM.map, true);

                /*$("#map_ad_lat").val(obj.lat);
                $("#map_ad_lng").val(obj.lng);*/
                //set values
                LocalizadorGmap.GmapsPointsInterest.Set({
                    lat : obj.lat,
                    lng : obj.lng
                });

            }

        },
        Set : function(obj){

            //form
            if(obj.long_name) $("#map_ad_cityname").val(obj.long_name);
            if(obj.short_name) $("#map_ad_countryNameShort").val(obj.short_name);
            if(obj.long_name) $("#map_ad_countryNameLong").val(obj.long_name);
            if(obj.address) $("#map_ad_address").val(obj.address);
            if(obj.lat) $("#map_ad_lat").val(obj.lat);
            if(obj.lng) $("#map_ad_lng").val(obj.lng);

            //link
            if(obj.lat && obj.lng) $('#map_ad_imglink img').attr('src', 'https://maps.googleapis.com/maps/api/staticmap?markers='+obj.lat+','+obj.lng+'&size=200x120&;zoom=14&key='+gMaps_ADM.keyMap);
            if(obj.lat) $("#map_ad_latlink").text(obj.lat);
            if(obj.lng) $("#map_ad_lnglink").text(obj.lng);

        },
        Unset : function(){

            $("#map_ad_search_bar").val('');
            $("#map_ad_cityname").val('');
            $("#map_ad_countryNameShort").val('');
            $("#map_ad_countryNameLong").val('');
            //$("#map_ad_placeName").val('');
            $("#map_ad_address").val('');
            $("#map_ad_lat").val('');
            $("#map_ad_lng").val('');

        },
        Close : function(){

            $('#mapPunInt').fadeOut();

        }

    }

}