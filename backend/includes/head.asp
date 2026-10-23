<!-- #include file="conn.asp" -->
<% 
versionNoCache = "0.0.8"
 
Dim logueado, titulo

if session("log") then logueado=true
%>
        <meta charset="UTF-8">
        <title><%=Application("SITE_NAME")%></title>
        
        <base href="<%=Application("URL_BASE")%>">
        
        <meta content='width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no' name='viewport'>
        <!-- Bootstrap 3.3.2 -->
        <!-- <link href="bootstrap/css/bootstrap.min.css" rel="stylesheet" type="text/css" /> -->

        <!-- Bootstrap 3.3.7 -->
        <link href="bootstrap/3.3.7/bootstrap.min.css" rel="stylesheet" type="text/css" /> 

        <!-- FontAwesome 4.3.0 -->
        <link href="bootstrap/font-awesome/4.3.0/css/font-awesome.min.css" rel="stylesheet" type="text/css" />
        <!-- Ionicons 2.0.0 -->
        <!--<link href="http://code.ionicframework.com/ionicons/2.0.0/css/ionicons.min.css" rel="stylesheet" type="text/css" />-->

        <!-- CSS Autocomplete -->
        <link rel="stylesheet" href="plugins/easy-autocomplete/easy-autocomplete.min.css"> 

		<!-- FontAwesome 4.7.0 -->
		<link rel="stylesheet" type="text/css" href="Upload_New/font-awesome/4.7.0/css/font-awesome.min.css">
		<!-- VanillaUpload 0.0.1 -->
		<link rel="stylesheet" type="text/css" href="Upload_New/css/bootstrap/style.css">
		<script src="Upload_New/js/uploadNew.js?v=0.0.5"></script>

        <link href="plugins/jQueryUI/css.1.11.2/jquery-ui.css" rel="stylesheet" type="text/css" />
