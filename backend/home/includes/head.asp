<!-- #include file="../../includes/conn.asp" -->
<%
versionNoCache = "0.0.12"

basePage = Application("URL_BASE") 
rootPage = "default.asp"
%> 

        <meta charset="utf-8" />
        <base href="<%=basePage%>">

        <meta name="viewport" content="width=device-width, initial-scale=1, maximum-scale=1" />
        <meta name="description" content="" />
        <meta name="author" content="" />
        <!--[if IE]>
            <meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1">
            <![endif]-->
        <title>Hotel y Pago</title>
        <!-- BOOTSTRAP CORE STYLE CSS -->
        <link href="home/css/bootstrap.css" rel="stylesheet" />
        <!-- FONT AWESOME STYLE CSS -->
        <link href="home/font-awesome-4.7.0/css/font-awesome.min.css" rel="stylesheet" />
        <!-- CUSTOM STYLE CSS -->
        <link href="home/css/style.css?v=<%=versionNoCache%>" rel="stylesheet" />
        <!-- GOOGLE FONT -->
        <link href='http://fonts.googleapis.com/css?family=Open+Sans:400,700' rel='stylesheet' type='text/css' />