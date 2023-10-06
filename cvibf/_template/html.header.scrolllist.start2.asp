<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<!--	<meta content="text/html; charset=UTF-8" http-equiv="Content-Type"/> -->
	<meta content="text/html; charset=iso-8859-1" http-equiv="Content-Type"/>
	<link rel="stylesheet" type="text/css" media="all" href="/main.css" />
	<link rel="stylesheet" type="text/css" href="/style.css" />
	<link rel="shortcut icon" href="/favicon.ico" />
	<link rel="icon" type="image/png" href="/image/favicon.png" />

	<meta name="viewport" content="width=device-width" />
	<meta content="text/html; charset=iso-8859-1" http-equiv="Content-Type"/>

	<title><% =sPageTitle %></title>

	<!-- / CSS IMPORTS \ -->
	<link rel="stylesheet" href="/res/css/style.css">
	<link rel="stylesheet" href="//code.jquery.com/ui/1.10.2/themes/smoothness/jquery-ui.css" />
	<!--[if lte IE 7]>
	<link rel="stylesheet" type="text/css" href="/res/css/ie7.css" />
	<![endif]-->
	<!--[if ie 8]>
	<link rel="stylesheet" type="text/css" href="/res/css/ie8.css" />
	<![endif]-->
	<!--[if ie 9]>
	<link rel="stylesheet" type="text/css" href="/res/css/ie9.css" />
	<![endif]-->
	<!-- \ END CSS IMPORTS / -->


	<!-- / SCRIPTS IMPORTS \ -->
	<% 
	' load slider for the homepage:
	If sScriptFileName = "default.asp" Then 
		%><script src="//ajax.googleapis.com/ajax/libs/jquery/1.8.2/jquery.min.js"></script>
		<%
	Else
		%>
		<script src="//ajax.googleapis.com/ajax/libs/jquery/1.9.1/jquery.min.js"></script>
		<script src="//ajax.googleapis.com/ajax/libs/jqueryui/1.10.2/jquery-ui.min.js"></script>
		<%
	End If %>

	<!-- \ END SCRIPTS IMPORTS / -->