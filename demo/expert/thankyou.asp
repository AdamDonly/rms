<%@ LANGUAGE="VBSCRIPT" %>
<% Option Explicit
Response.Buffer=True
'--------------------------------------------------------------------
'
' CV registration.
' Full format. Confirmation
'
'--------------------------------------------------------------------
%>
<!--#include file="../dbc.asp"-->
<!--#include file="../fnc.asp"-->
<!--#include file="../_forms/frmInterface.asp"-->
<!--#include file="../../_common/expProfile.asp"-->
<% 
	CheckUserLogin sScriptFullNameAsParams

	Dim sUserLogin, sUserPassword, sUserPhone
%>

<html>
    <head>
        <title>Curriculum Vitae. Confirmation.</title>
        <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
        <link rel="stylesheet" type="text/css" href="../styles.css">
    </head>

    <body bgcolor="#FFFFFF" topmargin=0 leftmargin=0  marginheight=0 marginwidth=0>
        <% ShowTopMenu %>
            
            <% ShowMessageStart "info", 550 %>
                <b><%= sApplicationTitle %></b><br><br>
                Thank you for having updating your CV. 
            <% ShowMessageEnd %>

        <% CloseDBConnection %>
    </body>
</html>
