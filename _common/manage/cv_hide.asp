<%
'--------------------------------------------------------------------
'
' IBF CV registration.
' Hiding a duplicate CV
'
'--------------------------------------------------------------------
%>
<!--#include file="../cv_data.asp"-->
<% 
CheckUserLogin sScriptFullNameAsParams
CheckExpertID()

Dim sMessage
If Request.Form()>"" Then
	ShowStandardPageHeader	

	If iExpertID=CheckIntegerAndZero(Request.Form("exp_originalid")) Then
		sMessage="You cannot hide a duplicate CV with itself. Please provide the correct Original Expert ID."
	End If
	
	objTempRs=GetDataOutParamsSP("usp_AdmExpDuplicateHide", Array( _
		Array(, adInteger, , iExpertID), _
		Array(, adInteger, , CheckInteger(Request.Form("exp_originalid"))), _
		Array(, adLongVarWChar, 20000, CheckString(Request.Form("exp_comments")))), Array( _ 
		Array(, adInteger)))
	
	If objTempRs(0)>=1 Then
		sMessage="This duplicate copy of expert's CV (ID=" & iExpertID & ") was successfully hidden."
	End If
	%>
	<br><br>
	<p align="center"><% =sMessage %></p>
	<br><p align="center"><a href="<% =sApplicationHomePath %><% =sParams %>"><img src="<% =sHomePath %>image/bte_continue.gif" border=0></a></p>
	<%
	ShowStandardPageFooter
	Response.End
End If
%>
<html>
<head>
<title>Hide duplicate copy of CV</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" type="text/css" href="<% =sHomePath %>styles.css">
</head>
<body bgcolor="#FFFFFF" topmargin=0 leftmargin=0 marginheight=0 marginwidth=0>
<% ShowTopMenu %>

	<br>
	<% InputFormHeader 580, "HIDE DUPLICATE CV" %>
	<% InputBlockHeader "100%" %>
	<form method="post" action="cv_hide.asp<%=sParams%>" name="RegForm">
	<input type="hidden" name="id_Expert" value="<%=iExpertID%>">
		<% InputBlockSpace 4 %>
		<% InputBlockElementLeftStart %><p class="ftxt">Expert ID</p><% InputBlockElementLeftEnd %><% InputBlockElementMiddle %>
		<% InputBlockElementRightStart %><p><b>&nbsp;<% =iExpertID %></b></p><% InputBlockElementRightEnd %>
		<% InputBlockElementLeftStart %><p class="ftxt">Full&nbsp;name</p><% InputBlockElementLeftEnd %><% InputBlockElementMiddle %>
		<% InputBlockElementRightStart %><p>&nbsp;<% =sFullName %></p><% InputBlockElementRightEnd %>
		<% InputBlockElementLeftStart %><p class="ftxt">Date of birth</p><% InputBlockElementLeftEnd %><% InputBlockElementMiddle %>
		<% InputBlockElementRightStart %><p>&nbsp;<% =ConvertDateForText(sBirthDate, "&nbsp;", "DDMMYYYY") %></p><% InputBlockElementRightEnd %>
		<% InputBlockElementLeftStart %><p class="ftxt">Email</p><% InputBlockElementLeftEnd %><% InputBlockElementMiddle %>
		<% InputBlockElementRightStart %><p>&nbsp;<% =sUserEmail %></p><% InputBlockElementRightEnd %>
		<% InputBlockSpace 4 %>
	<% InputBlockFooter %>
	<% InputFormAfterBlock %>
	<% InputFormDualLine %>

	<% InputFormBeforeBlock 580 %>
	<% InputBlockHeader "100%" %>
		<% InputBlockSpace 4 %>
		<% InputBlockElementLeftStart %><p class="ftxt">Original Expert ID</p><% InputBlockElementLeftEnd %><% InputBlockElementMiddle %>
		<% InputBlockElementRightStart %>&nbsp;&nbsp;<input type="text" name="exp_originalid" size="10" style="width=75px;"><% InputBlockElementRightEnd %>
		<% InputBlockElementLeftStart %><p class="ftxt">Comments</p><% InputBlockElementLeftEnd %><% InputBlockElementMiddle %>
		<% InputBlockElementRightStart %>&nbsp;&nbsp;<textarea cols="34" style="width=355px;" name="exp_comments" rows=5 wrap="yes"></textarea>
		<p class="sml">&nbsp;&nbsp;Please provide your comments if necessary.</p><% InputBlockElementRightEnd %>
		<% InputBlockSpace 10 %>
	<% InputBlockFooter %>
	<% InputFormFooter %>
	<% InputFormSpace 12 %>

	<div align="center">
	<input type="image" src="<% =sHomePath %>image/bte_hideexpert.gif" vspace="0" border="0">
	</div>
</form>
</body>
</html>

