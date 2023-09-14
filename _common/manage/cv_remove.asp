<% 
'--------------------------------------------------------------------
'
' IBF CV registration.
' Removing expert from the database
'
'--------------------------------------------------------------------
%>
<!--#include file="../cv_data.asp"-->
<% 
CheckUserLogin sScriptFullNameAsParams
CheckExpertID()

If Request.Form()>"" Then
	sParams=ReplaceUrlParams(sParams, "id")
	ShowStandardPageHeader	

	objTempRs=GetDataOutParamsSP("usp_AdmExpRemove", Array( _
		Array(, adInteger, , iExpertID), _
		Array(, adInteger, , CheckInteger(Request.Form("exp_reason"))), _
		Array(, adLongVarWChar, 20000, CheckString(Request.Form("exp_comments")))), Array( _ 
		Array(, adInteger)))
	
	If objTempRs(0)>=1 Then
		Response.Write "<br><br><br><br><p align=""center"">The CV of the expert with ID " & iExpertID & " was successfully deleted."
	End If
	%><br><br>
	<a href="<% =sApplicationHomePath %><% =sParams %>"><img src="<% =sHomePath %>image/bte_continue.gif" border=0></a>
	<%
	ShowStandardPageFooter
	Response.End
End If
%>
<html>
<head>
<title>Remove expert's CV from the database</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" type="text/css" href="<% =sHomePath %>styles.css">
</head>
<body bgcolor="#FFFFFF" topmargin=0 leftmargin=0 marginheight=0 marginwidth=0>
<% ShowTopMenu %>

	<br>
	<% InputFormHeader 580, "REMOVE EXPERT FROM DATABASE" %>
	<% InputBlockHeader "100%" %>
	<form method="post" action="cv_remove.asp<%=sParams%>" name="RegForm">
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
		<% InputBlockElementLeftStart %><p class="ftxt">Reason</p><% InputBlockElementLeftEnd %><% InputBlockElementMiddle %>
		<% InputBlockElementRightStart %><p>
		<input type="radio" name="exp_reason" value="1" checked>&nbsp;This is not a real CV<br>
		<input type="radio" name="exp_reason" value="2">&nbsp;Expert doesn't want to have his CV being registered in the database<% InputBlockElementRightEnd %>
		<% InputBlockElementLeftStart %><p class="ftxt">Comments</p><% InputBlockElementLeftEnd %><% InputBlockElementMiddle %>
		<% InputBlockElementRightStart %><p><textarea cols="34" style="width=355px;" name="exp_comments" rows=5 wrap="yes"></textarea>
		<p class="sml">&nbsp;Please paste here full text of expert's email, <br>&nbsp;where he is asking to be removed from the database</p><% InputBlockElementRightEnd %>
		<% InputBlockSpace 10 %>
	<% InputBlockFooter %>
	<% InputFormFooter %>
	<% InputFormSpace 12 %>

	<div align="center">
	<input type="image" src="<% =sHomePath %>image/bte_removeexpert.gif" vspace="0" border="0">
	</div>
</form>
</body>
</html>

