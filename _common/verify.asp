<%
'--------------------------------------------------------------------
'
' CV registration.
' Short format.
'
'--------------------------------------------------------------------
Response.Expires = 0
Response.ExpiresAbsolute = Now() - 1
Response.CacheControl = "no-cache" 
%>
<!--#include file="_data/datGender.asp"-->
<!--#include file="_data/datPsnTitle.asp"-->
<!--#include file="_data/datMonth.asp"-->
<%
If sApplicationName<>"expert" Then
	CheckUserLogin sScriptFullNameAsParams
End If
%>
<!--#include file="../_common/expProfile.asp"-->
<html>
<head>
<title>New CV registration</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" type="text/css" href="<% =sHomePath %>styles.css">
<script type="text/javascript" src="/_scripts/js/main.js"></script>
<script type="text/javascript">
<!-- 
function validateForm() {
	var f=document.forms[0];
	if (!(f)) {
		return false; }
	if (!checkTextFieldValue(f.exp_firstname, "", "Please fill in expert's first name.", 1)) { return false }
	if (!checkTextFieldValue(f.exp_familyname, "", "Please fill in expert's family name.", 1)) { return false }
	<% If InStr(sUrl, "invite.asp")>0 Then %>
	if (!checkTextFieldValue(f.exp_email, "", "Please fill in expert primary email.", 1)) { return false }
	<% End If %>
	if ((f.exp_email.value) && (!validateEmail(f.exp_email.value))) {
		alert("Please retype e-mail address correctly");
        f.exp_email.select();        
		return;
	}
	f.submit();
}
-->
</script>
</head>
<body topmargin=0 leftmargin=0  marginheight=0 marginwidth=0>
<% ShowTopMenu %>

<% ShowMessageStart "info", 550 %>
Please complete this form in order to check if expert already exists in the database.
<br>Fields marked with <span class="rs">*</span> are mandatory.<br>
<% ShowMessageEnd %>

  <!-- Personal information -->
	<% InputFormHeader 580, "NEW EXPERT REGISTRATION" %>
	<% InputBlockHeader "100%" %>
	<form action="verify_results.asp<% =AddUrlParams(sParams, "act=" & sAction) %>"  method="post" onSubmit="validateForm(); return false;">
<% If bCvMultiLanguageActive = cMultiLanguageEnabled Then %>
		<% InputBlockSpace 4 %>
		<% InputBlockElementLeftStart %><p class="ftxt"><% = GetLabel(sCvLanguage, "CV language") %></p><% InputBlockElementLeftEnd %><% InputBlockElementMiddle %>
		<% InputBlockElementRightStart %>&nbsp;&nbsp;<select name="exp_language" size="1" style="width:130px;">
		<%
		Dim sTempLanguage
		For Each sTempLanguage in dictLanguage
			Response.Write "<option value=""" & sTempLanguage & """" 
			If sCvLanguage=sTempLanguage Then Response.Write " selected"
			Response.Write ">" & dictLanguage.Item(sTempLanguage) & "</option>"
		Next
		%>
		</select><% InputBlockElementRightEnd %>
		<% InputBlockSpace 4 %>
	<% InputBlockFooter %>
	<% InputFormAfterBlock %>
	<% InputFormDualLine %>

  <!-- Personal information -->
	<% InputFormBeforeBlock 580 %>
	<% InputBlockHeader "100%" %>
<% End If %>

		<% InputBlockSpace 4 %>
		<% InputBlockElementLeftStart %><p class="ftxt">First&nbsp;name</p><% InputBlockElementLeftEnd %><% InputBlockElementMiddle %>
		<% InputBlockElementRightStart %>&nbsp;&nbsp;<input type="text" name="exp_firstname" size="31" style="width:355px;" maxlength="100" value="">&nbsp;&nbsp;<span class="fcmp">*</span>&nbsp;<% InputBlockElementRightEnd %>
		<% InputBlockElementLeftStart %><p class="ftxt">Family&nbsp;name</p><% InputBlockElementLeftEnd %><% InputBlockElementMiddle %>
		<% InputBlockElementRightStart %>&nbsp;&nbsp;<input type="text" name="exp_familyname" size="31" style="width:355px;" maxlength="100" value="">&nbsp;&nbsp;<span class="fcmp">*</span>&nbsp;<% InputBlockElementRightEnd %>
		<% InputBlockElementLeftStart %><p class="ftxt">Primary email</p><% InputBlockElementLeftEnd %><% InputBlockElementMiddle %>
		<% InputBlockElementRightStart %>&nbsp;&nbsp;<input type="text" name="exp_email" size="31" style="width:355px;" maxlength="120" value=""><% If InStr(sUrl, "invite.asp")>0 Then %>&nbsp;&nbsp;<span class="fcmp">*</span>&nbsp;<% End If %><% InputBlockElementRightEnd %>
		<% InputBlockSpace 6 %>
	<% InputBlockFooter %>
	<% InputFormFooter %>
	<% InputFormSpace 12 %>

	<table width=580 cellspacing=0 cellpadding=0 border=0 align="center">
	<tr>
	<td width=580 align=left><img src="<% =sHomePath %>image/x.gif" width=170 height=1><input type="image" src="<% =sHomePath %>image/bte_submit.gif" border=0 alt="  Submit  ">
	&nbsp; &nbsp; &nbsp; &nbsp; <a href="<% =sApplicationHomePath & sParams %>"><img src="<% =sHomePath %>/image/bte_cancel.gif" border=0 alt="  Cancel  "</a>
	</td>
	</tr>
	</form>
	</table><br>

<% CloseDBConnection %>
</body>
</html>

