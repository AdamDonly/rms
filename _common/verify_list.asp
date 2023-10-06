<%
Dim sOrderBy
sOrderBy=UCase(Request.QueryString("ord"))
If sOrderBy<>"E" And sOrderBy<>"R" And sOrderBy<>"I" And sOrderBy<>"U" Then sOrderBy="A"

%>
<html>
<head>
<title>List of experts with simular details</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" type="text/css" href="<% =sHomePath %>styles.css">
</head>
<body bgcolor="#FFFFFF" topmargin=0 leftmargin=0 marginheight=0 marginwidth=0>
<% ShowTopMenu %>

<% ShowMessageStart "info", 580 %>
<b>There are some CVs similar to the one you are going to encode.</b><br>
If the CV you are going to register is in the list click on expert's name to update that CV.

<% ShowMessageEnd %>

<%
sTempParams = ""
sTempParams = ReplaceUrlParams(sTempParams, "act=push")
sTempParams = ReplaceUrlParams(sTempParams, "url=" & sUrl)

sTempParams = ReplaceUrlParams(sTempParams, "exp_language=" & Request.QueryString("exp_language"))
sTempParams = ReplaceUrlParams(sTempParams, "exp_firstname=" & Request.QueryString("exp_firstname"))
sTempParams = ReplaceUrlParams(sTempParams, "exp_familyname=" & Request.QueryString("exp_familyname"))
sTempParams = ReplaceUrlParams(sTempParams, "exp_email=" & Request.QueryString("exp_email"))
%>

	<table width="96%" cellpadding="0" cellspacing="0" border="0" align="center">
	<tr><td bgcolor="#003399">
	<table width="100%" cellpadding="2" cellspacing="1" border="0">
	<tr bgcolor="#E0F3FF">
	<td <%If sOrderBy="I" Then Response.Write " bgcolor=""#99CCFF""" %> width="40" align="center"><p class="sml">Expert&nbsp;ID</p></td>
	<td <%If sOrderBy="A" Then Response.Write " bgcolor=""#99CCFF""" %>><p class="sml" width="40%"><b><% If sOrderBy<>"A" Then %><a href="<%=sScriptFileName & ReplaceUrlParams(sParams, "ord=A") %>"><% End If %>Surname, FirstName MiddleName (Title)</a></b></td>
	<td <%If sOrderBy="R" Then Response.Write " bgcolor=""#99CCFF""" %>><p class="sml" width="50">Date&nbsp;of&nbsp;birth</td>
	<td <%If sOrderBy="U" Then Response.Write " bgcolor=""#99CCFF""" %>><p class="sml" width="50">Email</td>
	<td width="100"><p class="sml">CV&nbsp;language(s)</b></td>
	<td width="100"><p class="sml">Status</b></td>
	<td width="100"><p class="sml">New&nbsp;CV&nbsp;language</b></td>
	</tr>
<%
Dim sCVTypeColor, sCVTypeText
While Not objResult.Eof
sCVTypeColor="#FFFFFF"

	Response.Write "<tr bgcolor=""" & sCVTypeColor & """>"
	Response.Write "<td align=""center""><p class=""mt"">" & objResult("id_Expert") & "</td><td><p class=""mt"">"
	Response.Write "<a href=""register6.asp?id=" & objResult("id_Expert") & """ target=_blank>" & objResult("psnLastName") & ", " & objResult("psnFirstName") & " " & objResult("psnMiddleName") & " (" & objResult("ptlName") & ")</a>"
	Response.Write "</p></td>"
	Response.Write "<td><p class=""mt"">" & objResult("psnBirthDate") & "</td>"
	Response.Write "<td><p class=""mt"">" & objResult("Email") & "</td>"
	Response.Write "<td><p class=""mt"">" & "<a href=""register6.asp?id=" & objResult("id_Expert") & """ target=_blank>" & ReplaceIfEmpty(dictLanguage.Item(Trim(objResult("Lng"))), objResult("Lng")) & "</a></p>"

		Dim bCvLanguageAlreadyRegistered
		bCvLanguageAlreadyRegistered = 0
		Dim objOtherRs
		Set objOtherRs=GetDataRecordsetSP("usp_ExpertIdDetailsSelect", Array( _
			Array(, adInteger, , objResult("id_Expert"))))

		While Not objOtherRs.EOF
			If (objOtherRs("Lng") = Trim(Request.Form("exp_language")) Or objOtherRs("Lng2") = Trim(Request.Form("exp_language"))) Then
				bCvLanguageAlreadyRegistered = 1
			End If
			Response.Write "<p class=""mt"">" & "<a href=""register6.asp?id=" & objOtherRs("id_Expert2") & """ target=_blank>" & ReplaceIfEmpty(dictLanguage.Item(Trim(objOtherRs("Lng2"))), objOtherRs("Lng2")) & "</a></p>"
			objOtherRs.MoveNext
		WEnd
		Set objOtherRs = Nothing

	Response.Write "</td>"
		sCVTypeText=""			
		If objResult("expDeleted")=True Then sCVTypeText=sCVTypeText & " <b>DELETED</b> |"
		If objResult("expRemoved")=True Then sCVTypeText=sCVTypeText & " <b>DELETED</b> |"
		If objResult("expApproved")=True Then sCVTypeText=sCVTypeText & " <b>APPROVED BY EXPERT</b> |"
	Response.Write "<td><p class=""mt"">" & sCVTypeText & "</td>"
	%>
	<td><p class="mt">
		<% If bCvLanguageAlreadyRegistered = 0 Then %>
		<a href="<% =sScriptFileName & ReplaceUrlParams(sTempParams, "lnglink=" & objResult("id_Expert")) & "&" & Request.Form() %>">Register&nbsp;expert's&nbsp;CV<br />in&nbsp;a&nbsp;new&nbsp;language</a>
		<% Else %>
			CV in the selected language is already registered. Click&nbsp;on&nbsp;language&nbsp;to&nbsp;update.
		<% End If %>
	</p></td>
	</tr>
	<% objResult.MoveNext
WEnd %>
</table>
</td></tr>
</table>

<br>
<p align=center><b>If you would like to register a CV for the same expert but in a language not present in the list, click on the appropriate link.</b>
<br>This option will create a new expert in the database - but it will be linked with the original CV.</p>
<br>


<p align=center><a href="<% =sScriptFileName & sTempParams & "&" & Request.Form() %>"><b>If the expert is not in the list - continue encoding new expert</b></a>
<br>This option will create a new expert in the database - <br>please use it if you are sure that the expert is not presented in the list above.</p>
