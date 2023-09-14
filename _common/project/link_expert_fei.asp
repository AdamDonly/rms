<%
'--------------------------------------------------------------------
'
' Link expert.
'
'--------------------------------------------------------------------
Response.Expires = 0
Response.ExpiresAbsolute = Now() - 1
Response.CacheControl = "no-cache" 
%>
<!--#include file="../_data/datMonth.asp"-->
<!--#include file="../_data/datExpertStatus.asp"-->
<!--#include file="../_data/datCurrency.asp"-->
<!--#include virtual="/_common/_class/main.asp"-->
<!--#include virtual="/_common/_class/project.asp"-->
<!--#include virtual="/_common/_class/expert.asp"-->
<!--#include virtual="/_common/_class/expert.project.asp"-->
<!--#include virtual="/_common/expProfile.asp"-->
<%
If sApplicationName<>"expert" Then
	CheckUserLogin sScriptFullNameAsParams
Else 
	Response.Redirect sApplicationHomePath
End If

Dim iProjectID
iProjectID=CheckIntegerAndZero(Request.QueryString("idproject"))

Dim sUserPhone
iExpertID=CheckIntegerAndZero(Request.QueryString("idexpert"))
LoadExpertProfile(iExpertID)

Dim objExpertProject
Set objExpertProject = New CExpertProject
objExpertProject.Project.ID=iProjectID
objExpertProject.Project.LoadData
objExpertProject.Expert.ID=iExpertID
objExpertProject.LoadData

If iProjectID=0 Then
	Dim objProjectList
	Set objProjectList = New CProjectList
	objProjectList.LoadData
End If

%>
<html>
<head>
<title>Add expert on project</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" type="text/css" href="<% =sHomePath %>styles.css">
<script type="text/javascript" src="/_scripts/js/main.js"></script>
<script type="text/javascript">
<!-- 
function validateForm() {
	var f=document.forms[0];
	if (!(f)) {
		return false; }
	if (!checkTextFieldValue(f.project_title, "", "Please fill in project title.", 1)) { return false }
	f.submit();
}
-->
</script>
</head>
<body topmargin=0 leftmargin=0  marginheight=0 marginwidth=0>
<% ShowTopMenu %>
  <!-- Personal information -->
	<% InputFormHeader 580, "ADD EXPERT ON PROJECT" %>
	<% InputBlockHeader "100%" %>
	<form action="link_expert_save.asp<% =AddUrlParams(sParams, "act=" & sAction) %>"  method="post" onSubmit="validateForm(); return false;">
		<% InputBlockSpace 4 %>
	<% If iProjectID=0 Then %>
		<% InputBlockElementLeftStart %><p class="ftxt">Project</p><% InputBlockElementLeftEnd %><% InputBlockElementMiddle %>
		<% InputBlockElementRightStart %>&nbsp;&nbsp;<select size="1" name="idproject" style="width:355px;">
		<option value="0"> </option>
		<%  For i=0 to objProjectList.Count-1
			Response.Write("<option value=" & objProjectList.Item(i).ID & ">" & objProjectList.Item(i).Title & "</option>")
		Next %>
		</select>&nbsp;&nbsp;<span class="fcmp">*</span>&nbsp;<% InputBlockElementRightEnd %>
	<% Else %>
		<input type="hidden" name="idproject" value="<% =objExpertProject.Project.ID %>">
		<% InputBlockElementLeftStart %><p class="ftxt">Project&nbsp;title</p><% InputBlockElementLeftEnd %><% InputBlockElementMiddle %>
		<% InputBlockElementRightStart %><p class="ftxtright"><b><% =objExpertProject.Project.Title %></b></p><% InputBlockElementRightEnd %>
	<% End If %>
		<input type="hidden" name="idexpert" value="<% =objExpertProject.Expert.ID %>">
		<% InputBlockElementLeftStart %><p class="ftxt">Expert&nbsp;name</p><% InputBlockElementLeftEnd %><% InputBlockElementMiddle %>
		<% InputBlockElementRightStart %><p class="ftxtright"><b><% =sFullName %></b></p><% InputBlockElementRightEnd %>
		<% InputBlockElementLeftStart %><p class="ftxt">Expert&nbsp;status</p><% InputBlockElementLeftEnd %><% InputBlockElementMiddle %>
		<% InputBlockElementRightStart %>&nbsp;&nbsp;<select size="1" name="expert_status" style="width:355px;">
		<option value="0"> </option>
		<%  For i=UBound(arrExpertStatusID) to LBound(arrExpertStatusID) step -1
		If objExpertProject.Status.ID=arrExpertStatusID(i) Then
			Response.Write("<option value=" & arrExpertStatusID(i) & " selected>" & arrExpertStatusTitle(i) & "</option>")
		Else 
			Response.Write("<option value=" & arrExpertStatusID(i) & ">" & arrExpertStatusTitle(i) & "</option>")
		End If
		Next %>
		</select>&nbsp;&nbsp;<span class="fcmp">*</span>&nbsp;<% InputBlockElementRightEnd %>
		<% InputBlockElementLeftStart %><p class="ftxt">Expert&nbsp;fee</p><% InputBlockElementLeftEnd %><% InputBlockElementMiddle %>
		<% InputBlockElementRightStart %>&nbsp;&nbsp;<input type="text" name="expert_fee" size="31" style="width:60px;" maxlength="8" value="<% =objExpertProject.Fee.Value %>">&nbsp;<select size="1" name="expert_fee_currency" style="width:55px; margin-top:0px;">
		<option value="0"> </option>
		<%  For i=LBound(arrCurrencyCode) to UBound(arrCurrencyCode)
		If objExpertProject.Fee.CurrencyCode=arrCurrencyCode(i) Then
			Response.Write("<option value=" & arrCurrencyCode(i) & " selected>" & arrCurrencyCode(i) & "</option>")
		Else 
			Response.Write("<option value=" & arrCurrencyCode(i) & ">" & arrCurrencyCode(i) & "</option>")
		End If
		Next %>
		</select><% InputBlockElementRightEnd %>
		<% InputBlockElementLeftStart %><p class="ftxt">Provided by (day, month)</p><% InputBlockElementLeftEnd %><% InputBlockElementMiddle %>
		<% InputBlockElementRightStart %>&nbsp;&nbsp;<input type="text" name="expert_provided_company" size="31" style="width:355px;" maxlength="100" value="<% =objExpertProject.ProvidedCompany %>"><% InputBlockElementRightEnd %>
		<% InputBlockElementLeftStart %><p class="ftxt">Comments (position, other financial information, etc.)</p><% InputBlockElementLeftEnd %><% InputBlockElementMiddle %>
		<% InputBlockElementRightStart %>&nbsp;&nbsp;<textarea name="expert_comments" size="31" rows="5" style="width:355px;"><% =objExpertProject.Comments %></textarea><% InputBlockElementRightEnd %>
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

