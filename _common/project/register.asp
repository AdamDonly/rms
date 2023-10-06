<%
'--------------------------------------------------------------------
'
' Project registration.
'
'--------------------------------------------------------------------
Response.Expires = 0
Response.ExpiresAbsolute = Now() - 1
Response.CacheControl = "no-cache" 
%>
<!--#include file="../_data/datMonth.asp"-->
<!--#include file="../_data/datProjectStatus.asp"-->
<!--#include virtual="/_common/_class/main.asp"-->
<!--#include virtual="/_common/_class/project.asp"-->
<%
If sApplicationName<>"expert" Then
	CheckUserLogin sScriptFullNameAsParams
Else 
	Response.Redirect sApplicationHomePath
End If

Dim iProjectID
iProjectID=CheckIntegerAndZero(Request.QueryString("idproject"))

Dim objProject
Set objProject = New CProject
objProject.ID=iProjectID
objProject.LoadData

If Not objProject.Status.Name>"" Then
	Dim sSuggestedProjectStatus, iSuggestedProjectStatusID
	sSuggestedProjectStatus=Request.QueryString("project_type")
	If sSuggestedProjectStatus="tendering" Then
		iSuggestedProjectStatusID=121
	ElseIf sSuggestedProjectStatus="running" Then
		iSuggestedProjectStatusID=202
	ElseIf sSuggestedProjectStatus="closed" Then
		iSuggestedProjectStatusID=301
	Else
		iSuggestedProjectStatusID=0
	End If
End If

sParams=ReplaceUrlParams(sParams, "project_type")
%>
<html>
<head>
<title><% =GetLabel(sInterfaceLanguage, "Project registration") %></title>
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
	<% InputFormHeader 580, GetLabel(sInterfaceLanguage, "PROJECT REGISTRATION") %>
	<% InputBlockHeader "100%" %>
	<form action="register_save.asp<% =AddUrlParams(sParams, "act=" & sAction) %>"  method="post" onSubmit="validateForm(); return false;">
	<input type="hidden" name="idproject" value="<% =objProject.ID %>">
		<% InputBlockSpace 4 %>
		<% InputBlockElementLeftStart %><p class="ftxt"><% =GetLabel(sInterfaceLanguage, "Project title") %></p><% InputBlockElementLeftEnd %><% InputBlockElementMiddle %>
		<% InputBlockElementRightStart %>&nbsp;&nbsp;<input type="text" name="project_title" size="31" style="width:355px;" maxlength="100" value="<% =objProject.Title %>">&nbsp;&nbsp;<span class="fcmp">*</span>&nbsp;<% InputBlockElementRightEnd %>
		<% InputBlockElementLeftStart %><p class="ftxt"><% =GetLabel(sInterfaceLanguage, "Project status") %></p><% InputBlockElementLeftEnd %><% InputBlockElementMiddle %>
		<% InputBlockElementRightStart %>&nbsp;&nbsp;<select size="1" name="project_status" style="width:355px;">
		<option value="0"> </option>
		<%  For i=LBound(arrProjectStatusID) to UBound(arrProjectStatusID)
		If objProject.Status.ID=arrProjectStatusID(i) Or iSuggestedProjectStatusID=arrProjectStatusID(i) Then
			Response.Write("<option value=" & arrProjectStatusID(i) & " selected>" & arrProjectStatusTitle(i) & "</option>")
		Else 
			Response.Write("<option value=" & arrProjectStatusID(i) & ">" & arrProjectStatusTitle(i) & "</option>")
		End If
		Next %>
		</select>&nbsp;&nbsp;<span class="fcmp">*</span>&nbsp;<% InputBlockElementRightEnd %>
		<% InputBlockElementLeftStart %><p class="ftxt"><% =GetLabel(sInterfaceLanguage, "Reference") %></p><% InputBlockElementLeftEnd %><% InputBlockElementMiddle %>
		<% InputBlockElementRightStart %>&nbsp;&nbsp;<input type="text" name="project_reference" size="31" style="width:355px;" maxlength="100" value="<% =objProject.Reference %>"><% InputBlockElementRightEnd %>
		<% InputBlockElementLeftStart %><p class="ftxt"><% =GetLabel(sInterfaceLanguage, "Country / Region") %></p><% InputBlockElementLeftEnd %><% InputBlockElementMiddle %>
		<% InputBlockElementRightStart %>&nbsp;&nbsp;<input type="text" name="project_country" size="31" style="width:355px;" maxlength="100" value="<% =objProject.Location %>"><% InputBlockElementRightEnd %>
		<% InputBlockElementLeftStart %><p class="ftxt"><% =GetLabel(sInterfaceLanguage, "Description") %></p><% InputBlockElementLeftEnd %><% InputBlockElementMiddle %>
		<% InputBlockElementRightStart %>&nbsp;&nbsp;<textarea name="project_description" size="31" rows="4" style="width:355px;"><% =objProject.Description %></textarea><% InputBlockElementRightEnd %>
		<% InputBlockElementLeftStart %><p class="ftxt"><% =GetLabel(sInterfaceLanguage, "Deadline") %></p><% InputBlockElementLeftEnd %><% InputBlockElementMiddle %>
		<% InputBlockElementRightStart %>&nbsp;&nbsp;<select name="project_deadline_day" size="1">
		<option value="0"><% =GetLabel(sInterfaceLanguage, "Day") %></option>
		<% 
		Dim iProjectDeadlineDay
		If IsDate(objProject.Deadline) Then iProjectDeadlineDay=Day(objProject.Deadline)
		For i=1 to 31 
			If iProjectDeadlineDay=i Then 
				Response.Write("<option value=" & i & " selected>" & i & "</option>")
			Else
				Response.Write("<option value=" & i & ">" & i & "</option>")
			End If			
		Next %>
		</select>
		<select name="project_deadline_month" size=1>
		<option value="0" selected><% =GetLabel(sInterfaceLanguage, "Month") %></option>
		<% 
		Dim iProjectDeadlineMonth
		If IsDate(objProject.Deadline) Then iProjectDeadlineMonth=Month(objProject.Deadline)
		For i=1 to UBound(arrMonthID)
			If iProjectDeadlineMonth=i Then
				Response.Write("<option value=" & arrMonthID(i) &" selected>"& arrMonthName(i) &"</option>")
			Else
				Response.Write("<option value=" & arrMonthID(i) &">"& arrMonthName(i) &"</option>")
			End If
		Next %>
		</select>
		<select name="project_deadline_year" size="1">
		<option value="0"><% =GetLabel(sInterfaceLanguage, "Year") %></option>
		<% Dim iCurrentYear
		iCurrentYear=Year(Date)
		Dim iProjectDeadlineYear
		If IsDate(objProject.Deadline) Then iProjectDeadlineYear=Year(objProject.Deadline)
		
		For i=-4 to 3
			If iProjectDeadlineYear=iCurrentYear+i Then
				Response.Write("<option value=" & (iCurrentYear+i) & " selected>"& (iCurrentYear+i) & "</option>")
			Else
				Response.Write("<option value=" & (iCurrentYear+i) & ">"& (iCurrentYear+i) & "</option>")
			End If
		Next %>
		</select>&nbsp;<% InputBlockElementRightEnd %>
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

