<% 
'--------------------------------------------------------------------
'
' Project details
'
'--------------------------------------------------------------------
%>
<%
' Check user's access rights
If sApplicationName<>"expert" Then
	CheckUserLogin sScriptFullNameAsParams
Else 
	Response.Redirect sApplicationHomePath
End If
%>
<!--#include file="../_data/datMonth.asp"-->
<!--#include virtual="/_common/_class/main.asp"-->
<!--#include virtual="/_common/_class/project.asp"-->
<!--#include virtual="/_common/_class/expert.asp"-->
<!--#include virtual="/_common/_class/status_cv.asp"-->
<!--#include virtual="/_common/_class/expert.status_cv.asp"-->
<%
Dim sCVFormat, bCvValidForMemberOrExpert

sParams=ReplaceUrlParams(sParams, "ord")

Dim iProjectID
iProjectID=CheckIntegerAndZero(Request.QueryString("idproject"))

Dim objProject
Set objProject = New CProject
objProject.ID=iProjectID
objProject.LoadData
%>
<html>
<head>
<title>Project details</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" type="text/css" href="<% =sHomePath %>styles.css">
<script type="text/javascript">
<!-- 
function deleteProject(idproject, total_experts) {
	if (!idproject) {
		return;
	}
	var m; // message

	if (total_experts && total_experts>0) {
		m = 'There are some experts registered on this project.\nPlease delete all linked experts first.'
		alert(m);
		return;
	}
	
	m = 'Are you sure you want to permanently delete this project?'
	if (confirm(m)) {
		var s = "register_save.asp<% =ReplaceUrlParams(sParams, "act=delete") %>"
		self.location.replace(s);
	}
}

function deleteExpertFromProject(idexpert, idproject) {
	if ((!idexpert) || (!idproject)) {
		return;
	}
	
	m = 'Are you sure you want to permanently delete this expert for the project?'
	if (confirm(m)) {
		var s = "link_expert_save.asp<% =ReplaceUrlParams(ReplaceUrlParams(sParams, "act=delete"), "idexpert") %>&idexpert=" + idexpert
		self.location.replace(s);
	}
}

-->
</script>
</head>
<body bgcolor="#FFFFFF" topmargin=0 leftmargin=0  marginheight=0 marginwidth=0>
<% ShowTopMenu %>

<br>
<table width="98%" cellpadding=0 cellspacing=0 border=0 align="center">
<tr><td width="85%" valign="top">
<%
' Project details
	WriteDataTitle "PROJECT DETAILS"
	ShowExpertsBlockSubTitle "98%", "52", "ex0"
	ShowUserNoticesViewHeader "100%", 180
	ShowUserNoticesViewSpacer 4

	ShowUserNoticesViewText "</b>Title", "<b>" & objProject.Title & "</b>"
	ShowUserNoticesViewText "</b>Reference", objProject.Reference
	ShowUserNoticesViewText "</b>Short name", objProject.Name
	ShowUserNoticesViewText "</b>Status", objProject.Status.Name
	ShowUserNoticesViewText "</b>Country / Region", objProject.Location
	ShowUserNoticesViewText "</b>Deadline", ConvertDateForText(objProject.Deadline, "&nbsp;", "DDMMYYYY")
	ShowUserNoticesViewText "</b>Description", objProject.Description

	ShowUserNoticesViewSpacer 5
	ShowUserNoticesViewFooter
	ShowExpertsBlockFooter "98%", "52", "ex0"
	
' Experts
	WriteDataTitle "EXPERTS"
	
Dim iTotalExpertsNumber
Dim iTotalPages, iTotalRecords, iCurrentPage, iCurrentRow, sRowColor, iSearchQueryID, bSaveSearchLog, j
Dim sOrderBy, sSearchString

sOrderBy=UCase(Request.QueryString("ord"))
If sOrderBy<>"E" And sOrderBy<>"A" And sOrderBy<>"R" And sOrderBy<>"I" And sOrderBy<>"U" And sOrderBy<>"B" Then sOrderBy="S"


iCurrentPage=Request.QueryString("page")
If Not IsNumeric(iCurrentPage) or iCurrentPage="" Then
	iCurrentPage=1
Else
	iCurrentPage=CInt(iCurrentPage)
End If

sSearchString=Request.QueryString("srch")

Set objTempRs=GetDataRecordsetSP("usp_ProjectExpertListSelect", Array( _
	Array(, adInteger, , iProjectID), _
	Array(, adVarChar, 100, sAction), _
	Array(, adVarChar, 255, sSearchString), _
	Array(, adVarChar, 100, sOrderBy)))

iTotalExpertsNumber=objTempRs.RecordCount
Dim sExpertFullName, sExpertEmail
%>

<%
If Not objTempRs.Eof Then
	iCurrentRow=0
	objTempRs.PageSize=50
	iTotalRecords=objTempRs.RecordCount
	iTotalPages=objTempRs.PageCount
	objTempRs.AbsolutePage=CInt(iCurrentPage)
	ShowNavigationPages iCurrentPage, iTotalPages, sParams
End If
%>
	
	<table width="98%" cellpadding="0" cellspacing="0" border="0" align="center">
	<tr><td bgcolor="#003399">
	<table width="100%" cellpadding="2" cellspacing="1" border="0">
	<tr bgcolor="#E0F3FF">
	<td <%If sOrderBy="A" Then Response.Write " bgcolor=""#99CCFF""" %> width="40%"><p class="sml"><b><% If sOrderBy<>"A" Then %><a href="<%=sScriptFileName & ReplaceUrlParams(sParams, "ord=A") %>"><% End If %>Name, FirstName MiddleName (Title)</a></b></td>
	<td <%If sOrderBy="R" Then Response.Write " bgcolor=""#99CCFF""" %> width="50"><p class="sml"><b><% If sOrderBy<>"R" Then %><a href="<%=sScriptFileName & ReplaceUrlParams(sParams, "ord=R") %>"><% End If %>Registered</a></b></td>
	<td <%If sOrderBy="S" Then Response.Write " bgcolor=""#99CCFF""" %> width="140"><p class="sml"><b><% If sOrderBy<>"S" Then %><a href="<%=sScriptFileName & ReplaceUrlParams(sParams, "ord=S") %>"><% End If %>Status</a></b></td>
	<td width="50"><p class="sml"><b>Fee</b></td>
	<td width="200"><p class="sml"><b>Comments</b></td>
	<td width="20"><p class="sml"><b>Modify</b></td>
	<td width="20"><p class="sml"><b>Delete</b></td>
	</tr>

	<% While Not objTempRs.Eof And iCurrentRow<objTempRs.PageSize %>
		<% 
		sExpertFullName = objTempRs("psnLastName") & ", " & objTempRs("psnFirstName") & " " & objTempRs("psnMiddleName") & " (" & objTempRs("ptlName") & ")"
		sExpertEmail=objTempRs("Email")
		If sApplicationName="external" Then
			If sContactDetailsExternally=cNameObfuscated Then
				sExpertFullName=ObfuscateString(objTempRs("psnLastName")) & ", " & ObfuscateString(objTempRs("psnFirstName")) & " " & ObfuscateString(objTempRs("psnMiddleName")) & " (" & objTempRs("ptlName") & ")"
				sExpertEmail=ObfuscateEmail(objTempRs("Email"))
			End If
			If sContactDetailsExternally=cNameHidden Then
				sExpertFullName=""
				sExpertEmail=""
			End If
		End If
		%>
		<tr bgcolor="#FFFFFF">
		<td><p class="mt"><a href="../view/cv_view.asp?id=<% =objTempRs("id_Expert") %>" target=_blank><% =sExpertFullName %></a></b></p></td>
		<td><p class="sml"><% =ConvertDateForText(objTempRs("epjCreateDate"), "&nbsp;", "DDMMYYYY") %>&nbsp;</td>
		<td><p class="sml"><% =objTempRs("exsTitle") %></p></td>
		<td><p class="sml"><% =objTempRs("epjFee") %><% If Not IsNull(objTempRs("epjFee")) Then %>&nbsp;<% =objTempRs("epjFeeCurrency") %><% End If %></p></td>
		<td><p class="sml"><% =objTempRs("epjComments") %></td>
		<td align="center"><a href="link_expert.asp<% =AddUrlParams(sParams, "idexpert=" & objTempRs("id_Expert")) %>"><img src="<% =sHomePath %>image/vn_updt.gif" width="15" height="15" border="0" alt="Modify"></a></td>
		<td align="center"><a href="javascript:deleteExpertFromProject(<% =objTempRs("id_Expert") %>, <% =iProjectID %>)"><img src="<% =sHomePath %>image/vn_del.gif" width="15" height="15" border="0" alt="Delete"></a></td>
		</tr>
		<%
		iCurrentRow=iCurrentRow+1
		objTempRs.MoveNext
	WEnd
	%>
	
	<tr bgcolor="#FFFFFF"><td colspan="10"><p class="mt">
	<% If iTotalExpertsNumber>0 Then %>
		Total: <b><%=ShowEntityPlural(iTotalExpertsNumber, "expert", "experts", "&nbsp;") %></b><% If Len(sSearchString)>0 Then Response.Write " matching search criteria"%>
	<% Else %>
		There are no experts provided on this project.
	<% End If %>
	</p></td></tr>
	</table>
	</td></tr>
	</table>
<%
objTempRs.Close
Set objTempRs=Nothing
%>
	
</td>
<td width="5%">&nbsp;&nbsp;</td>
<td width="20%" valign="top">
	<!-- Feature boxes -->
	<img src="<% =sHomePath %>image/x.gif" width=1 height=23><br />

	<% ShowFeatureBoxHeader("Project options") %>
	<p class="sml" style="padding: 2px 5px;">Some information about the project is missing?</p>
	<div align="center"><a href="../project/register.asp<% =ReplaceUrlParams(sParams, "idproject=" & iProjectID) %>"><img src="<% =sHomePath %>image/bte_updateproject152.gif" vspace="4" border="0"></a></div>
	<% ShowFeatureBoxDelimiter %>
	<p class="sml" style="padding: 2px 5px;">To remove the project without experts from the system</p>
	<div align="center"><a href="javascript:deleteProject(<% =iProjectID %>, <% =iTotalExpertsNumber %>)"><img src="<% =sHomePath %>image/bte_deleteproject152.gif" vspace="4" border="0"></a></div>
	<% ShowFeatureBoxFooter %>
	<br>
	
	<% ShowFeatureBoxHeader("Add experts") %>
	<p class="sml" style="padding: 2px 5px;">In order to add experts on this project</p>
	<div align="center"><a href="../search/exp_search.asp<% =ReplaceUrlParams(sParams, "idproject=" & iProjectID) %>"><img src="<% =sHomePath %>image/bte_searchexp152.gif" vspace="4" border="0"></a></div>
	<% ShowFeatureBoxFooter %>
	<br>
	
	
</td>
</tr>

<% CloseDBConnection %>
</body>
</html>
