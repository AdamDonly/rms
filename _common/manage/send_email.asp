<%
'--------------------------------------------------------------------
'
' List of experts in the database
'
'--------------------------------------------------------------------
%>
<!--#include virtual="/_common/_template/asp.header.notimeout.asp"-->
<!--#include file="../_data/datMonth.asp"-->
<!--#include virtual="/_common/_class/main.asp"-->
<!--#include virtual="/_common/_class/expert.asp"-->
<!--#include virtual="/_common/_class/status_cv.asp"-->
<!--#include virtual="/_common/_class/expert.status_cv.asp"-->
<% 
CheckUserLogin sScriptFullName

Dim iTotalExpertsNumber
Dim iTotalPages, iTotalRecords, iCurrentPage, iCurrentRow, sRowColor, iSearchQueryID, sSelect, bSaveSearchLog, j
Dim lstDuplicateIDs, arrDuplicateIDs, sDuplicates
Dim sCellStyle, sOrderBy, sSearchString
Dim sLastExperienceMonthFrom, sLastExperienceYearFrom, sLastExperienceMonthTo, sLastExperienceYearTo
Dim sCvModifiedMonthFrom, sCvModifiedYearFrom, sCvModifiedMonthTo, sCvModifiedYearTo

sOrderBy=UCase(Request.QueryString("ord"))
If sOrderBy<>"E" And sOrderBy<>"R" And sOrderBy<>"I" And sOrderBy<>"U" Then sOrderBy="A"

sSearchString=Request.QueryString("srch")

sLastExperienceMonthFrom=CheckInt(Request.QueryString("last_experience_from_month"))
sLastExperienceYearFrom=CheckInt(Request.QueryString("last_experience_from_year"))
sLastExperienceMonthTo=CheckInt(Request.QueryString("last_experience_to_month"))
sLastExperienceYearTo=CheckInt(Request.QueryString("last_experience_to_year"))

sCvModifiedMonthFrom=CheckInt(Request.QueryString("modified_from_month"))
sCvModifiedYearFrom=CheckInt(Request.QueryString("modified_from_year"))
sCvModifiedMonthTo=CheckInt(Request.QueryString("modified_to_month"))
sCvModifiedYearTo=CheckInt(Request.QueryString("modified_to_year"))

iSearchQueryID=CheckInt(Request.QueryString("qid"))
sSelect=Request.QueryString("select")
%>

<html>
<head>
<title><%=sApplicationName%>. Send emails to experts</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" type="text/css" href="<% =sHomePath %>styles.css">
<script>
	function ConfirmEmail() {
		if (confirm("Do you want to send emails to all the experts matching the selected criteria?"))
			{ document.forms[0].submit(); }
	}
</script>
</head>
<body bgcolor="#FFFFFF" topmargin=0 leftmargin=0 marginheight=0 marginwidth=0>
<% ShowTopMenu %>

<%
iCurrentPage=Request.QueryString("page")
If Not IsNumeric(iCurrentPage) or iCurrentPage="" Then
	iCurrentPage=1
Else
	iCurrentPage=CInt(iCurrentPage)
End If

If iSearchQueryID>0 And sSelect="all" Then

	UpdateMemberSearchQuery iMemberID, iSearchQueryID

	Set objTempRs=GetDataRecordsetSP("usp_MmbExpQuerySelect", Array( _
			Array(, adInteger, , iMemberID), _
			Array(, adInteger, , Null), _
			Array(, adInteger, , iSearchQueryID)))	
		
ElseIf iSearchQueryID>0 Then
	Set objTempRs=GetDataRecordsetSP("usp_MmbExpQuerySelect", Array( _
			Array(, adInteger, , iMemberID), _
			Array(, adInteger, , Null), _
			Array(, adInteger, , iSearchQueryID)))	
Else
	Set objTempRs=GetDataRecordsetSP("usp_AdmExpAllListExtraModifiedSelect", Array( _
		Array(, adVarChar, 250, Null), _
		Array(, adInteger, , 0), _
		Array(, adInteger, , 0), _
		Array(, adVarChar, 100, sAction), _
		Array(, adVarChar, 255, sSearchString), _
		Array(, adVarChar, 100, sOrderBy), _
		Array(, adVarChar, 16, ConvertDMYForSql(sLastExperienceYearFrom, sLastExperienceMonthFrom, 1)), _
		Array(, adVarChar, 16, ConvertDMYForSql(sLastExperienceYearTo, sLastExperienceMonthTo, 31)), _
		Array(, adVarChar, 16, ConvertDMYForSql(sCvModifiedYearFrom, sCvModifiedMonthFrom, 1)), _
		Array(, adVarChar, 16, ConvertDMYForSql(sCvModifiedYearTo, sCvModifiedMonthTo, 31)) _
		))
End If

iTotalExpertsNumber=objTempRs.RecordCount

If Not objTempRs.Eof Then
	iCurrentRow=0
	objTempRs.PageSize=50
	iTotalRecords=objTempRs.RecordCount
	iTotalPages=objTempRs.PageCount
	objTempRs.AbsolutePage=CInt(iCurrentPage)
	sParams=AddUrlParams(sParams, "act=" & sAction)
	%>

	<br>
	<table width="96%" cellpadding="0" cellspacing="0" border="0" align="center">
	<tr><td bgcolor="#003399">
	<table width="100%" cellpadding="2" cellspacing="1" border="0">
<%	
End If
objTempRs.Close
Set objTempRs=Nothing
%>
<tr bgcolor="#FFFFFF"><td colspan="8"><p class="mt"><b>Send emails to <%=ShowEntityPlural(iTotalExpertsNumber, "expert", "experts", "&nbsp;") %></b><% If sAction="" And InStr(sScriptFullName, "/cvassortis/") Then Response.Write " matching the selected criteria"%><br><br>
	<table width="70%" cellspacing="0" cellpadding="0" border="0" align="center">
	<form action="do_send_email.asp?<% =Request.QueryString %>" method="post" onSubmit="ConfirmEmail(); return false;">
	<tr height="29">
	<td width="15%"><p class="mt">From</p></td>
	<td width="85%">
	<select size="1" width="85%" name="message_from" id="message_from">
	<!--<option value="">&nbsp; </option>-->
	<option value="<% =sEmailClient%>"><% =sEmailClient %></option>
	<!--
	<%
	On Error Resume Next
	If Len(sEmailClientCopy)>2 Then
	%>
	<option value="<% =sEmailClientCopy%>"><% =sEmailClientCopy %></option>
	<%
	End If
	On Error GoTo 0
	%>
	-->
	</select>
	</td>
	</tr>
	<tr height="29">
	<td width="15%"><p class="mt">To</p></td>
	<td width="85%">
	<select size="1" width="85%" name="message_to" id="message_to">
<!--
	<option value="expert_all">Experts (all emails)</option>
	<option value="expert_primary">Experts (primary emails only)</option>
-->
	<option value="expert_all">Selected experts</option>
	<%
	On Error Resume Next
	If Len(sEmailClientCopy)>2 Then
	%>
	<option value="<% =sEmailClientCopy%>"><% =sEmailClientCopy %></option>
	<%
	End If
	On Error GoTo 0
	%>
	<option value="<% =sEmailClient%>"><% =sEmailClient %></option>
	</select>
	</td>
	</tr>
	<tr height="29">
	<td><p class="mt">Subject</p></td>
	<td><input type="text" size="85%" name="message_subject" id="message_subject"></td>
	</tr>
	<tr height="267">
	<td><p class="mt">Body</p></td>
	<td><textarea cols="64%" rows="16" name="message_body" id="message_body"></textarea></td>
	</tr>
	<tr height="29">
	<td>&nbsp;</td>
	<td><input type="submit" value=" &nbsp; &nbsp; Send &nbsp; &nbsp; "> &nbsp; &nbsp;
	<% If iSearchQueryID>0 Then %>
	<input type="button" value=" &nbsp; Cancel &nbsp; " onClick="javascript:window.close();"></td>
	<% Else %>
	<input type="button" value=" &nbsp; Cancel &nbsp; " onClick="javascript:window.location.href='cv_list.asp?<%= Request.QueryString %>'"></td>
	<% End If %>
	</tr>
	</form>
	</table>
	<br>
</td></tr>
</table>
</td></tr>
</table>

</body>
</html>
