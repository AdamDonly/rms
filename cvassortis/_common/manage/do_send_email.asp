<%
'--------------------------------------------------------------------
'
' List of experts in the database
'
'--------------------------------------------------------------------
%>
<!--#include virtual="/_common/_template/asp.header.nocache.asp"-->
<!--#include virtual="/_common/_template/asp.header.notimeout.asp"-->

<!--#include file="../_data/datMonth.asp"-->
<!--#include virtual="/_common/_class/main.asp"-->
<!--#include virtual="/_common/_class/expert.asp"-->
<!--#include virtual="/_common/_class/status_cv.asp"-->
<!--#include virtual="/_common/_class/expert.status_cv.asp"-->
<% 
CheckUserLogin sScriptFullName

Dim iTotalExpertsNumber
Dim iTotalPages, iTotalRecords, iCurrentPage, iCurrentRow, sRowColor, iSearchQueryID, bSaveSearchLog, j
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
%>

<%
If iSearchQueryID>0 Then
Set objTempRs=GetDataRecordsetSP("usp_MmbExpListQuerySelect", Array( _
		Array(, adInteger, , iMemberID), _
		Array(, adInteger, , Null), _
		Array(, adInteger, , iSearchQueryID)))	
Else
Set objTempRs=GetDataRecordsetSP("usp_ExpertListSelect", Array( _
	Array(, adInteger, , objUserCompanyDB.ID), _
	Array(, adVarChar, 100, Null), _
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

Dim sSender, sRecipient, sSubject, sBody
Dim sSenderActive, sRecipientActive, sSubjectActive, sBodyActive
sSender=Request.Form("message_from")
sSenderActive=sSender
sRecipient=Request.Form("message_to")
sRecipientActive=sRecipient
sSubject=Request.Form("message_subject")
sSubjectActive=sSubject
sBody=Request.Form("message_body")
sBody=ReadTextFile("\_mails\_SystemHeader.htm") + "<p>" + ConvertTextForEmail(sBody) + "</p>" + ReadTextFile("\_mails\_SystemFooter.htm")
sBodyActive=sBody

%>

<html>
<head>
<title><%=sApplicationName%>. Send emails to experts</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" type="text/css" href="/old/styles.css">
</head>
<body bgcolor="#FFFFFF" topmargin=0 leftmargin=0 marginheight=0 marginwidth=0>
<% ShowTopMenu %>

<br>
<table width="96%" cellpadding="0" cellspacing="0" border="0" align="center">
<tr><td bgcolor="#003399">
<table width="100%" cellpadding="2" cellspacing="1" border="0">
<tr bgcolor="#FFFFFF"><td colspan="8"><br>

<%
If Request.Form()>"" Then
While Not objTempRs.Eof 
' Recipient
	If Not InStr(ReplaceIfEmpty(sRecipient, ""), "@")>0 Then sRecipientActive=objTempRs("Email")
	
' Subject
	If InStr(sSubject, "[") Then
		sSubjectActive=sSubject
		sSubjectActive=Replace(sSubjectActive, "[name]", Trim(objTempRs("ptlName") & " " & objTempRs("psnFirstName") & " " & objTempRs("psnLastName")), 1, -1, vbTextCompare)
	End If
' Body
	If InStr(sBody, "[") Then
		sBodyActive=Trim(sBody)
		sBodyActive=Replace(sBodyActive, "[name]", Trim(objTempRs("ptlName") & " " & objTempRs("psnFirstName") & " " & objTempRs("psnLastName")), 1, -1, vbTextCompare)
	End If
	
	If sRecipientActive>"" And InStr(sRecipientActive, "@")>0 Then
		Response.Write Server.HtmlEncode("Email from " & sSenderActive & " to " & sRecipientActive & " Subject:" & sSubjectActive)
		On Error Resume Next
		SendEmail sSenderActive, sRecipientActive, sSubjectActive, sBodyActive, ""
		
		Response.Write " - OK<br>"
		On Error GoTo 0
		Response.Flush
	End If
	
	objTempRs.MoveNext
WEnd
%>
<script>
	<%
	sParams=AddUrlParams(sParams, "act=" & sAction)
	sParams=AddUrlParams(sParams, "done=1")
	%>
	window.location.href='<%= sScriptFileName & sParams %>';
</script>
<%
End If

objTempRs.Close
Set objTempRs=Nothing

If Request.QueryString("done")="1" Then Response.Write "<p align=""center"">Done. " & ShowEntityPlural(iTotalExpertsNumber, "email was", "emails were", " ") & " sent.</p>"
%>
<br>
</td></tr>
</table>
</td></tr>
</table>


</body>
</html>

