<%
'--------------------------------------------------------------------
'
' List of experts in the database
'
'--------------------------------------------------------------------
%>
<!--#include file="../_data/datMonth.asp"-->
<!--#include virtual="/_common/_class/main.asp"-->
<!--#include virtual="/_common/_class/expert.asp"-->
<!--#include virtual="/_common/_class/status_cv.asp"-->
<!--#include virtual="/_common/_class/expert.status_cv.asp"-->
<!--#include virtual="/_common/_class/document.asp"-->
<%
sTempParams=sParams
sTempParams=ReplaceUrlParams(sTempParams, "act=" & sAction)

' Remove inactive url params
sParams=ReplaceUrlParams(sParams, "srch")
sParams=ReplaceUrlParams(sParams, "ord")
sParams=ReplaceUrlParams(sParams, "id")

' Check UserID
CheckUserLogin sScriptFullNameAsParams

Dim iTotalExpertsNumber
Dim iTotalPages, iTotalRecords, iCurrentPage, iCurrentRow, sRowColor, iSearchQueryID, bSaveSearchLog, j
Dim sOrderBy, sSearchString

sOrderBy=UCase(Request.QueryString("ord"))
If sOrderBy<>"E" And sOrderBy<>"R" And sOrderBy<>"I" And sOrderBy<>"U" And sOrderBy<>"B" Then sOrderBy="A"

Dim sLastExperienceMonthFrom, sLastExperienceYearFrom, sLastExperienceMonthTo, sLastExperienceYearTo
Dim sCvModifiedMonthFrom, sCvModifiedYearFrom, sCvModifiedMonthTo, sCvModifiedYearTo

sLastExperienceMonthFrom=CheckInt(Request.QueryString("last_experience_from_month"))
sLastExperienceYearFrom=CheckInt(Request.QueryString("last_experience_from_year"))
sLastExperienceMonthTo=CheckInt(Request.QueryString("last_experience_to_month"))
sLastExperienceYearTo=CheckInt(Request.QueryString("last_experience_to_year"))

sCvModifiedMonthFrom=CheckInt(Request.QueryString("modified_from_month"))
sCvModifiedYearFrom=CheckInt(Request.QueryString("modified_from_year"))
sCvModifiedMonthTo=CheckInt(Request.QueryString("modified_to_month"))
sCvModifiedYearTo=CheckInt(Request.QueryString("modified_to_year"))

sSearchString=Request.QueryString("srch")

Response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=cv_list.xls"
%>
<html>
<head>
<title>List of experts</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<body bgcolor="#FFFFFF" topmargin=0 leftmargin=0 marginheight=0 marginwidth=0>

<%
iCurrentPage=Request.QueryString("page")
If Not IsNumeric(iCurrentPage) or iCurrentPage="" Then
	iCurrentPage=1
Else
	iCurrentPage=CInt(iCurrentPage)
End If

Dim iShowRemoved
If sAction="all" Or sAction="deleted" Then 
	iShowRemoved=1
Else
	iShowRemoved=0
End If

Set objTempRs=GetDataRecordsetSP("usp_AdmExpAllListExtraModifiedSelect", Array( _
	Array(, adVarChar, 250, Null), _
	Array(, adInteger, , 0), _
	Array(, adInteger, , iShowRemoved), _
	Array(, adVarChar, 100, sAction), _
	Array(, adVarChar, 255, sSearchString), _
	Array(, adVarChar, 100, sOrderBy), _
	Array(, adVarChar, 16, ConvertDMYForSql(sLastExperienceYearFrom, sLastExperienceMonthFrom, 1)), _
	Array(, adVarChar, 16, ConvertDMYForSql(sLastExperienceYearTo, sLastExperienceMonthTo, 31)), _
	Array(, adVarChar, 16, ConvertDMYForSql(sCvModifiedYearFrom, sCvModifiedMonthFrom, 1)), _
	Array(, adVarChar, 16, ConvertDMYForSql(sCvModifiedYearTo, sCvModifiedMonthTo, 31)) _
	))
	
iTotalExpertsNumber=objTempRs.RecordCount
Dim sExpertFullName, sExpertEmail

If Not objTempRs.Eof Then
	%>

	<table width="96%" cellpadding="0" cellspacing="0" border="0" align="center">
	<tr><td bgcolor="#003399">
	<table width="100%" cellpadding="2" cellspacing="1" border="0">
	<tr bgcolor="#E0F3FF">
	<td width="40" align="center"><p class="sml"><b>Expert&nbsp;ID</b></p></td>
	<td width="50"><p class="sml"><b>Title</b></td>
	<td width="15%"><p class="sml"><b>Last name</b></td>
	<td width="15%"><p class="sml"><b>First name</b></td>
	<td width="50"><p class="sml"><b>Registered</b></td>
	<td width="50"><p class="sml"><b>Modified</b></td>
	<td width="50"><p class="sml"><b>Last&nbsp;experience</b></p></td>
	<td><p class="sml">Email</b></td>
	<td><p class="sml">Phone</b></td>
	<td width="20"><p class="sml">Status</b></td>
	<% If bCvMultiLanguageActive = cMultiLanguageEnabled Then %>
		<td width="20"><p class="sml">CV&nbsp;language</b></td>
	<% End If %>
	<% If bCvTypeActive = cCvTypeEnabled Then %>
		<td width="100"><p class="sml">Type</b></td>
	<% End If %>
	<td width="100"><p class="sml">Comments</b></td>
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
		<td align="center"><p class="sml"><% =objTempRs("id_Expert") %></td>
		<td><p class="mt"><% =objTempRs("ptlName") %></a></b></p></td>
		<td><p class="mt"><% =objTempRs("psnLastName") %></a></b></p></td>
		<td><p class="mt"><% =objTempRs("psnFirstName") %></a></b></p></td>
		<td><p class="sml"><% =ConvertDateForText(objTempRs("expCreateDate"), "&nbsp;", "DDMMYYYY") %>&nbsp;</td>
		<td><p class="sml"><% =ConvertDateForText(objTempRs("expLastUpdate"), "&nbsp;", "DDMMYYYY") %>&nbsp;</td>
		<td><p class="sml"><% =ConvertDateForText(objTempRs("wkeEndDate"), "&nbsp;", "MonthYear") %>&nbsp;</td>
<!--
		<td><p class="sml"><% =ConvertDateForText(objTempRs("psnBirthDate"), "&nbsp;", "DayMonthYear") %>&nbsp;</td>
-->
		<td><p class="sml"><% =sExpertEmail %></td>
		<td><p class="sml"> <% =objTempRs("Phone") %></td>

	<% If iShowRemoved=0 Then
		' Showing a status
		Dim objExpertStatusCV
		Set objExpertStatusCV = New CExpertStatusCV
		objExpertStatusCV.Expert.ID=objTempRs("id_Expert")
		objExpertStatusCV.LoadData

		Response.Write "<td><p class=""sml"">" 
		If IsObject(objExpertStatusCV.Status) Then
			Response.Write objExpertStatusCV.Status.Name
		End If

		Response.Write "</p></td>" 
		%>
	<% If bCvMultiLanguageActive = cMultiLanguageEnabled Then %>
		<td><p class="sml"><% If Len(objTempRs("Lng"))>1 Then %><% =dictLanguage.Item(Trim(objTempRs("Lng"))) %><% Else %><% =objTempRs("Lng") %><% End If %></td>
	<% End If %>
	<% If bCvTypeActive = cCvTypeEnabled Then %>
			<td><p class="sml"><% =objTempRs("KgCvFile") %></td>
	<% End If %>
		<td><p class="sml"><% =objTempRs("expComments") %></td>
	<%
	End If

		Response.Write "</tr>"
		objTempRs.MoveNext
	WEnd
End If
objTempRs.Close
Set objTempRs=Nothing
%>
</table>
</td></tr>
</table>

</body>
</html>
