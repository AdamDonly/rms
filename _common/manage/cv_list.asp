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

sSearchString=Request.QueryString("srch")
%>
<html>
<head>
<title>List of experts</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" type="text/css" href="<% =sHomePath %>styles.css">
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

Dim iShowRemoved
If sAction="all" Or sAction="deleted" Then 
	iShowRemoved=1
Else
	iShowRemoved=0
End If

Set objTempRs=GetDataRecordsetSP("usp_AdmExpAllListSelect", Array( _
	Array(, adVarChar, 250, Null), _
	Array(, adInteger, , 1), _
	Array(, adInteger, , iShowRemoved), _
	Array(, adVarChar, 100, sAction), _
	Array(, adVarChar, 255, sSearchString), _
	Array(, adVarChar, 100, sOrderBy)))

iTotalExpertsNumber=objTempRs.RecordCount
Dim sExpertFullName, sExpertEmail

If Not objTempRs.Eof Then
	iCurrentRow=0
	objTempRs.PageSize=50
	iTotalRecords=objTempRs.RecordCount
	iTotalPages=objTempRs.PageCount
	objTempRs.AbsolutePage=CInt(iCurrentPage)
	ShowNavigationPages iCurrentPage, iTotalPages, sParams
	%>

	<table width="96%" cellpadding="0" cellspacing="0" border="0" align="center">
	<tr><td bgcolor="#003399">
	<table width="100%" cellpadding="2" cellspacing="1" border="0">
	<tr bgcolor="#FFFFFF">
	<form method="get" action="<%=sScriptFileName & sTempParams%>">
	<td colspan="10">
	<input type="hidden" name="act" value=<%=sAction%>>
	<p class="mt"><b>Search for experts using ID, First name, Surname or Email</b><br>
	<input type="text" name="srch" size=34" value="<% =sSearchString %>"> &nbsp; 
	<input type="submit" value="Search" >&nbsp;
	<input type="button" value=" Reset" onClick="javascript:window.location.href='<%=sScriptFileName & ReplaceUrlParams("","act=" & sAction) %>'">
	</p></td>
	</form>
	</tr>
	<tr bgcolor="#E0F3FF">
	<td <%If sOrderBy="I" Then Response.Write " bgcolor=""#99CCFF""" %> width="40" align="center"><p class="sml"><b><% If sOrderBy<>"I" Then %><a href="<%=sScriptFileName & ReplaceUrlParams(sTempParams, "ord=I") %>"><% End If %>Expert&nbsp;ID</b></p></td>
	<td width="30%" <%If sOrderBy="A" Then Response.Write " bgcolor=""#99CCFF""" %>><p class="sml"><b><% If sOrderBy<>"A" Then %><a href="<%=sScriptFileName & ReplaceUrlParams(sTempParams, "ord=A") %>"><% End If %>Name, FirstName MiddleName (Title)</a></b></td>
	<td width="50" <%If sOrderBy="R" Then Response.Write " bgcolor=""#99CCFF""" %>><p class="sml"><b><% If sOrderBy<>"R" Then %><a href="<%=sScriptFileName & ReplaceUrlParams(sTempParams, "ord=R") %>"><% End If %>Registered</b></td>
	<td width="50" <%If sOrderBy="U" Then Response.Write " bgcolor=""#99CCFF""" %>><p class="sml"><b><% If sOrderBy<>"U" Then %><a href="<%=sScriptFileName & ReplaceUrlParams(sTempParams, "ord=U") %>"><% End If %>Modified</b></td>
	<td width="50" <%If sOrderBy="E" Then Response.Write " bgcolor=""#99CCFF""" %>><p class="sml"><b><% If sOrderBy<>"E" Then %><a href="<%=sScriptFileName & ReplaceUrlParams(sTempParams, "ord=E") %>"><% End If %>Last&nbsp;experience</b></p></td>
<!--	
	<td width="50" <%If sOrderBy="B" Then Response.Write " bgcolor=""#99CCFF""" %>><p class="sml"><b><% If sOrderBy<>"B" Then %><a href="<%=sScriptFileName & ReplaceUrlParams(sTempParams, "ord=B") %>"><% End If %>Birthday</b></p></td>
-->	
	<td><p class="sml" width="20%">Email</b></td>
	<td width="20"><p class="sml">Status</b></td>
	<td width="30%"><p class="sml">Comments</b></td>
	<% If iShowRemoved=1 Then %>
	<td width="20"><p class="sml">Options</b></td>
	<% End If %>
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
		<td><p class="mt"><a href="../register/register6.asp?id=<% =objTempRs("id_Expert") %>" target=_blank><% =sExpertFullName %></a></b></p></td>
		<td><p class="sml"><% =ConvertDateForText(objTempRs("expCreateDate"), "&nbsp;", "DDMMYYYY") %>&nbsp;</td>
		<td><p class="sml"><% =ConvertDateForText(objTempRs("expLastUpdate"), "&nbsp;", "DDMMYYYY") %>&nbsp;</td>
		<td><p class="sml"><% =ConvertDateForText(objTempRs("wkeEndDate"), "&nbsp;", "MonthYear") %>&nbsp;</td>
<!--
		<td><p class="sml"><% =ConvertDateForText(objTempRs("psnBirthDate"), "&nbsp;", "DayMonthYear") %>&nbsp;</td>
-->
		<td><p class="sml"><% =sExpertEmail %></td>

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
		<td><p class="sml"><a href="../register/comments.asp?id=<% =objTempRs("id_Expert") %>"><img src="<% =sHomePath %>image/vn_updt.gif" width="15" height="15" align="left" hspace="6" vspace="0" border="0" alt="Edit comments for <% =objTempRs("psnLastName") & ", " & objTempRs("psnFirstName") & " " & objTempRs("psnMiddleName") & " (" & objTempRs("ptlName") & ")" %>"></a><% =objTempRs("expComments") %></td>
		<%
	ElseIf iShowRemoved=1 Then
		%>
		<td><p class="sml fcmp">
		<% If objTempRs("expRemoved")=True Then %>
			Deleted<br /><% =ConvertDateForText(objTempRs("expRemovedDate"), "&nbsp;", "DDMMYYYY") %>
		<% End If %>
		<% If objTempRs("expDeleted")=True Then %>
			Deleted<br /><% =ConvertDateForText(objTempRs("expDeletedDate"), "&nbsp;", "DDMMYYYY") %>
		<% End If %>
		</p></td>
		<td><p class="sml fcmp">
		<% If objTempRs("expRemoved")=True Then %>
			<% =objTempRs("expRemovedComments") %>
		<% End If %>
		<% If objTempRs("expDeleted")=True Then %>
			<% =objTempRs("expDeletedComments") %>
		<% End If %>
		</p></td>
		<td><p class="sml">
		<% If objTempRs("expRemoved")=True Or objTempRs("expDeleted")=True Then %>
			<a href="cv_restore.asp?id=<% =objTempRs("id_Expert") %>">Restore&nbsp;CV</a>
		<% End If %>
		</p></td>
		<%
	End If

		Response.Write "</tr>"
		iCurrentRow=iCurrentRow+1
		objTempRs.MoveNext
	WEnd
End If
objTempRs.Close
Set objTempRs=Nothing
%>
<tr bgcolor="#FFFFFF"><td colspan="10"><p class="mt">Total: <b><%=ShowEntityPlural(iTotalExpertsNumber, "expert", "experts", "&nbsp;") %></b><% If Len(sSearchString)>0 Then Response.Write " matching search criteria"%></p></td></tr>
</table>
</td></tr>
</table>

<% ShowNavigationPages iCurrentPage, iTotalPages, sParams %>

</body>
</html>
