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
sTempParams=sParams
sTempParams=ReplaceUrlParams(sTempParams, "act=" & sAction)

' Remove inactive url params
sParams=ReplaceUrlParams(sParams, "srch")
sParams=ReplaceUrlParams(sParams, "ord")
sParams=ReplaceUrlParams(sParams, "id")

' Check UserID
CheckUserLogin sScriptFullNameAsParams

objConn.Close
objConn.Open "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Data Source=MAGNESA;Initial Catalog=" & objUserCompanyDB.DatabasePath & ";"

Dim iTotalExpertsNumber
Dim iTotalPages, iTotalRecords, iCurrentPage, iCurrentRow, sRowColor, iSearchQueryID, bSaveSearchLog, j
Dim lstDuplicateIDs, arrDuplicateIDs, sDuplicates
Dim sCellStyle, sOrderBy, sSearchString
Dim sLastExperienceMonthFrom, sLastExperienceYearFrom, sLastExperienceMonthTo, sLastExperienceYearTo
Dim sCvModifiedMonthFrom, sCvModifiedYearFrom, sCvModifiedMonthTo, sCvModifiedYearTo

sOrderBy=UCase(Request.QueryString("ord"))
If sOrderBy<>"E" And sOrderBy<>"F" And sOrderBy<>"L" And sOrderBy<>"C" And sOrderBy<>"B" Then sOrderBy="P"

sSearchString=Request.QueryString("srch")

sLastExperienceMonthFrom=CheckInt(Request.QueryString("last_experience_from_month"))
sLastExperienceYearFrom=CheckInt(Request.QueryString("last_experience_from_year"))
sLastExperienceMonthTo=CheckInt(Request.QueryString("last_experience_to_month"))
sLastExperienceYearTo=CheckInt(Request.QueryString("last_experience_to_year"))

sCvModifiedMonthFrom=CheckInt(Request.QueryString("modified_from_month"))
sCvModifiedYearFrom=CheckInt(Request.QueryString("modified_from_year"))
sCvModifiedMonthTo=CheckInt(Request.QueryString("modified_to_month"))
sCvModifiedYearTo=CheckInt(Request.QueryString("modified_to_year"))

CheckUserLogin sScriptFullName
%>
<!--#include virtual="/_template/html.header.asp"-->
<body>
	<!-- header -->
	<!--#include virtual="/_template/page.header.asp"-->
</div>

<%
iCurrentPage=Request.QueryString("page")
If Not IsNumeric(iCurrentPage) or iCurrentPage="" Then
	iCurrentPage=1
Else
	iCurrentPage=CInt(iCurrentPage)
End If

Dim iShowProposedIbf, iShowProposedOthers
If sAction="all" Then 
	iShowProposedIbf=1
	iShowProposedOthers=1
	Response.Write "<br><p class=""ttl"">List of experts on MPIS without CV in DB (proposed by all companies)</p>"
ElseIf sAction="ibf" Then 
	iShowProposedIbf=1
	iShowProposedOthers=0
	Response.Write "<br><p class=""ttl"">List of experts on MPIS without CV in DB (proposed by IBF)</p>"
ElseIf sAction="others" Then 
	iShowProposedIbf=0
	iShowProposedOthers=1
	Response.Write "<br><p class=""ttl"">List of experts on MPIS without CV in DB (proposed by other partners)</p>"
Else
	iShowProposedIbf=0
	iShowProposedOthers=0
End If

Dim sLinkMpisContact
sLinkMpisContact="http://www.ibf.be/fwc/cgi/eproxy.exe/consult?PAGE=ContConsultFrameset.htm&TARGET=CONT&DETAIL=DEFAULT&KEY="

Set objTempRs=GetDataRecordsetSP("usp_AdmExpMpisNoCvSelect", Array( _
	Array(, adInteger, , iShowProposedIbf), _
	Array(, adInteger, , iShowProposedOthers), _
	Array(, adVarChar, 255, sSearchString)))

iTotalExpertsNumber=objTempRs.RecordCount

If Not objTempRs.Eof Then
	iCurrentRow=0
	objTempRs.PageSize=50
	iTotalRecords=objTempRs.RecordCount
	iTotalPages=objTempRs.PageCount
	objTempRs.AbsolutePage=CInt(iCurrentPage)
	sParams=AddUrlParams(sParams, "act=" & sAction)
	ShowNavigationPages iCurrentPage, iTotalPages, sParams
	%>

<div class="frame blue" align="center">
	<table class="results" style="width: 98%">
	<tr>
	<form method="get" action="<%=sScriptFileName & sParams%>">
	<td colspan="10">
		<table width="100%" cellpadding="0" cellspacing="0" border="0">
		<tr>
		<td width="380">
		<input type="hidden" name="act" value=<%=sAction%>>
		<p class="mt" style="margin: 6px, 5px;"><b>Contact ID, email, first or family name</b><br>
		<input type="text" name="srch" size="55" value="<%=sSearchString%>"> &nbsp; 
		</td>
		<td width="*">
		<br/>
		<input type="submit" value="Search" >&nbsp;
		<input type="button" value=" Reset" onClick="javascript:window.location.href='<%=sScriptFileName & "?act=" & sAction %>'">&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
		</p>
		</td>
		</tr>
		</table>
	</td>
	</form>
	</tr>
	
	<tr>
	<th <%If sOrderBy="I" Then Response.Write " bgcolor=""#99CCFF""" %> width="40" align="center"><p class="sml"><b><% If sOrderBy<>"I" Then %><c href="<%=sScriptFileName & ReplaceUrlParams(sParams, "ord=I") %>"><% End If %>Contact&nbsp;ID</b></p></td>
	<th <%If sOrderBy="A" Then Response.Write " bgcolor=""#99CCFF""" %>><p class="sml" width="40%"><b><% If sOrderBy<>"A" Then %><c href="<%=sScriptFileName & ReplaceUrlParams(sParams, "ord=A") %>"><% End If %>Surname, FirstName</a></b></td>
	<th <%If sOrderBy="F" Then Response.Write " bgcolor=""#99CCFF""" %>><p class="sml" width="50"><b><% If sOrderBy<>"F" Then %><c href="<%=sScriptFileName & ReplaceUrlParams(sParams, "ord=F") %>"><% End If %>Proposed<br>first&nbsp;time</b></td>
	<th <%If sOrderBy="L" Then Response.Write " bgcolor=""#99CCFF""" %>><p class="sml" width="50"><b><% If sOrderBy<>"L" Then %><c href="<%=sScriptFileName & ReplaceUrlParams(sParams, "ord=L") %>"><% End If %>Proposed<br>last&nbsp;time</b></td>
	<th <%If sOrderBy="P" Then Response.Write " bgcolor=""#99CCFF""" %>><p class="sml" width="50"><b><% If sOrderBy<>"P" Then %><c href="<%=sScriptFileName & ReplaceUrlParams(sParams, "ord=P") %>"><% End If %>Number<br>of&nbsp;projects</b></p></td>
	<th <%If sOrderBy="C" Then Response.Write " bgcolor=""#99CCFF""" %>><p class="sml" width="50"><b><% If sOrderBy<>"C" Then %><c href="<%=sScriptFileName & ReplaceUrlParams(sParams, "ord=C") %>"><% End If %>Proposed by companies</b></p></td>
	<th><p class="sml" width="20%">Email(s)</b></td>
	<th width="20"><p class="sml">Link<br>to&nbsp;CV</b></td>
	<th width="200"><p class="sml">Flag / Comments</b></td>

	<th width="20"><p class="sml">Options</b></td>
	</tr>

	<% While Not objTempRs.Eof And iCurrentRow<objTempRs.PageSize %>
	<tr class="tr_results<% If iCurrentRow Mod 2 <> 0 Then %> odd<% End If %>">
		<% Response.Write "<td align=""center""><p class=""sml"">" & objTempRs("IDCONTACT") & "</td><td><p class=""mt""><a class=""list"" href=""" & sLinkMpisContact & objTempRs("IDCONTACT") & """ target=_blank>" & objTempRs("NAME") & ", " & objTempRs("FIRSTNAME") & "</a></b></p></td>"

		'Response.Write "<td><p class=""sml"">" & ConvertDateForText(objTempRs("eacRegDate"), "&nbsp;", "DDMMYYYY") & "&nbsp;</td>"
		Response.Write "<td><p class=""sml"">" & ConvertDateForText(objTempRs("DATE_PROPOSED_FIRST"), "&nbsp;", "DDMMYYYY") & "&nbsp;</td>"
		Response.Write "<td><p class=""sml"">" & ConvertDateForText(objTempRs("DATE_PROPOSED_LAST"), "&nbsp;", "DDMMYYYY") & "&nbsp;</td>"
		Response.Write "<td><p class=""sml"">" & objTempRs("NUMBER_PROJECTS_PROPOSED") & "&nbsp;</td>"
		%>
		<td><p class="sml">
			<% =objTempRs("PROPOSED_BY_COMPANIES") %>&nbsp;
			<% If Len(objTempRs("SUBCONTRACTED_BY_COMPANIES"))>2 Then %>
			<br /><b>Subcontracted</b>: <% =objTempRs("SUBCONTRACTED_BY_COMPANIES") %>
			<% End If %>
		</p></td>
		<%
		Response.Write "<td><p class=""sml"">" & sCellStyle & objTempRs("EMAIL") & "</td>"
		%>
		<td><p class="sml fcmp">
			<a class="list" href="download.asp?uid=<% =objTempRs("CONTACT_CV_UID") %>">CV</a>
		</p></td>
		<td>
			<%
			If Len(objTempRs("CONTACT_RELATION_STATUS_FLAG"))>2 Then
				%>
				<p class="sml">
				<img src="/image/flag_<% =objTempRs("CONTACT_RELATION_STATUS_FLAG") %>.gif" alt="<% =objTempRs("CONTACT_RELATION_STATUS_VALUE") %>" vspace="3" width="7" height="12" border="0" align="left">
				&nbsp;<b><% =objTempRs("CONTACT_RELATION_STATUS_VALUE") %></b>
				<% If Len(objTempRs("CONTACT_RELATION_COMMENTS"))>2 Then %>
					(<% =objTempRs("CONTACT_RELATION_COMMENTS") %>)
				<% End If %>
				</p>
				<%
			End If
			%>
			<%
			If objTempRs("CONTACT_DETAILS_NOTFOUND")=1 Then
			%>
			<p class="sml fcmp"><b>Contact details cannot be found.</b></p>
			<p class="sml fcmp"><% = objTempRs("CONTACT_DETAILS_COMMENTS") %></p>
			<%
			End If
			%>
			<p class="sml" style="margin-top: 6px;"><a class="list" href="cv_contactmemo.asp?id_contact=<% =objTempRs("IDCONTACT") %>" target="_blanc"><img src="image/vn_updt.gif" width="15" height="15" align="left" hspace="6" vspace="0" border="0" alt="Edit comments for <% =objTempRs("NAME") & ", " & objTempRs("FIRSTNAME") %>"></a><% =objTempRs("CONTACT_KG_COMMENTS") %></p>
		</p></td>
		<td>
		<p class="sml"><a class="list" href="cv_contactid.asp?id_contact=<% =objTempRs("IDCONTACT") %>" target="_blank">CV is registered. Set CV ID.</a></p>
		<p class="sml" style="margin-top: 4px;"><a class="list" href="cv_notfound.asp?id_contact=<% =objTempRs("IDCONTACT") %>" target="_blank">Contact&nbsp;details&nbsp;cannot&nbsp;be&nbsp;found.</a></p>
		</td>
		<%
		Response.Write "</tr>"
		iCurrentRow=iCurrentRow+1
		objTempRs.MoveNext
	WEnd
End If
objTempRs.Close
Set objTempRs=Nothing
%>
<tr bgcolor="#FFFFFF"><td colspan="10"><p class="mt">Total: <b><%=ShowEntityPlural(iTotalExpertsNumber, "expert", "experts", "&nbsp;") %></b><% If sAction="" And InStr(sScriptFullName, "/cvassortis/") Then Response.Write " visible on assortis.com"%></p></td></tr>
</table>
</div><br />
<% ShowNavigationPages iCurrentPage, iTotalPages, sParams %>

	<!-- footer -->
	<!--#include virtual="/_template/page.footer.asp"-->
<% CloseDBConnection %>
</body>
<!--#include virtual="/_template/html.footer.asp"-->
