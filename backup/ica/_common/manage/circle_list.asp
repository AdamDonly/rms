<%
'--------------------------------------------------------------------
'
' List of experts in the database
'
'--------------------------------------------------------------------
%>
<!--#include virtual="/_common/_template/asp.header.notimeout.asp"-->
<!--#include file="../_data/datMonth.asp"-->
<!--#include virtual="/fnc_exp.asp"-->
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

' If the users role is companyAdmin or CvContact then lets load the list of experts


Dim iTotalExpertsNumber
Dim iTotalPages, iTotalRecords, iCurrentPage, iCurrentRow, sRowColor, iSearchQueryID, bSaveSearchLog, j, iIsActive
Dim sOrderBy, sSearchString

sOrderBy = UCase(Request.QueryString("ord"))
If sOrderBy <> "E" _
And sOrderBy <> "R" _
And sOrderBy <> "I" _
And sOrderBy <> "U" _
And sOrderBy <> "B" _
And sOrderBy <> "CIRCLE" _
And sOrderBy <> "TOPEXPERT" _
And sOrderBy <> "MANAGER" _
Then 
	sOrderBy = "A"
End If

sSearchString=Request.QueryString("srch")
If Len(sSearchString)>4 Then 
	If InStr(sSearchString, "-")=4 Then
		sSearchString=CheckIntegerAndZero(Mid(sSearchString, 5, Len(sSearchString)))
	End If
End If
%>
<!--#include virtual="/_template/html.header.start.asp"-->
<script type="text/javascript">
function removeExpertCircle(expert_uid, expert_topexpert, expert_name) {
	var message = 'Are you sure you want to remove ' + expert_name + '\nfrom your Experts Circle?'
	if (expert_topexpert) {
		message += '\nTop expert account for this expert will also be disabled.'
	}
	if (confirm(message)) {
		location.replace('/backoffice/view/cv_circle_fields.asp?uid=' + expert_uid + '&act=remove');
	}
}
function removeTopExpert(expert_uid, expert_topexpert, expert_name)
{
	if (confirm('Are you sure you want to remove the top expert account of expert ' + expert_name + '?')) {
		location.replace('<%= sIcaServerProtocol & sIcaServer %>/Intranet/RemoveTopExpert?uid=' + expert_uid);
	}
}
</script>
</head>
<body>
	<script>
		function loadUpdateExpertManager(iCircleId, iCompanyId, iUserID) {					
			$.ajax({
				url: '../../svc/expertcircle_expertmanagers.asp',
				data: { id: iCircleId, companyId: iCompanyId, userId: iUserID },
				cache: false,
				success: function (data) {
					$('#uem-placeholder-' + iCircleId).html(data);
					$('#updateExpertLink-' + iCircleId).hide();
				},
				error: function (jqXHR, textStatus, err) {
					alert('Error updating Expert Manager');
				}
			});
		}

		function onUpdateManagerComplete(iUserId,iCircleId,managersName) {			
			// clear ddl
			$('#uem-placeholder-' + iCircleId).html("");
			// show and update the anchor tag			
			$('#updateExpertLink-' + iCircleId)[0].onclick = function () { loadUpdateExpertManager(iCircleId, "<%=iUserCompanyID%>",iUserId); };
			$('#updateExpertLink-' + iCircleId).html("<small>" + managersName + "</small>")
			$('#updateExpertLink-' + iCircleId).show();
		}
	</script>
	<!-- header -->
	<!--#include virtual="/_template/page.header.asp"-->
	<div id="content" class="searchform">
	<div class="colCCCCCC uprCse f17 spc01 botMrgn10">MY EXPERTS CIRCLE</div>
<%
iCurrentPage = Request.QueryString("page")
If Not IsNumeric(iCurrentPage) or iCurrentPage="" Then
	iCurrentPage = 1
Else
	iCurrentPage = CInt(iCurrentPage)
End If

Dim iShowRemoved
If sAction="all" Or sAction="deleted" Then 
	iShowRemoved = 1
Else
	iShowRemoved = 0
End If

Set objTempRs = GetDataRecordsetSP("usp_" & sIcaServerSqlPrefix & "CompanyExpertCircleListSelect", Array( _
	Array(, adInteger, , iUserCompanyID), _
	Array(, adVarChar, 100, Null), _
	Array(, adVarChar, 100, sAction), _
	Array(, adVarChar, 255, sSearchString), _
	Array(, adVarChar, 100, sOrderBy) _
	))
	
iTotalExpertsNumber = objTempRs.RecordCount
Dim sExpertFullName, sExpertEmail, sExpertLastRemovalDate, sExpertLastRemovalTopExpertDate

If Not objTempRs.Eof Then
	iCurrentRow=0
	objTempRs.PageSize=50
	iTotalRecords=objTempRs.RecordCount
	iTotalPages=objTempRs.PageCount
	objTempRs.AbsolutePage=CInt(iCurrentPage)
	ShowNavigationPages iCurrentPage, iTotalPages, sParams
End If
%>
<div class="frame blue" align="center">
	<table class="results blue-table">
	<tr class="search-filter-row">
	<form method="get" action="<%=sScriptFileName & sParams%>">
	<td colspan="<% If iShowRemoved=1 Then %>9<% Else %>8<% End If %>" style="padding-left:0;">
		<table>
		<tr>
		<td width="360" style="padding-left:0;">
		<input type="hidden" name="act" value=<%=sAction%>>
		<p class="mt" style="margin:2px 2px 2px 0;"><b>Search for experts using ID, First name, Family name or Email</b></p>
		<input type="text" name="srch" style="width:344px" value="<%=sSearchString%>"> &nbsp; 
	
		</td>
		<td width="*" style="vertical-align:bottom;padding-bottom:5px;">
		<input type="submit" value="Search" class="red-button" />&nbsp;&nbsp;
		<input type="button" value=" Reset" class="red-button" onClick="javascript:window.location.href='<%=sScriptFileName & "?act=" & sAction %>'"/>
		</p>
		</td>
		</tr>
		</table>
	</td>
	</form>
	</tr>
	<tr class="header-row">
	<th <%If sOrderBy="I" Then Response.Write " bgcolor=""#99CCFF""" %> width="40" align="center">&nbsp;<% If sOrderBy<>"I" Then %><a href="<%=sScriptFileName & ReplaceUrlParams(sTempParams, "ord=I") %>"><u><% End If %>Expert&nbsp;ID</u></a>&nbsp;&nbsp;</th>
	<th width="20%" <%If sOrderBy="A" Then Response.Write " bgcolor=""#99CCFF""" %>><% If sOrderBy<>"A" Then %><a href="<%=sScriptFileName & ReplaceUrlParams(sTempParams, "ord=A") %>"><u><% End If %>SURNAME, First Name (Title)</a></th>
	<th width="200">Email</th>
	<th width="120" <%If sOrderBy="CIRCLE" Then Response.Write " bgcolor=""#99CCFF""" %>><% If sOrderBy <> "CIRCLE" Then %><a href="<%=sScriptFileName & ReplaceUrlParams(sTempParams, "ord=CIRCLE") %>"><u><% End If %>Expert&nbsp;Circle&nbsp;Details</u></a></th>
	<th width="120" <%If sOrderBy="TOPEXPERT" Then Response.Write " bgcolor=""#99CCFF""" %>><% If sOrderBy <> "TOPEXPERT" Then %><a href="<%=sScriptFileName & ReplaceUrlParams(sTempParams, "ord=TOPEXPERT") %>"><u><% End If %>Top&nbsp;Expert&nbsp;Details</u></a></th>
	<th <%If sOrderBy="MANAGER" Then Response.Write " bgcolor=""#99CCFF""" %>><% If sOrderBy <> "MANAGER" Then %><a href="<%=sScriptFileName & ReplaceUrlParams(sTempParams, "ord=MANAGER") %>"><u><% End If %>Expert&nbsp;Manager</u></a></th>
	<th width="20%">Fields&nbsp;of&nbsp;Expertise</th>
	</tr>

	<% While Not objTempRs.Eof And iCurrentRow<objTempRs.PageSize %>
		<% 
		Set objExpertDB = objExpertDBList.Find(objTempRs("id_Database"), "ID")
		
		sExpertFullName = objTempRs("psnLastName")
		If Len(objTempRs("psnFirstName")) > 0 Then
			sExpertFullName = sExpertFullName & ", " & objTempRs("psnFirstName")
		End If
		If Len(objTempRs("ptlName")) > 0 Then
			sExpertFullName = sExpertFullName & " (" & objTempRs("ptlName") & ")"
		End If
		sExpertEmail = objTempRs("Email")
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
		If Not IsNull(objTempRs("lastRemovedDate")) Then
			sExpertLastRemovalDate = ConvertDateForText(objTempRs("lastRemovedDate"), "&nbsp;", "DDMMYYY")
		Else
			sExpertLastRemovalDate = ""
		End If
		'
		If Not IsNull(objTempRs("lastRemovedTopExpertDate")) Then
			sExpertLastRemovalTopExpertDate = ConvertDateForText(objTempRs("lastRemovedTopExpertDate"), "&nbsp;", "DDMMYYY")
		Else
			sExpertLastRemovalTopExpertDate = ""
		End If
		Dim iTopExpertStatus
		iTopExpertStatus = CheckIntegerAndZero(objTempRs("TOPEXPERTSTATUS"))
		'GetExpertCompanyTopExpertByUid(objTempRs("uid_Expert"), iUserCompanyID, iUserID)
		If Not IsNull(objTempRs("IsActive")) Then
			iIsActive = CheckIntegerAndZero(objTempRs("IsActive"))
		Else 
			iIsActive = 0
		End If
		%>
		<tr class="tr_results<% If iCurrentRow Mod 2 <> 0 Then %> odd<% End If %><% If iIsActive = 0 Then %> not-active<% End If %>">
		<td class="number"><% =objExpertDB.DatabaseCode %><% =objTempRs("id_Expert") %></td>
		<td><a class="list" href="../register/register6.asp?uid=<% =objTempRs("uid_Expert") %>" target=_blank><% =sExpertFullName %></a></td>
		<td><small><% =sExpertEmail %></small></td>
		<td><small>Added&nbsp;<% =ConvertDateForText(objTempRs("ExpertCircleDateCreate"), "&nbsp;", "DDMMYYY") %>
			<% If (iUserCompanyRoleID = cUserRoleCompanyAdministrator Or _
				iUserCompanyRoleID = cUserRoleGlobalAdministrator Or _
				iUserCompanyRoleID = cUserRoleCVContactPoint Or _
				iUserID = CheckIntegerAndZero(objTempRs("addedByUserID"))) _
				And (iIsActive = 1) Then
				%><br/><a href='javascript:removeExpertCircle("<% =objTempRs("uid_Expert") %>", <% =iTopExpertStatus %>, "<% =sExpertFullName %>")' class="add-del">Delete</a>
			<% End If 
			If Len(sExpertLastRemovalDate) > 0 Then
				%><br/>Deleted&nbsp;<%=sExpertLastRemovalDate %><br/>
				<%
			End If 
			If iIsActive = 0 Then
				If Len(sExpertLastRemovalDate) = 0 Then
					%><br/><%
				End If
				%><a href="/backoffice/view/cv_circle_fields.asp?uid=<% =objTempRs("uid_Expert") %>" class="add-del">Re-add</a><%
			End If %>
		</small></td>
		<td><small>
			<%
			If iTopExpertStatus = 1 Then %>
				<img src="/image/file_top.gif" width=18 height=17 border=0 hspace=3 align="left">Added <% If Not IsNull(objTempRs("TOPEXPERTAPPROVED")) Then %><%=ConvertDateForText(objTempRs("TOPEXPERTAPPROVED"), "&nbsp;", "DDMMYYY") %><% End If %>
				<% ' the user who added the expert to the circle, should not be able to remove the top expert:
				If (iUserCompanyRoleID = cUserRoleCompanyAdministrator Or _
					iUserCompanyRoleID = cUserRoleGlobalAdministrator Or _
					iUserCompanyRoleID = cUserRoleCVContactPoint) _
				Then
					%><br/><a href='javascript:removeTopExpert("<% =objTempRs("uid_Expert") %>", <% =iTopExpertStatus %>, "<% =sExpertFullName %>")' class="add-del">Delete</a>
				<% End If %>
			<% ElseIf iTopExpertStatus = 2 Then %>
				<img src="/image/file_top.gif" width=18 height=17 border=0 hspace=3 align="left">Requested <% If Not IsNull(objTempRs("TOPEXPERTREQUESTED")) Then %><%=ConvertDateForText(objTempRs("TOPEXPERTREQUESTED"), "&nbsp;", "DDMMYYY") %><% End If %>
			<% ElseIf iIsActive = 1 Then 
				If Len(sExpertLastRemovalTopExpertDate) > 0 Then 
					%>Deleted&nbsp;<%=sExpertLastRemovalTopExpertDate %><br/>
					<%
				End If
				If (iUserCompanyRoleID = cUserRoleCompanyAdministrator Or _
					iUserCompanyRoleID = cUserRoleGlobalAdministrator Or _
					iUserCompanyRoleID = cUserRoleCVContactPoint Or _
					iUserID = CheckIntegerAndZero(objTempRs("addedByUserID"))) _ 
				Then
					%>
					<a href="<% =sIcaServerProtocol & sIcaServer %>/Intranet/Dashboard?act=terequest&val=<% =objTempRs("uid_Expert") %>"><%If Len(sExpertLastRemovalTopExpertDate) > 0 Then %>Re-add<% Else %>Add<% End If %></a>
					<%
				End If
			End If %>
		</small></td>
		<td> 
			<% 
			
			If iUserCompanyRoleID = cUserRoleCompanyAdministrator Or iUserCompanyRoleID = cUserRoleCVContactPoint Then %>
				<a id="updateExpertLink-<%=objTempRs("IDCIRCLE")%>" href="#" style="color:#74001b;" onclick="loadUpdateExpertManager('<%=objTempRs("IDCIRCLE")%>', <%=iUserCompanyID%>, <%=objTempRs("addedByUserID")%>);"><small><% =objTempRs("addedByUserFullName") %></small></a>
				<div id="uem-placeholder-<%=objTempRs("IDCIRCLE")%>"></div>
			<% Else %>
				<small><% =objTempRs("addedByUserFullName") %></small>
			<% End If %>
		</td>
		<td><small>
		<%
		Dim bCanChangeExpSelection
		bCanChangeExpSelection = 0
		If (iUserCompanyRoleID = cUserRoleCompanyAdministrator Or _
			iUserCompanyRoleID = cUserRoleGlobalAdministrator Or _
			iUserCompanyRoleID = cUserRoleCVContactPoint Or _
			iUserID = CheckIntegerAndZero(objTempRs("addedByUserID"))) _ 
		Then
			bCanChangeExpSelection = 1
		End If
		Set objTempRs2 = GetDataRecordsetSP("usp_" & sIcaServerSqlPrefix & "ExpertCompanyExpertCircleFieldsSelect", Array( _
			Array(, adVarChar, 40, objTempRs("uid_Expert")), _
			Array(, adInteger, , iUserCompanyID)))
		If Not objTempRs2.Eof Then 
			Dim iExpertCircleSectorCount, iExpertCircleCountryCount, iExpertCircleDonorCount
			iExpertCircleSectorCount = CheckIntegerAndZero(objTempRs2("SectorCount"))
			iExpertCircleCountryCount = CheckIntegerAndZero(objTempRs2("CountryCount"))
			iExpertCircleDonorCount = CheckIntegerAndZero(objTempRs2("DonorCount"))
			If bCanChangeExpSelection = 1 Then
				%><a href="/backoffice/view/cv_circle_fields.asp?uid=<% =objTempRs("uid_Expert") %>" target="_blank">
				<%
			End If 
			%>
			<% =ShowEntityPlural(iExpertCircleSectorCount, "sector", "sectors", " ") %> /
			<% =ShowEntityPlural(iExpertCircleCountryCount, "country", "countries", " ") %> /
			ALL donors
			<%
			If bCanChangeExpSelection = 1 Then 
				%></a>
				<%
			End If 
		Else 
			If bCanChangeExpSelection = 1 Then
				%><a class="list" href="/backoffice/view/cv_circle_fields.asp?uid=<% =objTempRs("uid_Expert") %>" target="_blank">Update fields selection</a>
				<%
			End If
		End If
		%>
		</small></td>
		</tr>
		<%
		Set objTempRs2 = Nothing
		iCurrentRow=iCurrentRow+1
		objTempRs.MoveNext
	WEnd
objTempRs.Close
Set objTempRs=Nothing
%>
<tr class="grid"><td colspan="<% If iShowRemoved=1 Then %>9<% Else %>8<% End If %>"><p class="mt">Total: <b><%=ShowEntityPlural(iTotalExpertsNumber, "expert", "experts", "&nbsp;") %></b><% If Len(sSearchString)>0 Then Response.Write " matching search criteria"%></p></td></tr>
</table>
</div><br />
<% ShowNavigationPages iCurrentPage, iTotalPages, sParams %>
</div>
	<!-- footer -->
	<!--#include virtual="/_template/page.footer.asp"-->
<% CloseDBConnection %>
</body>
<!--#include virtual="/_template/html.footer.asp"-->
