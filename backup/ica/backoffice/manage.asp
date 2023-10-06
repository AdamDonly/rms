<%@ LANGUAGE="VBSCRIPT" %>
<% Option Explicit %>
<!--#include file="../dbc.asp"-->
<!--#include file="../fnc.asp"-->
<!--#include file="../fnc_exp.asp"-->
<!--#include file="../_forms/frmInterface.asp"-->
<!--#include file="../../_common/_class/main.asp"-->
<!--#include file="../../_common/_class/project.asp"-->
<%
' Check UserID 
CheckUserLogin sScriptFullNameAsParams

If bIsUserTopExpert Then
	Response.Redirect("/backoffice/mycv/register6.asp")
	Response.End()
End If

sParams=ReplaceUrlParams(sParams, "idproject")
sParams=ReplaceUrlParams(sParams, "idexpert")

' If user can only search and view CVs - redirect to the search page
If iUserAccessMaskExperts=aUserAccessMaskView Then
	Response.Redirect "/backoffice/search/exp_search.asp"
End If
%>
<!--#include virtual="/_template/html.header.asp"-->
<body>
	<!-- header -->
	<!--#include virtual="/_template/page.header.asp"-->

	<!-- content -->
	<td id="content" class="searchform">
<div class="colCCCCCC uprCse f17 spc01 botMrgn10">MANAGE <% =objUserCompanyDB.DatabaseTitle %> DATABASE OF EXPERTS</div>

	<% If (iUserAccessMaskExperts=aUserAccessMaskNoAccess) Then %>
	<div class="dialog dlgMedium floatLeft">
		<span class="header" style="font-size:16.8px">Access denied</span>
		<div class="body">
		<ul><h4>
			<li style="font-weight: normal; padding: 12px 0;"><b>You don't have access to the experts section!</b></br>In order to get the access please request it to your company administrator.</li>
			</h4>
		</ul>
		</div>
	</div>
	<% End If %>

	<% If (iUserAccessMaskExperts And aUserAccessMaskEdit) Then %>
		<table class="sqbtn">
			<tr>
			<td><a href="manage/cv_list.asp<% =sParams %>">All experts visible in the database</a></td>
			<td><a href="manage/cv_list.asp<% =ReplaceUrlParams(sParams, "act=registeredweek") %>">Experts registered this week</a></td>
			<td><a href="manage/cv_list.asp<% =ReplaceUrlParams(sParams, "act=registeredmonth") %>">Experts registered this month</a></td>
			<td><a href="manage/cv_list.asp<% =ReplaceUrlParams(sParams, "act=updatedover12&ord=E") %>">Experts with CVs not updated for the past 12 months</a></td>
			<% If iMemberAccessExperts = cMemberAccessExpertsOwnOnly Then %>
				<td><a href="manage/cv_list.asp<% =ReplaceUrlParams(sParams, "act=deleted") %>">Hidden experts</a></td>
			<% Else %>
				<td><a href="manage/cv_list.asp<% =ReplaceUrlParams(sParams, "act=deleted") %>">Deleted experts</a></td>
			<% End If %>
			</tr>
		</table>
		<script type="text/javascript">
			$(function () {
				$('table.sqbtn tr td').click(function () {
					console.log($(this).find('a').length);
					window.location = $(this).find('a').attr('href');
					return false;
				});
			});
		</script>
	<% End If %>
	
</div>
	<!-- footer -->
	<!--#include file="../_template/page.footer.asp"-->
<% CloseDBConnection %>
</body>
<!--#include file="../_template/html.footer.asp"-->
