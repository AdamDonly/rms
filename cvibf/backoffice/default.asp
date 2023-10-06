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
If bUserMaskExperts=aUserMaskView Then
	Response.Redirect "/backoffice/search/exp_search.asp"
End If
%>
<!--#include virtual="/_template/html.header.asp"-->
<body>
<div id="holder">
	<!-- header -->
	<!--#include virtual="/_template/page.header.asp"-->

	<!-- content -->
	<div id="content" class="fullscreen">
	<div class="dialog dlgMedium floatLeft">
		<span class="header" style="border: 0"><% =GetExpCount("all") %> experts in ICA Members' Databases</span>
	</div>

	<div class="dialog dlgMedium floatRight">
		<div class="body">
		<ul><h4>
			<li style="font-weight:normal; padding:12px 0;"><a href="register/exists.asp<% =sParams %>">Verify if expert already exists in ICA Members' Databases</a></li>
			</h4>
		</ul>
		</div>
	</div>

	<div class="clear"></div>

	<% If (bUserMaskExperts=aUserMaskNoAccess) Then %>
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

	<% If (bUserMaskExperts And aUserMaskView) Then %>
	<div class="dialog dlgMedium floatLeft">
		<span class="header" style="font-size:16.8px">Search</span>
		<div class="body">
		<ul><h4>
			<li style="font-weight: normal; padding: 12px 0;"><a href="search/exp_search.asp<% =sParams %>">Search for experts</a></li>
			</h4>
		</ul>
		</div>
	</div>
	<% End If %>

	<% If (bUserMaskExperts And aUserMaskAdd) Then %>
	<div class="dialog dlgMedium floatRight">
		<span class="header" style="font-size:16.8px">Register expert in <% =objUserCompanyDB.DatabaseTitle %> database</span>
		<div class="body">
		<ul><h4>
			<li style="font-weight:normal; padding:12px 0;"><a href="register/register.asp<% =sParams %>">Complete CV registration</a></li>
			</h4>
		</ul>
		</div>
	</div>
	<% End If %>

	<% If (bUserMaskExperts And aUserMaskEdit) Then %>
	<div class="dialog dlgMedium floatLeft">
		<span class="header" style="font-size:16.8px">Manage <% =objUserCompanyDB.DatabaseTitle %> database</span>
		<div class="body">
		<ul><h4>
			<li style="font-weight:normal; padding:4px 0; padding-top: 12px;"><a href="manage/cv_list.asp<% =sParams %>">List of all experts visible in the database</a></li>
			<li style="font-weight:normal; padding:4px 0;"><a href="manage/cv_list.asp<% =ReplaceUrlParams(sParams, "act=registeredweek") %>">List of experts registered this week</a></li>
			<li style="font-weight:normal; padding:4px 0;"><a href="manage/cv_list.asp<% =ReplaceUrlParams(sParams, "act=registeredmonth") %>">List of experts registered this month</a></li>

			<li style="font-weight:normal; padding:4px 0; padding-top: 12px;"><a href="manage/cv_list.asp<% =ReplaceUrlParams(sParams, "act=updatedrange3to6&ord=E") %>">List of experts CVs not updated for the last 3 to 6 months</a></li>
			<li style="font-weight:normal; padding:4px 0;"><a href="manage/cv_list.asp<% =ReplaceUrlParams(sParams, "act=updatedrange6to12&ord=E") %>">List of experts CVs not updated for the last 6 to 12 months</a></li>
			<li style="font-weight:normal; padding:4px 0;"><a href="manage/cv_list.asp<% =ReplaceUrlParams(sParams, "act=updatedrange12to24&ord=E") %>">List of experts CVs not updated for the last 12 to 24 months</a></li>
			<li style="font-weight:normal; padding:4px 0;"><a href="manage/cv_list.asp<% =ReplaceUrlParams(sParams, "act=updatedover24&ord=E") %>">List of experts CVs not updated for more than 24 months</a></li>

			<li style="font-weight:normal; padding:4px 0; padding-top: 12px;"><a href="manage/cv_list.asp<% =ReplaceUrlParams(sParams, "act=deleted") %>">List of deleted experts</a></li>
			</h4>
		</ul>
		</div>
	</div>
	<% End If %>

	<% If (bUserMaskExperts And aUserMaskEdit) Then %>
	<div class="dialog dlgMedium floatRight">
		<span class="header" style="font-size:16.8px">MPIS experts</span>
		<div class="body">
		<ul><h4>
			<li style="font-weight:normal; padding:4px 0; padding-top: 12px;"><a href="manage/cv_mpis.asp?act=ibf<% =sParams %>">List of experts on MPIS without CV in DB (proposed by IBF)</a></li>
			<li style="font-weight:normal; padding:4px 0;"><a href="manage/cv_mpis.asp?act=others<% =sParams %>">List of experts on MPIS without CV in DB (proposed by partners)</a></li>
			</h4>
		</ul>
		</div>
	</div>
	<% End If %>
</div>
	<!-- footer -->
	<!--#include file="../_template/page.footer.asp"-->
<% CloseDBConnection %>
</body>
<!--#include file="../_template/html.footer.asp"-->
