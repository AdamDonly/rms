<!--#include file="../cv_data.asp"-->
<% 
CheckUserLogin sScriptFullNameAsParams
CheckExpertID()

Dim sMessage
If Request.Form()>"" Then
	ShowStandardPageHeader	
	
	' Mark expert as copied to the new DB
	objTempRs=GetDataOutParamsSP("usp_Ica_ExpertCopy", Array( _
		Array(, adInteger, , objExpertDB.ID), _
		Array(, adInteger, , iExpertID), _
		Array(, adInteger, , objUserCompanyDB.ID), _
		Array(, adLongVarWChar, 20000, CheckString(Request.Form("exp_comments")))), Array( _ 
		Array(, adInteger)))
	
	If objTempRs(0)>=1 Then
		sMessage="This expert's CV (ID=" & objExpertDB.DatabaseCode & iExpertID & ") was successfully copied."
	End If
	
	' Save copy log
	'On Error Resume Next
		objTempRs=UpdateRecordSP("usp_Ica_ExpertCopyLog", Array( _
			Array(, adInteger, , objExpertDB.ID), _
			Array(, adInteger, , iExpertID), _
			Array(, adInteger, , objUserCompanyDB.ID), _
			Array(, adVarWChar, 255, CheckString(Request.QueryString("exp_firstname"))), _
			Array(, adVarWChar, 255, CheckString(Request.QueryString("exp_familyname"))), _
			Array(, adVarWChar, 150, CheckString(Request.QueryString("exp_email"))), _
			Array(, adVarWChar, 3, CheckString(Request.QueryString("exp_language")))))
	'On Error GoTo 0
	%>
	<br><br>
	<p align="center"><% =sMessage %></p>
	<br><p align="center"><a href="<% =sApplicationHomePath %>view/cv_view.asp?uid=<% =sExpertUID %>"><img src="<% =sHomePath %>image/bte_continue.gif" border=0></a></p>
	<%
	ShowStandardPageFooter
	Response.End
End If
%>
<!--#include virtual="/_template/html.header.asp"-->
<body>
<div id="holder">
	<!-- header -->
	<!--#include virtual="/_template/page.header.asp"-->

	<!-- content -->
	<div id="content" class="searchform">

		<h2 class="service_title">Curriculum Vitae. <span class="service_slogan">Expert ID: <% =objExpertDB.DatabaseCode %><%=iExpertID%></span>
		</h2><br/>

	<form method="post" action="<% =sScriptFullName %>">
	<input type="hidden" name="id_Expert" value="<% =objExpertDB.DatabaseCode %><% =iExpertID %>">
		<div class="box search blue">
		<h3><span class="left">&nbsp;</span><span class="right">&nbsp;</span>Copy CV to <% =objUserCompanyDB.DatabaseTitle %> database</h3>
		<table class="search_form" width="100%" cellspacing=0 cellpadding=0 border=0>
		<tr>
		<td class="field splitter"><label>Provided&nbsp;name</label></td>
		<td class="value blue"><p><% =Request.QueryString("exp_firstname") & " " & Request.QueryString("exp_familyname") %></p></td>
		</tr>
		<tr class="last">
		<td class="field splitter"><label>Provided email</label></td>
		<td class="value blue"><p><% =Request.QueryString("exp_email") %></p></td>
		</tr>
		</table>
		<table class="search_form" width="100%" cellspacing=0 cellpadding=0 border=0>
		<tr>
		<td class="field splitter"><label>Comments</label></td>
		<td class="value blue"><textarea cols="34" style="width:355px;" name="exp_comments" rows=5 wrap="yes"></textarea>
		<p class="sml">&nbsp;&nbsp;Please provide your comments if necessary.</p></td>
		</tr>
		</table>
		</div>

		<div class="spacebottom">
		<input type="image" class="button first" src="/image/bte_copyexpert.gif" name="btnSubmit" id="btnSubmit" alt="Copy expert" border=0>
		</div>
		</form>

	</div>
</div>
	<!-- footer -->
	<!--#include virtual="/_template/page.footer.asp"-->
<% CloseDBConnection %>
</body>
<!--#include virtual="/_template/html.footer.asp"-->
