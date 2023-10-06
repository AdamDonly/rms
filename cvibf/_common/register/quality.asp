<%
'--------------------------------------------------------------------
'
' Quality of expert profile
'
'--------------------------------------------------------------------
%>
<!--#include virtual="/_template/asp.header.nocache.asp"-->
<!--#include file="../cv_data.asp"-->
<% 
CheckUserLogin sScriptFullNameAsParams
CheckExpertID()

Dim sMessage
sMessage=""
If Request.Form()>"" Then
	sMessage=Left(CheckString(Request.Form("exp_comments")), 40000)
	objTempRs=UpdateRecordSPWithConn(objConnCustom, "usp_ExpertCvQualityReportUpdate", Array( _
		Array(, adInteger, , iExpertID), _
		Array(, adLongVarWChar, 40000, sMessage)))
		
	PrepareEmailTemplate "emlUsrReportCvQuality.htm", _
		";;sExpertCvDatabase=" & UCase(objExpertDB.Database) & _
		";;sExpertID=" & objExpertDB.DatabaseCode & iExpertID & _
		";;sDetails=" & ConvertText(sMessage) & _
		";;sUserName=" & sUserFullName & _
		";;sUserCompany=" & sUserCompany & _
		";;sUserIpAddress=" & sUserIpAddress
	'SendEmail sEmailCvipSystem, sEmailClientCopy, sEmailSubject, sEmailBody, "info"
	'SendEmail sEmailCvipSystem, objExpertDB.ContactEmail, sEmailSubject, sEmailBody, "info"
	'SendEmail sEmailCvipSystem, "jozicic@icaworld.net", sEmailSubject, sEmailBody, "info"
	SendEmail sEmailCvipSystem, "imc@ibf.be", sEmailSubject, sEmailBody, "info"
	%>
	<!--#include virtual="/_common/_template/page.close.asp"-->
	<%
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
		</h2>
		<% ShowMessageStart "info", 580 %>
		Please provide all the necessary details regarding the quality of this CV.
		<% ShowMessageEnd %><br/>
		

	<form method="post" action="<% =sScriptFullName %>">
	<input type="hidden" name="id_Expert" value="<% =objExpertDB.DatabaseCode %><% =iExpertID %>">
		<div class="box search blue">
		<h3><span class="left">&nbsp;</span><span class="right">&nbsp;</span>CV Quality Issue Report</h3>
		<table class="search_form" width="100%" cellspacing=0 cellpadding=0 border=0>
		<tr>
		<td class="field splitter"><label>Origin of CV</label></td>
		<td class="value blue"><p><% =UCase(objExpertDB.Database) %>
		</p></td>
		</tr>
		<% If (bCvValidForMemberOrExpert=1 Or bCvValidForMemberOrExpert=5) Then %>
			<tr>
			<td class="field splitter"><label>Full&nbsp;name</label></td>
			<td class="value blue"><p><% =sFullName %></p></td>
			</tr>
		<% Else %>
			<tr class="last">
			<td class="field splitter"><label>Expert</label></td>
			<td class="value blue"><p><% =objExpertDB.DatabaseCode %><% =iExpertID %></p></td>
			</tr>
		<% End If %>
		</table>
		<table class="search_form" width="100%" cellspacing=0 cellpadding=0 border=0>
		<tr>
		<td class="field splitter"><label for="exp_comments">Remarks on the quality</label></td>
		<td class="value blue"><textarea cols="34" style="width: 355px;" name="exp_comments" rows=12 wrap="yes"><%=sMessage%></textarea></td>
		</tr>
		</table>
		</div>

		<div class="spacebottom">
		<input type="image" class="button first" src="/image/bte_savecont.gif" name="btnSubmit" id="btnSubmit" alt="Save & continue" border=0>
		</div>
		</form>

	</div>
</div>
	<!-- footer -->
	<!--#include virtual="/_template/page.footer.asp"-->
<% CloseDBConnection %>
</body>
<!--#include virtual="/_template/html.footer.asp"-->
