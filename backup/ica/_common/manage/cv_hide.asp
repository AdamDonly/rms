<%
'--------------------------------------------------------------------
'
' Hiding a duplicate CV
'
'--------------------------------------------------------------------
%>
<!--#include virtual="/_template/asp.header.nocache.asp"-->
<!--#include file="../cv_data.asp"-->
<% 
CheckUserLogin sScriptFullNameAsParams
CheckExpertID()

Dim sMessage
If Request.Form()>"" Then
	ShowStandardPageHeader	

	If iExpertID=CheckIntegerAndZero(Request.Form("exp_originalid")) Then
		sMessage="You cannot hide a duplicate CV with itself. Please provide the correct Original Expert ID."
	End If
	
	objTempRs=GetDataOutParamsSP("usp_AdmExpDuplicateHide", Array( _
		Array(, adInteger, , iExpertID), _
		Array(, adInteger, , CheckInteger(Request.Form("exp_originalid"))), _
		Array(, adLongVarWChar, 20000, CheckString(Request.Form("exp_comments")))), Array( _ 
		Array(, adInteger)))
	
	If objTempRs(0)>=1 Then
		sMessage="Expert's CV (ID=" & objExpertDB.DatabaseCode & iExpertID & ") was successfully hidden."
	End If
	%>
	<br><br>
	<p align="center"><% =sMessage %></p>
	<br><p align="center"><a href="<% =sApplicationHomePath %>"><img src="<% =sHomePath %>image/bte_continue.gif" border=0></a></p>
	<%
	ShowStandardPageFooter
	Response.End
End If
%>
<!--#include virtual="/_template/html.header.asp"-->
<body>
	<!-- header -->
	<!--#include virtual="/_template/page.header.asp"-->

	<!-- content -->
	<div id="content" class="searchform">

		<div class="colCCCCCC uprCse f17 spc01 botMrgn10"><span class="service_title">Curriculum Vitae.</span> Expert ID: <% =objExpertDB.DatabaseCode %><%=iExpertID%></div>

	<form method="post" action="<% =sScriptFullName %>">
	<input type="hidden" name="id_Expert" value="<% =objExpertDB.DatabaseCode %><% =iExpertID %>">
		<div class="box search blue">
		<h3>Hide duplicate CV</h3>
		<table class="search_form" width="100%" cellspacing=0 cellpadding=0 border=0>
		<tr>
		<td class="field splitter"><label>Full&nbsp;name</label></td>
		<td class="value blue"><p><% =sFullName %></p></td>
		</tr>
		<tr>
		<td class="field splitter"><label>Date of birth</label></td>
		<td class="value blue"><p><% =ConvertDateForText(sBirthDate, "&nbsp;", "DDMMYYYY") %></p></td>
		</tr>
		<tr class="last">
		<td class="field splitter"><label>Email</label></td>
		<td class="value blue"><p><% =sEmail %></p></td>
		</tr>
		</table>
		<table class="search_form" width="100%" cellspacing=0 cellpadding=0 border=0>
		<tr>
		<td class="field splitter"><label>Original Expert ID</label></td>
		<td class="value blue"><input type="text" name="exp_originalid" size="10" style="width:75px;"></td>
		</tr>
		<tr>
		<td class="field splitter"><label for="exp_comments">Comments</label></td>
		<td class="value blue"><textarea cols="34" style="width:355px;" name="exp_comments" rows=5 wrap="yes"></textarea>
		<p class="sml">&nbsp;&nbsp;Please provide your comments if necessary.</p></td>
		</tr>
		</table>
		</div>

		<div class="spacebottom">
		<input type="submit" name="btnSubmit" id="btnSubmit" class="red-button under-right-col w150" value="Hide this duplicate CV" />
		</div>
		</form>

	</div>
	<!-- footer -->
	<!--#include virtual="/_template/page.footer.asp"-->
<% CloseDBConnection %>
</body>
<!--#include virtual="/_template/html.footer.asp"-->
