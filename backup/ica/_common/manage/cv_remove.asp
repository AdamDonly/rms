<% 
'--------------------------------------------------------------------
'
' Removing expert from the database
'
'--------------------------------------------------------------------
%>
<!--#include file="../cv_data.asp"-->
<% 
CheckUserLogin sScriptFullNameAsParams
CheckExpertID()
%>
<%
If Request.Form()>"" Then
	sParams=ReplaceUrlParams(sParams, "id")
	ShowStandardPageHeader	

	objTempRs = GetDataOutParamsSPWithConn(objConnCustom, "usp_AdmExpRemove", Array( _
		Array(, adInteger, , iExpertID), _
		Array(, adInteger, , CheckInteger(Request.Form("exp_reason"))), _
		Array(, adLongVarWChar, 20000, CheckString(Request.Form("exp_comments")))), Array( _ 
		Array(, adInteger)))
	
	If objTempRs(0)>=1 Then
		Response.Write "<br><br><br><br><p align=""center"">The CV of the expert with ID " & objExpertDB.DatabaseCode & iExpertID & " was successfully deleted."
	End If
	%><br><br>
	<a href="<% =sApplicationHomePath %>"><img src="<% =sHomePath %>image/bte_continue.gif" border=0></a>
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

	<div class="colCCCCCC uprCse f17 spc01 botMrgn10"><span class="service_title">Curriculum Vitae. </span>Expert ID: <% =objExpertDB.DatabaseCode %><%=iExpertID%></div>

	<form method="post" action="<% =sScriptFullName %>">
	<input type="hidden" name="id_Expert" value="<% =objExpertDB.DatabaseCode %><% =iExpertID %>">
		<div class="box search blue">
		<h3>Remove Expert CV</h3>
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
		<td class="field splitter"><label>Reason</label></td>
		<td class="value blue">
		<p><input type="radio" name="exp_reason" value="1" checked>&nbsp;This is not a real CV</p>
		<p><input type="radio" name="exp_reason" value="2" style="margin-top:8px;">&nbsp;Expert doesn't want to have his CV being registered in the database</p></td>
		</tr>
		<tr>
		<td class="field splitter"><label>Comments</label></td>
		<td class="value blue"><textarea cols="34" style="width:355px;" name="exp_comments" rows=5 wrap="yes"></textarea>
		<p class="sml">&nbsp;&nbsp;Please paste here full text of expert's email, <br>&nbsp;where he / she is asking to be removed from the database.</p></td>
		</tr>
		</table>
		</div>

		<div class="spacebottom">
		<input type="submit" class="red-button under-right-col w150" name="btnSubmit" id="btnSubmit" value="Remove this expert" />
		</div>
		</form>

	</div>
	<!-- footer -->
	<!--#include virtual="/_template/page.footer.asp"-->
<% CloseDBConnection %>
</body>
<!--#include virtual="/_template/html.footer.asp"-->
