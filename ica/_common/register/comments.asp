<%
'--------------------------------------------------------------------
'
' Comments on expert profile
'
'--------------------------------------------------------------------
%>
<!--#include virtual="/_template/asp.header.nocache.asp"-->
<!--#include file="../cv_data.asp"-->
<% 
CheckUserLogin sScriptFullNameAsParams
CheckExpertID()

Dim sMessage
If Request.Form() > "" Then
' OLD comments functionality:
'	sComments = Left(CheckString(Request.Form("exp_comments")), 40000)
'	objTempRs=UpdateRecordSPWithConn(objConnCustom, "usp_ExpCvvCommentsUpdate", Array( _
'		Array(, adInteger, , iExpertID), _
'		Array(, adLongVarWChar, 40000, sComments)))
'		
'	Response.Redirect "register6.asp" & sParams
	
	' new comments:
	Dim iUserID_Creator, bIsPublic
	sExpertUID = CheckString(Request.Form("uid"))
	sComments = Left(CheckString(Request.Form("expertcomment")), 40000)
	bIsPublic = CheckIntegerAndZero(Request.Form("ispublic"))
	
	' set user's Assortis User ID depending on objExpertDB.DatabasePath:
	If objExpertDB.DatabasePath = "assortis2db" Then
		iUserID_Creator = iAssortisUserID
	Else
		iUserID_Creator = iUserID
	End If

	If iUserID_Creator > 0 And iExpertID > 0 And sComments > "" Then
		objTempRs = UpdateRecordSPWithConn(objConnCustom, "usp_ExpertCommentAddEdit", Array( _
			Array(, adInteger, , iExpertID), _
			Array(, adInteger, , iUserID_Creator), _
			Array(, adLongVarWChar, 40000, sComments), _
			Array(, adBoolean, , bIsPublic)))
			
		Response.Write "OK"
	Else
		Response.Write "ERROR_2"
	End If
Else
	Response.Write "ERROR_1"
End If

Response.End

' below is the old comments form, not used, so it can be deleted if will not be used anymore:
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
		<h3>Personal information</h3>
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
		<td class="field splitter"><label for="exp_comments">Comments</label></td>
		<td class="value blue"><textarea cols="34" style="width: 355px;" name="exp_comments" rows=12 wrap="yes"><%=sComments%></textarea></td>
		</tr>
		</table>
		</div>

		<div class="spacebottom">
		<input type="submit" class="red-button w125 under-right-col" name="btnSubmit" id="btnSubmit" value="Save & continue" />
		</div>
		</form>

	</div>
	<!-- footer -->
	<!--#include virtual="/_template/page.footer.asp"-->
<% CloseDBConnection %>
</body>
<!--#include virtual="/_template/html.footer.asp"-->
