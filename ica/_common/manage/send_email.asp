<%
'--------------------------------------------------------------------
'
' List of experts in the database
'
'--------------------------------------------------------------------
%>
<!--#include virtual="/_template/asp.header.notimeout.asp"-->

<!--#include file="../_data/datMonth.asp"-->
<!--#include virtual="/_common/_class/main.asp"-->
<!--#include virtual="/_common/_class/expert.asp"-->
<!--#include virtual="/_common/_class/status_cv.asp"-->
<!--#include virtual="/_common/_class/expert.status_cv.asp"-->
<% 
CheckUserLogin sScriptFullName

Dim iTotalExpertsNumber
Dim iTotalPages, iTotalRecords, iCurrentPage, iCurrentRow, sRowColor, iSearchQueryID, sSelect, bSaveSearchLog, j
Dim lstDuplicateIDs, arrDuplicateIDs, sDuplicates
Dim sCellStyle, sOrderBy, sSearchString
Dim sLastExperienceMonthFrom, sLastExperienceYearFrom, sLastExperienceMonthTo, sLastExperienceYearTo
Dim sCvModifiedMonthFrom, sCvModifiedYearFrom, sCvModifiedMonthTo, sCvModifiedYearTo

sOrderBy=UCase(Request.QueryString("ord"))
If sOrderBy<>"E" And sOrderBy<>"R" And sOrderBy<>"I" And sOrderBy<>"U" Then sOrderBy="A"

sSearchString=Request.QueryString("srch")

sLastExperienceMonthFrom=CheckInt(Request.QueryString("last_experience_from_month"))
sLastExperienceYearFrom=CheckInt(Request.QueryString("last_experience_from_year"))
sLastExperienceMonthTo=CheckInt(Request.QueryString("last_experience_to_month"))
sLastExperienceYearTo=CheckInt(Request.QueryString("last_experience_to_year"))

sCvModifiedMonthFrom=CheckInt(Request.QueryString("modified_from_month"))
sCvModifiedYearFrom=CheckInt(Request.QueryString("modified_from_year"))
sCvModifiedMonthTo=CheckInt(Request.QueryString("modified_to_month"))
sCvModifiedYearTo=CheckInt(Request.QueryString("modified_to_year"))

iSearchQueryID=CheckInt(Request.QueryString("qid"))
sSelect=Request.QueryString("select")
%>

<html>
<head>
<title><%=sApplicationName%>. Send emails to experts</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" type="text/css" href="/style.css">
<link href="/Resources/Styles/ica-experts.css" rel="stylesheet" type="text/css" media="all"/>
<script src="/Resources/Scripts/oth/jquery.js" type="text/javascript"></script>
<script src="/Resources/Scripts/tinymce/tinymce.min.js" type="text/javascript"></script>
<script type="text/javascript">
	function ConfirmEmail() {
		if (confirm("Do you want to send emails to all the experts matching the selected criteria?"))
			{ document.forms[0].submit(); }
	}

	$(function () {
		tinymce.init({
			selector: "textarea",
			entity_encoding: "raw",
			plugins: [
				"advlist autolink lists link charmap anchor",
				"searchreplace fullscreen",
				"insertdatetime table contextmenu paste textcolor colorpicker textpattern"
			],

			menubar: false,
			statusbar: false,
			toolbar: "undo redo | styleselect | fontselect | fontsizeselect | alignleft aligncenter alignright alignjustify | bullist numlist outdent indent | forecolor backcolor | link unlink anchor | insertdatetime | table | hr removeformat | subscript superscript | charmap emoticons | ltr rtl | visualchars visualblocks nonbreaking pagebreak",

			toolbar_items_size: 'small'
		});
	});
</script>
</head>
<body style="text-align:left">
<% ShowTopMenu %>

<%
iCurrentPage=Request.QueryString("page")
If Not IsNumeric(iCurrentPage) or iCurrentPage="" Then
	iCurrentPage=1
Else
	iCurrentPage=CInt(iCurrentPage)
End If

If iSearchQueryID>0 And sSelect="all" Then

	UpdateMemberSearchQuery iMemberID, iSearchQueryID

	Set objTempRs=GetDataRecordsetSP("usp_MmbExpQuerySelect", Array( _
			Array(, adInteger, , iMemberID), _
			Array(, adInteger, , Null), _
			Array(, adInteger, , iSearchQueryID)))	
		
ElseIf iSearchQueryID>0 Then
	Set objTempRs=GetDataRecordsetSP("usp_MmbExpQuerySelect", Array( _
			Array(, adInteger, , iMemberID), _
			Array(, adInteger, , Null), _
			Array(, adInteger, , iSearchQueryID)))	
Else
	Set objTempRs=GetDataRecordsetSP("usp_ExpertListSelect", Array( _
		Array(, adInteger, , objUserCompanyDB.ID), _
		Array(, adVarChar, 100, Null), _
		Array(, adInteger, , 0), _
		Array(, adInteger, , 0), _
		Array(, adVarChar, 100, sAction), _
		Array(, adVarChar, 255, sSearchString), _
		Array(, adVarChar, 100, sOrderBy), _
		Array(, adVarChar, 16, ConvertDMYForSql(sLastExperienceYearFrom, sLastExperienceMonthFrom, 1)), _
		Array(, adVarChar, 16, ConvertDMYForSql(sLastExperienceYearTo, sLastExperienceMonthTo, 31)), _
		Array(, adVarChar, 16, ConvertDMYForSql(sCvModifiedYearFrom, sCvModifiedMonthFrom, 1)), _
		Array(, adVarChar, 16, ConvertDMYForSql(sCvModifiedYearTo, sCvModifiedMonthTo, 31)) _
		))
End If

iTotalExpertsNumber=objTempRs.RecordCount

If Not objTempRs.Eof Then
	iCurrentRow=0
	objTempRs.PageSize=50
	iTotalRecords=objTempRs.RecordCount
	iTotalPages=objTempRs.PageCount
	objTempRs.AbsolutePage=CInt(iCurrentPage)
	sParams=AddUrlParams(sParams, "act=" & sAction)
End If
objTempRs.Close
Set objTempRs=Nothing
	%>
	<div class="colCCCCCC uprCse f17 spc01 botMrgn10" style="padding:5px 0 0 5px;">Send emails to <span class="service_title"><%=ShowEntityPlural(iTotalExpertsNumber, "expert", "experts", "&nbsp;") %></span><% If sAction="" And InStr(sScriptFullName, "/cvassortis/") Then Response.Write " matching the selected criteria"%></div>
	
	<table cellpadding="0" cellspacing="0" class="send-email-table">
		<form action="do_send_email.asp?<% =Request.QueryString %>" method="post" onSubmit="ConfirmEmail(); return false;">
			<input type="hidden" name="message_from" id="message_from" value="<%=sUserEmail %>"/>
			<input type="hidden" name="message_to" id="message_to" value="expert_all"/>
			<tr>
				<td class="label">From</td>
				<td><%=sUserEmail %></td>
			</tr>
			<tr>
				<td class="label">To</td>
				<td>All experts in the search result</td>
			</tr>
			<tr>
				<td class="label">Subject</td>
				<td><input type="text" name="message_subject" id="message_subject" class="inputFldWideR" /></td>
			</tr>
			<tr>
				<td class="label">Body</td>
				<td><textarea rows="16" name="message_body" id="message_body" class="inputFldWideR"></textarea></td>
			</tr>
			<tr>
				<td>&nbsp;</td>
				<td><input type="submit" value="Send" class="red-button"/> &nbsp; &nbsp;
					<input type="button" value="Cancel" onclick="javascript:window.close();" class="red-button"/>
				</td>
			</tr>
		</form>
	</table>
</body>
</html>
