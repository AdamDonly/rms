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
Dim sOrderBy, sSearchString

sOrderBy=UCase(Request.QueryString("ord"))
If sOrderBy<>"E" And sOrderBy<>"R" And sOrderBy<>"I" And sOrderBy<>"U" And sOrderBy<>"B" Then sOrderBy="A"

Dim sLastExperienceMonthFrom, sLastExperienceYearFrom, sLastExperienceMonthTo, sLastExperienceYearTo
Dim sCvModifiedMonthFrom, sCvModifiedYearFrom, sCvModifiedMonthTo, sCvModifiedYearTo

sLastExperienceMonthFrom=CheckInt(Request.QueryString("last_experience_from_month"))
sLastExperienceYearFrom=CheckInt(Request.QueryString("last_experience_from_year"))
sLastExperienceMonthTo=CheckInt(Request.QueryString("last_experience_to_month"))
sLastExperienceYearTo=CheckInt(Request.QueryString("last_experience_to_year"))

sCvModifiedMonthFrom=CheckInt(Request.QueryString("modified_from_month"))
sCvModifiedYearFrom=CheckInt(Request.QueryString("modified_from_year"))
sCvModifiedMonthTo=CheckInt(Request.QueryString("modified_to_month"))
sCvModifiedYearTo=CheckInt(Request.QueryString("modified_to_year"))

sSearchString=Request.QueryString("srch")
If Len(sSearchString)>4 Then 
	If InStr(sSearchString, "-")=4 Then
		sSearchString=CheckIntegerAndZero(Mid(sSearchString, 5, Len(sSearchString)))
	End If
End If
%>
<!--#include virtual="/_template/html.header.asp"-->
<body>
	<!-- header -->
	<!--#include virtual="/_template/page.header.asp"-->
	<div id="content" class="searchform">
	<div class="colCCCCCC uprCse f17 spc01 botMrgn10">MANAGE <% =objUserCompanyDB.DatabaseTitle %> DATABASE OF EXPERTS</div>
<%
iCurrentPage=Request.QueryString("page")
If Not IsNumeric(iCurrentPage) or iCurrentPage="" Then
	iCurrentPage=1
Else
	iCurrentPage=CInt(iCurrentPage)
End If

Dim iShowRemoved
If sAction="all" Or sAction="deleted" Then 
	iShowRemoved=1
Else
	iShowRemoved=0
End If

Set objTempRs=GetDataRecordsetSP("usp_ExpertListSelect", Array( _
	Array(, adInteger, , objUserCompanyDB.ID), _
	Array(, adVarChar, 100, Null), _
	Array(, adInteger, , 0), _
	Array(, adInteger, , iShowRemoved), _
	Array(, adVarChar, 100, sAction), _
	Array(, adVarChar, 255, sSearchString), _
	Array(, adVarChar, 100, sOrderBy), _
	Array(, adVarChar, 16, ConvertDMYForSql(sLastExperienceYearFrom, sLastExperienceMonthFrom, 1)), _
	Array(, adVarChar, 16, ConvertDMYForSql(sLastExperienceYearTo, sLastExperienceMonthTo, 31)), _
	Array(, adVarChar, 16, ConvertDMYForSql(sCvModifiedYearFrom, sCvModifiedMonthFrom, 1)), _
	Array(, adVarChar, 16, ConvertDMYForSql(sCvModifiedYearTo, sCvModifiedMonthTo, 31)) _
	))
	
iTotalExpertsNumber=objTempRs.RecordCount
Dim sExpertFullName, sExpertEmail

If Not objTempRs.Eof Then
	iCurrentRow = 0
	objTempRs.PageSize = 50
	If sAction = "show1000" Then objTempRs.PageSize = 1000
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
		<table style="width:100%">
		<tr>
		<td width="360" style="padding-left:0;">
		<input type="hidden" name="act" value=<%=sAction%>>
		<p class="mt" style="margin:2px;"><b>Search for experts using ID, First name, Family name or Email</b></p>
		<input type="text" name="srch" style="width:344px" value="<%=sSearchString%>"> &nbsp; 
		<p class="mt" style="margin:8px 2px 2px 2px;"><b>Last experience</b></p>
		<p class="mt" style="margin:2px;">from 
		<select name="last_experience_from_month" size="1" style="width:90px"><option></option><% For i=1 to UBound(arrMonthID)%><% Response.Write "<option value=""" & arrMonthID(i) & """"%><% If arrMonthID(i)=sLastExperienceMonthFrom Then Response.Write " selected" %><% Response.Write ">" & arrMonthName(i) & "</option>"%><% Next %></select>
		<select name="last_experience_from_year" size="1"><option></option><% For i=0 to Year(Date())-2002 %><% Response.Write "<option value=""" & (Year(Date())-i) & """"%><% If (Year(Date())-i)=sLastExperienceYearFrom Then Response.Write " selected" %><% Response.Write ">" & (Year(Date())-i) & "</option>"%><% Next %></select> &nbsp; 
		to 
		<select name="last_experience_to_month" size="1" style="width:90px"><option></option><% For i=1 to UBound(arrMonthID)%><% Response.Write "<option value=""" & arrMonthID(i) & """"%><% If arrMonthID(i)=sLastExperienceMonthTo Then Response.Write " selected" %><% Response.Write ">" & arrMonthName(i) & "</option>"%><% Next %></select>
		<select name="last_experience_to_year" size="1"><option></option><% For i=0 to Year(Date())-2002 %><% Response.Write "<option value=""" & (Year(Date())-i) & """"%><% If (Year(Date())-i)=sLastExperienceYearTo Then Response.Write " selected" %><% Response.Write ">" & (Year(Date())-i) & "</option>"%><% Next %></select> &nbsp;</p>

		<p class="mt" style="margin:8px 2px 2px 2px;"><b>CV modified</b></p>
		<p class="mt" style="margin:2px;">from 
		<select name="modified_from_month" size="1" style="width:90px"><option></option><% For i=1 to UBound(arrMonthID)%><% Response.Write "<option value=""" & arrMonthID(i) & """"%><% If arrMonthID(i)=sCvModifiedMonthFrom Then Response.Write " selected" %><% Response.Write ">" & arrMonthName(i) & "</option>"%><% Next %></select>
		<select name="modified_from_year" size="1"><option></option><% For i=0 to Year(Date())-2002 %><% Response.Write "<option value=""" & (Year(Date())-i) & """"%><% If (Year(Date())-i)=sCvModifiedYearFrom Then Response.Write " selected" %><% Response.Write ">" & (Year(Date())-i) & "</option>"%><% Next %></select> &nbsp; 
		to 
		<select name="modified_to_month" size="1" style="width:90px"><option></option><% For i=1 to UBound(arrMonthID)%><% Response.Write "<option value=""" & arrMonthID(i) & """"%><% If arrMonthID(i)=sCvModifiedMonthTo Then Response.Write " selected" %><% Response.Write ">" & arrMonthName(i) & "</option>"%><% Next %></select>
		<select name="modified_to_year" size="1"><option></option><% For i=0 to Year(Date())-2002 %><% Response.Write "<option value=""" & (Year(Date())-i) & """"%><% If (Year(Date())-i)=sCvModifiedYearTo Then Response.Write " selected" %><% Response.Write ">" & (Year(Date())-i) & "</option>"%><% Next %></select> &nbsp;</p>	
	
		</td>
		<td width="650" style="vertical-align: bottom;padding-bottom:4px;">
		<input type="submit" value="Search" class="red-button" />&nbsp;&nbsp;
		<input type="button" value=" Reset" class="red-button" onClick="javascript:window.location.href='<%=sScriptFileName & "?act=" & sAction %>'" />&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
		<input type="button" value="Send emails" class="red-button w95 floatRight" style="margin-right:-8px;" onclick="window.open('send_email.asp?<%= Request.QueryString %>', 'email-experts', 'width=600,height=500,location=0')" />
		<br class="clear" />
		</td>
		</tr>
		</table>
	</td>
	</form>
	</tr>
	<tr class="header-row">
	<th <%If sOrderBy="I" Then Response.Write " bgcolor=""#99CCFF""" %> width="40" align="center">&nbsp;<% If sOrderBy<>"I" Then %><a href="<%=sScriptFileName & ReplaceUrlParams(sTempParams, "ord=I") %>"><u><% End If %>Expert&nbsp;ID</u></a>&nbsp;&nbsp;</th>
	<th width="30%" <%If sOrderBy="A" Then Response.Write " bgcolor=""#99CCFF""" %>><% If sOrderBy<>"A" Then %><a href="<%=sScriptFileName & ReplaceUrlParams(sTempParams, "ord=A") %>"><u><% End If %>Name,&nbsp;First&nbsp;name&nbsp;(Title)</a></th>
	<th width="50" <%If sOrderBy="R" Then Response.Write " bgcolor=""#99CCFF""" %>><% If sOrderBy<>"R" Then %><a href="<%=sScriptFileName & ReplaceUrlParams(sTempParams, "ord=R") %>"><u><% End If %>Registered</th>
	<th width="50" <%If sOrderBy="U" Then Response.Write " bgcolor=""#99CCFF""" %>><% If sOrderBy<>"U" Then %><a href="<%=sScriptFileName & ReplaceUrlParams(sTempParams, "ord=U") %>"><u><% End If %>Modified</th>
	<th width="50" <%If sOrderBy="E" Then Response.Write " bgcolor=""#99CCFF""" %>><% If sOrderBy<>"E" Then %><a href="<%=sScriptFileName & ReplaceUrlParams(sTempParams, "ord=E") %>"><u><% End If %>Last&nbsp;experience</th>
<!--	
	<td width="50" <%If sOrderBy="B" Then Response.Write " bgcolor=""#99CCFF""" %>><p class="sml"><b><% If sOrderBy<>"B" Then %><a href="<%=sScriptFileName & ReplaceUrlParams(sTempParams, "ord=B") %>"><% End If %>Birthday</b></p></td>
-->	
	<th>Email</b></th>
	<th width="20">Status</b></th>
	<th width="30%">Comments</b></th>
	<% If iShowRemoved=1 Then %>
	<th width="20">Options</b></th>
	<% End If %>
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

		sExpertEmail=objTempRs("Email")
		If sApplicationName="external" Then
			If sContactDetailsExternally=cNameObfuscated Then
				sExpertFullName=ObfuscateString(objTempRs("psnLastName")) & ", " & ObfuscateString(objTempRs("psnFirstName")) & " " & ObfuscateString(objTempRs("psnMiddleName")) & " (" & objTempRs("ptlName") & ")"
				sExpertEmail=ObfuscateEmail(objTempRs("Email"))
			End If
			If sContactDetailsExternally = cNameHidden Then
				sExpertFullName = ""
				sExpertEmail = ""
			End If
		End If
		%>
		<tr class="tr_results<% If iCurrentRow Mod 2 <> 0 Then %> odd<% End If %>">
		<td class="number"><% =objExpertDB.DatabaseCode %><% =objTempRs("id_Expert") %></td>
		<td><a class="list" href="../register/register6.asp?uid=<% =objTempRs("uid_Expert") %>" target=_blank><% =sExpertFullName %></a></td>
		<td class="date"><small><% =ConvertDateForText(objTempRs("expCreateDate"), "&nbsp;", "DDMMYYYY") %>&nbsp;</small></td>
		<td class="date"><small><% =ConvertDateForText(objTempRs("expLastUpdate"), "&nbsp;", "DDMMYYYY") %>&nbsp;</small></td>
		<td class="date"><small><% =ConvertDateForText(objTempRs("wkeEndDate"), "&nbsp;", "MonthYear") %>&nbsp;</small></td>
		<td><small><% =sExpertEmail %></small></td>

	<% If iShowRemoved = 0 Then
		' Showing a status
		Dim objExpertStatusCV
		Set objExpertStatusCV = New CExpertStatusCV
		objExpertStatusCV.Expert.ID=objTempRs("id_Expert")
		objExpertStatusCV.LoadData

		Response.Write "<td><p class=""sml"">" 
		If IsObject(objExpertStatusCV.Status) Then
			Response.Write objExpertStatusCV.Status.Name
		End If

		Response.Write "</p></td>" 
		%>
		<td>
			<% ' OLD experts comments:
		'	If objUserCompanyDB.ID = objExpertDB.ID Then 
		'		% > <small><a href="../register/comments.asp?uid=< % =objTempRs("uid_Expert") % >"><img src="< % =sHomePath % >image/vn_updt.gif" width="15" height="15" align="left" hspace="6" vspace="0" border="0" alt="Edit comments for < % =objTempRs("psnLastName") & ", " & objTempRs("psnFirstName") & " " & objTempRs("psnMiddleName") & " (" & objTempRs("ptlName") & ")" % >"></a>< % =objTempRs("expComments") % ></small><%
		'	End If
		
			%><div class="comment-container"><%
				' NEW comments version:
				Set objTempRs2 = GetDataRecordsetSP("usp_ExpertCommentsSelect", Array(Array(, adInteger, , objTempRs("id_Expert"))))
				Dim sIsMyCommentFound, bCanAddComment
				sIsMyCommentFound = 0
				bCanAddComment = 1
				If objExpertDB.DatabasePath = "assortis2db" And (IsNull(iAssortisUserID) Or iAssortisUserID < 1) Then
					bCanAddComment = 0
				End If
				If Not objTempRs2.Eof Then
					While Not objTempRs2.Eof
						Dim bShowOthesComment
						bShowOthesComment = 1
						' add/edit "My" comment 
						If Not IsNull(objTempRs2("CommentorIcaUserId")) Then
							If CStr(iUserID) = CStr(objTempRs2("CommentorIcaUserId")) Then
								' Or iUserID = objTempRs2("id_User_ExpCommentsManager") Then
								If bCanAddComment = 1 Then
									%><a href="javascript:void(0)" class="icon-edit-comment" 
										data-expUid="<% =objTempRs("uid_Expert") %>" 
										data-expName="<%=sExpertFullName %>" 
										data-commentorUserId="<%=objTempRs2("id_User_Commentor") %>" 
										data-expId="<%=objTempRs("id_Database") %>-<%=objTempRs("id_Expert") %>"
										data-isPublic="0" title="Edit comment"><img src="<% =sHomePath %>image/vn_updt.gif" hspace="6" vspace="0" alt="" /></a>
									<div class="comment" data-commentorUserId="<%=objTempRs2("id_User_Commentor") %>" data-expUid="<% =objTempRs("uid_Expert") %>"><span class="comment-title red">ME:</span><br/><div class="my-comment"><% =objTempRs2("Comment") %></div></div>
									<% 
								End If
								sIsMyCommentFound = 1
								bShowOthesComment = 1
							End If
						End If
						
						If bShowOthesComment = 1 Then
							' for now - show other comments only within the member:
							If IsNull(iAssortisMemberID) Then iAssortisMemberID = 0
							If IsNull(iUserCompanyID) Then iUserCompanyID = 0
							If CInt(objTempRs2("CommentorIcaCompanyId")) = cInt(iUserCompanyID) Or CInt(objTempRs2("CommentorAssortisMemberId")) = CInt(iAssortisMemberID) Then
								%><div class="comment"><span class="comment-title"><% =objTempRs2("CommentorUserName") %>:</span><br/><% =objTempRs2("Comment") %></div><%
							End If
						End If 
						
						objTempRs2.MoveNext
					Wend
				End If
				
				If  bCanAddComment = 1 And sIsMyCommentFound = 0 Then
					%><a href="javascript:void(0)" class="icon-edit-comment" 
						data-expUid="<% =objTempRs("uid_Expert") %>" 
						data-expName="<%=sExpertFullName %>" 
						data-commentorUserId="<%=iUserID %>" 
						data-expId="<%=objTempRs("id_Database") %>-<%=objTempRs("id_Expert") %>"
						data-isPublic="0" title="Edit comment"><img src="<% =sHomePath %>image/vn_updt.gif" hspace="6" vspace="0" alt="" /></a>
					<div class="comment hidden" data-commentorUserId="<%=iUserID %>" data-expUid="<% =objTempRs("uid_Expert") %>"><span class="comment-title red">ME:</span><br/><div class="my-comment"></div></div>
					<% 
				End If
					
				objTempRs2.Close
				Set objTempRs2 = Nothing
				%>
				<div class="comment">
					<%
					' show these only for IBF, because they came from Assortis DB:
					If objTempRs("expComments") > "" And iUserCompanyID = 2 Then
						%><span class="comment-title">Old comment:</span><br/><% =objTempRs("expComments") %>
						<%
					End If %>
				</div>
			</div>
		</td>
		<%
	ElseIf iShowRemoved = 1 Then
		%>
		<td><p class="sml fcmp">
		<% If objTempRs("expRemoved")=1 Then %>
			Deleted<br /><% =ConvertDateForText(objTempRs("expRemovedDate"), "&nbsp;", "DDMMYYYY") %>
		<% End If %>
		<% If objTempRs("expDeleted")=1 Then %>
			Deleted<br /><% =ConvertDateForText(objTempRs("expDeletedDate"), "&nbsp;", "DDMMYYYY") %>
		<% End If %>
		</p></td>
		<td><p class="sml fcmp">
		<% On Error Resume Next %>
		<% If objTempRs("expRemoved")=1 Then %>
			<% =objTempRs("expRemovedComments") %>
		<% End If %>
		<% If objTempRs("expDeleted")=1 Then %>
			<% =objTempRs("expDeletedComments") %>
		<% End If %>
		<% On Error GoTo 0 %>
		</p></td>
		<td><p class="sml">
		<% If objTempRs("expRemoved")=1 Or objTempRs("expDeleted")=1 Then %>
			<a href="cv_restore.asp?uid=<% =objTempRs("uid_Expert") %>">Restore&nbsp;CV</a>
		<% End If %>
		</p></td>
		<%
	End If

		Response.Write "</tr>"
		iCurrentRow=iCurrentRow+1
		objTempRs.MoveNext
	WEnd
objTempRs.Close
Set objTempRs=Nothing
%>
<tr class="grid"><td colspan="<% If iShowRemoved=1 Then %>9<% Else %>8<% End If %>"><p class="mt">Total: <b><%=ShowEntityPlural(iTotalExpertsNumber, "expert", "experts", "&nbsp;") %></b><% If Len(sSearchString)>0 Then Response.Write " matching search criteria"%></p></td></tr>
</table>
<script type="text/javascript">	
$(function () {
	$('.icon-edit-comment').click(function (e) {
		$("#comment-dialog").show().css('top', (e.pageY + 10)).css('left', (e.pageX - $("#comment-dialog").width()));
		$("#comment-dialog textarea").text($(this).closest('.comment-container').find('.my-comment').text()).focus();
		$("#comment-dialog input[name='uid']").val($(this).attr('data-expUid'));
		$("#comment-dialog input[name='id_Expert']").val($(this).attr('data-expId'));
		$("#comment-dialog input[name='userid']").val($(this).attr('data-commentorUserId'));
		if ($(this).attr('data-isPublic') == '1' || $(this).attr('data-isPublic') == 'true')
		{
			$("#comment-dialog input[name='ispublic']").attr('checked', 'checked');
		}
		$("#comment-dialog .dialog-header").html('Edit your comment for <strong>' + ($(this).attr('data-expName') != '' ? $(this).attr('data-expName') : 'expert') + '</strong>');
	});
	
	$('#comment-dialog .btn-cancel').click(function () {
		$("#comment-dialog form")[0].reset();
		$("#comment-dialog").hide();
	});
	
	$('#commentform').submit(function (e) {
		e.preventDefault();
		
		$.ajax({
			cache: false,
			url: '../register/comments.asp?uid=' + $("#comment-dialog input[name='uid']").val(),
			type: "POST",
			commentorUserId: $("#comment-dialog input[name='userid']").val(),
			expUid: $("#comment-dialog input[name='uid']").val(),
			data: {
				uid: $("#comment-dialog input[name='uid']").val(),
				expertcomment: $("#comment-dialog textarea[name='expertcomment']").val(),
				id_Expert: $("#comment-dialog input[name='id_Expert']").val(),
				userid: $("#comment-dialog input[name='userid']").val(),
				ispublic: $("#comment-dialog input[name='ispublic']").is(':checked') ? 1 : 0,
			},
			success: function (result) {
				if (result == "OK")
				{
					$('.comment-container .comment[data-commentorUserId="' + this.commentorUserId + '"][data-expUid="' + this.expUid + '"]').removeClass('hidden');
					$('.comment-container .comment[data-commentorUserId="' + this.commentorUserId + '"][data-expUid="' + this.expUid + '"] .my-comment').html($("#comment-dialog textarea[name='expertcomment']").val());
					$("#comment-dialog form")[0].reset();
					$("#comment-dialog").hide();
				}
				else
				{
					alert("Error while processing your comment.");
				}
			},
			error: function () {
				alert("Error while processing your comment.");
			}
		});
	});
});

</script>
</div><br />
<% ShowNavigationPages iCurrentPage, iTotalPages, sParams %>
</div>
	<!-- footer -->
	<!--#include virtual="/_template/page.footer.asp"-->
<% CloseDBConnection %>
<div id="comment-dialog">
	<div class="dialog-header">Edit your comment for expert</div>
	<div class="dialog-content">
		<form id="commentform" name="commentform" method="post">
			<input type="hidden" name="uid" />
			<input type="hidden" name="id_Expert" />
			<input type="hidden" name="userid" />
			<textarea name="expertcomment"></textarea><br/>
			<%
			' hide the option for now:
			If 1 = 2 Then
				%><input type="checkbox" name="ispublic" value="1" /> Make my comment visible for all ICA members in search results. (if not checked, the comment will be visible only for the users within My Organisation)
				<%
			Else
				%><input type="hidden" name="ispublic" value="0" />
				<%
			End If %>
			<div class="dialog-button-line">
				<input type="submit" class="btn-save red-button floatRight w125" name="btnSubmit" id="btnSubmit" value="Save &amp; Close" />
				<a href="javascript:void(0)" class="btn-cancel grey-button floatLeft">Cancel</a>
				<br class="clear" />
			</div>
		</form>
	</div>
</div>
</body>
<!--#include virtual="/_template/html.footer.asp"-->
