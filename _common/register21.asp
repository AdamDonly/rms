<% 
'--------------------------------------------------------------------
'
' CV registration.
' Full format. Training.
'
'--------------------------------------------------------------------
Response.Expires = 0
Response.ExpiresAbsolute = Now() - 1
Response.CacheControl = "no-cache" 
%>
<!--#include file="_data/datEduSubject.asp"-->
<!--#include file="_data/datEduType.asp"-->
<!--#include file="_data/datMonth.asp"-->
<%
If sApplicationName<>"expert" Then
	CheckUserLogin sScriptFullNameAsParams
End If
CheckExpertID()
' The page is accessible only if expert id is valid
If iExpertID<=0 Then Response.Redirect(sApplicationHomePath & "register.asp" & sParams)

sParams=ReplaceUrlParams(sParams, "eduid")
Dim bExpEduNewRecord, iExpEduID, sExpEduOther, sExpEduTitle, sExpEduAchievements, sExpEduStartDate, sExpEduEndDate, bFlagSelected

iExpEduID=Request.QueryString("eduid")
If IsNumeric(iExpEduID) And iExpEduID>"" Then
	iExpEduID=CLng(iExpEduID)
Else
	iExpEduID=0
End If

If IsNumeric(iExpEduID) And iExpEduID>"" And sAction="delete" Then
	' Deleting data on projects 	
	objTempRs=UpdateRecordSP("usp_ExpCvvEducationDelete", Array( _
		Array(, adInteger, , iExpertID), _
		Array(, adInteger, , iExpEduID)))
	Response.Redirect(sScriptFileName & ReplaceUrlParams(sParams, "eduid"))
End If

If Request.Form()>"" Then

	iExpEduID=CheckString(Request.Form("id_Edu1"))
	sExpEduOther=CheckString(Request.Form("exp_otr_other"))
	sExpEduTitle=CheckString(Request.Form("exp_otr_desc"))
	sExpEduAchievements=CheckString(Request.Form("exp_otr_name"))
	sExpEduStartDate=ConvertDMYForSQL(CheckString(Request.Form("exp_otr_syear")), CheckString(Request.Form("exp_otr_smonth")), 1)
	sExpEduEndDate=ConvertDMYForSQL(CheckString(Request.Form("exp_otr_eyear")), CheckString(Request.Form("exp_otr_emonth")), 28)

	bExpEduNewRecord=Len(Trim(sExpEduOther)) + Len(Trim(sExpEduTitle)) + Len(Trim(sExpEduAchievements)) + Len(ReplaceIfEmpty(Trim(sExpEduStartDate), "")) + Len(ReplaceIfEmpty(Trim(sExpEduEndDate), ""))

	If IsNumeric(iExpEduID) And iExpEduID>"" And iExpEduID<>"0" Then
		objTempRs=UpdateRecordSP("usp_ExpCvvTrainingUpdate", Array( _
			Array(, adInteger, , iExpertID), _
			Array(, adInteger, , iExpEduID), _
			Array(, adInteger, , 2), _
			Array(, adVarWChar, 255, sExpEduOther), _
			Array(, adVarWChar, 255, sExpEduTitle), _
			Array(, adVarWChar, 255, sExpEduAchievements), _
			Array(, adVarChar, 16, sExpEduStartDate), _
			Array(, adVarChar, 16, sExpEduEndDate), _
			Array(, adVarChar, 3, "Eng")))
	ElseIf bExpEduNewRecord>0 Then
		objTempRs=InsertRecordSP("usp_ExpCvvTrainingInsert", Array( _
			Array(, adInteger, , iExpertID), _
			Array(, adInteger, , 2), _
			Array(, adVarWChar, 255, sExpEduOther), _
			Array(, adVarWChar, 255, sExpEduTitle), _
			Array(, adVarWChar, 255, sExpEduAchievements), _
			Array(, adVarChar, 16, sExpEduStartDate), _
			Array(, adVarChar, 16, sExpEduEndDate), _
			Array(, adVarChar, 3, "Eng")),"-")
	End If

	If Request.Form("exp_otr_continue")="0" then
		Response.Redirect "register21.asp" & sParams
	Else
		Response.Redirect "register3.asp" & sParams
        End if   
End If
%>

<html>
<head>
<title><% =GetLabel(sCvLanguage, "CV registration") %>. <% =GetLabel(sCvLanguage, "Training") %></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" type="text/css" href="<% =sHomePath %>styles.css">
<script type="text/javascript" src="/_scripts/js/main.js"></script>
<script language="JavaScript">
<!--
function validateForm() {
	var f=document.RegForm;
<% If Len(sBackOffice)<3 Then %>
	if (f.exp_otr_other.value!="" || f.exp_otr_name.value!="" || f.exp_otr_desc.value!="") { 
		AddTraining(1);
	} else 
<% End If %>
	{ f.submit(); }
}

function AddTraining(cont_next) {
	var f=document.RegForm;
	if (cont_next!=1) { 
		f.exp_otr_continue.value="0"; 
	}
<% If sApplicationName="external" Then %>
	f.submit();
	return;
<% End If %>
	
<% If Len(sBackOffice)<3 Then %>
	if (f.exp_otr_name.value=="") {
		alert("<% =GetLabel(sCvLanguage, "Please fill in the training title") %>"); document.RegForm.exp_otr_name.select(); return;
	}
	var start_month=parseInt(f.exp_otr_smonth.options[f.exp_otr_smonth.selectedIndex].value);
	var start_year=parseInt(f.exp_otr_syear.options[f.exp_otr_syear.selectedIndex].value);
	var end_month=parseInt(f.exp_otr_emonth.options[f.exp_otr_emonth.selectedIndex].value);
	var end_year=parseInt(f.exp_otr_eyear.options[f.exp_otr_eyear.selectedIndex].value);
	if (end_month==0 || end_year==0) {
		alert("<% =GetLabel(sCvLanguage, "Please fill in the training end date") %>"); return;
	}
	if ((start_year>end_year) || (start_year>0 && start_month>0 && start_year==end_year && start_month>end_month)) {
		alert("<% =GetLabel(sCvLanguage, "Please fill in the training dates properly") %>"); return;
	}
<% End If %>
	f.submit();
}
// -->
</script>

</head>
<body bgcolor="#FFFFFF" topmargin=0 leftmargin=0  marginheight=0 marginwidth=0>
<% ShowTopMenu %>
<% ShowRegistrationProgressBar "CV", 3 %>

  <!--  Other training -->
	<br><table width=580 cellspacing=0 cellpadding=0 border=0 align="center">
	<form method="post" action="<%=sScriptFileName & sParams%>" name="RegForm">
	<input type="hidden" name="exp_otr_continue" value="1">
	<tr height=1><td width=1 bgcolor="<%=colFormHeaderTop%>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td>
	<td width=578 bgcolor="<%=colFormHeaderTop%>"><img src="<% =sHomePath %>image/x.gif" width=498 height=1></td>
	<td width=1 bgcolor="<%=colFormHeaderBottom%>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td></tr>
	<tr height=18><td width=1 bgcolor="<%=colFormHeaderTop%>"><img src="<% =sHomePath %>image/x.gif" width=1 height=18></td>
	<td width=578 bgcolor="<%=colFormHeaderMain%>"><img src="<% =sHomePath %>image/x.gif" width=498 height=1><br><p class="fttl"><img src="<% =sHomePath %>image/<% =imgFormBullet %>" width=7 height=7 align=left vspace=3 hspace=8>TRAINING</p></td>
	<td width=1 bgcolor="<%=colFormHeaderBottom%>"><img src="<% =sHomePath %>image/x.gif" width=1 height=18></td></tr>
	<tr height=1><td width=1 bgcolor="<%=colFormHeaderTop%>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td>
	<td width=578 bgcolor="<%=colFormHeaderBottom%>"><img src="<% =sHomePath %>image/x.gif" width=498 height=1></td>
	<td width=1 bgcolor="<%=colFormHeaderBottom%>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td></tr>
	<tr height=1><td width=1 bgcolor="<% =colFormHeaderSplitter %>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td>
	<td width=578 bgcolor="<% =colFormHeaderSplitter %>"><img src="<% =sHomePath %>image/x.gif" width=498 height=1></td>
	<td width=1 bgcolor="<% =colFormHeaderSplitter %>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td></tr>
   
	<% ' List of training
	Set objTempRs=GetDataRecordsetSP("usp_ExpCvvEducationSelect", Array( _
		Array(, adInteger, , iExpertID), _
		Array(, adInteger, , 2)))
	If Not objTempRs.Eof Then %>
	<tr>
	<td width=1 bgcolor="<%=colFormBodyText%>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td>
	<td width=579 colspan=2 bgcolor="<%=colFormBodyRight%>" valign="top">
		<table width="100%" cellspacing=0 cellpadding=0 border=0>
		<tr>
		<td>
		<table cellspacing=1 cellpadding=1 align="center" width="100%" border=0 bgcolor="<%=colFormBodyRight%>">
		<tr height=20>
			<td width=15 bgcolor="<%=colFormHeaderTop%>"><p>N.</p></td>
			<td width=410 bgcolor="<%=colFormHeaderTop%>"><p><% =GetLabel(sCvLanguage, "Title") %></p></td>
			<td width=60 bgcolor="<%=colFormHeaderTop%>"><p><% =GetLabel(sCvLanguage, "Start date") %></p></td>
			<td width=60 bgcolor="<%=colFormHeaderTop%>"><p><% =GetLabel(sCvLanguage, "End date") %></p></td>
			<td width=15 bgcolor="<%=colFormHeaderTop%>"><p><% =GetLabel(sCvLanguage, "Modify") %></p></td>
			<td width=15 bgcolor="<%=colFormHeaderTop%>"><p><% =GetLabel(sCvLanguage, "Delete") %></p></td>
		</tr>	
		<% i=1
		while not objTempRs.EOF 
		%>
		  <tr height=20>
		    <td bgcolor="<%=colFormBodyText%>"><p align="center"><%=i%></td>
		    <td bgcolor="<%=colFormBodyText%>"><p><a href="register21.asp<%=AddUrlParams(sParams, "eduid=" & objTempRs("id_ExpEdu"))%>&act=update"><% If objTempRs("id_ExpEdu")=iExpEduID Then %><b><img src="<% =sHomePath %>image/vn_v.gif" width=8 height=12 border=0 hspace=0 align="left"><% End If %><%=CheckSpaces(ReplaceIfEmpty(objTempRs("eduDiploma1Eng"), "Not specified"), 45) %></a></td>
		    <td bgcolor="<%=colFormBodyText%>"><p><%=ConvertDateForText(objTempRs("eduStartDate"), "&nbsp;", "MMYYYY") %></td>
		    <td bgcolor="<%=colFormBodyText%>"><p><%=ConvertDateForText(objTempRs("eduEndDate"), "&nbsp;", "MMYYYY") %></td>
		    <td bgcolor="<%=colFormBodyText%>" align="center"><% If objTempRs("id_ExpEdu")=iExpEduID Then %><img src="<% =sHomePath %>image/vn_updte.gif" width=15 height=15 border=0 hspace=0 alt="Updating" align="center"><% Else %><a href="register21.asp<%=AddUrlParams(sParams, "eduid=" & objTempRs("id_ExpEdu"))%>&act=update"><img src="<% =sHomePath %>image/vn_updt.gif" width=15 height=15 border=0 hspace=0 alt="Update this record" align="center"></a><% End If %></td>
		    <td bgcolor="<%=colFormBodyText%>" align="center"><a href="register21.asp<%=AddUrlParams(sParams, "eduid=" & objTempRs("id_ExpEdu"))%>&act=delete"><img src="<% =sHomePath %>image/vn_del.gif" width=15 height=15 border=0 hspace=0 alt="Delete this record" align="center"></a></td>
		  </tr>
		<% i=i+1
		objTempRs.MoveNext
		WEnd 
		%>
		</table>
		</td>
		</tr>
		</table>
	</td>
	</tr>

	<tr height=1><td width=1 bgcolor="<% =colFormHeaderSplitter %>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td>
	<td width=578 bgcolor="<% =colFormHeaderSplitter %>"><img src="<% =sHomePath %>image/x.gif" width=498 height=1></td>
	<td width=1 bgcolor="<% =colFormHeaderSplitter %>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td></tr>

	<% End If 
	objTempRs.Close%>
	
	<% If iExpEduID>0 Then
	Set objTempRs=GetDataRecordsetSP("usp_ExpCvvEducationInfoSelect", Array( _
		Array(, adInteger, , iExpertID), _
		Array(, adInteger, , iExpEduID), _
		Array(, adInteger, , 2)))
	If Not objTempRs.Eof Then
		sExpEduOther=objTempRs("eduOtherEng")
		sExpEduTitle=objTempRs("eduDiploma1Eng")
		sExpEduStartDate=objTempRs("eduStartDate")
		sExpEduEndDate=objTempRs("eduEndDate")
		sExpEduAchievements=objTempRs("eduDescriptionEng")
	End If
	objTempRs.Close
	Set objTempRs=Nothing
	End If %>
    
	<tr><td width=1 bgcolor="<%=colFormBodyText%>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td>
	<td width=578 bgcolor="<%=colFormHeaderTop%>" valign="top">
		<table width="100%" cellspacing=0 cellpadding=0 border=0>
		<tr><td width=170><p class="ftxt"><% =GetLabel(sCvLanguage, "Title") %></td>
		<td bgcolor="<%=colFormBodyRight%>" width=1><img src="x.gif" width=1 height=36></td>
		<td bgcolor="<%=colFormBodyText%>" width=407><img src="x.gif" width=1 height=6><br>
		&nbsp;&nbsp;<input type="text" name="exp_otr_name" size="31" style="width:355px;" value="<%=sExpEduTitle%>" maxlength="255">&nbsp;&nbsp;<span class="fcmp">*</span></td></tr>

		<tr><td width=170><p class="ftxt"><% =GetLabel(sCvLanguage, "Skills / Qualifications") %></td>
		<td bgcolor="<%=colFormBodyRight%>" width=1><img src="x.gif" width=1 height=24></td>
		<td bgcolor="<%=colFormBodyText%>" width=407>
		&nbsp;&nbsp;<input type="text" name="exp_otr_other" size="31" style="width:355px;" value="<%=sExpEduOther%>" maxlength="255"></td></tr>
		
		<tr><td width=170><p class="ftxt"><% =GetLabel(sCvLanguage, "Start date") %></td>
		<td bgcolor="<%=colFormBodyRight%>" width=1><img src="x.gif" width=1 height=24></td>
		<td bgcolor="<%=colFormBodyText%>" width=407>
		&nbsp;&nbsp;<select name="exp_otr_smonth" size=1>
		<option value="0" selected><% =GetLabel(sCvLanguage, "Month") %></option>
		<%  For i=1 to UBound(arrMonthID)
		If IsDate(sExpEduStartDate) Then
			If Month(sExpEduStartDate)=arrMonthID(i) Then
				bFlagSelected=" selected"
			Else
				bFlagSelected=""
			End If
		End If	
		Response.Write("<option value=""" & arrMonthID(i) & """" & bFlagSelected & ">" & arrMonthName(i) & "</option>")
		Next %>
		</select>
		<select name="exp_otr_syear" size="1">
		<option value="0" Selected><% =GetLabel(sCvLanguage, "Year") %></option>
		<% For i=0 to 60 
		If IsDate(sExpEduStartDate) Then
			If Year(sExpEduStartDate)=Year(Date())-i then
				bFlagSelected=" selected"
			Else
				bFlagSelected=""
			End If
		End If	
		Response.Write("<option value=""" & Year(Date())-i & """" & bFlagSelected & ">" & Year(Date())-i & "</option>")
		Next %>
		</select>
		</td></tr>

		<tr><td width=170><p class="ftxt"><% =GetLabel(sCvLanguage, "End date") %></td>
		<td bgcolor="<%=colFormBodyRight%>" width=1><img src="x.gif" width=1 height=24></td>
		<td bgcolor="<%=colFormBodyText%>" width=407>
		&nbsp;&nbsp;<select name="exp_otr_emonth" size=1>
		<option value="0" selected><% =GetLabel(sCvLanguage, "Month") %></option>
		<% For i=1 to UBound(arrMonthID)
		If IsDate(sExpEduEndDate) Then
			If Month(sExpEduEndDate)=arrMonthID(i) Then
				bFlagSelected=" selected"
			Else
				bFlagSelected=""
			End If
		End If	
		Response.Write("<option value=""" & arrMonthID(i) & """" & bFlagSelected & ">" & arrMonthName(i) & "</option>")
		Next %>
		</select>
		<select name="exp_otr_eyear" size="1">
		<option value="0" Selected><% =GetLabel(sCvLanguage, "Year") %></option>
		<% For i=0 to 60 
		If IsDate(sExpEduEndDate) Then
			If Year(sExpEduEndDate)=Year(Date())-i then
				bFlagSelected=" selected"
			Else
				bFlagSelected=""
			End If
		End If	
		Response.Write("<option value=""" & Year(Date())-i & """" & bFlagSelected & ">" & Year(Date())-i & "</option>")
		Next %>
		</select>&nbsp;&nbsp;<span class="fcmp">*</span>
		</td></tr>

		<tr><td width=170><p class="ftxt"><% =GetLabel(sCvLanguage, "Achievements") %></td>
		<td bgcolor="<%=colFormBodyRight%>" width=1><img src="x.gif" width=1 height=24></td>
		<td bgcolor="<%=colFormBodyText%>" width=407>&nbsp;&nbsp;<input type="text" name="exp_otr_desc" size="31" style="width:355px;" value="<%=sExpEduAchievements%>" maxlength="255">&nbsp;&nbsp;
		<img src="<% =sHomePath %>image/x.gif" width=358 height=1 vspace=3><br>
		</td></tr>
		</table>
	</td>
	<td width=1 bgcolor="<%=colFormBodyRight%>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td></tr>

	<tr height=1><td width=1 bgcolor="<%=colFormBodyText%>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td>
	<td width=578 bgcolor="<%=colFormHeaderMain%>"><img src="<% =sHomePath %>image/x.gif" width=498 height=1></td>
	<td width=1 bgcolor="<%=colFormBodyRight%>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td></tr>
	<tr height=1><td width=1 bgcolor="<% =colFormHeaderSplitter %>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td>
	<td width=578 bgcolor="<% =colFormHeaderSplitter %>"><img src="<% =sHomePath %>image/x.gif" width=498 height=1></td>
	<td width=1 bgcolor="<% =colFormHeaderSplitter %>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td></tr>
	</table><br>

	<table width=580 cellspacing=0 cellpadding=0 border=0 align="center">
	<tr><td width=380 align=left>
	<img src="<% =sHomePath %>image/x.gif" width=170 height=1><a href="javascript:AddTraining(0);"><img src="<% =sHomePath %>image/bte_<% If iExpEduID>0 Then %>save<% Else %>add<% End If %>training.gif" name="Add this training" alt="Add this training to the list of your training" border=0 onSubmit="javascript:document.RegForm.exp_otr_continue.value='0'"></a>

	<input type="hidden" name="id_Edu1" value="<%=iExpEduID%>">
	</td>
	<td width=200 height=1 align="right"><a href="javascript:validateForm()"><img src="<% =sHomePath %>image/bte_savecont.gif" name="Continue" alt="Save and continue" border=0></a></td>
	</tr>
	</form>
	</table> 
<br><br>

<% CloseDBConnection %>
</body>
</html>

