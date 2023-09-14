<%
'--------------------------------------------------------------------
'
' CV registration.
' Full format. Education.
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
Dim bExpEduNewRecord, iExpEduID, sExpEduInstitution, sExpEduLocation, iExpEduDiplomaID, sExpEduDiplomaText, iExpEduSubjectID, sExpEduSubjectText, sExpEduStartDate, sExpEduEndDate, bDateValueSelected

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
	iExpEduID=CheckString(Request.Form("id_Edu"))
	sExpEduInstitution=CheckString(Request.Form("exp_Inst_name"))
	sExpEduLocation=""
	iExpEduSubjectID=CheckString(Request.Form("exp_edu_subj"))
	sExpEduSubjectText=CheckString(Request.Form("exp_edu_subj1"))
	iExpEduDiplomaID=CheckString(Request.Form("exp_edu_diploma"))
	sExpEduDiplomaText=CheckString(Request.Form("exp_edu_diploma1"))
	sExpEduStartDate=ConvertDMYForSQL(CheckString(Request.Form("exp_edu_syear")), CheckString(Request.Form("exp_edu_smonth")), 1)
	sExpEduEndDate=ConvertDMYForSQL(CheckString(Request.Form("exp_edu_eyear")), CheckString(Request.Form("exp_edu_emonth")), 28)

	bExpEduNewRecord=Len(Trim(sExpEduInstitution)) + Len(Trim(sExpEduLocation)) + Len(Trim(sExpEduSubjectText)) + Len(Trim(sExpEduDiplomaText)) + Len(ReplaceIfEmpty(Trim(sExpEduStartDate), "")) + Len(ReplaceIfEmpty(Trim(sExpEduEndDate), ""))
	If IsNumeric(iExpEduSubjectID) Then bExpEduNewRecord=bExpEduNewRecord + iExpEduSubjectID
	If IsNumeric(iExpEduDiplomaID) Then bExpEduNewRecord=bExpEduNewRecord + iExpEduDiplomaID

	If IsNumeric(iExpEduID) And iExpEduID>"" And iExpEduID<>"0" Then
		objTempRs=UpdateRecordSP("usp_ExpCvvEducationUpdate", Array( _
			Array(, adInteger, , iExpertID), _
			Array(, adInteger, , iExpEduID), _
			Array(, adInteger, , 1), _
			Array(, adSmallInt, , iExpEduSubjectID), _
			Array(, adSmallInt, , iExpEduDiplomaID), _
			Array(, adVarWChar, 255, sExpEduSubjectText), _
			Array(, adVarWChar, 255, sExpEduDiplomaText), _
			Array(, adVarChar, 3, "Eng"), _
			Array(, adVarWChar, 255, sExpEduInstitution), _
			Array(, adVarWChar, 255, sExpEduLocation), _
			Array(, adVarChar, 16, sExpEduStartDate), _
			Array(, adVarChar, 16, sExpEduEndDate)))
	ElseIf bExpEduNewRecord>0 Then
		objTempRs=InsertRecordSP("usp_ExpCvvEducationInsertNew", Array( _
			Array(, adInteger, , iExpertID), _
			Array(, adInteger, , 1), _
			Array(, adSmallInt, , iExpEduSubjectID), _
			Array(, adSmallInt, , iExpEduDiplomaID), _
			Array(, adVarWChar, 255, sExpEduSubjectText), _
			Array(, adVarWChar, 255, sExpEduDiplomaText), _
			Array(, adVarChar, 3, "Eng"), _
			Array(, adVarWChar, 255, sExpEduInstitution), _
			Array(, adVarWChar, 255, sExpEduLocation), _
			Array(, adVarChar, 16, sExpEduStartDate), _
			Array(, adVarChar, 16, sExpEduEndDate)),"-")
	End If

	If Request.Form("exp_edu_continue")="0" then
		Response.Redirect "register2.asp" & sParams
	Else
		Response.Redirect "register21.asp" & sParams
        End if   
End If
%>

<html>
<head>
<title>CV registration. Education</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" type="text/css" href="<% =sHomePath %>styles.css">
<script type="text/javascript" src="/_scripts/js/main.js"></script>
<script language="JavaScript">
<!--
function validateForm() { 
	var f=document.RegForm;
<% If Len(sBackOffice)<3 Then %>
	if (f.exp_Inst_name.value!="" || f.exp_edu_subj.selectedIndex>0 || f.exp_edu_syear.selectedIndex>0) { 
		AddEducation(1); 
	} else 
<% End If %>
	{ f.submit(); }
}

function AddEducation(cont_next) {  
	var f=document.RegForm;
	if (cont_next!=1) {
		f.exp_edu_continue.value="0"; 
	}
<% If sApplicationName="external" Then %>
	f.submit();
	return;
<% End If %>

<% If Len(sBackOffice)<3 Then %>
	if (f.exp_Inst_name.value=="") {
		alert("Please fill in the institution name."); document.RegForm.exp_Inst_name.focus(); return;
	}
	var start_month=parseInt(f.exp_edu_smonth.options[f.exp_edu_smonth.selectedIndex].value);
	var start_year=parseInt(f.exp_edu_syear.options[f.exp_edu_syear.selectedIndex].value);
	var end_month=parseInt(f.exp_edu_emonth.options[f.exp_edu_emonth.selectedIndex].value);
	var end_year=parseInt(f.exp_edu_eyear.options[f.exp_edu_eyear.selectedIndex].value);
	if (start_month==0 || start_year==0) {
		alert("Please fill in the education start date."); return;
	}
	if (end_month==0 || end_year==0) {
		alert("Please fill in the education end date."); return;
	}
	if ((start_year>end_year) || (start_year>0 && start_month>0 && start_year==end_year && start_month>end_month)) {
		alert("Please fill in the education dates properly."); return;
	}
	if (f.exp_edu_diploma.selectedIndex==0 && f.exp_edu_diploma1.value=="") {
		alert("Please specify a type of diploma or degree obtained."); return;
	}
	if (f.exp_edu_subj.selectedIndex==0) {
		alert("Please specify the education subject."); return;
	}
<% End If %>  

	if (cont_next!=1) {
		f.exp_edu_continue.value="0"; 
	}
	f.submit();
}
-->
</script>

</head>
<body bgcolor="#FFFFFF" topmargin=0 leftmargin=0  marginheight=0 marginwidth=0>
<% ShowTopMenu %>
<% ShowRegistrationProgressBar "CV", 2 %>

  <!-- Education -->
	<br><table width=580 cellspacing=0 cellpadding=0 border=0 align="center">
	<form method="post" action="<%=sScriptFileName & sParams%>" name="RegForm">
	<input type="hidden" name="exp_edu_continue" value="1">
	<tr height=1><td width=1 bgcolor="<%=colFormHeaderTop%>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td>
	<td width=578 bgcolor="<%=colFormHeaderTop%>"><img src="<% =sHomePath %>image/x.gif" width=498 height=1></td>
	<td width=1 bgcolor="<%=colFormHeaderBottom%>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td></tr>
	<tr height=18><td width=1 bgcolor="<%=colFormHeaderTop%>"><img src="<% =sHomePath %>image/x.gif" width=1 height=18></td>
	<td width=578 bgcolor="<%=colFormHeaderMain%>"><img src="<% =sHomePath %>image/x.gif" width=498 height=1><br><p class="fttl"><img src="<% =sHomePath %>image/<% =imgFormBullet %>" width=7 height=7 align=left vspace=3 hspace=8>EDUCATION</p></td>
	<td width=1 bgcolor="<%=colFormHeaderBottom%>"><img src="<% =sHomePath %>image/x.gif" width=1 height=18></td></tr>
	<tr height=1><td width=1 bgcolor="<%=colFormHeaderTop%>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td>
	<td width=578 bgcolor="<%=colFormHeaderBottom%>"><img src="<% =sHomePath %>image/x.gif" width=498 height=1></td>
	<td width=1 bgcolor="<%=colFormHeaderBottom%>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td></tr>
	<tr height=1><td width=1 bgcolor="<% =colFormHeaderSplitter %>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td>
	<td width=578 bgcolor="<% =colFormHeaderSplitter %>"><img src="<% =sHomePath %>image/x.gif" width=498 height=1></td>
	<td width=1 bgcolor="<% =colFormHeaderSplitter %>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td></tr>

	<%
	Set objTempRs=GetDataRecordsetSP("usp_ExpCvvEducationSelect", Array( _
		Array(, adInteger, , iExpertID), _
		Array(, adInteger, , 1)))
	If Not objTempRs.Eof Then %>
	<tr>                                                           
	<td width=1 bgcolor="<%=colFormBodyText%>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td>
	<td width=579 colspan=2 bgcolor="<%=colFormBodyRight%>" valign="top">
		<table width="100%" cellspacing=0 cellpadding=0 border=0>
		<tr>
		<td>
		<table cellspacing=1 cellpadding=1 align="center" width="100%" border=0 bgcolor="<%=colFormBodyRight%>">
		<tr height=20>
			<td width=10 bgcolor="<%=colFormHeaderTop%>"><p>No.</td>
			<td width=215 bgcolor="<%=colFormHeaderTop%>"><p>Institution&nbsp;Name</td>
			<td width=60 bgcolor="<%=colFormHeaderTop%>"><p>Start&nbsp;Date</td>
			<td width=60 bgcolor="<%=colFormHeaderTop%>"><p>End&nbsp;Date</td>
			<td width=200 bgcolor="<%=colFormHeaderTop%>"><p>Subject</td>
			<td width=15 bgcolor="<%=colFormHeaderTop%>"><p>Modify</td>
			<td width=15 bgcolor="<%=colFormHeaderTop%>"><p>Delete</td>
		</tr>	
		<% i=1
		While Not objTempRs.Eof %>
		  <tr height=20>
		    <td bgcolor="<%=colFormBodyText%>"><p align="center"><%=i%></td>
		    <td bgcolor="<%=colFormBodyText%>"><p><a href="register2.asp<%=AddUrlParams(sParams, "eduid=" & objTempRs("id_ExpEdu"))%>&act=update"><% If objTempRs("id_ExpEdu")=iExpEduID Then %><b><img src="<% =sHomePath %>image/vn_v.gif" width=8 height=12 border=0 hspace=0 align="left"><% End If %><%=CheckSpaces(ReplaceIfEmpty(objTempRs("InstNameEng"),"Not specified"), 30) %></a></td>
		    <td bgcolor="<%=colFormBodyText%>"><p><%=ConvertDateForText(objTempRs("eduStartDate"), "&nbsp;", "MMYYYY") %></td>
		    <td bgcolor="<%=colFormBodyText%>"><p><%=ConvertDateForText(objTempRs("eduEndDate"), "&nbsp;", "MMYYYY") %></td>
		    <td bgcolor="<%=colFormBodyText%>"><p><%=objTempRs("edsDescriptionEng") %></td>
		    <td bgcolor="<%=colFormBodyText%>" align="center"><% If objTempRs("id_ExpEdu")=iExpEduID Then %><img src="<% =sHomePath %>image/vn_updte.gif" width=15 height=15 border=0 hspace=0 alt="Updating" align="center"><% Else %><a href="register2.asp<%=AddUrlParams(sParams, "eduid=" & objTempRs("id_ExpEdu"))%>&act=update"><img src="<% =sHomePath %>image/vn_updt.gif" width=15 height=15 border=0 hspace=0 alt="Update this record" align="center"></a><% End If %></td>
		    <td bgcolor="<%=colFormBodyText%>" align="center"><a href="register2.asp<%=AddUrlParams(sParams, "eduid=" & objTempRs("id_ExpEdu"))%>&act=delete"><img src="<% =sHomePath %>image/vn_del.gif" width=15 height=15 border=0 hspace=0 alt="Delete this record" align="center"></a></td>
		  </tr>
		<% i=i+1
		objTempRs.MoveNext
		WEnd %>
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
	objTempRs.Close
	Response.Flush()
	
	If iExpEduID>0 Then
	Set objTempRs=GetDataRecordsetSP("usp_ExpCvvEducationInfoSelect", Array( _
		Array(, adInteger, , iExpertID), _
		Array(, adInteger, , iExpEduID), _
		Array(, adInteger, , 1)))

	If Not objTempRs.Eof then
		sExpEduInstitution=objTempRs("InstNameEng")
                sExpEduLocation=objTempRs("InstLocationEng")

		iExpEduDiplomaID=objTempRs("eduDiploma")
		sExpEduDiplomaText=objTempRs("eduDiploma1Eng")
		
		iExpEduSubjectID=objTempRs("id_eduSubject")
		sExpEduSubjectText=objTempRs("id_eduSubject1Eng")

		sExpEduStartDate=objTempRs("eduStartDate")
		sExpEduEndDate=objTempRs("eduEndDate")
	End If
	objTempRs.Close
	Set objTempRs=Nothing
	End If
	%>

	<tr><td width=1 bgcolor="<%=colFormBodyText%>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td>
	<td width=578 bgcolor="<%=colFormHeaderTop%>" valign="top">
		<table width="100%" cellspacing=0 cellpadding=0 border=0>
		<tr><td width=170><p class="ftxt">Institution&nbsp;name</td>
		<td bgcolor="<%=colFormBodyRight%>" width=1><img src="x.gif" width=1 height=36></td>
		<td bgcolor="<%=colFormBodyText%>" width=407><img src="x.gif" width=1 height=6><br>
		&nbsp;&nbsp;<input type="text" name="exp_Inst_name" size="31" style="width:355px;" value="<%=sExpEduInstitution%>" maxlength="255">&nbsp;&nbsp;<span class="fcmp">*</span></td></tr>

		<tr><td width=170><p class="ftxt">Start&nbsp;date</td>
		<td bgcolor="<%=colFormBodyRight%>" width=1><img src="x.gif" width=1 height=24></td>
		<td bgcolor="<%=colFormBodyText%>" width=407>
		&nbsp;&nbsp;<select name="exp_edu_smonth" size=1>
		<option value="0" selected>Month</option>
		<%  For i=1 to UBound(arrMonthID)
		If IsDate(sExpEduStartDate) Then
			If Month(sExpEduStartDate)=arrMonthID(i) Then
				bDateValueSelected=" selected"
			Else
				bDateValueSelected=""
			End If
		End If	
		Response.Write("<option value=""" & arrMonthID(i) & """" & bDateValueSelected & ">" & arrMonthName(i) & "</option>")
		Next %>
		</select>
		<select name="exp_edu_syear" size="1">
		<option value="0" Selected>Year</option>
		<% For i=0 to 60 
		If IsDate(sExpEduStartDate) Then
			If Year(sExpEduStartDate)=Year(Date())-i then
				bDateValueSelected=" selected"
			Else
				bDateValueSelected=""
			End If
		End If	
		Response.Write("<option value=""" & Year(Date())-i & """" & bDateValueSelected & ">" & Year(Date())-i & "</option>")
		Next %>
		</select>&nbsp;&nbsp;<span class="fcmp">*</span>
		</td></tr>

		<tr><td width=170><p class="ftxt">End&nbsp;date</td>
		<td bgcolor="<%=colFormBodyRight%>" width=1><img src="x.gif" width=1 height=24></td>
		<td bgcolor="<%=colFormBodyText%>" width=407>
		&nbsp;&nbsp;<select name="exp_edu_emonth" size=1>
		<option value="0" selected>Month</option>
		<%   For i=1 to UBound(arrMonthID)
		If IsDate(sExpEduEndDate) Then
			If Month(sExpEduEndDate)=arrMonthID(i) Then
				bDateValueSelected=" selected"
			Else
				bDateValueSelected=""
			End If
		End If	
		Response.Write("<option value=""" & arrMonthID(i) & """" & bDateValueSelected & ">" & arrMonthName(i) & "</option>")
		Next %>
		</select>
		<select name="exp_edu_eyear" size="1">
		<option value="0" Selected>Year</option>
		<% For i=0 to 65 
		If IsDate(sExpEduEndDate) Then
			If Year(sExpEduEndDate)=Year(Date())+3-i then
				bDateValueSelected=" selected"
			Else
				bDateValueSelected=""
			End If
		End If	
		Response.Write("<option value=""" & Year(Date())+3-i & """" & bDateValueSelected & ">" & Year(Date())+3-i & "</option>")
		Next %>
		</select>&nbsp;&nbsp;<span class="fcmp">*</span>
		</td></tr>

		<tr><td width=170><p class="ftxt">Type of Diploma /<br>Degree obtained</td>
		<td bgcolor="<%=colFormBodyRight%>" width=1><img src="x.gif" width=1 height=12></td>
		<td bgcolor="<%=colFormBodyText%>" width=407>
		&nbsp;&nbsp;<select size="1" name="exp_edu_diploma" style="width:355px;">
		<option value="0" selected> Please select </option>
		<%  For i=LBound(arrEduTypeID) to UBound(arrEduTypeID)
		If iExpEduDiplomaID=arrEduTypeID(i) Then
			Response.Write("<option value=" & arrEduTypeID(i) & " selected>" & arrEduTypeTitle(i) & "</option>")
		Else 
			Response.Write("<option value=" & arrEduTypeID(i) & ">" & arrEduTypeTitle(i) & "</option>")
		End If
		Next %>
		</select>&nbsp;&nbsp;<span class="fcmp">*</span>
		</td>
		</tr>
		
		<tr><td width=170><p class="ftxt">If other please specify</td>
		<td bgcolor="<%=colFormBodyRight%>" width=1><img src="x.gif" width=1 height=24></td>
		<td bgcolor="<%=colFormBodyText%>" width=407>&nbsp;&nbsp;<input type="text" name="exp_edu_diploma1" value="<%=sExpEduDiplomaText%>" size="31" style="width:355px;" maxlength="255"></td></tr>
		
		<tr><td width=170><p class="ftxt">Subject</td>
		<td bgcolor="<%=colFormBodyRight%>" width=1><img src="x.gif" width=1 height=24></td>
		<td bgcolor="<%=colFormBodyText%>" width=407>
		&nbsp;&nbsp;<select name="exp_edu_subj" size="1" style="width:355px;">
		<option value="0" selected> Please select </option>
		<%  For i=1 to UBound(arrEduSubjectID) 
		If iExpEduSubjectID=arrEduSubjectID(i) then
			Response.Write("<option value=" & arrEduSubjectID(i) & " selected>" & arrEduSubjectTitle(i) & "</option>")
		Else 
			Response.Write("<option value=" & arrEduSubjectID(i) & ">" & arrEduSubjectTitle(i) & "</option>")
		End If
		Next %>
		</select>&nbsp;&nbsp;<span class="fcmp">*</span>&nbsp;&nbsp;
		</td></tr>

		<tr><td width=170><p class="ftxt">if needed, please specify the exact title of your degree</td>
		<td bgcolor="<%=colFormBodyRight%>" width=1><img src="x.gif" width=1 height=24></td>
		<td bgcolor="<%=colFormBodyText%>" width=407>&nbsp;&nbsp;<input type="text" name="exp_edu_subj1" value="<%=sExpEduSubjectText%>" size="31" style="width:355px;" maxlength="255">

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
	<img src="<% =sHomePath %>image/x.gif" width=170 height=1><a href="javascript:AddEducation(0);"><img src="<% =sHomePath %>image/bte_<% If iExpEduID>0 Then %>save<% Else %>add<% End If %>qualif.gif" name="Add this education" alt="Add this education to the list of your qualifications" border=0></a>
	<input type="hidden" name="id_Edu" value="<%=iExpEduID%>">
	</td>
	<td width=200 height=1 align="right"><a href="javascript:validateForm();"><img src="<% =sHomePath %>image/bte_savecont.gif" name="Continue" alt="Save and continue"  border=0></a></td>
	</tr>
	</form>
	</table><br>
	    
<% CloseDBConnection %>
</body>
</html>
