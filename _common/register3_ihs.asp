<% 
'--------------------------------------------------------------------
'
' CV registration.
' Full format. Experience.
'
'--------------------------------------------------------------------
Response.Expires = 0
Response.ExpiresAbsolute = Now() - 1
Response.CacheControl = "no-cache" 
%>
<!--#include file="_data/datMonth.asp"-->
<%
If sApplicationName<>"expert" Then
	CheckUserLogin sScriptFullNameAsParams
End If
CheckExpertID()
' The page is accessible only if expert id is valid
If iExpertID<=0 Then Response.Redirect(sApplicationHomePath & "register.asp" & sParams)

sParams=ReplaceUrlParams(sParams, "prjid")
Dim bExpPrjNewRecord, iExpPrjID, sExpPrjTitle, iExpType, sExpPrjStartDate, sExpPrjEndDate, bExpPrjOngoing, sExpPrjOrganisation, sExpPrjPosition, sExpPrjBeneficiary, sExpPrjReferences, sExpPrjDescription, sExpPrjDonor, sPrjDescription
Dim sExpRefFirstName, sExpRefLastName, sExpRefPhone, sExpRefEmail, sExpRefExtended
Dim sCountries, sDonors, sSectors, iTotalCountries, iTotalDonors, iTotalSectors, bFlagSelected

iExpPrjID=Request.QueryString("prjid")
If IsNumeric(iExpPrjID) And iExpPrjID>"" Then
	iExpPrjID=CLng(iExpPrjID)
Else
	iExpPrjID=0
End If

If IsNumeric(iExpPrjID) And iExpPrjID>"" And sAction="delete" Then
	' Deleting data on projects 	
	objTempRs=UpdateRecordSP("usp_ExpCvvExperienceDelete", Array( _
		Array(, adInteger, , iExpertID), _
		Array(, adInteger, , iExpPrjID)))
	Response.Redirect(sScriptFileName & ReplaceUrlParams(sParams, "prjid"))
End If

If Request.Form()>"" Then

	iExpPrjID=CheckString(Request.Form("ProjectId"))
	sExpPrjTitle=Left(CheckString(Request.Form("proj_title")), 255)
	iExpType=CheckIntegerAndNull(Request.Form("exp_type"))
	iExpType=ReplaceIfEmpty(iExpType, 1)
	sExpPrjStartDate=ConvertDMYForSQL(CheckString(Request.Form("exp_syear")), CheckString(Request.Form("exp_smonth")), 1)
	sExpPrjEndDate=ConvertDMYForSQL(CheckString(Request.Form("exp_eyear")), CheckString(Request.Form("exp_emonth")), 1)
	If Request.Form("exp_e_ongoing")="on" Then
		bExpPrjOngoing=1
	Else
		bExpPrjOngoing=0
	End If
	
	sExpPrjOrganisation=Left(CheckString(Request.Form("exp_OrgName")), 200)
	sExpPrjPosition=Left(CheckString(Request.Form("exp_position")), 255)
	sExpPrjBeneficiary=Left(CheckString(Request.Form("exp_benf")), 200)
	sExpPrjReferences=Left(CheckString(Request.Form("exp_clientref")), 200)
	
	sExpRefFirstName=Left(CheckString(Request.Form("exp_ref_firstname")), 255)
	sExpRefLastName=Left(CheckString(Request.Form("exp_ref_lastname")), 255)
	sExpRefPhone=Left(CheckString(Request.Form("exp_ref_phone")), 150)
	sExpRefEmail=Left(CheckString(Request.Form("exp_ref_email")), 150)
	sExpRefExtended=CheckString(Request.Form("exp_ref_extended"))
	If sExpRefExtended="on" Then
		sExpRefExtended=1
	Else 
		sExpRefExtended=Null
	End If
	
	sPrjDescription=CheckString(Request.Form("exp_prj_descr"))
	sExpPrjDescription=CheckString(Request.Form("exp_wke_descr"))
	sExpPrjDonor=Left(CheckString(Request.Form("wke_don_other")), 255)

	bExpPrjNewRecord=Len(Trim(sExpPrjTitle)) + Len(Trim(sExpPrjOrganisation)) + Len(Trim(sExpPrjPosition)) + Len(Trim(sExpPrjBeneficiary)) + Len(Trim(sExpPrjReferences)) + Len(Trim(sExpPrjDescription)) + Len(Trim(sExpPrjDonor)) + Len(ReplaceIfEmpty(Trim(sExpPrjStartDate), "")) + Len(ReplaceIfEmpty(Trim(sExpPrjEndDate), ""))

	If IsNumeric(iExpPrjID) And iExpPrjID>0 Then
		objTempRs=UpdateRecordSP("usp_ExpCvvExperienceUpdate", Array( _
		Array(, adVarChar, 3, "Eng"), _
		Array(, adInteger, , iExpertID), _
		Array(, adInteger, , iExpPrjID), _
		Array(, adVarWChar, 255, sExpPrjTitle), _
		Array(, adVarWChar, 200, sExpPrjOrganisation), _
		Array(, adVarWChar, 255, sExpPrjPosition), _
		Array(, adVarWChar, 200, sExpPrjBeneficiary), _
		Array(, adVarWChar, 200, sExpPrjReferences), _
		Array(, adVarWChar, 255, sExpRefFirstName), _
		Array(, adVarWChar, 255, sExpRefLastName), _
		Array(, adVarChar, 150, sExpRefPhone), _
		Array(, adVarWChar, 150, sExpRefEmail), _
		Array(, adInteger, , sExpRefExtended), _
		Array(, adLongVarChar, 20000, sExpPrjDescription), _
		Array(, adLongVarChar, 20000, sPrjDescription), _
		Array(, adVarWChar, 255, sExpPrjDonor), _
		Array(, adVarChar, 16, sExpPrjStartDate), _
		Array(, adVarChar, 16, sExpPrjEndDate),_
		Array(, adInteger, , bExpPrjOngoing),_
		Array(, adInteger, , iExpType)))

	ElseIf bExpPrjNewRecord>0 Then 

	' Saving project data
	objTempRs=GetDataOutParamsSP("usp_ExpCvvExperienceInsert", Array( _
		Array(, adVarChar, 3, "Eng"), _
		Array(, adInteger, , iExpertID), _
		Array(, adVarWChar, 255, sExpPrjTitle), _
		Array(, adVarWChar, 200, sExpPrjOrganisation), _
		Array(, adVarWChar, 255, sExpPrjPosition), _
		Array(, adVarWChar, 200, sExpPrjBeneficiary), _
		Array(, adVarWChar, 200, sExpPrjReferences), _
		Array(, adVarWChar, 255, sExpRefFirstName), _
		Array(, adVarWChar, 255, sExpRefLastName), _
		Array(, adVarChar, 150, sExpRefPhone), _
		Array(, adVarWChar, 150, sExpRefEmail), _
		Array(, adInteger, , sExpRefExtended), _
		Array(, adLongVarChar, 20000, sExpPrjDescription), _
		Array(, adLongVarChar, 20000, sPrjDescription), _
		Array(, adVarWChar, 255, sExpPrjDonor), _
		Array(, adVarChar, 16, sExpPrjStartDate), _
		Array(, adVarChar, 16, sExpPrjEndDate),_
		Array(, adInteger, , bExpPrjOngoing),_
		Array(, adInteger, , iExpType)), _
		Array( Array(, adInteger)))
	iExpPrjID=objTempRs(0)	
	Set objTempRs=Nothing
	End If

	sCountries=CheckString(Request.Form("mmb_cou_hid"))
	sDonors=CheckString(Request.Form("mmb_don_hid"))
	sSectors=CheckString(Request.Form("mmb_sct_hid"))

	' Removing the number of total selected items in every field
	sCountries=Mid(sCountries, InStr(sCountries,",")+1,Len(sCountries))
	sDonors=Mid(sDonors, InStr(sDonors,",")+1,Len(sDonors))
	sSectors=Mid(sSectors, InStr(sSectors,",")+1,Len(sSectors))
	                      

	' Saving countries of interest (for non activated account)
	objTempRs=DeleteRecordSP("usp_ExpCvvExperienceCouDelete", Array( _ 
		Array(, adInteger, , iExpertID), _
		Array(, adInteger, , iExpPrjID)))
	Set objTempRs=Nothing
	objTempRs=GetDataOutParamsSP("usp_ExpCvvExperienceCouInsert", Array( _
		Array(, adVarChar, 2000, sCountries), _
		Array(, adInteger, , iExpertID), _
		Array(, adInteger, , iExpPrjID)), _
		Array( Array(, adInteger)))
	iTotalCountries=objTempRs(0)
	Set objTempRs=Nothing

	' Saving donors of interest (for non activated account)
	objTempRs=DeleteRecordSP("usp_ExpCvvExperienceDonDelete", Array( _ 
		Array(, adInteger, , iExpertID), _
		Array(, adInteger, , iExpPrjID)))
	Set objTempRs=Nothing
	objTempRs=GetDataOutParamsSP("usp_ExpCvvExperienceDonInsert", Array( _ 
		Array(, adVarChar, 2000, sDonors), _
		Array(, adInteger, , iExpertID), _
		Array(, adInteger, , iExpPrjID)), _
		Array( Array(, adInteger)))
	iTotalDonors=objTempRs(0)
	Set objTempRs=Nothing


	' Saving sectors of interest (for non activated account)
	objTempRs=DeleteRecordSP("usp_ExpCvvExperienceSctDelete", Array( _ 
		Array(, adInteger, , iExpertID), _
		Array(, adInteger, , iExpPrjID)))
	Set objTempRs=Nothing
	objTempRs=GetDataOutParamsSP("usp_ExpCvvExperienceSctInsert", Array( _ 
		Array(, adVarChar, 2000, sSectors), _
		Array(, adInteger, , iExpertID), _
		Array(, adInteger, , iExpPrjID)), _
		Array( Array(, adInteger)))
	iTotalSectors=objTempRs(0)
	Set objTempRs=Nothing

	If Request.Form("exp_wke_continue")="0" then
		Response.Redirect "register3.asp" & sParams
	Else
		Response.Redirect "register4.asp" & sParams
        End if   
End If
%>

<html>
<head>
<title><% =GetLabel(sCvLanguage, "CV registration") %>. <% =GetLabel(sCvLanguage, "Professional experience") %></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<% InsertScrollStyles %>
<link rel="stylesheet" type="text/css" href="<% =sHomePath %>styles.css">
<script type="text/javascript" src="/_scripts/js/main.js"></script>
<% InsertJSScrollFunctions 0, 0 %>
<script language="JavaScript" src="/_scripts/js/asr.js"></script>
<script language="JavaScript" src="/_scripts/js/lib.js"></script>
<script language="JavaScript">
<!--
function validateForm() {
	var f=document.RegForm;
<% If Len(sBackOffice)<3 Then %>
	if (f.proj_title.value!="" || f.exp_OrgName.value!="") {
		AddProject(1);
	} else 
<% End If %>
	{ f.submit(); }
}

function AddProject(cont_next) {
<%
Dim sUserSalutation
If sApplicationName="expert" Then
	sUserSalutation="your"
Else
	sUserSalutation="expert's"
End If
%>
var f=document.RegForm;
	if (cont_next!=1) {
		f.exp_wke_continue.value="0"; 
	}
<% If sApplicationName<>"external" Then %>
<% 
On Error Resume Next
	BeforeClientValidationCvRegistrationStep3
On Error GoTo 0
%>
<% If Len(sBackOffice)<3 Then %>
    if (f.proj_title.value=="" && f.exp_OrgName.value=="") {
		alert("<% =GetLabel(sCvLanguage, "Please specify the project title or the name of the company or organisation") %>"); f.proj_title.select(); return;
	}
	var start_month=f.exp_smonth.options[f.exp_smonth.selectedIndex].value;
	var start_year=f.exp_syear.options[f.exp_syear.selectedIndex].value;
	var end_month=f.exp_emonth.options[f.exp_emonth.selectedIndex].value;
	var end_year=f.exp_eyear.options[f.exp_eyear.selectedIndex].value;
	var end_ongoing=(f.exp_e_ongoing.checked?1:0);
	if (start_month==0 || start_year==0) {
		alert("Please fill in the experience start date."); return;
	}
	if ((end_month==0 || end_year==0) && (end_ongoing==0)) {
		alert("Please fill in the experience end date or tick a checkbox for ongoing experience."); return;
	}
	var start_date = new Date(start_year, start_month, 1);
	var end_date = new Date(end_year, end_month, 28);
	if ((start_date>end_date) && (end_ongoing==0)) {
		alert("Please fill in the experience dates properly."); return;
	}
	if (!checkTextFieldValue(f.exp_position, "", "Please fill in <% =sUserSalutation %> position.", 1)) { return }
	
	if (!checkTextFieldLength(f.exp_prj_descr, 25000, "<% =GetLabel(sCvLanguage, "Please make text of main project features shorter") %>", 1)) { return }
	if (!checkTextFieldValue(f.exp_wke_descr, "", "<% =GetLabel(sCvLanguage, "Please fill in a description of the tasks assigned") %>", 1)) { return }
	if (!checkTextFieldLength(f.exp_wke_descr, 25000, "<% =GetLabel(sCvLanguage, "Please make text of a description of the tasks assigned shorter") %>", 1)) { return }
	
    if(mmb_cou<1)  { alert("<% =GetLabel(sCvLanguage, "Please select at least one country") %>"); return; }
    if(mmb_cou>30) { alert("<% =GetLabel(sCvLanguage, "You cannot select more than 30 countries for one project") %>"); return; }
    if(mmb_sct<1)  { alert("<% =GetLabel(sCvLanguage, "Please select at least one sub-sector of expertise") %>"); return; }
    if(mmb_sct>50) { alert("<% =GetLabel(sCvLanguage, "You cannot select more than 50 sectors for one project") %>"); return; }
<% 
On Error Resume Next
	AfterClientValidationCvRegistrationStep3
On Error GoTo 0
%>
<% End If %>
<% End If %>

	f.mmb_cou_hid.value=mmb_cou+jNtInt;
	f.mmb_don_hid.value=mmb_don+jOrgInt;
	f.mmb_sct_hid.value=mmb_sct+jExTInt;

	f.submit();
}
// -->
</script>
</head>

<body bgcolor="#FFFFFF" topmargin=0 leftmargin=0 marginheight=0 marginwidth=0 onLoad="RestoreInt();">
<% ShowTopMenu %>
<% ShowRegistrationProgressBar "CV", 4 %>

<br />
  <!-- Key qualifications -->
	<table width=580 cellspacing=0 cellpadding=0 border=0 align="center">
	<tr height=1><td width=1 bgcolor="<%=colFormHeaderTop%>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td>
	<td width=578 bgcolor="<%=colFormHeaderTop%>"><img src="<% =sHomePath %>image/x.gif" width=498 height=1></td>
	<td width=1 bgcolor="<%=colFormHeaderBottom%>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td></tr>
	<tr height=18><td width=1 bgcolor="<%=colFormHeaderTop%>"><img src="<% =sHomePath %>image/x.gif" width=1 height=18></td>
	<td width=578 bgcolor="<%=colFormHeaderMain%>"><img src="<% =sHomePath %>image/x.gif" width=498 height=1><br><p class="fttl"><img src="<% =sHomePath %>image/<% =imgFormBullet %>" width=7 height=7 align=left vspace=3 hspace=8><% =GetLabel(sCvLanguage, "KEY QUALIFICATION AND SPECIFIC EXPERIENCE") %></p></td>
	<td width=1 bgcolor="<%=colFormHeaderBottom%>"><img src="<% =sHomePath %>image/x.gif" width=1 height=18></td></tr>
	<tr height=1><td width=1 bgcolor="<%=colFormHeaderTop%>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td>
	<td width=578 bgcolor="<%=colFormHeaderBottom%>"><img src="<% =sHomePath %>image/x.gif" width=498 height=1></td>
	<td width=1 bgcolor="<%=colFormHeaderBottom%>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td></tr>
	<tr height=1><td width=1 bgcolor="<% =colFormHeaderSplitter %>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td>
	<td width=578 bgcolor="<% =colFormHeaderSplitter %>"><img src="<% =sHomePath %>image/x.gif" width=498 height=1></td>
	<td width=1 bgcolor="<% =colFormHeaderSplitter %>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td></tr>

	<%
	Set objTempRs=GetDataRecordsetSP("usp_ExpCvvExperienceSelect", Array( _
		Array(, adInteger, , iExpertID)))
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
			<td width=60 bgcolor="<%=colFormHeaderTop%>"><p><% =GetLabel(sCvLanguage, "Start date") %></p></td>
			<td width=60 bgcolor="<%=colFormHeaderTop%>"><p><% =GetLabel(sCvLanguage, "End date") %></p></td>
			<td width=215 bgcolor="<%=colFormHeaderTop%>"><p><% =GetLabel(sCvLanguage, "Project / Organisation") %></p></td>
			<td width=205 bgcolor="<%=colFormHeaderTop%>"><p><% =GetLabel(sCvLanguage, "Position") %>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</p></td>
			<td width=15 bgcolor="<%=colFormHeaderTop%>"><p><% =GetLabel(sCvLanguage, "Modify") %></p></td>
			<td width=15 bgcolor="<%=colFormHeaderTop%>"><p><% =GetLabel(sCvLanguage, "Delete") %></p></td>
		</tr>	
		<% i=1
		While Not objTempRs.Eof %>
		  <tr height=20>
		    <td bgcolor="<%=colFormBodyText%>"><p align="center"><%=i%></td>
		    <td bgcolor="<%=colFormBodyText%>"><p><%=ConvertDateForText(objTempRs("wkeStartDate"), "&nbsp;", "MMYYYY") %></td>
		    <td bgcolor="<%=colFormBodyText%>"><p>
				<% If objTempRs("wkeEndDateOpen")=1 Then
					Response.Write GetLabel(sCvLanguage, "Ongoing")
					If IsDate(objTempRs("wkeEndDate")) And Not IsNull(objTempRs("wkeEndDate")) Then Response.Write "<br/>("
				End If
				If IsDate(objTempRs("wkeEndDate")) And Not IsNull(objTempRs("wkeEndDate")) Then
					Response.Write ConvertDateForText(objTempRs("wkeEndDate"), "&nbsp;", "MMYYYY")
					If objTempRs("wkeEndDateOpen")=1 Then Response.Write ")"
				End If %>
			</td>
		    <td bgcolor="<%=colFormBodyText%>"><p><a href="register3.asp<%=AddUrlParams(sParams, "prjid=" & objTempRs("id_ExpWke"))%>&act=update"><% If objTempRs("id_ExpWke")=iExpPrjID Then %><b><img src="<% =sHomePath %>image/vn_v.gif" width=8 height=12 border=0 hspace=0 align="left"><% End If %><%=CheckSpaces(ReplaceIfEmpty(objTempRs("wkePrjTitleEng"), ReplaceIfEmpty(objTempRs("wkeOrgNameEng"), "Not specified")), 30) %></a></td>
		    <td bgcolor="<%=colFormBodyText%>"><p><%=CheckSpaces(objTempRs("wkePositionEng"), 20) %></td>
		    <td bgcolor="<%=colFormBodyText%>" align="center"><% If objTempRs("id_ExpWke")=iExpPrjID Then %><img src="<% =sHomePath %>image/vn_updte.gif" width=15 height=15 border=0 hspace=0 alt="Updating" align="center"><% Else %><a href="register3.asp<%=AddUrlParams(sParams, "prjid=" & objTempRs("id_ExpWke"))%>&act=update"><img src="<% =sHomePath %>image/vn_updt.gif" width=15 height=15 border=0 hspace=0 alt="Update this record" align="center"></a><% End If %></td>
		    <td bgcolor="<%=colFormBodyText%>" align="center"><a href="register3.asp<%=AddUrlParams(sParams, "prjid=" & objTempRs("id_ExpWke"))%>&act=delete"><img src="<% =sHomePath %>image/vn_del.gif" width=15 height=15 border=0 hspace=0 alt="Delete this record" align="center"></a></td>
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
	objTempRs.Close %>

	<form method="post" action="<%=sScriptFileName & sParams%>" name="RegForm">
	<input type="hidden" name="exp_wke_continue" value="1">
	<% If iExpPrjID>0 Then
	On Error Resume Next
	Set objTempRs=GetDataRecordsetSP("usp_ExpCvvExperienceInfoSelect", Array( _
		Array(, adInteger, , iExpertID), _
		Array(, adInteger, , iExpPrjID)))

	If Not objTempRs.Eof Then
		iExpType=objTempRs("TypeofWke")
		sExpPrjTitle=objTempRs("wkePrjTitleEng")
		sExpPrjStartDate=objTempRs("wkeStartDate")
		sExpPrjEndDate=objTempRs("wkeEndDate")
		bExpPrjOngoing=objTempRs("wkeEndDateOpen")
		sExpPrjOrganisation=objTempRs("wkeOrgNameEng")
		sExpPrjPosition=objTempRs("wkePositionEng")
		sExpPrjBeneficiary=objTempRs("wkeBnfNameEng")
		
		sExpRefFirstName=objTempRs("wkeRefFirstName")
		sExpRefLastName=objTempRs("wkeRefLastName")
		sExpRefPhone=objTempRs("wkeRefPhone")
		sExpRefEmail=objTempRs("wkeRefEmail")
		sExpRefExtended=objTempRs("wkeRefExtended")
		
		sExpPrjReferences=objTempRs("wkeClientRefEng")
		sExpPrjDescription=objTempRs("wkeDescriptionEng")
		sPrjDescription=objTempRs("wkeProjectDescription")
		sExpPrjDonor=objTempRs("wkeDonorEng")
	End If
	objTempRs.Close
	On Error GoTo 0
	Set objTempRs=Nothing %>
<!--#include file="../_common/expWorkFields.asp"-->
	<% End If %>

	<input type="hidden" name="exp_wke_status" value="new">
	<input type="hidden" name="ProjectId" value="<%=iExpPrjID%>">
	<input type="hidden" name="reg_type" value="">
	<input type="hidden" name="Ref" value="">
	<input type="hidden" name="ComeBack" value="no">
	<input type="hidden" name="mmb_cou_hid" value=''>
	<input type="hidden" name="mmb_don_hid" value=''>
	<input type="hidden" name="mmb_sct_hid" value=''>

   	<tr><td width=1 bgcolor="<%=colFormBodyText%>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1><br></td>
	<td width=578 bgcolor="<%=colFormHeaderTop%>" valign="top">
		<table width="100%" cellspacing=0 cellpadding=0 border=0>
		<tr><td width=170><p class="ftxt"><% =GetLabel(sCvLanguage, "Type of experience") %></td>
		<td bgcolor="<%=colFormBodyRight%>" width=1><img src="x.gif" width=1 height=36></td>
		<td bgcolor="<%=colFormBodyText%>" width=407><img src="x.gif" width=1 height=6><br>
		<p class="txt"><input type="radio" name="exp_type" value=2 <% If iExpType=2 Then %>checked="checked"<% End If %>>&nbsp;&nbsp;Employment</p>
		<p class="txt"><input type="radio" name="exp_type" value=3 <% If iExpType=3 Then %>checked="checked"<% End If %>>&nbsp;&nbsp;Work performed</p>
		</td></tr>

		<tr><td width=170><p class="ftxt"><% =GetLabel(sCvLanguage, "Project title (Reg)") %></td>
		<td bgcolor="<%=colFormBodyRight%>" width=1><img src="x.gif" width=1 height=36></td>
		<td bgcolor="<%=colFormBodyText%>" width=407>
		&nbsp;&nbsp;<input type="text" name="proj_title" size=31 style="width:355px;" maxlength=255 value="<%=sExpPrjTitle%>">&nbsp;&nbsp;<span class="fcmp">*</span>&nbsp;&nbsp;</td></tr>

		<tr><td width=170 valign="top"><p class="ftxt"><% =GetLabel(sCvLanguage, "Main project features (Reg)") %></td>
		<td bgcolor="<%=colFormBodyRight%>" width=1><img src="x.gif" width=1 height=72></td>
		<td bgcolor="<%=colFormBodyText%>" width=407>&nbsp;&nbsp;<textarea cols="34" style="width:355px;" name="exp_prj_descr" rows=3 wrap="yes"><%=sPrjDescription%></textarea></td></tr>
		
		<tr><td width=170><p class="ftxt"><% =GetLabel(sCvLanguage, "Company / Organisation") %></td>
		<td bgcolor="<%=colFormBodyRight%>" width=1><img src="x.gif" width=1 height=24></td>
		<td bgcolor="<%=colFormBodyText%>" width=407>&nbsp;&nbsp;<input type="text" name="exp_OrgName" size=31 style="width:355px;" maxlength=200 value="<%=sExpPrjOrganisation%>">&nbsp;&nbsp;<span class="fcmp">*</span>&nbsp;&nbsp;</td></tr>

		<tr><td width=170><p class="ftxt"><% =GetLabel(sCvLanguage, "Start date") %></td>
		<td bgcolor="<%=colFormBodyRight%>" width=1><img src="x.gif" width=1 height=24></td>
		<td bgcolor="<%=colFormBodyText%>" width=407>
		&nbsp;&nbsp;<select name="exp_smonth" size=1>
		<option value="0" selected><% =GetLabel(sCvLanguage, "Month") %></option>
		<% For i=1 to UBound(arrMonthID)
		If IsDate(sExpPrjStartDate) Then
			If Month(sExpPrjStartDate)=arrMonthID(i) Then
				bFlagSelected=" selected"
			Else
				bFlagSelected=""
			End If
		End If	
		Response.Write("<option value=""" & arrMonthID(i) & """" & bFlagSelected & ">" & arrMonthName(i) & "</option>")
		Next %>
		</select>
		<select name="exp_syear" size="1">
		<option value="0"><% =GetLabel(sCvLanguage, "Year") %></option>
		<% For i=-1 To 60
		If IsDate(sExpPrjStartDate) Then
			If Year(sExpPrjStartDate)=Year(Date())-i then
				bFlagSelected=" selected"
			Else
				bFlagSelected=""
			End If
		End If	
		Response.Write("<option value=""" & Year(Date())-i & """" & bFlagSelected & ">" & (Year(Date())-i) & "</option>")
		Next %>
		</select>&nbsp;<span class="fcmp">*</span>&nbsp;&nbsp;
		</td></tr>

		<tr><td width=170><p class="ftxt"><% =GetLabel(sCvLanguage, "End date") %></td>
		<td bgcolor="<%=colFormBodyRight%>" width=1><img src="x.gif" width=1 height=24></td>
		<td bgcolor="<%=colFormBodyText%>" width=407>
		&nbsp;&nbsp;<select name="exp_emonth" size=1>
		<option value="0" selected><% =GetLabel(sCvLanguage, "Month") %></option>
		<% For i=1 to UBound(arrMonthID)
		If IsDate(sExpPrjEndDate) Then
			If Month(sExpPrjEndDate)=arrMonthID(i) Then
				bFlagSelected=" selected"
			Else
				bFlagSelected=""
			End If
		End If	
		Response.Write("<option value=""" & arrMonthID(i) & """" & bFlagSelected & ">" & arrMonthName(i) & "</option>")
		Next %>
		</select>
		<select name="exp_eyear" size="1">
		<option value="0"><% =GetLabel(sCvLanguage, "Year") %></option>
		<% For i=-3 To 60
		If IsDate(sExpPrjStartDate) Then
			If Year(sExpPrjEndDate)=Year(Date())-i then
				bFlagSelected=" selected"
			Else
				bFlagSelected=""
			End If
		End If	
		Response.Write("<option value=""" & Year(Date())-i & """" & bFlagSelected & ">" & (Year(Date())-i) & "</option>")
		Next %>
		</select>&nbsp; &nbsp; &nbsp;
		<input type="checkbox" name="exp_e_ongoing" id="exp_e_ongoing" <% If bExpPrjOngoing=1 Then %> checked<% End If %>>&nbsp;<span class="ftxt"><label for="exp_e_ongoing"><% =GetLabel(sCvLanguage, "Ongoing") %></label></span>
		</td></tr>

		<tr><td width=170><p class="ftxt"><% =GetLabel(sCvLanguage, "Position / Responsibility") %></td>
		<td bgcolor="<%=colFormBodyRight%>" width=1><img src="x.gif" width=1 height=24></td>
		<td bgcolor="<%=colFormBodyText%>" width=407>&nbsp;&nbsp;<input type="text" name="exp_position" size=31 style="width:355px;" maxlength=255 value="<%=sExpPrjPosition%>">&nbsp;&nbsp;<span class="fcmp">*</span>&nbsp;&nbsp;</td></tr>

		<tr><td width=170><p class="ftxt"><% =GetLabel(sCvLanguage, "Beneficiary") %></td>
		<td bgcolor="<%=colFormBodyRight%>" width=1><img src="x.gif" width=1 height=24></td>
		<td bgcolor="<%=colFormBodyText%>" width=407>&nbsp;&nbsp;<input type="text" name="exp_benf" size=31 style="width:355px;" maxlength=200 value="<%=sExpPrjBeneficiary%>">

		<!--
		<tr><td width=170><p class="ftxt"><% =GetLabel(sCvLanguage, "Client references") %></td>
		<td bgcolor="<%=colFormBodyRight%>" width=1><img src="x.gif" width=1 height=24></td>
		<td bgcolor="<%=colFormBodyText%>" width=407>&nbsp;&nbsp;<input type="text" name="exp_clientRef" size=31 style="width:355px;" maxlength=200 value="<%=sExpPrjReferences%>"></td></tr>
		-->
		<tr><td width=170><p class="ftxt"><% =GetLabel(sCvLanguage, "Reference") %>&nbsp;</nobr>(contact&nbsp;person)</br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<% =GetLabel(sCvLanguage, "First name") %></td>
		<td bgcolor="<%=colFormBodyRight%>" width=1><img src="x.gif" width=1 height=24></td>
		<td bgcolor="<%=colFormBodyText%>" width=407>&nbsp;&nbsp;<input type="text" name="exp_ref_firstname" size=31 style="width:355px;" maxlength=255 value="<%=sExpRefFirstName%>"></td></tr>

		<tr><td width=170><p class="ftxt">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<% =GetLabel(sCvLanguage, "Last name") %></td>
		<td bgcolor="<%=colFormBodyRight%>" width=1><img src="x.gif" width=1 height=24></td>
		<td bgcolor="<%=colFormBodyText%>" width=407>&nbsp;&nbsp;<input type="text" name="exp_ref_lastname" size=31 style="width:355px;" maxlength=255 value="<%=sExpRefLastName%>"></td></tr>

		<tr><td width=170><p class="ftxt">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<% =GetLabel(sCvLanguage, "Phone") %></td>
		<td bgcolor="<%=colFormBodyRight%>" width=1><img src="x.gif" width=1 height=24></td>
		<td bgcolor="<%=colFormBodyText%>" width=407>&nbsp;&nbsp;<input type="text" name="exp_ref_phone" size=31 style="width:355px;" maxlength=150 value="<%=sExpRefPhone%>"></td></tr>

		<tr><td width=170><p class="ftxt">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<% =GetLabel(sCvLanguage, "Email") %></td>
		<td bgcolor="<%=colFormBodyRight%>" width=1><img src="x.gif" width=1 height=24></td>
		<td bgcolor="<%=colFormBodyText%>" width=407>&nbsp;&nbsp;<input type="text" name="exp_ref_email" size=31 style="width:355px;" maxlength=150 value="<%=sExpRefEmail%>">	</td></tr>

		<tr><td width=170 valign="top"><p class="ftxt"><% =GetLabel(sCvLanguage, "Brief description of tasks") %></td>
		<td bgcolor="<%=colFormBodyRight%>" width=1><img src="x.gif" width=1 height=95></td>
		<td bgcolor="<%=colFormBodyText%>" width=407>&nbsp;&nbsp;<textarea cols="34" style="width:355px;" name="exp_wke_descr" rows=5 wrap="yes"><%=sExpPrjDescription%></textarea>&nbsp;&nbsp;<span class="fcmp">*</span>&nbsp;&nbsp;</td></tr>

		<tr><td width=170><p class="ftxt"><% =GetLabel(sCvLanguage, "Funding agency") %></td>
		<td bgcolor="<%=colFormBodyRight%>" width=1><img src="x.gif" width=1 height=24></td>
		<td bgcolor="<%=colFormBodyText%>" width=407>&nbsp;&nbsp;<input type="text" name="wke_don_other" size=31 style="width:355px;" maxlength=255 value="<%=sExpPrjDonor%>">

		<img src="<% =sHomePath %>image/x.gif" width=358 height=1 vspace=2><br></td></tr>
		</table>
	</td>
	<td width=1 bgcolor="<%=colFormBodyRight%>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td></tr>

	<tr height=1><td width=1 bgcolor="<%=colFormBodyText%>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td>
	<td width=578 bgcolor="<%=colFormHeaderMain%>"><img src="<% =sHomePath %>image/x.gif" width=498 height=1></td>
	<td width=1 bgcolor="<%=colFormBodyRight%>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td></tr>
	<tr height=1><td width=1 bgcolor="<% =colFormHeaderSplitter %>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td>
	<td width=578 bgcolor="<% =colFormHeaderSplitter %>"><img src="<% =sHomePath %>image/x.gif" width=498 height=1></td>
	<td width=1 bgcolor="<% =colFormHeaderSplitter %>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td></tr>
	</table>

<!-- Funding agancies section -->
<% ShowDonScrollBox "",  "(" & GetLabel(sCvLanguage, "Select funding agency from list") & ")", 0, 0, 1, 1, 0, 0 %><br>

<!-- Countries section -->
<% ShowCouScrollBox GetLabel(sCvLanguage, "SELECT PROJECT'S COUNTRIES"),  "", 1, 0, 1, 1, 0 %><br>
	
<!-- Sectors section -->
<% ShowSctScrollBox GetLabel(sCvLanguage, "SELECT PROJECT'S SUB-SECTORS"),  "", 1, 0, 1, 1, 0 %><br>

	<!-- Add project button -->
	<table width=576 cellspacing=0 cellpadding=0 border=0 align="center">
	<tr><td>
	<img src="<% =sHomePath %>image/x.gif" width=170 height=1><a href="javascript:AddProject(0);"><img src="<% =sHomePath %>image/bte_<% If iExpPrjID>0 Then %>save<% Else %>add<% End If %>project.gif" name="Add this project" alt="Add this project to the list of managed projects"  border=0></a>
	</td>
	<td width="*" height=1 align="right"><a href="javascript:validateForm();"><img src="<% =sHomePath %>image/bte_savecont.gif" name="Continue" alt="Save & Continue" border=0></a></td>
	</table><br>
	</form>

<script language=JavaScript type=text/javascript>
scrollInit(1,1,1);
document.RegForm.mmb_cou_hid.value='<%=mNt & mNtInt %>';
document.RegForm.mmb_don_hid.value='<%=mOrg & mOrgInt %>';
document.RegForm.mmb_sct_hid.value='<%=mExT & mExTInt %>';
</script>

<% CloseDBConnection %>
</body>
</html>
