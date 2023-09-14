<% 
'--------------------------------------------------------------------
'
' CV registration.
' Full format. Experience.
'
'--------------------------------------------------------------------
%>
<!--#include virtual="/_template/asp.header.nocache.asp"-->
<!--#include file="_data/datMonth.asp"-->
<%
If sApplicationName<>"expert" Then
	CheckUserLogin sScriptFullNameAsParams
End If
CheckExpertID()

' Log: 34 - Update expert
If Request.Form()>"" Then
	iLogResult = LogActivity(34, "ExpertID=" & Cstr(iExpertID) & " SavedStep: 4", "", "")
End If

Dim objConnCustom
Set objConnCustom = Server.CreateObject("ADODB.Connection")
objConnCustom.Open "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Data Source=MAGNESA;Initial Catalog=" & objExpertDB.DatabasePath & ";"

' The page is accessible only if expert id is valid
If iExpertID<=0 Then Response.Redirect(sApplicationHomePath & "register.asp" & sParams)

sParams=ReplaceUrlParams(sParams, "prjid")
Dim bExpPrjNewRecord, iExpPrjID, sExpPrjTitle, iExpType, sExpPrjStartDate, sExpPrjEndDate, bExpPrjOngoing, sExpPrjOrganisation, sExpPrjPosition, sExpPrjBeneficiary, sExpPrjReferences, sExpPrjDescription, sExpPrjDonor, sPrjDescription
Dim sExpRefName, sExpRefPosition, sExpRefPhone, sExpRefEmail, sExpRefExtended
Dim sCountries, sDonors, sSectors, iTotalCountries, iTotalDonors, iTotalSectors, bFlagSelected

iExpPrjID=Request.QueryString("prjid")
If IsNumeric(iExpPrjID) And iExpPrjID>"" Then
	iExpPrjID=CLng(iExpPrjID)
Else
	iExpPrjID=0
End If

If IsNumeric(iExpPrjID) And iExpPrjID>"" And sAction="delete" Then
	' Deleting data on projects 	
	objTempRs=UpdateRecordSPWithConn(objConnCustom, "usp_ExpCvvExperienceDelete", Array( _
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
	
	sExpRefName=Left(CheckString(Request.Form("exp_ref_name")), 255)
	sExpRefPosition=Left(CheckString(Request.Form("exp_ref_position")), 255)
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
		objTempRs=UpdateRecordSPWithConn(objConnCustom, "usp_ExpCvvExperienceUpdate", Array( _
		Array(, adVarChar, 3, "Eng"), _
		Array(, adInteger, , iExpertID), _
		Array(, adInteger, , iExpPrjID), _
		Array(, adVarWChar, 255, sExpPrjTitle), _
		Array(, adVarWChar, 200, sExpPrjOrganisation), _
		Array(, adVarWChar, 255, sExpPrjPosition), _
		Array(, adVarWChar, 200, sExpPrjBeneficiary), _
		Array(, adVarWChar, 200, sExpPrjReferences), _
		Array(, adVarWChar, 255, sExpRefName), _
		Array(, adVarWChar, 255, sExpRefPosition), _
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
	objTempRs=GetDataOutParamsSPWithConn(objConnCustom, "usp_ExpCvvExperienceInsert", Array( _
		Array(, adVarChar, 3, "Eng"), _
		Array(, adInteger, , iExpertID), _
		Array(, adVarWChar, 255, sExpPrjTitle), _
		Array(, adVarWChar, 200, sExpPrjOrganisation), _
		Array(, adVarWChar, 255, sExpPrjPosition), _
		Array(, adVarWChar, 200, sExpPrjBeneficiary), _
		Array(, adVarWChar, 200, sExpPrjReferences), _
		Array(, adVarWChar, 255, sExpRefName), _
		Array(, adVarWChar, 255, sExpRefPosition), _
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
	objTempRs=DeleteRecordSPWithConn(objConnCustom, "usp_ExpCvvExperienceCouDelete", Array( _ 
		Array(, adInteger, , iExpertID), _
		Array(, adInteger, , iExpPrjID)))
	Set objTempRs=Nothing
	objTempRs=GetDataOutParamsSPWithConn(objConnCustom, "usp_ExpCvvExperienceCouInsert", Array( _
		Array(, adVarChar, 2000, sCountries), _
		Array(, adInteger, , iExpertID), _
		Array(, adInteger, , iExpPrjID)), _
		Array( Array(, adInteger)))
	iTotalCountries=objTempRs(0)
	Set objTempRs=Nothing

	' Saving donors of interest (for non activated account)
	objTempRs=DeleteRecordSPWithConn(objConnCustom, "usp_ExpCvvExperienceDonDelete", Array( _ 
		Array(, adInteger, , iExpertID), _
		Array(, adInteger, , iExpPrjID)))
	Set objTempRs=Nothing
	objTempRs=GetDataOutParamsSPWithConn(objConnCustom, "usp_ExpCvvExperienceDonInsert", Array( _ 
		Array(, adVarChar, 2000, sDonors), _
		Array(, adInteger, , iExpertID), _
		Array(, adInteger, , iExpPrjID)), _
		Array( Array(, adInteger)))
	iTotalDonors=objTempRs(0)
	Set objTempRs=Nothing


	' Saving sectors of interest (for non activated account)
	objTempRs=DeleteRecordSPWithConn(objConnCustom, "usp_ExpCvvExperienceSctDelete", Array( _ 
		Array(, adInteger, , iExpertID), _
		Array(, adInteger, , iExpPrjID)))
	Set objTempRs=Nothing
	objTempRs=GetDataOutParamsSPWithConn(objConnCustom, "usp_ExpCvvExperienceSctInsert", Array( _ 
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
<!--#include virtual="/_template/html.header.scrolllist.start2.asp"-->
<% InsertJsHelpers 0, 0 %>
<script type="text/javascript" src="/_scripts/js/main.js"></script>
<script language="JavaScript">
<!--
function validateForm() {
	var f = document.RegForm;
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
	var f = document.RegForm;
	f.mmb_cou_hid.value = '0,' + (GetScrollboxSelection("divCouSelector", "cou_") || '0');
	f.mmb_don_hid.value = '0,' + (GetScrollboxSelection("divDonSelector", "don_") || '0');
	f.mmb_sct_hid.value = '0,' + (GetScrollboxSelection("divSctSelector", "sct_") || '0');

	var mmb_cou = (f.mmb_cou_hid.value.split(',')[0] || 0);
	var mmb_sct = (f.mmb_sct_hid.value.split(',')[0] || 0);
	var mmb_don = (f.mmb_don_hid.value.split(',')[0] || 0);

	if (cont_next!=1) {
		f.exp_wke_continue.value="0"; 
	}
<% If sApplicationName<>"external" And sApplicationName<>"backoffice" Then %>
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

	f.submit();
}
// -->
</script>
<script language="JavaScript" src="../../../_scripts/js/asr.js"></script>
<script language="JavaScript" src="../../../_scripts/js/lib.js"></script>
<% InsertScrollStyles %>
<link rel="stylesheet" type="text/css" media="all" href="/scroll.css" />
</head>
<body onload="RestoreInt();" onunload="">
<div align="center">
<div id="holder">
	<!-- header -->
	<!--#include virtual="/_template/page.header.asp"-->

	<!-- content -->
	<div id="content" class="searchform">
	<% 
	If Not bIsMyCV Then 
		%><h2 class="service_title">Curriculum Vitae. <span class="service_slogan">Expert ID: <% =objExpertDB.DatabaseCode %><%=iExpertID%></span></h2><br/>
		<% 
	End If

	ShowRegistrationProgressBar "CV", 4
	%>

	<form method="post" action="<%=sScriptFileName & sParams%>" name="RegForm">
		<div class="box search blue" style="padding-bottom: 0;">
		<h3><span class="left">&nbsp;</span><span class="right">&nbsp;</span><% =GetLabel(sCvLanguage, "Key qualification and specific experience") %></h3>
		<table class="search_form" width="100%" cellspacing=0 cellpadding=0 border=0>
  
	<%
	Set objTempRs=GetDataRecordsetSPWithConn(objConnCustom, "usp_ExpCvvExperienceSelect", Array( _
		Array(, adInteger, , iExpertID)))
	If Not objTempRs.Eof Then %>
	<tr>
	<td colspan=2>
		<table class="results" style="border-left:0; border-right:0;">
		<tr class="tr_results">
		<th class="number"><p>N.</p></td>
		<th class="date"><p><% =GetLabel(sCvLanguage, "Start date") %></p></td>
		<th class="date"><p><% =GetLabel(sCvLanguage, "End date") %></p></td>
		<th width=215><p><% =GetLabel(sCvLanguage, "Project / Organisation") %></p></td>
		<th width=205><p><% =GetLabel(sCvLanguage, "Position") %>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</p></td>
		<th width=15><p><% =GetLabel(sCvLanguage, "Modify") %></p></td>
		<th width=15><p><% =GetLabel(sCvLanguage, "Delete") %></p></td>
		</tr>	
		<% i=1
		While Not objTempRs.Eof %>
		<tr class="tr_results<% If i Mod 2 = 0 Then %> odd<% End If %>">
		<td class="number"><%=i%>.</td>
		<td class="date"><%=ConvertDateForText(objTempRs("wkeStartDate"), "&nbsp;", "MMYYYY") %></td>
		<td class="date">
				<% If objTempRs("wkeEndDateOpen")=1 Then
					Response.Write GetLabel(sCvLanguage, "Ongoing")
					If IsDate(objTempRs("wkeEndDate")) And Not IsNull(objTempRs("wkeEndDate")) Then Response.Write "<br/>("
				End If
				If IsDate(objTempRs("wkeEndDate")) And Not IsNull(objTempRs("wkeEndDate")) Then
					Response.Write ConvertDateForText(objTempRs("wkeEndDate"), "&nbsp;", "MMYYYY")
					If objTempRs("wkeEndDateOpen")=1 Then Response.Write ")"
				End If %>
		</td>
		<td><a href="register3.asp<%=AddUrlParams(sParams, "prjid=" & objTempRs("id_ExpWke"))%>&act=update"><% If objTempRs("id_ExpWke")=iExpPrjID Then %><b><img src="<% =sHomePath %>image/vn_v.gif" width=8 height=12 border=0 hspace=0 align="left"><% End If %><%=CheckSpaces(ReplaceIfEmpty(objTempRs("wkePrjTitleEng"), ReplaceIfEmpty(objTempRs("wkeOrgNameEng"), "Not specified")), 30) %></a></td>
		<td><%=CheckSpaces(objTempRs("wkePositionEng"), 20) %></td>
		<td align="center"><% If objTempRs("id_ExpWke")=iExpPrjID Then %><img src="<% =sHomePath %>image/vn_updte.gif" width=15 height=15 border=0 hspace=0 alt="Updating" align="center"><% Else %><a href="register3.asp<%=AddUrlParams(sParams, "prjid=" & objTempRs("id_ExpWke"))%>&act=update"><img src="<% =sHomePath %>image/vn_updt.gif" width=15 height=15 border=0 hspace=0 alt="Update this record" align="center"></a><% End If %></td>
		<td align="center"><a href="register3.asp<%=AddUrlParams(sParams, "prjid=" & objTempRs("id_ExpWke"))%>&act=delete"><img src="<% =sHomePath %>image/vn_del.gif" width=15 height=15 border=0 hspace=0 alt="Delete this record" align="center"></a></td>
		</tr>
		<% i=i+1
		objTempRs.MoveNext
		WEnd %>
		</table>
	</td>
	</tr>
	<% End If 
	objTempRs.Close %>
	<% If iExpPrjID>0 Then
	On Error Resume Next
	Set objTempRs=GetDataRecordsetSPWithConn(objConnCustom, "usp_ExpCvvExperienceInfoSelect", Array( _
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
		
		sExpRefName=objTempRs("wkeRefName")
		sExpRefPosition=objTempRs("wkeRefPosition")
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
	<input type="hidden" name="exp_wke_continue" value="1">
	<input type="hidden" name="exp_wke_status" value="new">
	<input type="hidden" name="ProjectId" value="<%=iExpPrjID%>">
	<input type="hidden" name="reg_type" value="">
	<input type="hidden" name="Ref" value="">
	<input type="hidden" name="ComeBack" value="no">
	<input type="hidden" name="mmb_cou_hid" value=''>
	<input type="hidden" name="mmb_don_hid" value=''>
	<input type="hidden" name="mmb_sct_hid" value=''>

		<tr>
		<td class="field splitter"><label for="proj_title"><% =GetLabel(sCvLanguage, "Type of experience (Reg)") %></label></td>
		<td class="value blue">
		<p class="txt"><input type="radio" name="exp_type" value=2 <% If iExpType=2 Then %>checked="checked"<% End If %>>&nbsp;&nbsp;Employment</p>
		<p class="txt"><input type="radio" name="exp_type" value=3 <% If iExpType=3 Then %>checked="checked"<% End If %>>&nbsp;&nbsp;Work performed</p>
		</td>
		</tr>
		<tr>
		<tr>
		<td class="field splitter"><label for="proj_title"><% =GetLabel(sCvLanguage, "Project title (Reg)") %></label></td>
		<td class="value blue"><input type="text" id="proj_title" name="proj_title" size=31 style="width:355px;" maxlength=255 value="<%=sExpPrjTitle%>">&nbsp;&nbsp;<span class="fcmp">*</span>&nbsp;</td>
		</tr>
		<tr>
		<td class="field splitter"><label for="proj_title"><% =GetLabel(sCvLanguage, "Main project features (Reg)") %></label></td>
		<td class="value blue"><textarea cols="34" style="width:355px;" name="exp_prj_descr" rows=3 wrap="yes"><%=sPrjDescription%></textarea></td>
		</tr>
		<tr>
		<td class="field splitter"><label for="exp_OrgName"><% =GetLabel(sCvLanguage, "Company / Organisation") %></label></td>
		<td class="value blue"><input type="text" id="exp_OrgName" name="exp_OrgName" size=31 style="width:355px;" maxlength=200 value="<%=sExpPrjOrganisation%>">&nbsp;&nbsp;<span class="fcmp">*</span>&nbsp;</td>
		</tr>
		<tr>
		<td class="field splitter"><label for="exp_smonth"><% =GetLabel(sCvLanguage, "Start date") %></label></td>
		<td class="value blue"><select id="exp_smonth" name="exp_smonth" size=1>
		<option value="0"><% =GetLabel(sCvLanguage, "Month") %></option>
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
		</select>&nbsp;<span class="fcmp">*</span>&nbsp;</td>
		</tr>
		<tr>
		<td class="field splitter"><label for="exp_emonth"><% =GetLabel(sCvLanguage, "End date") %></label></td>
		<td class="value blue"><select id="exp_emonth" name="exp_emonth" size=1>
		<option value="0"><% =GetLabel(sCvLanguage, "Month") %></option>
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
		<tr>
		<td class="field splitter"><label for="exp_position"><% =GetLabel(sCvLanguage, "Position / Responsibility") %></label></td>
		<td class="value blue"><input type="text" id="exp_position" name="exp_position" size=31 style="width:355px;" maxlength=255 value="<%=sExpPrjPosition%>">&nbsp;&nbsp;<span class="fcmp">*</span>&nbsp;</td>
		</tr>
		<tr>
		<td class="field splitter"><label for="exp_benf"><% =GetLabel(sCvLanguage, "Beneficiary") %></label></td>
		<td class="value blue"><input type="text" id="exp_benf" name="exp_benf" size=31 style="width:355px;" maxlength=200 value="<%=sExpPrjBeneficiary%>"></td>
		</tr>
		<tr>
		<td class="field splitter"><label for="exp_ref_name"><% =GetLabel(sCvLanguage, "Reference") %>&nbsp;</nobr>(person)</br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<% =GetLabel(sCvLanguage, "Full name") %></label></td>
		<td class="value blue"><input type="text" id="exp_ref_name" name="exp_ref_name" size=31 style="width:355px;" maxlength=255 value="<%=sExpRefName%>"></td>
		</tr>
		<tr>
		<td class="field splitter"><label for="exp_ref_position">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<% =GetLabel(sCvLanguage, "Position") %></label></td>
		<td class="value blue"><input type="text" id="exp_ref_position" name="exp_ref_position" size=31 style="width:355px;" maxlength=255 value="<%=sExpRefPosition%>"></td>
		</tr>
		<tr>
		<td class="field splitter"><label for="exp_ref_phone">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<% =GetLabel(sCvLanguage, "Phone") %></label></td>
		<td class="value blue"><input type="text" id="exp_ref_phone" name="exp_ref_phone" size=31 style="width:355px;" maxlength=150 value="<%=sExpRefPhone%>"></td>
		</tr>
		<tr>
		<td class="field splitter"><label for="exp_ref_email">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<% =GetLabel(sCvLanguage, "Email") %></label></td>
		<td class="value blue"><input type="text" id="exp_ref_email" name="exp_ref_email" size=31 style="width:355px;" maxlength=150 value="<%=sExpRefEmail%>"></td>
		</tr>
		<tr>
		<td class="field splitter"><label for="exp_clientref"><% =GetLabel(sCvLanguage, "Other references") %></label></td>
		<td class="value blue"><input type="text" id="exp_clientref" name="exp_clientref" size=31 style="width:355px;" maxlength=150 value="<% =sExpPrjReferences %>"></td>
		</tr>
		<tr>
		<td class="field splitter"><label for="exp_wke_descr"><% =GetLabel(sCvLanguage, "Brief description of tasks") %></label></td>
		<td class="value blue"><textarea cols="34" style="width:355px;" id="exp_wke_descr" name="exp_wke_descr" rows=18 wrap="yes"><% =sExpPrjDescription %></textarea>&nbsp;&nbsp;<span class="fcmp">*</span>&nbsp;</td>
		</tr>
		<tr class="last">
		<td class="field splitter"><label for="wke_don_other"><% =GetLabel(sCvLanguage, "Funding agency") %></label></td>
		<td class="value blue"><input type="text" id="wke_don_other" name="wke_don_other" size=31 style="width:355px;" maxlength=255 value="<% =sExpPrjDonor %>"></td>
		</tr>
		</table>
		</div>

		<div class="mini_wraper">
			<%			
			'Funding agancies
			ShowDonScrollBox "Select project funding agency",  "<p class=""fsml2"">(Select funding agency from the list or specify in the field above if it is not in the list)</p>", 1, 0, 0, 0 %><br />
			<%
			' Sectors
			ShowSctScrollBox "Select project sub-sectors",  "", 1, 0, 0 %><br />
			<%
			' Countries
			ShowCouScrollBox "Select project countries",  "", 1, 0, 1, 1, 0 %><br />

		<div class="spacebottom">
		<a href="javascript:AddProject(0);"><img class="button first" src="<% =sHomePath %>image/bte_<% If iExpPrjID>0 Then %>save<% Else %>add<% End If %>project.gif" name="Add this project" alt="Add this project to the list of managed projects"  border=0></a>
		<a href="javascript:validateForm();"><img class="button last" src="<% =sHomePath %>image/bte_savecont.gif" name="Continue" alt="Save and continue"  border=0></a>
		</div>
		</form><br />

	</div>
</div>
</div>
	<!-- footer -->
	<!--#include file="../_template/page.footer.asp"-->
<script language=JavaScript type=text/javascript>
	document.RegForm.mmb_cou_hid.value='<%=mNt & mNtInt %>';
	document.RegForm.mmb_don_hid.value='<%=mOrg & mOrgInt %>';
 	document.RegForm.mmb_sct_hid.value='<%=mExT & mExTInt %>';
</script>
<% CloseDBConnection %>

<% WriteFilterTableScript %>
<link href="/res/css/jquery.mCustomScrollbar.css" rel="stylesheet" type="text/css" />
<script src="/res/scripts/jquery.mCustomScrollbar.concat.min.js"></script>
</body>
<!--#include file="../_template/html.footer.asp"-->

