<%
'--------------------------------------------------------------------
'
' CV registration.
' Full format. Education.
'
'--------------------------------------------------------------------
%>
<!--#include virtual="/_template/asp.header.nocache.asp"-->
<!--#include file="_data/datEduSubject.asp"-->
<!--#include file="_data/datEduType.asp"-->
<!--#include file="_data/datMonth.asp"-->
<%
If sApplicationName<>"expert" Then
	CheckUserLogin sScriptFullNameAsParams
End If
CheckExpertID()

' Log: 34 - Update expert
If Request.Form()>"" Then
	iLogResult = LogActivity(34, "ExpertID=" & Cstr(iExpertID) & " SavedStep: 2", "", "")
End If

Dim objConnCustom
Set objConnCustom = Server.CreateObject("ADODB.Connection")
objConnCustom.Open "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Data Source=MAGNESA;Initial Catalog=" & objExpertDB.DatabasePath & ";"

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
	objTempRs=UpdateRecordSPWithConn(objConnCustom, "usp_ExpertEducationDelete", Array( _
		Array(, adInteger, , iExpertID), _
		Array(, adInteger, , iExpEduID)))
	Response.Redirect(sScriptFileName & ReplaceUrlParams(sParams, "eduid"))
End If

If Request.Form()>"" Then
	iExpEduID=CheckString(Request.Form("id_Edu"))
	sExpEduInstitution=CheckString(Request.Form("exp_inst_name"))
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
		objTempRs=UpdateRecordSPWithConn(objConnCustom, "usp_ExpertEducationUpdate", Array( _
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
		objTempRs=InsertRecordSPWithConn(objConnCustom, "usp_ExpertEducationInsert", Array( _
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
<!--#include virtual="/_template/html.header.start.asp"-->
<script type="text/javascript" src="/_scripts/js/main.js"></script>
<script language="JavaScript">
<!--
function validateForm() { 
	var f=document.RegForm;
<% If Len(sBackOffice)<3 Then %>
	if (f.exp_inst_name.value!="" || f.exp_edu_subj.selectedIndex>0 || f.exp_edu_syear.selectedIndex>0) { 
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
<% If sApplicationName="external" Or sApplicationName="backoffice" Then %>
	f.submit();
	return;
<% End If %>
	
<% If Len(sBackOffice)<3 Then %>
	if (f.exp_inst_name.value=="") {
		alert("<% =GetLabel(sCvLanguage, "Please fill in the institution name") %>"); document.RegForm.exp_inst_name.focus(); return;
	}
	var start_month=parseInt(f.exp_edu_smonth.options[f.exp_edu_smonth.selectedIndex].value);
	var start_year=parseInt(f.exp_edu_syear.options[f.exp_edu_syear.selectedIndex].value);
	var end_month=parseInt(f.exp_edu_emonth.options[f.exp_edu_emonth.selectedIndex].value);
	var end_year=parseInt(f.exp_edu_eyear.options[f.exp_edu_eyear.selectedIndex].value);
	if (end_month==0 || end_year==0) {
		alert("<% =GetLabel(sCvLanguage, "Please fill in the education end date") %>"); return;
	}
	if ((start_year>end_year) || (start_year>0 && start_month>0 && start_year==end_year && start_month>end_month)) {
		alert("<% =GetLabel(sCvLanguage, "Please fill in the education dates properly") %>"); return;
	}
	if (f.exp_edu_diploma.selectedIndex==0 && f.exp_edu_diploma1.value=="") {
		alert("<% =GetLabel(sCvLanguage, "Please specify a type of diploma or degree obtained") %>"); return;
	}
	if (f.exp_edu_subj.selectedIndex==0) {
		alert("<% =GetLabel(sCvLanguage, "Please specify the education subject") %>"); return;
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
<body>
	<!-- header -->
	<!--#include virtual="/_template/page.header.asp"-->

	<!-- content -->
	<div id="content" class="searchform">
	<% 
	If Not bIsMyCV Then 
		%><div id="hdrUpdatedList" class="colCCCCCC uprCse f17 spc01 botMrgn10"><span class="service_title">Curriculum Vitae.</span> Expert ID: <% =objExpertDB.DatabaseCode %><%=iExpertID%></div>
		<% 
	Else
		%><div class="colCCCCCC uprCse f17 spc01 botMrgn10"><span class="service_title">Curriculum Vitae</span></div>
		<% 
	End If
	
	ShowRegistrationProgressBar "CV", 2 
	%>
		
	<form method="post" action="<%=sScriptFileName & sParams%>" name="RegForm">
	<input type="hidden" name="exp_edu_continue" value="1">
		<div class="box search blue">
		<h3><% =GetLabel(sCvLanguage, "Education") %></h3>
		<table class="search_form" width="100%" cellspacing=0 cellpadding=0 border=0>
	<%
	Set objTempRs=GetDataRecordsetSPWithConn(objConnCustom, "usp_ExpertEducationSelect", Array( _
		Array(, adInteger, , iExpertID), _
		Array(, adInteger, , 1)))
	If Not objTempRs.Eof Then %>
	<tr>
	<td colspan=2>
		<table class="results" style="border-left:0; border-right:0;">
		<tr class="tr_results">
		<th class="number"><p>N.</p></td>
		<th width=215><p><% =GetLabel(sCvLanguage, "Institution") %></p></td>
		<th class="date"><p><% =GetLabel(sCvLanguage, "Start date") %></p></td>
		<th class="date"><p><% =GetLabel(sCvLanguage, "End date") %></p></td>
		<th width=200><p><% =GetLabel(sCvLanguage, "Subject") %></p></td>
		<th width=15><p><% =GetLabel(sCvLanguage, "Modify") %></p></td>
		<th width=15><p><% =GetLabel(sCvLanguage, "Delete") %></p></td>
		</tr>	
		<% i=1
		While Not objTempRs.Eof %>
		<tr class="tr_results<% If i Mod 2 = 0 Then %> odd<% End If %>">
		<td><p align="center"><%=i%>.</td>
		<td><p><a href="register2.asp<%=AddUrlParams(sParams, "eduid=" & objTempRs("id_ExpEdu"))%>&act=update"><% If objTempRs("id_ExpEdu")=iExpEduID Then %><b><img src="<% =sHomePath %>image/vn_v.gif" width=8 height=12 border=0 hspace=0 align="left"><% End If %><%=CheckSpaces(ReplaceIfEmpty(objTempRs("InstNameEng"),"Not specified"), 30) %></a></td>
		<td><p><%=ConvertDateForText(objTempRs("eduStartDate"), "&nbsp;", "MMYYYY") %></td>
		<td><p><%=ConvertDateForText(objTempRs("eduEndDate"), "&nbsp;", "MMYYYY") %></td>
		<td><p><%=objTempRs("edsDescriptionEng") %></td>
		<td align="center"><% If objTempRs("id_ExpEdu")=iExpEduID Then %><img src="<% =sHomePath %>image/vn_updte.gif" width=15 height=15 border=0 hspace=0 alt="Updating" align="center"><% Else %><a href="register2.asp<%=AddUrlParams(sParams, "eduid=" & objTempRs("id_ExpEdu"))%>&act=update"><img src="<% =sHomePath %>image/vn_updt.gif" width=15 height=15 border=0 hspace=0 alt="Update this record" align="center"></a><% End If %></td>
		<td align="center"><a href="register2.asp<%=AddUrlParams(sParams, "eduid=" & objTempRs("id_ExpEdu"))%>&act=delete"><img src="<% =sHomePath %>image/vn_del.gif" width=15 height=15 border=0 hspace=0 alt="Delete this record" align="center"></a></td>
		</tr>
		<% i=i+1
		objTempRs.MoveNext
		WEnd %>
		</table>
	</td>
	</tr>
	<% End If 
	objTempRs.Close
	Response.Flush()
	
	If iExpEduID>0 Then
	Set objTempRs=GetDataRecordsetSPWithConn(objConnCustom, "usp_ExpertEducationSelect", Array( _
		Array(, adInteger, , iExpertID), _
		Array(, adInteger, , 1), _
		Array(, adInteger, , iExpEduID)))

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
		<input type="hidden" name="id_Edu" value="<%=iExpEduID%>">
		<tr>
		<td class="field splitter"><label for="exp_inst_name"><% =GetLabel(sCvLanguage, "Institution name") %></label></td>
		<td class="value blue"><input type="text" id="exp_inst_name" name="exp_inst_name" size="31" style="width:355px;" value="<%=sExpEduInstitution%>" maxlength="255">&nbsp;&nbsp;<span class="fcmp">*</span></td>
		</tr>
		<tr>
		<td class="field splitter"><label for="exp_edu_smonth"><% =GetLabel(sCvLanguage, "Start date") %></label></td>
		<td class="value blue"><select id="exp_edu_smonth" name="exp_edu_smonth" size=1>
		<option value="0" selected><% =GetLabel(sCvLanguage, "Month") %></option>
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
		<option value="0" Selected><% =GetLabel(sCvLanguage, "Year") %></option>
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
		</select></td>
		</tr>
		<tr>
		<td class="field splitter"><label for="exp_edu_emonth"><% =GetLabel(sCvLanguage, "End date") %></label></td>
		<td class="value blue"><select id="exp_edu_emonth" name="exp_edu_emonth" size=1>
		<option value="0" selected><% =GetLabel(sCvLanguage, "Month") %></option>
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
		<option value="0" Selected><% =GetLabel(sCvLanguage, "Year") %></option>
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
		</select>&nbsp;&nbsp;<span class="fcmp">*</span></td>
		</tr>
		<tr>
		<td class="field splitter"><label for="exp_edu_diploma"><% =GetLabel(sCvLanguage, "Type of diploma") %></label></td>
		<td class="value blue"><select size="1" id="exp_edu_diploma" name="exp_edu_diploma" style="width:355px;">
		<option value="0" selected> <% =GetLabel(sCvLanguage, "Please select") %> </option>
		<%  For i=LBound(arrEduTypeID) to UBound(arrEduTypeID)

		If iExpEduDiplomaID=arrEduTypeID(i) Then
			Response.Write("<option value=" & arrEduTypeID(i) & " selected>" & arrEduTypeTitle(i) & "</option>")
		Else 
			Response.Write("<option value=" & arrEduTypeID(i) & ">" & arrEduTypeTitle(i) & "</option>")
		End If
		Next %>
		</select>&nbsp;&nbsp;<span class="fcmp">*</span></td>
		</tr>
		<tr>
		<td class="field splitter"><label for="exp_edu_diploma1"><% =GetLabel(sCvLanguage, "If other please specify") %></label></td>
		<td class="value blue"><input type="text" id="exp_edu_diploma1" name="exp_edu_diploma1" value="<%=sExpEduDiplomaText%>" size="31" style="width:355px;" maxlength="255"></td>
		</tr>
		<tr>
		<td class="field splitter"><label for="exp_edu_subj"><% =GetLabel(sCvLanguage, "Subject") %></label></td>
		<td class="value blue"><select id="exp_edu_subj" name="exp_edu_subj" size="1" style="width:355px;">
		<option value="0" selected> <% =GetLabel(sCvLanguage, "Please select") %> </option>
		<%  For i=1 to UBound(arrEduSubjectID) 
		If iExpEduSubjectID=arrEduSubjectID(i) then
			Response.Write("<option value=" & arrEduSubjectID(i) & " selected>" & arrEduSubjectTitle(i) & "</option>")
		Else 
			Response.Write("<option value=" & arrEduSubjectID(i) & ">" & arrEduSubjectTitle(i) & "</option>")
		End If
		Next %>
		</select>&nbsp;&nbsp;<span class="fcmp">*</span>&nbsp;</td>
		</tr>
		<tr>
		<td class="field splitter"><label for="exp_edu_subj1"><% =GetLabel(sCvLanguage, "If needed, please specify the exact title of your degree") %></label></td>
		<td class="value blue"><input type="text" id="exp_edu_subj1" name="exp_edu_subj1" value="<%=sExpEduSubjectText%>" size="31" style="width:355px;" maxlength="255"></td>
		</tr>
		</table>
		</div>

		<div class="spacebottom">
		<a href="javascript:AddEducation(0);" class="red-button w125 under-right-col" title="Add this education to the list of your qualifications"><% If iExpEduID>0 Then %>save<% Else %>add<% End If %> qualification</a>
		<a href="javascript:validateForm();" class="red-button w125 next-btn">save & continue</a>
		</div>
		</form>

	</div>
	<!-- footer -->
	<!--#include file="../_template/page.footer.asp"-->
<% CloseDBConnection %>
</body>
<!--#include file="../_template/html.footer.asp"-->

