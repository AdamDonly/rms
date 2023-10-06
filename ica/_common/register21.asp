<% 
'--------------------------------------------------------------------
'
' CV registration.
' Full format. Training.
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
	iLogResult = LogActivity(34, "ExpertID=" & Cstr(iExpertID) & " SavedStep: 3", "", "")
End If

Dim objConnCustom
Set objConnCustom = Server.CreateObject("ADODB.Connection")
objConnCustom.Open "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Data Source=MAGNESA;Initial Catalog=" & objExpertDB.DatabasePath & ";"

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
	objTempRs=UpdateRecordSPWithConn(objConnCustom, "usp_ExpertEducationDelete", Array( _
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
		objTempRs=UpdateRecordSPWithConn(objConnCustom, "usp_ExpCvvTrainingUpdate", Array( _
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
		objTempRs=InsertRecordSPWithConn(objConnCustom, "usp_ExpCvvTrainingInsert", Array( _
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
<!--#include virtual="/_template/html.header.start.asp"-->
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
<% If sApplicationName="external" Or sApplicationName="backoffice" Then %>
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

	ShowRegistrationProgressBar "CV", 3
	%>
		
	<form method="post" action="<%=sScriptFileName & sParams%>" name="RegForm">
	<input type="hidden" name="exp_otr_continue" value="1">
		<div class="box search blue">
		<h3><% =GetLabel(sCvLanguage, "Training") %></h3>
		<table class="search_form" width="100%" cellspacing=0 cellpadding=0 border=0>
  
	<% ' List of training
	Set objTempRs=GetDataRecordsetSPWithConn(objConnCustom, "usp_ExpertEducationSelect", Array( _
		Array(, adInteger, , iExpertID), _
		Array(, adInteger, , 2)))
	If Not objTempRs.Eof Then %>
	<tr>
	<td colspan=2>
		<table class="results" style="border-left:0; border-right:0;">
		<tr class="tr_results">
		<th class="number"><p>N.</p></td>
		<th width=410><p><% =GetLabel(sCvLanguage, "Title") %></p></td>
		<th class="date"><p><% =GetLabel(sCvLanguage, "Start date") %></p></td>
		<th class="date"><p><% =GetLabel(sCvLanguage, "End date") %></p></td>
		<th width=15><p><% =GetLabel(sCvLanguage, "Modify") %></p></td>
		<th width=15><p><% =GetLabel(sCvLanguage, "Delete") %></p></td>
		</tr>	
		<% i=1
		while not objTempRs.EOF 
		%>
		<tr class="tr_results<% If i Mod 2 = 0 Then %> odd<% End If %>">
		<td><p align="center"><%=i%>.</td>
		<td><p><a href="register21.asp<%=AddUrlParams(sParams, "eduid=" & objTempRs("id_ExpEdu"))%>&act=update"><% If objTempRs("id_ExpEdu")=iExpEduID Then %><b><img src="<% =sHomePath %>image/vn_v.gif" width=8 height=12 border=0 hspace=0 align="left"><% End If %><%=CheckSpaces(ReplaceIfEmpty(objTempRs("eduDiploma1Eng"), "Not specified"), 45) %></a></td>
		<td><p><%=ConvertDateForText(objTempRs("eduStartDate"), "&nbsp;", "MMYYYY") %></td>
		<td><p><%=ConvertDateForText(objTempRs("eduEndDate"), "&nbsp;", "MMYYYY") %></td>
		<td align="center"><% If objTempRs("id_ExpEdu")=iExpEduID Then %><img src="<% =sHomePath %>image/vn_updte.gif" width=15 height=15 border=0 hspace=0 alt="Updating" align="center"><% Else %><a href="register21.asp<%=AddUrlParams(sParams, "eduid=" & objTempRs("id_ExpEdu"))%>&act=update"><img src="<% =sHomePath %>image/vn_updt.gif" width=15 height=15 border=0 hspace=0 alt="Update this record" align="center"></a><% End If %></td>
		<td align="center"><a href="register21.asp<%=AddUrlParams(sParams, "eduid=" & objTempRs("id_ExpEdu"))%>&act=delete"><img src="<% =sHomePath %>image/vn_del.gif" width=15 height=15 border=0 hspace=0 alt="Delete this record" align="center"></a></td>
		</tr>
		<% i=i+1
		objTempRs.MoveNext
		WEnd 
		%>
		</table>
	</td>
	</tr>
	<% End If 
	objTempRs.Close%>
	
	<% If iExpEduID>0 Then
	Set objTempRs=GetDataRecordsetSPWithConn(objConnCustom, "usp_ExpertEducationSelect", Array( _
		Array(, adInteger, , iExpertID), _
		Array(, adInteger, , 2), _
		Array(, adInteger, , iExpEduID)))
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
		<input type="hidden" name="id_Edu1" value="<%=iExpEduID%>">
		<tr>
		<td class="field splitter"><label for="exp_otr_name"><% =GetLabel(sCvLanguage, "Title") %></label></td>
		<td class="value blue"><input type="text" id="exp_otr_name" name="exp_otr_name" size="31" style="width:355px;" value="<%=sExpEduTitle%>" maxlength="255">&nbsp;&nbsp;<span class="fcmp">*</span></td>
		</tr>
		<tr>
		<td class="field splitter"><label for="exp_otr_other"><% =GetLabel(sCvLanguage, "Skills / Qualifications") %></label></td>
		<td class="value blue"><input type="text" id="exp_otr_other" name="exp_otr_other" size="31" style="width:355px;" value="<%=sExpEduOther%>" maxlength="255"></td>
		</tr>
		<tr>
		<td class="field splitter"><label for="exp_otr_smonth"><% =GetLabel(sCvLanguage, "Start date") %></label></td>
		<td class="value blue"><select id="exp_otr_smonth" name="exp_otr_smonth" size=1>
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
		</select></td>
		</tr>
		<tr>
		<td class="field splitter"><label for="exp_otr_emonth"><% =GetLabel(sCvLanguage, "End date") %></label></td>
		<td class="value blue"><select id="exp_otr_emonth" name="exp_otr_emonth" size=1>
		<option value="0"><% =GetLabel(sCvLanguage, "Month") %></option>
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
		<option value="0"><% =GetLabel(sCvLanguage, "Year") %></option>
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
		</select>&nbsp;&nbsp;<span class="fcmp">*</span></td>
		</tr>
		<tr class="last">
		<td class="field splitter"><label for="exp_otr_desc"><% =GetLabel(sCvLanguage, "Achievements") %></label></td>
		<td class="value blue"><input type="text" id="exp_otr_desc" name="exp_otr_desc" size="31" style="width:355px;" value="<%=sExpEduAchievements%>" maxlength="255"></td>
		</tr>
		</table>
		</div>
		
		<div class="spacebottom">
		<a href="javascript:AddTraining(0);" class="red-button w125 under-right-col" title="Add this training to the list of your training"><% If iExpEduID>0 Then %>save<% Else %>add<% End If %> training</a>
		<a href="javascript:validateForm();" class="red-button w125 next-btn">Save & continue</a>
		</div>
		</form>

	</div>
	<!-- footer -->
	<!--#include file="../_template/page.footer.asp"-->
<% CloseDBConnection %>
</body>
<!--#include file="../_template/html.footer.asp"-->
