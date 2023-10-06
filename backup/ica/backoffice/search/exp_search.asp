<%@ LANGUAGE="VBSCRIPT" %>
<% 'Option Explicit
'--------------------------------------------------------------------
'
' Search for experts
'
'--------------------------------------------------------------------
%>
<!--#include virtual="/dbc.asp"-->
<!--#include virtual="/fnc.asp"-->
<!--#include virtual="/fnc_exp.asp"-->
<!--#include virtual="/_common/_data/en/lib.asp"-->
<!--#include virtual="/_common/_data/datCountry.asp"-->
<!--#include virtual="/_common/_data/datLngName.asp"-->
<!--#include virtual="/_common/_data/datLngLevel.asp"-->
<!--#include virtual="/_common/_data/datEduSubject.asp"-->
<!--#include virtual="/_forms/frmInterface.asp"-->
<!--#include virtual="/_forms/frmScrollBoxNew.asp"-->
<%
' Check user's access rights
CheckUserLogin sScriptFullNameAsParams

Dim ExpertIds
Dim iNumberCVsSelected, iNumberCVsDownloaded, iNumberCVsDownloadedFromSelected, iNumberCVsSubscribedFor, iNumberCVsInOptimalPackage, iNumberCVsInNextPackage

iProjectID=CheckIntegerAndZero(Request.QueryString("idproject"))
sParams=ReplaceUrlParams(sParams, "idproject=" & iProjectID)

Response.Cookies("ExpertIds")=""
ExpertIds="0"
%>
<!--#include virtual="/_template/html.header.scrolllist.start2.asp"-->
	
<% InsertJsHelpers 0, 0 %>

<script language="JavaScript" type="text/javascript">
document.cookie = 'ExpertIds=';
function Continue() {
	var f = document.RegForm;

	f.mmb_cou_hid.value = '0,' + (GetScrollboxSelection("divCouSelector", "cou_") || '0');
	f.mmb_don_hid.value = '0,' + (GetScrollboxSelection("divDonSelector", "don_") || '0');
	f.mmb_sct_hid.value = '0,' + (GetScrollboxSelection("divSctSelector", "sct_") || '0');

	// at least one criteria should be filled in
	if ((f.mmb_cou_hid.value.replace(/,0/gi,'')=='0') 
		&& (f.mmb_don_hid.value.replace(/,0/gi,'')=='0')
		&& (f.mmb_sct_hid.value.replace(/,0/gi,'')=='0')
		&& (f.srch_firstname.value.length<2)
		&& (f.srch_surname.value.length<2)
		&& (f.srch_query.value.length<2)
		&& (f.currentlyin.selectedIndex<1)
		&& (f.nationality.selectedIndex<1)
		&& (f.subject.selectedIndex<1)
		&& (f.nativelng.selectedIndex<1)
		&& (f.seniority.selectedIndex<1)
	<% If bUserAccessMethodology = True Then %>
		&& (!f.methodology.checked)
	<% End If %>
		) {	
			alert('Please fill in search criteria.');
			return false;
		}
	
	f.action="exp_results.asp";
	f.submit();
}
</script>
<script language="JavaScript" src="../../../_scripts/js/asr.js"></script>
<script language="JavaScript" src="../../../_scripts/js/lib.20170726.js"></script>
<% InsertScrollStyles %>
<link rel="stylesheet" type="text/css" media="all" href="/scroll.css" />
</head>
<body onload="RestoreInt();" onunload="">

	<!-- header -->
	<!--#include file="../../_template/page.header.asp"-->

	<!-- content -->
	<div id="content" class="searchform">
	
<div id="hdrUpdatedList" class="colCCCCCC uprCse f17 spc01 botMrgn10">SEARCH FOR EXPERTS</div>
<div class="col666666 itlc botMrgn10">
	<% 
	' display message about Assortis CVs in the result, if the user has EDB access:
	Dim assortisEdbTextAddition
	assortisEdbTextAddition = ""
	If bAssortisSubscriptionEdbActive = True Then
		assortisEdbTextAddition = "<br/><br/>You can also access the CVs registered in the Assortis Database of Experts."
	End if

	' No access to other expert databases OR 10% access:
	'OLD CHECK:If iMemberAccessExperts = cMemberAccessExpertsOwnOnly Then 
	If iUserAccessExpertsOtherDbNoAccess = 1 OR iUserAccessExpertsOtherDbRestricted = 1 Then
		Dim memberExpertsCount
		memberExpertsCount = GetExpCount(objUserCompanyDB.Database)
		
		If isNull(memberExpertsCount) Then memberExpertsCount = 0 End If
		%><b>Search among <% =ShowEntityPlural(memberExpertsCount, "expert", "experts", " ") %><% If memberExpertsCount = 0 Then %>s<% End If %> in <% =objUserCompanyDB.DatabaseTitle %>’s Expert Database<%
			If iUserAccessExpertsOtherDbRestricted = 1 Then
				%> and <span style="color:#d22c29">random 10% of the ICA Common Database of Experts</span><%
			End If
			%>.</b><%=assortisEdbTextAddition %><br/><br/>
		If you wish to get full access and search through all <% =GetExpCount("all") %> experts in the ICA Common Database of Experts<br>please contact your <a href="mailto:info@icaworld.net">ICA Team</a>.
		<% 
	' Full access:
	ElseIf iUserAccessExpertsTab = 1 And iUserAccessExpertsOtherDbNoAccess = 0 And iUserAccessExpertsOtherDbRestricted = 0 Then
		%>Search among <% =GetExpCount("all") %> experts in ICA Members’ Databases<%=assortisEdbTextAddition %>
		<% 
	End If %>
</div>

<!-- Keywords search -->
	<form action="exp_results.asp" method="post" name="RegForm" id="RegForm" onSubmit="Continue(); return false;">
		<div class="box search blue">
		<h3>Keywords search</h3>
		<table class="search_form" width="100%" cellspacing="0" cellpadding="0" border="0">

			<tr><td class="field splitter" style="width: 25%">Expert name</td>
			<td class="value blue" style="width: 75%">
				<table width="400" cellpadding="0" cellspacing="0" border="0">
				<tr valign="top">
				<td width="153"><input type="text" name="srch_firstname" size=13 maxlength=100 style="width:160px;margin-right:5px;" /></td>
				<td width="200" colspan="2"><input type="text" name="srch_surname" size=13 maxlength=7500 style="width:220px;" /></td>
				</tr>
				<tr>
				<td><p class="sml">&nbsp;First name</p></td>
				<td><p class="sml">&nbsp;Family&nbsp;name&nbsp;</p></td>
				<td align="right"><p class="sml" align="right">ID&nbsp;&nbsp;</p></td>
				</tr>
				</table>
			</td>
			</tr>
			
			<tr class="first last"><td class="field splitter"><label for="srch_query">Keyword search</label></td>
			<td class="value blue"><select id="srch_querytype" name="srch_querytype" style="width:160px;">
				<option selected value="all of the words from">all of the words</option>
				<option value="any of the words from">any of the words</option>
				<option value="the exact phrase">the exact phrase</option>
				<option value="boolean expression">boolean expression</option></select>&nbsp;&nbsp;
				<input type="text" name="srch_query" size=23 maxlength=100 style="width:220px;">
				<p class="sml">&nbsp;This searches the entire content of all the online CVs.</p>
			</td>
			</tr>
		</table>
		</div>

		<div class="spacebottom">
			<input type="submit" class="red-button" style="margin-left:25%;margin-right:15px;" value="Search">
			<a href="<%=sScriptFullName & AddUrlParams(sParams, "act=" & sAccessType) %>" class="red-button">Clear all</a>
		</div>

		<div class="mini_wraper">
			<%
			' Sectors
			ShowSctScrollBox "Sectors of experts' experience",  "", 1, 0, 0 %><br />
			<%
			' Countries
			ShowCouScrollBox "Countries of experts' experience",  "", 1, 0, 1, 1, 0 %><br />
			<%			
			'Funding agancies
			ShowDonScrollBox "Funding agencies of experts' experience",  "", 1, 1, 0, 0 %><br />

  <!-- Search options -->
		<div class="box search blue">
		<h3>Search options</h3>
		<table class="search_form" width="100%" cellspacing=0 cellpadding=0 border=0>
		<% If iMemberAccessExperts = cMemberAccessExpertsOwnOnly Then %>
			<input type="hidden" name="database" value="<% =objUserCompanyDB.Database %><% If bAssortisSubscriptionEdbActive Then %>,assortis<% End If %>">
		<% Else %>
			<tr><td class="field splitter" style="width: 25%"><label for="database">Database(s)</label></td>
			<td class="value blue" style="width: 75%"><select id="database" name="database" size=6 multiple style="width:355px;">
			<option value="">- all -</option>
			<% objExpertDBList.ShowSelectItems "", "Database", "" %>
			</select>
			<p class="sml">(press [Ctrl] for multiple selection)</p>
			</td></tr>
		<% End If %>
		
 
		<tr><td class="field splitter" style="width: 25%"><label for="pastyears">Timeframe of past<br />relevant experience</label></td>
		<td class="value blue" style="width: 75%"><select id="pastyears" name="pastyears" size=1 style="width: 355px;">
		<option value="0" selected> </option>
		<%
			Response.Write("<option value=""1"">12 months</option>" & vbCrLf)
			Response.Write("<option value=""2"">2 years</option>" & vbCrLf)
			Response.Write("<option value=""4"">4 years</option>" & vbCrLf)
			Response.Write("<option value=""100"">&gt;4 years (entire career)</option>" & vbCrLf)
		%>
		</select>
		</td></tr>

		<tr><td class="field splitter"><label for="pastprojects">Consider experience on<br />at least (n° of projects)</label></td>
		<td class="value blue"><select id="pastprojects" name="pastprojects" size=1 style="width: 355px;">
		<option value="0" selected> </option>
		<option value="5">2 projects</option>
		<option value="5">5 projects</option>
		<option value="10">10 projects</option>
		</select>
		</td></tr>

		<tr><td class="field splitter"><label for="currentlyin">Currently working in</label></td>
		<td class="value blue"><select id="currentlyin" name="currentlyin" size=1 style="width: 355px;">
		<option value="0"> </option>
		<%
		For i=1 To Ubound(arrCountryID)
			Response.Write("<option value=" & arrCountryID(i) & ">" & arrCountryName(i) & "</option>" & vbCrLf)
		Next
		%>
		</select>
		</td></tr>


		<tr><td class="field splitter"><label for="nationality">Nationality</label></td>
		<td class="value blue"><select id="nationality" name="nationality" size=6 multiple style="width: 355px;">
		<option value="">- all -</option>
		<option value="1100">EU Member States</option>
		<option value="1122">ACP Countries</option>
		<option value="1123">ENPI Countries</option>
		<option value="1124">IPA Countries</option>
		<!--
		<option value="1114">MEDA Countries</option>
		<option value="1112">TACIS Countries</option>
		-->
		<%
		For i = LBound(arrCountryID) To UBound(arrCountryID) - 1
			If arrCountryEU(i)=1 Then Response.Write("<option value=" & arrCountryID(i) & ">" & arrCountryName(i) & "</option>" & vbCrLf)
		Next
		For i = LBound(arrCountryID) To UBound(arrCountryID) - 1
			If arrCountryEU(i)=0 Then Response.Write("<option value=" & arrCountryID(i) & ">" & arrCountryName(i) & "</option>" & vbCrLf)
		Next
		%>
		</select>
		<p class="sml">(press [Ctrl] for multiple selection)</p>
		</td></tr>

		<tr><td class="field splitter"><label for="subject">Education</label></td>
		<td class="value blue"><select id="subject" name="subject" multiple size="4" style="width: 355px;">
		<option value="">- all -</option>
		<%
		For i = 1 To UBound(arrEduSubjectID) - 1
			Response.Write("<option value=" & arrEduSubjectID(i) & ">" & arrEduSubjectTitle(i) & "</option>" & vbCrLf)
		Next
		%>
		</select>
		<p class="sml">(press [Ctrl] for multiple selection)</p>
		</td></tr>

		<tr><td class="field splitter"><label for="nativelng">Languages</label></td>
		<td class="value blue">
		<table><tr valign="top">
			<td>
			<select id="nativelng" name="nativelng" multiple size="4" style="width: 220px;">
			<option value="">- all -</option>
			<%
			For i = 1 To UBound(arrLanguageID)
				Response.Write("<option value=" & arrLanguageID(i) & ">" & arrLanguageTitle(i) & "</option>" & vbCrLf)
			Next
			%>
			</select></td>
			<td>&nbsp;</td>
			<td><p class="sml">Level of knowledge</p>
			<select id="language_level" name="language_level" size="1" style="width: 120px;">
			<option></option>
			<%
			For i = 1 To Ubound(arrLanguageLevelID)
				Response.Write("<option value=" & arrLanguageLevelID(i)*3 & ">" & arrLanguageLevelTitle(i) & "</option>" & vbCrLf)
			Next
			%>
			</select></td>
		</tr></table>
		<p class="sml">(press [Ctrl] for multiple selection)</p>
		</td></tr>

		<tr class="last"><td class="field splitter"><label for="seniority">Seniority</label></td>
		<td class="value blue"><select id="seniority" name="seniority" style="width: 220px;">
		<option value=""> </option>
		<option value="0 AND 5">less than 5 years</option>
		<option value="5 AND 100">over 5 years</option>
		<option value="5 AND 10">between 5 and 10 years</option>
		<option value="0 AND 10">less than 10 years</option>
		<option value="10 AND 100">over 10 years</option>
		<option value="10 AND 15">between 10 and 15 years</option>
		<option value="0 AND 15">less than 15 years</option>
		<option value="15 AND 100">over 15 years</option>
		<option value="0 AND 20">less than 20 years</option>
		<option value="20 AND 100">over 20 years</option>
		</select>
		</td></tr>

	<% If bCvMultiLanguageActive = cMultiLanguageEnabled Then %>
		<tr class="last"><td class="field splitter"><label for="cv_language">CV language</label></td>
		<td class="value blue"><select name="cv_language" id="cv_language" style="width:130px;">
		<option value=""></option>
		<%
		Dim sTempLanguage
		For Each sTempLanguage in dictLanguage
			Response.Write "<option value=""" & sTempLanguage & """" 
			Response.Write ">" & dictLanguage.Item(sTempLanguage) & "</option>"
		Next
		%>
		</select>
		</td></tr>
	<% End If %>

	<% If bCvTypeActive = cCvTypeEnabled Then %>
		<tr class="last"><td class="field splitter"><label for="cv_type">CV type</label></td>
		<td class="value blue"><select name="cv_type" id="cv_type" style="width:130px;">
		<option value=""></option>
		<%
		Dim sTempCvType
		For Each sTempCvType in dictCvType
			Response.Write "<option value=""" & sTempCvType & """" 
			Response.Write ">" & dictCvType.Item(sTempCvType) & "</option>"
		Next
		%>
		</select>
		</td></tr>
	<% End If %>

	<% If bUserAccessMethodology = True Then %>
		<tr class="last"><td class="field splitter"><label for="methodology">Methodology writer</label></td>
		<td class="value blue"><input type="checkbox" name="methodology" id="methodology">
		</td></tr>

		<tr class="last"><td class="field splitter"><label for="metho_flag">Methodology writer status</label></td>
		<td class="value blue"><select name="metho_flag" id="metho_flag" style="width:130px;">
		<option value=""></option>
		<option value="10">Blue</option>
		<option value="20">Green</option>
		<option value="80">Yellow</option>
		<option value="90">Red</option>
		<option value="50">Grey</option>
		</select>
		</td></tr>

	<% End If %>

		</table>
		</div>

		<input type="submit" class="red-button" style="margin-left:25%;margin-right:15px;" value="Search">
		<a href="<%=sScriptFullName & AddUrlParams(sParams, "act=" & sAccessType) %>" class="red-button">Clear all</a>
		<br/><br/>
		
	<input type="hidden" name="mmb_cou_hid">
	<input type="hidden" name="mmb_don_hid">
	<input type="hidden" name="mmb_sct_hid">
	<input type="hidden" name="srch_type" value="advanced">
	</form>
	</div>

	<!-- footer -->
	<!--#include virtual="/_template/page.footer.asp"-->
<% CloseDBConnection %>

<% WriteFilterTableScript %>
<link href="/res/css/jquery.mCustomScrollbar.css" rel="stylesheet" type="text/css" />
<script src="/res/scripts/jquery.mCustomScrollbar.concat.min.js"></script>
</body>
<!--#include virtual="/_template/html.footer.asp"-->
