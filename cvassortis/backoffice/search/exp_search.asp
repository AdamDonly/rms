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
<!--#include file="../../../_common/_data/en/lib.asp"-->
<!--#include virtual="/_forms/frmInterface.asp"-->
<!--#include virtual="/_forms/frmScrollBox.asp"-->
<!--#include file="../../../_common/_data/datCountry.asp"-->
<!--#include file="../../../_common/_data/datLngName.asp"-->
<!--#include file="../../../_common/_data/datLngLevel.asp"-->
<!--#include file="../../../_common/_data/datEduSubject.asp"-->
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
<!--#include virtual="/_template/html.header.scrolllist.start.asp"-->
	
<% InsertJSScrollFunctions 0, 0 %>
<script language="JavaScript">
<!--
document.cookie = 'ExpertIds=';
function Continue() 
{  
	var f = document.forms[0];
	f.mmb_cou_hid.value='0'+jNtInt;
	f.mmb_don_hid.value='0'+jOrgInt;
	f.mmb_sct_hid.value='0'+jExTInt;

	// at least one criteria should be filled in
	if ((f.mmb_cou_hid.value.replace(/,0/gi,'')=='0') 
		&& (f.mmb_don_hid.value.replace(/,0/gi,'')=='0')
		&& (f.mmb_sct_hid.value.replace(/,0/gi,'')=='0')
		&& (f.srch_firstname.value.length<2)
		&& (f.srch_surname.value.length<2)
		&& (f.srch_query.value.length<2)
		&& (f.currentlyin.selectedIndex<1)
		&& (f.nationality.selectedIndex<0)
		&& (f.subject.selectedIndex<0)
		&& (f.nativelng.selectedIndex<0)
		&& (f.seniority.selectedIndex<1)
		) {	
			alert('Please fill in search criteria.');
			return false;
		}
	
	f.action="exp_results.asp";
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
	<!--#include file="../../_template/page.header.asp"-->

	<!-- content -->
	<div id="content" class="searchform">
		<h2 class="service_title">Search for experts</h2>

<!-- Keywords search -->
	<form action="exp_results.asp" method="post" name="RegForm" id="RegForm" onSubmit="Continue(); return false;">
		<div class="box search blue">
		<h3><span class="left">&nbsp;</span><span class="right">&nbsp;</span>Keywords search</h3>
		<table class="search_form" width="100%" cellspacing="0" cellpadding="0" border="0">

			<tr><td class="field splitter">Expert name</td>
			<td class="value blue">
				<table width="360" cellpadding="0" cellspacing="0" border="0">
				<tr valign="top">
				<td width="153"><input type="text" name="srch_firstname" size=13 maxlength=100 style="width:145px;"></td>
				<td width="200" colspan="2"><input type="text" name="srch_surname" size=13 maxlength=7500 style="width:200px;"></td>
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
			<td class="value blue"><select id="srch_querytype" name="srch_querytype" style="width:145px;">
				<option selected value="all of the words from">all of the words</option><option value="any of the words from">any of the words</option><option value="the exact phrase">the exact phrase</option><option value="boolean expression">boolean expression</option></select>&nbsp;&nbsp;<input type="text" name="srch_query" size=23 maxlength=100 style="width:200px;">
				<p class="sml">&nbsp;This searches the entire content of all the online CVs.</p>
			</td>
			</tr>
		</table>
		</div>

		<div class="spacebottom">
		<input class="button first" type="image" src="<% =sHomePath %>image/bte_search.gif" name="Search" alt="Search">
		<a href="<%=sScriptFullName & AddUrlParams(sParams, "act=" & sAccessType) %>"><img class="button" src="<% =sHomePath %>image/bte_clearall.gif" alt="Clear all"></a>
		</div>

  <!-- Sectors section -->
	<% ShowSctScrollBox "Sectors of experts experience",  "", 1, 0, 1, 1, 0 %><br />

  <!-- Countries section -->
	<% ShowCouScrollBox "Countries of experts experience",  "", 1, 0, 1, 1, 0 %><br />
	
  <!-- Funding agancies section -->
	<% ShowDonScrollBox "Funding agencies of experts experience",  "", 1, 0, 1, 1, 0, 0 %><br />

  <!-- Search options -->
		<div class="box search blue">
		<h3><span class="left">&nbsp;</span><span class="right">&nbsp;</span>Search options</h3>
		<table class="search_form" width="100%" cellspacing=0 cellpadding=0 border=0>

		<% If iUserCompanyMemberType=cMemberTypeTechnicalAssociate Then %>
			<input type="hidden" name="database" value="<% =objExpertDBList.Find(iUserCompanyID, "CompanyID").Database %>">
		<% Else %>
			<tr><td class="field splitter"><label for="database">Database(s)</label></td>
			<td class="value blue"><select id="database" name="database" size=6 multiple style="width:355px;">
			<% objExpertDBList.ShowSelectItems "", "Database", "" %>
			</select>
			<p class="sml">(press [Ctrl] for multiple selection)</p>
			</td></tr>
		<% End If %>
		
 
		<tr><td class="field splitter"><label for="pastyears">Timeframe of past<br />relevant experience</label></td>
		<td class="value blue"><select id="pastyears" name="pastyears" size=1 style="width: 355px;">
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
		<option value="1100">EU Member States</option>
		<option value="1122">ACP Countries</option>
		<option value="1123">ENPI Countries</option>
		<option value="1124">IPA Countries</option>
		<!--
		<option value="1114">MEDA Countries</option>
		<option value="1112">TACIS Countries</option>
		-->
		<%
		For i=1 To Ubound(arrCountryID)
			If arrCountryEU(i)=1 Then Response.Write("<option value=" & arrCountryID(i) & ">" & arrCountryName(i) & "</option>" & vbCrLf)
		Next
		For i=1 To Ubound(arrCountryID)
			If arrCountryEU(i)=0 Then Response.Write("<option value=" & arrCountryID(i) & ">" & arrCountryName(i) & "</option>" & vbCrLf)
		Next
		%>
		</select>
		<p class="sml">(press [Ctrl] for multiple selection)</p>
		</td></tr>

		<tr><td class="field splitter"><label for="subject">Education</label></td>
		<td class="value blue"><select id="subject" name="subject" multiple size="4" style="width: 355px;">
		<%
		For i=1 To Ubound(arrEduSubjectID)
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
			<%
			For i=1 To Ubound(arrLanguageID)
				Response.Write("<option value=" & arrLanguageID(i) & ">" & arrLanguageTitle(i) & "</option>" & vbCrLf)
			Next
			%>
			</select></td>
			<td>&nbsp;</td>
			<td><p class="sml">Level of knowledge</p>
			<select id="language_level" name="language_level" size="1" style="width: 120px;">
			<option></option>
			<%
			For i=1 To Ubound(arrLanguageLevelID)
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

		</table>
		</div>

		<input class="button first" type="image" src="<% =sHomePath %>image/bte_search.gif" name="Search" alt="Search">
		<a href="<%=sScriptFullName & AddUrlParams(sParams, "act=" & sAccessType) %>"><img class="button" src="<% =sHomePath %>image/bte_clearall.gif" alt="Clear all"></a>
		<br/><br/>
		
	<input type="hidden" name="mmb_sct_hid">
	<input type="hidden" name="mmb_don_hid">
	<input type="hidden" name="mmb_cou_hid">
	<input type="hidden" name="srch_type" value="advanced">
	</form>
	</div>
</div>
</div>
	<!-- footer -->
	<!--#include virtual="/_template/page.footer.asp"-->
<script language=JavaScript type=text/javascript>
scrollInit(1,1,1);
</script>
<% CloseDBConnection %>
</body>
<!--#include virtual="/_template/html.footer.asp"-->
