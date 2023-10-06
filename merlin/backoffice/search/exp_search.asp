<%@ LANGUAGE="VBSCRIPT" %>
<% 'Option Explicit
'--------------------------------------------------------------------
'
' Search for experts
'
'--------------------------------------------------------------------
%>
<!--#include file="../../dbc.asp"-->
<!--#include file="../../fnc.asp"-->
<!--#include file="../../../_common/_data/en/lib.asp"-->
<!--#include file="../../_forms/frmInterface.asp"-->
<!--#include file="../../_forms/frmScrollBox.asp"-->
<!--#include file="../../../_common/_data/datCountry.asp"-->
<!--#include file="../../../_common/_data/datLanguage.asp"-->
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
<html>
<head>
<title>Expert database. Advanced search</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">

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

	// any criteria have to be filled in
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
		&& (f.pastyears.selectedIndex<1)
		&& (f.pastprojects.selectedIndex<1)
	<% If bCvMultiLanguageActive = cMultiLanguageEnabled Then %>
		&& (f.cv_language.selectedIndex<1)
	<% End If %>
	<% If bCvTypeActive = cCvTypeEnabled Then %>
		&& (f.cv_type.selectedIndex<1)
	<% End If %>
		) {	
			alert('Please fill in a search criteria.');
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
<link rel="stylesheet" type="text/css" href="<% =sHomePath %>styles.css">
</head>

<body bgcolor="#FFFFFF" topmargin=0 leftmargin=0 marginheight=0 marginwidth=0 onLoad="RestoreInt();">
<% ShowTopMenu %>
<% ShowWaitMessage %>

<!-- The registration progress bar -->
<% If iMmbAccountStatusExp>0 Then ShowRegistrationProgressBar "EXP" & iMmbAccountStatusExp, 1 %>

<!-- Keywords search -->
	<% ShowInputFormHeader 580, "KEYWORDS SEARCH" %>
	<form action="exp_results.asp" method="post" name="RegForm" onSubmit="Continue(); return false;">
		<tr><td width="170" valign="top"><p class="ftxt">Expert name</p></td>
		<td bgcolor="<% =colFormBodyRight %>" width=1><img src="x.gif" width=1 height=32></td>
		<td bgcolor="<% =colFormBodyText %>" width=407><img src="x.gif" width=1 height=4><br>
		
		<table width="360" cellpadding="0" cellspacing="0" border="0">
		<tr valign="top">
		<td width="160">&nbsp;&nbsp;<input type="text" name="srch_firstname" size=13 maxlength=100 style="width:145px;"></td>
		<td width="220" colspan="2"><input type="text" name="srch_surname" size=13 maxlength=100 style="width:200px;"></td>
		</tr>
		<tr>
		<td><p class="sml">&nbsp;&nbsp;First name</p></td>
		<td><p class="sml">&nbsp;Surname&nbsp;</p></td>
		<td align="right"><p class="sml" align="right">ID&nbsp;</p></td>
		</tr>
		</table>
		<img src="<% =sHomePath %>image/x.gif" width=358 height=1 vspace=3><br>
		</td></tr>
		
		<tr><td width="170" valign="top"><p class="ftxt">Keyword search</p></td>
		<td bgcolor="<% =colFormBodyRight %>" width=1><img src="x.gif" width=1 height=32></td>
		<td bgcolor="<% =colFormBodyText %>" width=407><img src="x.gif" width=1 height=4><br>
		&nbsp;&nbsp;<select name="srch_querytype" style="width:145px;">
		<option selected value="all of the words from">all of the words</option><option value="any of the words from">any of the words</option><option value="the exact phrase">the exact phrase</option><option value="boolean expression">boolean expression</option></select>&nbsp;&nbsp;<input type="text" name="srch_query" size=23 maxlength=100 style="width:200px;">
		<p class="sml">&nbsp;This searches the entire content of all the online CVs.</p>
		<img src="<% =sHomePath %>image/x.gif" width=358 height=1 vspace=3><br>
		</td></tr>
	<% ShowInputFormFooter 580 %>
	<br>

	<table width=580 cellspacing=0 cellpadding=0 border=0 align="center">
	<tr>
	<td width=400 align=left>
	<img src="<% =sHomePath %>image/x.gif" width=170 height=1><input type="image" src="<% =sHomePath %>image/bte_search.gif" name="search" alt="Search" border=0></a>
	<img src="<% =sHomePath %>image/x.gif" width=30 height=1><a href="<%=sScriptFileName & sParams%>"><img src="<% =sHomePath %>image/bte_clearall.gif" border=0 alt="Clear All" name="Clear All"></a>
	</td>
	<td width=180 align="center"></td>
	</tr>
	</table><br>


  <!-- Sectors section -->
	<% ShowSctScrollBox "SECTORS OF EXPERTS' EXPERIENCE",  "", 1, 0, 1, 1, 0 %><br>

  <!-- Countries section -->
	<% ShowCouScrollBox "COUNTRIES OF EXPERTS' EXPERIENCE",  "", 1, 0, 1, 1, 0 %><br>
	
  <!-- Funding agancies section -->
	<% ShowDonScrollBox "FUNDING AGENCIES OF EXPERTS' EXPERIENCE",  "", 1, 0, 1, 1, 0, 0 %><br>

  <!-- Search options -->
	<% ShowInputFormHeader 580, "SEARCH OPTIONS" %>
		<tr><td width="170" valign="top"><p class="ftxt">Professional experience<br>gained in the past</p></td>
		<td bgcolor="<% =colFormBodyRight %>" width=1><img src="x.gif" width=1 height=32></td>
		<td bgcolor="<% =colFormBodyText %>" width=407><img src="x.gif" width=1 height=4><br>
		&nbsp;&nbsp;<select name="pastyears" size=1 style="width:355px;">
		<option value="0" selected> </option>
		<%
			Response.Write("<option value=""1"">12 months</option>" & vbCrLf)
			Response.Write("<option value=""2"">2 years</option>" & vbCrLf)
			Response.Write("<option value=""4"">4 years</option>" & vbCrLf)
			Response.Write("<option value=""100"">&gt;4 years (entire career)</option>" & vbCrLf)
		%>
		</select><img src="<% =sHomePath %>image/x.gif" width=1 height=9><br>
		</td></tr>

		<tr><td width="170" valign="top"><p class="ftxt">Consider experience on<br>at least</p></td>
		<td bgcolor="<% =colFormBodyRight %>" width=1><img src="x.gif" width=1 height=32></td>
		<td bgcolor="<% =colFormBodyText %>" width=407><img src="x.gif" width=1 height=4><br>
		&nbsp;&nbsp;<select name="pastprojects" size=1 style="width:355px;">
		<option value="0" selected> </option>
		<option value="5">5 projects</option>
		<option value="10">10 projects</option>
		<option value="15">15 projects</option>
		</select>
		</td></tr>

		<tr><td width="170" valign="top"><p class="ftxt">Currently working in</p></td>
		<td bgcolor="<% =colFormBodyRight %>" width=1><img src="x.gif" width=1 height=32></td>
		<td bgcolor="<% =colFormBodyText %>" width=407><img src="x.gif" width=1 height=4><br>
		&nbsp;&nbsp;<select name="currentlyin" size=1 style="width:355px;">
		<option value="0"> </option>
		<%
		For i=1 To Ubound(arrCountryID)
			Response.Write("<option value=" & arrCountryID(i) & ">" & arrCountryName(i) & "</option>" & vbCrLf)
		Next
		%>
		</select><br><img src="<% =sHomePath %>image/x.gif" width=1 height=9><br>
		</td></tr>

		<tr><td width="170" valign="top"><p class="ftxt">Nationality</p></td>
		<td bgcolor="<% =colFormBodyRight %>" width=1><img src="x.gif" width=1 height=32></td>
		<td bgcolor="<% =colFormBodyText %>" width=407><img src="x.gif" width=1 height=4><br>
		&nbsp;&nbsp;<select name="nationality" size=6 multiple style="width:355px;">
		<option value="1100">EU Member States</option>
		<option value="1112">TACIS Countries</option>
		<option value="1113">CARDS Countries</option>
		<option value="1114">MEDA Countries</option>
		<%
		For i=1 To Ubound(arrCountryID)
			If arrCountryEU(i)=1 Then Response.Write("<option value=" & arrCountryID(i) & ">" & arrCountryName(i) & "</option>" & vbCrLf)
		Next
		For i=1 To Ubound(arrCountryID)
			If arrCountryEU(i)=0 Then Response.Write("<option value=" & arrCountryID(i) & ">" & arrCountryName(i) & "</option>" & vbCrLf)
		Next
		%>
		</select>
		<p class="sml">&nbsp;&nbsp;(press [Ctrl] for multiple selection)</p><br>
		</td></tr>

		<tr><td width="170" valign="top"><p class="ftxt">Education</p></td>
		<td bgcolor="<% =colFormBodyRight %>" width=1><img src="x.gif" width=1 height=32></td>
		<td bgcolor="<% =colFormBodyText %>" width=407>
		&nbsp;&nbsp;<select  size="4" name="subject" multiple style="width:355px;">
		<%
		For i=1 To Ubound(arrEduSubjectID)
			Response.Write("<option value=" & arrEduSubjectID(i) & ">" & arrEduSubjectTitle(i) & "</option>" & vbCrLf)
		Next
		%>
		</select>
		<p class="sml">&nbsp;&nbsp;(press [Ctrl] for multiple selection)</p><br>
		</td></tr>

		<tr><td width="170" valign="top"><p class="ftxt">Languages</p></td>
		<td bgcolor="<% =colFormBodyRight %>" width=1><img src="x.gif" width=1 height=32></td>
		<td bgcolor="<% =colFormBodyText %>" width=407>
		&nbsp;&nbsp;<select  size="4" name="nativelng" multiple style="width:220px;">
		<%
		For i=1 To Ubound(arrLanguageID)
			Response.Write("<option value=" & arrLanguageID(i) & ">" & arrLanguageTitle(i) & "</option>" & vbCrLf)
		Next
		%>
		</select>
		<p class="sml">&nbsp;&nbsp;(press [Ctrl] for multiple selection)<br /></p>
		</td></tr>

		<tr><td width="170" valign="top"><p class="ftxt">Seniority</p></td>
		<td bgcolor="<% =colFormBodyRight %>" width=1><img src="x.gif" width=1 height=32></td>
		<td bgcolor="<% =colFormBodyText %>" width=407>
		&nbsp;&nbsp;<select name="seniority" style="width:130px;">
		<option value=""> </option>
		<option value="0 AND 5">less than 5 years</option>
		<option value="5 AND 100">over 5 years</option>
		<option value="10 AND 100">over 10 years</option>
		<option value="15 AND 100">over 15 years</option>
		<option value="20 AND 100">over 20 years</option>
		</select>
		</td></tr>
		
	<% If bCvMultiLanguageActive = cMultiLanguageEnabled Then %>
		<tr><td width="170" valign="top"><p class="ftxt">CV language</p></td>
		<td bgcolor="<% =colFormBodyRight %>" width=1><img src="x.gif" width=1 height=32></td>
		<td bgcolor="<% =colFormBodyText %>" width=407>
		&nbsp;&nbsp;<select name="cv_language" style="width:130px;">
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
		<tr><td width="170" valign="top"><p class="ftxt">CV type</p></td>
		<td bgcolor="<% =colFormBodyRight %>" width=1><img src="x.gif" width=1 height=32></td>
		<td bgcolor="<% =colFormBodyText %>" width=407>
		&nbsp;&nbsp;<select name="cv_type" style="width:130px;">
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
	<% ShowInputFormFooter 580 %><br>

	<table width=580 cellspacing=0 cellpadding=0 border=0 align="center">
	<tr>
	<td width=400 align=left>
	<img src="<% =sHomePath %>image/x.gif" width=170 height=1><input type="image" src="<% =sHomePath %>image/bte_search.gif" name="search" alt="Search" border=0></a>
	<img src="<% =sHomePath %>image/x.gif" width=30 height=1><a href="<%=sScriptFileName & sParams%>"><img src="<% =sHomePath %>image/bte_clearall.gif" border=0 alt="Clear All" name="Clear All"></a>
	</td>
	<td width=180 align="center"></td>
	</tr>
	<input type="hidden" name="idproject" value="<% =iProjectID %>">
	<input type="hidden" name="mmb_sct_hid">
	<input type="hidden" name="mmb_don_hid">
	<input type="hidden" name="mmb_cou_hid">
	<input type="hidden" name="srch_type" value="advanced">
	</form>
	</table><br>

 
<SCRIPT language=JavaScript type=text/javascript>
scrollInit(1,1,1);
</SCRIPT>

<% HideWaitMessage %>
<% CloseDBConnection %>
</body>
</html>
