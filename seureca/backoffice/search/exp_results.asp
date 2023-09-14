<%@ LANGUAGE="VBSCRIPT" %>
<% Option Explicit 
'--------------------------------------------------------------------
'
' Search for Experts
'
'--------------------------------------------------------------------
Response.Expires = 0
Response.ExpiresAbsolute = Now() - 1
Response.CacheControl = "no-cache"
%>
<!--#include file="../../dbc.asp"-->
<!--#include file="../../fnc.asp"-->
<!--#include file="../../fnc_exp.asp"-->
<!--#include file="../../../_common/_data/en/lib_seureca.asp"-->
<!--#include file="../../_forms/frmInterface.asp"-->
<!--#include file="../../../_common/_data/datMonth.asp"-->
<%
' Remove inactive url params
sParams=ReplaceUrlParams(sParams, "qid")
sParams=ReplaceUrlParams(sParams, "page")
sParams=ReplaceUrlParams(sParams, "x")
sParams=ReplaceUrlParams(sParams, "y")
sParams=ReplaceUrlParams(sParams, "idproject")
'sParams=ReplaceUrlParams(sParams, "srch_orderby")
sParams=ReplaceUrlParams(sParams, "srch_queryadd")

' Check user's access rights
CheckUserLogin sScriptFullNameAsParams

Dim iMmbAccountStatusExp
iMmbAccountStatusExp=1

Dim iSearchQueryID, sSearchFullText, sSearchSeniority, j
Dim iNumberCVsSelected, iNumberCVsDownloaded, iNumberCVsDownloadedFromSelected, iNumberCVsSubscribedFor, iNumberCVsInOptimalPackage, iNumberCVsInNextPackage, iNumberCVsTotalInResult

Dim colExpBorder, colExpNum, colExpTitle, colExpColumn, colExpDetails0, colExpDetails1
Dim sAvailability, sPreferences, sKeyQualifications
Dim sSearchExperts, sSearchFirstname, sSearchSurname, sSearchKeywords, sSearchKeywordsType, sSearchDB, bSaveSearchLog
Dim sSearchNationality, sSearchEduSubject, sSearchNativeLng, sSearchOtherLng
Dim sSearchSectors, arrSearchSectors, sSearchMainSectors, arrSearchMainSectors
Dim sSearchCountries, arrSearchCountries, sSearchRegions, arrSearchRegions
Dim sSearchDonors, arrSearchDonors, sSearchCurrentlyIn, iSearchPastYears, iSearchPastProjects
Dim sExpertIds, arrExpertIds, iCurrentPage, iTotalRecords, iTotalPages, sSearchKeywordsAdd, sSearchOrderBy
Dim sSearchAction
Dim sCvFolder

Dim iNumberCVsDownloadedToday

Dim iNoticeLimitDownloadedToday
Dim iStopLimitDownloadedToday
iNoticeLimitDownloadedToday=10
iStopLimitDownloadedToday=30

iSearchQueryID=ReplaceIfEmpty(Request.Form("qid"), Request.QueryString("qid"))
bSaveSearchLog=1

iCurrentPage=Request.QueryString("page")
If Not IsNumeric(iCurrentPage) or iCurrentPage="" Then
	iCurrentPage=1
Else
	iCurrentPage=CInt(iCurrentPage)
End If

Dim iProjectID
If Not (iSearchQueryID>0 And IsNumeric(iSearchQueryID)) Then

	iProjectID=CheckIntegerAndZero(Request.Form("idproject"))
	sParams=ReplaceUrlParams(sParams, "idproject=" & iProjectID)
	sParams=ReplaceUrlParams(sParams, "qid=" & iSearchQueryID)

	sSearchFullText=""
	sSearchExperts=Request.QueryString("eid")
	If sSearchExperts>"" Then
		sSearchFullText=" </i>expert CVs     "
	End If

	sSearchFirstname=CheckString(Request.Form("srch_firstname"))
	sSearchSurname=CheckString(Request.Form("srch_surname"))

	sSearchKeywords=CheckString(Request.Form("srch_query"))
	If sSearchKeywords>"" Then sSearchKeywords=Left(sSearchKeywords, 100)
	sSearchKeywordsHighlight=sSearchKeywords
	If sSearchKeywords>"" Then
		sSearchKeywords=Left(sSearchKeywords, 150)
		sSearchKeywordsType=CheckString(Request.Form("srch_querytype"))
	End If

	sSearchKeywords=CreateSearchString(sSearchKeywords, sSearchKeywordsType)

	sSearchNationality=CheckString(Request.Form("nationality"))
	sSearchEduSubject=CheckString(Request.Form("subject"))
	sSearchCurrentlyIn=CheckInteger(Request.Form("currentlyin"))
	iSearchPastYears=CheckInteger(Request.Form("pastyears"))
	If iSearchPastYears=0 Or iSearchPastYears=100 Then iSearchPastYears=Null

	iSearchPastProjects=CheckInteger(Request.Form("pastprojects"))
	If iSearchPastProjects="" Or iSearchPastProjects=0 Or iSearchPastProjects=100 Then iSearchPastProjects=Null
	
	sSearchNativeLng=CheckString(Request.Form("nativelng"))
	sSearchOtherLng=CheckString(Request.Form("otherlng"))
	sSearchSeniority=CheckString(Request.Form("seniority"))

	sSearchSectors=CheckString(Request.Form("mmb_sct_hid"))
	sSearchSectors=Replace(sSearchSectors,",0","")
	If Len(sSearchSectors)>2 Then sSearchSectors=Right(sSearchSectors, Len(sSearchSectors)-2)
	If sSearchSectors="0" Then
		sSearchSectors=""
	End If
	If sSearchSectors>"" Then
		arrSearchSectors=Split(sSearchSectors,",")
	End If
	If Request.Form("srch_msectors")>"" Then
		sSearchMainSectors=CheckString(Request.Form("srch_msectors"))
	Else
		sSearchMainSectors=CheckString(Request.QueryString("srch_msectors"))
	End If
	If sSearchMainSectors="0" Then
		sSearchMainSectors=""
		sSearchFullText=sSearchFullText & "all sectors AND "
	End If
	If sSearchMainSectors>"" Then
		arrSearchMainSectors=Split(sSearchMainSectors,",")
	End If


	sSearchCountries=CheckString(Request.Form("mmb_cou_hid"))
	sSearchCountries=Replace(sSearchCountries,",0","")
	If Len(sSearchCountries)>2 Then sSearchCountries=Right(sSearchCountries, Len(sSearchCountries)-2)
	If sSearchCountries="0" Then
		sSearchCountries=""
	End If
	If sSearchCountries>"" Then
		arrSearchCountries=Split(sSearchCountries,",")
	End If
	
	If Request.Form("srch_regions")>"" Then
		sSearchRegions=CheckString(Request.Form("srch_regions"))
	Else
		sSearchRegions=CheckString(Request.QueryString("srch_regions"))
	End If
	If sSearchRegions="0" Then
		sSearchRegions=""
		sSearchFullText=sSearchFullText & "all regions AND "
	End If
	If sSearchRegions=>"" Then
		arrSearchRegions=Split(sSearchRegions,",")
	End If


	sSearchDonors=CheckString(Request.Form("mmb_don_hid"))
	sSearchDonors=Replace(sSearchDonors,",0","")
	If Len(sSearchDonors)>2 Then sSearchDonors=Right(sSearchDonors, Len(sSearchDonors)-2)
	If sSearchDonors="0" Then
		sSearchDonors=""
	End If
	If sSearchDonors>"" Then
		arrSearchDonors=Split(sSearchDonors,",")
	End If


	If (sSearchSectors>"0") Then
		For i=1 To UBound(arrSearchSectors)
			For j=0 To aExT
			If aExTCode(j)=CInt(arrSearchSectors(i)) Then
			sSearchFullText=sSearchFullText & aExTInfo(j) & ", "
			End If
			Next
		Next
               	If Len(sSearchFullText)>2 Then sSearchFullText=Left(sSearchFullText,Len(sSearchFullText)-2) & " AND "
	End If
	If (sSearchMainSectors>"") Then
		For i=0 To UBound(arrSearchMainSectors)
			For j=0 To aExF
			If aExFCode(j)=CInt(arrSearchMainSectors(i)) Then
			sSearchFullText=sSearchFullText & aExFInfo(j) & ", "
			End If
			Next
		Next
               	If Len(sSearchFullText)>2 Then sSearchFullText=Left(sSearchFullText,Len(sSearchFullText)-2) & " AND "
	End If

	If (sSearchCountries>"0" And sSearchCountries>"") Then
		For i=1 To UBound(arrSearchCountries)
			For j=0 To aNt
			If aNtCode(j)=CLng(arrSearchCountries(i)) Then
			sSearchFullText=sSearchFullText & aNtInfo(j) & ", "
			End If
			Next
		Next
               	If Len(sSearchFullText)>2 Then sSearchFullText=Left(sSearchFullText,Len(sSearchFullText)-2) & " AND "
	End If
	If (sSearchRegions>"") Then
		For i=1 To Request.Form("srch_regions").Count
			For j=0 To aGZ
			If aGZnCode(j)=CInt(Request.Form("srch_regions")(i)) Then
			sSearchFullText=sSearchFullText & aGZnInfo(j) & ", "
			End If
			Next
		Next
               	If Len(sSearchFullText)>2 Then sSearchFullText=Left(sSearchFullText,Len(sSearchFullText)-2) & " AND "
	End If

	If (sSearchDonors>"0") Then
		For i=1 To UBound(arrSearchDonors)
			For j=0 To aOrg
			If aOrgCode(j)=CLng(arrSearchDonors(i)) Then
			sSearchFullText=sSearchFullText & aOrgInfo(j) & ", "
			End If
			Next
		Next
               	If Len(sSearchFullText)>2 Then sSearchFullText=Left(sSearchFullText,Len(sSearchFullText)-2) & " AND "
	End If

	If Len(sSearchFullText)>5 Then
		sSearchFullText="<i>" & Left(sSearchFullText,Len(sSearchFullText)-5) & "</i>"
	ElseIf sSearchFullText="" Then
		sSearchFullText=" your search query"
	End If

	sCvLanguage=Left(Request.Form("cv_language"), 3)
	sCvFolder=Left(Request.Form("cv_type"), 150)
	
	objTempRs=GetDataOutParamsSP("usp_MmbExpSearchFirstSelect", Array( _
		Array(, adInteger, , iMemberID), _
		Array(, adVarChar, 150, sSearchExperts), _
		Array(, adVarWChar, 150, sSearchFirstname), _
		Array(, adVarWChar, 150, sSearchSurname), _
		Array(, adVarWChar, 1000, sSearchKeywords), _
		Array(, adVarChar, 1400, sSearchNationality), _
		Array(, adVarChar, 350, sSearchEduSubject), _
		Array(, adVarChar, 700, sSearchNativeLng), _
		Array(, adVarChar, 700, sSearchOtherLng), _
		Array(, adVarChar, 50, sSearchSeniority), _
		Array(, adVarChar, 1000, sSearchCountries), _
		Array(, adVarChar, 150, sSearchRegions), _
		Array(, adVarChar, 1600, sSearchSectors), _
		Array(, adVarChar, 150, sSearchMainSectors), _
		Array(, adVarChar, 200, sSearchDonors), _
		Array(, adVarChar, 10, sSearchDB), _
		Array(, adInteger, , sSearchCurrentlyIn), _
		Array(, adInteger, , iSearchPastYears), _
		Array(, adInteger, , iSearchPastProjects), _
		Array(, adVarChar, 3, sCvLanguage), _
		Array(, adVarChar, 150, sCvFolder), _
		Array(, adTinyInt, , 0), _
		Array(, adTinyInt, , 0), _
		Array(, adTinyInt, , bSaveSearchLog)), _
		Array(Array(, adInteger)))

	iSearchQueryID=objTempRs(0)

	' To avoid a page with 'Page has Expired' message
	sTempParams=sParams
	sTempParams=ReplaceUrlParams(sTempParams, "qid=" & iSearchQueryID) 
	sTempParams=ReplaceUrlParams(sTempParams, "page=1")
	sTempParams=ReplaceUrlParams(sTempParams, "txt=" & sSearchKeywordsHighlight)

	Response.Clear
	Response.Redirect sScriptFileName & sTempParams

	'Set objTempRs=GetDataRecordsetSP("usp_MmbExpSearchRepeatSelect", Array( _
	'	Array(, adInteger, , iMemberID), _
	'	Array(, adInteger, , iSearchQueryID), _
	'	Array(, adVarWChar, 255, Null), _
	'	Array(, adVarChar, 50, Null)))

Else
	Response.Flush

	iProjectID=CheckIntegerAndZero(Request.QueryString("idproject"))
	sParams=ReplaceUrlParams(sParams, "idproject=" & iProjectID)
	sParams=ReplaceUrlParams(sParams, "qid=" & iSearchQueryID)

	sSearchAction=Request.QueryString("act")
	iSearchQueryID=Request.QueryString("qid")
	sSearchKeywordsAdd=Request.QueryString("srch_queryadd")
	sSearchOrderBy=Request.QueryString("srch_orderby")

	Set objTempRs=GetDataRecordsetSP("usp_MmbExpSearchRepeatSelect", Array( _
		Array(, adInteger, , iMemberID), _
		Array(, adInteger, , iSearchQueryID), _
		Array(, adVarWChar, 255, ReplaceIfEmpty(CreateSearchString(sSearchKeywordsAdd, "all of the words from"),Null)), _
		Array(, adVarChar, 50, ReplaceIfEmpty(sSearchOrderBy,Null))))
	sSearchFullText=Request.Form("srch_fullquery")

End If

If Request.Form("ExpertIds")<>"" Then
	sExpertIds=Request.Form("ExpertIds")
ElseIf Request.Cookies("ExpertIds")>"" Then
	sExpertIds=Request.Cookies("ExpertIds")
ElseIf sSearchExperts>"" Then
	sExpertIds=sSearchExperts
Else
	sExpertIds="0"
End If

' Clean up the list from already downloaded experts
	objTempRs2=GetDataOutParamsSP("usp_MmbExpDownloadedCleanup", Array( _
		Array(, adInteger, , iMemberID), _
		Array(, adVarChar, 4000, sExpertIds)), _
		Array(Array(, adVarChar, 4000)))
	sExpertIds=objTempRs2(0)

arrExpertIds=Split(sExpertIds,",")

If sExpertIds>"0" Then
	sExpertIds=Replace(Trim(sExpertIds), " ,", ",")
	sExpertIds=Replace(sExpertIds, ",", " ,") & " "
End If

sSearchKeywordsHighlight=Request.QueryString("txt")

'If Not InStr(sParams, "qid=")>1 Then sParams=AddUrlParams(sParams, "qid=" & iSearchQueryID)
%>


<html>
<head>
<script language="JavaScript">
function CheckExpertIds(expert_id) {
var selected_expert_ids='';
var active_expert_element='id_' + expert_id;
var active_expert_string=',' + (expert_id || 0);
	if (document.forms[0].ExpertIds) {
		selected_expert_ids=document.forms[0].ExpertIds.value;

		if (document.forms[0][active_expert_element] && document.forms[0][active_expert_element].checked) {
		
			if (!VerifyCvsDownloadedToday(<% =iStopLimitDownloadedToday %>)) {
				if (document.forms[0][active_expert_element] && document.forms[0][active_expert_element].checked) {
					document.forms[0][active_expert_element].checked=false;
				}
			} else {
				selected_expert_ids += active_expert_string;
			}

		} else {
			selected_expert_ids = selected_expert_ids.replace(active_expert_string, '');
		}
		document.forms[0].ExpertIds.value=selected_expert_ids;
		document.cookie = 'ExpertIds=' + selected_expert_ids;
	}
}	

function VerifyCvsDownloadedToday(limit) {
var selected_expert_ids='';
	if (document.forms[0].ExpertIds) {
		selected_expert_ids=document.forms[0].ExpertIds.value;

		if ((<% =ReplaceIfEmpty(iNumberCVsDownloadedToday, 0) %> >= limit) || (<% =ReplaceIfEmpty(iNumberCVsDownloadedToday, 0) %> + selected_expert_ids.split(',').length > limit)) {
			alert('You cannot download more CVs!\nPlease contact us for further information regarding this case.');
			return false;
		}
	}
	return true;
}

function DownloadCV(cv_num, eid, lng) {
	var f = document.forms[0];
	var obj_cv = f["cvformat" + cv_num];
	var frm = obj_cv.options[obj_cv.selectedIndex].value;
	// The default CV format is EC format
	if (obj_cv.selectedIndex==0) {
		frm="EC";
	}
	var url = "../view/cv_view" + frm + ".asp<% =ReplaceUrlParams(ReplaceUrlParams(sParams, "id"), "t=1") %>&id=" + eid;
	window.location.href=url;
}

function Register(regtype) {
	if (!VerifyCvsDownloadedToday(<% =iStopLimitDownloadedToday %>+1)) {
		return;
	}
	if (document.forms[0].ExpertIds.value.replace(' /g', '').length<=1) {
		alert("Please select at least one expert");
		return;
	} 
	document.forms[0].action="exp_register.asp?act=selected"; 
	document.forms[0].submit(); 
} 
// -->
</script>
<script type="text/javascript" src="/_scripts/js/jquery.js" charset="utf-8"></script>
<script type="text/javascript" charset="utf-8">
// Search query id
var query_id="<% =iSearchQueryID %>";
// Number of selected experts
var expert_query_selection_count; 
// The comma separated list with selected experts
var expert_query_selection_list;
// The array of selected experts
var expert_query_selection = [];

$(document).ready(function(){

	InitialiseSelection();
	
	// Initialise the number of experts selected
	function InitialiseSelection() {
		expert_query_selection_count=0; 

		expert_query_selection_list = $("#expert_query_selection_list").text();
		if (expert_query_selection_list.length>1) {
			expert_query_selection = expert_query_selection_list.split(",");
			expert_query_selection_count=expert_query_selection.length; 
			
			$("#expert_query_selection_count").text(expert_query_selection_count);
		} else {
			expert_query_selection_count = 0;
			$("#expert_query_selection_count").text("No");
		}

		// Check selected experts
		for (var i=0; i<expert_query_selection_count; i++) {
			$("#expert_" + expert_query_selection[i] + "_query_" + query_id + "_selection").attr("checked", true);
		}
		
		RefreshClearSelection();	
	}
	
	// Uninitialise the experts selected
	function UninitialiseSelection() {
		expert_query_selection_count=0; 

		expert_query_selection_list = $("#expert_query_selection_list").text();
		if (expert_query_selection_list.length>1) {
			expert_query_selection = expert_query_selection_list.split(",");
			expert_query_selection_count=expert_query_selection.length; 
		}

		// Uncheck selected experts
		for (var i=0; i<expert_query_selection_count; i++) {
			$("#expert_" + expert_query_selection[i] + "_query_" + query_id + "_selection").attr("checked", false);
		}
		
		// Empty the list
		expert_query_selection = [];
		$("#expert_query_selection_list").text("");
		
		// Empty the db
		$.get("http://<% =sScriptServerName & sScriptBaseName %>exp_select.asp?act=0&qid=" + query_id);
	}
	
	function RefreshClearSelection() {
		// If there are selected experts - show clear selection button
		if (expert_query_selection_count>0) {
			$("#expert_query_selection_remove").show();
		} else {
			$("#expert_query_selection_remove").hide();
		}
	}

	$(".expert_query_selection").click(function(){
		// Get expert and query ids from dom element id
		var element_id=$(this).attr("id").replace("_selection", "");
		var expert_id=element_id.replace("expert_", "");
		expert_id=expert_id.substring(0, expert_id.indexOf("_"));

		// Get the current number of experts selected
		var expert_query_selection_count_text = $("#expert_query_selection_count").text();
		expert_query_selection_count=0; 
		
		if (!isNaN(expert_query_selection_count_text)) { 
			expert_query_selection_count=parseInt(expert_query_selection_count_text); 
		}
		
		// Save the selection in the db
		if ($("#expert_" + expert_id + "_query_" + query_id + "_selection").attr("checked")==true) {
			// Add expert id to the array
			expert_query_selection.push(expert_id);
			expert_query_selection_list = expert_query_selection.join(",");
			// Update the list
			$("#expert_query_selection_list").text(expert_query_selection_list);
			expert_query_selection_count = expert_query_selection_count + 1;
			// Update the db
			$.get("http://<% =sScriptServerName & sScriptBaseName %>exp_select.asp?act=1&eid=" + expert_id + "&qid=" + query_id)
		} else {
			// Remove expert id from the array
			var pos = expert_query_selection.indexOf(expert_id);
			if (pos>=0) {
				expert_query_selection.splice(pos, 1);
			}
			expert_query_selection_list = expert_query_selection.join(",");
			// Update the list
			$("#expert_query_selection_list").text(expert_query_selection_list);
			expert_query_selection_count = expert_query_selection_count - 1;
			// Update the db
			$.get("http://<% =sScriptServerName & sScriptBaseName %>exp_select.asp?act=0&eid=" + expert_id + "&qid=" + query_id)
		}

		if (expert_query_selection_count<0) { 
			expert_query_selection_count=0; 
			}
		
		// Show the current number of experts selected
		$("#expert_query_selection_count").text(expert_query_selection_count);
		
		RefreshClearSelection();
	});


	$("#a_expert_query_selection_remove").click(function(){
		UninitialiseSelection();
		InitialiseSelection();
	});
});

function SendEmailAllExperts() {
	url="../manage/send_email.asp<% =ReplaceUrlParams(sParams, "qid=" & iSearchQueryID) %>&select=all"
	var wnd=window.open(url, "cvip_email", "width=800, height=600, left=" + ((screen.availWidth-650)/2) + ", top=" + ((screen.availHeight-450)/2) + ", scrollbars=no, status=no, location=no, menubar=no, toolbar=no, resizable=yes");
	wnd.focus();
	
	self.location.reload();
}

function SendEmailExperts() {
	if (expert_query_selection_count==0) {
		alert("Please select experts by ticking the corresponding checkboxes before sending emails.");
		return;
	}

	url="../manage/send_email.asp<% =ReplaceUrlParams(sParams, "qid=" & iSearchQueryID) %>"
	var wnd=window.open(url, "cvip_email", "width=800, height=600, left=" + ((screen.availWidth-650)/2) + ", top=" + ((screen.availHeight-450)/2) + ", scrollbars=no, status=no, location=no, menubar=no, toolbar=no, resizable=yes");
	wnd.focus();
}	
</script>

<title>Search result - experts database</title>

<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" type="text/css" href="<% =sHomePath %>styles.css">
</head>

<body bgcolor="#FFFFFF" topmargin=0 leftmargin=0 marginheight=0 marginwidth=0>
<% ShowTopMenu %>

<% If objTempRs.Eof Then %>
<!-- [i] No Results -->
	<br><br>
	<% ShowMessageStart "info", 620 %>
	There are <span class="rs">no experts</span> matching your search query.<br>
	Please revise your query and try again. If you have questions please contact us at <a href="mailto:info@assortis.com">info@assortis.com</a>.
	<% ShowMessageEnd %>
	<br><div align="center"><a href="exp_search.asp<% =sParams %>"><img src="<% =sHomePath %>image/bte_searchagain.gif" name="Search again" height=18 width=105 alt="Search again" border=0></a></div>

<% Else 
	If IsNull(objTempRs(0)) Then
		Response.Write("<br><br>")
		ShowMessageStart "info", 500 %>
			The searh query is expired. Please revise your query and try again. <br>If you have questions please contact us at <a href="mailto:info@assortis.com">info@assortis.com</a>.
		<% ShowMessageEnd
		Response.End
	End If
	iNumberCVsTotalInResult=objTempRs.RecordCount %>

   <!-- [i] Results overview -->
	<% ShowMessageStart "info", 650 %>
	<% If iNumberCVsTotalInResult=1 Then %>There is <span class="rs">1</span> expert<% Else %>There are <span class="rs"><%=iNumberCVsTotalInResult%></span> experts<% End If %> matching your search query.
	<div align="center"><a href="exp_search.asp<% =sParams %>"><img src="<% =sHomePath %>image/bte_searchagain.gif"  name="Search again" height=18 width=105 alt="Search again" border=0 hspace=10 vspace=3></a></div>
	</ul>
	<% ShowMessageEnd %>

	</p>

	<%
	objTempRs.PageSize=10
	iTotalRecords=objTempRs.RecordCount
	iTotalPages=objTempRs.PageCount

	ShowNavigationPages iCurrentPage, iTotalPages, sParams
%>
	<!--
	<table width=750 cellspacing=0 cellpadding=0 border=0 align="center"><tr>
	<td width="45%" align="left" valign="top"><a href="#"><img src="<% =sHomePath %>image/bte_searchagain.gif" height=18 width=105 alt="Search again" border=0 hspace=20 vspace=12 align="left"></a><img src="<% =sHomePath %>image/x.gif" width=40 hieght=1></td>
	<td><% ShowNavigationPages iCurrentPage, iTotalPages, sParams %>
	<td width="50%" align="right" valign="top"><a href="exp_register.asp<% =sParams %>"><img src="<% =sHomePath %>image/bte_downloadselected.gif" height=18 width=166 alt="Confirm CVs selection" border=0 hspace=20 vspace=12 align="right"></a><img src="<% =sHomePath %>image/x.gif" width=40 hieght=1></td>
	</tr></table>
	-->

<table width="98%" cellpadding=0 cellspacing=0 border=0 align="center">
<form name="ExpList" method="post">
<input type="hidden" name="qid" value="<%=iSearchQueryID%>"><input type="hidden" name="srch_fullquery" value='<%=sSearchFullText%>'>

<% sExpertIds=ReplaceIfEmpty(sExpertIds, "0") %>
<input type=hidden name="ExpertIds" value="<%=sExpertIds%>">

<tr><td width="85%" valign="top">

	<% ShowExpertsBlock 1, iCurrentPage %>

</td>
</form>
<td width="5%">&nbsp;&nbsp;</td>
<td width="20%" valign="top">
   <!-- Feature boxes -->

	<form name="search" method="get" action="<% =sScriptFileName %>">
	<input type="hidden" name="qid" value="<% =iSearchQueryID %>">
	<input type="hidden" name="idproject" value="<% =iProjectID %>">
	<% ShowFeatureBoxHeader "Refine&nbsp;your&nbsp;search" %>
	<p class="sml" align="center">Order these results by</p>
	<img src="<% =sHomePath %>image/x.gif" width=1 height=3><br>
	<div align="center"><select style="font-face: Arial; font-size:8.5pt;" name="srch_orderby" size=1>
	<option value="expRank"<% If sSearchOrderBy="expRank" Then %> selected<% End If %>>Relevance &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;</option>
	<option value="expLastName"<% If sSearchOrderBy="expLastName" Then %> selected<% End If %>>Last name</option>
	<option value="expSeniority"<% If sSearchOrderBy="expSeniority" Then %> selected<% End If %>>Seniority</option>
	</select></div>
	<img src="<% =sHomePath %>image/x.gif" width=1 height=6><br>
	<div align="center"><input type="image" src="<% =sHomePath %>image/bte_order.gif" alt="Search for experts" height=18 vspace=0 border=0></a></div>
	<% ShowFeatureBoxDelimiter %>
	<p class="sml" align="center">Search within results</p>
	<img src="<% =sHomePath %>image/x.gif" width=1 height=3><br>
	<div align="center"><input type="text" style="font-face: Arial; font-size:8.5pt;" name="srch_queryadd" size=16 value="<%=sSearchKeywordsAdd%>"></div>
	<img src="<% =sHomePath %>image/x.gif" width=1 height=6><br>
	<div align="center"><input type="image" src="<% =sHomePath %>image/bte_search.gif" alt="Search for experts" height=18 vspace=0 border=0></a></div>
	<% ShowFeatureBoxFooter %>
	</form>
	
	<% ShowFeatureBoxHeader "Advanced&nbsp;options" %>
	<div align="center"><p class="sml"><span id="expert_query_selection_count">No</span> expert(s) selected <span id="expert_query_selection_remove">(<a href="#" id="a_expert_query_selection_remove">X</a>)</span></p></div>
	<img src="../../image/x.gif" width=1 height=5><br>
	<div align="center"><p class="sml"><a href="javascript:SendEmailExperts();">Send emails to selected experts</a></p></div>
	<img src="../../image/x.gif" width=1 height=15><br>
	<div align="center"><p class="sml"><a href="javascript:SendEmailAllExperts();">Send emails to all experts</a></p></div>
	<img src="../../image/x.gif" width=1 height=5><br>
	<span id="expert_query_selection_list" style="display: none;"><% =ShowMemberExpertQuery(iMemberID, 0, iSearchQueryID) %></span>
	<% ShowFeatureBoxFooter %>

</td>
</tr>
</table>

<% ShowNavigationPages iCurrentPage, iTotalPages, sParams %>
<% End If 
objTempRs.Close %>

<% CloseDBConnection %>
</body>
</html>

