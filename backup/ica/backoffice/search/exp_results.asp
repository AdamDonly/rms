<%@ LANGUAGE="VBSCRIPT" %>
<% Option Explicit 
'--------------------------------------------------------------------
'
' Search for Experts
'
'--------------------------------------------------------------------
%>
<!--#include virtual = "/_template/asp.header.nocache.asp"-->
<!--#include virtual = "/_template/asp.header.notimeout.asp"-->
<!--#include file="../../dbc.asp"-->
<!--#include file="../../fnc.asp"-->
<!--#include file="../../fnc_exp.asp"-->
<!--#include file="../../fnc_log.asp"-->
<!--#include file="../../_forms/frmInterface.asp"-->
<!--#include virtual = "/_common/_data/en/lib.asp"-->
<!--#include file="../../../_common/_data/datMonth.asp"-->
<!--#include virtual = "/_common/_data/datFlagMetho.asp"-->
<!--#include virtual = "/_common/_class/main.asp"-->
<!--#include virtual = "/_common/_class/expert.cv.language.asp"-->
<%
' Remove inactive url params
sParams=ReplaceUrlParams(sParams, "qid")
sParams=ReplaceUrlParams(sParams, "page")
sParams=ReplaceUrlParams(sParams, "x")
sParams=ReplaceUrlParams(sParams, "y")
sParams=ReplaceUrlParams(sParams, "idproject")
'sParams=ReplaceUrlParams(sParams, "srch_orderby")
'sParams=ReplaceUrlParams(sParams, "srch_queryadd")

' Check user's access rights
CheckUserLogin sScriptFullNameAsParams

Dim iMmbAccountStatusExp
iMmbAccountStatusExp=1

Dim iSearchQueryID, sSearchFullText, sSearchSeniority, j
Dim iNumberCVsSelected, iNumberCVsDownloaded, iNumberCVsDownloadedFromSelected, iNumberCVsSubscribedFor, iNumberCVsInOptimalPackage, iNumberCVsInNextPackage, iNumberCVsTotalInResult

Dim colExpBorder, colExpNum, colExpTitle, colExpColumn, colExpDetails0, colExpDetails1
Dim sAvailability, sPreferences, sKeyQualifications
Dim sSearchExperts, sSearchFirstname, sSearchSurname, sSearchKeywords, sSearchKeywordsType, sSearchDB, bSaveSearchLog, bSearchInhouseAgreed
Dim sSearchNationality, sSearchEduSubject, sSearchNativeLng, sSearchOtherLng, iSearchLanguageLevel
Dim sSearchSectors, arrSearchSectors, sSearchMainSectors, arrSearchMainSectors, bSearchSectorsSimultaneously
Dim sSearchCountries, arrSearchCountries, sSearchRegions, arrSearchRegions
Dim sSearchDonors, arrSearchDonors, sSearchAllDonors, sSearchCurrentlyIn, iSearchPastYears, iSearchPastProjects
Dim sExpertIds, arrExpertIds, iCurrentPage, iTotalRecords, iTotalPages, sSearchKeywordsAdd, sSearchOrderBy
Dim sSearchAction
Dim sCvFolder
Dim iSearchMethodologyWriters, iSearchMethodologyFlag

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
	If Len(sSearchSurname)>=2 And IsNumeric(sSearchSurname) Then
		sSearchExperts=sSearchSurname
		sSearchSurname=""
	End If
	
	Dim sSelectedExpertDatabaseCode
	sSelectedExpertDatabaseCode=""
	If Len(sSearchSurname)>4 Then 
		If InStr(sSearchSurname, "-")=4 Then
			sSelectedExpertDatabaseCode=Left(sSearchSurname, 4)
			sSearchExperts=CheckIntegerAndZero(Mid(sSearchSurname, 5, Len(sSearchSurname)))
			sSearchSurname=""

			Set objExpertDB = objExpertDBList.Find(sSelectedExpertDatabaseCode, "DatabaseCode")
			sSearchDatabases = objExpertDB.Database
		End If
	End If
	
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

	iSearchLanguageLevel=CheckInteger(Request.Form("language_level"))
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

	bSearchSectorsSimultaneously = 0
	If Request.Form("sectors_simultaneously")="on" Then
		bSearchSectorsSimultaneously = 1
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

	sSearchAllDonors = CheckString(Request.Form("don_0"))
	If sSearchAllDonors = "on" Then
		sSearchDonors = ""
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

	Dim sSearchDatabases
	If Len(Request.Form("database"))>0 Then
		sSearchDatabases = CheckString(Request.Form("database"))
		If InStr(sSearchDatabases, "all")>0 Then sSearchDatabases=""
	End If

	If bUserAccessMethodology Then
		iSearchMethodologyWriters = CheckboxOnOrZero(Request.Form("methodology"))
		iSearchMethodologyFlag = CheckIntegerAndNull(Request.Form("metho_flag"))
	Else
		iSearchMethodologyWriters = 0
		iSearchMethodologyFlag = 0
	End If

	objTempRs=GetDataOutParamsSP("usp_Ica_MemberExpertMethoSearchFirstSelect", Array( _
		Array(, adVarChar, 1500, sSearchDatabases), _
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
		Array(, adInteger, , iSearchLanguageLevel), _
		Array(, adVarChar, 3, sCvLanguage), _
		Array(, adInteger, , bSearchSectorsSimultaneously), _
		Array(, adTinyInt, , 0), _
		Array(, adTinyInt, , 0), _
		Array(, adTinyInt, , iSearchMethodologyWriters), _
		Array(, adTinyInt, , iSearchMethodologyFlag), _
		Array(, adTinyInt, , bSaveSearchLog)), _
		Array(Array(, adInteger)))

	iSearchQueryID=objTempRs(0)

	' To avoid a page with 'Page has Expired' message
	sTempParams=sParams
	sTempParams=ReplaceUrlParams(sTempParams, "qid=" & iSearchQueryID) 
	sTempParams=ReplaceUrlParams(sTempParams, "page=1")
	sTempParams=ReplaceUrlParams(sTempParams, "txt=" & sSearchKeywordsHighlight)

	' log search (36):
	iLogResult = LogActivity(36, "SearchQueryID=" & Cstr(iSearchQueryID), "", "")

	Response.Clear
	Response.Redirect sScriptFileName & sTempParams

Else
	Response.Flush

	iProjectID=CheckIntegerAndZero(Request.QueryString("idproject"))
	sParams=ReplaceUrlParams(sParams, "idproject=" & iProjectID)
	sParams=ReplaceUrlParams(sParams, "qid=" & iSearchQueryID)

	sSearchAction=Request.QueryString("act")
	iSearchQueryID=Request.QueryString("qid")
	sSearchKeywordsAdd=Request.QueryString("srch_queryadd")
	sSearchOrderBy=Request.QueryString("srch_orderby")

	Set objTempRs=GetDataRecordsetSP("usp_Ica_MemberExpertMethoSearchRepeatSelect", Array( _
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

If Not InStr(sParams, "qid=")>1 Then sParams=AddUrlParams(sParams, "qid=" & iSearchQueryID)
%>
<!--#include virtual="/_template/html.header.start.asp"-->

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

function DownloadCV(cv_num, euid, lng, db) {
	var f = document.forms[0];
	var obj_cv = f["cvformat" + cv_num];
	var frm = obj_cv.options[obj_cv.selectedIndex].value;
	euid = (f["uid" + cv_num] && f["uid" + cv_num].value) || euid;
	var url = "../view/cv_view" + frm + ".asp<% =ReplaceUrlParams(ReplaceUrlParams(sParams, "uid"), "t=1") %>&uid=" + euid;
	// To avoid a bug with IE8
	euid = euid.substring(1, 9);
	var wnd = window.open(url, "cv" + euid, "scrollbars=yes, status=yes, location=yes, menubar=yes, toolbar=yes, resizable=yes");
	wnd.focus();
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

<% If bUserAccessMethodology Then %>
function editMethodology(expertUid, databaseId, expertId) {
	var url = "exp_methodology.asp<% =ReplaceUrlParams(ReplaceUrlParams(sParams, "uid"), "t=1") %>&uid=" + expertUid;
	var euid = expertUid.substring(1, 9);
	var wnd = window.open(url, "metho" + euid, "width=800, height=655, left=400, top=100, scrollbars=no, status=no, location=no, menubar=no, toolbar=no, resizable=yes");
	wnd.focus();
}
<% End If %>
</script>
</head>
<body>
	<!-- header -->
	<!--#include file="../../_template/page.header.asp"-->

	<!-- content -->
	<div id="content" class="workscreen">
		<div id="hdrUpdatedList" class="colCCCCCC uprCse f17 spc01 botMrgn10">Search results</div>

<% If objTempRs.Eof Then %>
<!-- [i] No Results -->
	<br />
	<div class="col666666 botMrgn10">
		There is no expert matching your search query.<br />
		Please revise your query and try again. If you have questions please contact us at <a href="mailto:info@icaworld.net">info@icaworld.net</a>.
	</div>
	<br />
	<div class="alignCenter"><a href="exp_search.asp" class="red-button w95">New Search</a></div>

<% Else 
	If IsNull(objTempRs(0)) Then
		%><br /><br />
		<div class="col666666 botMrgn10">
			The searh query is expired. Please revise your query and try again. <br />If you have questions please contact us at <a href="mailto:info@icaworld.net">info@icaworld.net</a>.
		</div><% 
		Response.End
	End If
	iNumberCVsTotalInResult=objTempRs.RecordCount %>

   <!-- [i] Results overview -->
	<div class="col666666 botMrgn10">
		<a href="exp_search.asp<% =sParams %>" class="red-button floatRight w95" style="margin-top:-32px;">Search Again</a>
		<% If iNumberCVsTotalInResult=1 Then %>There is 1 expert<% Else %>There are <%=iNumberCVsTotalInResult%> experts<% End If %> matching your search query.
	</div>

	<%
	objTempRs.PageSize=10
	iTotalRecords=objTempRs.RecordCount
	iTotalPages=objTempRs.PageCount

	ShowNavigationPages iCurrentPage, iTotalPages, sParams
%>

<form name="exp_list_form" id="exp_list_form" method="post">
<input type="hidden" name="qid" value="<%=iSearchQueryID%>"><input type="hidden" name="srch_fullquery" value='<%=sSearchFullText%>'>

<% sExpertIds=ReplaceIfEmpty(sExpertIds, "0") %>
<input type=hidden name="ExpertIds" value="<%=sExpertIds%>">

<% ShowIcaExpertsBlock 1, iCurrentPage %>

</form>

<% ShowNavigationPages iCurrentPage, iTotalPages, sParams %>
<% End If %>

	</div>
	<div id="rightspace">
   <!-- Feature boxes -->

	<form name="search" method="get" action="<%=sScriptFileName%>">
	<input type="hidden" name="qid" value="<%=iSearchQueryID%>">
	<% ShowFeatureBoxHeader "Refine&nbsp;your&nbsp;search" %>
	<div class="content">
	<p class="sml" align="center">Order these results by</p>
	<div align="center" style="margin-bottom:10px;"><select style="font-face: Arial; font-size:8.5pt;" name="srch_orderby" size=1>
	<option value="expRank"<% If sSearchOrderBy="expRank" Then %> selected<% End If %>>Relevance &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;</option>
	<option value="expProfession"<% If sSearchOrderBy="expProfession" Then %> selected<% End If %>>Profession</option>
	<option value="expSeniority"<% If sSearchOrderBy="expSeniority" Then %> selected<% End If %>>Seniority</option>
	</select></div>
	<div align="center" style="margin-bottom:5px;"><a href="javascript:void(0)" onclick="$(this).closest('form').submit()" class="red-button">Order</a></div>
	<% ShowFeatureBoxDelimiter %>
	<p class="sml" align="center">Search within results</p>
	<input type="hidden" name="txt" value="<% =Request.QueryString("txt") %>">
	<div align="center"><input type="text" style="font-face: Arial; font-size:8.5pt;" name="srch_queryadd" size=16 value="<% =sSearchKeywordsAdd %>"></div>
	<p class="sml" align="center">(work experience only)</p>
	<div align="center"><a href="javascript:void(0)" onclick="$(this).closest('form').submit()" class="red-button">Search</a></div>
	</div>
	<% ShowFeatureBoxFooter %>
	</form>

	<% If iUserID = 135 Then
	'	If iNumberCVsTotalInResult > 0 Then
	'		objTempRs.MoveFirst
	'		While Not objTempRs.Eof
	'			Set objExpertDB = objExpertDBList.Find(objTempRs("DB"), "Database")
	'			Response.Write objExpertDB.DatabaseCode & objTempRs("id_Expert") & "<br>"
	'			objTempRs.MoveNext
	'		Wend
	'	End If
	End If %>

	</div>
<% objTempRs.Close %>

	<!-- footer -->
	<!--#include file="../../_template/page.footer.asp"-->
<% CloseDBConnection %>
</body>
<!--#include file="../../_template/html.footer.asp"-->

