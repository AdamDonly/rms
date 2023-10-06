<%@ LANGUAGE="VBSCRIPT" %>
<% 'Option Explicit
'--------------------------------------------------------------------
'
'--------------------------------------------------------------------
%>
<!--#include virtual = "/dbc.asp"-->
<!--#include virtual = "/fnc.asp"-->
<!--#include virtual = "/fnc_exp.asp"-->
<!--#include file="cv_data.asp"-->
<!--#include virtual = "/_common/_data/en/lib.asp"-->
<!--#include virtual = "/_common/_data/datCountry.asp"-->
<!--#include virtual = "/_common/_data/datProcurementType.asp"-->
<!--#include virtual = "/_forms/frmInterface.asp"-->
<!--#include virtual = "/_forms/frmScrollBoxNew.asp"-->
<%
' Check user's access rights
CheckUserLogin sScriptFullNameAsParams
%>
<!--#include file="cv_circle_data.asp"-->
<!--#include file="cv_circle_save.asp"-->

<!--#include virtual="/_template/html.header.scrolllist.start2.asp"-->

<% InsertJsHelpers 0, 0 %>

<%
Dim iTopExpertStatus
iTopExpertStatus = GetExpertCompanyTopExpertByUid(sCvUID, iUserCompanyID, iUserID)
%>
<script language="JavaScript" type="text/javascript">
function Continue() {
	var f = document.RegForm;

	f.mmb_cou_hid.value = '0,' + (GetScrollboxSelection("divCouSelector", "cou_") || '0');
	f.mmb_don_hid.value = '0,' + (GetScrollboxSelection("divDonSelector", "don_") || '0');
	f.mmb_sct_hid.value = '0,' + (GetScrollboxSelection("divSctSelector", "sct_") || '0');
	<%
	If iTopExpertStatus = 1 Then
		%>
		if ($('input[type="checkbox"][id^="sct_"]:checked').length > 7)
		{
			alert("The selected sectors should not be more than 7 for Top Experts.");
			return false;
		}
		<%
	End If
	%>
	// at least one criteria should be filled in
	if ((f.mmb_cou_hid.value.replace(/,0/gi,'')=='0') 
		&& (f.mmb_don_hid.value.replace(/,0/gi,'')=='0')
		&& (f.mmb_sct_hid.value.replace(/,0/gi,'')=='0')
		) {	
			alert('Please specify sectors, countries and funding agancies of interest for this expert.');
			return false;
		}
	
	f.action = "cv_circle_fields2.asp";
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
	<!--#include virtual = "/_template/page.header.asp"-->

	<!-- content -->
	<div id="content" class="searchform">
	
	<% 
	If Not bIsMyCV Then 
		%><div class="colCCCCCC uprCse f17 spc01 botMrgn10"><%
		If iTopExpertStatus = 1 Then
			%><span class="service_title">Top Expert</span><%
		Else
			%><span class="service_title"><% If Not bIsCompanyCircleExpert Then %>Add to <% End If %>My Experts Circle</span>
			<%
		End If
		%><br/><% =sFullNameWithSpaces %>
		<br/>Expert ID: <% =objExpertDB.DatabaseCode %><%=iCvID%>
		</div>

		<p>Please select sectors and countries for informing this expert about projects and vacancies<br>
		or <a href="<% = sScriptFileName & ReplaceUrlParams(sParams, "act=fromcv") %>">reload the selection of sectors and countries from the expert's CV</a></p>
	<%
	End If
	%>

<!-- Keywords search -->
	<form action="cv_circle_fields2.asp" method="post" name="RegForm" id="RegForm" onSubmit="Continue(); return false;">


		<div class="mini_wraper">
			<div id="divSctSelector" class="filter_table_flex">
				<div class="filter_table_header"><h3>Procurement types</h3></div>
				<div style="padding:10px;">
					<%
					Dim bIsProcTypeSelected
					For i = LBound(arrProcurementTypeID) To UBound(arrProcurementTypeID)
						If arrProcurementTypeName(i) > "" Then
							bIsProcTypeSelected = 0
							If mPT > 0 Then
								For j = 0 To mPT
									If mProcTypeIDs(j) = arrProcurementTypeID(i) Then
										bIsProcTypeSelected = 1
									End If
								Next
							End If
							%><div style="display:inline-block;width:22%;"><input type="checkbox" id="proctype_<%=arrProcurementTypeID(i) %>" name="mmb_proctypes" value="<%=arrProcurementTypeID(i) %>" <% If arrProcurementTypeID(i) = 1 Or bIsProcTypeSelected = 1 Then %>checked="checked"<% End If %>> <label style="vertical-align:baseline;" for="proctype_<%=arrProcurementTypeID(i) %>"><%=arrProcurementTypeName(i) %></label></div>
							<%
						End If
					Next %>
				</div>
			</div><br />
			<%
			' Sectors
			ShowSctScrollBox "Sectors of interest",  "", 1, 0, 0 %><br />
			<%
			' Countries
			ShowCouScrollBox "Countries of interest",  "", 1, 1, 1, 1, 0 %><br />
			<%			
			'Funding agancies - no need to select them, as by default all funding agencies are selected
			' ShowDonScrollBox "Funding agencies of interest",  "", 1, 1, 0, 0 %><br />

		</div>

		<input type="submit" class="red-button under-right-col-25perc" name="btnSubmit" value="Submit" />
		<a href="<%=sScriptFilelName & AddUrlParams(sParams, "act=clear") %>"class="red-button next-btn">Clear all</a>
		<br/><br/>
		
	<input type="hidden" name="mmb_cou_hid">
	<input type="hidden" name="mmb_don_hid">
	<input type="hidden" name="mmb_sct_hid">
	<input type="hidden" name="uid" value="<% =sCvUID %>">
	</form>
	</div>
	<!-- footer -->
	<!--#include virtual="/_template/page.footer.asp"-->
<script language=JavaScript type=text/javascript>
<% If sAction = "clear" Then %>
	document.RegForm.mmb_cou_hid.value='';
	document.RegForm.mmb_don_hid.value='';
 	document.RegForm.mmb_sct_hid.value='';
<% Else %>
	document.RegForm.mmb_cou_hid.value='<%=mNt & mNtInt %>';
	document.RegForm.mmb_don_hid.value='<%=mOrg & mOrgInt %>';
 	document.RegForm.mmb_sct_hid.value='<%=mExT & mExTInt %>';
<% End If %>
</script>
<% CloseDBConnection %>

<% WriteFilterTableScript %>
<link href="/res/css/jquery.mCustomScrollbar.css" rel="stylesheet" type="text/css" />
<script src="/res/scripts/jquery.mCustomScrollbar.concat.min.js"></script>
</body>
<!--#include virtual="/_template/html.footer.asp"-->
