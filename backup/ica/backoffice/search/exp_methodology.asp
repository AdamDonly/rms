<!--#include file="../../dbc.asp"-->
<!--#include file="../../fnc.asp"-->
<!--#include file="../../fnc_exp.asp"-->
<!--#include virtual = "/_common/_data/datFlagMetho.asp"-->

<!--#include virtual="/_common/cv_data.asp"-->
<!--#include virtual="/_common/_class/main.asp"-->
<% 
If Not bUserAccessModifyMethodology Then
	Response.Redirect "/"
	Response.End
End If
If bCvValidForMemberOrExpert = aClientSecurityCvViewEnabled Or bCvValidForMemberOrExpert = aClientSecurityCvViewAll Then
Else
	Response.Redirect "/"
	Response.End
End If

Dim	iExpertMethoSent, _
	dExpertMethoSentDate, _
	dExpertMethoAnswerDate, _
	bExpertMethoShowAll, _
	bExpertMethoTA, _
	bExpertMethoFWC, _
	iExpertMethoCountType, _
	iExpertMethoCountTa, _
	iExpertMethoCountFwc, _
	iExpertMethoCountGrant, _
	iExpertMethoContribRev, _
	iExpertMethoContribTech, _
	iExpertMethoContribFull, _
	iExpertMethoContribRead, _
	bExpertMethoEN, _
	bExpertMethoFR, _
	bExpertMethoSP, _
	bExpertMethoPT, _
	bExpertMethoDE, _
	bExpertMethoProofread, _
	sExpertMethoComments, _
	sExpertMethoExpertComments, _
	sExpertMethoKeywords, _
	sExpertMethoExpertKeywords, _
	sExpertMethoDonors, _
	sExpertMethoSelectedFlag, _
	bUpdated

bUpdated = False
If Request.Form() > "" Then
	sExpertMethoSelectedFlag = CheckIntegerAndNull(Request.Form("flag"))

	iExpertMethoCountTa = CheckIntegerAndZero(Request.Form("count_ta"))
	iExpertMethoCountFwc = CheckIntegerAndZero(Request.Form("count_fwc"))
	iExpertMethoCountGrant = CheckIntegerAndZero(Request.Form("count_grant"))
	iExpertMethoContribRev = CheckboxOnOrZero(Request.Form("contrib_rev"))
	iExpertMethoContribTech = CheckboxOnOrZero(Request.Form("contrib_tech"))
	iExpertMethoContribFull = CheckboxOnOrZero(Request.Form("contrib_full"))
	iExpertMethoContribRead = CheckboxOnOrZero(Request.Form("contrib_read"))

	bExpertMethoShowAll = CheckboxOnOrZero(Request.Form("show_all"))
	bExpertMethoTA = CheckboxOnOrZero(Request.Form("project_ta"))
	bExpertMethoFWC = CheckboxOnOrZero(Request.Form("project_fwc"))
	bExpertMethoEN = CheckboxOnOrZero(Request.Form("language_en"))
	bExpertMethoFR = CheckboxOnOrZero(Request.Form("language_fr"))
	bExpertMethoSP = CheckboxOnOrZero(Request.Form("language_sp"))
	bExpertMethoPT = CheckboxOnOrZero(Request.Form("language_pt"))
	bExpertMethoDE = CheckboxOnOrZero(Request.Form("language_de"))
	bExpertMethoProofread = CheckIntegerAndNull(Request.Form("proofread"))
	
	sExpertMethoComments = CheckStringLength(HtmlEncode(Request.Form("comments")), 10000)
	sExpertMethoKeywords = CheckStringLength(HtmlEncode(Request.Form("keywords")), 10000)
	sExpertMethoDonors = CheckStringLength(HtmlEncode(Request.Form("donors")), 10000)

'	UpdateRecordSP "usp_Ica_ExpertMethodologyFlagUpdate", Array( _
'		Array(, adInteger, , objExpertDB.ID), _
'		Array(, adInteger, , iCvID), _
'		Array(, adInteger, , sExpertMethoSelectedFlag), _
'		Array(, adVarWChar, 10000, sExpertMethoComments))

	UpdateRecordSP "usp_Ica_ExpertMethodologyFlagFullUpdate2", Array( _
		Array(, adInteger, , objExpertDB.ID), _
		Array(, adInteger, , iCvID), _
		Array(, adInteger, , sExpertMethoSelectedFlag), _
		Array(, adInteger, , bExpertMethoShowAll), _
		Array(, adInteger, , iExpertMethoCountTa), _
		Array(, adInteger, , iExpertMethoCountFwc), _
		Array(, adInteger, , iExpertMethoCountGrant), _
		Array(, adInteger, , iExpertMethoContribRev), _
		Array(, adInteger, , iExpertMethoContribTech), _
		Array(, adInteger, , iExpertMethoContribFull), _
		Array(, adInteger, , iExpertMethoContribRead), _
		Array(, adInteger, , bExpertMethoEN), _
		Array(, adInteger, , bExpertMethoFR), _
		Array(, adInteger, , bExpertMethoSP), _
		Array(, adInteger, , bExpertMethoPT), _
		Array(, adInteger, , bExpertMethoDE), _
		Array(, adVarWChar, 10000, sExpertMethoComments), _
		Array(, adVarWChar, 10000, sExpertMethoKeywords), _
		Array(, adVarWChar, 10000, sExpertMethoDonors) _
		)

	bUpdated = True
	%>
	<!--#include virtual = "/_template/page.close.asp"-->
<%
Else

	On Error Resume Next
	Set objTempRs = GetDataRecordsetSP("usp_Ica_ExpertMethodologySelect", Array( _
		Array(, adInteger, , objExpertDB.ID), _
		Array(, adInteger, , iCvID)))
	If Not objTempRs.Eof Then
		iExpertMethoCountType = objTempRs("expMethoCountType")
		iExpertMethoCountTa = objTempRs("expMethoCountTa")
		iExpertMethoCountFwc = objTempRs("expMethoCountFwc")
		iExpertMethoCountGrant = objTempRs("expMethoCountGrant")
		iExpertMethoContribRev = objTempRs("expMethoContribRev")
		iExpertMethoContribTech = objTempRs("expMethoContribTech")
		iExpertMethoContribFull = objTempRs("expMethoContribFull")
		iExpertMethoContribRead = objTempRs("expMethoContribRead")
		bExpertMethoShowAll = objTempRs("expMethoShowAll")
		bExpertMethoTA = objTempRs("expMethoTA")
		bExpertMethoFWC = objTempRs("expMethoFWC")
		bExpertMethoEN = objTempRs("expMethoEN")
		bExpertMethoFR = objTempRs("expMethoFR")
		bExpertMethoSP = objTempRs("expMethoSP")
		bExpertMethoPT = objTempRs("expMethoPT")
		bExpertMethoDE = objTempRs("expMethoDE")
		bExpertMethoProofread = objTempRs("expProofread")
		sExpertMethoExpertComments = objTempRs("expMethoExpComments")
		sExpertMethoComments = objTempRs("expMethoComments")

		sExpertMethoExpertKeywords = objTempRs("expMethoExpKeywords")
		sExpertMethoKeywords = objTempRs("expMethoKeywords")
		sExpertMethoDonors = objTempRs("expMethoDonors")

		sExpertMethoSelectedFlag = objTempRs("expMethoSelectedFlag")
	End If
	objTempRs.Close
	Set objTempRs = Nothing
	On Error GoTo 0
End If
%>

<!--#include virtual="/_template/html.header.asp"-->
<body>
<div>

	<!-- content -->
	<h2 class="service_title" style="margin: 8px 0; "><% =sFullName %> <span class="service_slogan">Expert ID: <% =objExpertDB.DatabaseCode %><%=iCvID%></span></h2>

<% If bUpdated = False Then %>

	<div style="margin: 5px 20px; text-align: left;">
	<form action="<%=AddUrlParams(sParams, "act=" & sAction) %>"  method="post" id="exp_form" name="exp_form">

	<table width="100%" border="0">
	<tr style="height: 30px;">
		<td width="20%"><p><label class="inline" for="flag">Flag</label></p></td>
		<td width="40%">
		<div class="value">
			<select name="flag" id="flag" style="width: 120px;">
			<option></option>
			<% ShowFlagMethoSelectItems sExpertMethoSelectedFlag, "id", "" %>
			</select>
		</div>
		<td width="40%">
		<div class="value">
			<input type="checkbox" name="show_all" id="show_all" value="1" <% If bExpertMethoShowAll = 1 Then %>checked<% End If %>> 
			<label class="inline" for="show_all"><span style="font-size: 8pt; padding-top: 5px;">Show full profile on a preview / search results</span></label> &nbsp; 
		</div>
		</td>
	</tr>

	<tr style="height: 25px;">
		<td width="20%"><p>Methodologies #</p></td>
		<td width="40%"></td>
		<td width="40%"><p>Documents</p></td>
	</tr>
	<tr style="height: 25px;">
		<td width="20%"><p><label class="inline" for="count_ta" style="color: #aaa;">&nbsp; &nbsp; All (old)</label></p></td>
		<td>
		<div class="value">
		<select name="count_type" id="count_type" style="width: 280px;" disabled readonly>
			<option value="0"></option>
			<option value="1" <% If iExpertMethoCountType = 1 Then %>selected<% End If %>>a) between 1 and 4</option>
			<option value="2" <% If iExpertMethoCountType = 2 Then %>selected<% End If %>>b) between 5 and 8</option>
			<option value="3" <% If iExpertMethoCountType = 3 Then %>selected<% End If %>>c) more than 8</option>
		</select>
		</div>
		</td>
		<td rowspan="6" style="vertical-align: top; padding-top: 5px;">
			<%
			Dim objDocumentList
			Set objDocumentList = New CDocumentList
			objDocumentList.LoadDocumentListByExpertID iCvID, "", aDocumentTypeIdMethodologySupport
			
			If objDocumentList.Count = 0 Then
			%>
				<p class="sml" style="padding: 2px 5px;">There are no documents uploaded yet.</p>
			<% 
			Else
				ShowDocumentListEditTable objDocumentList
			End If
			Set objDocumentList = Nothing
			%>
			<div align="center"><a href="/backoffice/search/doc_methodology.asp<% =ReplaceUrlParams(sParams, "document=0") %>" class="red-button w125">Upload document</a></div>

		</td>
	</tr>

	<tr style="height: 25px;">
		<td><p><label class="inline" for="count_ta">&nbsp; &nbsp; TA</label></p></td>
		<td>
		<div class="value">
		<select name="count_ta" id="count_ta" style="width: 280px;">
			<option value="0"></option>
			<option value="1" <% If iExpertMethoCountTa = 1 Then %>selected<% End If %>>a) between 1 and 4</option>
			<option value="2" <% If iExpertMethoCountTa = 2 Then %>selected<% End If %>>b) between 5 and 8</option>
			<option value="3" <% If iExpertMethoCountTa = 3 Then %>selected<% End If %>>c) more than 8</option>
		</select>
		</div>
		</td>
	</tr>

	<tr style="height: 25px;">
		<td><p><label class="inline" for="count_fwc">&nbsp; &nbsp; FWC</label></p></td>
		<td>
		<div class="value">
		<select name="count_fwc" id="count_fwc" style="width: 280px;">
			<option value="0"></option>
			<option value="1" <% If iExpertMethoCountFwc = 1 Then %>selected<% End If %>>a) between 1 and 4</option>
			<option value="2" <% If iExpertMethoCountFwc = 2 Then %>selected<% End If %>>b) between 5 and 8</option>
			<option value="3" <% If iExpertMethoCountFwc = 3 Then %>selected<% End If %>>c) more than 8</option>
		</select>
		</div>
		</td>
	</tr>

	<tr style="height: 25px;">
		<td><p><label class="inline" for="count_grant">&nbsp; &nbsp; Grants</label></p></td>
		<td>
		<div class="value">
		<select name="count_grant" id="count_grant" style="width: 280px;">
			<option value="0"></option>
			<option value="1" <% If iExpertMethoCountGrant = 1 Then %>selected<% End If %>>a) between 1 and 4</option>
			<option value="2" <% If iExpertMethoCountGrant = 2 Then %>selected<% End If %>>b) between 5 and 8</option>
			<option value="3" <% If iExpertMethoCountGrant = 3 Then %>selected<% End If %>>c) more than 8</option>
		</select>
		</div>
		</td>
	</tr>

	<tr>
		<td style="vertical-align: top; padding-top: 5px;"><p><label class="inline" for="count_type">Contribution</label></p></td>
		<td style="padding-top: 5px;">
		<div class="value">
			<p style="margin-bottom: 5px;">
				<input type="checkbox" name="contrib_rev" id="contrib_rev" value="1" <% If iExpertMethoContribRev = 1 Then %>checked<% End If %>> <label class="inline" for="contrib_rev">Reviewing others' contributions</label> &nbsp; &nbsp;
			</p>
			<p style="margin-bottom: 5px;">
				<input type="checkbox" name="contrib_tech" id="contrib_tech" value="1" <% If iExpertMethoContribTech = 1 Then %>checked<% End If %>> <label class="inline" for="contrib_tech">Contributing with technical inputs</label> &nbsp; &nbsp; <br>
			</p>
			<p style="margin-bottom: 5px;">
				<input type="checkbox" name="contrib_full" id="contrib_full" value="1" <% If iExpertMethoContribFull = 1 Then %>checked<% End If %>> <label class="inline" for="contrib_full">Writing methodologies in full</label> &nbsp; &nbsp; <br>
			</p>
			<p style="margin-bottom: 5px;">
				<input type="checkbox" name="contrib_read" id="contrib_read" value="1" <% If iExpertMethoContribRead = 1 Then %>checked<% End If %>> <label class="inline" for="contrib_read">Proofreading and editing</label> &nbsp; &nbsp; 
			</p>
		</div>
		</td>
	</tr>

	<% 
	If bExpertMethoTA = 1 _
	Or bExpertMethoFwc = 1 _
	Then %>
	<tr style="height: 25px;">
		<td><p><label class="inline" for="project_ta" style="color: #aaa;">Type of projects (old)</label></p></td>
		<td>
		<div class="value">
			<p>
				<input type="checkbox" disabled readOnly name="project_ta" id="project_ta" value="ta" <% If bExpertMethoTA = 1 Then %>checked<% End If %>> <label class="inline" for="project_ta">Technical Assistance</label> &nbsp; &nbsp; 
				<input type="checkbox" disabled readOnly name="project_fwc" id="project_fwc" value="fwc" <% If bExpertMethoFWC = 1 Then %>checked<% End If %>> <label class="inline" for="project_fwc">Framework Contracts</label> &nbsp; &nbsp;
			</p>
		</div>
		</td>
	</tr>
	<%
	End If 
	%>

	<tr style="height: 25px;">
		<td><p><label class="inline" for="count_type">Languages</label></p></td>
		<td colspan="2">
		<div class="value">
			<p>
				<input type="checkbox" name="language_en" id="language_en" value="en" <% If bExpertMethoEN = 1 Then %>checked<% End If %>> <label class="inline" for="language_en">English</label> &nbsp; &nbsp; 
				<input type="checkbox" name="language_fr" id="language_fr" value="fr" <% If bExpertMethoFR = 1 Then %>checked<% End If %>> <label class="inline" for="language_fr">French</label> &nbsp; &nbsp; 
				<input type="checkbox" name="language_sp" id="language_sp" value="sp" <% If bExpertMethoSP = 1 Then %>checked<% End If %>> <label class="inline" for="language_sp">Spanish</label> &nbsp; &nbsp; 
				<input type="checkbox" name="language_pt" id="language_pt" value="pt" <% If bExpertMethoPT = 1 Then %>checked<% End If %>> <label class="inline" for="language_pt">Portuguese</label> &nbsp; &nbsp; 
				<input type="checkbox" name="language_de" id="language_de" value="de" <% If bExpertMethoDE = 1 Then %>checked<% End If %>> <label class="inline" for="language_de">German</label> &nbsp; &nbsp; 
			</p>
		</div>
		</td>
	</tr>

	<tr>
		<td valign="top"><p><label class="inline" for="comments">Internal comments</label><br><small>only &lt;strong&gt; and &lt;br&gt; html tags are allowed</small></p></td>
		<td colspan="2">
		<div class="value">
			<textarea class="inputarea" name="comments" id="comments" cols="31" style="width: 90%; height: 72px" rows="3" wrap="yes"><% =sExpertMethoComments %></textarea>
		</div>
		</td>
	</tr>
	<tr style="height: 25px;">
		<td valign="top"><p><label class="inline" for="">Expert's comments</label></p></td>
		<td colspan="2">
		<div class="value">
			<% =Server.HtmlEncode(ReplaceIfEmpty(sExpertMethoExpertComments, "-")) %>
		</div>
		</td>
	</tr>

	<tr>
		<td valign="top"><p><label class="inline" for="keywords">Writing experience</label><br><small>only &lt;strong&gt; and &lt;br&gt; html tags are allowed</small></p></td>
		<td colspan="2">
		<div class="value">
			<textarea class="inputarea" name="keywords" id="keywords" cols="31" style="width: 90%; height: 72px" rows="3" wrap="yes"><% =sExpertMethoKeywords %></textarea>
		</div>
		</td>
	</tr>
	<tr style="height: 25px;">
		<td valign="top"><p><label class="inline" for="">Expert is comfortable to write on</label></p></td>
		<td colspan="2">
		<div class="value">
			<% =Server.HtmlEncode(ReplaceIfEmpty(sExpertMethoExpertKeywords, "-")) %>
		</div>
		</td>
	</tr>

	<tr>
		<td valign="top"><p><label class="inline" for="donors">Other donors</label></p></td>
		<td colspan="2">
		<div class="value">
			<textarea class="inputarea" name="donors" id="donors" cols="31" style="width: 90%; height: 72px" rows="3" wrap="yes"><% =sExpertMethoDonors %></textarea>
		</div>
		</td>
	</tr>

	<tr>
		<td>&nbsp;</td>
		<td colspan="2">
		<div class="spacetop20 spacebottom">
		<div class="medium primary btn">
			<input type="submit" value="Save">
		</div>
		</div>
	</tr>
	</table>

	</form>
	<% 
End If
%>

</div>
<% CloseDBConnection %>
</body>
<!--#include virtual="/_template/html.footer.asp"-->
