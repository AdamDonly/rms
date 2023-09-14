<%@ LANGUAGE="VBSCRIPT" %>
<% Option Explicit %>

<!--#include file="../../dbc.asp"-->
<!--#include file="../../fnc.asp"-->
<!--#include file="../../_forms/frmInterface.asp"-->
<!--#include file="../../fnc_exp.asp"-->

<%
'--------------------------------------------------------------------
'
' Top Expert documents upload.
'
'--------------------------------------------------------------------
Response.Expires = 0
Response.ExpiresAbsolute = Now() - 1
Response.CacheControl = "no-cache" 

If sApplicationName<>"expert" Then
	CheckUserLogin sScriptFullNameAsParams
End If
%>
<!--#include file="../../_common/cv_data.asp"-->
<!--#include virtual="/_template/html.header.start.asp"-->
<script type="text/javascript" src="/_scripts/js/main.js"></script>
<script type="text/javascript">
<!--
function validateForm() {
	var f = document.forms[0];
	if (!(f)) {
		return false;
	}
	if (!checkTextFieldValue(f.exp_docTitle, "", "Please fill in document title.", 1)) { return false }

	if (!checkTextFieldValue(f.cvEng, "", "Please click Browse and select a file to upload.", 1)) { return false }
	f.submit();
}
-->
</script>
</head>
<body>
<div id="holder">
	<!-- header -->
	<!--#include virtual="/_template/page.header.asp"-->

	<!-- content -->
	<div id="content" class="workscreen">

		<h2 class="service_title">Upload your documents. <span class="service_slogan">Add your certificates and diplomas.</span></h2>

		<% ShowMessageStart "info", 450 %>
		
		<% ShowMessageEnd %><br/>

		<form enctype="multipart/form-data" action="docsave.asp"  method="post" onSubmit="validateForm(); return false;">
		<div class="box search blue">
		<h3><span class="left">&nbsp;</span><span class="right">&nbsp;</span>New document</h3>
		<table class="search_form" width="100%" cellspacing=0 cellpadding=0 border=0>
		<tr>
		<td class="field splitter"><label for="exp_firstname">Document title</label></td>
		<td class="value blue"><input type="text" id="exp_docTitle" name="exp_docTitle" size="31" style="width: 300px;" maxlength="255" />&nbsp;<font size="3" color="#CC0000">*</font></td>
		</tr>
		<tr class="last">
		<td class="field splitter"><label for="cvEng">Choose file</label></td>
		<td class="value blue"><input type="file" name="cvEng" accept="/image." size="24"/></td>
		</tr>
		</table>
		</div>
		
		<div class="spacebottom">
		<input type="image" class="button first" src="/image/bte_submit.gif" name="btnSubmit" id="btnSubmit" alt="Submit">
		<a href="/"><img class="button" src="/image/bte_cancel.gif" border=0 alt="Cancel"></a>
		</div>
		</form>
	</div>
	<div id="rightspace">
		<!-- feature boxes -->
		<%
		If bIsMyCV Then
			ShowTopExpCVFeatureBox
	'	Else
	'		ShowExpCVFeatureBox
		End If
		%>	
	</div>
</div>
	<!-- footer -->
	<!--#include virtual="/_template/page.footer.asp"-->
<% CloseDBConnection %>
</body>
<!--#include virtual="/_template/html.footer.asp"-->

