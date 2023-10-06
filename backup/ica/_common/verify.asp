<%
'--------------------------------------------------------------------
'
' CV registration.
' Short format.
'
'--------------------------------------------------------------------
Response.Expires = 0
Response.ExpiresAbsolute = Now() - 1
Response.CacheControl = "no-cache" 
%>
<!--#include file="_data/datGender.asp"-->
<!--#include file="_data/datPsnTitle.asp"-->
<!--#include file="_data/datMonth.asp"-->
<%
If sApplicationName<>"expert" Then
	CheckUserLogin sScriptFullNameAsParams
End If
%>
<!--#include file="../_common/expProfile.asp"-->
<!--#include virtual="/_template/html.header.start.asp"-->
<script type="text/javascript" src="/_scripts/js/main.js"></script>
<script type="text/javascript">
<!-- 
function validateForm() {
	var f=document.forms[0];
	if (!(f)) {
		return false; }
	if (!checkTextFieldValue(f.exp_firstname, "", "Please fill in expert's first name.", 1)) { return false; }
	if (!checkTextFieldValue(f.exp_familyname, "", "Please fill in expert's family name.", 1)) { return false; }
	if (!checkTextFieldValue(f.exp_email, "", "Please fill in expert email.", 1)) { return false; }
	if ((f.exp_email.value) && (!validateEmail(f.exp_email.value))) {
		alert("Please retype e-mail address correctly");
        f.exp_email.select();        
		return false;
	}
	if (!checkTextFieldValue(f.cvEng, "", "Please click Browse and select file with expert's CV to attach.", 1)) { return false; }
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

<div class="colCCCCCC uprCse f17 spc01 botMrgn10">REGISTER EXPERT IN <% =objUserCompanyDB.DatabaseTitle %> DATABASE</div>
<div class="botMrgn5 col666666 itlc botMrgn10">
Check if a CV already exists in ICA Members’ Databases. <br>All fields on the below form need to be filled in and expert’s CV must be attached.
</div>

<div class="col333399 itlc botMrgn15">
<p>If your CV is not found in other ICA Members’ Databases then you will be taken to the full CV registration page.<br>
If your CV is already in other ICA Members’ Databases then you will only need to update your CV <br>in the second step to add also it to your organisation database.</p>
</div>

	
		<form enctype="multipart/form-data" action="verify_results.asp<% =AddUrlParams(sParams, "act=" & sAction) %>"  method="post" onsubmit="validateForm(); return false;">
		<div class="box search blue">
		<h3>Personal information &amp; Contact details</h3>
		<table class="search_form" width="100%" cellspacing=0 cellpadding=0 border=0>
<% If bCvMultiLanguageActive = cMultiLanguageEnabled Then %>
		<tr class="first">
		<td class="field splitter"><label for="exp_language"><% = GetLabel(sCvLanguage, "CV language") %></label></td>
		<td class="value blue"><select name="exp_language" size="1" onChange="submitForm();" style="width:130px;">
		<%
		Dim sTempLanguage
		For Each sTempLanguage in dictLanguage
			Response.Write "<option value=""" & sTempLanguage & """" 
			If sCvLanguage=sTempLanguage Then Response.Write " selected"
			Response.Write ">" & dictLanguage.Item(sTempLanguage) & "</option>"
		Next
		%>
		</select></td>
		</tr>
		</table>

		<table class="search_form" width="100%" cellspacing=0 cellpadding=0 border=0>
<% End If %>
		<tr class="first">
		<td class="field splitter"><label for="exp_firstname">First&nbsp;name</label></td>
		<td class="value blue"><input type="text" id="exp_firstname" name="exp_firstname" size=31 style="width: 300px;" maxlength=255 value="">&nbsp;<font size=3 color="#CC0000">*</font></td>
		</tr>
		<tr>
		<td class="field splitter"><label for="exp_familyname">Family&nbsp;name</label></td>
		<td class="value blue"><input type="text" id="exp_familyname" name="exp_familyname" size=31 style="width: 300px;" maxlength=250 value="">&nbsp;<font size=3 color="#CC0000">*</font></td>
		</tr>
		<tr>
		<td class="field splitter"><label for="exp_email">Primary email</label></td>
		<td class="value blue"><input type="text" id="exp_email" name="exp_email" size=31 style="width: 300px;" maxlength=50 value="">&nbsp;<font size=3 color="#CC0000">*</font></td>
		</tr>
		<tr class="last">
		<td class="field splitter"><label for="cvEng">Attach CV</label></td>
		<td class="value blue"><input type="file" name="cvEng" accept="/image." size="24"></td>
		</tr>
		</table>
		</div>
		
		<div class="spacebottom">
			<input type="submit" id="btnSubmit" name="btnSubmit" class="red-button under-right-col" value="Submit" />
			<a href="/" class="red-button next-btn">Cancel</a>
		</div>
		</form>
	</div>
	<!-- footer -->
	<!--#include virtual="/_template/page.footer.asp"-->
<% CloseDBConnection %>
</body>
<!--#include virtual="/_template/html.footer.asp"-->
