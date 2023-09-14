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
CheckExpertID

Dim sUserPhone
Dim sFlagSelected
%>
<!--#include file="../_common/expProfile.asp"-->
<% LoadExpertProfile(iExpertID) %>
<html>
<head>
<title>Quick CV registration</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" type="text/css" href="<% =sHomePath %>styles.css">
<script type="text/javascript" src="/_scripts/js/main.js"></script>
<script type="text/javascript">
<!-- 
function validateForm() {
	var f=document.forms[0];
	if (!(f)) {
		return false; }
<% If sApplicationName="expert" Then %>
	if (!checkSelectFieldIndex(f.exp_title, 0, "Please select your salutation", 1)) { return false }
	if (!checkTextFieldValue(f.exp_firstname, "", "Please fill in your first name.", 1)) { return false }
	if (!checkTextFieldValue(f.exp_familyname, "", "Please fill in your family name.", 1)) { return false }
	if (f.exp_dbirth.selectedIndex > 0 && f.exp_mbirth.selectedIndex > 0 && f.exp_ybirth.selectedIndex > 0) { 
		if (!checkDateComposition(f.exp_ybirth.value, f.exp_mbirth.value, f.exp_dbirth.value, "Please fill in the date of your birth properly.")) { return false }
	}
	if (!checkTextFieldValue(f.exp_phone, "", "Please specify your primary phone number.", 1)) { return false }
	if (!checkTextFieldValue(f.exp_email, "", "Please specify your primary email.", 1)) { return false }
	if (!validateEmail(f.exp_email.value)) {
		alert("Please retype your e-mail address correctly");
        f.exp_email.select();        
		return;
	}
	if (!checkTextFieldValue(f.Availability, "", "Please specify your availability for this year.", 1)) { return false }
	if (!checkTextFieldLength(f.Availability, 5000, "Please make text of your availability shorter.", 1)) { return false }
	if (!checkTextFieldValue(f.cvEng, "", "Please click Browse and select file with your CV to attach.", 1)) { return false }
<% Else %>
	if (!checkTextFieldValue(f.exp_firstname, "", "Please fill in expert's first name.", 1)) { return false }
	if (!checkTextFieldValue(f.exp_familyname, "", "Please fill in expert's family name.", 1)) { return false }
	if (!checkTextFieldValue(f.exp_email, "", "Please specify expert's primary email.", 1)) { return false }
	if (!validateEmail(f.exp_email.value)) {
		alert("Please retype e-mail address correctly");
        f.exp_email.select();        
		return;
	}
<% End If %>
	f.submit();
}
-->
</script>
</head>
<body topmargin=0 leftmargin=0  marginheight=0 marginwidth=0>
<% ShowTopMenu %>

<% ShowMessageStart "info", 550 %>
<% If sApplicationName="expert" And iExpertID=0 Then %>
	If you have already registered your profile, please <a href="login.asp<% =AddUrlParams(sParams, "url=" + sScriptFullName) %>">log in to update your details</a>.<br><br>
<% End If %>
Please complete this registration form. 
Fields marked with <span class="rs">*</span> are mandatory.<br>
<% ShowMessageEnd %>

  <!-- Personal information -->
	<% InputFormHeader 580, "PERSONAL INFORMATION &amp; CONTACT DETAILS" %>
	<% InputBlockHeader "100%" %>
	<form enctype="multipart/form-data" action="quick_save.asp<%=AddUrlParams(sParams, "act=" & sAction) %>"  method="post" name="register" onSubmit="validateForm(); return false;">
		<% InputBlockSpace 4 %>
		<% InputBlockElementLeftStart %><p class="ftxt">Salutation</p><% InputBlockElementLeftEnd %><% InputBlockElementMiddle %>
		<% InputBlockElementRightStart %>&nbsp;&nbsp;<select name="exp_title" size="1">
		<option value="0"> Please select  </option>
		<% For i=1 to UBound(arrPersonTitleID)
		sFlagSelected=""
		If IsNumeric(iTitleID) And iTitleID>"" Then
			If CInt(iTitleID)=arrPersonTitleID(i) Then
				sFlagSelected=" selected"
			End If
		End If
		Response.Write("<option value=""" & arrPersonTitleID(i) & """" & sFlagSelected & ">"& arrPersonTitle(i) &"</option>")
		Next %>
		</select><% If sApplicationName="expert" Then %><% =sFieldCompulsoryMark %><% End If %><% InputBlockElementRightEnd %>
		<% InputBlockElementLeftStart %><p class="ftxt">First&nbsp;name</p><% InputBlockElementLeftEnd %><% InputBlockElementMiddle %>
		<% InputBlockElementRightStart %>&nbsp;&nbsp;<input type="text" name="exp_firstname" size="31" style="width:355px;" maxlength="100" value="<% =sFirstName %>">&nbsp;&nbsp;<span class="fcmp">*</span>&nbsp;<% InputBlockElementRightEnd %>
		<% InputBlockElementLeftStart %><p class="ftxt">Family&nbsp;name</p><% InputBlockElementLeftEnd %><% InputBlockElementMiddle %>
		<% InputBlockElementRightStart %>&nbsp;&nbsp;<input type="text" name="exp_familyname" size="31" style="width:355px;" maxlength="100" value="<% =sLastName %>">&nbsp;&nbsp;<span class="fcmp">*</span>&nbsp;<% InputBlockElementRightEnd %>
		<% InputBlockElementLeftStart %><p class="ftxt">Date&nbsp;of&nbsp;birth</p><% InputBlockElementLeftEnd %><% InputBlockElementMiddle %>
		<% InputBlockElementRightStart %>&nbsp;&nbsp;<select name="exp_dbirth" size="1">
		<option value="0">Day</option>
		<% For i=1 to 31 
			If iBirthDay=i Then
				Response.Write("<option value=" & i & " selected>" & i & "</option>")
			Else
				Response.Write("<option value=" & i & ">" & i & "</option>")
			End If
		Next %>
		</select>
		<select name="exp_mbirth" size=1>
		<option value="0" selected>Month</option>
		<% For i=1 to UBound(arrMonthID)
			If iBirthMonth=arrMonthID(i) Then
				Response.Write("<option value=" & arrMonthID(i) &" selected>"& arrMonthName(i) &"</option>")
			Else 
				Response.Write("<option value=" & arrMonthID(i) &">"& arrMonthName(i) &"</option>")
			End If
		Next %>
		</select>
		<select name="exp_ybirth" size="1">
		<option value="0">Year</option>
		<% Dim iCurrentYear
		iCurrentYear=Year(Date)
		For i=16 to 96 
			If iBirthYear=(iCurrentYear-i) Then
				Response.Write("<option value=" & (iCurrentYear-i) & " selected>"& (iCurrentYear-i) & "</option>")
			Else 
				Response.Write("<option value=" & (iCurrentYear-i) & ">"& (iCurrentYear-i) & "</option>")
			End if
		Next %>
		</select>&nbsp;<% InputBlockElementRightEnd %>
		<% InputBlockElementLeftStart %><p class="ftxt">Primary phone</p><% InputBlockElementLeftEnd %><% InputBlockElementMiddle %>
		<% InputBlockElementRightStart %>&nbsp;&nbsp;<input type="text" name="exp_phone" size="31" style="width:355px;" maxlength="40" value="<% =sUserPhone %>"><% If sApplicationName="expert" Then %><% =sFieldCompulsoryMark %><% End If %><% InputBlockElementRightEnd %>
		<% InputBlockElementLeftStart %><p class="ftxt">Primary email</p><% InputBlockElementLeftEnd %><% InputBlockElementMiddle %>
		<% InputBlockElementRightStart %>&nbsp;&nbsp;<input type="text" name="exp_email" size="31" style="width:355px;" maxlength="120" value="<% =sUserEmail %>"><% =sFieldCompulsoryMark %><% InputBlockElementRightEnd %>
		<% InputBlockSpace 6 %>
	<% InputBlockFooter %>
	<% InputFormFooter %>
	<% InputFormSpace 12 %>

  <!-- Current availability -->
	<% ShowInputFormHeader 580, "CURRENT AVAILABILITY &amp; ASSIGNMENT PREFERENCES" %>
	<% InputBlockHeader "100%" %>
		<% InputBlockSpace 6 %>
		<% InputBlockElementLeftStart %><p class="ftxt">Availability<br><span class="sml3">(available from month / year)</span></p><% InputBlockElementLeftEnd %><% InputBlockElementMiddle %>
		<% InputBlockElementRightStart %>&nbsp;&nbsp;<textarea cols="31" style="width:355px;" NAME="Availability" rows=4 wrap="yes"><% =sAvailability %></textarea><% If sApplicationName="expert" Then %><% =sFieldCompulsoryMark %><% End If %><% InputBlockElementRightEnd %>
		<% InputBlockElementLeftStart %><p class="ftxt">Preferred project duration</p><% InputBlockElementLeftEnd %><% InputBlockElementMiddle %>
		<% InputBlockElementRightStart %><p class="txt">
		<input type="checkbox" name="shortterm" value="1" <% If iShortterm=1 Then %> checked<% End If %>> Short-term missions &nbsp;&nbsp; <input type="checkbox" name="longterm" value="1"  <% If iLongterm=1 Then %> checked<% End If %>> Long-term missions<% InputBlockElementRightEnd %>
		<% InputBlockSpace 6 %>
	<% InputBlockFooter %>
	<% InputFormFooter %>
	<% InputFormSpace 12 %>

	<% ShowInputFormHeader 580, "ATTACH CV" %>
	<% InputBlockHeader "100%" %>
		<% InputBlockSpace 6 %>
		<% InputBlockElementLeftStart %><p class="ftxt">Upload your current CV</p><% InputBlockElementLeftEnd %><% InputBlockElementMiddle %>
		<% InputBlockElementRightStart %>&nbsp;&nbsp;<input type="file" name="exp_cv" size="41">&nbsp;&nbsp;<span class="fcmp">*</span>&nbsp;<% InputBlockElementRightEnd %>
		<% InputBlockSpace 6 %>
	<% InputBlockFooter %>
	<% InputFormFooter %>
	<% InputFormSpace 15 %>

	<table width=580 cellspacing=0 cellpadding=0 border=0 align="center">
	<tr>
	<td width=580 align=left><img src="<% =sHomePath %>image/x.gif" width=170 height=1><input type="image" src="<% =sHomePath %>image/bte_submit.gif" border=0 alt="  Submit  ">
	&nbsp; &nbsp; &nbsp; &nbsp; <a href="<% =sApplicationHomePath & sParams %>"><img src="<% =sHomePath %>/image/bte_cancel.gif" border=0 alt="  Cancel  "</a>
	</td>
	</tr>
	</form>
	</table><br>

<% CloseDBConnection %>
</body>
</html>

