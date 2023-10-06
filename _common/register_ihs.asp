<% 
'--------------------------------------------------------------------
'
' CV registration.
' Full format. Personal information.
'
'--------------------------------------------------------------------
Response.Expires = 0
Response.ExpiresAbsolute = Now() - 1
Response.CacheControl = "no-cache" 
%>
<!--#include file="_data/datGender.asp"-->
<!--#include file="_data/datPsnTitle.asp"-->
<!--#include file="_data/datPsnStatus.asp"-->
<!--#include file="_data/datMonth.asp"-->
<!--#include file="_data/datCountry.asp"-->
<%
If sApplicationName<>"expert" Then
	CheckUserLogin sScriptFullNameAsParams
End If
CheckExpertID
%>
<!--#include file="../_common/expProfile.asp"-->
<%
Dim sUserLogin, sUserPassword, sUserPhone
Dim sFlagSelected, j
Dim objResult, iResult

If Request.Form()>"" Then
	objResult=SaveExpertFullProfile(iExpertID, Request.Form())

	On Error Resume Next
	' Execute the script from the active application
		AfterCvRegistrationStep1 objResult
	On Error GoTo 0
	
	If sApplicationName="expert" Then
		If objResult(0)=0 Then
			If objResult(1)>0 Then
				iExpertID=objResult(1)
				iUserID=objResult(2)
				sUserLogin=objResult(3)
				sUserPassword=objResult(4)

				' Login active user
				objTempRs2=UpdateRecordSP("usp_LogSessionUser", _
					Array(Array(, adVarChar, 40, sSessionID), Array(, adInteger, , iUserID)))
		
			End If
		End If
	End If
	Set objResult=Nothing

	Response.Redirect "register2.asp" & sParams
End If                                                       

LoadExpertProfile iExpertID
LoadExpertNationality iExpertID
%>
<html>
<head>
<title>CV registration. Personal information</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" type="text/css" href="<% =sHomePath %>styles.css">
<script type="text/javascript" src="/_scripts/js/main.js"></script>
<script type="text/javascript">
<!-- 
function validateForm() {
<%
Dim sUserSalutation
If sApplicationName="expert" Then
	sUserSalutation="your"
Else
	sUserSalutation="expert's"
End If
%>
	var f=document.forms[0];
	if (!(f)) {
		return false; }
<% If sApplicationName="external" Then %>
	f.submit();
	return;
<% End If %>
	if (!checkSelectFieldIndex(f.exp_title, 0, "Please select <% =sUserSalutation %> title.", 1)) { return false }
	if (!checkTextFieldValue(f.exp_firstname, "", "Please fill in <% =sUserSalutation %> first name.", 1)) { return false }
	if (!checkTextFieldValue(f.exp_familyname, "", "Please fill in <% =sUserSalutation %> family name.", 1)) { return false }
	if (f.exp_dbirth.selectedIndex > 0 && f.exp_mbirth.selectedIndex > 0 && f.exp_ybirth.selectedIndex > 0) { 
		if (!checkDateComposition(f.exp_ybirth.value, f.exp_mbirth.value, f.exp_dbirth.value, "Please fill in the date of <% =sUserSalutation %> birth properly.")) { return false }
	}
	if (!checkTextFieldValue(f.newloc, "", "Please specify <% =sUserSalutation %> nationality.", 0)) { return false }
<% If sApplicationName="expert" Then %>
	if (!(f.exp_gender[0].checked || f.exp_gender[1].checked)) {
		alert("Please specify <% =sUserSalutation %> gender.");
		return false }
	if (!checkTextFieldValue(f.exp_phone, "", "Please specify <% =sUserSalutation %> primary phone number.", 1)) { return false }
		
<% End If %>
		
	if (!checkTextFieldValue(f.exp_email, "", "Please specify <% =sUserSalutation %> primary email.", 1)) { return false }
	if (!validateEmail(f.exp_email.value)) {
		alert("Please retype <% =sUserSalutation %> email address correctly.");
        f.exp_email.select();        
		return;
   }
   
<% If sApplicationName="expert" Then %>
	if (!checkTextFieldValue(f.exp_curr_Position, "", "Please fill in <% =sUserSalutation %> current position.", 1)) { return false }
<% End If %>
	if (!checkTextFieldValue(f.exp_key_qualif, "", "Please fill in <% =sUserSalutation %> key qualifications.", 1)) { return false }
	if (!checkTextFieldLength(f.exp_key_qualif, 25000, "Please make text of key qualifications shorter.", 1)) { return false }
<% If sApplicationName="expert" Then %>
	if (!checkTextFieldValue(f.exp_wke_years, "", "Please specify the years of <% =sUserSalutation %> professional experience.", 1)) { return false }
<% End If %>
	f.submit();
}

//////////////////:starts from here /////////////:::
/////////////////////////////////////
    function moveloc(intDirection)
    {
    	var heading = '';
    	var msg = '';
    	var flag;
    	var arrnew = new Array();
    
    	with (document.locations){
    		if (intDirection){
    			//Add it to mylocations
    			msg = '';
    			for(var x=0;x<locations.length;x++){
    				var opt = locations.options[x];
    				if (opt.selected){
    					flag = 1;
    					//if more then 20 then alert and exit.
    					if (mylocations.length > 19){
    						alert("You are only allowed 20 locations.");
    						break;	
    					}
    					//check if option exists if not add it
    					for (var y=0;y<mylocations.length;y++){
    						var myopt = mylocations.options[y];
    						if (myopt.value == opt.value){	
    							flag = 0;
    						}
    					}
    					if (flag){
    						//This is not a duplicate so add it to the select box.
    						mylocations.options[mylocations.options.length] = new Option(opt.text, opt.value, 0, 0);
    					}
    				}
    			}
    		}else{
    			//Delete it from my locations
    			for(var x=mylocations.length-1;x>=0;x--){
    				var opt = mylocations.options[x];
    				if (opt.selected){									
    					//Remove it from the select box
    					mylocations.options[x] = null;
    				}
    			}		
    		}
    
    		//Fill hidden field with new values
    		for (var y=0;y<mylocations.length;y++){
    			arrnew[y] = mylocations.options[y].value
    		}			
    		newloc.value = arrnew.join();
    
    	}
    }
    
function SetNationality() {
	var arrnew = new Array();
	with (document.locations){
		
		//Fill hidden field with new values
		for (var y=0;y<mylocations.length;y++){
			arrnew[y] = mylocations.options[y].value
		}			
		newloc.value = arrnew.join();

	}
	return 1;
}
-->
</script>
</head>

<body bgcolor="#FFFFFF" topmargin=0 leftmargin=0  marginheight=0 marginwidth=0>
<% ShowTopMenu %>
<% ShowRegistrationProgressBar "CV", 1 %>

  <!-- [i] CV online -->
<% ShowMessageStart "info", 580 %>
<% If sApplicationName="expert" And iExpertID=0 Then %>
	If you have already registered your profile, please <a href="login.asp<% =AddUrlParams(sParams, "url=" + sScriptFullName) %>">log in to update your details</a>.<br><br>
<% End If %>
Please fill in all the relevant information and as many details on your experience as possible. <br>Fields marked with <span class="fcmp">*</span> are required information.
<br><br>
You can always go back and edit each section by clicking on the menu at the top.
<br>
<% ShowMessageEnd %>

  <!-- Personal information -->
	<% InputFormHeader 580, "PERSONAL INFORMATION" %>
	<% InputBlockHeader "100%" %>
	<form method="post" action="register.asp<% =sParams %>" name="locations" onSubmit="validateForm(); return false;">
	<input type="hidden" name="reg_type" value="">
	<input type="hidden" name="newloc" value="">
	<input type="hidden" name="id_Person" value="<% =iPersonID %>">
		<% InputBlockSpace 4 %>
		<% InputBlockElementLeftStart %><p class="ftxt">Title</p><% InputBlockElementLeftEnd %><% InputBlockElementMiddle %>
		<% InputBlockElementRightStart %>&nbsp;&nbsp;<select name="exp_title" size="1">
		<option value="0" selected> Please select  </option>
		<% For i=1 to UBound(arrPersonTitleID)
		sFlagSelected=""
		If IsNumeric(iTitleID) And iTitleID>"" Then
			If CInt(iTitleID)=arrPersonTitleID(i) Then
				sFlagSelected=" selected"
			End If
		End If
		Response.Write("<option value=""" & arrPersonTitleID(i) & """" & sFlagSelected & ">"& arrPersonTitle(i) &"</option>")
		Next %>
		</select>&nbsp;&nbsp;<span class="fcmp">*</span>&nbsp;<% InputBlockElementRightEnd %>
		<% InputBlockElementLeftStart %><p class="ftxt">First&nbsp;name</p><% InputBlockElementLeftEnd %><% InputBlockElementMiddle %>
		<% InputBlockElementRightStart %>&nbsp;&nbsp;<input type="text" name="exp_firstname" size="45" style="width:355px;" maxlength=255 value="<% =sFirstName %>">&nbsp;&nbsp;<span class="fcmp">*</span>&nbsp;<% InputBlockElementRightEnd %>
		<% InputBlockElementLeftStart %><p class="ftxt">Middle&nbsp;name</p><% InputBlockElementLeftEnd %><% InputBlockElementMiddle %>
		<% InputBlockElementRightStart %>&nbsp;&nbsp;<input type="text" name="exp_middlename" size="45" style="width:355px;" maxlength=255 value="<% =sMiddleName %>"><% InputBlockElementRightEnd %>
		<% InputBlockElementLeftStart %><p class="ftxt">Family&nbsp;name</p><% InputBlockElementLeftEnd %><% InputBlockElementMiddle %>
		<% InputBlockElementRightStart %>&nbsp;&nbsp;<input type="text" name="exp_familyname" size="45" style="width:355px;" maxlength=255 value="<% =sLastName %>">&nbsp;&nbsp;<span class="fcmp">*</span>&nbsp;<% InputBlockElementRightEnd %>
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
		</select><% If sApplicationName="expert" Then %>&nbsp;&nbsp;<span class="fcmp">*</span><% End If %>&nbsp;<% InputBlockElementRightEnd %>
		<% InputBlockElementLeftStart %><p class="ftxt">Place&nbsp;of&nbsp;birth</p><% InputBlockElementLeftEnd %><% InputBlockElementMiddle %>
		<% InputBlockElementRightStart %>&nbsp;&nbsp;<input type="text" name="exp_birthplace" size="45" style="width:355px;" maxlength=255 value="<% =sBirthPlace %>">&nbsp;<% InputBlockElementRightEnd %>
		<% InputBlockElementLeftStart %><p class="ftxt">Nationality</p><% InputBlockElementLeftEnd %><% InputBlockElementMiddle %>
		<% InputBlockElementRightStart %>&nbsp;&nbsp;<select name="locations" multiple rows="4" size="4" style="width:355px;">
		<% For i=0 To UBound(arrCountryID)-1
		sFlagSelected=""
		If IsArray(arrExpNationalityID) Then
			For j=1 To Ubound(arrExpNationalityID)
				If arrExpNationalityID(j)=arrCountryID(i) Then
					sFlagSelected=" selected"
				End If
			Next
		End If
		Response.Write ("<option value=""" & arrCountryID(i) & """" & sFlagSelected & ">"& arrCountryName(i) & "</option>")
		Next %>
		</select>&nbsp;&nbsp;<span class="fcmp">*</span>&nbsp;&nbsp;
		<p class="sml">&nbsp;&nbsp;Click on <b>Add</b> to add a selected nationality to your list.<br>
                &nbsp;&nbsp;If you want to remove a nationality, highlight it and click on <b>Remove</b></p>
    			<table width="100%" cellpadding="0" cellspacing="0" border="0">
    			<tr>
    			<td width="50%" valign="top" align="center"><p>
    			<a href="javascript:moveloc(1)">Add</a>
    			</td>
    			<td width="50%" valign="top" align="center"><p>
    			<a href="javascript:moveloc(0)">Remove</a> 
    			</td>
    			</table>
		&nbsp;&nbsp;<select name="mylocations" multiple style="width:355px;" rows="3" size="3">
		<% 	
		If IsArray(arrExpNationalityID) Then
			For i=1 To Ubound(arrExpNationalityID)
				For j=0 To UBound(arrCountryID)-1
					If arrExpNationalityID(i)=arrCountryID(j) Then
						If arrCountryName(j)>"" Then Response.Write ("<option value="& arrCountryID(j) & ">"& arrCountryName(j) & "</option>")
					End If
				Next
			Next
		End If
		%>
		</select><% InputBlockElementRightEnd %>
		<% InputBlockElementLeftStart %><p class="ftxt">Gender</p><% InputBlockElementLeftEnd %><% InputBlockElementMiddle %>
		<% InputBlockElementRightStart %><p><% if iGenderID="1" then 
		 	Response.Write("<input type='radio' name='exp_gender' value='1' checked>")
		else
			Response.Write("<input type='radio' name='exp_gender' value='1'>")
		end if %>
		<b>male &nbsp; &nbsp; 
		<% if iGenderID="2" then 
		 	Response.Write("<input type='radio' name='exp_gender' value='2' checked>")
		else
			Response.Write("<input type='radio' name='exp_gender' value='2'>")
		end if %>
		female<% If sApplicationName="expert" Then %>&nbsp;&nbsp;</b><span class="fcmp">*</span><% End If %>&nbsp;<% InputBlockElementRightEnd %>
		<% InputBlockElementLeftStart %><p class="ftxt">Marital&nbsp;status</p><% InputBlockElementLeftEnd %><% InputBlockElementMiddle %>
		<% InputBlockElementRightStart %>&nbsp;&nbsp;<select name="exp_marstatus" size="1">
		<option value="0" selected> - Please select - </option>
		<% For i=1 to UBound(arrMaritalStatusID)
		if iMaritalStatusID=arrMaritalStatusID(i) then
			Response.Write("<option value=" & arrMaritalStatusID(i) &" selected>"& arrMaritalStatusTitle(i) &"</option>")
		else
			Response.Write("<option value=" & arrMaritalStatusID(i) &">"& arrMaritalStatusTitle(i) &"</option>")
		end if
		Next %>
		</select><% InputBlockElementRightEnd %>
		<% InputBlockElementLeftStart %><p class="ftxt">Primary phone</p><% InputBlockElementLeftEnd %><% InputBlockElementMiddle %>
		<% InputBlockElementRightStart %>&nbsp;&nbsp;<input type="text" name="exp_phone" size="45" style="width:355px;" maxlength=50 value="<% =sUserPhone %>"><% If sApplicationName="expert" Then %>&nbsp;&nbsp;<span class="fcmp">*</span><% End If %>&nbsp;<% InputBlockElementRightEnd %>
		<% InputBlockElementLeftStart %><p class="ftxt">Primary email</p><% InputBlockElementLeftEnd %><% InputBlockElementMiddle %>
		<% InputBlockElementRightStart %>&nbsp;&nbsp;<input type="text" name="exp_email" maxlength=120 size="45" style="width:355px;" value="<% =sUserEmail %>">&nbsp;&nbsp;<span class="fcmp">*</span>&nbsp;<% InputBlockElementRightEnd %>
		<% InputBlockSpace 4 %>
	<% InputBlockFooter %>
	<% InputFormAfterBlock %>
	<% InputFormDualLine %>

   <!-- Professional information  -->
	<% InputFormBeforeBlock 580 %>
	<% InputBlockHeader "100%" %>
		<% InputBlockSpace 4 %>
		
<!--
		<% InputBlockElementLeftStart %><p class="ftxt">Profession</p><% InputBlockElementLeftEnd %><% InputBlockElementMiddle %>
		<% InputBlockElementRightStart %>&nbsp;&nbsp;<input type="text" name="exp_Prof" size="45" style="width:355px;" maxlength=255 value="<% =sProfession %>">&nbsp;<% InputBlockElementRightEnd %>
-->		
		<% InputBlockElementLeftStart %><p class="ftxt">Current&nbsp;position</p><% InputBlockElementLeftEnd %><% InputBlockElementMiddle %>
		<% InputBlockElementRightStart %>&nbsp;&nbsp;<input type="text" name="exp_curr_Position" size="45" style="width:355px;" maxlength=255  value="<% =sCurrPosition %>"><% If sApplicationName="expert" Then %>&nbsp;&nbsp;<span class="fcmp">*</span><% End If %>&nbsp;<% InputBlockElementRightEnd %>
		<% InputBlockElementLeftStart %><p class="ftxt">Key&nbsp;qualifications</p><% InputBlockElementLeftEnd %><% InputBlockElementMiddle %>
		<% InputBlockElementRightStart %>&nbsp;&nbsp;<textarea cols="34" style="width:355px;" name="exp_key_qualif" rows=4 wrap="yes"><% =sKeyQualifications %></textarea>&nbsp;&nbsp;<span class="fcmp">*</span>&nbsp;<% InputBlockElementRightEnd %>
		<% InputBlockElementLeftStart %><p class="ftxt">Years of professional experience</p><% InputBlockElementLeftEnd %><% InputBlockElementMiddle %>
		<% InputBlockElementRightStart %>&nbsp;&nbsp;<input type="text" name="exp_wke_years" size="5" maxlength=2 onBlur="checkNumeric(this, 'Please enter only numbers for your experience', 1)" value="<% =iExperienceYears %>"><% If sApplicationName="expert" Then %>&nbsp;&nbsp;<span class="fcmp">*</span><% End If %>&nbsp;&nbsp;<span class="sml">(use only numbers)</span><% InputBlockElementRightEnd %>
		<% InputBlockSpace 6 %>
	<% InputBlockFooter %>
	<% InputFormFooter %>
	<% InputFormSpace 12 %>

	<table width=580 cellspacing=0 cellpadding=0 border=0 align="center">
	<tr>
        <td width="300" align="right"><img src="<% =sHomePath %>image/x.gif" width=170 height=1><input type="image" src="<% =sHomePath %>image/bte_savecont.gif" name="Save & continue" alt="Save & continue" border=0></td>
	</tr>
	</form>
	</table><br>
<script type="text/javascript">
<!--
    SetNationality();
//-->
</script>

<% CloseDBConnection %>
</body>
</html>
