<% 
'--------------------------------------------------------------------
'
' CV registration.
' Full format. Personal information.
'
'--------------------------------------------------------------------
%>
<!--#include virtual="/_template/asp.header.nocache.asp"-->
<%
If sApplicationName<>"expert" Then
	CheckUserLogin sScriptFullNameAsParams
End If
CheckExpertID

' Log: 34 - Update expert
If Request.Form()>"" Then
	iLogResult = LogActivity(34, "ExpertID=" & Cstr(iExpertID) & " SavedStep: 1", "", "")
End If
%>
<!--#include file="../_common/expProfile.asp"-->
<%
Dim sUserLogin, sUserPassword, sUserPhone
Dim sFlagSelected, j
Dim objResult, iResult

If Request.Form("exp_firstname")>"" And Request.Form("exp_familyname")>"" Then
	objResult=SaveExpertFullProfile(objExpertDB.DatabaseCode, iExpertID, Request.Form())

	On Error Resume Next
	' Execute the script from the active application
		AfterCvRegistrationStep1 objResult
	' Save custom fields
		sCvLanguage=Left(Request.Form("exp_language"), 3)
		sCvFolder=Left(Request.Form("exp_type"), 150)

		If Len(sCvLanguage)>0 Or Len(sCvFolder)>0 Then
			If sApplicationName="expert" Then
				SaveExpertLanguage iExpertID, sCvLanguage
			Else
				SaveExpertLanguageAndFolder iExpertID, sCvLanguage, sCvFolder
			End If
		End If
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

	If Request.Form("next")="0" Then
	Else
		Response.Redirect "register2.asp" & sParams
	End If
End If

LoadExpertProfile objExpertDB.DatabaseCode, iExpertID
LoadExpertNationality objExpertDB.DatabaseCode, iExpertID
%>
<!--#include file="_data/datGender.asp"-->
<!--#include file="_data/datPsnTitle.asp"-->
<!--#include file="_data/datPsnStatus.asp"-->
<!--#include file="_data/datMonth.asp"-->
<!--#include file="_data/datCountry.asp"-->
<!--#include virtual="/_template/html.header.start.asp"-->
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
		return false; 
	}
<% If sApplicationName="external" Or sApplicationName="backoffice" Then %>
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

function submitForm() {
	var f=document.forms[0];
	if (f.next) {
		f.next.value = 0;
	}
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
	<% 
	If Not bIsMyCV Then 
		%><div id="hdrUpdatedList" class="colCCCCCC uprCse f17 spc01 botMrgn10"><span class="service_title">Curriculum Vitae.</span> Expert ID: <% =objExpertDB.DatabaseCode %><%=iExpertID%></div>
		<% 
	Else
		%><div class="colCCCCCC uprCse f17 spc01 botMrgn10"><span class="service_title">Curriculum Vitae</span></div>
		<% 
	End If 
	
	ShowRegistrationProgressBar "CV", 1 
	%>

  <!-- [i] CV online -->
<% ShowMessageStart "info", 580 %>
<% If sApplicationName="expert" And iExpertID=0 Then %>
	<% = GetLabel(sCvLanguage, "If you have already registered your profile") %><br /><br />
<% End If %>
<% = GetLabel(sCvLanguage, "Please fill in all the relevant information") %><br /><% = GetLabel(sCvLanguage, "Fields marked with *") %>
<br /><br />
<% = GetLabel(sCvLanguage, "You can always go back") %>
<br />
<% ShowMessageEnd %>

  <!-- Personal information -->
	<form method="post" action="register.asp<% =sParams %>" name="locations" onSubmit="validateForm(); return false;">
	<input type="hidden" name="reg_type" value=""><input type="hidden" name="next" value="1">
	<input type="hidden" name="newloc" value="">
	<input type="hidden" name="id_Person" value="<% =iPersonID %>">
		<div class="box search blue">
		<h3><% =GetLabel(sCvLanguage, "Personal information") %></h3>
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
		<td class="field splitter"><label for="exp_title"><% = GetLabel(sCvLanguage, "Personal title") %></label></td>
		<td class="value blue"><select name="exp_title" size="1" style="width:130px;">
		<option value="0" selected> <% = GetLabel(sCvLanguage, "Please select") %> </option>
		<% For i=1 to UBound(arrPersonTitleID)
		sFlagSelected=""
		If IsNumeric(iTitleID) And iTitleID>"" Then
			If CInt(iTitleID)=arrPersonTitleID(i) Then
				sFlagSelected=" selected"
			End If
		End If
		Response.Write("<option value=""" & arrPersonTitleID(i) & """" & sFlagSelected & ">"& arrPersonTitle(i) &"</option>")
		Next %>
		</select>&nbsp;&nbsp;<span class="fcmp">*</span>&nbsp;</td>
		</tr>
		<tr>
		<td class="field splitter"><label for="exp_firstname"><% = GetLabel(sCvLanguage, "First name") %></label></td>
		<td class="value blue"><input type="text" id="exp_firstname" name="exp_firstname" size="45" style="width:355px;" maxlength=255 value="<% =sFirstName %>">&nbsp;&nbsp;<span class="fcmp">*</span>&nbsp;</td>
		</tr>
		<tr>
		<td class="field splitter"><label for="exp_middlename"><% = GetLabel(sCvLanguage, "Middle name") %></label></td>
		<td class="value blue"><input type="text" id="exp_middlename" name="exp_middlename" size="45" style="width:355px;" maxlength=255 value="<% =sMiddleName %>"></td>
		</tr>
		<tr>
		<td class="field splitter"><label for="exp_familyname"><% = GetLabel(sCvLanguage, "Family name") %></label></td>
		<td class="value blue"><input type="text" id="exp_familyname" name="exp_familyname" size="45" style="width:355px;" maxlength=255 value="<% =sLastName %>">&nbsp;&nbsp;<span class="fcmp">*</span>&nbsp;</td>
		</tr>
		<tr>
		<td class="field splitter"><label for="exp_dbirth"><% = GetLabel(sCvLanguage, "Date of birth") %></label></td>
		<td class="value blue"><select id="exp_dbirth" name="exp_dbirth" size="1">
		<option value="0"><% = GetLabel(sCvLanguage, "Day") %></option>
		<% For i=1 to 31 
			If iBirthDay=i Then
				Response.Write("<option value=" & i & " selected>" & i & "</option>")
			Else
				Response.Write("<option value=" & i & ">" & i & "</option>")
			End If
		Next %>
		</select>
		<select name="exp_mbirth" size=1>
		<option value="0" selected><% = GetLabel(sCvLanguage, "Month") %></option>
		<% For i=1 to UBound(arrMonthID)
			If iBirthMonth=arrMonthID(i) Then
				Response.Write("<option value=" & arrMonthID(i) &" selected>"& arrMonthName(i) &"</option>")
			Else 
				Response.Write("<option value=" & arrMonthID(i) &">"& arrMonthName(i) &"</option>")
			End If
		Next %>
		</select>
		<select name="exp_ybirth" size="1">
		<option value="0"><% = GetLabel(sCvLanguage, "Year") %></option>
		<% Dim iCurrentYear
		iCurrentYear=Year(Date)
		For i=16 to 96 
			If iBirthYear=(iCurrentYear-i) Then
				Response.Write("<option value=" & (iCurrentYear-i) & " selected>"& (iCurrentYear-i) & "</option>")
			Else 
				Response.Write("<option value=" & (iCurrentYear-i) & ">"& (iCurrentYear-i) & "</option>")
			End if
		Next %>
		</select><% If sApplicationName="expert" Then %>&nbsp;&nbsp;<span class="fcmp">*</span><% End If %>&nbsp;</td>
		</tr>
		<tr>
		<td class="field splitter"><label for="exp_birthplace"><% = GetLabel(sCvLanguage, "Place of birth") %></label></td>
		<td class="value blue"><input type="text" id="exp_birthplace" name="exp_birthplace" size="45" style="width:355px;" maxlength=255 value="<% =sBirthPlace %>"></td>
		</tr>
		<tr>
		<td class="field splitter"><label for="locations"><% = GetLabel(sCvLanguage, "Nationality") %></label></td>
		<td class="value blue"><select id="locations" name="locations" multiple rows="4" size="4" style="width:355px;">
		<% For i=0 To UBound(arrCountryID)-1
		sFlagSelected=""
		If IsArray(arrExpNationalityID) Then
			For j=LBound(arrExpNationalityID) To Ubound(arrExpNationalityID)
				If CheckIntegerAndNull(arrExpNationalityID(j))=arrCountryID(i) Then
					sFlagSelected=" selected"
				End If
			Next
		End If
		Response.Write ("<option value=""" & arrCountryID(i) & """" & sFlagSelected & ">"& arrCountryName(i) & "</option>")
		Next %>
		</select>&nbsp;&nbsp;<span class="fcmp">*</span>&nbsp;&nbsp;
		<p class="sml">&nbsp;&nbsp;<% = GetLabel(sCvLanguage, "Add nationality") %><br />
                &nbsp;&nbsp;<% = GetLabel(sCvLanguage, "Remove nationality") %></p>
    			<table width="355" cellpadding="1" cellspacing="1" border="0">
    			<tr>
    			<td width="50%" valign="top" align="center"><p>
    			<a href="javascript:moveloc(1)"><% = GetLabel(sCvLanguage, "Add") %></a>
    			</td>
    			<td width="50%" valign="top" align="center"><p>
    			<a href="javascript:moveloc(0)"><% = GetLabel(sCvLanguage, "Remove") %></a> 
    			</td>
    			</table>
		<select name="mylocations" multiple style="width:355px;" rows="3" size="3">
		<% 	
		If IsArray(arrExpNationalityID) Then
			For i=LBound(arrExpNationalityID) To Ubound(arrExpNationalityID)
				For j=0 To UBound(arrCountryID)-1
					If CheckIntegerAndNull(arrExpNationalityID(i))=arrCountryID(j) Then
						If arrCountryName(j)>"" Then Response.Write ("<option value="& arrCountryID(j) & ">"& arrCountryName(j) & "</option>")
					End If
				Next
			Next
		End If
		%>
		</select></td>
		</tr>
		<tr>
		<td class="field splitter"><label for="exp_gender"><% = GetLabel(sCvLanguage, "Gender") %></label></td>
		<td class="value blue"><p>
		<input type="radio" name="exp_gender" value="1" <% If iGenderID="1" Then Response.Write "checked" %>>
		<b><% = GetLabel(sCvLanguage, "male") %> &nbsp; &nbsp;</b>
		<input type="radio" name="exp_gender" value="2" <% If iGenderID="2" Then Response.Write "checked" %>>
		<b><% = GetLabel(sCvLanguage, "female") %>
		<% If sApplicationName="expert" Then %>&nbsp;&nbsp;</b><span class="fcmp">*</span><% End If %>&nbsp;
	</td>
		</tr>
		<tr>
		<td class="field splitter"><label for="exp_marstatus"><% = GetLabel(sCvLanguage, "Marital status") %></label></td>
		<td class="value blue"><select id="exp_marstatus" name="exp_marstatus" size="1">
		<option value="0" selected> </option>
		<% For i=1 to UBound(arrMaritalStatusID)
		if iMaritalStatusID=arrMaritalStatusID(i) then
			Response.Write("<option value=" & arrMaritalStatusID(i) &" selected>"& arrMaritalStatusTitle(i) &"</option>")
		else
			Response.Write("<option value=" & arrMaritalStatusID(i) &">"& arrMaritalStatusTitle(i) &"</option>")
		end if
		Next %>
		</select></td>
		</tr>
		<tr>
		<td class="field splitter"><label for="exp_phone"><% = GetLabel(sCvLanguage, "Primary phone") %></label></td>
		<td class="value blue"><input type="text" id="exp_phone" name="exp_phone" size="45" style="width:355px;" maxlength=50 value="<% =sPhone %>"><% If sApplicationName="expert" Then %>&nbsp;&nbsp;<span class="fcmp">*</span><% End If %>&nbsp;</td>
		</tr>
		<tr class="last">
		<td class="field splitter"><label for="exp_email"><% = GetLabel(sCvLanguage, "Primary email") %></label></td>
		<td class="value blue"><input type="text" id="exp_email" name="exp_email" maxlength=120 size="45" style="width:355px;" value="<% =sEmail %>">&nbsp;&nbsp;<span class="fcmp">*</span>&nbsp;</td>
		</tr>
		</table>

		<table class="search_form" width="100%" cellspacing=0 cellpadding=0 border=0>
		<tr class="first">
		<td class="field splitter"><label for="exp_curr_Position"><% = GetLabel(sCvLanguage, "Current position") %></label></td>
		<td class="value blue"><input type="text" id="exp_curr_Position" name="exp_curr_Position" size="45" style="width:355px;" maxlength=255  value="<% =sCurrPosition %>"><% If sApplicationName="expert" Then %>&nbsp;&nbsp;<span class="fcmp">*</span><% End If %>&nbsp;</td>
		</tr>
		<tr>
		<td class="field splitter"><label for="exp_key_qualif"><% = GetLabel(sCvLanguage, "Key qualifications") %></label></td>
		<td class="value blue"><textarea id="exp_key_qualif" name="exp_key_qualif" cols="34" style="width:355px;" rows=4 wrap="yes"><% =sKeyQualifications %></textarea>&nbsp;&nbsp;<span class="fcmp">*</span>&nbsp;</td>
		</tr>
		<tr>
		<td class="field splitter"><label for="exp_wke_years"><% = GetLabel(sCvLanguage, "Years of professional experience") %></label></td>
		<td class="value blue"><input type="text" id="exp_wke_years" name="exp_wke_years" size="5" maxlength=2 onBlur="checkNumeric(this, 'Please enter only numbers for your experience', 1)" value="<% =iExperienceYears %>"><% If sApplicationName="expert" Then %>&nbsp;&nbsp;<span class="fcmp">*</span><% End If %>&nbsp;&nbsp;<span class="sml">(<% = GetLabel(sCvLanguage, "use only numbers") %>)</span></td>
		</tr>
		</table>
		</div>

		<div class="spacebottom">
			<input type="submit" id="btnSubmit" name="btnSubmit" class="red-button w125 under-right-col" value="Save & continue" />
		</div>
		</form>

	</div>

<!-- footer -->
<!--#include file="../_template/page.footer.asp"-->

<% CloseDBConnection %>
<script type="text/javascript">
<!--
    SetNationality();
//-->
</script>
</body>
<!--#include file="../_template/html.footer.asp"-->
