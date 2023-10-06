<% 
'--------------------------------------------------------------------
'
' CV registration.
' Full format. Addresses and availability.
'
'--------------------------------------------------------------------
%>
<!--#include virtual="/_template/asp.header.nocache.asp"-->
<!--#include file="_data/datCountry.asp"-->
<% 
If sApplicationName<>"expert" Then
	CheckUserLogin sScriptFullNameAsParams
End If
CheckExpertID()

' Log: 34 - Update expert
If Request.Form()>"" Then
	iLogResult = LogActivity(34, "ExpertID=" & Cstr(iExpertID) & " SavedStep: 6", "", "")
End If

Dim objConnCustom
Set objConnCustom = Server.CreateObject("ADODB.Connection")
objConnCustom.Open "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Data Source=MAGNESA;Initial Catalog=" & objExpertDB.DatabasePath & ";"

' The page is accessible only if expert id is valid
If iExpertID<=0 Then Response.Redirect(sApplicationHomePath & "register.asp" & sParams)

Dim sEmail, sPhone
Dim sExpAvailability, iExpShortterm, iExpLongterm, sExpMembership, sExpOtherSkills, sExpPublications, sExpReferences, sPreferences
Dim bExpPermanentAddressExists, iExpPermanentAddressID, sExpPermanentStreet, sExpPermanentPostcode, sExpPermanentCity, iExpPermanentCountryID, sExpPermanentPhone, sExpPermanentMobile, sExpPermanentFax, sExpPermanentEmail, sExpPermanentWebsite
Dim bExpCurrentAddressExists, iExpCurrentAddressID, sExpCurrentStreet, sExpCurrentPostcode, sExpCurrentCity, iExpCurrentCountryID, sExpCurrentPhone, sExpCurrentMobile, sExpCurrentFax, sExpCurrentEmail, sExpCurrentWebsite
Dim bEmailAlreadySent

If Request.Form()>"" Then
	
	If Request.Form("shortterm")<>"" Then
		iExpShortterm=1
	Else
		iExpShortterm=0
	End If  
 	If Request.Form("longterm")<>"" then
		iExpLongterm=1
	Else
		iExpLongterm=0
	End If  

	sExpAvailability=Left(CheckString(Request.Form("availability")), 400)
	sPreferences=Left(CheckString(Request.Form("preferences")), 50000)
	
	sExpOtherSkills=Left(CheckString(Request.Form("otherskills")), 50000)
	sExpMembership=Left(CheckString(Request.Form("memberships")), 50000)
	sExpPublications=Left(CheckString(Request.Form("publications")), 50000)
	sExpReferences=Left(CheckString(Request.Form("references")), 50000)

	objTempRs=UpdateRecordSPWithConn(objConnCustom, "usp_ExpertAvailabilityUpdate", Array( _
		Array(, adVarChar, 3, "Eng"), _
		Array(, adInteger, , iExpertID), _
		Array(, adVarWChar, 400, sExpAvailability), _
		Array(, adTinyInt, , iExpShortterm), _
		Array(, adTinyInt, , iExpLongterm), _
		Array(, adLongVarWChar, 50000, sPreferences), _
		Array(, adLongVarWChar, 50000, sExpOtherSkills), _
		Array(, adLongVarWChar, 50000, sExpMembership), _
		Array(, adLongVarWChar, 50000, sExpPublications), _
		Array(, adLongVarWChar, 50000, sExpReferences)))

	' Permanent address
	iExpPermanentAddressID=CheckString(Request.Form("exp_pma_adrid"))
	sExpPermanentStreet=Left(CheckString(Request.Form("exp_pma_street")), 255)
	sExpPermanentPostcode=Left(CheckString(Request.Form("exp_pma_postcode")), 50)
	sExpPermanentCity=Left(CheckString(Request.Form("exp_pma_city")), 250)
	iExpPermanentCountryID=ReplaceIfEmpty(CheckString(Request.Form("exp_pma_cou")), Null)
	sExpPermanentPhone=Left(CheckString(Request.Form("exp_pma_phone")), 250)
	sExpPermanentMobile=Left(CheckString(Request.Form("exp_pma_mobile")), 250)
	sExpPermanentFax=Left(CheckString(Request.Form("exp_pma_fax")), 250)
	sExpPermanentEmail=Left(CheckString(Request.Form("exp_pma_email")), 250)
	sExpPermanentWebsite=Left(CheckString(Request.Form("exp_pma_web")), 255)

	bExpPermanentAddressExists=Len(Trim(sExpPermanentStreet)) + Len(Trim(sExpPermanentPostcode)) + Len(Trim(sExpPermanentCity)) + Len(Trim(sExpPermanentPhone)) + Len(Trim(sExpPermanentMobile)) + Len(Trim(sExpPermanentFax)) + Len(Trim(sExpPermanentEmail)) + Len(Trim(sExpPermanentWebsite))
        If IsNumeric(iExpPermanentCountryID) Then bExpPermanentAddressExists=bExpPermanentAddressExists + iExpPermanentCountryID
	If IsNumeric(bExpPermanentAddressExists) Then 
		bExpPermanentAddressExists=CInt(bExpPermanentAddressExists)
	Else
		bExpPermanentAddressExists=0
	End If

	If IsNumeric(iExpPermanentAddressID) And iExpPermanentAddressID>"" And iExpPermanentAddressID<>"0" Then
		objTempRs=UpdateRecordSPWithConn(objConnCustom, "usp_ExpCvvAddressUpdate", Array( _
			Array(, adVarChar, 3, "Eng"), _
			Array(, adInteger, , iExpertID), _
			Array(, adInteger, , 1), _
			Array(, adInteger, , iExpPermanentAddressID), _
			Array(, adVarWChar, 255, sExpPermanentStreet), _
			Array(, adVarWChar, 50, sExpPermanentPostcode), _
			Array(, adVarWChar, 150, sExpPermanentCity), _
			Array(, adInteger, , iExpPermanentCountryID), _
			Array(, adVarWChar, 150, sExpPermanentPhone), _
			Array(, adVarWChar, 150, sExpPermanentMobile), _
			Array(, adVarWChar, 150, sExpPermanentFax), _
			Array(, adVarWChar, 150, sExpPermanentEmail), _
			Array(, adVarWChar, 255, sExpPermanentWebsite)))

	ElseIf bExpPermanentAddressExists>0 Then
		objTempRs=InsertRecordSPWithConn(objConnCustom, "usp_ExpCvvAddressInsert", Array( _
			Array(, adVarChar, 3, "Eng"), _
			Array(, adInteger, , iExpertID), _
			Array(, adInteger, , 1), _
			Array(, adVarWChar, 255, sExpPermanentStreet), _
			Array(, adVarWChar, 50, sExpPermanentPostcode), _
			Array(, adVarWChar, 150, sExpPermanentCity), _
			Array(, adInteger, , iExpPermanentCountryID), _
			Array(, adVarWChar, 150, sExpPermanentPhone), _
			Array(, adVarWChar, 150, sExpPermanentMobile), _
			Array(, adVarWChar, 150, sExpPermanentFax), _
			Array(, adVarWChar, 150, sExpPermanentEmail), _
			Array(, adVarWChar, 255, sExpPermanentWebsite)),"-")
	End If

	' Current address
	iExpCurrentAddressID=CheckString(Request.Form("exp_cra_adrid"))
	sExpCurrentStreet=Left(CheckString(Request.Form("exp_cra_street")), 255)
	sExpCurrentPostcode=Left(CheckString(Request.Form("exp_cra_postcode")), 50)
	sExpCurrentCity=Left(CheckString(Request.Form("exp_cra_city")), 150)
	iExpCurrentCountryID=ReplaceIfEmpty(CheckString(Request.Form("exp_cra_cou")), Null)
	sExpCurrentPhone=Left(CheckString(Request.Form("exp_cra_phone")), 150)
	sExpCurrentMobile=Left(CheckString(Request.Form("exp_cra_mobile")), 150)
	sExpCurrentFax=Left(CheckString(Request.Form("exp_cra_fax")), 150)
	sExpCurrentEmail=Left(CheckString(Request.Form("exp_cra_email")), 150)
	sExpCurrentWebsite=Left(CheckString(Request.Form("exp_cra_web")), 255)

	bExpCurrentAddressExists=Len(Trim(sExpCurrentStreet)) + Len(Trim(sExpCurrentPostcode)) + Len(Trim(sExpCurrentCity)) + Len(Trim(sExpCurrentPhone)) + Len(Trim(sExpCurrentMobile)) + Len(Trim(sExpCurrentFax)) + Len(Trim(sExpCurrentEmail)) + Len(Trim(sExpCurrentWebsite))
        If IsNumeric(iExpCurrentCountryID) Then bExpCurrentAddressExists=bExpCurrentAddressExists + iExpCurrentCountryID

	If IsNumeric(iExpCurrentAddressID) And iExpCurrentAddressID>"" And iExpCurrentAddressID<>"0" Then
		objTempRs=UpdateRecordSPWithConn(objConnCustom, "usp_ExpCvvAddressUpdate", Array( _
			Array(, adVarChar, 3, "Eng"), _
			Array(, adInteger, , iExpertID), _
			Array(, adInteger, , 3), _
			Array(, adInteger, , iExpCurrentAddressID), _
			Array(, adVarWChar, 255, sExpCurrentStreet), _
			Array(, adVarWChar, 50, sExpCurrentPostcode), _
			Array(, adVarWChar, 150, sExpCurrentCity), _
			Array(, adInteger, , iExpCurrentCountryID), _
			Array(, adVarWChar, 150, sExpCurrentPhone), _
			Array(, adVarWChar, 150, sExpCurrentMobile), _
			Array(, adVarWChar, 150, sExpCurrentFax), _
			Array(, adVarWChar, 150, sExpCurrentEmail), _
			Array(, adVarWChar, 255, sExpCurrentWebsite)))

	ElseIf bExpCurrentAddressExists>0 Then
		objTempRs=InsertRecordSPWithConn(objConnCustom, "usp_ExpCvvAddressInsert", Array( _
			Array(, adVarChar, 3, "Eng"), _
			Array(, adInteger, , iExpertID), _
			Array(, adInteger, , 3), _
			Array(, adVarWChar, 255, sExpCurrentStreet), _
			Array(, adVarWChar, 50, sExpCurrentPostcode), _
			Array(, adVarWChar, 150, sExpCurrentCity), _
			Array(, adInteger, , iExpCurrentCountryID), _
			Array(, adVarWChar, 150, sExpCurrentPhone), _
			Array(, adVarWChar, 150, sExpCurrentMobile), _
			Array(, adVarWChar, 150, sExpCurrentFax), _
			Array(, adVarWChar, 150, sExpCurrentEmail), _
			Array(, adVarWChar, 255, sExpCurrentWebsite)),"-")
	End If 	
	
	If sApplicationName="expert" Then
		Response.Redirect "thankyou.asp" & sParams
	Else
		Response.Redirect "register6.asp" & sParams
	End If
End If
%>
<!--#include virtual="/_template/html.header.start.asp"-->
<script type="text/javascript" src="/_scripts/js/main.js"></script>
<script language="JavaScript">
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
<% If sApplicationName="external" Or sApplicationName="backoffice" Then %>
	f.submit();
	return;
<% End If %>
		
	if (!checkTextFieldValue(f.exp_pma_street, "", "Please fill in a street of <% =sUserSalutation %> permanent address.", 1)) { return false }
	if (!checkTextFieldValue(f.exp_pma_city, "", "Please fill in a city of <% =sUserSalutation %> permanent address.", 1)) { return false }
	if (!checkTextFieldValue(f.exp_pma_postcode, "", "Please fill in a postcode of <% =sUserSalutation %> permanent address.", 1)) { return false }
	if (!checkSelectFieldIndex(f.exp_pma_cou, 0, "Please select a country of <% =sUserSalutation %> permanent address.", 1)) { return false }

	if (!checkTextFieldValue(f.exp_pma_phone, "", "Please fill in <% =sUserSalutation %> permanent phone number.", 1)) { return false }

	if (!checkTextFieldValue(f.exp_pma_email, "", "Please specify <% =sUserSalutation %> permanent email.", 1)) { return false }
	if (!validateEmail(f.exp_pma_email.value)) {
		alert("Please retype <% =sUserSalutation %> permanent email correctly");
        f.exp_pma_email.select();        
		return;
	}
	
	if (f.exp_cra_email.value.length>0 && !validateEmail(f.exp_cra_email.value)) {
		alert("Please retype <% =sUserSalutation %> current email correctly");
        f.exp_cra_email.select();        
		return;
   }
	if (!checkTextFieldLength(f.Availability, 400, "<% =GetLabel(sCvLanguage, "Please make text of your availibility shorter") %>", 1)) { return; }

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
	<div id="content" class="searchform">
	<% 
	If Not bIsMyCV Then 
		%><h2 class="service_title">Curriculum Vitae. <span class="service_slogan">Expert ID: <% =objExpertDB.DatabaseCode %><%=iExpertID%></span></h2><br/>
		<% 
	End If
	
	ShowRegistrationProgressBar "CV", 6
	%>

<!-- Permanent address -->
	<form method="post" name="RegForm" action="<%=sScriptFileName & sParams %>">
		<div class="box search blue">
		<h3><span class="left">&nbsp;</span><span class="right">&nbsp;</span><% =GetLabel(sCvLanguage, "Permanent address") %></h3>
		<table class="search_form" width="100%" cellspacing=0 cellpadding=0 border=0>
	<% Set objTempRs=GetDataRecordsetSPWithConn(objConnCustom, "usp_ExpCvvAddressSelect", Array( _
		Array(, adInteger, , iExpertID), _
		Array(, adInteger, , 1)))

	If Not objTempRs.Eof Then 
		iExpPermanentAddressID=objTempRs("id_Address")
		sExpPermanentStreet=objTempRs("adrStreetEng")
		sExpPermanentCity=objTempRs("adrCityEng")
		sExpPermanentPostcode=objTempRs("adrPostCode")
		iExpPermanentCountryID=objTempRs("id_Country")
		sExpPermanentPhone=objTempRs("adrPhone")
		sExpPermanentMobile=objTempRs("adrMobile")
		sExpPermanentFax=objTempRs("adrFax")
		sExpPermanentEmail=objTempRs("adrEmail")
		sExpPermanentWebsite=objTempRs("adrWeb")
	End If 
	objTempRs.close

	Set objTempRs=GetDataRecordsetSPWithConn(objConnCustom, "usp_ExpCvvAddressSelect", Array( _
		Array(, adInteger, , iExpertID), _
		Array(, adInteger, , 3)))
   
	If Not objTempRs.Eof Then 
		iExpCurrentAddressID=objTempRs("id_Address")
		sExpCurrentStreet=objTempRs("adrStreetEng")
		sExpCurrentCity=objTempRs("adrCityEng")
		sExpCurrentPostcode=objTempRs("adrPostCode")
		iExpCurrentCountryID=objTempRs("id_Country")
		sExpCurrentPhone=objTempRs("adrPhone")
		sExpCurrentMobile=objTempRs("adrMobile")
		sExpCurrentFax=objTempRs("adrFax")
		sExpCurrentEmail=objTempRs("adrEmail")
		sExpCurrentWebsite=objTempRs("adrWeb")
	End If 
	objTempRs.Close 

	If CheckLength(sExpCurrentEmail)=0 Then
		If (Not (InStr(sExpPermanentEmail, sEmail)>0 Or InStr(sEmail, sExpPermanentEmail)>0)) Or CheckLength(sExpPermanentEmail)=0 Then
			sExpCurrentEmail=sEmail
		End If
	End If
	%>
	<input type="hidden" name="exp_pma_adrid" value="<%=iExpPermanentAddressID%>">
	<input type="hidden" name="exp_cra_adrid" value="<%=iExpCurrentAddressID%>">
		<tr>
		<td class="field splitter"><label for="exp_pma_street"><% =GetLabel(sCvLanguage, "Street") %></label></td>
		<td class="value blue"><input type="text" id="exp_pma_street" name="exp_pma_street" size=31 style="width:355px;" maxlength=255 value="<%=sExpPermanentStreet%>">&nbsp;&nbsp;<span class="fcmp">*</span>&nbsp;</td>
		</tr>
		<tr>
		<td class="field splitter"><label for="exp_pma_city"><% =GetLabel(sCvLanguage, "City") %></label></td>
		<td class="value blue"><input type="text" id="exp_pma_city" name="exp_pma_city" size=31 style="width:355px;" maxlength=150 value="<%=sExpPermanentCity%>">&nbsp;&nbsp;<span class="fcmp">*</span>&nbsp;</td>
		</tr>
		<tr>
		<td class="field splitter"><label for="exp_pma_postcode"><% =GetLabel(sCvLanguage, "Postcode") %></label></td>
		<td class="value blue"><input type="text" id="exp_pma_postcode" name="exp_pma_postcode" maxlength=50 size=31 style="width:355px;" value="<%=sExpPermanentPostcode%>">&nbsp;&nbsp;<span class="fcmp">*</span>&nbsp;</td>
		</tr>
		<tr>
		<td class="field splitter"><label for="exp_pma_cou"><% =GetLabel(sCvLanguage, "Country") %></label></td>
		<td class="value blue"><select id="exp_pma_cou" name="exp_pma_cou" size="1" style="width:355px;">
		<option value=" ">  <% =GetLabel(sCvLanguage, "Please select") %> &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; </option>
		<% For i=LBound(arrCountryID) To UBound(arrCountryID)-1
		If iExpPermanentCountryID=arrCountryID(i) Then
			Response.Write ("<option value="& arrCountryID(i) & " selected>"& arrCountryName(i) & "</option>")
		Else
			Response.Write ("<option value="& arrCountryID(i) & ">"& arrCountryName(i) & "</option>")
		End If
		Next %>
		</select>&nbsp;&nbsp;<span class="fcmp">*</span>&nbsp;</td>
		</tr>
		<tr>
		<td class="field splitter"><label for="exp_pma_phone"><% =GetLabel(sCvLanguage, "Phone") %></label></td>
		<td class="value blue"><input type="text" id="exp_pma_phone" name="exp_pma_phone" size=31 style="width:355px;" maxlength=150 value="<%=sExpPermanentPhone%>">&nbsp;&nbsp;<span class="fcmp">*</span>&nbsp;</td>
		</tr>
		<tr>
		<td class="field splitter"><label for="exp_pma_mobile"><% =GetLabel(sCvLanguage, "Mobile") %></label></td>
		<td class="value blue"><input type="text" id="exp_pma_mobile" name="exp_pma_mobile" size=31 style="width:355px;" maxlength=150 value="<%=sExpPermanentMobile%>"></td>
		</tr>
		<tr>
		<td class="field splitter"><label for="exp_pma_fax"><% =GetLabel(sCvLanguage, "Fax") %></label></td>
		<td class="value blue"><input type="text" id="exp_pma_fax" name="exp_pma_fax" size=31 style="width:355px;" maxlength=150 value="<%=sExpPermanentFax%>"></td>
		</tr>
		<tr>
		<td class="field splitter"><label for="exp_pma_email"><% =GetLabel(sCvLanguage, "Email") %></label></td>
		<td class="value blue"><input type="text" id="exp_pma_email" name="exp_pma_email" maxlength=150 size=31 style="width:355px;" value="<%=sExpPermanentEmail%>">&nbsp;&nbsp;<span class="fcmp">*</span>&nbsp;</td>
		</tr>
		<tr class="last">
		<td class="field splitter"><label for="exp_pma_web"><% =GetLabel(sCvLanguage, "Website") %></label></td>
		<td class="value blue"><input type="text" id="exp_pma_web" name="exp_pma_web" size=31 style="width:355px;"  maxlength=255 value="<%=sExpPermanentWebsite%>"></td>
		</tr>
		</table>
		</div><br />

		<div class="box search blue">
		<h3><span class="left">&nbsp;</span><span class="right">&nbsp;</span><% =GetLabel(sCvLanguage, "Current address (if different)") %></h3>
		<table class="search_form" width="100%" cellspacing=0 cellpadding=0 border=0>
		<tr>
		<td class="field splitter"><label for="exp_cra_street"><% =GetLabel(sCvLanguage, "Street") %></label></td>
		<td class="value blue"><input type="text" id="exp_cra_street" name="exp_cra_street" size=31 style="width:355px;" maxlength=255 value="<%=sExpCurrentStreet%>"></td>
		</tr>
		<tr>
		<td class="field splitter"><label for="exp_cra_city"><% =GetLabel(sCvLanguage, "City") %></label></td>
		<td class="value blue"><input type="text" id="exp_cra_city" name="exp_cra_city" size=31 style="width:355px;" maxlength=150 value="<%=sExpCurrentCity%>"></td>
		</tr>
		<tr>
		<td class="field splitter"><label for="exp_cra_postcode"><% =GetLabel(sCvLanguage, "Postcode") %></label></td>
		<td class="value blue"><input type="text" id="exp_cra_postcode" name="exp_cra_postcode" maxlength=50 size=31 style="width:355px;" value="<%=sExpCurrentPostcode%>"></td>
		</tr>
		<tr>
		<td class="field splitter"><label for="exp_cra_cou"><% =GetLabel(sCvLanguage, "Country") %></label></td>
		<td class="value blue"><select id="exp_cra_cou" name="exp_cra_cou" size="1" style="width:355px;">
		<option value=" ">  <% =GetLabel(sCvLanguage, "Please select") %> &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; </option>
		<% For i=LBound(arrCountryID) To UBound(arrCountryID)-1
		If iExpPermanentCountryID=arrCountryID(i) Then
			Response.Write ("<option value="& arrCountryID(i) & " selected>"& arrCountryName(i) & "</option>")
		Else
			Response.Write ("<option value="& arrCountryID(i) & ">"& arrCountryName(i) & "</option>")
		End If
		Next %>
		</select></td>
		</tr>
		<tr>
		<td class="field splitter"><label for="exp_cra_phone"><% =GetLabel(sCvLanguage, "Phone") %></label></td>
		<td class="value blue"><input type="text" id="exp_cra_phone" name="exp_cra_phone" size=31 style="width:355px;" maxlength=150 value="<%=sExpCurrentPhone%>"></td>
		</tr>
		<tr>
		<td class="field splitter"><label for="exp_cra_mobile"><% =GetLabel(sCvLanguage, "Mobile") %></label></td>
		<td class="value blue"><input type="text" id="exp_cra_mobile" name="exp_cra_mobile" size=31 style="width:355px;" maxlength=150 value="<%=sExpCurrentMobile%>"></td>
		</tr>
		<tr>
		<td class="field splitter"><label for="exp_cra_fax"><% =GetLabel(sCvLanguage, "Fax") %></label></td>
		<td class="value blue"><input type="text" id="exp_cra_fax" name="exp_cra_fax" size=31 style="width:355px;" maxlength=150 value="<%=sExpCurrentFax%>"></td>
		</tr>
		<tr>
		<td class="field splitter"><label for="exp_cra_email"><% =GetLabel(sCvLanguage, "Email") %></label></td>
		<td class="value blue"><input type="text" id="exp_cra_email" name="exp_cra_email" maxlength=150 size=31 style="width:355px;" value="<%=sExpCurrentEmail%>"></td>
		</tr>
		<tr class="last">
		<td class="field splitter"><label for="exp_cra_web"><% =GetLabel(sCvLanguage, "Website") %></label></td>
		<td class="value blue"><input type="text" id="exp_cra_web" name="exp_cra_web" size=31 style="width:355px;"  maxlength=255 value="<%=sExpCurrentWebsite%>"></td>
		</tr>
		</table>
		</div>

	<% Set objTempRs=GetDataRecordsetSPWithConn(objConnCustom, "usp_ExpCvvExpInfoSelect", Array( _
		Array(, adInteger, , iExpertID)))

	If Not objTempRs.Eof Then 

		sExpAvailability=objTempRs("expAvailabilityEng")
		iExpShortterm = objTempRs("expShortterm")
		iExpLongterm = objTempRs("expLongterm")
		sPreferences = objTempRs("expPreferences")		

		sExpOtherSkills=objTempRs("expOtherSkills")
		sExpMembership=objTempRs("expMemberProfEng")
		sExpPublications=objTempRs("expPublicationsEng")
		sExpReferences=objTempRs("expReferencesEng")
	End If
	objTempRs.Close %>

<!-- [i] please specify -->
<% ShowMessageStart "info", 580 %>
	<% =GetLabel(sCvLanguage, "Please specify availability") %>
<% ShowMessageEnd %><br />
    
<!-- Current availability -->
		<div class="box search blue">
		<h3><span class="left">&nbsp;</span><span class="right">&nbsp;</span><% =GetLabel(sCvLanguage, "Current availability") %></h3>
		<table class="search_form" width="100%" cellspacing=0 cellpadding=0 border=0>
		<tr>
		<td class="field splitter"><label for="Availability"><% =GetLabel(sCvLanguage, "Availability") %></label></td>
		<td class="value blue"><textarea cols="31" style="width:355px;" id="Availability" name="Availability" rows=4 wrap="yes"><%=sExpAvailability%></textarea></td>
		</tr>
		</table>
		</div>


<!-- [i] please select -->
<% ShowMessageStart "info", 580 %>
	<% =GetLabel(sCvLanguage, "Please state your preferences") %>
<% ShowMessageEnd %><br />


<!-- Assignment Preferences -->
		<div class="box search blue">
		<h3><span class="left">&nbsp;</span><span class="right">&nbsp;</span><% =GetLabel(sCvLanguage, "Assignment preferences") %></h3>
		<table class="search_form" width="100%" cellspacing=0 cellpadding=0 border=0>
		<tr>
		<td class="field splitter"><label><% =GetLabel(sCvLanguage, "Preferences") %></label></td>
		<td class="value blue"><input type="checkbox" name="shortterm" value="1" <% if iExpShortterm=1 then %> checked <%end if %> > <% =GetLabel(sCvLanguage, "Short-term missions") %> &nbsp;<input type="checkbox" name="longterm" value="1" <% if iExpLongterm=1 then %> checked <%end if %>> <% =GetLabel(sCvLanguage, "Long-term missions") %></td>
		</tr>
		<tr>
		<td class="field splitter"><label><% =GetLabel(sCvLanguage, "Other preferences") %> (<% =GetLabel(sCvLanguage, "location") %>, etc.)</label></td>
		<td class="value blue"><textarea cols="31" style="width:355px;" name="preferences" rows=4 wrap="yes"><%=sPreferences%></textarea></td>
		</tr>
		</table>
		</div><br />

<!-- Other section -->
		<div class="box search blue">
		<h3><span class="left">&nbsp;</span><span class="right">&nbsp;</span><% =GetLabel(sCvLanguage, "Miscellaneous") %></h3>
		<table class="search_form" width="100%" cellspacing=0 cellpadding=0 border=0>
		<tr>
		<td class="field splitter"><label><% =GetLabel(sCvLanguage, "Membership in professional bodies") %></label></td>
		<td class="value blue"><textarea cols="31" style="width:355px;" name="memberships" rows=4 wrap="yes"><%=sExpMembership%></textarea></td>
		</tr>
		<tr>
		<td class="field splitter"><label><% =GetLabel(sCvLanguage, "Publications") %></label></td>
		<td class="value blue"><textarea cols="31" style="width:355px;" name="publications" rows=6 wrap="yes"><%=sExpPublications%></textarea></td>
		</tr>
		<tr>
		<td class="field splitter"><label><% =GetLabel(sCvLanguage, "References") %></label></td>
		<td class="value blue"><textarea cols="31" style="width:355px;" name="references" rows=4 wrap="yes"><%=sExpReferences%></textarea></td>
		</tr>
		</table>
		</div>

		<div class="spacebottom">
		<input type="image" src="<% =sHomePath %>image/bte_savecont.gif" class="button first" name="Continue" alt="Save & Continue" border="0" onClick="validateForm(); return false;">
		</div>
		</form>

	</div>
</div>
	<!-- footer -->
	<!--#include file="../_template/page.footer.asp"-->
<% CloseDBConnection %>
</body>
<!--#include file="../_template/html.footer.asp"-->
