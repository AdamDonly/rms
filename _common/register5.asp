<% 
'--------------------------------------------------------------------
'
' CV registration.
' Full format. Addresses and availability.
'
'--------------------------------------------------------------------
Response.Expires = 0
Response.ExpiresAbsolute = Now() - 1
Response.CacheControl = "no-cache" 
%>
<!--#include file="_data/datCountry.asp"-->
<% 
If sApplicationName<>"expert" Then
	CheckUserLogin sScriptFullNameAsParams
End If
CheckExpertID()
' The page is accessible only if expert id is valid
If iExpertID<=0 Then Response.Redirect(sApplicationHomePath & "register.asp" & sParams)

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

	objTempRs=UpdateRecordSP("usp_ExpCvvAvailabilityUpdate", Array( _
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
		objTempRs=UpdateRecordSP("usp_ExpCvvAddressUpdate", Array( _
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
		objTempRs=InsertRecordSP("usp_ExpCvvAddressInsert", Array( _
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
		objTempRs=UpdateRecordSP("usp_ExpCvvAddressUpdate", Array( _
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
		objTempRs=InsertRecordSP("usp_ExpCvvAddressInsert", Array( _
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
<html>
<head>
<title><% =GetLabel(sCvLanguage, "CV registration") %>. <% =GetLabel(sCvLanguage, "Contact details &amp; Availability") %></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" type="text/css" href="<% =sHomePath %>styles.css">
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
<% If sApplicationName="external" Then %>
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

<body bgcolor="#FFFFFF" topmargin=0 leftmargin=0  marginheight=0 marginwidth=0>
<% ShowTopMenu %>
<% ShowRegistrationProgressBar "CV", 6 %>

<!-- Permanent address -->
	<br><table width=580 cellspacing=0 cellpadding=0 border=0 align="center">
	<form method="post" name="RegForm" action="<%=sScriptFileName & sParams %>">
 
	<tr height=1><td width=1 bgcolor="<%=colFormHeaderTop%>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td>
	<td width=578 bgcolor="<%=colFormHeaderTop%>"><img src="<% =sHomePath %>image/x.gif" width=498 height=1></td>
	<td width=1 bgcolor="<%=colFormHeaderBottom%>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td></tr>
	<tr height=18><td width=1 bgcolor="<%=colFormHeaderTop%>"><img src="<% =sHomePath %>image/x.gif" width=1 height=18></td>
	<td width=578 bgcolor="<%=colFormHeaderMain%>"><img src="<% =sHomePath %>image/x.gif" width=498 height=1><br><p class="fttl"><img src="<% =sHomePath %>image/<% =imgFormBullet %>" width=7 height=7 align=left vspace=3 hspace=8><% =GetLabel(sCvLanguage, "PERMANENT ADDRESS") %></p></td>
	<td width=1 bgcolor="<%=colFormHeaderBottom%>"><img src="<% =sHomePath %>image/x.gif" width=1 height=18></td></tr>
	<tr height=1><td width=1 bgcolor="<%=colFormHeaderTop%>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td>
	<td width=578 bgcolor="<%=colFormHeaderBottom%>"><img src="<% =sHomePath %>image/x.gif" width=498 height=1></td>
	<td width=1 bgcolor="<%=colFormHeaderBottom%>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td></tr>
	<tr height=1><td width=1 bgcolor="<% =colFormHeaderSplitter %>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td>
	<td width=578 bgcolor="<% =colFormHeaderSplitter %>"><img src="<% =sHomePath %>image/x.gif" width=498 height=1></td>
	<td width=1 bgcolor="<% =colFormHeaderSplitter %>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td></tr>

	<% Set objTempRs=GetDataRecordsetSP("usp_ExpCvvAddressSelect", Array( _
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

	Set objTempRs=GetDataRecordsetSP("usp_ExpCvvAddressSelect", Array( _
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
		If (Not (InStr(sExpPermanentEmail, sUserEmail)>0 Or InStr(sUserEmail, sExpPermanentEmail)>0)) Or CheckLength(sExpPermanentEmail)=0 Then
			sExpCurrentEmail=sUserEmail
		End If
	End If
	%>

	<tr><td width=1 bgcolor="<%=colFormBodyText%>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td>
	<td width=578 bgcolor="<%=colFormHeaderTop%>" valign="top">
		<table width="100%" cellspacing=0 cellpadding=0 border=0>
		<tr><td width=170><p class="ftxt"><% =GetLabel(sCvLanguage, "Street") %></td>
		<td bgcolor="<%=colFormBodyRight%>" width=1><img src="x.gif" width=1 height=24></td>
		<td bgcolor="<%=colFormBodyText%>" width=407><img src="x.gif" width=1 height=6><br>
		&nbsp;&nbsp;<input type="text" name="exp_pma_street" size=31 style="width:355px;" maxlength=255 value="<%=sExpPermanentStreet%>">&nbsp;&nbsp;<span class="fcmp">*</span>&nbsp;&nbsp;</td></tr>

		<tr><td width=170><p class="ftxt"><% =GetLabel(sCvLanguage, "City") %></td>
		<td bgcolor="<%=colFormBodyRight%>" width=1><img src="x.gif" width=1 height=24></td>
		<td bgcolor="<%=colFormBodyText%>" width=407>&nbsp;&nbsp;<input type="text" name="exp_pma_city" size=31 style="width:355px;" maxlength=150 value="<%=sExpPermanentCity%>">&nbsp;&nbsp;<span class="fcmp">*</span>&nbsp;&nbsp;</td></tr>

		<tr><td width=170><p class="ftxt"><% =GetLabel(sCvLanguage, "Postcode") %></td>
		<td bgcolor="<%=colFormBodyRight%>" width=1><img src="x.gif" width=1 height=24></td>
		<td bgcolor="<%=colFormBodyText%>" width=407>&nbsp;&nbsp;<input type="text" name="exp_pma_postcode" maxlength=50 size=31 style="width:355px;" value="<%=sExpPermanentPostcode%>">&nbsp;&nbsp;<span class="fcmp">*</span>&nbsp;&nbsp;</td></tr>

		<tr><td width=170><p class="ftxt"><% =GetLabel(sCvLanguage, "Country") %></td>
		<td bgcolor="<%=colFormBodyRight%>" width=1><img src="x.gif" width=1 height=24></td>
		<td bgcolor="<%=colFormBodyText%>" width=407>
		&nbsp;&nbsp;<select name="exp_pma_cou" size="1" style="width:355px;">
		<option value=" ">  <% =GetLabel(sCvLanguage, "Please select") %> &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; </option>
		<% For i=LBound(arrCountryID) To UBound(arrCountryID)-1
		If iExpPermanentCountryID=arrCountryID(i) Then
			Response.Write ("<option value="& arrCountryID(i) & " selected>"& arrCountryName(i) & "</option>")
		Else
			Response.Write ("<option value="& arrCountryID(i) & ">"& arrCountryName(i) & "</option>")
		End If
		Next %>
		</select>&nbsp;&nbsp;<span class="fcmp">*</span>&nbsp;&nbsp;</td></tr>

		<tr><td width=170><p class="ftxt"><% =GetLabel(sCvLanguage, "Phone") %></td>
		<td bgcolor="<%=colFormBodyRight%>" width=1><img src="x.gif" width=1 height=24></td>
		<td bgcolor="<%=colFormBodyText%>" width=407>&nbsp;&nbsp;<input type="text" name="exp_pma_phone" size=31 style="width:355px;" maxlength=150 value="<%=sExpPermanentPhone%>">&nbsp;&nbsp;<span class="fcmp">*</span>&nbsp;&nbsp;</td></tr>

		<tr><td width=170><p class="ftxt"><% =GetLabel(sCvLanguage, "Mobile") %></td>
		<td bgcolor="<%=colFormBodyRight%>" width=1><img src="x.gif" width=1 height=24></td>
		<td bgcolor="<%=colFormBodyText%>" width=407>&nbsp;&nbsp;<input type="text" name="exp_pma_mobile" size=31 style="width:355px;" maxlength=150 value="<%=sExpPermanentMobile%>"></td></tr>

		<tr><td width=170><p class="ftxt"><% =GetLabel(sCvLanguage, "Fax") %></td>
		<td bgcolor="<%=colFormBodyRight%>" width=1><img src="x.gif" width=1 height=24></td>
		<td bgcolor="<%=colFormBodyText%>" width=407>&nbsp;&nbsp;<input type="text" name="exp_pma_fax" size=31 style="width:355px;" maxlength=150 value="<%=sExpPermanentFax%>"></td></tr>

		<tr><td width=170><p class="ftxt"><% =GetLabel(sCvLanguage, "Email") %></td>
		<td bgcolor="<%=colFormBodyRight%>" width=1><img src="x.gif" width=1 height=24></td>
		<td bgcolor="<%=colFormBodyText%>" width=407>&nbsp;&nbsp;<input type="text" name="exp_pma_email" maxlength=150 size=31 style="width:355px;" value="<%=sExpPermanentEmail%>">&nbsp;&nbsp;<span class="fcmp">*</span>&nbsp;&nbsp;</td></tr>

		<tr><td width=170><p class="ftxt"><% =GetLabel(sCvLanguage, "Website") %></td>
		<td bgcolor="<%=colFormBodyRight%>" width=1><img src="x.gif" width=1 height=24></td>
		<td bgcolor="<%=colFormBodyText%>" width=407>&nbsp;&nbsp;<input type="text" name="exp_pma_web" size=31 style="width:355px;"  maxlength=255 value="<%=sExpPermanentWebsite%>">

		<img src="<% =sHomePath %>image/x.gif" width=358 height=1 vspace=3><br>
		</td></tr>
		</table>
	</td>
	<td width=1 bgcolor="<%=colFormBodyRight%>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td></tr>

	<tr height=1><td width=1 bgcolor="<%=colFormBodyText%>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td>
	<td width=578 bgcolor="<%=colFormHeaderMain%>"><img src="<% =sHomePath %>image/x.gif" width=498 height=1></td>
	<td width=1 bgcolor="<%=colFormBodyRight%>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td></tr>
	<tr height=1><td width=1 bgcolor="<% =colFormHeaderSplitter %>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td>
	<td width=578 bgcolor="<% =colFormHeaderSplitter %>"><img src="<% =sHomePath %>image/x.gif" width=498 height=1></td>
	<td width=1 bgcolor="<% =colFormHeaderSplitter %>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td></tr>
	</table><br>

	<input type="hidden" name="exp_pma_adrid" value="<%=iExpPermanentAddressID%>">


  <!-- Current address -->
	<table width=580 cellspacing=0 cellpadding=0 border=0 align="center">
	<tr height=1><td width=1 bgcolor="<%=colFormHeaderTop%>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td>
	<td width=578 bgcolor="<%=colFormHeaderTop%>"><img src="<% =sHomePath %>image/x.gif" width=498 height=1></td>
	<td width=1 bgcolor="<%=colFormHeaderBottom%>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td></tr>
	<tr height=18><td width=1 bgcolor="<%=colFormHeaderTop%>"><img src="<% =sHomePath %>image/x.gif" width=1 height=18></td>
	<td width=578 bgcolor="<%=colFormHeaderMain%>"><img src="<% =sHomePath %>image/x.gif" width=498 height=1><br><p class="fttl"><img src="<% =sHomePath %>image/<% =imgFormBullet %>" width=7 height=7 align=left vspace=3 hspace=8><% =GetLabel(sCvLanguage, "CURRENT ADDRESS (if different)") %></p></td>
	<td width=1 bgcolor="<%=colFormHeaderBottom%>"><img src="<% =sHomePath %>image/x.gif" width=1 height=18></td></tr>
	<tr height=1><td width=1 bgcolor="<%=colFormHeaderTop%>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td>
	<td width=578 bgcolor="<%=colFormHeaderBottom%>"><img src="<% =sHomePath %>image/x.gif" width=498 height=1></td>
	<td width=1 bgcolor="<%=colFormHeaderBottom%>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td></tr>
	<tr height=1><td width=1 bgcolor="<% =colFormHeaderSplitter %>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td>
	<td width=578 bgcolor="<% =colFormHeaderSplitter %>"><img src="<% =sHomePath %>image/x.gif" width=498 height=1></td>
	<td width=1 bgcolor="<% =colFormHeaderSplitter %>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td></tr>

	<tr><td width=1 bgcolor="<%=colFormBodyText%>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td>
	<td width=578 bgcolor="<%=colFormHeaderTop%>" valign="top">
		<table width="100%" cellspacing=0 cellpadding=0 border=0>
		<tr><td width=170><p class="ftxt"><% =GetLabel(sCvLanguage, "Street") %></td>
		<td bgcolor="<%=colFormBodyRight%>" width=1><img src="x.gif" width=1 height=24></td>
		<td bgcolor="<%=colFormBodyText%>" width=407><img src="x.gif" width=1 height=6><br>
		&nbsp;&nbsp;<input type="text" name="exp_cra_street" size=31 style="width:355px;" maxlength=255 value="<%=sExpCurrentStreet%>">&nbsp;&nbsp;</td></tr>

		<tr><td width=170><p class="ftxt"><% =GetLabel(sCvLanguage, "City") %></td>
		<td bgcolor="<%=colFormBodyRight%>" width=1><img src="x.gif" width=1 height=24></td>
		<td bgcolor="<%=colFormBodyText%>" width=407>&nbsp;&nbsp;<input type="text" name="exp_cra_city" size=31 style="width:355px;" maxlength=150 value="<%=sExpCurrentCity%>">&nbsp;&nbsp;</td></tr>

		<tr><td width=170><p class="ftxt"><% =GetLabel(sCvLanguage, "Postcode") %></td>
		<td bgcolor="<%=colFormBodyRight%>" width=1><img src="x.gif" width=1 height=24></td>
		<td bgcolor="<%=colFormBodyText%>" width=407>&nbsp;&nbsp;<input type="text" name="exp_cra_postcode" maxlength=50 size=31 style="width:355px;" value="<%=sExpCurrentPostcode%>">&nbsp;&nbsp;</td></tr>

		<tr><td width=170><p class="ftxt"><% =GetLabel(sCvLanguage, "Country") %></td>
		<td bgcolor="<%=colFormBodyRight%>" width=1><img src="x.gif" width=1 height=24></td>
		<td bgcolor="<%=colFormBodyText%>" width=407>
		&nbsp;&nbsp;<select name="exp_cra_cou" size="1" style="width:355px;">
		<option value=""> <% =GetLabel(sCvLanguage, "Please select") %> &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; </option>
		<% For i=LBound(arrCountryID) To UBound(arrCountryID)-1
		If iExpCurrentCountryID=arrCountryID(i) Then
			Response.Write ("<option value="& arrCountryID(i) & " selected>"& arrCountryName(i) & "</option>")
		Else
			Response.Write ("<option value="& arrCountryID(i) & ">"& arrCountryName(i) & "</option>")
		End If
		Next %>
		</select>&nbsp;&nbsp;</td></tr>

		<tr><td width=170><p class="ftxt"><% =GetLabel(sCvLanguage, "Phone") %></td>
		<td bgcolor="<%=colFormBodyRight%>" width=1><img src="x.gif" width=1 height=24></td>
		<td bgcolor="<%=colFormBodyText%>" width=407>&nbsp;&nbsp;<input type="text" name="exp_cra_phone" size=31 style="width:355px;" maxlength=150 value="<%=sExpCurrentPhone%>">&nbsp;&nbsp;</td></tr>

		<tr><td width=170><p class="ftxt"><% =GetLabel(sCvLanguage, "Mobile") %></td>
		<td bgcolor="<%=colFormBodyRight%>" width=1><img src="x.gif" width=1 height=24></td>
		<td bgcolor="<%=colFormBodyText%>" width=407>&nbsp;&nbsp;<input type="text" name="exp_cra_mobile" size=31 style="width:355px;" maxlength=150 value="<%=sExpCurrentMobile%>"></td></tr>

		<tr><td width=170><p class="ftxt"><% =GetLabel(sCvLanguage, "Fax") %></td>
		<td bgcolor="<%=colFormBodyRight%>" width=1><img src="x.gif" width=1 height=24></td>
		<td bgcolor="<%=colFormBodyText%>" width=407>&nbsp;&nbsp;<input type="text" name="exp_cra_fax" size=31 style="width:355px;" maxlength=150 value="<%=sExpCurrentFax%>"></td></tr>

		<tr><td width=170><p class="ftxt"><% =GetLabel(sCvLanguage, "Email") %></td>
		<td bgcolor="<%=colFormBodyRight%>" width=1><img src="x.gif" width=1 height=24></td>
		<td bgcolor="<%=colFormBodyText%>" width=407>&nbsp;&nbsp;<input type="text" name="exp_cra_email" maxlength=150 size=31 style="width:355px;" value="<%=sExpCurrentEmail%>"></td></tr>

		<tr><td width=170><p class="ftxt"><% =GetLabel(sCvLanguage, "Website") %></td>
		<td bgcolor="<%=colFormBodyRight%>" width=1><img src="x.gif" width=1 height=24></td>
		<td bgcolor="<%=colFormBodyText%>" width=407>&nbsp;&nbsp;<input type="text" name="exp_cra_web" size=31 style="width:355px;"  maxlength=255 value="<%=sExpCurrentWebsite%>">

		<img src="<% =sHomePath %>image/x.gif" width=358 height=1 vspace=3><br>
		</td></tr>
		</table>
	</td>
	<td width=1 bgcolor="<%=colFormBodyRight%>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td></tr>
	<input type="hidden" name="exp_cra_adrid" value="<%=iExpCurrentAddressID%>">

	<tr height=1><td width=1 bgcolor="<%=colFormBodyText%>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td>
	<td width=578 bgcolor="<%=colFormHeaderMain%>"><img src="<% =sHomePath %>image/x.gif" width=498 height=1></td>
	<td width=1 bgcolor="<%=colFormBodyRight%>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td></tr>
	<tr height=1><td width=1 bgcolor="<% =colFormHeaderSplitter %>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td>
	<td width=578 bgcolor="<% =colFormHeaderSplitter %>"><img src="<% =sHomePath %>image/x.gif" width=498 height=1></td>
	<td width=1 bgcolor="<% =colFormHeaderSplitter %>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td></tr>
	</table><br><br>


<!-- Red horisontal line -->
	<table width=100% cellspacing=0 cellpadding=0 border=0 align="center">
	<tr height=2><td bgcolor="<%=colFormHeaderMain%>"><img src="<% =sHomePath %>image/x.gif" width=600 height=2><br></td></tr>
	</table><br>

	<% Set objTempRs=GetDataRecordsetSP("usp_ExpCvvExpInfoSelect", Array( _
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
<% ShowMessageEnd %>
    
<!-- Current availability -->
	<table width=580 cellspacing=0 cellpadding=0 border=0 align="center">
	<tr height=1><td width=1 bgcolor="<%=colFormHeaderTop%>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td>
	<td width=578 bgcolor="<%=colFormHeaderTop%>"><img src="<% =sHomePath %>image/x.gif" width=498 height=1></td>
	<td width=1 bgcolor="<%=colFormHeaderBottom%>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td></tr>
	<tr height=18><td width=1 bgcolor="<%=colFormHeaderTop%>"><img src="<% =sHomePath %>image/x.gif" width=1 height=18></td>
	<td width=578 bgcolor="<%=colFormHeaderMain%>"><img src="<% =sHomePath %>image/x.gif" width=498 height=1><br><p class="fttl"><img src="<% =sHomePath %>image/<% =imgFormBullet %>" width=7 height=7 align=left vspace=3 hspace=8><% =GetLabel(sCvLanguage, "CURRENT AVAILABILITY") %></p></td>
	<td width=1 bgcolor="<%=colFormHeaderBottom%>"><img src="<% =sHomePath %>image/x.gif" width=1 height=18></td></tr>
	<tr height=1><td width=1 bgcolor="<%=colFormHeaderTop%>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td>
	<td width=578 bgcolor="<%=colFormHeaderBottom%>"><img src="<% =sHomePath %>image/x.gif" width=498 height=1></td>
	<td width=1 bgcolor="<%=colFormHeaderBottom%>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td></tr>
	<tr height=1><td width=1 bgcolor="<% =colFormHeaderSplitter %>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td>
	<td width=578 bgcolor="<% =colFormHeaderSplitter %>"><img src="<% =sHomePath %>image/x.gif" width=498 height=1></td>
	<td width=1 bgcolor="<% =colFormHeaderSplitter %>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td></tr>

	<tr><td width=1 bgcolor="<%=colFormBodyText%>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td>
	<td width=578 bgcolor="<%=colFormHeaderTop%>" valign="top">
		<table width="100%" cellspacing=0 cellpadding=0 border=0>
		<tr><td width=170><p class="ftxt"><% =GetLabel(sCvLanguage, "Availability") %></td>
		<td bgcolor="<%=colFormBodyRight%>" width=1><img src="x.gif" width=1 height=24></td>
		<td bgcolor="<%=colFormBodyText%>" width=407><img src="x.gif" width=1 height=6><br>
		&nbsp;&nbsp;<textarea cols="31" style="width:355px;" name="Availability" rows=4 wrap="yes"><%=sExpAvailability%></textarea>

		<img src="<% =sHomePath %>image/x.gif" width=358 height=1 vspace=3><br>
		</td></tr>
		</table>
	</td>
	<td width=1 bgcolor="<%=colFormBodyRight%>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td></tr>

	<tr height=1><td width=1 bgcolor="<%=colFormBodyText%>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td>
	<td width=578 bgcolor="<%=colFormHeaderMain%>"><img src="<% =sHomePath %>image/x.gif" width=498 height=1></td>
	<td width=1 bgcolor="<%=colFormBodyRight%>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td></tr>
	<tr height=1><td width=1 bgcolor="<% =colFormHeaderSplitter %>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td>
	<td width=578 bgcolor="<% =colFormHeaderSplitter %>"><img src="<% =sHomePath %>image/x.gif" width=498 height=1></td>
	<td width=1 bgcolor="<% =colFormHeaderSplitter %>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td></tr>
	</table><br>


<!-- [i] please select -->
<% ShowMessageStart "info", 580 %>
	<% =GetLabel(sCvLanguage, "Please state your preferences") %>
<% ShowMessageEnd %>


<!-- Assignment Preferences -->
	<table width=580 cellspacing=0 cellpadding=0 border=0 align="center">
	<tr height=1><td width=1 bgcolor="<%=colFormHeaderTop%>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td>
	<td width=578 bgcolor="<%=colFormHeaderTop%>"><img src="<% =sHomePath %>image/x.gif" width=498 height=1></td>
	<td width=1 bgcolor="<%=colFormHeaderBottom%>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td></tr>
	<tr height=18><td width=1 bgcolor="<%=colFormHeaderTop%>"><img src="<% =sHomePath %>image/x.gif" width=1 height=18></td>
	<td width=578 bgcolor="<%=colFormHeaderMain%>"><img src="<% =sHomePath %>image/x.gif" width=498 height=1><br><p class="fttl"><img src="<% =sHomePath %>image/<% =imgFormBullet %>" width=7 height=7 align=left vspace=3 hspace=8><% =GetLabel(sCvLanguage, "ASSIGNMENT PREFERENCES") %></p></td>
	<td width=1 bgcolor="<%=colFormHeaderBottom%>"><img src="<% =sHomePath %>image/x.gif" width=1 height=18></td></tr>
	<tr height=1><td width=1 bgcolor="<%=colFormHeaderTop%>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td>
	<td width=578 bgcolor="<%=colFormHeaderBottom%>"><img src="<% =sHomePath %>image/x.gif" width=498 height=1></td>
	<td width=1 bgcolor="<%=colFormHeaderBottom%>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td></tr>
	<tr height=1><td width=1 bgcolor="<% =colFormHeaderSplitter %>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td>
	<td width=578 bgcolor="<% =colFormHeaderSplitter %>"><img src="<% =sHomePath %>image/x.gif" width=498 height=1></td>
	<td width=1 bgcolor="<% =colFormHeaderSplitter %>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td></tr>

	<tr><td width=1 bgcolor="<%=colFormBodyText%>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td>
	<td width=578 bgcolor="<%=colFormHeaderTop%>" valign="top">
		<table width="100%" cellspacing=0 cellpadding=0 border=0>
		<tr><td width=170><p class="ftxt"><% =GetLabel(sCvLanguage, "Preferences") %></td>
		<td bgcolor="<%=colFormBodyRight%>" width=1><img src="x.gif" width=1 height=24></td>
		<td bgcolor="<%=colFormBodyText%>" width=407><p><img src="x.gif" width=1 height=6><br>
		&nbsp;&nbsp;<input type="checkbox" name="shortterm" value="1" <% if iExpShortterm=1 then %> checked <%end if %> > <% =GetLabel(sCvLanguage, "Short-term missions") %> &nbsp;<input type="checkbox" name="longterm" value="1" <% if iExpLongterm=1 then %> checked <%end if %>> <% =GetLabel(sCvLanguage, "Long-term missions") %>
		</td></tr>
		<tr><td width=170><p class="ftxt"><% =GetLabel(sCvLanguage, "Other preferences") %> (<% =GetLabel(sCvLanguage, "location") %>, etc.)</td>
		<td bgcolor="<%=colFormBodyRight%>" width=1><img src="x.gif" width=1 height=24></td>
		<td bgcolor="<%=colFormBodyText%>" width=407><img src="x.gif" width=1 height=6><br>
		&nbsp;&nbsp;<textarea cols="31" style="width:355px;" name="preferences" rows=4 wrap="yes"><%=sPreferences%></textarea>
		<img src="<% =sHomePath %>image/x.gif" width=358 height=1 vspace=3><br>
		</td></tr>
		</table>
	</td>
	<td width=1 bgcolor="<%=colFormBodyRight%>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td></tr>

	<tr height=1><td width=1 bgcolor="<%=colFormBodyText%>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td>
	<td width=578 bgcolor="<%=colFormHeaderMain%>"><img src="<% =sHomePath %>image/x.gif" width=498 height=1></td>
	<td width=1 bgcolor="<%=colFormBodyRight%>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td></tr>
	<tr height=1><td width=1 bgcolor="<% =colFormHeaderSplitter %>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td>
	<td width=578 bgcolor="<% =colFormHeaderSplitter %>"><img src="<% =sHomePath %>image/x.gif" width=498 height=1></td>
	<td width=1 bgcolor="<% =colFormHeaderSplitter %>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td></tr>
	</table><br><br>

<!-- Other section -->

	<table width=580 cellspacing=0 cellpadding=0 border=0 align="center">
	<tr height=1><td width=1 bgcolor="<%=colFormHeaderTop%>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td>
	<td width=578 bgcolor="<%=colFormHeaderTop%>"><img src="<% =sHomePath %>image/x.gif" width=498 height=1></td>
	<td width=1 bgcolor="<%=colFormHeaderBottom%>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td></tr>
	<tr height=18><td width=1 bgcolor="<%=colFormHeaderTop%>"><img src="<% =sHomePath %>image/x.gif" width=1 height=18></td>
	<td width=578 bgcolor="<%=colFormHeaderMain%>"><img src="<% =sHomePath %>image/x.gif" width=498 height=1><br><p class="fttl"><img src="<% =sHomePath %>image/<% =imgFormBullet %>" width=7 height=7 align=left vspace=3 hspace=8><% =GetLabel(sCvLanguage, "MISCELLANEOUS") %></p></td>
	<td width=1 bgcolor="<%=colFormHeaderBottom%>"><img src="<% =sHomePath %>image/x.gif" width=1 height=18></td></tr>
	<tr height=1><td width=1 bgcolor="<%=colFormHeaderTop%>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td>
	<td width=578 bgcolor="<%=colFormHeaderBottom%>"><img src="<% =sHomePath %>image/x.gif" width=498 height=1></td>
	<td width=1 bgcolor="<%=colFormHeaderBottom%>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td></tr>
	<tr height=1><td width=1 bgcolor="<% =colFormHeaderSplitter %>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td>
	<td width=578 bgcolor="<% =colFormHeaderSplitter %>"><img src="<% =sHomePath %>image/x.gif" width=498 height=1></td>
	<td width=1 bgcolor="<% =colFormHeaderSplitter %>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td></tr>
  
	<tr><td width=1 bgcolor="<%=colFormBodyText%>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td>
	<td width=578 bgcolor="<%=colFormHeaderTop%>" valign="top">
		<table width="100%" cellspacing=0 cellpadding=0 border=0>
		<tr><td width=170><p class="ftxt"><% =GetLabel(sCvLanguage, "Other skills") %></td>
		<td bgcolor="<%=colFormBodyRight%>" width=1><img src="x.gif" width=1 height=24></td>
		<td bgcolor="<%=colFormBodyText%>" width=407><img src="x.gif" width=1 height=6><br>
		&nbsp;&nbsp;<textarea cols="31" style="width:355px;" name="otherskills" rows=4 wrap="yes"><%=sExpOtherSkills%></textarea></td></tr>
		<tr><td width=170><p class="ftxt"><% =GetLabel(sCvLanguage, "Membership in professional bodies") %></td>
		<td bgcolor="<%=colFormBodyRight%>" width=1><img src="x.gif" width=1 height=24></td>
		<td bgcolor="<%=colFormBodyText%>" width=407>
		&nbsp;&nbsp;<textarea cols="31" style="width:355px;" name="memberships" rows=4 wrap="yes"><%=sExpMembership%></textarea></td></tr>
		<tr><td width=170><p class="ftxt"><% =GetLabel(sCvLanguage, "Publications") %></td>
		<td bgcolor="<%=colFormBodyRight%>" width=1><img src="x.gif" width=1 height=24></td>
		<td bgcolor="<%=colFormBodyText%>" width=407>&nbsp;&nbsp;<textarea cols="31" style="width:355px;" name="publications" rows=6 wrap="yes"><%=sExpPublications%></textarea></td></tr>
		<tr><td width=170><p class="ftxt"><% =GetLabel(sCvLanguage, "References") %></td>
		<td bgcolor="<%=colFormBodyRight%>" width=1><img src="x.gif" width=1 height=24></td>
		<td bgcolor="<%=colFormBodyText%>" width=407>&nbsp;&nbsp;<textarea cols="31" style="width:355px;" name="references" rows=4 wrap="yes"><%=sExpReferences%></textarea>
		<img src="<% =sHomePath %>image/x.gif" width=358 height=1 vspace=3><br>
		</td></tr>
		</table>
	</td>
	<td width=1 bgcolor="<%=colFormBodyRight%>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td></tr>

	<tr height=1><td width=1 bgcolor="<%=colFormBodyText%>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td>
	<td width=578 bgcolor="<%=colFormHeaderMain%>"><img src="<% =sHomePath %>image/x.gif" width=498 height=1></td>
	<td width=1 bgcolor="<%=colFormBodyRight%>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td></tr>
	<tr height=1><td width=1 bgcolor="<% =colFormHeaderSplitter %>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td>
	<td width=578 bgcolor="<% =colFormHeaderSplitter %>"><img src="<% =sHomePath %>image/x.gif" width=498 height=1></td>
	<td width=1 bgcolor="<% =colFormHeaderSplitter %>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td></tr>
	</table><br>


	<table width=576 cellspacing=0 cellpadding=0 border=0 align="center">
	<tr>
	<td width="100%" align="left">
	<img src="<% =sHomePath %>image/x.gif" width=170 height=1><input type="image" src="<% =sHomePath %>image/bte_savecont.gif" name="Continue" alt="Save & Continue" border="0" onClick="validateForm(); return false;">
	</td>
	</tr>
	</form>
	</table><br>

<% CloseDBConnection %>
</body>
</html>
