<%
'--------------------------------------------------------------------
'
' Expert's CV. View in AFB format
'
'--------------------------------------------------------------------
%>
<!--#include file="cv_data.asp"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" type="text/css" href="<% =sHomePath %>styles.css">
<title>CV in African Development Bank format</title>
</head>

<body bgcolor="#FFFFFF" topmargin=0 leftmargin=0 marginheight=0 marginwidth=0>
<% ShowTopMenu %>

<p class="ttl" align="center"><b>
<% If iCvID=4238 Then %>
	Sample Curriculum Vitae* - African Development Bank format
<% Else %>
        Curriculum Vitae - African Development Bank format
<% End If %>
</b></p>

<table width="98%" cellpadding=0 cellspacing=0 border=0 align="center">
<tr><td width="85%" valign="top">
<br />

<%
' Personal information
	WriteTableHeader
	WriteDataRow "<b>" & GetLabel(sCvLanguage, "Proposed Position") & "</b>:", ""
	WriteDataRow "<b>" & GetLabel(sCvLanguage, "Name of Firm") & "</b>:", ""
	WriteDataRow "<b>" & GetLabel(sCvLanguage, "Name of Staff") & "</b>:", sFullName
	WriteDataRow "<b>" & GetLabel(sCvLanguage, "Profession") & "</b>:", sProfession
	WriteDataRow "<b>" & GetLabel(sCvLanguage, "Date of Birth") & "</b>:", ConvertDateForText(sBirthDate, " ", "DayMonthYear") 
	WriteDataRow "<b>" & GetLabel(sCvLanguage, "Years with Firm") & "</b>:", ""
	WriteDataRow "<b>" & GetLabel(sCvLanguage, "Nationality") & "</b>:", sNationality 
	WriteSpaceRow

	WriteDataRow "<b>" & GetLabel(sCvLanguage, "Membership in Professional Societies") & "</b>:", sMemberships
	WriteSpaceRow
	WriteDataRow "<b>" & GetLabel(sCvLanguage, "Detailed Tasks Assigned (AFB)") & "</b>:", ""
	WriteSpaceRow
	WriteDataRow "<b>" & GetLabel(sCvLanguage, "Key Qualifications (AFD)") & "</b>:", ConvertText(sKeyQualification)
	WriteTableFooter

' Education
%>
	<table cellspacing=0 cellpadding=0 align="center" width="97%">
	<tr><td width="100%" bgcolor="#FFFFFF"><p class="txt"><b><% =GetLabel(sCvLanguage, "Education") %></b>:</td></tr>
	</table>
<%
	WriteGridTableHeader
	WriteGridDataRow 4, "25%#|#13%#|#13%#|#50%", "<b>" & GetLabel(sCvLanguage, "Institution") & "</b>#|#<b>" & GetLabel(sCvLanguage, "Start date") & "</b>#|#<b>" & GetLabel(sCvLanguage, "End date") & "</b>#|#<b>" & GetLabel(sCvLanguage, "Degree(s) or Diploma(s) obtained") & "</b>"

	If objRsExpEdu.Eof Then
		WriteGridDataRow 4, "", "-#|#-#|#-#|#-"
	Else
        i=1
	While Not objRsExpEdu.Eof
		WriteGridDataRow 4, "", objRsExpEdu("InstNameEng") & "#|#" & ConvertDateForText(objRsExpEdu("eduStartDate"), "&nbsp;", "MMYYYY") & "#|#" & ConvertDateForText(objRsExpEdu("eduEndDate"), "&nbsp;", "MMYYYY") & "#|#" & Trim(objRsExpEdu("edtDescriptionEng") & " " & objRsExpEdu("eduDiploma1Eng")) & "<br>" & Trim(objRsExpEdu("edsDescriptionEng") & " " & objRsExpEdu("id_EduSubject1Eng"))
		i=i+1
		objRsExpEdu.MoveNext
	WEnd
	End If 
	objRsExpEdu.Close  
	Set objRsExpEdu=Nothing
	WriteTableFooter


' Employment records
%>
	<table cellspacing=0 cellpadding=0 align="center" width="97%">
	<tr><td width="100%" bgcolor="#FFFFFF"><p class="txt"><b><% =GetLabel(sCvLanguage, "Employment Record") %></b>:</td></tr>
	</table><br>
<%	
	While Not objRsExpWke.Eof
		WriteGridTableHeader
		Dim sExperienceEndDate
		sExperienceEndDate=""
		If objRsExpWke("wkeEndDateOpen")=1 Then
			sExperienceEndDate=sExperienceEndDate & GetLabel(sCvLanguage, "Ongoing")
			If IsDate(objRsExpWke("wkeEndDate")) And Not IsNull(objRsExpWke("wkeEndDate")) Then sExperienceEndDate=sExperienceEndDate & " ("
		End If
		If IsDate(objRsExpWke("wkeEndDate")) And Not IsNull(objRsExpWke("wkeEndDate")) Then
			sExperienceEndDate=sExperienceEndDate & ConvertDateForText(objRsExpWke("wkeEndDate"), " ", "MonthYear")
			If objRsExpWke("wkeEndDateOpen")=1 Then sExperienceEndDate=sExperienceEndDate & ")"
		End If
		
		WriteGridDataRow 2, "25%#|#75%", "<b>" & GetLabel(sCvLanguage, "Dates") & "</b>" & "#|#" & ConvertDateForText(objRsExpWke("wkeStartDate"), "&nbsp;", "MonthYear") & " - " & sExperienceEndDate
		WriteGridDataRow 2, "", "<b>" & GetLabel(sCvLanguage, "Employer") & "</b>"  & "#|#" &  objRsExpWke("wkeOrgNameEng")

		sCountries = GetExpertExperienceCountryList(iCvID, objRsExpWke("id_ExpWke"), sCvLanguage)
		WriteGridDataRow 2, "", "<b>" & GetLabel(sCvLanguage, "Location") & "</b>" & "#|#" & sCountries
		WriteGridDataRow 2, "", "<b>" & GetLabel(sCvLanguage, "Position held") & "</b>" & "#|#" & objRsExpWke("wkePositionEng")
		dflag=0
		sDescription=ConvertText(objRsExpWke("wkeDescriptionEng"))
		WriteGridDataRow 2, "", "<b>" & GetLabel(sCvLanguage, "Description of duties") & ":</b>" & "#|#" & sDescription
		WriteTableFooter

	objRsExpWke.MoveNext
	WEnd
	objRsExpWke.Close
	Set objRsExpWke=Nothing

' Languages
%>
	<table cellspacing=0 cellpadding=0 align="center" width="97%">
	<tr><td width="100%" bgcolor="#FFFFFF"><p class="txt"><b>9.	<% =GetLabel(sCvLanguage, "Languages") %></b>:</td></tr>
	</table>
<%	
	If (Not objRsExpLngOther.Eof) Or (Not objRsExpLngNative.Eof) Then
		WriteGridTableHeader
		WriteGridDataRow 4, "25%#|#25%#|#25%#|#25%", "<b>" & GetLabel(sCvLanguage, "Language") & "</b>#|#<b>" & GetLabel(sCvLanguage, "Reading") & "</b>#|#<b>" & GetLabel(sCvLanguage, "Speaking") & "</b>#|#<b>" & GetLabel(sCvLanguage, "Writing") & "</b>"
		While Not objRsExpLngNative.Eof
			On Error Resume Next
				sTempLanguage = ReplaceIfEmpty(objRsExpLngNative("lngName" & sCvLanguage), objRsExpLngNative("lngNameEng"))
			On Error GoTo 0
			WriteGridDataRow 4, "", sTempLanguage & "#|#" & arrLanguageLevelTitle(objRsExpLngNative("exlReading")) & "#|#" & arrLanguageLevelTitle(objRsExpLngNative("exlSpeaking")) & "#|#" & arrLanguageLevelTitle(objRsExpLngNative("exlWriting"))
			objRsExpLngNative.MoveNext
		WEnd
		While Not objRsExpLngOther.Eof
			On Error Resume Next
				sTempLanguage = ReplaceIfEmpty(objRsExpLngOther("lngName" & sCvLanguage), objRsExpLngOther("lngNameEng"))
			On Error GoTo 0
			WriteGridDataRow 4, "", sTempLanguage & "#|#" & arrLanguageLevelTitle(objRsExpLngOther("exlReading")) & "#|#" & arrLanguageLevelTitle(objRsExpLngOther("exlSpeaking")) & "#|#" & arrLanguageLevelTitle(objRsExpLngOther("exlWriting"))
			objRsExpLngOther.MoveNext
		WEnd
		WriteTableFooter
	End If
	objRsExpLngNative.Close
	Set objRsExpLngNative=Nothing	
	objRsExpLngOther.Close
	Set objRsExpLngOther=Nothing	
	%>

	<table cellspacing=0 cellpadding=0 align="center" width="97%">
	<tr><td width="100%" bgcolor="#FFFFFF"><p class="txt"><b><% =GetLabel(sCvLanguage, "Certification") %>:</b> <br><br><% =GetLabel(sCvLanguage, "I, the undersigned, certify (AFB)...") %><br><br>____________________________________________________&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Date:________________<br><small><i>[<% =GetLabel(sCvLanguage, "Signature of staff member or authorized representative of the staff") %>]</i></small> &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; <small><i><% =GetLabel(sCvLanguage, "Day/Month/Year") %></i></small><br></td></tr>
	</table><br>
	

</td>
<td width="5%">&nbsp;&nbsp;</td>
<td width="20%" valign="top">
   <!-- Feature boxes -->
	<% ShowExpCVFeatureBox %>

</td>
</tr>
</table>

<% CloseDBConnection %>
</body>
</html>
