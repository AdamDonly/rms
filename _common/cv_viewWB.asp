<%
'--------------------------------------------------------------------
'
' Expert's CV. View in WB format
'
'--------------------------------------------------------------------
%>
<!--#include file="cv_data.asp"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" type="text/css" href="<% =sHomePath %>styles.css">
<title>CV in World Bank Format</title>
</head>

<body bgcolor="#FFFFFF" topmargin=0 leftmargin=0 marginheight=0 marginwidth=0>
<% ShowTopMenu %>

<p class="ttl" align="center"><b>Curriculum Vitae - World Bank format</b></p>

<table width="98%" cellpadding=0 cellspacing=0 border=0 align="center">
<tr><td width="85%" valign="top">
<br />

<%
' Personal information
	WriteTableHeader
	WriteDataRow "<b>1. " & GetLabel(sCvLanguage, "Proposed Position") & ":</b>", "<small><i>[" & GetLabel(sCvLanguage, "only one candidate...") & "]</i></small>"
	WriteDataRow "<b>2. " & GetLabel(sCvLanguage, "Name of Firm") & ":</b>", "<small><i>[" & GetLabel(sCvLanguage, "insert name of firm proposing the staff") & "]</i></small>"
	WriteDataRow "<b>3. " & GetLabel(sCvLanguage, "Name of Staff") & ":</b>", sFullName
	WriteDataRow "<b>4. " & GetLabel(sCvLanguage, "Date of Birth") & ":</b>", ConvertDateForText(sBirthDate, " ", "DayMonthYear") & " &nbsp;  &nbsp;  &nbsp;  &nbsp;  &nbsp;  &nbsp;  &nbsp;  &nbsp;  &nbsp;  &nbsp;  &nbsp;  &nbsp;  &nbsp;  &nbsp;  &nbsp; <b>" & GetLabel(sCvLanguage, "Nationality") & ":</b> &nbsp;  &nbsp; " & sNationality 
	WriteSpaceRow
	WriteDataRow "&nbsp;&nbsp;&nbsp;&nbsp;<b>" & GetLabel(sCvLanguage, "Address") & ":<br></b>", sPermAddress
	WriteTableFooter

' Education
%>
	<table cellspacing=0 cellpadding=0 align="center" width="97%">
	<tr><td width="100%" bgcolor="#FFFFFF"><p class="txt"><b>5.	<% =GetLabel(sCvLanguage, "Education (WB)") %></b> <small><i>[<% =GetLabel(sCvLanguage, "Indicate college/university...") %>]</i></small> :</td></tr>
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

' Membership of Professional Associations:
%>
	<table cellspacing=0 cellpadding=0 align="center" width="97%">
	<tr><td width="100%" bgcolor="#FFFFFF"><p class="txt"><b>6. <% =GetLabel(sCvLanguage, "Membership of Professional Associations") %></b>: <% =sMemberships %></td></tr>
	</table><br>
<%
	
' Other training
%>
	<table cellspacing=0 cellpadding=0 align="center" width="97%">
	<tr><td width="100%" bgcolor="#FFFFFF"><p class="txt"><b>7.	<% =GetLabel(sCvLanguage, "Other Training") %></b> <small><i>[<% =GetLabel(sCvLanguage, "Indicate significant training...") %>]</i></small> :</td></tr>
	</table>
<%
	WriteGridTableHeader
	WriteGridDataRow 4, "25%#|#13%#|#13%#|#50%", "<b>" & GetLabel(sCvLanguage, "Institution") & "</b>#|#<b>" & GetLabel(sCvLanguage, "Start date") & "</b>#|#<b>" & GetLabel(sCvLanguage, "End date") & "</b>#|#<b>" & GetLabel(sCvLanguage, "Degree(s) or Diploma(s) obtained") & "</b>"
	If objRsExpTrn.Eof Then
		WriteGridDataRow 4, "", "-#|#-#|#-#|#-"
	Else
        i=1
	While Not objRsExpTrn.eof
		WriteGridDataRow 4, "", objRsExpTrn("InstNameEng") & "#|#" & ConvertDateForText(objRsExpTrn("eduStartDate"), "&nbsp;", "MMYYYY") & "#|#" & ConvertDateForText(objRsExpTrn("eduEndDate"), "&nbsp;", "MMYYYY") & "#|#" & objRsExpTrn("edtDescriptionEng") & " " & objRsExpTrn("eduDiploma1Eng") & Trim(objRsExpTrn("edsDescriptionEng") & " " & objRsExpTrn("id_EduSubject1Eng"))
		i=i+1
		objRsExpTrn.MoveNext
	WEnd
	End If 
	objRsExpTrn.Close  
	Set objRsExpTrn=Nothing
	WriteTableFooter

' Countries of Work Experience
	Dim sExperienceCountries
	sExperienceCountries = GetExpertExperienceCountryList(iCvID, Null, sCvLanguage)
%>
	<table cellspacing=0 cellpadding=0 align="center" width="97%">
	<tr><td width="100%" bgcolor="#FFFFFF"><p class="txt"><b>8.	<% =GetLabel(sCvLanguage, "Countries of Work Experience") %></b> <small><i>[<% =GetLabel(sCvLanguage, "List countries where staff has worked in the last ten years") %>]</i></small> : <% =sExperienceCountries %></td></tr>
	</table><br>
<%	
	
' Languages
%>
	<table cellspacing=0 cellpadding=0 align="center" width="97%">
	<tr><td width="100%" bgcolor="#FFFFFF"><p class="txt"><b>9.	<% =GetLabel(sCvLanguage, "Languages") %></b> <small><i>[<% =GetLabel(sCvLanguage, "For each language indicate proficiency...") %>]</i></small> :</td></tr>
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

' Employment records
%>
	<table cellspacing=0 cellpadding=0 align="center" width="97%">
	<tr><td width="100%" bgcolor="#FFFFFF"><p class="txt"><b>10. <% =GetLabel(sCvLanguage, "Employment Record") %></b> <small><i>[<% =GetLabel(sCvLanguage, "Starting with present position...") %>]</i></small> :</td></tr>
	</table><br>
<%	
	While Not objRsExpWke.Eof
		If objRsExpWke("TypeofWke")=2 Or (objRsExpWke("TypeofWke")<=1 And Len(ReplaceIfEmpty(objRsExpWke("wkePrjTitleEng"), ""))<=1) Then
		'If Len(ReplaceIfEmpty(objRsExpWke("wkePrjTitleEng"), ""))<=1 Then
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

			'sCountries = GetExpertExperienceCountryList(iCvID, objRsExpWke("id_ExpWke"), sCvLanguage)
			'WriteGridDataRow 2, "", "<b>" & GetLabel(sCvLanguage, "Location") & "</b>" & "#|#" & sCountries
			WriteGridDataRow 2, "", "<b>" & GetLabel(sCvLanguage, "Position held") & "</b>" & "#|#" & objRsExpWke("wkePositionEng")
			dflag=0
			'sDescription=ConvertText(objRsExpWke("wkeDescriptionEng"))
			'WriteGridDataRow 2, "", "<b>" & GetLabel(sCvLanguage, "Description of duties") & ":</b>" & "#|#" & sDescription
		WriteTableFooter
		End If

	objRsExpWke.MoveNext
	WEnd
	'objRsExpWke.Close
	'Set objRsExpWke=Nothing

	
	WriteGridTableHeader
	WriteGridDataRow 2, "25%#|#75%", "<b>11. " & GetLabel(sCvLanguage, "Detailed Tasks Assigned") & "</b><br><small><i>[" & GetLabel(sCvLanguage, "List all tasks to be performed under this assignment") & "]</i></small>" & "#|#" & "<b>12. " & GetLabel(sCvLanguage, "Work Undertaken that Best Illustrates Capability to Handle the Tasks Assigned") & "</b><br><small><i>[" & GetLabel(sCvLanguage, "Among the assignments in which the staffs have been involved...") & ".]</i></small><br><br>"
	%>
	<tr>
	<td width="25%" bgcolor="#FFFFFF" valign="top"></small></td>
	<td width="75%" bgcolor="#FFFFFF" valign="top">
	<%
		objRsExpWke.MoveFirst
		While Not objRsExpWke.Eof
			If objRsExpWke("TypeofWke")=3 Or (objRsExpWke("TypeofWke")<=1 And Len(ReplaceIfEmpty(objRsExpWke("wkePrjTitleEng"), ""))>1) Then
			'If Len(ReplaceIfEmpty(objRsExpWke("wkePrjTitleEng"), ""))>1 Then
			
				Dim sYear
				If IsNull(objRsExpWke("wkeStartDate")) Then
					sYear=Year(objRsExpWke("wkeEndDate"))
				ElseIf IsNull(objRsExpWke("wkeEndDate")) Then
					sYear=Year(objRsExpWke("wkeStartDate"))
				ElseIf DateDiff("yyyy", objRsExpWke("wkeStartDate"), objRsExpWke("wkeEndDate"))<=0 Then
					sYear=Year(objRsExpWke("wkeStartDate"))
				Else 
					sYear=Year(objRsExpWke("wkeStartDate")) & " - " & Year(objRsExpWke("wkeEndDate"))
				End If			
				
				Response.Write "<p class=""txt""><b>" & GetLabel(sCvLanguage, "Name of assignment or project") & ": </b>" & objRsExpWke("wkePrjTitleEng") & "</p>"
				
				Response.Write "<p class=""txt""><b>" & GetLabel(sCvLanguage, "Year") & ": </b>" & sYear & "</p>"
				sCountries = GetExpertExperienceCountryList(iCvID, objRsExpWke("id_ExpWke"), sCvLanguage)
				Set objTempRs2=Nothing
				Response.Write "<p class=""txt""><b>" & GetLabel(sCvLanguage, "Location") & ": </b>" & sCountries & "</p>"
				Response.Write "<p class=""txt""><b>" & GetLabel(sCvLanguage, "Client") & ": </b>" & objRsExpWke("wkeBnfNameEng") & "</p>"
				Response.Write "<p class=""txt""><b>" & GetLabel(sCvLanguage, "Main project features") & ": </b>" & objRsExpWke("wkeProjectDescription") & "</p>"
				Response.Write "<p class=""txt""><b>" & GetLabel(sCvLanguage, "Position held") & ": </b>" & objRsExpWke("wkePositionEng") & "</p>"
				dflag=0
				sDescription=ConvertText(objRsExpWke("wkeDescriptionEng"))
				Response.Write "<p class=""txt""><b>" & GetLabel(sCvLanguage, "Activities performed") & ": </b>" & sDescription & "</p>"
				Response.Write "<br />"
			End If
		objRsExpWke.MoveNext
		WEnd
		objRsExpWke.Close
		Set objRsExpWke=Nothing
	%>
	</td>
	</tr>
	<%
	WriteTableFooter
	
	
'	WriteTableHeader
'	WriteDataRow "Detailed Tasks Assigned:", "<font color=""#C0C0C0"">(fill in with your data)</font><br><br>"
'	WriteDataRow "<b>Key&nbsp;qualifications:</b>", ConvertText(sKeyQualification)
'	WriteTableFooter
	
	%>

	<table cellspacing=0 cellpadding=0 align="center" width="97%">
	<tr><td width="100%" bgcolor="#FFFFFF"><p class="txt"><b>13. <% =GetLabel(sCvLanguage, "Certification") %>:</b> <br><br><% =GetLabel(sCvLanguage, "I, the undersigned, certify...") %><br><br>____________________________________________________&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Date:________________<br><small><i>[<% =GetLabel(sCvLanguage, "Signature of staff member or authorized representative of the staff") %>]</i></small> &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; <small><i><% =GetLabel(sCvLanguage, "Day/Month/Year") %></i></small><br><br><br><% =GetLabel(sCvLanguage, "Full name of authorized representative") %>: ___________________________________________</td></tr>
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
