<%
'--------------------------------------------------------------------
'
' Expert's CV. View in Tripleline format
'
'--------------------------------------------------------------------
%>
<!--#include file="cv_data.asp"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" type="text/css" href="<% =sHomePath %>styles.css">
<title>CV in Tripleline format</title>
</head>

<body bgcolor="#FFFFFF" topmargin=0 leftmargin=0 marginheight=0 marginwidth=0>
<% ShowTopMenu %>

<p class="ttl" align="center"><b>Curriculum Vitae - Tripleline format</b></p>

<table width="98%" cellpadding=0 cellspacing=0 border=0 align="center">
<tr><td width="85%" valign="top">
<br />

<%
sLastName=Replace(ConvertText(sLastName), "      ", "")
sFirstName=Replace(ConvertText(sFirstName), "      ", "")
sFullNameWithSpaces=sFirstName & " " & sLastName
sFullName=Replace(ConvertText(sFullName), "      ", "")

	WriteTableHeader
	
	WriteSimpleRowNoBorder "<h3 style=""font-face: Arial; color: #369865; margin-left: 25%"">" & sFullNameWithSpaces & "</font><h3>"
	WriteRowBorder
	WriteDataRow "Position", ""
	WriteRowBorder
	WriteDataRow "Profile", ConvertText(sKeyQualification)
	WriteRowBorder
	WriteDataRow "Key Skills", ""
	WriteRowBorder
	
' Nationality & languages
	Dim sLanguagesList
	sLanguagesList = ""
	If (Not objRsExpLngOther.Eof) Or (Not objRsExpLngNative.Eof) Then
		sLanguagesList = sLanguagesList & GetListStart()
		While Not objRsExpLngNative.Eof
			sLanguagesList = sLanguagesList & GetListItemStart() & objRsExpLngNative("lngNameEng") & " (native)" & GetListItemEnd()
			objRsExpLngNative.MoveNext
		WEnd

		While Not objRsExpLngOther.Eof
			sLanguagesList = sLanguagesList & GetListItemStart() & objRsExpLngOther("lngNameEng") & " (" & LCase(arrLanguageLevelTitle(objRsExpLngOther("exlAverage"))) & ")" & GetListItemEnd()
			objRsExpLngOther.MoveNext
		WEnd
		sLanguagesList = sLanguagesList & GetListEnd()
	End If
	objRsExpLngNative.Close
	Set objRsExpLngNative=Nothing	
	objRsExpLngOther.Close
	Set objRsExpLngOther=Nothing	

	WriteDataRow2 "Nationality", sNationality, "Languages", sLanguagesList
	WriteRowBorder
	
' Education	
	Dim sEducationList
	sEducationList = ""
	If Not objRsExpEdu.Eof Then
		sEducationList = sEducationList & GetListStart()
		While Not objRsExpEdu.Eof
			sEduSubject = ""
			If Len(objRsExpEdu("edsDescriptionEng"))>0 And objRsExpEdu("edsDescriptionEng")<>"Other" Then 
				sEduSubject=sEduSubject & Trim(objRsExpEdu("edsDescriptionEng"))
			End If
			If Len(objRsExpEdu("id_EduSubject1Eng"))>0 Then
				sEduSubject=sEduSubject & Trim(" " & objRsExpEdu("id_EduSubject1Eng"))
			End If
			If Len(sEduSubject)>0 Then
				sEduSubject = " in " & sEduSubject
			End If

			sEducationList = sEducationList & _
				GetListItemStart() & _
				Trim(EducationTypeTitleByID(objRsExpEdu("eduDiploma")) & _ 
				sEduSubject & _
				", " & _
				objRsExpEdu("InstLocationEng") & " " & _
				ConvertDateForText(objRsExpEdu("eduEndDate"), " ", "Year"))
				
			objRsExpEdu.MoveNext
			
			If Not objRsExpEdu.Eof Then
				sEducationList = sEducationList & GetListItemEnd()
			Else
				sEducationList = sEducationList & GetListItemEndLast()
			End If
		WEnd
		sEducationList = sEducationList & GetListEnd()
	End If
	objRsExpEdu.Close
	Set objRsExpEdu=Nothing	

	WriteDataRow "Qualifications", sEducationList
	WriteRowBorder
	
' Country Experience
	Dim sCountryList
	sCountryList = GetExpertExperienceContinentCountryGroupedList2(iCvID, Null, sCvLanguage, "<font color=""#369865""><b>", "", ": </b></font> ", "", ",")
	
	WriteDataRow "Country Experience", sCountryList
	WriteRowBorder
	
	
' Clients
	Dim sClientList
	sClientList = GetExpertExperienceOrganisationList(iCvID, Null, Null)
	WriteDataRow "Clients", sClientList
	
' Employment History
	WriteSimpleRow "<font color=""#369865""><b>Employment History</b></font>"

	Dim sExperienceEndDate, sExperiencePeriod, sExperienceTitle
	If not objRsExpWke.Eof Then
	While Not objRsExpWke.Eof
		If IsDate(objRsExpWke("wkeStartDate")) And Not IsNull(objRsExpWke("wkeStartDate")) Then
			sExperiencePeriod = Year(objRsExpWke("wkeStartDate"))
		Else 
			sExperiencePeriod = ""
		End If
		If objRsExpWke("wkeEndDateOpen")=1 Then
			sExperienceEndDate = "present"
		ElseIf IsDate(objRsExpWke("wkeEndDate")) And Not IsNull(objRsExpWke("wkeEndDate")) Then
			sExperienceEndDate = Year(objRsExpWke("wkeEndDate"))
		Else 
			sExperienceEndDate = ""
		End If

		If Len(sExperiencePeriod)>0 And Len(sExperienceEndDate)>0 Then
			sExperiencePeriod = sExperiencePeriod & "-" & sExperienceEndDate
		Else
			sExperiencePeriod = sExperiencePeriod & sExperienceEndDate
		End If

		sExperienceTitle = ""
		If Len(objRsExpWke("wkePositionEng"))>0 Then
			sExperienceTitle = sExperienceTitle & objRsExpWke("wkePositionEng") & ", "
		End If
		If Len(objRsExpWke("wkeOrgNameEng"))>0 Then
			sExperienceTitle = sExperienceTitle & objRsExpWke("wkeOrgNameEng") & ", "
		End If
		
		objTempRs2=GetDataOutParamsSP("usp_ExpCvvExperienceCouSelect", Array( _
			Array(, adInteger, , iCvID), Array(, adInteger, , objRsExpWke("id_ExpWke")), Array(, adVarChar, 10, "listshort")), Array( _
			Array(, adVarWChar, 4000)))
		sCountries=ConvertText(objTempRs2(0))
		Set objTempRs2=Nothing
		If Len(sCountries)>0 Then
			sExperienceTitle = sExperienceTitle & sCountries & ", "
		End If

		If Len(sExperienceTitle)>0 Then
			sExperienceTitle = Left(sExperienceTitle, Len(sExperienceTitle)-2)
		End If
		
		WriteDataRow sExperiencePeriod, sExperienceTitle
	objRsExpWke.MoveNext
	WEnd
	End If


' Selected Experience
	WriteSimpleRow "<font color=""#369865""><b>Selected Experience</b></font>"
	If Not objRsExpWke.Eof Then
	objRsExpWke.MoveFirst
	While Not objRsExpWke.Eof
		If IsDate(objRsExpWke("wkeStartDate")) And Not IsNull(objRsExpWke("wkeStartDate")) Then
			sExperiencePeriod = Year(objRsExpWke("wkeStartDate"))
		Else 
			sExperiencePeriod = ""
		End If
		If IsDate(objRsExpWke("wkeEndDate")) And Not IsNull(objRsExpWke("wkeEndDate")) Then
			sExperienceEndDate = Year(objRsExpWke("wkeEndDate"))
		Else 
			sExperienceEndDate = ""
		End If

		If Len(sExperiencePeriod)>0 And Len(sExperienceEndDate)>0 Then
			If sExperiencePeriod <> sExperienceEndDate Then
				sExperiencePeriod = sExperiencePeriod & "-" & sExperienceEndDate
			End If
		Else
			sExperiencePeriod = sExperiencePeriod & sExperienceEndDate
		End If
		sExperiencePeriod = "" & sExperiencePeriod & "</b></font>"

		objTempRs2=GetDataOutParamsSP("usp_ExpCvvExperienceCouSelect", Array( _
			Array(, adInteger, , iCvID), Array(, adInteger, , objRsExpWke("id_ExpWke")), Array(, adVarChar, 10, "listshort")), Array( _
			Array(, adVarWChar, 4000)))
		sCountries=ConvertText(objTempRs2(0))
		Set objTempRs2=Nothing
		If Len(sCountries)>0 Then
			sExperiencePeriod = sExperiencePeriod & "<br />" & sCountries
		End If

		sExperienceTitle = ""
		If Len(objRsExpWke("wkePrjTitleEng"))>0 Then
			sExperienceTitle = sExperienceTitle & objRsExpWke("wkePrjTitleEng") & ", "
		End If
		If Len(objRsExpWke("wkeOrgNameEng"))>0 Then
			sExperienceTitle = sExperienceTitle & objRsExpWke("wkeOrgNameEng") & ", "
		End If
		If Len(objRsExpWke("wkePositionEng"))>0 Then
			sExperienceTitle = sExperienceTitle & objRsExpWke("wkePositionEng") & ", "
		End If
		
		If Len(sExperienceTitle)>0 Then
			sExperienceTitle = Left(sExperienceTitle, Len(sExperienceTitle)-2)
		End If
		
		sDescription=ConvertText(objRsExpWke("wkeDescriptionEng"))
		
		sExperienceTitle = "<font color=""#369865""><b>" & sExperienceTitle & "</b></font><br />" & sDescription
		
		WriteDataRow sExperiencePeriod, sExperienceTitle
	objRsExpWke.MoveNext
	WEnd
	End If
	objRsExpWke.Close
	Set objRsExpWke=Nothing

	
' Publications
	If sPublications>"" Then
		WriteSimpleRow "<font color=""#369865""><b>Publications</b></font>"
		WriteSimpleRowNoBorder "</b></font>" & ConvertText(sPublications)
	End If


' Other relevant information
	WriteSimpleRow "<font color=""#369865""><b>Other relevant information</b></font>"
	If sMemberships>"" Then
		WriteSimpleRowNoBorder "Membership:</b></font>  " & ConvertText(sMemberships)
	End If
	If sOtherSkills>"" Then
		WriteSimpleRowNoBorder "Other skills:</b></font>  " & ConvertText(sOtherSkills)
	End If

	WriteTableFooter
%>

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



