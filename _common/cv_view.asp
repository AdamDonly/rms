<%
'--------------------------------------------------------------------
'
' Expert's CV. View in assortis.com format
' With or without contact details
'
'--------------------------------------------------------------------
%>
<!--#include file="cv_data.asp"-->
<!--#include virtual="/_common/_class/main.asp"-->
<!--#include virtual="/_common/_class/project.asp"-->
<!--#include virtual="/_common/_class/expert.asp"-->
<!--#include virtual="/_common/_class/expert.project.asp"-->
<html>
<head>
<meta name="robots" content="noarchive">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" type="text/css" href="<% =sHomePath %>styles.css">
<title>CV in assortis.com format</title>
</head>
<body bgcolor="#FFFFFF" topmargin=0 leftmargin=0 marginheight=0 marginwidth=0>
<% ShowTopMenu %>

<p class="ttl" align="center"><b>
<% If bCvValidForMemberOrExpert=1 Or bCvValidForMemberOrExpert=5 Then %>
	Curriculum Vitae - assortis.com format
<% Else %>
	<table width="100%"><tr>
	<td width="30%"><p class="ttl">Expert ID: <%=iCvID%></p></td>
	<td width="70%"><p class="ttl"><b>Curriculum Vitae - contact detail free version</b></p></td>
	</tr></table>
<% End If %>
</b></p>

<table width="98%" cellpadding=0 cellspacing=0 border=0 align="center">
<tr><td width="85%" valign="top">

<%
' Personal information
	WriteDataTitle GetLabel(sCvLanguage, "PERSONAL INFORMATION")
	ShowExpertsBlockSubTitle "98%", "52", "ex"  & bCvValidForMemberOrExpert
	ShowUserNoticesViewHeader "99%", 180
	ShowUserNoticesViewSpacer 4

	If bCvValidForMemberOrExpert=1 Then
		ShowUserNoticesViewText GetLabel(sCvLanguage, "Personal title"), sTitle
		ShowUserNoticesViewText GetLabel(sCvLanguage, "First name"), sFirstName
		If sMiddleName>"" Then ShowUserNoticesViewText GetLabel(sCvLanguage, "Middle name"), sMiddleName
		ShowUserNoticesViewText GetLabel(sCvLanguage, "Family name"), sLastName
		ShowUserNoticesViewText GetLabel(sCvLanguage, "Date of birth"), ConvertDateForText(sBirthDate, "&nbsp;", "DayMonthYear")
		If sBirthPlace>"" Then ShowUserNoticesViewText GetLabel(sCvLanguage, "Place of birth"), sBirthPlace
	End If
	ShowUserNoticesViewText GetLabel(sCvLanguage, "Nationality"), sNationality
	If iGender>0 And IsNumeric(iGender) Then ShowUserNoticesViewText GetLabel(sCvLanguage, "Gender"), arrGenderTitle(iGender)
	If iMaritalStatus>0 And IsNumeric(iMaritalStatus) Then ShowUserNoticesViewText GetLabel(sCvLanguage, "Marital status"), arrMaritalStatusTitle(iMaritalStatus)

	ShowUserNoticesViewSpacer 5
	ShowUserNoticesViewFooter
	ShowExpertsBlockFooter "98%", "52", "ex"  & bCvValidForMemberOrExpert

' Selection on internal projects
Dim objExpertProjectList
Set objExpertProjectList = New CExpertProjectList
objExpertProjectList.Expert.ID = iCvID
objExpertProjectList.LoadData
If objExpertProjectList.Count>0 Then
	WriteDataTitle "SELECTION ON INTERNAL PROJECTS"
	For i=0 To objExpertProjectList.Count-1
		ShowExpertsBlockSubTitle "98%", "52", "ex"  & bCvValidForMemberOrExpert
		ShowUserNoticesViewHeader "99%", 180
		ShowUserNoticesViewSpacer 4
		objExpertProjectList.Item(i).Project.LoadData
		ShowUserNoticesViewText "</b>Project", objExpertProjectList.Item(i).Project.Title
		ShowUserNoticesViewText "</b>Country / Region", objExpertProjectList.Item(i).Project.Location
		ShowUserNoticesViewText "</b>Status", objExpertProjectList.Item(i).Status.Name
		ShowUserNoticesViewText "</b>Fee", objExpertProjectList.Item(i).Fee.Value & " " & objExpertProjectList.Item(i).Fee.CurrencyCode
		ShowUserNoticesViewText "</b>Registration date", ConvertDateForText(objExpertProjectList.Item(i).ProvidedDate, "&nbsp;", "MMYYYY")
		ShowUserNoticesViewText "</b>Comments", objExpertProjectList.Item(i).Comments
		ShowUserNoticesViewSpacer 5
		ShowUserNoticesViewFooter
		ShowExpertsBlockFooter "98%", "52", "ex"  & bCvValidForMemberOrExpert
	Next
End If 
Set objExpertProjectList = Nothing

' Education	
	If Not objRsExpEdu.Eof Then
	WriteDataTitle GetLabel(sCvLanguage, "EDUCATION")
	While Not objRsExpEdu.Eof
		ShowExpertsBlockSubTitle "98%", "52", "ex"  & bCvValidForMemberOrExpert
		ShowUserNoticesViewHeader "99%", 180
		ShowUserNoticesViewSpacer 4
		If objRsExpEdu("InstNameEng")>"" Then ShowUserNoticesViewText GetLabel(sCvLanguage, "Institution"), objRsExpEdu("InstNameEng")
		If objRsExpEdu("InstLocationEng")>"" Then ShowUserNoticesViewText GetLabel(sCvLanguage, "Location"), objRsExpEdu("InstLocationEng")
		If IsDate(objRsExpEdu("eduStartDate")) Then ShowUserNoticesViewText GetLabel(sCvLanguage, "Start date"), ConvertDateForText(objRsExpEdu("eduStartDate"), "&nbsp;", "MMYYYY")
		If IsDate(objRsExpEdu("eduEndDate")) Then ShowUserNoticesViewText GetLabel(sCvLanguage, "End date"), ConvertDateForText(objRsExpEdu("eduEndDate"), "&nbsp;", "MMYYYY")
		If Not IsNull(objRsExpEdu("eduDiploma")) Then
			ShowUserNoticesViewText GetLabel(sCvLanguage, "Type of diploma"), Trim(EducationTypeTitleByID(objRsExpEdu("eduDiploma")) & " " & objRsExpEdu("eduDiploma1Eng"))
		Else
			ShowUserNoticesViewText GetLabel(sCvLanguage, "Type of diploma"), objRsExpEdu("eduDiploma1Eng")
		End If
		sEduSubject=""
		If Len(objRsExpEdu("edsDescriptionEng"))>0 And objRsExpEdu("edsDescriptionEng")<>"Other" Then 
			sEduSubject=sEduSubject & Trim(objRsExpEdu("edsDescriptionEng"))
		End If
		If Len(objRsExpEdu("id_EduSubject1Eng"))>0 Then
			sEduSubject=sEduSubject & Trim(" " & objRsExpEdu("id_EduSubject1Eng"))
		End If
		If Len(sEduSubject)>0 Then
			ShowUserNoticesViewText GetLabel(sCvLanguage, "Subject"), sEduSubject
		End If
		ShowUserNoticesViewSpacer 5
		ShowUserNoticesViewFooter
		ShowExpertsBlockFooter "98%", "52", "ex"  & bCvValidForMemberOrExpert
	objRsExpEdu.MoveNext
	WEnd
	End If
	objRsExpEdu.Close
	Set objRsExpEdu=Nothing

' Training
	If not objRsExpTrn.Eof Then
	WriteDataTitle GetLabel(sCvLanguage, "TRAINING")
	While Not objRsExpTrn.Eof
		ShowExpertsBlockSubTitle "98%", "52", "ex"  & bCvValidForMemberOrExpert
		ShowUserNoticesViewHeader "99%", 180
		ShowUserNoticesViewSpacer 4
		ShowUserNoticesViewText GetLabel(sCvLanguage, "Skills / Qualification"), objRsExpTrn("eduOtherEng")
		ShowUserNoticesViewText GetLabel(sCvLanguage, "Title"), objRsExpTrn("eduDiploma1Eng")
		ShowUserNoticesViewText GetLabel(sCvLanguage, "Start date"), ConvertDateForText(objRsExpTrn("eduStartDate"), "&nbsp;", "MMYYYY")
		ShowUserNoticesViewText GetLabel(sCvLanguage, "End date"), ConvertDateForText(objRsExpTrn("eduEndDate"), "&nbsp;", "MMYYYY")
		sAchievements=objRsExpTrn("eduDescriptionEng")
		If sAchievements>"" Then
			ShowUserNoticesViewText GetLabel(sCvLanguage, "Achievements"), sAchievements
		End If
		ShowUserNoticesViewSpacer 5
		ShowUserNoticesViewFooter
		ShowExpertsBlockFooter "98%", "52", "ex"  & bCvValidForMemberOrExpert
	objRsExpTrn.MoveNext
	WEnd
	End If
	objRsExpTrn.Close
	Set objRsExpTrn=Nothing

' Professional experience
	WriteDataTitle GetLabel(sCvLanguage, "PROFESSIONAL EXPERIENCE")
	ShowExpertsBlockSubTitle "98%", "52", "ex"  & bCvValidForMemberOrExpert
	ShowUserNoticesViewHeader "99%", 180
	ShowUserNoticesViewSpacer 4
	ShowUserNoticesViewText GetLabel(sCvLanguage, "Status"), sProfessionalStatus
	ShowUserNoticesViewText GetLabel(sCvLanguage, "Profession"), sProfession
	ShowUserNoticesViewText GetLabel(sCvLanguage, "Current position"), sPosition
	ShowUserNoticesViewText GetLabel(sCvLanguage, "Key qualifications"), sKeyQualification
	ShowUserNoticesViewText GetLabel(sCvLanguage, "Other skills"), sOtherSkills
	ShowUserNoticesViewText GetLabel(sCvLanguage, "Years of professional experience"), iProfYears
	ShowUserNoticesViewSpacer 5
	ShowUserNoticesViewFooter
	ShowExpertsBlockFooter "98%", "52", "ex"  & bCvValidForMemberOrExpert

' Employment records
	If not objRsExpWke.Eof Then
	WriteDataTitle GetLabel(sCvLanguage, "EMPLOYMENT RECORD AND COMPLETED PROJECTS")
	While Not objRsExpWke.Eof
		
		ShowExpertsBlockSubTitle "98%", "52", "ex"  & bCvValidForMemberOrExpert
		ShowUserNoticesViewHeader "99%", 180
		ShowUserNoticesViewSpacer 4
		ShowUserNoticesViewText GetLabel(sCvLanguage, "Project title"), "<b>" & HighlightKeywords(objRsExpWke("wkePrjTitleEng"), sSearchKeywordsHighlight) & "</b>"
		ShowUserNoticesViewText GetLabel(sCvLanguage, "Start date"), ConvertDateForText(objRsExpWke("wkeStartDate"), "&nbsp;", "MMYYYY")
		Dim sExperienceEndDate
		sExperienceEndDate=""
		If objRsExpWke("wkeEndDateOpen")=1 Then
			sExperienceEndDate=sExperienceEndDate & GetLabel(sCvLanguage, "Ongoing")
			If IsDate(objRsExpWke("wkeEndDate")) And Not IsNull(objRsExpWke("wkeEndDate")) Then sExperienceEndDate=sExperienceEndDate & " ("
		End If
		If IsDate(objRsExpWke("wkeEndDate")) And Not IsNull(objRsExpWke("wkeEndDate")) Then
			sExperienceEndDate=sExperienceEndDate &  ConvertDateForText(objRsExpWke("wkeEndDate"), "&nbsp;", "MMYYYY")
			If objRsExpWke("wkeEndDateOpen")=1 Then sExperienceEndDate=sExperienceEndDate & ")"
		End If
		ShowUserNoticesViewText GetLabel(sCvLanguage, "End date"), sExperienceEndDate
		ShowUserNoticesViewText GetLabel(sCvLanguage, "Company / Organisation"), objRsExpWke("wkeOrgNameEng")
		ShowUserNoticesViewText GetLabel(sCvLanguage, "Position / Responsibility"), "<b>" & objRsExpWke("wkePositionEng") & "</b>"
		ShowUserNoticesViewText GetLabel(sCvLanguage, "Beneficiary"), objRsExpWke("wkeBnfNameEng")
	sCountries = GetExpertExperienceCountryGroupedList(iCvID, objRsExpWke("id_ExpWke"), sCvLanguage)
	sSectors = GetExpertExperienceSectorGroupedList(iCvID, objRsExpWke("id_ExpWke"), sCvLanguage)

objTempRs2=GetDataOutParamsSP("usp_ExpCvvExperienceDonSelect", Array( _
	Array(, adInteger, , iCvID), Array(, adInteger, , objRsExpWke("id_ExpWke")), Array(, adVarChar, 10, "list")), Array( _
	Array(, adVarWChar, 4000)))
sDonors=objTempRs2(0)
Set objTempRs2=Nothing

		ShowUserNoticesViewText GetLabel(sCvLanguage, "Funding agencies"), sDonors
		ShowUserNoticesViewText GetLabel(sCvLanguage, "Countries"), sCountries
		ShowUserNoticesViewText GetLabel(sCvLanguage, "Sectors"), sSectors
		If objRsExpWke("wkeClientRefEng")>"" Then ShowUserNoticesViewText GetLabel(sCvLanguage, "Client references"), objRsExpWke("wkeClientRefEng")
		ShowUserNoticesViewText GetLabel(sCvLanguage, "Description of tasks"), ConvertText(objRsExpWke("wkeDescriptionEng"))

		ShowUserNoticesViewSpacer 5
		ShowUserNoticesViewFooter
		ShowExpertsBlockFooter "98%", "52", "ex"  & bCvValidForMemberOrExpert
	objRsExpWke.MoveNext
	WEnd
	End If
	objRsExpWke.Close
	Set objRsExpWke=Nothing

' Languages
	WriteDataTitle GetLabel(sCvLanguage, "LANGUAGES SKILLS")
	ShowExpertsBlockSubTitle "98%", "52", "ex"  & bCvValidForMemberOrExpert
	ShowUserNoticesViewHeader "99%", 180
	ShowUserNoticesViewSpacer 4
	ShowUserNoticesViewText GetLabel(sCvLanguage, "Language"), GetLabel(sCvLanguage, "Reading") & " / " & GetLabel(sCvLanguage, "Speaking") & " / " & GetLabel(sCvLanguage, "Writing")
	
	If (Not objRsExpLngOther.Eof) Or (Not objRsExpLngNative.Eof) Then
		While Not objRsExpLngNative.Eof
			On Error Resume Next
				sTempLanguage = ReplaceIfEmpty(objRsExpLngNative("lngName" & sCvLanguage), objRsExpLngNative("lngNameEng"))
			On Error GoTo 0
			ShowUserNoticesViewText sTempLanguage, GetLabel(sCvLanguage, "Native")
			objRsExpLngNative.MoveNext
		WEnd

		Dim sReading, sSpeaking, sWriting
		While Not objRsExpLngOther.Eof
			On Error Resume Next
				sTempLanguage = ReplaceIfEmpty(objRsExpLngOther("lngName" & sCvLanguage), objRsExpLngOther("lngNameEng"))
			On Error GoTo 0		
			If IsNumeric(objRsExpLngOther("exlReading")) And objRsExpLngOther("exlReading")>"" Then
				sReading = arrLanguageLevelTitle(objRsExpLngOther("exlReading"))
			Else
				sReading = ""
			End If
			If IsNumeric(objRsExpLngOther("exlSpeaking")) And objRsExpLngOther("exlSpeaking")>"" Then
				sSpeaking = arrLanguageLevelTitle(objRsExpLngOther("exlSpeaking"))
			Else
				sSpeaking = ""
			End If
			If IsNumeric(objRsExpLngOther("exlWriting")) And objRsExpLngOther("exlWriting")>"" Then
				sWriting = arrLanguageLevelTitle(objRsExpLngOther("exlWriting"))
			Else
				sWriting = ""
			End If
			ShowUserNoticesViewText sTempLanguage, sReading & " / " & sSpeaking & " / " & sWriting
			objRsExpLngOther.MoveNext
		WEnd
	End If
	objRsExpLngNative.Close
	Set objRsExpLngNative=Nothing	
	objRsExpLngOther.Close
	Set objRsExpLngOther=Nothing	
	
'	ShowUserNoticesViewText "Languages", sOtherLanguages
	ShowUserNoticesViewSpacer 5
	ShowUserNoticesViewFooter
	ShowExpertsBlockFooter "98%", "52", "ex"  & bCvValidForMemberOrExpert

' Other
	If bCvValidForMemberOrExpert=1 And (sMemberships>"" Or sPublications>"" Or sReferences>"") Then
		WriteDataTitle GetLabel(sCvLanguage, "OTHER")

		ShowExpertsBlockSubTitle "98%", "52", "ex"  & bCvValidForMemberOrExpert
		ShowUserNoticesViewHeader "99%", 180
		ShowUserNoticesViewSpacer 4
		If sMemberships>"" Then
			ShowUserNoticesViewText GetLabel(sCvLanguage, "Membership of professional bodies"), ConvertText(sMemberships)
		End If
		If sPublications>"" Then
			ShowUserNoticesViewText GetLabel(sCvLanguage, "Publications"), ConvertText(sPublications)
		End If
		If sReferences>"" Then
			ShowUserNoticesViewText GetLabel(sCvLanguage, "References"), ConvertText(sReferences)
		End If
		ShowUserNoticesViewSpacer 5
		ShowUserNoticesViewFooter
		ShowExpertsBlockFooter "98%", "52", "ex"  & bCvValidForMemberOrExpert
	ElseIf sAvailability>"" Or sPreferences>"" Then
		WriteDataTitle GetLabel(sCvLanguage, "OTHER")
	End If

	If sAvailability>"" Or sPreferences>"" Then 
	ShowExpertsBlockSubTitle "98%", "52", "ex"  & bCvValidForMemberOrExpert
	ShowUserNoticesViewHeader "99%", 180
	ShowUserNoticesViewSpacer 4
	If sAvailability>"" Then 
		ShowUserNoticesViewText GetLabel(sCvLanguage, "Availability"), sAvailability
	End If
	If sPreferences>"" Then 
		ShowUserNoticesViewText GetLabel(sCvLanguage, "Assignment preferences"), sPreferences
	End If
	ShowUserNoticesViewSpacer 5
	ShowUserNoticesViewFooter
	ShowExpertsBlockFooter "98%", "52", "ex"  & bCvValidForMemberOrExpert
	End If

If bCvValidForMemberOrExpert=1 Then
' Permanent address
	If bPermAddress Then
	WriteDataTitle GetLabel(sCvLanguage, "PERMANENT ADDRESS")
	ShowExpertsBlockSubTitle "98%", "52", "ex"  & bCvValidForMemberOrExpert
	ShowUserNoticesViewHeader "99%", 180
	ShowUserNoticesViewSpacer 4
	If sPermAddressStreet>"" Then
		ShowUserNoticesViewText GetLabel(sCvLanguage, "Street"), sPermAddressStreet
	End If
	If sPermAddressCity>"" Then
		ShowUserNoticesViewText GetLabel(sCvLanguage, "City"), sPermAddressCity
	End If
	If sPermAddressPostcode>"" Then
		ShowUserNoticesViewText GetLabel(sCvLanguage, "Postcode"), sPermAddressPostcode
	End If
	ShowUserNoticesViewText GetLabel(sCvLanguage, "Country"), sPermAddressCountry
	ShowUserNoticesViewText GetLabel(sCvLanguage, "Phone"), sPermAddressPhone
	If sPermAddressMobile>"" Then
		ShowUserNoticesViewText GetLabel(sCvLanguage, "Mobile"), sPermAddressMobile
	End If
	If sPermAddressFax>"" Then
		ShowUserNoticesViewText GetLabel(sCvLanguage, "Fax"), sPermAddressFax
	End If
	ShowUserNoticesViewText GetLabel(sCvLanguage, "Email"), sPermAddressEmail
	If sPermAddressWeb>"" Then
		ShowUserNoticesViewText GetLabel(sCvLanguage, "Website"), sPermAddressWeb
	End If
	ShowUserNoticesViewSpacer 5
	ShowUserNoticesViewFooter
	ShowExpertsBlockFooter "98%", "52", "ex"  & bCvValidForMemberOrExpert
	End If

' Current address
	If bCurAddress Then
	WriteDataTitle GetLabel(sCvLanguage, "CURRENT ADDRESS")
	ShowExpertsBlockSubTitle "98%", "52", "ex"  & bCvValidForMemberOrExpert
	ShowUserNoticesViewHeader "99%", 180
	ShowUserNoticesViewSpacer 4
	If sCurAddressStreet>"" Then
		ShowUserNoticesViewText GetLabel(sCvLanguage, "Street"), sCurAddressStreet
	End If
	If sCurAddressCity>"" Then
		ShowUserNoticesViewText GetLabel(sCvLanguage, "City"), sCurAddressCity
	End If
	If sCurAddressPostcode>"" Then
		ShowUserNoticesViewText GetLabel(sCvLanguage, "Postcode"), sCurAddressPostcode
	End If
	ShowUserNoticesViewText GetLabel(sCvLanguage, "Country"), sCurAddressCountry & " "
	ShowUserNoticesViewText GetLabel(sCvLanguage, "Phone"), sCurAddressPhone
	If sCurAddressMobile>"" Then
		ShowUserNoticesViewText GetLabel(sCvLanguage, "Mobile"), sCurAddressMobile
	End If
	If sCurAddressFax>"" Then
		ShowUserNoticesViewText GetLabel(sCvLanguage, "Fax"), sCurAddressFax
	End If
	ShowUserNoticesViewText GetLabel(sCvLanguage, "Email"), sCurAddressEmail
	If sCurAddressWeb>"" Then
		ShowUserNoticesViewText GetLabel(sCvLanguage, "Website"), sCurAddressWeb
	End If
	ShowUserNoticesViewSpacer 5
	ShowUserNoticesViewFooter
	ShowExpertsBlockFooter "98%", "52", "ex"  & bCvValidForMemberOrExpert
	End If
End If
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
