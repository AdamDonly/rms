<%
' Personal information
	WriteDataTitle GetLabel(sCvLanguage, "Personal information")
	ShowExpertsBlockSubTitle "98%", "52", "ex"  & bCvValidForMemberOrExpert

	If bCvValidForMemberOrExpert=1 Or bCvValidForMemberOrExpert=5 Then
		ShowUserNoticesViewText "</b>" & GetLabel(sCvLanguage, "Personal title"), sTitle
		ShowUserNoticesViewText "</b>" & GetLabel(sCvLanguage, "First name"), sFirstName
		If sMiddleName>"" Then ShowUserNoticesViewText "</b>" & GetLabel(sCvLanguage, "Middle name"), sMiddleName
		ShowUserNoticesViewText "</b>" & GetLabel(sCvLanguage, "Family name"), sLastName
		ShowUserNoticesViewText "</b>" & GetLabel(sCvLanguage, "Date of birth"), ConvertDateForText(sBirthDate, "&nbsp;", "DDMMYYYY")
		If sBirthPlace>"" Then ShowUserNoticesViewText "</b>" & GetLabel(sCvLanguage, "Place of birth"), sBirthPlace
	End If
	ShowUserNoticesViewText "</b>" & GetLabel(sCvLanguage, "Nationality"), sNationality
	If iGender>0 And IsNumeric(iGender) Then ShowUserNoticesViewText "</b>" & GetLabel(sCvLanguage, "Gender"), arrGenderTitle(iGender)
	If iMaritalStatus>0 And IsNumeric(iMaritalStatus) Then ShowUserNoticesViewText "</b>" & GetLabel(sCvLanguage, "Marital status"), arrMaritalStatusTitle(iMaritalStatus)

	ShowUserNoticesViewFooter
	ShowExpertsBlockFooter "98%", "52", "ex"  & bCvValidForMemberOrExpert

' Education	
	If Not objRsExpEdu.Eof Then
	WriteDataTitle GetLabel(sCvLanguage, "Education")
	While Not objRsExpEdu.Eof
		ShowExpertsBlockSubTitle "98%", "52", "ex"  & bCvValidForMemberOrExpert
		If objRsExpEdu("InstNameEng")>"" Then ShowUserNoticesViewText "</b>" & GetLabel(sCvLanguage, "Institution"), objRsExpEdu("InstNameEng")
		If objRsExpEdu("InstLocationEng")>"" Then ShowUserNoticesViewText "</b>" & GetLabel(sCvLanguage, "Location"), objRsExpEdu("InstLocationEng")
		If IsDate(objRsExpEdu("eduStartDate")) Then ShowUserNoticesViewText "</b>" & GetLabel(sCvLanguage, "Start date"), ConvertDateForText(objRsExpEdu("eduStartDate"), "&nbsp;", "MMYYYY")
		If IsDate(objRsExpEdu("eduEndDate")) Then ShowUserNoticesViewText "</b>" & GetLabel(sCvLanguage, "End date"), ConvertDateForText(objRsExpEdu("eduEndDate"), "&nbsp;", "MMYYYY")
		If Not IsNull(objRsExpEdu("eduDiploma")) Then
			ShowUserNoticesViewText "</b>" & GetLabel(sCvLanguage, "Type of diploma"), Trim(EducationTypeTitleByID(objRsExpEdu("eduDiploma")) & " " & objRsExpEdu("eduDiploma1Eng"))
		Else
			ShowUserNoticesViewText "</b>" & GetLabel(sCvLanguage, "Type of diploma"), objRsExpEdu("eduDiploma1Eng")
		End If
		sEduSubject=""
		If Len(objRsExpEdu("edsDescriptionEng"))>0 And objRsExpEdu("edsDescriptionEng")<>"Other" Then 
			sEduSubject=sEduSubject & " " & Trim(objRsExpEdu("edsDescriptionEng"))
		End If
		If Len(objRsExpEdu("id_EduSubject1Eng"))>0 Then
			sEduSubject=sEduSubject & " " & Trim(" " & objRsExpEdu("id_EduSubject1Eng"))
		End If
		If Len(sEduSubject)>0 Then
			ShowUserNoticesViewText "</b>" & GetLabel(sCvLanguage, "Subject"), sEduSubject
		End If
		ShowUserNoticesViewFooter
	objRsExpEdu.MoveNext
	WEnd
	ShowExpertsBlockFooter "98%", "52", "ex"  & bCvValidForMemberOrExpert
	End If
	objRsExpEdu.Close
	Set objRsExpEdu=Nothing

' Training
	If not objRsExpTrn.Eof Then
	WriteDataTitle GetLabel(sCvLanguage, "Training")
	While Not objRsExpTrn.Eof
		ShowExpertsBlockSubTitle "98%", "52", "ex"  & bCvValidForMemberOrExpert
		ShowUserNoticesViewText "</b>" & GetLabel(sCvLanguage, "Skills / Qualification"), objRsExpTrn("eduOtherEng")
		ShowUserNoticesViewText "</b>" & GetLabel(sCvLanguage, "Title"), objRsExpTrn("eduDiploma1Eng")
		ShowUserNoticesViewText "</b>" & GetLabel(sCvLanguage, "Start date"), ConvertDateForText(objRsExpTrn("eduStartDate"), "&nbsp;", "MMYYYY")
		ShowUserNoticesViewText "</b>" & GetLabel(sCvLanguage, "End date"), ConvertDateForText(objRsExpTrn("eduEndDate"), "&nbsp;", "MMYYYY")
		sAchievements=objRsExpTrn("eduDescriptionEng")
		If sAchievements>"" Then
			ShowUserNoticesViewText "</b>" & GetLabel(sCvLanguage, "Achievements"), sAchievements
		End If
		ShowUserNoticesViewFooter
	objRsExpTrn.MoveNext
	WEnd
	ShowExpertsBlockFooter "98%", "52", "ex"  & bCvValidForMemberOrExpert
	End If
	objRsExpTrn.Close
	Set objRsExpTrn=Nothing

' Professional experience
	WriteDataTitle GetLabel(sCvLanguage, "Professional experience")
	ShowExpertsBlockSubTitle "98%", "52", "ex"  & bCvValidForMemberOrExpert
	If bCvValidForMemberOrExpert=1 Or bCvValidForMemberOrExpert=5 Then
		ShowUserNoticesViewText "</b>" & GetLabel(sCvLanguage, "Status"), sProfessionalStatus
		ShowUserNoticesViewText "</b>" & GetLabel(sCvLanguage, "Profession"), sProfession
		ShowUserNoticesViewText "</b>" & GetLabel(sCvLanguage, "Current position"), sPosition
	End If
	ShowUserNoticesViewText "</b>" & GetLabel(sCvLanguage, "Key qualifications"), HighlightKeywords(sKeyQualification, sSearchKeywordsHighlight)
	ShowUserNoticesViewText "</b>" & GetLabel(sCvLanguage, "Other skills"), sOtherSkills
	ShowUserNoticesViewText "</b>" & GetLabel(sCvLanguage, "Years of professional experience"), iProfYears
	ShowUserNoticesViewFooter
	ShowExpertsBlockFooter "98%", "52", "ex"  & bCvValidForMemberOrExpert

' Employment records
	If not objRsExpWke.Eof Then
	WriteDataTitle GetLabel(sCvLanguage, "Employment record and completed projects")
	While Not objRsExpWke.Eof
		
		ShowExpertsBlockSubTitle "98%", "52", "ex"  & bCvValidForMemberOrExpert
		ShowUserNoticesViewText "</b>" & GetLabel(sCvLanguage, "Project title"), "<b>" & HighlightKeywords(objRsExpWke("wkePrjTitleEng"), sSearchKeywordsHighlight) & "</b>"
		ShowUserNoticesViewText "</b>" & GetLabel(sCvLanguage, "Start date"), ConvertDateForText(objRsExpWke("wkeStartDate"), "&nbsp;", "MMYYYY")
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
		ShowUserNoticesViewText "</b>" & GetLabel(sCvLanguage, "End date"), sExperienceEndDate
		ShowUserNoticesViewText "</b>" & GetLabel(sCvLanguage, "Company / Organisation"), objRsExpWke("wkeOrgNameEng")
		ShowUserNoticesViewText "</b>" & GetLabel(sCvLanguage, "Position / Responsibility"), HighlightKeywords(objRsExpWke("wkePositionEng"), sSearchKeywordsHighlight)
		ShowUserNoticesViewText "</b>" & GetLabel(sCvLanguage, "Beneficiary"), objRsExpWke("wkeBnfNameEng")

	sCountries = GetExpertExperienceCountryGroupedList(iCvID, objRsExpWke("id_ExpWke"), sCvLanguage)
	sSectors = GetExpertExperienceSectorGroupedList(iCvID, objRsExpWke("id_ExpWke"), sCvLanguage)

objTempRs2=GetDataOutParamsSP("usp_ExpCvvExperienceDonSelect", Array( _
	Array(, adInteger, , iCvID), Array(, adInteger, , objRsExpWke("id_ExpWke")), Array(, adVarChar, 10, "list")), Array( _
	Array(, adVarWChar, 4000)))
sDonors=objTempRs2(0)
Set objTempRs2=Nothing

		ShowUserNoticesViewText "</b>" & GetLabel(sCvLanguage, "Funding agencies"), sDonors
		ShowUserNoticesViewText "</b>" & GetLabel(sCvLanguage, "Countries"), sCountries
		ShowUserNoticesViewText "</b>" & GetLabel(sCvLanguage, "Sectors"), sSectors
		If objRsExpWke("wkeClientRefEng")>"" Then ShowUserNoticesViewText "</b>" & GetLabel(sCvLanguage, "Client references"), objRsExpWke("wkeClientRefEng")
		ShowUserNoticesViewText "</b>" & GetLabel(sCvLanguage, "Description of tasks"), HighlightKeywords(ConvertText(objRsExpWke("wkeDescriptionEng")), sSearchKeywordsHighlight)

		ShowUserNoticesViewFooter
	objRsExpWke.MoveNext
	WEnd
	ShowExpertsBlockFooter "98%", "52", "ex"  & bCvValidForMemberOrExpert
	End If
	objRsExpWke.Close
	Set objRsExpWke=Nothing

' Languages
	WriteDataTitle GetLabel(sCvLanguage, "Languages skills")
	ShowExpertsBlockSubTitle "98%", "52", "ex"  & bCvValidForMemberOrExpert
	ShowUserNoticesViewText "</b>" & GetLabel(sCvLanguage, "Language"), GetLabel(sCvLanguage, "Reading") & " / " & GetLabel(sCvLanguage, "Speaking") & " / " & GetLabel(sCvLanguage, "Writing")
	
	If (Not objRsExpLngOther.Eof) Or (Not objRsExpLngNative.Eof) Then
		While Not objRsExpLngNative.Eof
			On Error Resume Next
				sTempLanguage = ReplaceIfEmpty(objRsExpLngNative("lngName" & sCvLanguage), objRsExpLngNative("lngNameEng"))
			On Error GoTo 0
			ShowUserNoticesViewText "</b>" & sTempLanguage, GetLabel(sCvLanguage, "Native")
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
			ShowUserNoticesViewText "</b>" & sTempLanguage, sReading & " / " & sSpeaking & " / " & sWriting
			objRsExpLngOther.MoveNext
		WEnd
	End If
	objRsExpLngNative.Close
	Set objRsExpLngNative=Nothing	
	objRsExpLngOther.Close
	Set objRsExpLngOther=Nothing	
	
'	ShowUserNoticesViewText "</b>Languages", sOtherLanguages
	ShowUserNoticesViewFooter
	ShowExpertsBlockFooter "98%", "52", "ex"  & bCvValidForMemberOrExpert

' Other
	If (bCvValidForMemberOrExpert=1 Or bCvValidForMemberOrExpert=5) Then
		If sMemberships>"" Or sPublications>"" Or sReferences>"" Or sAvailability>"" Or sPreferences>"" Then 
			WriteDataTitle GetLabel(sCvLanguage, "Other")

			ShowExpertsBlockSubTitle "98%", "52", "ex"  & bCvValidForMemberOrExpert
			If sMemberships>"" Then
				ShowUserNoticesViewText "</b>" & GetLabel(sCvLanguage, "Membership of professional bodies"), ConvertText(sMemberships)
			End If
			If sPublications>"" Then
				ShowUserNoticesViewText "</b>" & GetLabel(sCvLanguage, "Publications"), ConvertText(sPublications)
			End If
			If sReferences>"" Then
				ShowUserNoticesViewText "</b>" & GetLabel(sCvLanguage, "References"), ConvertText(sReferences)
			End If
			
			ShowUserNoticesViewFooter
			ShowExpertsBlockFooter "98%", "52", "ex"  & bCvValidForMemberOrExpert
		End If
	End If

	If sAvailability>"" Or sPreferences>"" Then 
		WriteDataTitle GetLabel(sCvLanguage, "Availability & preferences")
		ShowExpertsBlockSubTitle "98%", "52", "ex"  & bCvValidForMemberOrExpert
		If sAvailability>"" Then 
			ShowUserNoticesViewText "</b>" & GetLabel(sCvLanguage, "Availability"), sAvailability

		End If
		If sPreferences>"" Then 
			ShowUserNoticesViewText "</b>" & GetLabel(sCvLanguage, "Assignment preferences"), sPreferences
		End If
		ShowUserNoticesViewFooter
		ShowExpertsBlockFooter "98%", "52", "ex"  & bCvValidForMemberOrExpert
	End If
	
	
If bCvValidForMemberOrExpert=1 Or bCvValidForMemberOrExpert=5 Then
' Permanent address
	WriteDataTitle GetLabel(sCvLanguage, "Permanent address")
	ShowExpertsBlockSubTitle "98%", "52", "ex"  & bCvValidForMemberOrExpert
	If sPermAddressStreet>"" Then
		ShowUserNoticesViewText "</b>" & GetLabel(sCvLanguage, "Street"), sPermAddressStreet
	End If
	If sPermAddressCity>"" Then
		ShowUserNoticesViewText "</b>" & GetLabel(sCvLanguage, "City"), sPermAddressCity
	End If
	If sPermAddressPostcode>"" Then
		ShowUserNoticesViewText "</b>" & GetLabel(sCvLanguage, "Postcode"), sPermAddressPostcode
	End If
	ShowUserNoticesViewText "</b>" & GetLabel(sCvLanguage, "Country"), sPermAddressCountry
	ShowUserNoticesViewText "</b>" & GetLabel(sCvLanguage, "Phone"), sPermAddressPhone
	If sPermAddressMobile>"" Then
		ShowUserNoticesViewText "</b>" & GetLabel(sCvLanguage, "Mobile"), sPermAddressMobile
	End If
	If sPermAddressFax>"" Then
		ShowUserNoticesViewText "</b>" & GetLabel(sCvLanguage, "Fax"), sPermAddressFax
	End If
	ShowUserNoticesViewText "</b>" & GetLabel(sCvLanguage, "Email"), sPermAddressEmail
	If sPermAddressWeb>"" Then
		ShowUserNoticesViewText "</b>" & GetLabel(sCvLanguage, "Website"), sPermAddressWeb
	End If
	ShowUserNoticesViewFooter
	ShowExpertsBlockFooter "98%", "52", "ex"  & bCvValidForMemberOrExpert

' Current address
	If bCurAddress Then
	WriteDataTitle GetLabel(sCvLanguage, "Current address")
	ShowExpertsBlockSubTitle "98%", "52", "ex"  & bCvValidForMemberOrExpert
	If sCurAddressStreet>"" Then
		ShowUserNoticesViewText "</b>" & GetLabel(sCvLanguage, "Street"), sCurAddressStreet
	End If
	If sCurAddressCity>"" Then
		ShowUserNoticesViewText "</b>" & GetLabel(sCvLanguage, "City"), sCurAddressCity
	End If
	If sCurAddressPostcode>"" Then
		ShowUserNoticesViewText "</b>" & GetLabel(sCvLanguage, "Postcode"), sCurAddressPostcode
	End If
	ShowUserNoticesViewText "</b>" & GetLabel(sCvLanguage, "Country"), sCurAddressCountry & " "
	ShowUserNoticesViewText "</b>" & GetLabel(sCvLanguage, "Phone"), sCurAddressPhone
	If sCurAddressMobile>"" Then
		ShowUserNoticesViewText "</b>" & GetLabel(sCvLanguage, "Mobile"), sCurAddressMobile
	End If
	If sCurAddressFax>"" Then
		ShowUserNoticesViewText "</b>" & GetLabel(sCvLanguage, "Fax"), sCurAddressFax
	End If
	ShowUserNoticesViewText "</b>" & GetLabel(sCvLanguage, "Email"), sCurAddressEmail
	If sCurAddressWeb>"" Then
		ShowUserNoticesViewText "</b>" & GetLabel(sCvLanguage, "Website"), sCurAddressWeb
	End If
	ShowUserNoticesViewFooter
	ShowExpertsBlockFooter "98%", "52", "ex"  & bCvValidForMemberOrExpert
	End If
End If
%>
