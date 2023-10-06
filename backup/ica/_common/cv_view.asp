<%
'--------------------------------------------------------------------
'
' Expert's CV. View in assortis.com format
' With or without contact details
'
'--------------------------------------------------------------------
sSearchKeywordsHighlight=Trim(Request.QueryString("txt") + " " + Request.QueryString("srch_queryadd"))
sSearchKeywordsHighlight=Replace(sSearchKeywordsHighlight, " AND ", " ")
sSearchKeywordsHighlight=Replace(sSearchKeywordsHighlight, " OR ", " ")
sSearchKeywordsHighlight=Replace(sSearchKeywordsHighlight, " NOT ", " ")
sSearchKeywordsHighlight=Replace(sSearchKeywordsHighlight, " NEAR ", " ")
sSearchKeywordsHighlight=Replace(sSearchKeywordsHighlight, "  ", " ")

If Len(sSearchKeywordsHighlight)>2 Then
	arrSearchKeywordsHighlight=Split(sSearchKeywordsHighlight, " ")
End If
%>
<!--#include file="cv_data.asp"-->
<%
' Log:
If bCvValidForMemberOrExpert > 0 Then
	' 38 - View CV Full
	iLogResult = LogActivity(38, "CVID=" & Cstr(iCvID) & " Format: ASR", "", "")
End If
%>
<!--#include virtual="/_common/_class/main.asp"-->
<!--#include virtual="/_common/_class/project.asp"-->
<!--#include virtual="/_common/_class/expert.asp"-->
<!--#include virtual="/_common/_class/expert.project.asp"-->
<!--#include virtual="/_common/_class/expert.cv.language.asp"-->

<!--#include virtual="/_template/html.header.asp"-->
<body class="cv-view">
	<!-- header -->
	<!--#include virtual="/_template/page.header.asp"-->

	<!-- content -->
	<div id="content" class="workscreen">
	<% 
	If Not bIsMyCV Then 
		RenderVacanciesSelector iCvID, sCvUID, objExpertDB.DatabaseCodePrimary
		
		%><div id="hdrUpdatedList" class="colCCCCCC uprCse f17 spc01 botMrgn10">
			<span class="service_title">Curriculum Vitae.</span> Expert ID: <% =objExpertDB.DatabaseCodePrimary %><%=iCvID%>
			<% If bCvValidForMemberOrExpert = aClientSecurityCvViewEnabled Or bCvValidForMemberOrExpert = aClientSecurityCvViewAll Then %>
			<% Else %>
				<br /><span style="font-size:0.7em;text-transform:none">Contact details free version</span>
			<% End If %>
		</div>
		<% 
	Else
		%><div class="colCCCCCC uprCse f17 spc01 botMrgn10"><span class="service_title">Curriculum Vitae</span></div>
	<%
	End If

' Personal information
	WriteDataTitle GetLabel(sCvLanguage, "Personal information")
	ShowExpertsBlockSubTitle "98%", "52", "ex"  & bCvValidForMemberOrExpert

	If bCvValidForMemberOrExpert = aClientSecurityCvViewEnabled Or bCvValidForMemberOrExpert = aClientSecurityCvViewAll Then

		ShowUserNoticesViewText GetLabel(sCvLanguage, "Personal title"), sTitle
		ShowUserNoticesViewText GetLabel(sCvLanguage, "First name"), sFirstName
		If sMiddleName>"" Then ShowUserNoticesViewText GetLabel(sCvLanguage, "Middle name"), sMiddleName
		ShowUserNoticesViewText GetLabel(sCvLanguage, "Family name"), sLastName
		ShowUserNoticesViewText GetLabel(sCvLanguage, "Date of birth"), ConvertDateForText(sBirthDate, "&nbsp;", "DayMonthYear")
		If sBirthPlace>"" Then ShowUserNoticesViewText GetLabel(sCvLanguage, "Place of birth"), sBirthPlace
	Else
		' In preview show only the year of birth
		If sBirthDate>"" And IsDate(sBirthDate) Then
			ShowUserNoticesViewText GetLabel(sCvLanguage, "Year of birth"), Year(sBirthDate)
		End If
	End If
	ShowUserNoticesViewText GetLabel(sCvLanguage, "Nationality"), sNationality
	If iGender>0 And IsNumeric(iGender) Then ShowUserNoticesViewText GetLabel(sCvLanguage, "Gender"), arrGenderTitle(iGender)
	If iMaritalStatus>0 And IsNumeric(iMaritalStatus) Then ShowUserNoticesViewText GetLabel(sCvLanguage, "Marital status"), arrMaritalStatusTitle(iMaritalStatus)

	ShowUserNoticesViewFooter
	ShowExpertsBlockFooter "98%", "52", "ex"  & bCvValidForMemberOrExpert

' Education	
	If Not objRsExpEdu.Eof Then
	WriteDataTitle GetLabel(sCvLanguage, "Education")
	While Not objRsExpEdu.Eof
		ShowExpertsBlockSubTitle "98%", "52", "ex"  & bCvValidForMemberOrExpert
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
			sEduSubject=sEduSubject & " " & Trim(objRsExpEdu("edsDescriptionEng"))
		End If
		If Len(objRsExpEdu("id_EduSubject1Eng"))>0 Then
			sEduSubject=sEduSubject & " " & Trim(" " & objRsExpEdu("id_EduSubject1Eng"))
		End If
		If Len(sEduSubject)>0 Then
			ShowUserNoticesViewText GetLabel(sCvLanguage, "Subject"), sEduSubject
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
		ShowUserNoticesViewText GetLabel(sCvLanguage, "Skills / Qualification"), objRsExpTrn("eduOtherEng")
		ShowUserNoticesViewText GetLabel(sCvLanguage, "Title"), objRsExpTrn("eduDiploma1Eng")
		ShowUserNoticesViewText GetLabel(sCvLanguage, "Start date"), ConvertDateForText(objRsExpTrn("eduStartDate"), "&nbsp;", "MMYYYY")
		ShowUserNoticesViewText GetLabel(sCvLanguage, "End date"), ConvertDateForText(objRsExpTrn("eduEndDate"), "&nbsp;", "MMYYYY")
		sAchievements=objRsExpTrn("eduDescriptionEng")
		If sAchievements>"" Then
			ShowUserNoticesViewText GetLabel(sCvLanguage, "Achievements"), sAchievements
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
	
	If bCvValidForMemberOrExpert=aClientSecurityCvViewEnabled Or bCvValidForMemberOrExpert=aClientSecurityCvViewAll Then
		ShowUserNoticesViewText GetLabel(sCvLanguage, "Status"), sProfessionalStatus
		ShowUserNoticesViewText GetLabel(sCvLanguage, "Profession"), HighlightKeywords(sProfession, sSearchKeywordsHighlight)
		ShowUserNoticesViewText GetLabel(sCvLanguage, "Current position"), HighlightKeywords(sPosition, sSearchKeywordsHighlight)
		ShowUserNoticesViewText GetLabel(sCvLanguage, "Key qualifications"), HighlightKeywords(sKeyQualification, sSearchKeywordsHighlight)
		ShowUserNoticesViewText GetLabel(sCvLanguage, "Other skills"), HighlightKeywords(sOtherSkills, sSearchKeywordsHighlight)
	Else
		sKeyQualification = RemoveContactDetailsIfNoAccess(sKeyQualification)
		ShowUserNoticesViewText GetLabel(sCvLanguage, "Key qualifications"), sKeyQualification
	End If
	ShowUserNoticesViewText GetLabel(sCvLanguage, "Years of professional experience"), iProfYears

	ShowUserNoticesViewFooter
	ShowExpertsBlockFooter "98%", "52", "ex"  & bCvValidForMemberOrExpert

' Employment records
	If Not objRsExpWke.Eof _
	And sScriptFileName<>"cv_verify.asp" Then
	WriteDataTitle GetLabel(sCvLanguage, "Employment record and completed projects")
	While Not objRsExpWke.Eof
		
		ShowExpertsBlockSubTitle "98%", "52", "ex"  & bCvValidForMemberOrExpert
		On Error Resume Next
			If Not (IsNull(objRsExpWke("wkeStartDate")) Or IsNull(objRsExpWke("wkeEndDate"))) Then
				Dim iDurationMonths, iDurationYears, sDuration
				iDurationMonths = DateDiff("m", objRsExpWke("wkeStartDate"), objRsExpWke("wkeEndDate"))+1
				' Projects of x years and more then 10 months could be rounded to x+1 years
				iDurationYears = DateDiff("yyyy", objRsExpWke("wkeStartDate"), DateAdd("m", 3, objRsExpWke("wkeEndDate")))
				sDuration = ""
				If iDurationMonths > 24 Then
					sDuration = ShowEntityPlural(iDurationYears, "year", "years", "&nbsp;")
				Else
					sDuration = ShowEntityPlural(iDurationMonths, "month", "months", "&nbsp;")
				End If
			End If

			Dim sExperienceEndDate
			sExperienceEndDate=""
			If objRsExpWke("wkeEndDateOpen") = 1 Then
				sExperienceEndDate = sExperienceEndDate & GetLabel(sCvLanguage, "Ongoing")
				If IsDate(objRsExpWke("wkeEndDate")) And Not IsNull(objRsExpWke("wkeEndDate")) Then sExperienceEndDate = sExperienceEndDate & " ("
			End If
			If IsDate(objRsExpWke("wkeEndDate")) And Not IsNull(objRsExpWke("wkeEndDate")) Then
				sExperienceEndDate = sExperienceEndDate &  ConvertDateForText(objRsExpWke("wkeEndDate"), "&nbsp;", "MMYYYY")
				If objRsExpWke("wkeEndDateOpen") = 1 Then sExperienceEndDate = sExperienceEndDate & ")"
			End If

			ShowUserNoticesViewText GetLabel(sCvLanguage, "Project title"), "<b>" & HighlightKeywords(RemoveContactDetailsIfNoAccess(objRsExpWke("wkePrjTitleEng")), sSearchKeywordsHighlight) & "</b>"
			If bCvValidForMemberOrExpert=aClientSecurityCvViewEnabled Or bCvValidForMemberOrExpert=aClientSecurityCvViewAll Then
				ShowUserNoticesViewText GetLabel(sCvLanguage, "Start date"), ConvertDateForText(objRsExpWke("wkeStartDate"), "&nbsp;", "MMYYYY")
				ShowUserNoticesViewText GetLabel(sCvLanguage, "End date"), sExperienceEndDate
				ShowUserNoticesViewText GetLabel(sCvLanguage, "Contractor"), objRsExpWke("wkeOrgNameEng")
			Else
				If sDuration > "" Then
					ShowUserNoticesViewText "</b>Assignment duration", sDuration
				End If
			End If
		ShowUserNoticesViewText GetLabel(sCvLanguage, "Position / Responsibility"), "<b>" & HighlightKeywords(objRsExpWke("wkePositionEng"), sSearchKeywordsHighlight) & "</b>"
		ShowUserNoticesViewText GetLabel(sCvLanguage, "Beneficiary"), objRsExpWke("wkeBnfNameEng")
		
		sCountries = GetExpertExperienceCountryGroupedList(iCvID, objRsExpWke("id_ExpWke"), sCvLanguage)
		sSectors = GetExpertExperienceSectorGroupedList(iCvID, objRsExpWke("id_ExpWke"), sCvLanguage)

		objTempRs2=GetDataOutParamsSPWithConn(objConnCustom, "usp_ExpCvvExperienceDonSelect", Array( _
			Array(, adInteger, , iCvID), Array(, adInteger, , objRsExpWke("id_ExpWke")), Array(, adVarChar, 10, "list")), Array( _
			Array(, adVarWChar, 4000)))
		sDonors=objTempRs2(0)
		Set objTempRs2=Nothing

		ShowUserNoticesViewText GetLabel(sCvLanguage, "Funding agency"), sDonors
		ShowUserNoticesViewText GetLabel(sCvLanguage, "Countries"), sCountries
		ShowUserNoticesViewText GetLabel(sCvLanguage, "Sectors"), sSectors
			
		If bCvValidForMemberOrExpert=aClientSecurityCvViewEnabled Or bCvValidForMemberOrExpert=aClientSecurityCvViewAll Then
			If objRsExpWke("wkeClientRefEng")>"" Then ShowUserNoticesViewText GetLabel(sCvLanguage, "Client references"), objRsExpWke("wkeClientRefEng")
		End If
		ShowUserNoticesViewText GetLabel(sCvLanguage, "Description of tasks"), HighlightKeywords(ConvertText(RemoveContactDetailsIfNoAccess(objRsExpWke("wkeDescriptionEng"))), sSearchKeywordsHighlight)

		ShowUserNoticesViewText " ", "&nbsp;"
		On Error GoTo 0
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
	
'	ShowUserNoticesViewText GetLabel(sCvLanguage, "Languages"), sOtherLanguages
	ShowUserNoticesViewFooter
	ShowExpertsBlockFooter "98%", "52", "ex"  & bCvValidForMemberOrExpert

' Other
	If (bCvValidForMemberOrExpert=1 Or bCvValidForMemberOrExpert=5) Then
		If sMemberships>"" Or sPublications>"" Or sReferences>"" Or sAvailability>"" Or sPreferences>"" Then 
			WriteDataTitle GetLabel(sCvLanguage, "Other")

			ShowExpertsBlockSubTitle "98%", "52", "ex"  & bCvValidForMemberOrExpert
		If sMemberships>"" Then
			ShowUserNoticesViewText GetLabel(sCvLanguage, "Membership of professional bodies"), ConvertText(sMemberships)
		End If
		If sPublications>"" Then
			ShowUserNoticesViewText GetLabel(sCvLanguage, "Publications"), ConvertText(sPublications)
		End If
		If sReferences>"" Then
			ShowUserNoticesViewText GetLabel(sCvLanguage, "References"), ConvertText(sReferences)
		End If
			
			ShowUserNoticesViewFooter
			ShowExpertsBlockFooter "98%", "52", "ex"  & bCvValidForMemberOrExpert
		End If

		If sAvailability>"" Or sPreferences>"" Then 
			WriteDataTitle GetLabel(sCvLanguage, "Availability & mission preferences")
			ShowExpertsBlockSubTitle "98%", "52", "ex"  & bCvValidForMemberOrExpert
			If sAvailability>"" Then 
				ShowUserNoticesViewText GetLabel(sCvLanguage, "Availability"), RemoveContactDetailsIfNoAccess(sAvailability)
			End If
			If sPreferences>"" Then 
				ShowUserNoticesViewText GetLabel(sCvLanguage, "Assignment preferences"), sPreferences
			End If
			ShowUserNoticesViewFooter
			ShowExpertsBlockFooter "98%", "52", "ex"  & bCvValidForMemberOrExpert
		End If
	
	
	' Permanent address
		WriteDataTitle GetLabel(sCvLanguage, "Permanent address")
		ShowExpertsBlockSubTitle "98%", "52", "ex"  & bCvValidForMemberOrExpert
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
		ShowUserNoticesViewFooter
		ShowExpertsBlockFooter "98%", "52", "ex"  & bCvValidForMemberOrExpert
	
	' Current address
		If bCurAddress Then
		WriteDataTitle GetLabel(sCvLanguage, "Current address")
		ShowExpertsBlockSubTitle "98%", "52", "ex"  & bCvValidForMemberOrExpert
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
		ShowUserNoticesViewFooter
		ShowExpertsBlockFooter "98%", "52", "ex"  & bCvValidForMemberOrExpert
		End If
	End If
%>

	</div>
	<div id="rightspace">
	<!-- feature boxes -->
	<% 
	If Not bIsMyCV Then
		ShowExpCVFeatureBox 
	Else
		ShowTopExpCVFeatureBox
	End If
	%>
	</div>
	<!-- footer -->
	<!--#include file="../_template/page.footer.asp"-->
<% CloseDBConnection %>
</body>
<!--#include file="../_template/html.footer.asp"-->
