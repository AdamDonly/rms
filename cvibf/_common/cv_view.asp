<%
'--------------------------------------------------------------------
'
' Expert's CV. View in assortis.com format
' With or without contact details
'
'--------------------------------------------------------------------
%>
<!--#include file="cv_data.asp"-->
<%
If bCvValidForMemberOrExpert=0 Then
	Response.Redirect(Replace(sScriptFullName, sScriptFileName, "cv_preview.asp"))
End If

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
<!--#include virtual="/_template/html.header.asp"-->
<body>
<div id="holder">
	<!-- header -->
	<!--#include virtual="/_template/page.header.asp"-->

	<!-- content -->
	<div id="content" class="workscreen">
	<% 
	If Not bIsMyCV Then 
		%><h2 class="service_title">Curriculum Vitae. <span class="service_slogan">Expert ID: <% =objExpertDB.DatabaseCode %><%=iCvID%>
		<% If bCvValidForMemberOrExpert=1 Or bCvValidForMemberOrExpert=5 Then %>
		<% Else %>
		<br />Contact details free version
		<% End If %>
		</span>
		</h2><br/>
	<%
	End If

' Personal information
	WriteDataTitle GetLabel(sCvLanguage, "Personal information")
	ShowExpertsBlockSubTitle "98%", "52", "ex"  & bCvValidForMemberOrExpert

	If bCvValidForMemberOrExpert=1 Or bCvValidForMemberOrExpert=5 Then
		ShowUserNoticesViewText GetLabel(sCvLanguage, "Personal title"), sTitle
		ShowUserNoticesViewText GetLabel(sCvLanguage, "First name"), sFirstName
		If sMiddleName>"" Then ShowUserNoticesViewText GetLabel(sCvLanguage, "Middle name"), sMiddleName
		ShowUserNoticesViewText GetLabel(sCvLanguage, "Family name"), sLastName
		ShowUserNoticesViewText GetLabel(sCvLanguage, "Date of birth"), ConvertDateForText(sBirthDate, "&nbsp;", "DayMonthYear")
		If sBirthPlace>"" Then ShowUserNoticesViewText GetLabel(sCvLanguage, "Place of birth"), sBirthPlace
	Else
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
			sEduSubject=sEduSubject & Trim(objRsExpEdu("edsDescriptionEng"))
		End If
		If Len(objRsExpEdu("id_EduSubject1Eng"))>0 Then
			sEduSubject=sEduSubject & Trim(" " & objRsExpEdu("id_EduSubject1Eng"))
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
	If bCvValidForMemberOrExpert=1 Or bCvValidForMemberOrExpert=5 Then
		ShowUserNoticesViewText GetLabel(sCvLanguage, "Status"), sProfessionalStatus
		ShowUserNoticesViewText GetLabel(sCvLanguage, "Profession"), sProfession
		ShowUserNoticesViewText GetLabel(sCvLanguage, "Current position"), sPosition
	End If
	ShowUserNoticesViewText GetLabel(sCvLanguage, "Key qualifications"), sKeyQualification
	ShowUserNoticesViewText GetLabel(sCvLanguage, "Other skills"), sOtherSkills
	ShowUserNoticesViewText GetLabel(sCvLanguage, "Years of professional experience"), iProfYears

	ShowUserNoticesViewFooter
	ShowExpertsBlockFooter "98%", "52", "ex"  & bCvValidForMemberOrExpert

' Employment records
	If Not objRsExpWke.Eof _
	And sScriptFileName<>"cv_verify.asp" Then
	WriteDataTitle GetLabel(sCvLanguage, "Employment record and completed projects")
	While Not objRsExpWke.Eof
		
		ShowExpertsBlockSubTitle "98%", "52", "ex"  & bCvValidForMemberOrExpert
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
		ShowUserNoticesViewText GetLabel(sCvLanguage, "Position / Responsibility"), "<b>" & HighlightKeywords(objRsExpWke("wkePositionEng"), sSearchKeywordsHighlight) & "</b>"
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
		ShowUserNoticesViewText GetLabel(sCvLanguage, "Description of tasks"), HighlightKeywords(ConvertText(objRsExpWke("wkeDescriptionEng")), sSearchKeywordsHighlight)

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
	
'	ShowUserNoticesViewText "Languages", sOtherLanguages
	ShowUserNoticesViewFooter
	ShowExpertsBlockFooter "98%", "52", "ex"  & bCvValidForMemberOrExpert

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
		ShowUserNoticesViewText GetLabel(sCvLanguage, "Availability"), sAvailability
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
</div>
	<!-- footer -->
	<!--#include file="../_template/page.footer.asp"-->
<% CloseDBConnection %>
</body>
<!--#include file="../_template/html.footer.asp"-->
