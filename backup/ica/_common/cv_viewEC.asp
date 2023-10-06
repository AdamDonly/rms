<%
'--------------------------------------------------------------------
'
' Expert's CV. View in EC format
'
'--------------------------------------------------------------------
%>
<!--#include file="cv_data.asp"-->
<%
' Log:
If bCvValidForMemberOrExpert > 0 Then
	' 38 - View CV Full
	iLogResult = LogActivity(38, "CVID=" & Cstr(iCvID) & " Format: " & sCvFormat, "", "")
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
		%><div id="hdrUpdatedList" class="colCCCCCC uprCse f17 spc01 botMrgn10"><span class="service_title">Curriculum Vitae.</span> Expert ID: <% =objExpertDB.DatabaseCode %><%=iCvID%>
		<br /><span style="font-size:0.7em;text-transform:none;color:#999">European Commission format</span>
		<% If bCvValidForMemberOrExpert=1 Or bCvValidForMemberOrExpert=5 Then %>
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
	WriteTableHeader
	WriteDataRow "1. " & GetLabel(sCvLanguage, "Family name") & ":", sTitleLastName
	WriteDataRow "2. " & GetLabel(sCvLanguage, "First name") & ":", sFirstName
	WriteDataRow "3. " & GetLabel(sCvLanguage, "Date of birth") & ":", ConvertDateForText(sBirthDate, " ", "DayMonthYear")
	WriteDataRow "4. " & GetLabel(sCvLanguage, "Nationality") & ":", sNationality 
	If iMaritalStatus>"" And IsNumeric(iMaritalStatus) Then
		WriteDataRow "5. " &  GetLabel(sCvLanguage, "Civil status") & ":", arrMaritalStatusTitle(iMaritalStatus)
	Else
		WriteDataRow "5. " &  GetLabel(sCvLanguage, "Civil status") & ":", ""
	End If
	WriteSpaceRow
		WriteDataRow "&nbsp;&nbsp;&nbsp;&nbsp;" & GetLabel(sCvLanguage, "Address") & ":<br /><br />&nbsp;&nbsp;&nbsp;&nbsp;(" &  GetLabel(sCvLanguage, "Phone") & " / " & GetLabel(sCvLanguage, "email") & ")", sPermAddress

	WriteTableFooter

' Education
	WriteTableHeader
	WriteDataRow "6. " &  GetLabel(sCvLanguage, "Education") & ":", " "
	WriteTableFooter
	Dim sPeriod
	
	WriteGridTableHeader
	WriteGridDataRow 2, "30%#|#70%", GetLabel(sCvLanguage, "Institution") & "<br />(" &  GetLabel(sCvLanguage, "Date from") & " - " &  GetLabel(sCvLanguage, "Date to") & ")#|#" & GetLabel(sCvLanguage, "Degree(s) or Diploma(s) obtained") & ":"
' Training. Start
	While Not objRsExpTrn.Eof
		sPeriod=""
		If Not (IsNull(objRsExpTrn("eduStartDate")) And IsNull(objRsExpTrn("eduEndDate"))) Then
			sPeriod=sPeriod & "("
			If Not (IsNull(objRsExpTrn("eduStartDate"))) Then
				sPeriod=sPeriod & ConvertDateForText(objRsExpTrn("eduStartDate"), "&nbsp;", "MMYYYY") & "-"
			End If
			sPeriod=sPeriod & ConvertDateForText(objRsExpTrn("eduEndDate"), "&nbsp;", "MMYYYY") & ")"
		End If
		WriteGridDataRow 2, "30%#|#70%",  objRsExpTrn("InstNameEng") & " " & sPeriod & "#|#" & objRsExpTrn("eduDiploma1Eng") & Trim(objRsExpTrn("edsDescriptionEng") & " " & objRsExpTrn("id_EduSubject1Eng"))

		objRsExpTrn.MoveNext
	WEnd
	objRsExpTrn.Close
	Set objRsExpTrn=Nothing
' Training. End

	While Not objRsExpEdu.Eof
		sPeriod=""
		If Not (IsNull(objRsExpEdu("eduStartDate")) And IsNull(objRsExpEdu("eduEndDate"))) Then
			sPeriod=sPeriod & "("
			If Not (IsNull(objRsExpEdu("eduStartDate"))) Then
				sPeriod=sPeriod & ConvertDateForText(objRsExpEdu("eduStartDate"), "&nbsp;", "MMYYYY") & " - "
			End If
			sPeriod=sPeriod & ConvertDateForText(objRsExpEdu("eduEndDate"), "&nbsp;", "MMYYYY") & ")"
		End If
		sEduSubject=""
		If Len(objRsExpEdu("edsDescriptionEng"))>0 And objRsExpEdu("edsDescriptionEng")<>"Other" Then 
			sEduSubject=sEduSubject & Trim(objRsExpEdu("edsDescriptionEng"))
		End If
		If Len(objRsExpEdu("id_EduSubject1Eng"))>0 Then
			sEduSubject=sEduSubject & Trim(" " & objRsExpEdu("id_EduSubject1Eng"))
		End If
		If Len(sEduSubject)>0 Then
			sEduSubject="<br />" & sEduSubject
		End If
		WriteGridDataRow 2, "30%#|#70%",  objRsExpEdu("InstNameEng") & "<br />" & sPeriod & "#|#" & Trim(objRsExpEdu("edtDescriptionEng") & " " & objRsExpEdu("eduDiploma1Eng")) & sEduSubject
		objRsExpEdu.MoveNext
	WEnd
	objRsExpEdu.Close  
	Set objRsExpEdu=Nothing
	
	WriteGridTableFooter

' Languages
	WriteTableHeader
	WriteDataRow1Column "7. " & GetLabel(sCvLanguage, "Languages skills EC"), " "
	WriteTableFooter

	If (Not objRsExpLngOther.Eof) Or (Not objRsExpLngNative.Eof) Then
		WriteGridTableHeader
		WriteGridDataRow 4, "25%#|#25%#|#25%#|#25%", GetLabel(sCvLanguage, "Language") & "#|#\qc " &  GetLabel(sCvLanguage, "Reading") & "#|#\qc " &  GetLabel(sCvLanguage, "Speaking") & "#|#\qc " &  GetLabel(sCvLanguage, "Writing")

		While Not objRsExpLngNative.Eof
			On Error Resume Next
				sTempLanguage = ReplaceIfEmpty(objRsExpLngNative("lngName" & sCvLanguage), objRsExpLngNative("lngNameEng"))
			On Error GoTo 0
			WriteGridDataRow 4, "", sTempLanguage & "#|#\qc" & SetECLanguageLevel(objRsExpLngNative("exlReading")) & "#|#\qc" & SetECLanguageLevel(objRsExpLngNative("exlSpeaking")) & "#|#\qc" & SetECLanguageLevel(objRsExpLngNative("exlWriting"))
			objRsExpLngNative.MoveNext
		WEnd

		While Not objRsExpLngOther.Eof
			On Error Resume Next
				sTempLanguage = ReplaceIfEmpty(objRsExpLngOther("lngName" & sCvLanguage), objRsExpLngOther("lngNameEng"))
			On Error GoTo 0
			WriteGridDataRow 4, "", sTempLanguage & "#|#\qc" & SetECLanguageLevel(objRsExpLngOther("exlReading")) & "#|#\qc" & SetECLanguageLevel(objRsExpLngOther("exlSpeaking")) & "#|#\qc" & SetECLanguageLevel(objRsExpLngOther("exlWriting"))
			objRsExpLngOther.MoveNext
		WEnd
		WriteGridTableFooter
	End If
	objRsExpLngNative.Close
	Set objRsExpLngNative=Nothing	
	objRsExpLngOther.Close
	Set objRsExpLngOther=Nothing	

' Membership
	WriteTableHeader
	WriteDataRow "8. " & GetLabel(sCvLanguage, "Membership of professional bodies") & ":", HighlightKeywords(sMemberships, sSearchKeywordsHighlight)
	WriteTableFooter

' Other skills
	WriteTableHeader
	WriteDataRow "9. " & GetLabel(sCvLanguage, "Other skills") & ":", HighlightKeywords(sOtherSkills, sSearchKeywordsHighlight)
	WriteTableFooter

	WriteTableHeader
	WriteDataRow "10. " &  GetLabel(sCvLanguage, "Present position") & ":", sPosition
	WriteDataRow "11. " &  GetLabel(sCvLanguage, "Years of professional experience") & ":", iProfYears
	WriteDataRow "12. " &  GetLabel(sCvLanguage, "Key qualifications") & ":", HighlightKeywords(sKeyQualification, sSearchKeywordsHighlight)
	WriteTableFooter

' Countries of work experience
	WriteTableHeader
	WriteDataRow1Column "13. " &  GetLabel(sCvLanguage, "Specific experience in the region") & ":", " "
	WriteTableFooter

	Set objRsExpCou=GetDataRecordsetSPWithConn(objConnCustom, "usp_ExpertEcCountrySelect", Array( _
		Array(, adInteger, , iCvID), _
		Array(, adVarChar, 80, sCvLanguage)))
	If Not objRsExpCou.Eof Then
	WriteGridTableHeader
	WriteGridDataRow 2, "46%#|#55%", GetLabel(sCvLanguage, "Country") & "#|#" & GetLabel(sCvLanguage, "Date from") & " - " & GetLabel(sCvLanguage, "Date to")

	While Not objRsExpCou.Eof
		arrStartDateValues=Split(objRsExpCou(1), "#-#")
		arrEndDateValues=Split(objRsExpCou(2), "#-#")
		arrPrjTitleValues=Split(objRsExpCou(3), "#-#")

		If UBound(arrStartDateValues)>0 Then
			k=0
			Dim MaxRows
			MaxRows=UBound(arrStartDateValues)
			'If UBound(arrPrjTitleValues)>MaxRows Then MaxRows=UBound(arrPrjTitleValues)
			
			WriteGridDataMultiRow 2, UBound(arrEndDateValues), "", objRsExpCou(0) & "#|#" & ConvertSQLDateToText(arrStartDateValues(k), "&nbsp;", "MMYYYY") & " - " & ConvertSQLDateToText(arrEndDateValues(k), "&nbsp;", "MMYYYY")
			On Error Resume Next
			For k=1 To MaxRows-1
				WriteGridDataRow 2, "", ConvertSQLDateToText(arrStartDateValues(k), "&nbsp;", "MMYYYY") & " - " & ConvertSQLDateToText(arrEndDateValues(k), "&nbsp;", "MMYYYY")
			Next
			On Error GoTo 0
		Else
			WriteGridDataRow 2, "", objRsExpCou(0) & "#|#" & ConvertSQLDateToText(objRsExpCou(1), "&nbsp;", "MMYYYY") & " - " & ConvertSQLDateToText(objRsExpCou(2), "&nbsp;", "MMYYYY")
		End If
		objRsExpCou.MoveNext
		Set arrRowsValues=Nothing
	WEnd
	End If
	WriteGridTableFooter
	objRsExpCou.Close
	Set objRsExpCou=Nothing	


' Employment records
	WriteTableHeader
	WriteDataRow "14. " & GetLabel(sCvLanguage, "Professional experience") & ":", " "
	WriteTableFooter

	If Not objRsExpWke.Eof Then
		WriteGridTableHeader
		WriteGridDataRow 5, "12%#|#10%#|#14%#|#14%#|#50%", GetLabel(sCvLanguage, "Date from") & " -<br /> " &  GetLabel(sCvLanguage, "Date to") & "#|#" & GetLabel(sCvLanguage, "Location") & "#|#" &  GetLabel(sCvLanguage, "Company and reference person") & "#|#" & GetLabel(sCvLanguage, "Position") & "#|#" & GetLabel(sCvLanguage, "Description")

		While Not objRsExpWke.Eof
			sCountries = GetExpertExperienceCountryList(iCvID, objRsExpWke("id_ExpWke"), sCvLanguage)

			dflag=0
			sDescription=ConvertText(objRsExpWke("wkeDescriptionEng"))
			If Len(objRsExpWke("wkePrjTitleEng"))>0 Then
				sDescription=ConvertText(objRsExpWke("wkePrjTitleEng")) & "<br />" & sDescription
			End If
			
			Dim sCompanyReference
			sCompanyReference=""
			If Len(objRsExpWke("wkeOrgNameEng"))>1 Then 
				sCompanyReference=sCompanyReference & objRsExpWke("wkeOrgNameEng") & "<br />"
			Else
				If Len(objRsExpWke("wkeBnfNameEng"))>1 Then sCompanyReference=sCompanyReference & objRsExpWke("wkeBnfNameEng") & "<br />"
			End If
			If Len(objRsExpWke("wkeRefFirstName"))>1 Or Len(objRsExpWke("wkeRefLastName"))>1 Then sCompanyReference=sCompanyReference & objRsExpWke("wkeRefFirstName") & " " & objRsExpWke("wkeRefLastName")
			If Len(objRsExpWke("wkeRefPhone"))>1 Or Len(objRsExpWke("wkeRefEmail"))>1 Then sCompanyReference=sCompanyReference & " (" & Trim(objRsExpWke("wkeRefPhone") & " " & objRsExpWke("wkeRefEmail")) & ")"
			
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
			WriteGridDataRow 5, "12%#|#10%#|#14%#|#14%#|#50%", "" & _
				ConvertDateForText(objRsExpWke("wkeStartDate"), " ", "MonthYear") & " - " & sExperienceEndDate & _
				"#|#" & sCountries & _
				"#|#" & sCompanyReference & _
				"#|#" & objRsExpWke("wkePositionEng") & _
				"#|#" & HighlightKeywords(sDescription, sSearchKeywordsHighlight)

			objRsExpWke.MoveNext
		WEnd

		WriteGridTableFooter
	End If
	objRsExpWke.Close
	Set objRsExpWke=Nothing


	WriteTableHeader
	WriteDataRow1Column "15. " &  GetLabel(sCvLanguage, "Other relevant information (e.g., Publications)") & ":", " "
	WriteDataRow1Column ConvertText(sPublications), ""
	If Len(sReferences)>1 Then
		WriteDataRow "", ""
		WriteDataRow GetLabel(sCvLanguage, "References") & ":", ConvertText(sReferences)
	End If
	WriteTableFooter

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
