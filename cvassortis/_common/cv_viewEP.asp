<%
'--------------------------------------------------------------------
'
' Expert's CV. View in EP format
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
	iLogResult = LogActivity(38, "CVID=" & Cstr(iCvID) & " Format: " & sCvFormat, "", "")
End If
%>
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
		<br />Europass format
		<% If bCvValidForMemberOrExpert=1 Or bCvValidForMemberOrExpert=5 Then %>
		<% Else %>
		<br />Contact details free version
		<% End If %>
		</span>
		</h2><br/>
		<%
	End If

' Personal information
	WriteTableHeader
	
	WriteDataRowWithFormat "<b>Europass<br>Curriculum Vitae</b>", "", "EP"
	WriteDataRowWithFormat "<b>" & GetLabel(sCvLanguage, "Personal information") & "</b>", "", "EP"
	WriteSpaceRow

	WriteDataRowWithFormat GetLabel(sCvLanguage, "Surname(s) / First name(s)"), "<b>" & Trim(sTitleLastName) & ", " & sFirstName & "</b>", "EP"
	WriteDataRowWithFormat GetLabel(sCvLanguage, "Address"), "<b>" & sPermAddress & "</b>", "EP"
	WriteDataRowWithFormat GetLabel(sCvLanguage, "Nationality"), sNationality, "EP"
	WriteDataRowWithFormat GetLabel(sCvLanguage, "Date of birth"), ConvertDateForText(sBirthDate, "&nbsp;", "DayMonthYear"), "EP"
	WriteSpaceRow
	WriteSpaceRow
	WriteDataRowWithFormat "<b>" & GetLabel(sCvLanguage, "Desired employment / Occupational field") & "</b>", "", "EP"
	WriteSpaceRow
	WriteSpaceRow
	
' Employment records
	WriteDataRowWithFormat "<b>" & GetLabel(sCvLanguage, "Work experience") & "</b>", "", "EP"
	WriteSpaceRow

	While Not objRsExpWke.Eof
		sCountries = GetExpertExperienceCountryList(iCvID, objRsExpWke("id_ExpWke"), sCvLanguage)

		dflag=0
		sDescription=ConvertText(objRsExpWke("wkeDescriptionEng"))

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

		WriteDataRowWithFormat GetLabel(sCvLanguage, "Dates"), "<b>" & ConvertDateForText(objRsExpWke("wkeStartDate"), " ", "MonthYear") & " - " & sExperienceEndDate & "</b>", "EP"
		WriteDataRowWithFormat GetLabel(sCvLanguage, "Occupation or position held"), objRsExpWke("wkePositionEng"), "EP"
		WriteDataRowWithFormat GetLabel(sCvLanguage, "Main activities and responsibilities"), sDescription, "EP"
		WriteDataRowWithFormat GetLabel(sCvLanguage, "Name and address of employer"), objRsExpWke("wkeOrgNameEng") &  "<br>" & sCountries, "EP"
		WriteDataRowWithFormat GetLabel(sCvLanguage, "Type of business or sector"), "", "EP"
		WriteSpaceRow
		WriteSpaceRow
		WriteSpaceRow

		objRsExpWke.MoveNext
	WEnd

	objRsExpWke.Close
	Set objRsExpWke=Nothing

	WriteSpaceRow
	
' Employment records
	WriteDataRowWithFormat "<b>" & GetLabel(sCvLanguage, "Education and training") & "</b>", "", "EP"
	WriteSpaceRow

' Training
	Dim sPeriod
	While Not objRsExpTrn.Eof
		sPeriod=""
		If Not (IsNull(objRsExpTrn("eduStartDate")) And IsNull(objRsExpTrn("eduEndDate"))) Then
			If Not (IsNull(objRsExpTrn("eduStartDate"))) Then
				sPeriod=sPeriod & ConvertDateForText(objRsExpTrn("eduStartDate"), "&nbsp;", "MonthYear") & " - "
			End If
			sPeriod=sPeriod & ConvertDateForText(objRsExpTrn("eduEndDate"), "&nbsp;", "MonthYear")
		End If

		WriteDataRowWithFormat GetLabel(sCvLanguage, "Dates"), "<b>" & sPeriod & "</b>", "EP"
		WriteDataRowWithFormat GetLabel(sCvLanguage, "Title of qualification awarded"), objRsExpTrn("eduDiploma1Eng"), "EP"
		WriteDataRowWithFormat GetLabel(sCvLanguage, "Principal subjects/occupational skills covered"), Trim(objRsExpTrn("edsDescriptionEng") & " " & objRsExpTrn("id_EduSubject1Eng")), "EP"
		WriteDataRowWithFormat GetLabel(sCvLanguage, "Name and type of organisation providing education and training"), objRsExpTrn("InstNameEng"), "EP"
		WriteDataRowWithFormat GetLabel(sCvLanguage, "Level in national or international classification"), "", "EP"
		WriteSpaceRow
		WriteSpaceRow
		WriteSpaceRow

		objRsExpTrn.MoveNext
	WEnd
	objRsExpTrn.Close
	Set objRsExpTrn=Nothing
' Education

	While Not objRsExpEdu.Eof
		sPeriod=""
		If Not (IsNull(objRsExpEdu("eduStartDate")) And IsNull(objRsExpEdu("eduEndDate"))) Then
			If Not (IsNull(objRsExpEdu("eduStartDate"))) Then
				sPeriod=sPeriod & ConvertDateForText(objRsExpEdu("eduStartDate"), "&nbsp;", "MonthYear") & " - "
			End If
			sPeriod=sPeriod & ConvertDateForText(objRsExpEdu("eduEndDate"), "&nbsp;", "MonthYear")
		End If

		WriteDataRowWithFormat GetLabel(sCvLanguage, "Dates"), "<b>" & sPeriod & "</b>", "EP"
		WriteDataRowWithFormat GetLabel(sCvLanguage, "Title of qualification awarded"), objRsExpEdu("eduDiploma1Eng"), "EP"
		WriteDataRowWithFormat GetLabel(sCvLanguage, "Principal subjects/occupational skills covered"), Trim(objRsExpEdu("edsDescriptionEng") & " " & objRsExpEdu("id_EduSubject1Eng")), "EP"
		WriteDataRowWithFormat GetLabel(sCvLanguage, "Name and type of organisation providing education and training"), objRsExpEdu("InstNameEng"), "EP"
		WriteDataRowWithFormat GetLabel(sCvLanguage, "Level in national or international classification"), "", "EP"
		WriteSpaceRow
		WriteSpaceRow
		WriteSpaceRow

		objRsExpEdu.MoveNext
	WEnd
	objRsExpEdu.Close  
	Set objRsExpEdu=Nothing
	
	WriteDataRowWithFormat "<b>" & GetLabel(sCvLanguage, "Personal skills and competences") & "</b>", "", "EP"
	WriteSpaceRow

' Languages
	Dim sNativeLanguage
	sNativeLanguage = ""
	While Not objRsExpLngNative.Eof
		sNativeLanguage = sNativeLanguage & objRsExpLngNative("lngNameEng") & ", "
		objRsExpLngNative.MoveNext
	WEnd
	If Len(sNativeLanguage)>2 Then sNativeLanguage=Left(sNativeLanguage, Len(sNativeLanguage)-2)

	WriteDataRowWithFormat GetLabel(sCvLanguage, "Mother tongue(s)"), "<b>" & sNativeLanguage & "</b>", "EP"
	WriteSpaceRow
	
	If (Not objRsExpLngOther.Eof) Or (Not objRsExpLngNative.Eof) Then
		WriteDataRowWithFormat GetLabel(sCvLanguage, "Other language(s)"), "", "EP"

		While Not objRsExpLngOther.Eof
			On Error Resume Next
				sTempLanguage = ReplaceIfEmpty(objRsExpLngOther("lngName" & sCvLanguage), objRsExpLngOther("lngNameEng"))
			On Error GoTo 0
			WriteDataRowWithFormat "", "<b>" & sTempLanguage & "</b>", "EP"
			WriteDataRowWithFormat GetLabel(sCvLanguage, "Understanding"), arrLanguageLevelTitle(objRsExpLngOther("exlReading")), "EP"
			WriteDataRowWithFormat GetLabel(sCvLanguage, "Speaking"), arrLanguageLevelTitle(objRsExpLngOther("exlSpeaking")), "EP"
			WriteDataRowWithFormat GetLabel(sCvLanguage, "Writing"), arrLanguageLevelTitle(objRsExpLngOther("exlWriting")), "EP"
			WriteSpaceRow
			objRsExpLngOther.MoveNext
		WEnd
	End If
	objRsExpLngNative.Close
	Set objRsExpLngNative=Nothing	
	objRsExpLngOther.Close
	Set objRsExpLngOther=Nothing	

	WriteDataRowWithFormat GetLabel(sCvLanguage, "Social skills and competences"), "", "EP"
	WriteSpaceRow

	WriteDataRowWithFormat GetLabel(sCvLanguage, "Organisational skills and competences"), "", "EP"
	WriteSpaceRow

	WriteDataRowWithFormat GetLabel(sCvLanguage, "Technical skills and competences"), sKeyQualification, "EP"
	WriteSpaceRow

	WriteDataRowWithFormat GetLabel(sCvLanguage, "Computer skills and competences"), "", "EP"
	WriteSpaceRow

	WriteDataRowWithFormat GetLabel(sCvLanguage, "Artistic skills and competences"), "", "EP"
	WriteSpaceRow

	WriteDataRowWithFormat GetLabel(sCvLanguage, "Other skills and competences"), "", "EP"
	WriteSpaceRow
	WriteSpaceRow

	WriteDataRowWithFormat GetLabel(sCvLanguage, "Driving licence(s)"), "", "EP"
	WriteSpaceRow
	WriteSpaceRow

	WriteDataRowWithFormat "<b>" & GetLabel(sCvLanguage, "Additional information") & "</b>", "", "EP"
	WriteDataRowWithFormat GetLabel(sCvLanguage, "Publications"), ConvertText(sPublications), "EP"
	WriteSpaceRow
	WriteDataRowWithFormat GetLabel(sCvLanguage, "Memberships"), ConvertText(sMemberships), "EP"
	WriteSpaceRow
	WriteDataRowWithFormat GetLabel(sCvLanguage, "References"), ConvertText(sReferences), "EP"
	WriteSpaceRow
	WriteSpaceRow

	WriteDataRowWithFormat "<b>" & GetLabel(sCvLanguage, "Annexes") & "</b>", "", "EP"
	WriteSpaceRow
	WriteSpaceRow

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
</div>
	<!-- footer -->
	<!--#include file="../_template/page.footer.asp"-->
<% CloseDBConnection %>
</body>
<!--#include file="../_template/html.footer.asp"-->
