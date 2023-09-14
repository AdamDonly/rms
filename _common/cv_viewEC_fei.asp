<%
'--------------------------------------------------------------------
'
' Expert's CV. View in EC format
'
'--------------------------------------------------------------------
%>
<!--#include file="cv_data.asp"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" type="text/css" href="<% =sHomePath %>styles.css">
<title>CV in European Commission format</title>
</head>

<body bgcolor="#FFFFFF" topmargin=0 leftmargin=0 marginheight=0 marginwidth=0>
<% ShowTopMenu %>

<p class="ttl" align="center"><b>Curriculum Vitae - European Commission format</b></p>

<table width="98%" cellpadding=0 cellspacing=0 border=0 align="center">
<tr><td width="85%" valign="top">
<br />

<%
' Personal information
	WriteTableHeader
	WriteDataRow "1. Surname:", sTitleLastName
	WriteDataRow "2. Name:", sFirstName
	If Len(sBirthPlace)>0 Then
		WriteDataRow "3.&nbsp;Date&nbsp;and&nbsp;place&nbsp;of&nbsp;birth:", ConvertDateForText(sBirthDate, "&nbsp;", "DDMMYYYY") & ", " & sBirthPlace
	Else
		WriteDataRow "3.&nbsp;Date&nbsp;of&nbsp;birth:", ConvertDateForText(sBirthDate, "&nbsp;", "DDMMYYYY")
	End If
	WriteDataRow "4. Nationality:", sNationality 
	If iMaritalStatus>"" And IsNumeric(iMaritalStatus) Then
		WriteDataRow "5. Civil status:", arrMaritalStatusTitle(iMaritalStatus)
	Else
		WriteDataRow "5. Civil status:", ""
	End If
	WriteTableFooter

' Education
	WriteTableHeader
	WriteDataRow "6. Education:", " "
	WriteTableFooter
	Dim sPeriod
	
	WriteGridTableHeader
	WriteGridDataRow 2, "30%#|#70%", "<b>Institution<br>[ Date from - Date to ]</b>#|#<b>Degree(s) or Diploma(s) obtained</b>"

' Training. Start
	While Not objRsExpTrn.Eof
		sPeriod=""
		If Not (IsNull(objRsExpTrn("eduStartDate")) And IsNull(objRsExpTrn("eduEndDate"))) Then
			sPeriod=sPeriod & "[ "
			If Not (IsNull(objRsExpTrn("eduStartDate"))) Then
				sPeriod=sPeriod & ConvertDateForText(objRsExpTrn("eduStartDate"), "&nbsp;", "MMYYYY") & "-"
			End If
			sPeriod=sPeriod & ConvertDateForText(objRsExpTrn("eduEndDate"), "&nbsp;", "MMYYYY") & " ]"
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
			sPeriod=sPeriod & "[ "
			If Not (IsNull(objRsExpEdu("eduStartDate"))) Then
				sPeriod=sPeriod & ConvertDateForText(objRsExpEdu("eduStartDate"), "&nbsp;", "MMYYYY") & " - "
			End If
			sPeriod=sPeriod & ConvertDateForText(objRsExpEdu("eduEndDate"), "&nbsp;", "MMYYYY") & " ]"
		End If
		sEduSubject=""
		If Len(objRsExpEdu("edsDescriptionEng"))>0 And objRsExpEdu("edsDescriptionEng")<>"Other" Then 
			sEduSubject=sEduSubject & Trim(objRsExpEdu("edsDescriptionEng"))
		End If
		If Len(objRsExpEdu("id_EduSubject1Eng"))>0 Then
			sEduSubject=sEduSubject & Trim(" " & objRsExpEdu("id_EduSubject1Eng"))
		End If
		If Len(sEduSubject)>0 Then
			sEduSubject="<br/>" & sEduSubject
		End If
		WriteGridDataRow 2, "30%#|#70%",  objRsExpEdu("InstNameEng") & "<br>" & sPeriod & "#|#" & Trim(objRsExpEdu("edtDescriptionEng") & " " & objRsExpEdu("eduDiploma1Eng")) & sEduSubject
		objRsExpEdu.MoveNext
	WEnd
	objRsExpEdu.Close  
	Set objRsExpEdu=Nothing
	
	WriteTableFooter

' Languages
	WriteTableHeader
	WriteDataRow "7.&nbsp;Languages&nbsp;skills:&nbsp;Indicate&nbsp;competence&nbsp;on&nbsp;a&nbsp;scale&nbsp;of&nbsp;1&nbsp;to&nbsp;5&nbsp; (1&nbsp;<nobr />-&nbsp;excellent;&nbsp;5&nbsp;<nobr />-&nbsp;basic)", " "
	WriteTableFooter

	If (Not objRsExpLngOther.Eof) Or (Not objRsExpLngNative.Eof) Then
		WriteGridTableHeader
		WriteGridDataRow 4, "25%#|#25%#|#25%#|#25%", "<b>Language</b>#|#\qc<b>Reading</b>#|#\qc<b>Speaking</b>#|#\qc<b>Writing</b>"

		While Not objRsExpLngNative.Eof
			WriteGridDataRow 4, "", objRsExpLngNative("lngNameEng") & "#|#\qc" & SetECLanguageLevel(objRsExpLngNative("exlReading")) & "#|#\qc" & SetECLanguageLevel(objRsExpLngNative("exlSpeaking")) & "#|#\qc" & SetECLanguageLevel(objRsExpLngNative("exlWriting"))
			objRsExpLngNative.MoveNext
		WEnd

		While Not objRsExpLngOther.Eof
			WriteGridDataRow 4, "", objRsExpLngOther("lngNameEng") & "#|#\qc" & SetECLanguageLevel(objRsExpLngOther("exlReading")) & "#|#\qc" & SetECLanguageLevel(objRsExpLngOther("exlSpeaking")) & "#|#\qc" & SetECLanguageLevel(objRsExpLngOther("exlWriting"))
			objRsExpLngOther.MoveNext
		WEnd
		WriteTableFooter
	End If
	objRsExpLngNative.Close
	Set objRsExpLngNative=Nothing	
	objRsExpLngOther.Close
	Set objRsExpLngOther=Nothing	


' Membership
	WriteTableHeader
	WriteDataRow "8.&nbsp;Membership&nbsp;of<br>&nbsp; &nbsp; &nbsp; &nbsp;professional&nbsp;bodies:", sMemberships
	WriteTableFooter

' Other skills
	WriteTableHeader
	WriteDataRow "9.&nbsp;Other&nbsp;skills:&nbsp;</b><br>&nbsp; &nbsp; &nbsp; &nbsp;(e.g.&nbsp;computer&nbsp;literacy,&nbsp;etc.)", sOtherSkills
	WriteTableFooter

	WriteTableHeader
	WriteDataRow "10.&nbsp;Present&nbsp;position:", sPosition
	WriteDataRow "11.&nbsp;Years&nbsp;of&nbsp;<br>&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; professional&nbsp;experience:", iProfYears
	WriteDataRow "12.&nbsp;Key&nbsp;qualifications:", sKeyQualification
	WriteTableFooter

' Countries of work experience
	WriteTableHeader
	WriteDataRow "13.&nbsp;Specific&nbsp;experience&nbsp;in&nbsp;non&nbsp;industrialised&nbsp;countries:", " "
	WriteTableFooter

	Set objRsExpCou=GetDataRecordsetSP("usp_ExpCvvECCouSelect", Array( _
		Array(, adInteger, , iCvID)))
	If Not objRsExpCou.Eof Then
	WriteGridTableHeader
	WriteGridDataRow 3, "22%#|#28%#|#50%", "<b>Country</b>#|#<b>Date:</b> from (month/year) to (month/year)#|#<b>Project title:</b>"

	While Not objRsExpCou.Eof

		arrStartDateValues=Split(objRsExpCou(1), "#-#")
		arrEndDateValues=Split(objRsExpCou(2), "#-#")
		arrPrjTitleValues=Split(objRsExpCou(3), "#-#")

		If UBound(arrEndDateValues)>0 Then
			k=0
			WriteGridDataMultiRow 3, UBound(arrEndDateValues)+1, "", objRsExpCou(0) & "#|#" & ConvertSQLDateToText(arrStartDateValues(k), "&nbsp;", "MMYYYY") & " - " & ConvertSQLDateToText(arrEndDateValues(k), "&nbsp;", "MMYYYY") & "#|#" & arrPrjTitleValues(k)
			For k=1 To UBound(arrEndDateValues)
				WriteGridDataRow 2, "", ConvertSQLDateToText(arrStartDateValues(k), "&nbsp;", "MMYYYY") & " - " & ConvertSQLDateToText(arrEndDateValues(k), "&nbsp;", "MMYYYY") & "#|#" & arrPrjTitleValues(k)
			Next
		Else
			WriteGridDataRow 3, "", objRsExpCou(0) & "#|#" & ConvertSQLDateToText(objRsExpCou(1), "&nbsp;", "MMYYYY") & " - " & ConvertSQLDateToText(objRsExpCou(2), "&nbsp;", "MMYYYY") & "#|#" & objRsExpCou(3)
		End If
		objRsExpCou.MoveNext
		Set arrRowsValues=Nothing
	WEnd
	End If
	WriteTableFooter
	objRsExpCou.Close
	Set objRsExpCou=Nothing	


' Employment records
	WriteTableHeader
	WriteDataRow "14.&nbsp;Professional&nbsp;experience:", " "
	WriteTableFooter

	If Not objRsExpWke.Eof Then
		WriteGridTableHeader
		WriteGridDataRow 5, "12%#|#10%#|#14%#|#14%#|#50%", "Date from -<br> Date to" & "#|#" & "Location" & "#|#" & "Company" & "#|#" & "Position" & "#|#" & "Description"

		While Not objRsExpWke.Eof
			objTempRs2=GetDataOutParamsSP("usp_ExpCvvExperienceCouSelect", Array( _
				Array(, adInteger, , iCvID), Array(, adInteger, , objRsExpWke("id_ExpWke")), Array(, adVarChar, 10, "listshort")), Array( _
				Array(, adVarWChar, 4000)))
			sCountries=Replace(objTempRs2(0),"<p>","<p class=""txt"">")
			Set objTempRs2=Nothing			

			dflag=0
			sDescription=ConvertText(objRsExpWke("wkeDescriptionEng"))
			
			Dim sCompanyReference
			sCompanyReference=""
			If Len(objRsExpWke("wkeOrgNameEng"))>1 Then sCompanyReference=sCompanyReference & objRsExpWke("wkeOrgNameEng") & "<br/>"
			If Len(objRsExpWke("wkeBnfNameEng"))>1 Then sCompanyReference=sCompanyReference & objRsExpWke("wkeBnfNameEng") & "<br/>"
			If Len(objRsExpWke("wkeRefFirstName"))>1 Or Len(objRsExpWke("wkeRefLastName"))>1 Then sCompanyReference=sCompanyReference & objRsExpWke("wkeRefFirstName") & " " & objRsExpWke("wkeRefLastName")
			If Len(objRsExpWke("wkeRefPhone"))>1 Or Len(objRsExpWke("wkeRefEmail"))>1 Then sCompanyReference=sCompanyReference & " (" & objRsExpWke("wkeRefPhone") & " " & objRsExpWke("wkeRefEmail") & ")"
			
			Dim sExperienceEndDate
			sExperienceEndDate=""
			If objRsExpWke("wkeEndDateOpen")=1 Then
				sExperienceEndDate=sExperienceEndDate & "Ongoing"
				If IsDate(objRsExpWke("wkeEndDate")) And Not IsNull(objRsExpWke("wkeEndDate")) Then sExperienceEndDate=sExperienceEndDate & " ("
			End If
			If IsDate(objRsExpWke("wkeEndDate")) And Not IsNull(objRsExpWke("wkeEndDate")) Then
				sExperienceEndDate=sExperienceEndDate & ConvertDateForText(objRsExpWke("wkeEndDate"), " ", "MonthYear")
				If objRsExpWke("wkeEndDateOpen")=1 Then sExperienceEndDate=sExperienceEndDate & ")"
			End If
			WriteGridDataRow 5, "12%#|#10%#|#18%#|#18%#|#45%", "" & _
				ConvertDateForText(objRsExpWke("wkeStartDate"), " ", "MonthYear") & " - " & sExperienceEndDate & _
				"#|#" & sCountries & _
				"#|#" & sCompanyReference & _
				"#|#" & objRsExpWke("wkePositionEng") & _
				"#|#" & sDescription 

			objRsExpWke.MoveNext
		WEnd

		WriteTableFooter
	End If
	objRsExpWke.Close
	Set objRsExpWke=Nothing


	WriteTableHeader
	WriteDataRow "15.&nbsp;Others:", " "
	WriteDataRow "15a.&nbsp;Publications and seminars:", ConvertText(sPublications)
	WriteDataRow "15b.&nbsp;References:", ConvertText(sReferences)
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
