<%
'--------------------------------------------------------------------
'
' Expert's CV. Save in EC format
'
'--------------------------------------------------------------------
Response.Buffer = True
%>
<!--#include file="cv_data.asp"-->
<%
Dim MaxRows

sFileType = LCase(Request.QueryString("ftype"))
If sFileType="doc" Or sFileType="prn" Then
	Response.ContentType = "application/vnd.ms-word"
	If sFileType="doc" then Response.AddHeader "Content-Disposition", "attachment; filename=" & sFileName & ".rtf"
	If sFileType="prn" then Response.AddHeader "Content-Disposition", "filename=" & sFileName & ".rtf"
End If
sLastName=Replace(ConvertText2RTF(sLastName), "      ", "")
sFullName=Replace(ConvertText2RTF(sFullName), "      ", "")

Response.Write("{\rtf1\ansi{\fonttbl{\f1\fswiss\fcharset0\fprq2 Arial;}{\f2\fswiss\fcharset186\fprq2 Arial;}}" & vbCrLf)
Response.Write("{\colortbl;\red0\green0\blue0;\red0\green0\blue0;\red0\green255\blue255;\red0\green255\blue0;\red255\green0\blue255;\red204\green0\blue0;\red255\green255\blue0;\red255\green255\blue255;\red0\green0\blue128;\red0\green128\blue128;\red0\green128\blue0;\red128\green0\blue128;\red128\green0\blue0;\red128\green128\blue0;\red128\green128\blue128;\red192\green192\blue192;}" & vbCrLf)
Response.Write("{\stylesheet{\ql \li0\ri0\widctlpar\faauto\rin0\lin0\itap0 \fs24\lang2057\langfe1033\cgrid\langnp2057\langfenp1033 \snext0 Normal;}{\*\cs10 \additive Default Paragraph Font;}{\s15\ql \li0\ri0\widctlpar\tqc\tx4153\tqr\tx8306\aspalpha\aspnum\faauto\adjustright\rin0\lin0\itap0 \fs24\lang2057\langfe1033\cgrid\langnp2057\langfenp1033 \sbasedon0 \snext15 header;}{\s16\ql \li0\ri0\widctlpar\tqc\tx4153\tqr\tx8306\aspalpha\aspnum\faauto\adjustright\rin0\lin0\itap0 \fs24\lang2057\langfe1033\cgrid\langnp2057\langfenp1033 \sbasedon0 \snext16 footer;}{\*\cs17 \additive \sbasedon10 page number;}}" & vbCrLf)
Response.Write("{\info{\title CV of " & sFullName & "}{\author assortis.com}{\company assortis.com}{\subject European Commission format}{\category Expert CV}{\keywords " & sProfession & "}{\doccomm Downloaded: " & ConvertDateForText(Date(), " ", "DDMMYYYY") & "}}\paperw11907\paperh16840\margl1797\margr1797\margt789\margb689\viewkind1\viewscale100\titlepg " & vbCrLf)

'Response.Write("{\footerf \pard\plain \s16\qr \li0\ri0\widctlpar\brdrt\brdrs\brdrw10\brsp20 \tqc\tx4153\tqr\tx8306\aspalpha\aspnum\faauto\adjustright\rin0\lin0\itap0 \fs24\lang2057\langfe1033\cgrid\langnp2057\langfenp1033 {\f1\fs16\cf9\lang2060\langfe1033\langnp2060 Page\~: }{\field{\*\fldinst {\cs17\f1\fs16\cf9  PAGE }}{\fldrslt {\cs17\f1\fs16\cf9\lang1024\langfe1024\noproof 2}}}{\f1\fs16\cf9\lang2060\langfe1033\langnp2060 \par }}{\*\pnseclvl1\pnucrm\pnstart1\pnindent720\pnhang{\pntxta .}}{\*\pnseclvl2\pnucltr\pnstart1\pnindent720\pnhang{\pntxta .}}{\*\pnseclvl3\pndec\pnstart1\pnindent720\pnhang{\pntxta .}}{\*\pnseclvl4\pnlcltr\pnstart1\pnindent720\pnhang{\pntxta )}}" & vbCrLf)
'Response.Write("{\footer \pard\plain \s16\qr \li0\ri0\widctlpar\brdrt\brdrs\brdrw10\brsp20 \tqc\tx4153\tqr\tx8306\aspalpha\aspnum\faauto\adjustright\rin0\lin0\itap0 \fs24\lang2057\langfe1033\cgrid\langnp2057\langfenp1033 {\f1\fs16\cf9\lang2060\langfe1033\langnp2060 Page\~: }{\field{\*\fldinst {\cs17\f1\fs16\cf9  PAGE }}{\fldrslt {\cs17\f1\fs16\cf9\lang1024\langfe1024\noproof 2}}}{\f1\fs16\cf9\lang2060\langfe1033\langnp2060 \par }}{\*\pnseclvl1\pnucrm\pnstart1\pnindent720\pnhang{\pntxta .}}{\*\pnseclvl2\pnucltr\pnstart1\pnindent720\pnhang{\pntxta .}}{\*\pnseclvl3\pndec\pnstart1\pnindent720\pnhang{\pntxta .}}{\*\pnseclvl4\pnlcltr\pnstart1\pnindent720\pnhang{\pntxta )}}" & vbCrLf)
Response.Write("\qc\f1\fs18\cf2 \b CURRICULUM VITAE\b0\par" & vbCrLf)
Response.Write("{" & vbCrLf)
Response.Write("\par\ql\f1\fs18\cf2" & vbCrLf)

' Personal information
	WriteTableHeader
	WriteDataRow "1. " &  GetLabel(sCvLanguage, "Family name") & ":", sTitleLastName
	WriteDataRow "2. " &  GetLabel(sCvLanguage, "First name") & ":", sFirstName
	WriteDataRow "3. " &  GetLabel(sCvLanguage, "Date of birth") & ":", ConvertDateForText(sBirthDate, " ", "DayMonthYear")
	WriteDataRow "4. " &  GetLabel(sCvLanguage, "Nationality") & ":", sNationality 
	If iMaritalStatus>"" And IsNumeric(iMaritalStatus) Then
		WriteDataRow "5. " &  GetLabel(sCvLanguage, "Civil status") & ":", arrMaritalStatusTitle(iMaritalStatus)
	Else
		WriteDataRow "5. " &  GetLabel(sCvLanguage, "Civil status") & ":", ""
	End If

	' Some clients might ask to drop the address from EC format
	WriteSpaceRow
	WriteDataRow "\~\~\~\~" & GetLabel(sCvLanguage, "Address") & ":\line \line \~ \~ (" &  GetLabel(sCvLanguage, "Phone")& " / " & GetLabel(sCvLanguage, "email") & ")", ConvertText2RTF(sPermAddress)

	WriteTableFooter

' Education
	WriteTableHeader
	WriteGridDataRow 1, "100%", "6. " &  GetLabel(sCvLanguage, "Education") & ":", 0, "0"
	WriteTableFooter
	Dim sPeriod
	
	WriteGridTableHeader
	WriteGridDataRow 2, "31%#|#70%", GetLabel(sCvLanguage, "Institution") & "\line " & GetLabel(sCvLanguage, "Date from") & " - " &  GetLabel(sCvLanguage, "Date to") & "#|#" & GetLabel(sCvLanguage, "Degree(s) or Diploma(s) obtained") & ":", 1, "1#|#1"
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
		WriteGridDataRow 2, "31%#|#70%",  objRsExpTrn("InstNameEng") & " " & sPeriod & "#|#" & objRsExpTrn("eduDiploma1Eng") & Trim(objRsExpTrn("edsDescriptionEng") & " " & objRsExpTrn("id_EduSubject1Eng")), 1, "1#|#1"

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
			sEduSubject="\line " & sEduSubject
		End If
		WriteGridDataRow 2, "31%#|#70%",  objRsExpEdu("InstNameEng") & "\line " & sPeriod & "#|#" & Trim(objRsExpEdu("edtDescriptionEng") & " " & objRsExpEdu("eduDiploma1Eng")) & sEduSubject, 1, "1#|#1"
		objRsExpEdu.MoveNext
	WEnd
	WriteTableFooter
	objRsExpEdu.Close  
	Set objRsExpEdu=Nothing
	

' Languages
	WriteTableHeader
	WriteGridDataRow 1, "100%", "7. " & GetLabel(sCvLanguage, "Languages skills EC"), 0, "0"
	WriteTableFooter

	If (Not objRsExpLngOther.Eof) Or (Not objRsExpLngNative.Eof) Then
		WriteGridTableHeader
		WriteGridDataRow 4, "31%#|#23%#|#23%#|#23%", GetLabel(sCvLanguage, "Language") & "#|#\qc " &  GetLabel(sCvLanguage, "Reading") & "#|#\qc " &  GetLabel(sCvLanguage, "Speaking") & "#|#\qc " &  GetLabel(sCvLanguage, "Writing"), 1, "1#|#1#|#1#|#1"

		While Not objRsExpLngNative.Eof
			On Error Resume Next
				sTempLanguage = ReplaceIfEmpty(objRsExpLngNative("lngName" & sCvLanguage), objRsExpLngNative("lngNameEng"))
			On Error GoTo 0
			WriteGridDataRow 4, "31%#|#23%#|#23%#|#23%", sTempLanguage & "#|#\qc " & SetECLanguageLevel(objRsExpLngNative("exlReading")) & "#|#\qc " & SetECLanguageLevel(objRsExpLngNative("exlSpeaking")) & "#|#\qc " & SetECLanguageLevel(objRsExpLngNative("exlWriting")), 1, "1#|#1#|#1#|#1"
			objRsExpLngNative.MoveNext
		WEnd

		While Not objRsExpLngOther.Eof
			On Error Resume Next
				sTempLanguage = ReplaceIfEmpty(objRsExpLngOther("lngName" & sCvLanguage), objRsExpLngOther("lngNameEng"))
			On Error GoTo 0
			WriteGridDataRow 4, "31%#|#23%#|#23%#|#23%", sTempLanguage & "#|#\qc " & SetECLanguageLevel(objRsExpLngOther("exlReading")) & "#|#\qc " & SetECLanguageLevel(objRsExpLngOther("exlSpeaking")) & "#|#\qc " & SetECLanguageLevel(objRsExpLngOther("exlWriting")), 1, "1#|#1#|#1#|#1"
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
	WriteGridDataRow 2, "31%#|#70%", "8. " &  GetLabel(sCvLanguage, "Membership of professional bodies") & ":" & "#|#" & sMemberships, 0, "0#|#0"
	WriteTableFooter

' Other skills
	WriteTableHeader
	WriteGridDataRow 2, "31%#|#70%", "9. " &  GetLabel(sCvLanguage, "Other skills") & ":\~\b0 " & "#|#" & sOtherSkills, 0, "0#|#0"
	WriteTableFooter

	WriteTableHeader
	WriteGridDataRow 2, "31%#|#70%", "10. " &  GetLabel(sCvLanguage, "Present position") & ":\line " & "#|#" & sPosition, 0, "0#|#0"
	WriteGridDataRow 2, "31%#|#70%", "11. " &  GetLabel(sCvLanguage, "Years of professional experience") & ":\line " & "#|#" & " \line " & iProfYears, 0, "0#|#0"
	WriteGridDataRow 2, "31%#|#70%", "12. " &  GetLabel(sCvLanguage, "Key qualifications") & ":\line " & "#|#" & sKeyQualification, 0, "0#|#0"
	WriteTableFooter

' Countries of work experience
	WriteTableHeader
	WriteGridDataRow 1, "100%", "13. " &  GetLabel(sCvLanguage, "Specific experience in the region") & ":", 0, "0"
	WriteTableFooter

	Set objRsExpCou=GetDataRecordsetSP("usp_ExpertEcCountrySelect", Array( _
		Array(, adInteger, , iCvID), _
		Array(, adVarChar, 80, sCvLanguage)))
	If Not objRsExpCou.Eof Then
	WriteGridTableHeader
	WriteGridDataRow 2, "46%#|#55%", GetLabel(sCvLanguage, "Country") & "#|#" & GetLabel(sCvLanguage, "Date from") & " - " & GetLabel(sCvLanguage, "Date to"), 1, "1#|#1"

	While Not objRsExpCou.Eof
		arrStartDateValues=Split(objRsExpCou(1), "#-#")
		arrEndDateValues=Split(objRsExpCou(2), "#-#")
		arrPrjTitleValues=Split(objRsExpCou(3), "#-#")

		If UBound(arrStartDateValues)>0 Then
			k=0
			MaxRows=UBound(arrStartDateValues)
			'If UBound(arrPrjTitleValues)>MaxRows Then MaxRows=UBound(arrPrjTitleValues)
			
			WriteGridDataMultiRow 2, "1s", "46%#|#55%", objRsExpCou(0) & "#|#" & ConvertSQLDateToText(arrStartDateValues(k), " ", "MMYYYY") & " - " & ConvertSQLDateToText(arrEndDateValues(k), " ", "MMYYYY"), 1, "1#|#1"
			On Error Resume Next
			For k=1 To MaxRows-1
				WriteGridDataMultiRow 2, "1c", "46%#|#55%", " #|#" & ConvertSQLDateToText(arrStartDateValues(k), " ", "MMYYYY") & " - " & ConvertSQLDateToText(arrEndDateValues(k), " ", "MMYYYY"), 1, "1#|#1"
			Next
			On Error GoTo 0
		Else
			WriteGridDataRow 2, "46%#|#55%", objRsExpCou(0) & "#|#" & ConvertSQLDateToText(objRsExpCou(1), " ", "MMYYYY") & " - " & ConvertSQLDateToText(objRsExpCou(2), " ", "MMYYYY"), 1, "1#|#1"
		End If
		objRsExpCou.MoveNext
		Set arrRowsValues=Nothing
	WEnd
	End If
	WriteTableFooter
	objRsExpCou.Close
	Set objRsExpCou=Nothing	

' Close active section and open a new one in landscape
Response.Write("\par \sect}" & vbCrLf)
Response.Write("\sectd \lndscpsxn\pgwsxn16838\pghsxn11906\margl1797\margr1797\margt789\margb689\viewkind1 ")
Response.Write("{" & vbCrLf)
	
' Employment records
	WriteTableHeader
	WriteGridDataRowLandscape 1, "100%", "14. " &  GetLabel(sCvLanguage, "Professional experience") & ":", 0, "0"
	WriteTableFooter

	If Not objRsExpWke.Eof Then
		WriteGridTableHeader
		WriteGridDataRowLandscape 5, "12%#|#10%#|#14%#|#14%#|#50%", GetLabel(sCvLanguage, "Date from") & " -\line " &  GetLabel(sCvLanguage, "Date to") & "#|#" & GetLabel(sCvLanguage, "Location") & "#|#" &  GetLabel(sCvLanguage, "Company and reference person") & "#|#" & GetLabel(sCvLanguage, "Position") & "#|#" & GetLabel(sCvLanguage, "Description"), 1, "1#|#1#|#1#|#1#|#1"

		While Not objRsExpWke.Eof
			sCountries = GetExpertExperienceCountryList(iCvID, objRsExpWke("id_ExpWke"), sCvLanguage)

			dflag=0
			sDescription=ConvertText2RTF(objRsExpWke("wkeDescriptionEng"))
			If Len(objRsExpWke("wkePrjTitleEng"))>0 Then
				sDescription=ConvertText2RTF(objRsExpWke("wkePrjTitleEng")) & "\line " & sDescription
			End If

			Dim sCompanyReference
			sCompanyReference=""
			If Len(objRsExpWke("wkeOrgNameEng"))>1 Then 
				sCompanyReference=sCompanyReference & objRsExpWke("wkeOrgNameEng") & "\line "
			Else
				If Len(objRsExpWke("wkeBnfNameEng"))>1 Then sCompanyReference=sCompanyReference & objRsExpWke("wkeBnfNameEng") & "\line "
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
			
			WriteGridDataRowLandscape 5, "12%#|#10%#|#14%#|#14%#|#50%", "" & _
				ConvertDateForText(objRsExpWke("wkeStartDate"), " ", "MonthYear") & " - " & sExperienceEndDate & _
				"#|#" & sCountries & _
				"#|#" & sCompanyReference & _
				"#|#" & objRsExpWke("wkePositionEng") & _
				"#|#" & sDescription & _
				"", 1, "1#|#1#|#1#|#1#|#1"

			objRsExpWke.MoveNext
		WEnd

		WriteTableFooter
	End If
	objRsExpWke.Close
	Set objRsExpWke=Nothing

	WriteTableHeader
	WriteGridDataRowLandscape 2, "99%#|#1%", "15. " &  GetLabel(sCvLanguage, "Other relevant information (e.g., Publications)") & ":" & "#|# ", 0, "0#|#0"
	WriteGridDataRowLandscape 2, "99%#|#1%", ConvertText(sPublications) & "#|# ", 0, "0#|#0"
	If Len(sReferences)>1 Then
		WriteGridDataRowLandscape 2, "99%#|#1%", "15. " &  GetLabel(sCvLanguage, "Other relevant information (e.g., Publications)") & ":" & "#|# ", 0, "0#|#0"
		WriteGridDataRowLandscape 2, "99%#|#1%", ConvertText(sPublications) & "#|# ", 0, "0#|#0"
	End If
	WriteTableFooter

Response.Write("\par }}" & vbCrLf)
%>
