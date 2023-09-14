<%
'--------------------------------------------------------------------
'
' Expert's CV. Save in assortis.com format
'
'--------------------------------------------------------------------
Response.Buffer = True
%>
<!--#include file="cv_data.asp"-->
<%
' Log: 35 - Download CV
iLogResult = LogActivity(35, "CVID=" & Cstr(iCvID) & " Format: ASR", "", "")

sFileType = LCase(Request.QueryString("ftype"))
If sFileType="doc" Then
	Response.ContentType = "application/msword"
	Response.AddHeader "Content-Disposition", "attachment; filename=" & sFileName & ".rtf"
End If
If sFileType="prn" Then
	Response.ContentType = "application/msword"
	Response.AddHeader "Content-Disposition", "inline; filename=" & sFileName & ".rtf"
End If
sLastName=Replace(ConvertText2RTF(sLastName), "      ", "")
sFirstName=Replace(ConvertText2RTF(sFirstName), "      ", "")
sFullNameWithSpaces=sFirstName & " " & sLastName
sFullName=Replace(ConvertText2RTF(sFullName), "      ", "")



Response.Write("{\rtf1\ansi{\fonttbl{\f1\fswiss\fcharset0\fprq2 Tahoma;}{\f2\fswiss\fcharset186\fprq2 Tahoma;}}" & vbCrLf)
Response.Write("{\colortbl;\red0\green0\blue0;\red0\green51\blue153;\red0\green255\blue255;\red0\green255\blue0;\red255\green0\blue255;\red204\green0\blue0;\red255\green255\blue0;\red255\green255\blue255;\red0\green0\blue128;\red0\green128\blue128;\red0\green128\blue0;\red128\green0\blue128;\red128\green0\blue0;\red128\green128\blue0;\red128\green128\blue128;\red192\green192\blue192;}" & vbCrLf)
Response.Write("{\stylesheet{\ql \li0\ri0\widctlpar\faauto\rin0\lin0\itap0 \fs24\lang2057\langfe1033\cgrid\langnp2057\langfenp1033 \snext0 Normal;}{\*\cs10 \additive Default Paragraph Font;}{\s15\ql \li0\ri0\widctlpar\tqc\tx4153\tqr\tx8306\aspalpha\aspnum\faauto\adjustright\rin0\lin0\itap0 \fs24\lang2057\langfe1033\cgrid\langnp2057\langfenp1033 \sbasedon0 \snext15 header;}{\s16\ql \li0\ri0\widctlpar\tqc\tx4153\tqr\tx8306\aspalpha\aspnum\faauto\adjustright\rin0\lin0\itap0 \fs24\lang2057\langfe1033\cgrid\langnp2057\langfenp1033 \sbasedon0 \snext16 footer;}{\*\cs17 \additive \sbasedon10 page number;}}" & vbCrLf)
Response.Write("{\info{\title CV of " & sFullName & "}{\author ICA}{\company ICA}{\subject assortis format}{\category Expert CV}{\keywords " & sProfession & "}{\doccomm Downloaded: " & ConvertDateForText(Date()," ", "DDMMYYYY") & "}}\paperw11907\paperh16840\margl1797\margr1797\margt789\margb689\viewkind1\viewscale100\titlepg " & vbCrLf)

Response.Write("{\footerf \pard\plain \s16\qr \li0\ri0\widctlpar\brdrt\brdrs\brdrw10\brsp20 \tqc\tx4153\tqr\tx8306\aspalpha\aspnum\faauto\adjustright\rin0\lin0\itap0 \fs24\lang2057\langfe1033\cgrid\langnp2057\langfenp1033 {\f1\fs16\cf9\lang2060\langfe1033\langnp2060 Page\~: }{\field{\*\fldinst {\cs17\f1\fs16\cf9  PAGE }}{\fldrslt {\cs17\f1\fs16\cf9\lang1024\langfe1024\noproof 2}}}{\f1\fs16\cf9\lang2060\langfe1033\langnp2060 \par }}{\*\pnseclvl1\pnucrm\pnstart1\pnindent720\pnhang{\pntxta .}}{\*\pnseclvl2\pnucltr\pnstart1\pnindent720\pnhang{\pntxta .}}{\*\pnseclvl3\pndec\pnstart1\pnindent720\pnhang{\pntxta .}}{\*\pnseclvl4\pnlcltr\pnstart1\pnindent720\pnhang{\pntxta )}}" & vbCrLf)
Response.Write("{\footer \pard\plain \s16\qr \li0\ri0\widctlpar\brdrt\brdrs\brdrw10\brsp20 \tqc\tx4153\tqr\tx8306\aspalpha\aspnum\faauto\adjustright\rin0\lin0\itap0 \fs24\lang2057\langfe1033\cgrid\langnp2057\langfenp1033 {\f1\fs16\cf9\lang2060\langfe1033\langnp2060 Page\~: }{\field{\*\fldinst {\cs17\f1\fs16\cf9  PAGE }}{\fldrslt {\cs17\f1\fs16\cf9\lang1024\langfe1024\noproof 2}}}{\f1\fs16\cf9\lang2060\langfe1033\langnp2060 \par }}{\*\pnseclvl1\pnucrm\pnstart1\pnindent720\pnhang{\pntxta .}}{\*\pnseclvl2\pnucltr\pnstart1\pnindent720\pnhang{\pntxta .}}{\*\pnseclvl3\pndec\pnstart1\pnindent720\pnhang{\pntxta .}}{\*\pnseclvl4\pnlcltr\pnstart1\pnindent720\pnhang{\pntxta )}}" & vbCrLf)
Response.Write("{" & vbCrLf)
Response.Write("\par\ql\f1\fs18\cf2" & vbCrLf)

'Set objFso=Server.CreateObject("Scripting.FileSystemObject")
'Set fInTemplate=objFso.OpenTextFile(Server.MapPath("\_common") & "\cv_image.vrf", 1)
'Response.Write fInTemplate.ReadAll
'Set fInTemplate=Nothing
'Set objFSO=Nothing

'Response.Write("{\cs17\f1\fs16\cf8  CV of " & sFullNameWithSpaces & " (  Downloaded: " & ConvertDateForText(Date(), " ", "DayMonthYear") & " )}{\f1\fs16\cf8 \par }}" & vbCrLf)
'Response.Write("{\footerf \pard\plain \s16\qr \li0\ri0\widctlpar\brdrt\brdrs\brdrw10\brsp20 \tqc\tx4153\tqr\tx8306\aspalpha\aspnum\faauto\adjustright\rin0\lin0\itap0 \fs24\lang2057\langfe1033\cgrid\langnp2057\langfenp1033 {\f1\fs16\cf9\lang2060\langfe1033\langnp2060 Page\~: }{\field{\*\fldinst {\cs17\f1\fs16\cf9  PAGE }}{\fldrslt {\cs17\f1\fs16\cf9\lang1024\langfe1024\noproof 2}}}{\f1\fs16\cf9\lang2060\langfe1033\langnp2060 \par }}{\*\pnseclvl1\pnucrm\pnstart1\pnindent720\pnhang{\pntxta .}}{\*\pnseclvl2\pnucltr\pnstart1\pnindent720\pnhang{\pntxta .}}{\*\pnseclvl3\pndec\pnstart1\pnindent720\pnhang{\pntxta .}}{\*\pnseclvl4\pnlcltr\pnstart1\pnindent720\pnhang{\pntxta )}}" & vbCrLf)
'Response.Write("{\footer \pard\plain \s16\qr \li0\ri0\widctlpar\brdrt\brdrs\brdrw10\brsp20 \tqc\tx4153\tqr\tx8306\aspalpha\aspnum\faauto\adjustright\rin0\lin0\itap0 \fs24\lang2057\langfe1033\cgrid\langnp2057\langfenp1033 {\f1\fs16\cf9\lang2060\langfe1033\langnp2060 Page\~: }{\field{\*\fldinst {\cs17\f1\fs16\cf9  PAGE }}{\fldrslt {\cs17\f1\fs16\cf9\lang1024\langfe1024\noproof 2}}}{\f1\fs16\cf9\lang2060\langfe1033\langnp2060 \par }}{\*\pnseclvl1\pnucrm\pnstart1\pnindent720\pnhang{\pntxta .}}{\*\pnseclvl2\pnucltr\pnstart1\pnindent720\pnhang{\pntxta .}}{\*\pnseclvl3\pndec\pnstart1\pnindent720\pnhang{\pntxta .}}{\*\pnseclvl4\pnlcltr\pnstart1\pnindent720\pnhang{\pntxta )}}" & vbCrLf)
'Response.Write("{" & vbCrLf)
'Response.Write("\qc\f1\fs18\cf6 \b Curriculum Vitae.\b0\cf0" & vbCrLf)

' Personal information
	WriteDataTitle GetLabel(sCvLanguage, "Personal information")

	WriteTableHeader
	WriteDataRow GetLabel(sCvLanguage, "Personal title"), sTitle
	WriteDataRow GetLabel(sCvLanguage, "First name"), sFirstName
	If sMiddleName>"" Then
		WriteDataRow GetLabel(sCvLanguage, "Middle name"), sMiddleName
	End If
	WriteDataRow GetLabel(sCvLanguage, "Family name"), sLastName
	WriteDataRow GetLabel(sCvLanguage, "Date of birth"), ConvertDateForText(sBirthDate, " ", "DDMMYYYY")
	If sBirthPlace>"" Then
		WriteDataRow ConvertText2RTF(GetLabel(sCvLanguage, "Place of birth")), sBirthPlace
	End If
	WriteDataRow GetLabel(sCvLanguage, "Nationality"), sNationality 
	If iGender>"" And IsNumeric(iGender) Then
		WriteDataRow GetLabel(sCvLanguage, "Gender"), arrGenderTitle(iGender)
	End If
	If iMaritalStatus>"" And IsNumeric(iMaritalStatus) Then
		WriteDataRow GetLabel(sCvLanguage, "Marital status"), arrMaritalStatusTitle(iMaritalStatus)
	End If
	WriteTableFooter


' Education	
	If Not objRsExpEdu.Eof Then
	WriteDataTitle GetLabel(sCvLanguage, "Education")
	While Not objRsExpEdu.Eof
		WriteTableHeader
		If objRsExpEdu("InstNameEng")>"" Then WriteDataRow GetLabel(sCvLanguage, "Institution"), objRsExpEdu("InstNameEng")
		If objRsExpEdu("InstLocationEng")>"" Then WriteDataRow GetLabel(sCvLanguage, "Location"), objRsExpEdu("InstLocationEng")
		If IsDate(objRsExpEdu("eduStartDate")) Then WriteDataRow ConvertText2RTF(GetLabel(sCvLanguage, "Start date")), ConvertDateForText(objRsExpEdu("eduStartDate"), " ", "MMYYYY")
		If IsDate(objRsExpEdu("eduEndDate")) Then WriteDataRow ConvertText2RTF(GetLabel(sCvLanguage, "End date")), ConvertDateForText(objRsExpEdu("eduEndDate"), " ", "MMYYYY")
		If Not IsNull(objRsExpEdu("eduDiploma")) Then
			WriteDataRow GetLabel(sCvLanguage, "Type of diploma(RTF)"), Trim(EducationTypeTitleByID(objRsExpEdu("eduDiploma")) & " " & objRsExpEdu("eduDiploma1Eng"))
		Else
			WriteDataRow GetLabel(sCvLanguage, "Type of diploma(RTF)"), objRsExpEdu("eduDiploma1Eng")
		End If
		sEduSubject=""
		If Len(objRsExpEdu("edsDescriptionEng"))>0 And objRsExpEdu("edsDescriptionEng")<>"Other" Then 
			sEduSubject=Trim(sEduSubject & objRsExpEdu("edsDescriptionEng"))
		End If
		If Len(objRsExpEdu("id_EduSubject1Eng"))>0 Then
			sEduSubject=Trim(sEduSubject & " " & objRsExpEdu("id_EduSubject1Eng"))
		End If
		If Len(sEduSubject)>0 Then
			WriteDataRow "Subject", sEduSubject
		End If		
		WriteTableFooter
	objRsExpEdu.MoveNext
	WEnd
	End If
	objRsExpEdu.Close
	Set objRsExpEdu=Nothing

' Training
	If not objRsExpTrn.Eof Then
	WriteDataTitle GetLabel(sCvLanguage, "Training")
	While Not objRsExpTrn.Eof
		WriteTableHeader
		WriteDataRow GetLabel(sCvLanguage, "Skills / Qualification"), objRsExpTrn("eduOtherEng")
		WriteDataRow GetLabel(sCvLanguage, "Title"), objRsExpTrn("eduDiploma1Eng")
		WriteDataRow ConvertText2RTF(GetLabel(sCvLanguage, "Start date")), ConvertDateForText(objRsExpTrn("eduStartDate"), " ", "MMYYYY")
		WriteDataRow ConvertText2RTF(GetLabel(sCvLanguage, "End date")), ConvertDateForText(objRsExpTrn("eduEndDate"), " ", "MMYYYY")
		sAchievements=ConvertText2RTF(objRsExpTrn("eduDescriptionEng"))
		If sAchievements>"" Then
			WriteDataRow GetLabel(sCvLanguage, "Achievements"), sAchievements
		End If
		WriteTableFooter
	objRsExpTrn.MoveNext
	WEnd
	End If
	objRsExpTrn.Close
	Set objRsExpTrn=Nothing

' Professional experience
	WriteDataTitle GetLabel(sCvLanguage, "Professional experience")
	WriteTableHeader
	WriteDataRow GetLabel(sCvLanguage, "Profession"), ConvertText2RTF(sProfession)
	WriteDataRow GetLabel(sCvLanguage, "Current position"), ConvertText2RTF(sPosition)
	WriteDataRow GetLabel(sCvLanguage, "Key qualifications"), ConvertText2RTF(sKeyQualification)
	WriteDataRow GetLabel(sCvLanguage, "Other skills"), ConvertText2RTF(sOtherSkills)
	WriteDataRow GetLabel(sCvLanguage, "Years of professional experience"), iProfYears
	WriteTableFooter

' Employment records
	If not objRsExpWke.Eof Then
	WriteDataTitle GetLabel(sCvLanguage, "Employment record and completed projects")
	While Not objRsExpWke.Eof
		
		WriteTableHeader
		WriteDataRow GetLabel(sCvLanguage, "Project title"), "\b " & objRsExpWke("wkePrjTitleEng") & "\b0 "
		WriteDataRow ConvertText2RTF(GetLabel(sCvLanguage, "Start date")), ConvertDateForText(objRsExpWke("wkeStartDate"), " ", "MMYYYY")
		Dim sExperienceEndDate
		sExperienceEndDate=""
		If objRsExpWke("wkeEndDateOpen")=1 Then
			sExperienceEndDate=sExperienceEndDate & GetLabel(sCvLanguage, "Ongoing")
			If IsDate(objRsExpWke("wkeEndDate")) And Not IsNull(objRsExpWke("wkeEndDate")) Then sExperienceEndDate=sExperienceEndDate & " ("
		End If
		If IsDate(objRsExpWke("wkeEndDate")) And Not IsNull(objRsExpWke("wkeEndDate")) Then
			sExperienceEndDate=sExperienceEndDate &  ConvertDateForText(objRsExpWke("wkeEndDate"), " ", "MMYYYY")
			If objRsExpWke("wkeEndDateOpen")=1 Then sExperienceEndDate=sExperienceEndDate & ")"
		End If
		WriteDataRow ConvertText2RTF(GetLabel(sCvLanguage, "End date")), sExperienceEndDate
		WriteDataRow GetLabel(sCvLanguage, "Company / Organisation"), objRsExpWke("wkeOrgNameEng")
		WriteDataRow GetLabel(sCvLanguage, "Position / Responsibility"), ConvertText2RTF(objRsExpWke("wkePositionEng"))
		WriteDataRow GetLabel(sCvLanguage, "Beneficiary"), ConvertText2RTF(objRsExpWke("wkeBnfNameEng"))

		sCountries = GetExpertExperienceCountryGroupedList(iCvID, objRsExpWke("id_ExpWke"), sCvLanguage)
		sSectors = GetExpertExperienceSectorGroupedList(iCvID, objRsExpWke("id_ExpWke"), sCvLanguage)

objTempRs2=GetDataOutParamsSPWithConn(objConnCustom, "usp_ExpCvvExperienceDonSelect", Array( _
	Array(, adInteger, , iCvID), Array(, adInteger, , objRsExpWke("id_ExpWke")), Array(, adVarChar, 10, "list")), Array( _
	Array(, adVarWChar, 4000)))
sDonors=objTempRs2(0)
Set objTempRs2=Nothing

		WriteDataRow GetLabel(sCvLanguage, "Funding agencies"), ConvertText2RTF(sDonors)
		WriteDataRow GetLabel(sCvLanguage, "Countries"), ConvertText2RTF(sCountries)
		WriteDataRow GetLabel(sCvLanguage, "Sectors"), ConvertText2RTF(sSectors)

		If objRsExpWke("wkeClientRefEng")>"" Then WriteDataRow GetLabel(sCvLanguage, "Client references"), ConvertText2RTF(objRsExpWke("wkeClientRefEng"))
		sDescription=ConvertText2RTF(objRsExpWke("wkeDescriptionEng"))
		WriteDataRow GetLabel(sCvLanguage, "Description of tasks(RTF)"), sDescription
		WriteTableFooter
	objRsExpWke.MoveNext
	WEnd
	End If
	objRsExpWke.Close
	Set objRsExpWke=Nothing

' Languages
	WriteDataTitle GetLabel(sCvLanguage, "Languages skills")
	WriteTableHeader
	WriteDataRow GetLabel(sCvLanguage, "Language"), GetLabel(sCvLanguage, "Reading") & " / " & GetLabel(sCvLanguage, "Speaking") & " / " & GetLabel(sCvLanguage, "Writing")
	
	If (Not objRsExpLngOther.Eof) Or (Not objRsExpLngNative.Eof) Then
		While Not objRsExpLngNative.Eof
			On Error Resume Next
				sTempLanguage = ReplaceIfEmpty(objRsExpLngNative("lngName" & sCvLanguage), objRsExpLngNative("lngNameEng"))
			On Error GoTo 0
			WriteDataRow sTempLanguage, GetLabel(sCvLanguage, "Native")
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
			WriteDataRow sTempLanguage, sReading & " / " & sSpeaking & " / " & sWriting
			objRsExpLngOther.MoveNext
		WEnd
	End If
	objRsExpLngNative.Close
	Set objRsExpLngNative=Nothing	
	objRsExpLngOther.Close
	Set objRsExpLngOther=Nothing	

	WriteTableFooter

' Other
	If sMemberships>"" Or sPublications>"" Or sReferences>"" Or sAvailability>"" Or sPreferences>"" Then
		WriteDataTitle GetLabel(sCvLanguage, "Other")
	End If
	WriteTableHeader
	If sMemberships>"" Then
		WriteDataRow "Membership of\line  professional bodies", ConvertText2RTF(sMemberships)
	End If
	If sPublications>"" Then
		WriteDataRow "Publications", ConvertText2RTF(sPublications)
	End If
	If sReferences>"" Then
		WriteDataRow "References", ConvertText2RTF(sReferences)
	End If
	WriteTableFooter

	If sAvailability>"" Or sPreferences>"" Then 
	WriteTableHeader
	If sAvailability>"" Then 
		WriteDataRow GetLabel(sCvLanguage, "Availability"), sAvailability
	End If
	If sPreferences>"" Then 
		WriteDataRow GetLabel(sCvLanguage, "Assignment preferences"), sPreferences
	End If
	WriteTableFooter
	End If

	' Permanent address
		WriteDataTitle GetLabel(sCvLanguage, "Permanent address")
		WriteTableHeader
		If sPermAddressStreet>"" Then
			WriteDataRow GetLabel(sCvLanguage, "Street"), sPermAddressStreet
		End If
		If sPermAddressCity>"" Then
			WriteDataRow GetLabel(sCvLanguage, "City"), sPermAddressCity
		End If
		If sPermAddressPostcode>"" Then
			WriteDataRow GetLabel(sCvLanguage, "Postcode"), sPermAddressPostcode
		End If
		WriteDataRow GetLabel(sCvLanguage, "Country"), sPermAddressCountry
		WriteDataRow GetLabel(sCvLanguage, "Phone"), sPermAddressPhone
		If sPermAddressMobile>"" Then
			WriteDataRow GetLabel(sCvLanguage, "Mobile"), sPermAddressMobile
		End If
		If sPermAddressFax>"" Then
			WriteDataRow GetLabel(sCvLanguage, "Fax"), sPermAddressFax
		End If
		WriteDataRow GetLabel(sCvLanguage, "Email"), sPermAddressEmail
		If sPermAddressWeb>"" Then
			WriteDataRow GetLabel(sCvLanguage, "Website"), sPermAddressWeb
		End If
		WriteTableFooter

	' Current address
		If bCurAddress Then
		WriteDataTitle GetLabel(sCvLanguage, "Current address")
		WriteTableHeader
		If sCurAddressStreet>"" Then
			WriteDataRow GetLabel(sCvLanguage, "Street"), sCurAddressStreet
		End If
		If sCurAddressCity>"" Then
			WriteDataRow GetLabel(sCvLanguage, "City"), sCurAddressCity
		End If
		If sCurAddressPostcode>"" Then
			WriteDataRow GetLabel(sCvLanguage, "Postcode"), sCurAddressPostcode
		End If
		WriteDataRow GetLabel(sCvLanguage, "Country"), sCurAddressCountry
		WriteDataRow GetLabel(sCvLanguage, "Phone"), sCurAddressPhone
		If sCurAddressMobile>"" Then
			WriteDataRow GetLabel(sCvLanguage, "Mobile"), sCurAddressMobile
		End If
		If sCurAddressFax>"" Then
			WriteDataRow GetLabel(sCvLanguage, "Fax"), sCurAddressFax
		End If
		WriteDataRow GetLabel(sCvLanguage, "Email"), sCurAddressEmail
		If sCurAddressWeb>"" Then
			WriteDataRow GetLabel(sCvLanguage, "Website"), sCurAddressWeb
		End If
		WriteTableFooter
		End If
	
Response.Write("\par }}" & vbCrLf)
%>


