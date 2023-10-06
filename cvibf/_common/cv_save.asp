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

sFileType=LCase(Request.QueryString("ftype"))
If sFileType="doc" Then
	Response.ContentType = "application/vnd.ms-word"
	Response.AddHeader "Content-Disposition", "attachment; filename=" & sFileName & ".rtf"
End If
If sFileType="prn" Then
	Response.ContentType = "application/vnd.ms-word"
	Response.AddHeader "Content-Disposition", "inline; filename=" & sFileName & ".rtf"
End If
sLastName=Replace(ConvertText2RTF(sLastName), "      ", "")
sFirstName=Replace(ConvertText2RTF(sFirstName), "      ", "")
sFullNameWithSpaces=sFirstName & " " & sLastName
sFullName=Replace(ConvertText2RTF(sFullName), "      ", "")



Response.Write("{\rtf1\ansi{\fonttbl{\f1\fswiss\fcharset0\fprq2 Tahoma;}{\f2\fswiss\fcharset186\fprq2 Tahoma;}}" & vbCrLf)
Response.Write("{\colortbl;\red0\green0\blue0;\red0\green51\blue153;\red0\green255\blue255;\red0\green255\blue0;\red255\green0\blue255;\red204\green0\blue0;\red255\green255\blue0;\red255\green255\blue255;\red0\green0\blue128;\red0\green128\blue128;\red0\green128\blue0;\red128\green0\blue128;\red128\green0\blue0;\red128\green128\blue0;\red128\green128\blue128;\red192\green192\blue192;}" & vbCrLf)
Response.Write("{\stylesheet{\ql \li0\ri0\widctlpar\faauto\rin0\lin0\itap0 \fs24\lang2057\langfe1033\cgrid\langnp2057\langfenp1033 \snext0 Normal;}{\*\cs10 \additive Default Paragraph Font;}{\s15\ql \li0\ri0\widctlpar\tqc\tx4153\tqr\tx8306\aspalpha\aspnum\faauto\adjustright\rin0\lin0\itap0 \fs24\lang2057\langfe1033\cgrid\langnp2057\langfenp1033 \sbasedon0 \snext15 header;}{\s16\ql \li0\ri0\widctlpar\tqc\tx4153\tqr\tx8306\aspalpha\aspnum\faauto\adjustright\rin0\lin0\itap0 \fs24\lang2057\langfe1033\cgrid\langnp2057\langfenp1033 \sbasedon0 \snext16 footer;}{\*\cs17 \additive \sbasedon10 page number;}}" & vbCrLf)
Response.Write("{\info{\title CV of " & sFullName & "}{\author Assortis CVIP}{\company Adetef}{\subject assortis format}{\category Expert CV}{\keywords " & sProfession & "}{\doccomm Downloaded: " & ConvertDateForText(Date()," ", "DDMMYYYY") & "}}\paperw11907\paperh16840\margl1797\margr1797\margt789\margb689\viewkind1\viewscale100\titlepg " & vbCrLf)

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
	WriteDataTitle "PERSONAL INFORMATION"

	WriteTableHeader
	WriteDataRow "Title", sTitle
	WriteDataRow "First name", sFirstName
	If sMiddleName>"" Then
		WriteDataRow "Middle name", sMiddleName
	End If
	WriteDataRow "Family name", sLastName
	WriteDataRow "Date of birth", ConvertDateForText(sBirthDate, " ", "DDMMYYYY")
	If sBirthPlace>"" Then
		WriteDataRow "Place of birth", sBirthPlace
	End If
	WriteDataRow "Nationality", sNationality 
	If iGender>"" And IsNumeric(iGender) Then
		WriteDataRow "Gender", arrGenderTitle(iGender)
	End If
	If iMaritalStatus>"" And IsNumeric(iMaritalStatus) Then
		WriteDataRow "Marital status", arrMaritalStatusTitle(iMaritalStatus)
	End If
	WriteTableFooter


' Education	
	If Not objRsExpEdu.Eof Then
	WriteDataTitle "EDUCATION"
	While Not objRsExpEdu.Eof
		WriteTableHeader
		WriteDataRow "Institution", objRsExpEdu("InstNameEng")
		WriteDataRow "Location", objRsExpEdu("InstLocationEng")
		WriteDataRow "Start date", ConvertDateForText(objRsExpEdu("eduStartDate"), " ", "MMYYYY")
		WriteDataRow "End date", ConvertDateForText(objRsExpEdu("eduEndDate"), " ", "MMYYYY")
		If Not IsNull(objRsExpEdu("eduDiploma")) Then
			WriteDataRow "Type of Diploma", Trim(EducationTypeTitleByID(objRsExpEdu("eduDiploma")) & " " & objRsExpEdu("eduDiploma1Eng"))
		Else
			WriteDataRow "Type of Diploma", objRsExpEdu("eduDiploma1Eng")
		End If
		sEduSubject=""
		If Len(objRsExpEdu("edsDescriptionEng"))>0 And objRsExpEdu("edsDescriptionEng")<>"Other" Then 
			sEduSubject=sEduSubject & Trim(objRsExpEdu("edsDescriptionEng"))
		End If
		If Len(objRsExpEdu("id_EduSubject1Eng"))>0 Then
			sEduSubject=sEduSubject & Trim(" " & objRsExpEdu("id_EduSubject1Eng"))
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
	WriteDataTitle "TRAINING"
	While Not objRsExpTrn.Eof
		WriteTableHeader
		WriteDataRow "Skills / Qualification", objRsExpTrn("eduOtherEng")
		WriteDataRow "Title", objRsExpTrn("eduDiploma1Eng")
		WriteDataRow "Start date", ConvertDateForText(objRsExpTrn("eduStartDate"), " ", "MMYYYY")
		WriteDataRow "End date", ConvertDateForText(objRsExpTrn("eduEndDate"), " ", "MMYYYY")
		sAchievements=ConvertText2RTF(objRsExpTrn("eduDescriptionEng"))
		If sAchievements>"" Then
			WriteDataRow "Achievements", sAchievements
		End If
		WriteTableFooter
	objRsExpTrn.MoveNext
	WEnd
	End If
	objRsExpTrn.Close
	Set objRsExpTrn=Nothing

' Professional experience
	WriteDataTitle "PROFESSIONAL EXPERIENCE"
	WriteTableHeader
	WriteDataRow "Profession", ConvertText2RTF(sProfession)
	WriteDataRow "Current position", ConvertText2RTF(sPosition)
	WriteDataRow "Key qualifications", ConvertText2RTF(sKeyQualification)
	WriteDataRow "Other skills", ConvertText2RTF(sOtherSkills)
	WriteDataRow "Years of professional\line  experience", iProfYears
	WriteTableFooter

' Employment records
	If not objRsExpWke.Eof Then
	WriteDataTitle "EMPLOYMENT RECORD AND COMPLETED PROJECTS"
	While Not objRsExpWke.Eof
		
		WriteTableHeader
		WriteDataRow "Project title", "\b " & objRsExpWke("wkePrjTitleEng") & "\b0 "
objTempRs2=GetDataOutParamsSP("usp_ExpCvvExperienceCouSelect", Array( _
	Array(, adInteger, , iCvID), Array(, adInteger, , objRsExpWke("id_ExpWke")), Array(, adVarChar, 10, "listshort")), Array( _
	Array(, adVarWChar, 4000)))
sCountries=ConvertText2RTF(objTempRs2(0))
Set objTempRs2=Nothing
		WriteDataRow "Country", sCountries
		WriteDataRow "Start date", ConvertDateForText(objRsExpWke("wkeStartDate"), " ", "MMYYYY")
		Dim sExperienceEndDate
		sExperienceEndDate=""
		If objRsExpWke("wkeEndDateOpen")=1 Then
			sExperienceEndDate=sExperienceEndDate & "Ongoing"
			If IsDate(objRsExpWke("wkeEndDate")) And Not IsNull(objRsExpWke("wkeEndDate")) Then sExperienceEndDate=sExperienceEndDate & " ("
		End If
		If IsDate(objRsExpWke("wkeEndDate")) And Not IsNull(objRsExpWke("wkeEndDate")) Then
			sExperienceEndDate=sExperienceEndDate &  ConvertDateForText(objRsExpWke("wkeEndDate"), " ", "MMYYYY")
			If objRsExpWke("wkeEndDateOpen")=1 Then sExperienceEndDate=sExperienceEndDate & ")"
		End If
		WriteDataRow "End date", sExperienceEndDate
		WriteDataRow "Company\~/\~Organisation", objRsExpWke("wkeOrgNameEng")
		WriteDataRow "Position\~/\~Responsibility", ConvertText2RTF(objRsExpWke("wkePositionEng"))
		WriteDataRow "Beneficiary", ConvertText2RTF(objRsExpWke("wkeBnfNameEng"))
		dflag=0
'		WriteDataRow "Funding\~agencies\~", GetDoners("Eng","lnkWke_Don","id_ExpWke",objRsExpWke("id_ExpWke")) & GetOtherDonor("Eng","lnkWke_Don","id_ExpWke",objRsExpWke("id_ExpWke"))

'		sSectors=GetSectors("Eng","lnkWke_Sct","id_ExpWke",objRsExpWke("id_ExpWke"))
'		sSectors=ConvertText2RTF(sSectors)
'		WriteDataRow "Sectors", sSectors

		WriteDataRow "Client\~references", ConvertText2RTF(objRsExpWke("wkeClientRefEng"))
		sDescription=ConvertText2RTF(objRsExpWke("wkeDescriptionEng"))
		WriteDataRow "Detail\~task assigned", sDescription
		WriteTableFooter
	objRsExpWke.MoveNext
	WEnd
	End If
	objRsExpWke.Close
	Set objRsExpWke=Nothing

' Languages
	WriteDataTitle "LANGUAGES SKILLS"
	WriteTableHeader
	WriteDataRow "Languages", "Reading / Speaking / Writing"
	
	If (Not objRsExpLngOther.Eof) Or (Not objRsExpLngNative.Eof) Then
		While Not objRsExpLngNative.Eof
			WriteDataRow objRsExpLngNative("lngNameEng"), "Native"
			objRsExpLngNative.MoveNext
		WEnd

		While Not objRsExpLngOther.Eof
			WriteDataRow objRsExpLngOther("lngNameEng"), arrLanguageLevelTitle(objRsExpLngOther("exlReading")) & " / " & arrLanguageLevelTitle(objRsExpLngOther("exlSpeaking")) & " / " & arrLanguageLevelTitle(objRsExpLngOther("exlWriting"))
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
		WriteDataTitle "OTHER"
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
		WriteDataRow "Availability", sAvailability
	End If
	If sPreferences>"" Then 
		WriteDataRow "Assignment preferences", sPreferences
	End If
	WriteTableFooter
	End If

' Permanent address
	WriteDataTitle "PERMANENT ADDRESS"
	WriteTableHeader
	If sPermAddressStreet>"" Then
		WriteDataRow "Street", sPermAddressStreet
	End If
	If sPermAddressCity>"" Then
		WriteDataRow "City", sPermAddressCity
	End If
	If sPermAddressPostcode>"" Then
		WriteDataRow "Postcode", sPermAddressPostcode
	End If
	WriteDataRow "Country", sPermAddressCountry
	WriteDataRow "Phone", sPermAddressPhone
	If sPermAddressMobile>"" Then
		WriteDataRow "Mobile", sPermAddressMobile
	End If
	If sPermAddressFax>"" Then
		WriteDataRow "Fax", sPermAddressFax
	End If
	WriteDataRow "E-mail", sPermAddressEmail
	If sPermAddressWeb>"" Then
		WriteDataRow "Web-site", sPermAddressWeb
	End If
	WriteTableFooter

' Current address
	If bCurAddress Then
	WriteDataTitle "CURRENT ADDRESS"
	WriteTableHeader
	If sCurAddressStreet>"" Then
		WriteDataRow "Street", sCurAddressStreet
	End If
	If sCurAddressCity>"" Then
		WriteDataRow "City", sCurAddressCity
	End If
	If sCurAddressPostcode>"" Then
		WriteDataRow "Postcode", sCurAddressPostcode
	End If
	WriteDataRow "Country", sCurAddressCountry
	WriteDataRow "Phone", sCurAddressPhone
	If sCurAddressMobile>"" Then
		WriteDataRow "Mobile", sCurAddressMobile
	End If
	If sCurAddressFax>"" Then
		WriteDataRow "Fax", sCurAddressFax
	End If
	WriteDataRow "E-mail", sCurAddressEmail
	If sCurAddressWeb>"" Then
		WriteDataRow "Web-site", sCurAddressWeb
	End If
	WriteTableFooter
	End If

Response.Write("\par }}" & vbCrLf)
%>


