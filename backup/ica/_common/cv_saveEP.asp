<%
'--------------------------------------------------------------------
'
' Expert's CV. Save in EP format
'
'--------------------------------------------------------------------
Response.Buffer = True
%>
<!--#include file="cv_data.asp"-->
<%
' Log: 35 - Download CV
iLogResult = LogActivity(35, "CVID=" & Cstr(iCvID) & " Format: " & sCvFormat, "", "")

Dim MaxRows

sFileType=LCase(Request.QueryString("ftype"))
If sFileType="doc" Or sFileType="prn" Then
	Response.ContentType = "application/vnd.ms-word"
	If sFileType="doc" then Response.AddHeader "Content-Disposition", "attachment; filename=" & sFileName & ".rtf"
	If sFileType="prn" then Response.AddHeader "Content-Disposition", "filename=" & sFileName & ".rtf"
End If
sLastName=Replace(ConvertText2RTF(sLastName), "      ", "")
sFullName=Replace(ConvertText2RTF(sFullName), "      ", "")


Response.Write("{\rtf1\ansi{\fonttbl{\f1\fswiss\fcharset0\fprq2 Arial Narrow;}{\f2\fswiss\fcharset186\fprq2 Arial Narrow;}}" & vbCrLf)
Response.Write("{\colortbl;\red0\green0\blue0;\red0\green0\blue0;\red0\green255\blue255;\red0\green255\blue0;\red255\green0\blue255;\red204\green0\blue0;\red255\green255\blue0;\red255\green255\blue255;\red0\green0\blue128;\red0\green128\blue128;\red0\green128\blue0;\red128\green0\blue128;\red128\green0\blue0;\red128\green128\blue0;\red128\green128\blue128;\red192\green192\blue192;}" & vbCrLf)
Response.Write("{\stylesheet{\ql \li0\ri0\widctlpar\faauto\rin0\lin0\itap0 \fs24\lang2057\langfe1033\cgrid\langnp2057\langfenp1033 \snext0 Normal;}{\*\cs10 \additive Default Paragraph Font;}{\s15\ql \li0\ri0\widctlpar\tqc\tx4153\tqr\tx8306\aspalpha\aspnum\faauto\adjustright\rin0\lin0\itap0 \fs24\lang2057\langfe1033\cgrid\langnp2057\langfenp1033 \sbasedon0 \snext15 header;}{\s16\ql \li0\ri0\widctlpar\tqc\tx4153\tqr\tx8306\aspalpha\aspnum\faauto\adjustright\rin0\lin0\itap0 \fs24\lang2057\langfe1033\cgrid\langnp2057\langfenp1033 \sbasedon0 \snext16 footer;}{\*\cs17 \additive \sbasedon10 page number;}}" & vbCrLf)
Response.Write("{\info{\title CV of " & sFullName & "}{\author assortis.com}{\company assortis.com}{\subject Europass format}{\category Expert CV}{\keywords " & sProfession & "}{\doccomm Downloaded: " & ConvertDateForText(Date(), " ", "DDMMYYYY") & "}}\paperw11907\paperh16840\margl851\margr1797\margt851\margb851\gutter0\viewkind1\viewscale100\titlepg " & vbCrLf)

'Response.Write("{\footerf \pard\plain \s16\qr \li0\ri0\widctlpar\brdrt\brdrs\brdrw10\brsp20 \tqc\tx4153\tqr\tx8306\aspalpha\aspnum\faauto\adjustright\rin0\lin0\itap0 \fs24\lang2057\langfe1033\cgrid\langnp2057\langfenp1033 {\f1\fs16\cf9\lang2060\langfe1033\langnp2060 Page\~: }{\field{\*\fldinst {\cs17\f1\fs16\cf9  PAGE }}{\fldrslt {\cs17\f1\fs16\cf9\lang1024\langfe1024\noproof 2}}}{\f1\fs16\cf9\lang2060\langfe1033\langnp2060 \par }}{\*\pnseclvl1\pnucrm\pnstart1\pnindent720\pnhang{\pntxta .}}{\*\pnseclvl2\pnucltr\pnstart1\pnindent720\pnhang{\pntxta .}}{\*\pnseclvl3\pndec\pnstart1\pnindent720\pnhang{\pntxta .}}{\*\pnseclvl4\pnlcltr\pnstart1\pnindent720\pnhang{\pntxta )}}" & vbCrLf)
'Response.Write("{\footer \pard\plain \s16\qr \li0\ri0\widctlpar\brdrt\brdrs\brdrw10\brsp20 \tqc\tx4153\tqr\tx8306\aspalpha\aspnum\faauto\adjustright\rin0\lin0\itap0 \fs24\lang2057\langfe1033\cgrid\langnp2057\langfenp1033 {\f1\fs16\cf9\lang2060\langfe1033\langnp2060 Page\~: }{\field{\*\fldinst {\cs17\f1\fs16\cf9  PAGE }}{\fldrslt {\cs17\f1\fs16\cf9\lang1024\langfe1024\noproof 2}}}{\f1\fs16\cf9\lang2060\langfe1033\langnp2060 \par }}{\*\pnseclvl1\pnucrm\pnstart1\pnindent720\pnhang{\pntxta .}}{\*\pnseclvl2\pnucltr\pnstart1\pnindent720\pnhang{\pntxta .}}{\*\pnseclvl3\pndec\pnstart1\pnindent720\pnhang{\pntxta .}}{\*\pnseclvl4\pnlcltr\pnstart1\pnindent720\pnhang{\pntxta )}}" & vbCrLf)
Response.Write("{" & vbCrLf)
Response.Write("\par\ql\f1\fs18\cf2" & vbCrLf)

Dim sFlag
Set objFso=Server.CreateObject("Scripting.FileSystemObject")
Set fInTemplate=objFso.OpenTextFile(Server.MapPath("\_common") & "\cv_europass.vrf", 1)
sFlag = fInTemplate.ReadAll
Set fInTemplate=Nothing
Set objFSO=Nothing

' Personal information
	WriteTableHeader
	WriteDataRow "\qr " & sFlag & " \line " & "\fs26\b Europass\line Curriculum Vitae\b0\line ", ""
	WriteDataRow "\qr\fs22\b " & GetLabel(sCvLanguage, "Personal information") & "\b0 ", ""
	WriteSpaceRow

	WriteDataRow "\qr " & GetLabel(sCvLanguage, "Surname(s) / First name(s)"), "\fs22\b " & sTitleLastName & ", " & sFirstName & "\b0 "
	If bUserIbfStaff=0 Then
		WriteDataRow "\qr\fs20 " & GetLabel(sCvLanguage, "Address"), "" & ConvertText2RTF(sPermAddress) & " "
	End If
	WriteDataRow "\qr " & GetLabel(sCvLanguage, "Nationality"), sNationality
	WriteDataRow "\qr " & GetLabel(sCvLanguage, "Date of birth"), ConvertDateForText(sBirthDate, " ", "DayMonthYear")

	WriteSpaceRow
	WriteSpaceRow
	WriteSpaceRow
	
' Employment records
	WriteDataRow "\line\fs22\qr\expnd0\b " & GetLabel(sCvLanguage, "Work experience") & "\b0\fs20 ", ""
	WriteSpaceRow

	While Not objRsExpWke.Eof
		sCountries = GetExpertExperienceCountryList(iCvID, objRsExpWke("id_ExpWke"), sCvLanguage)

		dflag=0
		sDescription=ConvertText2RTF(objRsExpWke("wkeDescriptionEng"))
		
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
		
		WriteDataRow "\qr\line " & GetLabel(sCvLanguage, "Dates"), "\line\b " & ConvertDateForText(objRsExpWke("wkeStartDate"), " ", "MonthYear") & " - " & sExperienceEndDate & "\b0 "
		WriteDataRow "\qr " & GetLabel(sCvLanguage, "Occupation or position held"), objRsExpWke("wkePositionEng")
		WriteDataRow "\qr " & GetLabel(sCvLanguage, "Main activities and responsibilities"), ConvertText2RTF(sDescription)
		WriteDataRow "\qr " & GetLabel(sCvLanguage, "Name and address of employer"), objRsExpWke("wkeOrgNameEng") &  "\line " & ConvertText2RTF(sCountries)
		WriteDataRow "\qr " & GetLabel(sCvLanguage, "Type of business or sector"), ""

		objRsExpWke.MoveNext
	WEnd

	objRsExpWke.Close
	Set objRsExpWke=Nothing

	WriteSpaceRow
	
' Employment records
	WriteDataRow "\line\fs22\expnd0\qr\b " & GetLabel(sCvLanguage, "Education and training") & "\b0\fs20 ", ""
	WriteSpaceRow

' Training
	Dim sPeriod
	While Not objRsExpTrn.Eof
		sPeriod=""
		If Not (IsNull(objRsExpTrn("eduStartDate")) And IsNull(objRsExpTrn("eduEndDate"))) Then
			If Not (IsNull(objRsExpTrn("eduStartDate"))) Then
				sPeriod=sPeriod & ConvertDateForText(objRsExpTrn("eduStartDate"), " ", "MonthYear") & " - "
			End If
			sPeriod=sPeriod & ConvertDateForText(objRsExpTrn("eduEndDate"), " ", "MonthYear")
		End If

		WriteDataRow "\qr\line " & GetLabel(sCvLanguage, "Dates"), "\line\b " & sPeriod & "\b0 "
		WriteDataRow "\qr " & GetLabel(sCvLanguage, "Title of qualification awarded"), objRsExpTrn("eduDiploma1Eng")
		WriteDataRow "\qr " & GetLabel(sCvLanguage, "Principal subjects/occupational skills covered"), Trim(objRsExpTrn("edsDescriptionEng") & " " & objRsExpTrn("id_EduSubject1Eng"))
		WriteDataRow "\qr " & GetLabel(sCvLanguage, "Name and type of organisation providing education and training"), objRsExpTrn("InstNameEng")
		WriteDataRow "\qr " & GetLabel(sCvLanguage, "Level in national or international classification"), ""

		objRsExpTrn.MoveNext
	WEnd
	objRsExpTrn.Close
	Set objRsExpTrn=Nothing
' Education

	While Not objRsExpEdu.Eof
		sPeriod=""
		If Not (IsNull(objRsExpEdu("eduStartDate")) And IsNull(objRsExpEdu("eduEndDate"))) Then
			If Not (IsNull(objRsExpEdu("eduStartDate"))) Then
				sPeriod=sPeriod & ConvertDateForText(objRsExpEdu("eduStartDate"), " ", "MonthYear") & " - "
			End If
			sPeriod=sPeriod & ConvertDateForText(objRsExpEdu("eduEndDate"), " ", "MonthYear")
		End If

		WriteDataRow "\qr\line " & GetLabel(sCvLanguage, "Dates"), "\line\b " & sPeriod & "\b0 "
		WriteDataRow "\qr " & GetLabel(sCvLanguage, "Title of qualification awarded"), objRsExpEdu("eduDiploma1Eng")
		WriteDataRow "\qr " & GetLabel(sCvLanguage, "Principal subjects/occupational skills covered"), Trim(objRsExpEdu("edsDescriptionEng") & " " & objRsExpEdu("id_EduSubject1Eng"))
		WriteDataRow "\qr " & GetLabel(sCvLanguage, "Name and type of organisation providing education and training"), objRsExpEdu("InstNameEng")
		WriteDataRow "\qr " & GetLabel(sCvLanguage, "Level in national or international classification"), ""

		objRsExpEdu.MoveNext
	WEnd
	objRsExpEdu.Close  
	Set objRsExpEdu=Nothing
	
	WriteDataRow "\line\fs22\qr\expnd0\b " & GetLabel(sCvLanguage, "Personal skills and competences") & "\b0\fs20 ", ""
	WriteSpaceRow

' Languages
	Dim sNativeLanguage
	sNativeLanguage = ""
	While Not objRsExpLngNative.Eof
		On Error Resume Next
			sNativeLanguage = sNativeLanguage & ReplaceIfEmpty(objRsExpLngNative("lngName" & sCvLanguage), objRsExpLngNative("lngNameEng")) & ", "
		On Error GoTo 0
		objRsExpLngNative.MoveNext
	WEnd
	If Len(sNativeLanguage)>2 Then sNativeLanguage=Left(sNativeLanguage, Len(sNativeLanguage)-2)

	WriteDataRow "\qr " & GetLabel(sCvLanguage, "Mother tongue(s)"), "\b " & sNativeLanguage & "\b0 "
	
	If (Not objRsExpLngOther.Eof) Then
		WriteDataRow "\qr\line " & GetLabel(sCvLanguage, "Other language(s)"), ""
		WriteGridTableHeader
		WriteGridDataRow 4, "30%#|#28%#|#28%#|#15%", "" & "#|#\qc\b " &  GetLabel(sCvLanguage, "Understanding") & "\b0#|#\qc\b " &  GetLabel(sCvLanguage, "Speaking(EP)") & "\b0#|#\qc\b " &  GetLabel(sCvLanguage, "Writing(EP)") & "\b0", 1, "0#|#1#|#1#|#1"
		WriteGridDataRow 6, "30%#|#14%#|#14%#|#14%#|#14%#|#15%", "" & "#|#\qc " &  GetLabel(sCvLanguage, "Listening") & "#|#\qc " &  GetLabel(sCvLanguage, "Reading(EP)") & "#|#\qc " &  GetLabel(sCvLanguage, "Spoken interaction") & "#|#\qc " &  GetLabel(sCvLanguage, "Spoken production") & "#|#\qc " &  "", 1, "0#|#1#|#1#|#1#|#1#|#1"

		While Not objRsExpLngOther.Eof
			On Error Resume Next
				sTempLanguage = ReplaceIfEmpty(objRsExpLngOther("lngName" & sCvLanguage), objRsExpLngOther("lngNameEng"))
			On Error GoTo 0
			WriteGridDataRow 6, "30%#|#14%#|#14%#|#14%#|#14%#|#15%", "\qr\b " &  sTempLanguage & " \b0 #|#\qc " & SetEPLanguageLevel(objRsExpLngOther("exlReading")) & "#|#\qc " & SetEPLanguageLevel(objRsExpLngOther("exlReading")) & "#|#\qc " & SetEPLanguageLevel(objRsExpLngOther("exlSpeaking")) & "#|#\qc " & SetEPLanguageLevel(objRsExpLngOther("exlSpeaking")) & "#|#\qc " & SetEPLanguageLevel(objRsExpLngOther("exlWriting")), 1, "0#|#1#|#1#|#1#|#1#|#1"
			objRsExpLngOther.MoveNext
		WEnd
		WriteTableFooter
		WriteDataRow "\qr\line ", "* " & GetLabel(sCvLanguage, "Common European Framework of Reference for Languages")
	End If

	objRsExpLngNative.Close
	Set objRsExpLngNative=Nothing	
	objRsExpLngOther.Close
	Set objRsExpLngOther=Nothing	
	
	WriteDataRow "\qr " & GetLabel(sCvLanguage, "Social skills and competences"), ""
	WriteSpaceRow

	WriteDataRow "\line\qr " & GetLabel(sCvLanguage, "Organisational skills and competences"), ""
	WriteSpaceRow

	WriteDataRow "\line\qr " & GetLabel(sCvLanguage, "Technical skills and competences"), ConvertText2RTF(sKeyQualification)
	WriteSpaceRow

	WriteDataRow "\line\qr " & GetLabel(sCvLanguage, "Computer skills and competences"), ""
	WriteSpaceRow

	WriteDataRow "\line\qr " & GetLabel(sCvLanguage, "Artistic skills and competences"), ""
	WriteSpaceRow

	WriteDataRow "\line\qr " & GetLabel(sCvLanguage, "Other skills and competences"), ""
	WriteSpaceRow

	WriteDataRow "\line\qr " & GetLabel(sCvLanguage, "Driving licence(s)"), ""
	WriteSpaceRow
	WriteSpaceRow

	WriteDataRow "\line\qr\fs22\expnd0\b " & GetLabel(sCvLanguage, "Additional information") & "\b0\fs20 ", ""
	WriteDataRow "\line\qr " & GetLabel(sCvLanguage, "Publications"), ConvertText2RTF(sPublications)
	WriteSpaceRow
	WriteDataRow "\line\qr " & GetLabel(sCvLanguage, "Memberships"), ConvertText2RTF(sMemberships)
	WriteSpaceRow
	WriteDataRow "\line\qr " & GetLabel(sCvLanguage, "References"), ConvertText2RTF(sReferences)
	WriteSpaceRow
	WriteSpaceRow

	WriteDataRow "\line\qr\fs22\expnd0\b " & GetLabel(sCvLanguage, "Annexes") & "\b0\fs20 ", ""
	WriteSpaceRow
	WriteSpaceRow

	WriteTableFooter

Response.Write("\par }}" & vbCrLf)
%>
