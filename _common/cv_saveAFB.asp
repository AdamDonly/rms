<%
'--------------------------------------------------------------------
'
' Expert's CV. Save in AFB format
'
'--------------------------------------------------------------------
Response.Buffer = True
%>
<!--#include file="cv_data.asp"-->
<%
sFileType=LCase(Request.QueryString("ftype"))
If sFileType="doc" Or sFileType="prn" Then
	Response.ContentType = "application/vnd.ms-word"
	Response.AddHeader "Content-Disposition", "attachment; filename=" & sFileName & ".rtf"
End If
sLastName=Replace(ConvertText2RTF(sLastName), "      ", "")
sFullName=Replace(ConvertText2RTF(sFullName), "      ", "")


Response.Write("{\rtf1\ansi{\fonttbl{\f1\fswiss\fcharset0\fprq2 Arial;}{\f2\fswiss\fcharset186\fprq2 Arial;}}" & vbCrLf)
Response.Write("{\colortbl;\red0\green0\blue0;\red0\green0\blue0;\red0\green255\blue255;\red0\green255\blue0;\red255\green0\blue255;\red204\green0\blue0;\red255\green255\blue0;\red255\green255\blue255;\red0\green0\blue128;\red0\green128\blue128;\red0\green128\blue0;\red128\green0\blue128;\red128\green0\blue0;\red128\green128\blue0;\red128\green128\blue128;\red192\green192\blue192;}" & vbCrLf)
Response.Write("{\stylesheet{\ql \li0\ri0\widctlpar\faauto\rin0\lin0\itap0 \fs24\lang2057\langfe1033\cgrid\langnp2057\langfenp1033 \snext0 Normal;}{\*\cs10 \additive Default Paragraph Font;}{\s15\ql \li0\ri0\widctlpar\tqc\tx4153\tqr\tx8306\aspalpha\aspnum\faauto\adjustright\rin0\lin0\itap0 \fs24\lang2057\langfe1033\cgrid\langnp2057\langfenp1033 \sbasedon0 \snext15 header;}{\s16\ql \li0\ri0\widctlpar\tqc\tx4153\tqr\tx8306\aspalpha\aspnum\faauto\adjustright\rin0\lin0\itap0 \fs24\lang2057\langfe1033\cgrid\langnp2057\langfenp1033 \sbasedon0 \snext16 footer;}{\*\cs17 \additive \sbasedon10 page number;}}" & vbCrLf)
Response.Write("{\info{\title CV of " & sFullName & "}{\author assortis.com}{\company assortis.com}{\subject African Development Bank format}{\category Expert CV}{\keywords " & sProfession & "}{\doccomm Downloaded: " & ConvertDateForText(Date(), " ", "DDMMYYY") & "}}\paperw11907\paperh16840\margl1797\margr1797\margt789\margb689\viewkind1\viewscale100\titlepg " & vbCrLf)

Response.Write("{\footerf \pard\plain \s16\qr \li0\ri0\widctlpar\brdrt\brdrs\brdrw10\brsp20 \tqc\tx4153\tqr\tx8306\aspalpha\aspnum\faauto\adjustright\rin0\lin0\itap0 \fs24\lang2057\langfe1033\cgrid\langnp2057\langfenp1033 {\f1\fs16\cf9\lang2060\langfe1033\langnp2060 Page\~: }{\field{\*\fldinst {\cs17\f1\fs16\cf9  PAGE }}{\fldrslt {\cs17\f1\fs16\cf9\lang1024\langfe1024\noproof 2}}}{\f1\fs16\cf9\lang2060\langfe1033\langnp2060 \par }}{\*\pnseclvl1\pnucrm\pnstart1\pnindent720\pnhang{\pntxta .}}{\*\pnseclvl2\pnucltr\pnstart1\pnindent720\pnhang{\pntxta .}}{\*\pnseclvl3\pndec\pnstart1\pnindent720\pnhang{\pntxta .}}{\*\pnseclvl4\pnlcltr\pnstart1\pnindent720\pnhang{\pntxta )}}" & vbCrLf)
Response.Write("{\footer \pard\plain \s16\qr \li0\ri0\widctlpar\brdrt\brdrs\brdrw10\brsp20 \tqc\tx4153\tqr\tx8306\aspalpha\aspnum\faauto\adjustright\rin0\lin0\itap0 \fs24\lang2057\langfe1033\cgrid\langnp2057\langfenp1033 {\f1\fs16\cf9\lang2060\langfe1033\langnp2060 Page\~: }{\field{\*\fldinst {\cs17\f1\fs16\cf9  PAGE }}{\fldrslt {\cs17\f1\fs16\cf9\lang1024\langfe1024\noproof 2}}}{\f1\fs16\cf9\lang2060\langfe1033\langnp2060 \par }}{\*\pnseclvl1\pnucrm\pnstart1\pnindent720\pnhang{\pntxta .}}{\*\pnseclvl2\pnucltr\pnstart1\pnindent720\pnhang{\pntxta .}}{\*\pnseclvl3\pndec\pnstart1\pnindent720\pnhang{\pntxta .}}{\*\pnseclvl4\pnlcltr\pnstart1\pnindent720\pnhang{\pntxta )}}" & vbCrLf)
Response.Write("\qc\f1\fs18\cf2 \b Format of Curriculum Vitae (CV) For Proposed Key Staff\b0\par" & vbCrLf)
Response.Write("{" & vbCrLf)
Response.Write("\par\ql\f1\fs18\cf2" & vbCrLf)

' Personal information
	WriteTableHeader
	WriteDataRow "\b " & GetLabel(sCvLanguage, "Proposed Position") & ":\b0 ", ""
	WriteDataRow "\b " & GetLabel(sCvLanguage, "Name of Firm") & ":\b0 ", ""
	WriteDataRow "\b " & GetLabel(sCvLanguage, "Name of Staff") & ":\b0 ", sFullName
	WriteDataRow "\b " & GetLabel(sCvLanguage, "Profession") & ":\b0 ", sProfession
	WriteDataRow "\b " & GetLabel(sCvLanguage, "Date of Birth") & ":\b0 ", ConvertDateForText(sBirthDate, " ", "DayMonthYear")
	WriteSpaceRow
	WriteDataRow "\b " & GetLabel(sCvLanguage, "Years with Firm") & ":\b0 ", ""
	WriteDataRow "\b " & GetLabel(sCvLanguage, "Nationality") & ":\b0 ", sNationality
	WriteDataRow "\b " & GetLabel(sCvLanguage, "Membership in Professional Societies") & ":\b0 ", ConvertText2RTF(sMemberships)
	WriteSpaceRow
	WriteDataRow "\b " & GetLabel(sCvLanguage, "Detailed Tasks Assigned (AFB)") & ":\b0 ", ""
	WriteSpaceRow
	WriteDataRow "\b " & GetLabel(sCvLanguage, "Key Qualifications (AFD)") & ":\b0", ConvertText2RTF(sKeyQualification)
	WriteTableFooter

' Education
	WriteTableHeader
	WriteSimpleRow "\b " & GetLabel(sCvLanguage, "Education") & "\b0 :"
	WriteTableFooterNoPar
	WriteGridTableHeader
	WriteGridDataRow 4, "31%#|#15%#|#15%#|#40%", "\b " & GetLabel(sCvLanguage, "Institution") & "\b0 #|#\b " & GetLabel(sCvLanguage, "Start date") & "\b0 #|#\b " & GetLabel(sCvLanguage, "End date") & "\b0 #|#\b " & GetLabel(sCvLanguage, "Degree(s) or Diploma(s) obtained") & "\b0 ", 1, "1#|#1#|#1#|#1"

	If objRsExpEdu.Eof Then
		WriteGridDataRow 4, "31%#|#15%#|#15%#|#40%", "-#|#-#|#-#|#-", 1, "1#|#1#|#1#|#1"
	Else
        i=1
	While Not objRsExpEdu.Eof
		WriteGridDataRow 4, "31%#|#15%#|#15%#|#40%", objRsExpEdu("InstNameEng") & "#|#" & ConvertDateForText(objRsExpEdu("eduStartDate"), " ", "MMYYYY") & "#|#" & ConvertDateForText(objRsExpEdu("eduEndDate"), " ", "MMYYYY") & "#|#" & Trim(objRsExpEdu("edtDescriptionEng") & " " & objRsExpEdu("eduDiploma1Eng")) & "\line " & Trim(objRsExpEdu("edsDescriptionEng") & " " & objRsExpEdu("id_EduSubject1Eng")), 1, "1#|#1#|#1#|#1"
		i=i+1
		objRsExpEdu.MoveNext
	WEnd
	End If 
	objRsExpEdu.Close  
	Set objRsExpEdu=Nothing
	WriteTableFooter

' Employment records
	WriteTableHeader
	WriteSimpleRow "\b " & GetLabel(sCvLanguage, "Employment Record") & "\b0:"
	WriteTableFooter

	While Not objRsExpWke.Eof
		WriteGridTableHeader
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
		
		WriteGridDataRow 2, "31%#|#70%", "\b " & GetLabel(sCvLanguage, "Dates") & "\b0 " & "#|#" & ConvertDateForText(objRsExpWke("wkeStartDate"), " ", "MonthYear") & " - " & sExperienceEndDate, 0, "0#|#0"
		WriteGridDataRow 2, "31%#|#70%", "\b " & GetLabel(sCvLanguage, "Employer") & "\b0 "  & "#|#" &  objRsExpWke("wkeOrgNameEng"), 0, "0#|#0"
		
		sCountries = GetExpertExperienceCountryList(iCvID, objRsExpWke("id_ExpWke"), sCvLanguage)
		WriteGridDataRow 2, "31%#|#70%", "\b " & GetLabel(sCvLanguage, "Location") & "\b0 "  & "#|#" &  sCountries, 0, "0#|#0"
		WriteGridDataRow 2, "31%#|#70%", "\b " & GetLabel(sCvLanguage, "Position held") & "\b0 " & "#|#" & objRsExpWke("wkePositionEng"), 0, "0#|#0"
		dflag=0
		sDescription=ConvertText2RTF(objRsExpWke("wkeDescriptionEng"))
		WriteGridDataRow 2, "31%#|#70%", "\b " & GetLabel(sCvLanguage, "Description of duties") & "\b0 "  & "#|#" &  sDescription, 0, "0#|#0"
		
		WriteTableFooter

	objRsExpWke.MoveNext
	WEnd
	objRsExpWke.Close
	Set objRsExpWke=Nothing

' Languages
	WriteTableHeader
	WriteGridDataRow 1, "100%", "\b " & GetLabel(sCvLanguage, "Languages") & "\b0:", 0, "0"
	WriteTableFooterNoPar
	If (Not objRsExpLngOther.Eof) Or (Not objRsExpLngNative.Eof) Then
	WriteGridTableHeader
	WriteGridDataRow 4, "31%#|#21%#|#22%#|#21%", "\b " & GetLabel(sCvLanguage, "Language") & "\b0 #|#\b " & GetLabel(sCvLanguage, "Reading") & "\b0 #|#\b " & GetLabel(sCvLanguage, "Speaking") & "\b0 #|#\b " & GetLabel(sCvLanguage, "Writing") & "\b0 ", 1, "1#|#1#|#1#|#1"
	End If
	While Not objRsExpLngNative.Eof
		On Error Resume Next
			sTempLanguage = ReplaceIfEmpty(objRsExpLngNative("lngName" & sCvLanguage), objRsExpLngNative("lngNameEng"))
		On Error GoTo 0
		WriteGridDataRow 4, "31%#|#21%#|#22%#|#21%", sTempLanguage & "#|#" & arrLanguageLevelTitle(objRsExpLngNative("exlReading")) & "#|#" & arrLanguageLevelTitle(objRsExpLngNative("exlSpeaking")) & "#|#" & arrLanguageLevelTitle(objRsExpLngNative("exlWriting")), 1, "1#|#1#|#1#|#1"
		objRsExpLngNative.MoveNext
	WEnd
	While Not objRsExpLngOther.Eof
		On Error Resume Next
			sTempLanguage = ReplaceIfEmpty(objRsExpLngOther("lngName" & sCvLanguage), objRsExpLngOther("lngNameEng"))
		On Error GoTo 0
		WriteGridDataRow 4, "31%#|#21%#|#22%#|#21%", sTempLanguage & "#|#" & arrLanguageLevelTitle(objRsExpLngOther("exlReading")) & "#|#" & arrLanguageLevelTitle(objRsExpLngOther("exlSpeaking")) & "#|#" & arrLanguageLevelTitle(objRsExpLngOther("exlWriting")), 1, "1#|#1#|#1#|#1"
		objRsExpLngOther.MoveNext
	WEnd
	WriteTableFooter
	objRsExpLngNative.Close
	Set objRsExpLngNative=Nothing	
	objRsExpLngOther.Close
	Set objRsExpLngOther=Nothing
	
	Response.Write ConvertText2RTF("\b " & GetLabel(sCvLanguage, "Certification") & ":\b0\line\line " & GetLabel(sCvLanguage, "I, the undersigned, certify (AFB)...") & "\line\line\line _________________________________________________________   Date: _________________\line \i " & GetLabel(sCvLanguage, "Signature of staff member or authorized representative of the staff") & " \i0                     \i " & GetLabel(sCvLanguage, "Day/Month/Year") & " \i0 ")

Response.Write("\par }}" & vbCrLf)
%>
