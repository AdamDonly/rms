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
' Log: 35 - Download CV
iLogResult = LogActivity(35, "CVID=" & Cstr(iCvID) & " Format: " & sCvFormat, "", "")

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
	WriteDataRow "Proposed position:", ""
	WriteDataRow "Name of Firm:", ""
	WriteDataRow "Name of Staff:", sFullName
	WriteDataRow "Profession:", sProfession
	WriteDataRow "Date of Birth:", ConvertDateForText(sBirthDate, " ", "DDMMYYY")
	
	WriteSpaceRow
	WriteDataRow "Address:\line \line (phone/fax/e-mail)", ConvertText2RTF(sPermAddress)
	
	WriteSpaceRow
	WriteSpaceRow
	WriteDataRow "Years with Firm/Entity:", ""
	WriteDataRow "Nationality:", sNationality 
	WriteDataRow "Membership\~in\line professional\~societies:", " \line " & sMemberships
	WriteDataRow "Detailed tasks assigned:", "" 
	WriteDataRow "Key qualifications:", ConvertText2RTF(sKeyQualification)
	WriteTableFooter

' Education
	WriteTableHeader
	WriteGridDataRow 1, "100%", "Education:", 0, "0"
	WriteTableFooter
	WriteGridTableHeader
	WriteGridDataRow 4, "31%#|#15%#|#15%#|#40%", "\b Institution\b0 #|#\b Start\~date\b0 #|#\b End\~date\b0 #|#\b Degree\~/\~Diploma\~obtained\b0 ", 1, "1#|#1#|#1#|#1"

	If objRsExpEdu.Eof Then
		WriteGridDataRow 4, "31%#|#15%#|#15%#|#40%", "-#|#-#|#-#|#-", 1, "1#|#1#|#1#|#1"
	Else
        i=1
	While Not objRsExpEdu.Eof
		WriteGridDataRow 4, "31%#|#15%#|#15%#|#40%", objRsExpEdu("InstNameEng") & "#|#" & ConvertDateForText(objRsExpEdu("eduStartDate"), " ", "MMYYY") & "#|#" & ConvertDateForText(objRsExpEdu("eduEndDate"), " ", "MMYYY") & "#|#" & Trim(objRsExpEdu("edtDescriptionEng") & " " & objRsExpEdu("eduDiploma1Eng")) & "\line " & Trim(objRsExpEdu("edsDescriptionEng") & " " & objRsExpEdu("id_EduSubject1Eng")), 1, "1#|#1#|#1#|#1"
		i=i+1
		objRsExpEdu.MoveNext
	WEnd
	End If 
	objRsExpEdu.Close  
	Set objRsExpEdu=Nothing
        i=1
	While Not objRsExpTrn.eof
		WriteGridDataRow 4, "31%#|#15%#|#15%#|#40%", objRsExpTrn("eduDiploma1Eng") & "#|#" & ConvertDateForText(objRsExpTrn("eduStartDate"), " ", "MMYYY") & "#|#" & ConvertDateForText(objRsExpTrn("eduEndDate"), " ", "MMYYY") & "#|#" & objRsExpTrn("edtDescriptionEng"), 1, "1#|#1#|#1#|#1"
		i=i+1
		objRsExpTrn.MoveNext
	WEnd
	objRsExpTrn.Close  
	Set objRsExpTrn=Nothing
	WriteTableFooter


' Employment records
	WriteTableHeader
	WriteGridDataRow 1, "100%", "Employment\~records:", 0, "0"
	WriteTableFooter

	While Not objRsExpWke.Eof
		
		WriteGridTableHeader
		WriteGridDataRow 2, "31%#|#70%", "\b Date\b0 " & "#|#" & ConvertDateForText(objRsExpWke("wkeStartDate"), " ", "MMYYY") & " - " & ConvertDateForText(objRsExpWke("wkeEndDate"), " ", "MMYYY"), 1, "1#|#1"
		WriteGridDataRow 2, "31%#|#70%", "\b Employer\b0 "  & "#|#" &  objRsExpWke("wkeOrgNameEng"), 1, "1#|#1"
		WriteGridDataRow 2, "31%#|#70%", "\b Position held\b0 " & "#|#" & objRsExpWke("wkePositionEng"), 1, "1#|#1"
		dflag=0
		sDescription=ConvertText(objRsExpWke("wkeDescriptionEng"))
		WriteGridDataRow 2, "31%#|#70%", "\b Description\~of\~duties\b0 " & "#|#" & sDescription, 1, "1#|#1"
		WriteTableFooter

	objRsExpWke.MoveNext
	WEnd
	objRsExpWke.Close
	Set objRsExpWke=Nothing

' Languages
	WriteTableHeader
	WriteGridDataRow 1, "100%", "Languages:", 0, "0"
	WriteTableFooter

	If (Not objRsExpLngOther.Eof) Or (Not objRsExpLngNative.Eof) Then
	WriteGridTableHeader
	WriteGridDataRow 4, "31%#|#21%#|#22%#|#21%", "\b Language\b0 #|#\b Reading\b0 #|#\b Speaking\b0 #|#\b Writing\b0 ", 1, "1#|#1#|#1#|#1"
	End If

	While Not objRsExpLngNative.Eof
		WriteGridDataRow 4, "31%#|#21%#|#22%#|#21%", objRsExpLngNative("lngNameEng") & "#|#" & arrLanguageLevelTitle(objRsExpLngNative("exlReading")) & "#|#" & arrLanguageLevelTitle(objRsExpLngNative("exlSpeaking")) & "#|#" & arrLanguageLevelTitle(objRsExpLngNative("exlWriting")), 1, "1#|#1#|#1#|#1"
		objRsExpLngNative.MoveNext
	WEnd

	While Not objRsExpLngOther.Eof
		WriteGridDataRow 4, "31%#|#21%#|#22%#|#21%", objRsExpLngOther("lngNameEng") & "#|#" & arrLanguageLevelTitle(objRsExpLngOther("exlReading")) & "#|#" & arrLanguageLevelTitle(objRsExpLngOther("exlSpeaking")) & "#|#" & arrLanguageLevelTitle(objRsExpLngOther("exlWriting")), 1, "1#|#1#|#1#|#1"
		objRsExpLngOther.MoveNext
	WEnd
	WriteTableFooter
	objRsExpLngNative.Close
	Set objRsExpLngNative=Nothing	
	objRsExpLngOther.Close
	Set objRsExpLngOther=Nothing	

	Response.Write ConvertText2RTF("<b>Certification:</b> <br><br>I, the undersigned, certify that to the best of my knowledge and belief, these data correctly describe me, my qualifications, and my experience.<br><br><br>__________________________________________________________<br>Signature of staff member and authorized representative of the firm &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; Date: ________")

Response.Write("\par }}" & vbCrLf)
%>
