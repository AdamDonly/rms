<%
'--------------------------------------------------------------------
'
' Expert's CV. Save in ADB format
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
Response.Write("{\info{\title CV of " & sFullName & "}{\author assortis.com}{\company assortis.com}{\subject Asian Development Bank format}{\category Expert CV}{\keywords " & sProfession & "}{\doccomm Downloaded: " & ConvertDateForText(Date()," ", "DDMMYYYY") & "}}\paperw11907\paperh16840\margl1797\margr1797\margt789\margb689\viewkind1\viewscale100\titlepg " & vbCrLf)

Response.Write("{\footerf \pard\plain \s16\qr \li0\ri0\widctlpar\brdrt\brdrs\brdrw10\brsp20 \tqc\tx4153\tqr\tx8306\aspalpha\aspnum\faauto\adjustright\rin0\lin0\itap0 \fs24\lang2057\langfe1033\cgrid\langnp2057\langfenp1033 {\f1\fs16\cf9\lang2060\langfe1033\langnp2060 Page\~: }{\field{\*\fldinst {\cs17\f1\fs16\cf9  PAGE }}{\fldrslt {\cs17\f1\fs16\cf9\lang1024\langfe1024\noproof 2}}}{\f1\fs16\cf9\lang2060\langfe1033\langnp2060 \par }}{\*\pnseclvl1\pnucrm\pnstart1\pnindent720\pnhang{\pntxta .}}{\*\pnseclvl2\pnucltr\pnstart1\pnindent720\pnhang{\pntxta .}}{\*\pnseclvl3\pndec\pnstart1\pnindent720\pnhang{\pntxta .}}{\*\pnseclvl4\pnlcltr\pnstart1\pnindent720\pnhang{\pntxta )}}" & vbCrLf)
Response.Write("{\footer \pard\plain \s16\qr \li0\ri0\widctlpar\brdrt\brdrs\brdrw10\brsp20 \tqc\tx4153\tqr\tx8306\aspalpha\aspnum\faauto\adjustright\rin0\lin0\itap0 \fs24\lang2057\langfe1033\cgrid\langnp2057\langfenp1033 {\f1\fs16\cf9\lang2060\langfe1033\langnp2060 Page\~: }{\field{\*\fldinst {\cs17\f1\fs16\cf9  PAGE }}{\fldrslt {\cs17\f1\fs16\cf9\lang1024\langfe1024\noproof 2}}}{\f1\fs16\cf9\lang2060\langfe1033\langnp2060 \par }}{\*\pnseclvl1\pnucrm\pnstart1\pnindent720\pnhang{\pntxta .}}{\*\pnseclvl2\pnucltr\pnstart1\pnindent720\pnhang{\pntxta .}}{\*\pnseclvl3\pndec\pnstart1\pnindent720\pnhang{\pntxta .}}{\*\pnseclvl4\pnlcltr\pnstart1\pnindent720\pnhang{\pntxta )}}" & vbCrLf)
Response.Write("\qr\f1\fs18\cf16 \b FORM TECH-3\b0\par\par" & vbCrLf)
Response.Write("\qc\f1\fs18\cf2 \b CURRICULUM VITAE (CV) FORMAT TO BE SUBMITTED WITH PROPOSAL\b0\par" & vbCrLf)
Response.Write("{" & vbCrLf)
Response.Write("\par\ql\f1\fs18\cf2" & vbCrLf)

' Personal information
	WriteTableHeader
	WriteDataRow "1. Proposed position \line for this project:", "\ql\f1\fs18\cf16 (only one candidate should be nominated for each position) \ql\f1\fs18\cf2"
	WriteDataRow "2. Name:", sFullName
	WriteDataRow "3. Date of birth:", ConvertDateForText(sBirthDate, " ", "DDMMYYYY")
	WriteDataRow "4. Nationality:", sNationality 
	sPermAddress=Replace(sPermAddress,"<p class=""txt"">","")
	sPermAddress=Replace(sPermAddress,"</p>","\line ")
	WriteDataRow "5. Personal address:", sPermAddress
	WriteTableFooter

' Education
	WriteTableHeader
	WriteGridDataRow 1, "100%", "6. Education:", 0, "0"
	WriteTableFooter
	WriteGridTableHeader
	WriteGridDataRow 4, "31%#|#15%#|#15%#|#40%", "\b Institution\b0 #|#\b Start\~date\b0 #|#\b End\~date\b0 #|#\b Degree\~/\~Diploma\~obtained\b0 ", 1, "1#|#1#|#1#|#1"

	If objRsExpEdu.Eof Then
		WriteGridDataRow 4, "31%#|#15%#|#15%#|#40%", "-#|#-#|#-#|#-", 1, "1#|#1#|#1#|#1"
	Else
        i=1
	While Not objRsExpEdu.Eof
		WriteGridDataRow 4, "31%#|#15%#|#15%#|#40%", objRsExpEdu("InstNameEng") & "#|#" & ConvertDateForText(objRsExpEdu("eduStartDate"), " ", "MMYYYY") & "#|#" &  ConvertDateForText(objRsExpEdu("eduEndDate"), " ", "MMYYYY") & "#|#" & Trim(objRsExpEdu("edtDescriptionEng") & " " & objRsExpEdu("eduDiploma1Eng")) & "\line " & Trim(objRsExpEdu("edsDescriptionEng") & " " & objRsExpEdu("id_EduSubject1Eng")), 1, "1#|#1#|#1#|#1"
		i=i+1
		objRsExpEdu.MoveNext
	WEnd
	End If 
	objRsExpEdu.Close  
	Set objRsExpEdu=Nothing
	WriteTableFooter


' Training
	WriteTableHeader
	WriteGridDataRow 1, "100%", "7. Other training:", 0, "0"
	WriteTableFooter
	WriteGridTableHeader
	WriteGridDataRow 4, "31%#|#15%#|#15%#|#40%", "\b Training\~title\b0 #|#\b Start\~date\b0 #|#\b End\~date\b0 #|#\b Degree\~/\~Diploma\~obtained\b0 ", 1, "1#|#1#|#1#|#1"

	If objRsExpTrn.eof Then
		WriteGridDataRow 4, "31%#|#15%#|#15%#|#40%", "-#|#-#|#-#|#-", 1, "1#|#1#|#1#|#1"
	Else
        i=1
	While Not objRsExpTrn.Eof
		WriteGridDataRow 4, "31%#|#15%#|#15%#|#40%", objRsExpTrn("eduDiploma1Eng") & "#|#" & ConvertDateForText(objRsExpTrn("eduStartDate"), " ", "MMYYYY") & "#|#" & ConvertDateForText(objRsExpTrn("eduEndDate"), " ", "MMYYYY") & "#|#" & objRsExpTrn("edtDescriptionEng"), 1, "1#|#1#|#1#|#1"
		i=i+1
		objRsExpTrn.MoveNext
	WEnd
	End If 
	objRsExpTrn.Close  
	Set objRsExpTrn=Nothing
	WriteTableFooter


' Languages
	WriteTableHeader
	WriteGridDataRow 1, "100%", "8.\~Languages\~and\~degree\~of\~proficency:", 0, "0"
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

' Membership
	WriteTableHeader
	WriteGridDataRow 1, "100%", "9.\~Membership\~in\~professional\~societies:", 0, "0"
	WriteGridDataRow 2, "3%#|#97%", " #|#" & sMemberships , 0, "0#|#0"
	WriteTableFooter

' Countries of work experience
	WriteTableHeader
	WriteGridDataRow 1, "100%", "10.\~Countries\~of\~work\~experience:", 0, "0"
	WriteTableFooter

	Set objRsExpCou=GetDataRecordsetSP("usp_ExpCvvADBCouSelect", Array( _
		Array(, adInteger, , iCvID)))
	If Not objRsExpCou.Eof Then
	WriteGridTableHeader
	WriteGridDataRow 2, "31%#|#70%", "\b Country\b0 #|#\b Period\b0 ", 1, "1#|#1"

	While Not objRsExpCou.Eof
		WriteGridDataRow 2, "31%#|#70%", objRsExpCou(0) & "#|#" & objRsExpCou(1), 1, "1#|#1"
		objRsExpCou.MoveNext
	WEnd
	End If
	WriteTableFooter
	objRsExpCou.Close
	Set objRsExpCou=Nothing	


' Employment records
	WriteTableHeader
	WriteGridDataRow 1, "100%", "11.\~Employment\~records:", 0, "0"
	WriteTableFooter

	While Not objRsExpWke.Eof
		
		WriteGridTableHeader
		WriteGridDataRow 2, "31%#|#70%", "\b Date\b0 " & "#|#" & ConvertDateForText(objRsExpWke("wkeStartDate"), " ", "MMYYYY") & " - " & ConvertDateForText(objRsExpWke("wkeEndDate"), " ", "MMYYYY"), 1, "1#|#1"
		WriteGridDataRow 2, "31%#|#70%", "\b Employer\b0 "  & "#|#" &  objRsExpWke("wkeOrgNameEng"), 1, "1#|#1"
		WriteGridDataRow 2, "31%#|#70%", "\b Position held\b0 " & "#|#" & objRsExpWke("wkePositionEng"), 1, "1#|#1"
		sDescription=ConvertText(objRsExpWke("wkeDescriptionEng"))
		WriteGridDataRow 2, "31%#|#70%", "\b Description\~of\~duties\b0 " & "#|#" & sDescription, 1, "1#|#1"
		WriteTableFooter

	objRsExpWke.MoveNext
	WEnd
	objRsExpWke.Close
	Set objRsExpWke=Nothing


	WriteTableHeader
	WriteGridDataRow 2, "31%#|#70%", "12.\~Detailed\~tasks\~assigned:#|#\f1\fs18\cf16 ( Work undertaken that best illustrates capability to handle the tasks assigned ) \f1\fs18\cf2 ", 0, "0#|#0"
	WriteTableFooter

	WriteTableHeader
	WriteGridDataRow 1, "100%", "13.\~Certification:", 0, "0"
	WriteTableFooter

	Response.Write ConvertText2RTF("<font color=""#C0C0C0"">(Please follow exactly the following format. Omission will be seen as noncompliance)</font><br><br>I, the undersigned, certify that <br>(i)&nbsp;&nbsp;&nbsp;&nbsp;I am not a former ADB Staff or if I am, I have retired/resigned from ADB for more than twelve (12) months ago; <br>(ii)&nbsp; &nbsp;I am not a close relative of ADB personnel; and <br>(iii) &nbsp;&nbsp;to the best of my knowledge and belief, this biodata correctly describes myself, my qualifications, and my experience. <br><br>I understand that any willful misstatement described herein may lead to my disqualification or dismissal, if engaged. I have been employed by [name of the firm] continuously for the last (12) months as regular full time staff.</b><br><br><br>Signature &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; Date of signing <font color=""#C0C0C0"">&nbsp; ( Day / Month / Year )</font>")

Response.Write("\par }}" & vbCrLf)
%>
