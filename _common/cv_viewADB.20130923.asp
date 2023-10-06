<%
'--------------------------------------------------------------------
'
' Expert's CV. View in ADB format
'
'--------------------------------------------------------------------
%>
<!--#include file="cv_data.asp"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<LINK REL="stylesheet" TYPE="text/css" href="<% =sHomePath %>styles.css">
<title>CV in Asian Development Bank Format</title>
</head>

<body bgcolor="#FFFFFF" topmargin=0 leftmargin=0 marginheight=0 marginwidth=0>
<% ShowTopMenu %>

<p class="ttl" align="center"><b>
<% If iCvID=4238 Then %>
	Sample Curriculum Vitae* - Asian Development Bank format
<% Else %>
        Curriculum Vitae - Asian Development Bank format
<% End If %>
</b></p>

<table width="98%" cellpadding=0 cellspacing=0 border=0 align="center">
<tr><td width="85%" valign="top">
<br />

<%
' Personal information
	WriteTableHeader
	WriteDataRow "1. Proposed position<br>&nbsp;&nbsp;&nbsp;&nbsp;for the project:", "<font color=""#C0C0C0"">(only one candidate should be nominated for each position)</font>"
	WriteDataRow "2. Name:", sFullName
	WriteDataRow "3. Date of birth:", ConvertDateForText(sBirthDate, "&nbsp;", "DDMMYYYY")
	WriteDataRow "4. Nationality:", sNationality 
	WriteDataRow "5. Personal address:", sPermAddress
	WriteTableFooter

' Education
	WriteTableHeader
	WriteDataRow "6. Education:", " "
	WriteTableFooter
	WriteGridTableHeader
	WriteGridDataRow 4, "25%#|#13%#|#13%#|#50%", "<b>Institution</b>#|#<b>Start&nbsp;date</b>#|#<b>End&nbsp;date</b>#|#<b>Degree&nbsp;/&nbsp;Diploma&nbsp;obtained</b>"

	If objRsExpEdu.Eof Then
		WriteGridDataRow 4, "", "-#|#-#|#-#|#-"
	Else
        i=1
	While Not objRsExpEdu.Eof
		WriteGridDataRow 4, "", objRsExpEdu("InstNameEng") & "#|#" & ConvertDateForText(objRsExpEdu("eduStartDate"), "&nbsp;", "MMYYYY") & "#|#" & ConvertDateForText(objRsExpEdu("eduEndDate"), "&nbsp;", "MMYYYY") & "#|#" & Trim(objRsExpEdu("edtDescriptionEng") & " " & objRsExpEdu("eduDiploma1Eng")) & "<br>" & Trim(objRsExpEdu("edsDescriptionEng") & " " & objRsExpEdu("id_EduSubject1Eng"))
																																																							 
		i=i+1
		objRsExpEdu.MoveNext
	WEnd
	End If 
	objRsExpEdu.Close  
	Set objRsExpEdu=Nothing
	WriteTableFooter


' Training
	WriteTableHeader
	WriteDataRow "7. Other training:", " "
	WriteTableFooter
	WriteGridTableHeader
	WriteGridDataRow 4, "25%#|#13%#|#13%#|#50%", "<b>Training&nbsp;title</b>#|#<b>Start&nbsp;date</b>#|#<b>End&nbsp;date</b>#|#<b>Degree&nbsp;/&nbsp;Diploma&nbsp;obtained</b>"

	If objRsExpTrn.eof Then
		WriteGridDataRow 4, "", "-#|#-#|#-#|#-"
	Else
        i=1
	While Not objRsExpTrn.Eof
		WriteGridDataRow 4, "", objRsExpTrn("eduDiploma1Eng") & "#|#" & ConvertDateForText(objRsExpTrn("eduStartDate"), "&nbsp;", "MMYYYY") & "#|#" & ConvertDateForText(objRsExpTrn("eduEndDate"), "&nbsp;", "MMYYYY") & "#|#" & objRsExpTrn("edtDescriptionEng")
		i=i+1
		objRsExpTrn.MoveNext
	WEnd
	End If 
	objRsExpTrn.Close  
	Set objRsExpTrn=Nothing
	WriteTableFooter

' Languages
	WriteTableHeader
	WriteDataRow "8.&nbsp;Languages&nbsp;and&nbsp;degree&nbsp;of&nbsp;proficency:", " "
	WriteTableFooter

	If (Not objRsExpLngOther.Eof) Or (Not objRsExpLngNative.Eof) Then
		WriteGridTableHeader
		WriteGridDataRow 4, "25%#|#25%#|#25%#|#25%", "<b>Language</b>#|#<b>Reading</b>#|#<b>Speaking</b>#|#<b>Writing</b>"

		While Not objRsExpLngNative.Eof
			WriteGridDataRow 4, "", objRsExpLngNative("lngNameEng") & "#|#" & arrLanguageLevelTitle(objRsExpLngNative("exlReading")) & "#|#" & arrLanguageLevelTitle(objRsExpLngNative("exlSpeaking")) & "#|#" & arrLanguageLevelTitle(objRsExpLngNative("exlWriting"))
			objRsExpLngNative.MoveNext
		WEnd

		While Not objRsExpLngOther.Eof
			WriteGridDataRow 4, "", objRsExpLngOther("lngNameEng") & "#|#" & arrLanguageLevelTitle(objRsExpLngOther("exlReading")) & "#|#" & arrLanguageLevelTitle(objRsExpLngOther("exlSpeaking")) & "#|#" & arrLanguageLevelTitle(objRsExpLngOther("exlWriting"))
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
	WriteDataRow "9.&nbsp;Membership&nbsp;in&nbsp;professional&nbsp;societies:", sMemberships
	WriteTableFooter

' Countries of work experience
	WriteTableHeader
	WriteDataRow "10.&nbsp;Countries&nbsp;of&nbsp;work&nbsp;experience:", " "
	WriteTableFooter

	Set objRsExpCou=GetDataRecordsetSP("usp_ExpCvvADBCouSelect", Array( _
		Array(, adInteger, , iCvID)))
	If Not objRsExpCou.Eof Then
	WriteGridTableHeader
	WriteGridDataRow 2, "25%#|#75%", "<b>Country</b>#|#<b>Period</b>"

	While Not objRsExpCou.Eof
		WriteGridDataRow 2, "", objRsExpCou(0) & "#|#" & objRsExpCou(1)
		objRsExpCou.MoveNext
	WEnd
	End If
	WriteTableFooter
	objRsExpCou.Close
	Set objRsExpCou=Nothing	


' Employment records
	WriteTableHeader
	WriteDataRow "11.&nbsp;Employment&nbsp;records:", " "
	WriteTableFooter

	While Not objRsExpWke.Eof
		
		WriteGridTableHeader
		Dim sExperienceEndDate
		sExperienceEndDate=""
		If objRsExpWke("wkeEndDateOpen")=1 Then
			sExperienceEndDate=sExperienceEndDate & "Ongoing"
			If IsDate(objRsExpWke("wkeEndDate")) And Not IsNull(objRsExpWke("wkeEndDate")) Then sExperienceEndDate=sExperienceEndDate & " ("
		End If
		If IsDate(objRsExpWke("wkeEndDate")) And Not IsNull(objRsExpWke("wkeEndDate")) Then
			sExperienceEndDate=sExperienceEndDate & ConvertDateForText(objRsExpWke("wkeEndDate"), " ", "MMYYYY")
			If objRsExpWke("wkeEndDateOpen")=1 Then sExperienceEndDate=sExperienceEndDate & ")"
		End If
			
		WriteGridDataRow 2, "", "<b>Project title</b>"  & "#|#" &  objRsExpWke("wkePrjTitleEng")
		WriteGridDataRow 2, "25%#|#75%", "<b>Date</b>" & "#|#" & ConvertDateForText(objRsExpWke("wkeStartDate"), "&nbsp;", "MMYYYY") & " - " & sExperienceEndDate
		sCountries = GetExpertExperienceCountryList(iCvID, objRsExpWke("id_ExpWke"), sCvLanguage)
		WriteGridDataRow 2, "", "<b>Location</b>"  & "#|#" &  sCountries
		WriteGridDataRow 2, "", "<b>Employer</b>"  & "#|#" &  objRsExpWke("wkeOrgNameEng")
		WriteGridDataRow 2, "", "<b>Project features</b>"  & "#|#" &  objRsExpWke("wkeProjectDescription")
		WriteGridDataRow 2, "", "<b>Position held</b>" & "#|#" & objRsExpWke("wkePositionEng")
		'dflag=0
		sDescription=ConvertText(objRsExpWke("wkeDescriptionEng"))
		WriteGridDataRow 2, "", "<b>Description of duties</b>" & "#|#" & sDescription
		WriteTableFooter

	objRsExpWke.MoveNext
	WEnd
	objRsExpWke.Close
	Set objRsExpWke=Nothing

	WriteTableHeader
	WriteDataRow "12.&nbsp;Detailed&nbsp;tasks&nbsp;assigned:", "<font color=""#C0C0C0""> ( Work undertaken that best illustrates capability to handle the tasks assigned ) </font>"
	WriteTableFooter
	%>

	<table cellspacing=0 cellpadding=0 align="center" width="98%">
	<tr><td width="100%" bgcolor="#FFFFFF"><p class="txt">13.&nbsp;Certification:</b> <br> <font color="#C0C0C0">(Please follow exactly the following format. Omission will be seen as noncompliance)</font><br><br>I, the undersigned, certify that <br>(i)&nbsp;&nbsp;&nbsp;&nbsp;I am not a former ADB Staff or if I am, I have retired/resigned from ADB for more than twelve (12) months ago; <br>(ii)&nbsp; &nbsp;I am not a close relative of ADB personnel; and <br>(iii) &nbsp;&nbsp;to the best of my knowledge and belief, this biodata correctly describes myself, my qualifications, and my experience. <br><br>I understand that any willful misstatement described herein may lead to my disqualification or dismissal, if engaged. I have been employed by [name of the firm] continuously for the last (12) months as regular full time staff.</b><br><br><br>Signature &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; Date of signing <font color="#C0C0C0">&nbsp; ( Day / Month / Year )</font></td></tr>
	</table><br>

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

