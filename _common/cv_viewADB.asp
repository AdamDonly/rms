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
%>
	<table cellspacing=0 cellpadding=0 align="center" width="97%">
	<tr><td width="100%" bgcolor="#FFFFFF"><p class="txt"><b>11. <% =GetLabel(sCvLanguage, "Employment Record") %></b> <small><i>[<% =GetLabel(sCvLanguage, "Starting with present position...") %>]</i></small> :</td></tr>
	</table><br>
<%	
	While Not objRsExpWke.Eof
		If Len(ReplaceIfEmpty(objRsExpWke("wkePrjTitleEng"), ""))<=1 Then
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
			
			WriteGridDataRow 2, "25%#|#75%", "<b>" & GetLabel(sCvLanguage, "Dates") & "</b>" & "#|#" & ConvertDateForText(objRsExpWke("wkeStartDate"), "&nbsp;", "MonthYear") & " - " & sExperienceEndDate
			WriteGridDataRow 2, "", "<b>" & GetLabel(sCvLanguage, "Employer") & "</b>"  & "#|#" &  objRsExpWke("wkeOrgNameEng")

			'sCountries = GetExpertExperienceCountryList(iCvID, objRsExpWke("id_ExpWke"), sCvLanguage)
			'WriteGridDataRow 2, "", "<b>" & GetLabel(sCvLanguage, "Location") & "</b>" & "#|#" & sCountries
			WriteGridDataRow 2, "", "<b>" & GetLabel(sCvLanguage, "Position held") & "</b>" & "#|#" & objRsExpWke("wkePositionEng")
			dflag=0
			'sDescription=ConvertText(objRsExpWke("wkeDescriptionEng"))
			'WriteGridDataRow 2, "", "<b>" & GetLabel(sCvLanguage, "Description of duties") & ":</b>" & "#|#" & sDescription
		WriteTableFooter
		End If

	objRsExpWke.MoveNext
	WEnd
	'objRsExpWke.Close
	'Set objRsExpWke=Nothing

	
	WriteGridTableHeader
	WriteGridDataRow 2, "25%#|#75%", "<b>12. " & GetLabel(sCvLanguage, "Detailed Tasks Assigned") & "</b><br><small><i>[" & GetLabel(sCvLanguage, "List all tasks to be performed under this assignment") & "]</i></small>" & "#|#" & "<b>13. " & GetLabel(sCvLanguage, "Work Undertaken that Best Illustrates Capability to Handle the Tasks Assigned") & "</b><br><small><i>[" & GetLabel(sCvLanguage, "Among the assignments in which the staffs have been involved...") & ".]</i></small><br><br>"
	%>
	<tr>
	<td width="25%" bgcolor="#FFFFFF" valign="top"></small></td>
	<td width="75%" bgcolor="#FFFFFF" valign="top">
	<%
		objRsExpWke.MoveFirst
		While Not objRsExpWke.Eof
			If Len(ReplaceIfEmpty(objRsExpWke("wkePrjTitleEng"), ""))>1 Then
			
				Dim sYear
				If IsNull(objRsExpWke("wkeStartDate")) Then
					sYear=Year(objRsExpWke("wkeEndDate"))
				ElseIf IsNull(objRsExpWke("wkeEndDate")) Then
					sYear=Year(objRsExpWke("wkeStartDate"))
				ElseIf DateDiff("yyyy", objRsExpWke("wkeStartDate"), objRsExpWke("wkeEndDate"))<=0 Then
					sYear=Year(objRsExpWke("wkeStartDate"))
				Else 
					sYear=Year(objRsExpWke("wkeStartDate")) & " - " & Year(objRsExpWke("wkeEndDate"))
				End If			
				
				Response.Write "<p class=""txt""><b>" & GetLabel(sCvLanguage, "Name of assignment or project") & ": </b>" & objRsExpWke("wkePrjTitleEng") & "</p>"
				
				Response.Write "<p class=""txt""><b>" & GetLabel(sCvLanguage, "Year") & ": </b>" & sYear & "</p>"
				sCountries = GetExpertExperienceCountryList(iCvID, objRsExpWke("id_ExpWke"), sCvLanguage)
				Set objTempRs2=Nothing
				Response.Write "<p class=""txt""><b>" & GetLabel(sCvLanguage, "Location") & ": </b>" & sCountries & "</p>"
				Response.Write "<p class=""txt""><b>" & GetLabel(sCvLanguage, "Client") & ": </b>" & objRsExpWke("wkeBnfNameEng") & "</p>"
				Response.Write "<p class=""txt""><b>" & GetLabel(sCvLanguage, "Main project features") & ": </b>" & objRsExpWke("wkeProjectDescription") & "</p>"
				Response.Write "<p class=""txt""><b>" & GetLabel(sCvLanguage, "Position held") & ": </b>" & objRsExpWke("wkePositionEng") & "</p>"
				dflag=0
				sDescription=ConvertText(objRsExpWke("wkeDescriptionEng"))
				Response.Write "<p class=""txt""><b>" & GetLabel(sCvLanguage, "Activities performed") & ": </b>" & sDescription & "</p>"
				Response.Write "<br />"
			End If
		objRsExpWke.MoveNext
		WEnd
		objRsExpWke.Close
		Set objRsExpWke=Nothing
	%>
	</td>
	</tr>
	<%
	WriteTableFooter
	%>

	<table cellspacing=0 cellpadding=0 align="center" width="98%">
	<tr><td width="100%" bgcolor="#FFFFFF"><p class="txt">
	14.&nbsp;Certification:</b><br><br>
	I, the undersigned, certify to the best of my knowledge and belief<br>
	(i) this CV correctly describes my qualifications and my experience<br>
	(ii) I am employed by the Executing or the Implementing Agency<br>
	(iii) I am a close relative of a current ADB staff member<br>
	(iv) I am the spouse of a current ADB staff member<br>
	(v) I am a former ADB staff member.<br>
	&nbsp; &nbsp; *	If yes, I retired from ADB over 12 months ago<br>
	(vi) I am part of the team who wrote the terms of reference for this consulting services assignment.<br>
	(vii) I am sanctioned (not eligible for engagement) by ADB.<br><br>
	
	I understand that any willful misstatement described herein may lead to my disqualification or dismissal, if engaged.<br><br><br>Signature of expert &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; Date of signing <font color="#C0C0C0">&nbsp; ( Day / Month / Year )</font></td></tr>
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

