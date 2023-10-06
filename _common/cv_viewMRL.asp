<%
'--------------------------------------------------------------------
'
' Expert's CV. View in Cabinet Merlin format
'
'--------------------------------------------------------------------
sForceCvLanguage = cLanguageFrench
%>
<!--#include file="cv_data.asp"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" type="text/css" href="<% =sHomePath %>styles.css">
<title>CV in Cabinet Merlin format</title>
</head>

<body bgcolor="#FFFFFF" topmargin=0 leftmargin=0 marginheight=0 marginwidth=0>
<% ShowTopMenu %>

<p class="ttl" align="center"><b>Curriculum Vitae - Cabinet Merlin format</b></p>

<table width="98%" cellpadding=0 cellspacing=0 border=0 align="center">
<tr><td width="85%" valign="top">
<br />

<%
	Response.Write "" & _
	"<table width=""55%"" border=0 bgcolor=""#0066CC"" cellpadding=1 cellspacing=0 align=center><tr><td>" & vbCrLf & _
	"  <table cellspacing=0 cellpadding=0 align=""center"" width=""100%"" bgcolor=""#EEEEEE"">" & vbCrLf
	colCellBG = "#EEEEEE"
	WriteDataRow1Column "<b>" & sLastName & sFirstName & "</b>", ""
	WriteDataRow1Column "Né en " & Year(sBirthDate), ""
	If Not objRsExpWke.Eof Then
		WriteDataRow1Column objRsExpWke("wkeOrgNameEng"), ""
		WriteDataRow1Column "Depuis " & Year(objRsExpWke("wkeStartDate")), ""
		WriteDataRow1Column objRsExpWke("wkePositionEng"), ""
	End If
	colCellBG = "#FFFFFF"
	WriteTableFooter

' DOMAINES DE COMPETENCES
	WriteTableHeader
	WriteDataRowHeader GetLabel(sCvLanguage, "DOMAINES DE COMPETENCES")
	WriteDataRow1Column sKeyQualification, ""
	WriteTableFooter

' FORMATION
	WriteTableHeader
	WriteDataRowHeader GetLabel(sCvLanguage, "FORMATION")
	WriteTableFooterWithoutSpace
	Dim sPeriod
	Dim sEduDiploma
	
	WriteGridTableHeader

	While Not objRsExpEdu.Eof
		sPeriod=""
		If Not (IsNull(objRsExpEdu("eduStartDate")) And IsNull(objRsExpEdu("eduEndDate"))) Then
			sPeriod=sPeriod & Year(ReplaceIfEmpty(objRsExpEdu("eduEndDate"), objRsExpEdu("eduStartDate"))) & ""
		End If
		sEduDiploma=""
		On Error Resume Next
			sEduDiploma = ReplaceIfEmpty(objRsExpEdu("edtDescription" & sCvLanguage), objRsExpEdu("edtDescriptionEng"))
		On Error GoTo 0
		
		sEduSubject=""
		On Error Resume Next
			sEduSubject = ReplaceIfEmpty(objRsExpEdu("edsDescription" & sCvLanguage), objRsExpEdu("edsDescriptionEng"))
		On Error GoTo 0
		If sEduSubject="Other" Then 
			sEduSubject=""
		End If
		If Len(objRsExpEdu("id_EduSubject1Eng"))>0 Then
			sEduSubject=sEduSubject & Trim(" " & objRsExpEdu("id_EduSubject1Eng"))
		End If
		If Len(sEduSubject)>0 Then
			sEduSubject="<br />" & sEduSubject
		End If
		
		WriteGridDataRow 4, "20%#|#40%#|#30%#|#10%", "<b>" & sPeriod & "</b>#|#" & Trim(sEduDiploma & " " & objRsExpEdu("eduDiploma1Eng") & " " & sEduSubject) & "#|#" & objRsExpEdu("InstNameEng") & "#|#" & objRsExpEdu("InstLocationEng")
		objRsExpEdu.MoveNext
	WEnd
	objRsExpEdu.Close  
	Set objRsExpEdu=Nothing
	WriteGridTableFooter

' EXPERIENCE PROFESSIONNELLE
	WriteTableHeader
	WriteDataRowHeader GetLabel(sCvLanguage, "EXPERIENCE PROFESSIONNELLE")
	WriteTableFooterWithoutSpace

	While Not objRsExpWke.Eof
		If objRsExpWke("TypeofWke")=2 Or (objRsExpWke("TypeofWke")<=1 And Len(ReplaceIfEmpty(objRsExpWke("wkePrjTitleEng"), ""))<=1) Then
		WriteGridTableHeader
			Dim sExperienceEndDate
			sExperienceEndDate=""
			If objRsExpWke("wkeEndDateOpen")=1 Then
				sExperienceEndDate=sExperienceEndDate & GetLabel(sCvLanguage, "Ongoing")
				If IsDate(objRsExpWke("wkeEndDate")) And Not IsNull(objRsExpWke("wkeEndDate")) Then sExperienceEndDate=sExperienceEndDate & " ("
			End If
			If IsDate(objRsExpWke("wkeEndDate")) And Not IsNull(objRsExpWke("wkeEndDate")) Then
				sExperienceEndDate=sExperienceEndDate & Year(objRsExpWke("wkeEndDate"))
				If objRsExpWke("wkeEndDateOpen")=1 Then sExperienceEndDate=sExperienceEndDate & ")"
			End If
			
			WriteGridDataRow 3, "20%#|#30%#|#50%", "<b>" & Year(objRsExpWke("wkeStartDate")) & " - " & sExperienceEndDate & "</b>#|#" &  objRsExpWke("wkeOrgNameEng") & "#|#" & objRsExpWke("wkePositionEng")
			dflag=0
		WriteTableFooter
		End If

	objRsExpWke.MoveNext
	WEnd
	
' COMPETENCES LINGUISTIQUES
	WriteTableHeader
	WriteDataRowHeader GetLabel(sCvLanguage, "COMPETENCES LINGUISTIQUES")
	WriteTableFooterWithoutSpace

	If (Not objRsExpLngOther.Eof) Or (Not objRsExpLngNative.Eof) Then
		WriteGridTableHeader
		WriteGridDataRow 4, "25%#|#25%#|#25%#|#25%", "<font color=""999999""><b>" & GetLabel(sCvLanguage, "Language") & "</b>#|#<font color=""999999""><b>" & GetLabel(sCvLanguage, "Reading") & "</b></font>#|#<font color=""999999""><b>" & GetLabel(sCvLanguage, "Writing") & "</b></font>#|#<font color=""999999""><b>" & GetLabel(sCvLanguage, "Speaking") & "</b></font>"
		While Not objRsExpLngNative.Eof
			On Error Resume Next
				sTempLanguage = ReplaceIfEmpty(objRsExpLngNative("lngName" & sCvLanguage), objRsExpLngNative("lngNameEng"))
			On Error GoTo 0
			WriteGridDataRow 4, "", sTempLanguage & "#|#" & arrLanguageLevelTitle(objRsExpLngNative("exlReading")) & "#|#" & arrLanguageLevelTitle(objRsExpLngNative("exlWriting")) & "#|#" & arrLanguageLevelTitle(objRsExpLngNative("exlSpeaking"))
			objRsExpLngNative.MoveNext
		WEnd
		While Not objRsExpLngOther.Eof
			On Error Resume Next
				sTempLanguage = ReplaceIfEmpty(objRsExpLngOther("lngName" & sCvLanguage), objRsExpLngOther("lngNameEng"))
			On Error GoTo 0
			WriteGridDataRow 4, "", sTempLanguage & "#|#" & arrLanguageLevelTitle(objRsExpLngOther("exlReading")) & "#|#" & arrLanguageLevelTitle(objRsExpLngOther("exlWriting")) & "#|#" & arrLanguageLevelTitle(objRsExpLngOther("exlSpeaking"))
			objRsExpLngOther.MoveNext
		WEnd
		WriteTableFooter
	End If
	objRsExpLngNative.Close
	Set objRsExpLngNative=Nothing	
	objRsExpLngOther.Close
	Set objRsExpLngOther=Nothing

' PRINCIPALES REFERENCES
	WriteTableHeader
	WriteDataRowHeader GetLabel(sCvLanguage, "PRINCIPALES REFERENCES")
	WriteTableFooterWithoutSpace

	objRsExpWke.MoveFirst
	While Not objRsExpWke.Eof
		If objRsExpWke("TypeofWke")=3 Or (objRsExpWke("TypeofWke")<=1 And Len(ReplaceIfEmpty(objRsExpWke("wkePrjTitleEng"), ""))>1) Then
		'If Len(ReplaceIfEmpty(objRsExpWke("wkePrjTitleEng"), ""))>1 Then
			WriteGridTableHeader
		
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
			
			WriteGridDataRow 3, "20%#|#30%#|#50%", "<b>" & sYear & "</b>" & "#|#" &  "<font color=""999999""><b>Maître d'Ouvrage</b></font>" & "#|#" & "<b>" & objRsExpWke("wkeOrgNameEng") & "</b>"
			WriteGridDataRow 3, "20%#|#30%#|#50%", "" & "#|#" &  "<font color=""999999""><b>Projet</b></font>" & "#|#" & objRsExpWke("wkePrjTitleEng")
			WriteGridDataRow 3, "20%#|#30%#|#50%", "" & "#|#" &  "<font color=""999999""><b>Fonction occupée</b></font>" & "#|#" & objRsExpWke("wkePositionEng")
			sDescription=ConvertText(objRsExpWke("wkeDescriptionEng"))
			WriteGridDataRow 3, "20%#|#30%#|#50%", "" & "#|#" &  "<font color=""999999""><b>Descriptif</b></font>" & "#|#" & sDescription
			
			WriteTableFooter			
		End If
	objRsExpWke.MoveNext
	WEnd
	objRsExpWke.Close
	Set objRsExpWke=Nothing
	

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



