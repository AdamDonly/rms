<%
'--------------------------------------------------------------------
'
' Expert's CV. Save in Cabinet Merlin format
'
'--------------------------------------------------------------------
Response.Buffer = True
%>
<!--#include file="cv_data.asp"-->
<%
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

Dim sHeader
Set objFso=Server.CreateObject("Scripting.FileSystemObject")
Set fInTemplate=objFso.OpenTextFile(Server.MapPath("\_common") & "\cv_mrl.vrf", 1)
sHeader = fInTemplate.ReadAll & vbCrLf
sHeader = Replace(sHeader, "<#Name#>", sFullNameWithSpaces)
sHeader = Replace(sHeader, "<#YearOfBirth#>", Year(sBirthDate))
sHeader = Replace(sHeader, "<#CurrentCompanyName#>", objRsExpWke("wkeOrgNameEng"))
sHeader = Replace(sHeader, "<#CurrentCompanyStartYear#>", Year(objRsExpWke("wkeStartDate")))
sHeader = Replace(sHeader, "<#CurrentPosition#>", objRsExpWke("wkePositionEng"))
sHeader = Replace(sHeader, "<#dLastUpdate#>", ConvertDateForText(objRsExpWke("wkeEndDate"), " ", "MonthYear"))
Response.Write sHeader
Set fInTemplate=Nothing
Set objFSO=Nothing

' DOMAINES DE COMPETENCES
	WriteTableHeader
	WriteDataRowHeader GetLabel(sCvLanguage, "DOMAINES DE COMPETENCES")
	WriteGridDataRow 1, "100%", sKeyQualification, 0, "0"
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
		
		WriteGridDataRow 4, "20%#|#40%#|#30%#|#10%", "\b " & sPeriod & "\b0 #|#" & Trim(sEduDiploma & " " & objRsExpEdu("eduDiploma1Eng") & " " & sEduSubject) & "#|#" & objRsExpEdu("InstNameEng") & "#|#" & objRsExpEdu("InstLocationEng"), 0, "0#|#0#|#0#|#0"
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
		'If Len(ReplaceIfEmpty(objRsExpWke("wkePrjTitleEng"), ""))<=1 Then
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
			
			WriteGridDataRow 3, "20%#|#30%#|#50%", "\b " & Year(objRsExpWke("wkeStartDate")) & " - " & sExperienceEndDate & "\b0 #|#" &  objRsExpWke("wkeOrgNameEng") & "#|#" & objRsExpWke("wkePositionEng"), 0, "0#|#0#|#0"
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
		WriteGridDataRow 4, "25%#|#25%#|#25%#|#25%", "\cf15\b " & GetLabel(sCvLanguage, "Language") & "\b0 #|#\cf15\b " & GetLabel(sCvLanguage, "Reading") & "\b0\cf1 #|#\cf15\b " & GetLabel(sCvLanguage, "Writing") & "\b0\cf1 #|#\cf15\b " & GetLabel(sCvLanguage, "Speaking") & "\b0\cf1 ", 0, "0#|#0#|#0#|#0"
		While Not objRsExpLngNative.Eof
			On Error Resume Next
				sTempLanguage = ReplaceIfEmpty(objRsExpLngNative("lngName" & sCvLanguage), objRsExpLngNative("lngNameEng"))
			On Error GoTo 0
			WriteGridDataRow 4, "25%#|#25%#|#25%#|#25%", sTempLanguage & "#|#" & arrLanguageLevelTitle(objRsExpLngNative("exlReading")) & "#|#" & arrLanguageLevelTitle(objRsExpLngNative("exlWriting")) & "#|#" & arrLanguageLevelTitle(objRsExpLngNative("exlSpeaking")), 0, "0#|#0#|#0#|#0"
			objRsExpLngNative.MoveNext
		WEnd
		While Not objRsExpLngOther.Eof
			On Error Resume Next
				sTempLanguage = ReplaceIfEmpty(objRsExpLngOther("lngName" & sCvLanguage), objRsExpLngOther("lngNameEng"))
			On Error GoTo 0
			WriteGridDataRow 4, "25%#|#25%#|#25%#|#25%", sTempLanguage & "#|#" & arrLanguageLevelTitle(objRsExpLngOther("exlReading")) & "#|#" & arrLanguageLevelTitle(objRsExpLngOther("exlWriting")) & "#|#" & arrLanguageLevelTitle(objRsExpLngOther("exlSpeaking")), 0, "0#|#0#|#0#|#0"
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
			
			WriteGridDataRow 3, "20%#|#30%#|#50%", "\b " & sYear & "\b0 " & "#|#" &  "\cf15\b Maître d'Ouvrage\b0\cf1 " & "#|#" & "\b " & objRsExpWke("wkeOrgNameEng") & "\b0 ", 0, "0#|#0#|#0"
			WriteGridDataRow 3, "20%#|#30%#|#50%", "" & "#|#" &  "\cf15\b Projet\b0\cf1 " & "#|#" & objRsExpWke("wkePrjTitleEng"), 0, "0#|#0#|#0"
			WriteGridDataRow 3, "20%#|#30%#|#50%", "" & "#|#" &  "\cf15\b Fonction occupée\b0\cf1 " & "#|#" & objRsExpWke("wkePositionEng"), 0, "0#|#0#|#0"
			sDescription=ConvertText(objRsExpWke("wkeDescriptionEng"))
			WriteGridDataRow 3, "20%#|#30%#|#50%", "" & "#|#" &  "\cf15\b Descriptif\b0\cf1 " & "#|#" & sDescription, 0, "0#|#0#|#0"
			
			WriteTableFooter			
		End If
	objRsExpWke.MoveNext
	WEnd
	objRsExpWke.Close
	Set objRsExpWke=Nothing
	
	
	
	
Response.Write("\pard }" & vbCrLf)
%>


