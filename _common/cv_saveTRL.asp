<%
'--------------------------------------------------------------------
'
' Expert's CV. Save in Tripleline format
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
Set fInTemplate=objFso.OpenTextFile(Server.MapPath("\_common") & "\cv_trl.vrf", 1)
sHeader = fInTemplate.ReadAll & vbCrLf
sHeader = Replace(sHeader, "<#Name#>", sFullNameWithSpaces)
Response.Write sHeader
Set fInTemplate=Nothing
Set objFSO=Nothing

	' The initial table header is written already through file header section (WriteTableHeader)
	
	WriteDataRow "Position", "\b1\up1\cf2  \b0\up0\cf1"
	WriteDataRow "Profile", ConvertText2RTF(sKeyQualification)
	WriteDataRow "Key Skills", ""
	
' Nationality & languages
	Dim sLanguagesList
	sLanguagesList = ""
	If (Not objRsExpLngOther.Eof) Or (Not objRsExpLngNative.Eof) Then
		sLanguagesList = sLanguagesList & GetListStart()
		While Not objRsExpLngNative.Eof
			sLanguagesList = sLanguagesList & GetListItemStart() & objRsExpLngNative("lngNameEng") & " (native)" & GetListItemEnd()
			objRsExpLngNative.MoveNext
		WEnd

		While Not objRsExpLngOther.Eof
			sLanguagesList = sLanguagesList & GetListItemStart() & objRsExpLngOther("lngNameEng") & " (" & LCase(arrLanguageLevelTitle(objRsExpLngOther("exlAverage"))) & ")" & GetListItemEnd()
			objRsExpLngOther.MoveNext
		WEnd
		sLanguagesList = sLanguagesList & GetListEnd()
	End If
	objRsExpLngNative.Close
	Set objRsExpLngNative=Nothing	
	objRsExpLngOther.Close
	Set objRsExpLngOther=Nothing	

	WriteDataRow2 "Nationality", sNationality, "Languages", sLanguagesList
	
' Education	
	Dim sEducationList
	sEducationList = ""
	If Not objRsExpEdu.Eof Then
		sEducationList = sEducationList & GetListStart()
		While Not objRsExpEdu.Eof
			sEduSubject = ""
			If Len(objRsExpEdu("edsDescriptionEng"))>0 And objRsExpEdu("edsDescriptionEng")<>"Other" Then 
				sEduSubject=sEduSubject & Trim(objRsExpEdu("edsDescriptionEng"))
			End If
			If Len(objRsExpEdu("id_EduSubject1Eng"))>0 Then
				sEduSubject=sEduSubject & Trim(" " & objRsExpEdu("id_EduSubject1Eng"))
			End If
			If Len(sEduSubject)>0 Then
				sEduSubject = " in " & sEduSubject
			End If

			sEducationList = sEducationList & _
				GetListItemStart() & _
				Trim(EducationTypeTitleByID(objRsExpEdu("eduDiploma")) & _ 
				sEduSubject & _
				", " & _
				objRsExpEdu("InstLocationEng") & " " & _
				Year(objRsExpEdu("eduEndDate")))
				
			objRsExpEdu.MoveNext
			
			If Not objRsExpEdu.Eof Then
				sEducationList = sEducationList & GetListItemEnd()
			Else
				sEducationList = sEducationList & GetListItemEndLast()
			End If
		WEnd
		sEducationList = sEducationList & GetListEnd()
	End If
	objRsExpEdu.Close
	Set objRsExpEdu=Nothing	

	WriteDataRow "Qualifications", sEducationList
	
' Country Experience
	Dim sCountryList
	sCountryList = GetListStart() & GetExpertExperienceContinentCountryGroupedList2(iCvID, Null, sCvLanguage, GetListItemStart() & "\b1\up1\cf2 ", GetListItemEnd(), ": \b0\up0\cf1 ", "", ", ") & GetListEnd()
	
	WriteDataRow "Country Experience", sCountryList
	
	
' Clients
	Dim sClientList
	sClientList = GetExpertExperienceOrganisationList(iCvID, Null, Null)
	WriteDataRow "Clients", sClientList
	
' Employment History
	WriteSimpleRow "\b1\up1\cf2\fs28 Employment History\b0\up0\cf1"

	Dim sExperienceEndDate, sExperiencePeriod, sExperienceTitle
	If not objRsExpWke.Eof Then
	While Not objRsExpWke.Eof
		If IsDate(objRsExpWke("wkeStartDate")) And Not IsNull(objRsExpWke("wkeStartDate")) Then
			sExperiencePeriod = Year(objRsExpWke("wkeStartDate"))
		Else 
			sExperiencePeriod = ""
		End If
		If objRsExpWke("wkeEndDateOpen")=1 Then
			sExperienceEndDate = "present"
		ElseIf IsDate(objRsExpWke("wkeEndDate")) And Not IsNull(objRsExpWke("wkeEndDate")) Then
			sExperienceEndDate = Year(objRsExpWke("wkeEndDate"))
		Else 
			sExperienceEndDate = ""
		End If

		If Len(sExperiencePeriod)>0 And Len(sExperienceEndDate)>0 Then
			If sExperiencePeriod <> sExperienceEndDate Then
				sExperiencePeriod = sExperiencePeriod & "-" & sExperienceEndDate
			End If
		Else
			sExperiencePeriod = sExperiencePeriod & sExperienceEndDate
		End If

		sExperienceTitle = ""
		If Len(objRsExpWke("wkePositionEng"))>0 Then
			sExperienceTitle = sExperienceTitle & objRsExpWke("wkePositionEng") & ", "
		End If
		If Len(objRsExpWke("wkeOrgNameEng"))>0 Then
			sExperienceTitle = sExperienceTitle & objRsExpWke("wkeOrgNameEng") & ", "
		End If
		
		objTempRs2=GetDataOutParamsSP("usp_ExpCvvExperienceCouSelect", Array( _
			Array(, adInteger, , iCvID), Array(, adInteger, , objRsExpWke("id_ExpWke")), Array(, adVarChar, 10, "listshort")), Array( _
			Array(, adVarWChar, 4000)))
		sCountries=ConvertText2RTF(objTempRs2(0))
		Set objTempRs2=Nothing
		If Len(sCountries)>0 Then
			sExperienceTitle = sExperienceTitle & sCountries & ", "
		End If

		If Len(sExperienceTitle)>0 Then
			sExperienceTitle = Left(sExperienceTitle, Len(sExperienceTitle)-2)
		End If
		
		WriteDataRow sExperiencePeriod, sExperienceTitle
	objRsExpWke.MoveNext
	WEnd
	End If

	WriteTableFooterNoPar
	WriteTableHeader

' Selected Experience
	WriteSimpleRow "\b1\up1\cf2\fs28 Selected Experience\b0\up0\cf1"
	If Not objRsExpWke.Eof Then
	objRsExpWke.MoveFirst
	While Not objRsExpWke.Eof
		If IsDate(objRsExpWke("wkeStartDate")) And Not IsNull(objRsExpWke("wkeStartDate")) Then
			sExperiencePeriod = Year(objRsExpWke("wkeStartDate"))
		Else 
			sExperiencePeriod = ""
		End If
		If IsDate(objRsExpWke("wkeEndDate")) And Not IsNull(objRsExpWke("wkeEndDate")) Then
			sExperienceEndDate = Year(objRsExpWke("wkeEndDate"))
		Else 
			sExperienceEndDate = ""
		End If

		If Len(sExperiencePeriod)>0 And Len(sExperienceEndDate)>0 Then
			If sExperiencePeriod <> sExperienceEndDate Then
				sExperiencePeriod = sExperiencePeriod & "-" & sExperienceEndDate
			End If
		Else
			sExperiencePeriod = sExperiencePeriod & sExperienceEndDate
		End If
		objTempRs2=GetDataOutParamsSP("usp_ExpCvvExperienceCouSelect", Array( _
			Array(, adInteger, , iCvID), Array(, adInteger, , objRsExpWke("id_ExpWke")), Array(, adVarChar, 10, "listshort")), Array( _
			Array(, adVarWChar, 4000)))
		sCountries=ConvertText2RTF(objTempRs2(0))
		Set objTempRs2=Nothing
		If Len(sCountries)>0 Then
			sExperiencePeriod = sExperiencePeriod & "\line\b0\up0\cf1 " & sCountries
		End If

		sExperienceTitle = ""
		If Len(objRsExpWke("wkePrjTitleEng"))>0 Then
			sExperienceTitle = sExperienceTitle & objRsExpWke("wkePrjTitleEng") & ", "
		End If
		If Len(objRsExpWke("wkeOrgNameEng"))>0 Then
			sExperienceTitle = sExperienceTitle & objRsExpWke("wkeOrgNameEng") & ", "
		End If
		If Len(objRsExpWke("wkePositionEng"))>0 Then
			sExperienceTitle = sExperienceTitle & objRsExpWke("wkePositionEng") & ", "
		End If
		
		If Len(sExperienceTitle)>0 Then
			sExperienceTitle = Left(sExperienceTitle, Len(sExperienceTitle)-2)
		End If
		
		sDescription=ConvertText2RTF(objRsExpWke("wkeDescriptionEng"))
		
		sExperienceTitle = "\f6\fs22\b1\up1\cf2 " & sExperienceTitle & "\line\b0\up0\cf1 " & sDescription
		
		WriteDataRow sExperiencePeriod, sExperienceTitle
	objRsExpWke.MoveNext
	WEnd
	End If
	objRsExpWke.Close
	Set objRsExpWke=Nothing

	
' Publications
	If sPublications>"" Then
		WriteSimpleRow "\b1\up1\cf2\fs28 Publications\b0\up0\cf1"
		WriteSimpleRow "\b0\up0\cf1 " & ConvertText2RTF(sPublications)
	End If


' Other relevant information
	WriteSimpleRow "\b1\up1\cf2\fs28 Other relevant information\b0\up0\cf1"
	If sMemberships>"" Then
		WriteSimpleRow "Membership:\b0\up0\cf1  " & ConvertText2RTF(sMemberships)
	End If
	If sOtherSkills>"" Then
		WriteSimpleRow "Other skills:\b0\up0\cf1  " & ConvertText2RTF(sOtherSkills)
	End If

	WriteTableFooterNoPar

Response.Write("\pard }}" & vbCrLf)
%>


