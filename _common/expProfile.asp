<%
Dim iPersonID, sFirstName, sMiddleName, sLastName, iTitleID, sTitle, sBirthPlace, dBirthDate, iBirthDay, iBirthMonth, iBirthYear, iGenderID, iMaritalStatusID
Dim sFullName
Dim sProfession, sCurrPosition, sKeyQualifications, iExperienceYears, sNationality, sRegistrationNumber
Dim iProfessionalStatusID
Dim sAvailability, iShortterm, iLongterm, sPreferences
Dim arrExpNationalityID
Dim bEmailExpertAccountSent
Dim sCvFolder
Dim sCVFormat, bCvValidForMemberOrExpert

' Load expert's profile
Function LoadExpertProfile(AExpertId)
	If AExpertID>0 Then
		Set objTempRs=GetDataRecordsetSP("usp_ExpCvvExpInfoSelect", Array( _
			Array(, adInteger, , AExpertID)))
		If Not objTempRs.Eof Then
			On Error Resume Next
				iPersonID=objTempRs("id_Person")
				iTitleID=objTempRs("id_psnTitle")
				sFirstName=objTempRs("psnFirstNameEng")
				sMiddleName=objTempRs("psnMiddleNameEng")
				sLastName=objTempRs("psnLastNameEng")
				
				If iTitleID>0 Then sTitle=arrPersonTitle(iTitleID)
					
				If Len(sTitle)>0 Then
					sFullName=sTitle & " " & sLastName
				Else
					sFullName=sFirstName & " " & sLastName
				End If

				sBirthPlace=objTempRs("psnBirthPlaceEng")
				dBirthDate=objTempRs("psnBirthDate")
				If IsDate(dBirthDate) Then
					iBirthDay=Day(dBirthDate)
					iBirthMonth=Month(dBirthDate)
					iBirthYear=Year(dBirthDate)
				End If
				iGenderID=objTempRs("psnGender")
				iMaritalStatusID=objTempRs("id_MaritalStatus")

				sProfession=objTempRs("expProfessionEng")
				iProfessionalStatusID=objTempRs("id_ProfessionalStatus")
				sCurrPosition=objTempRs("expCurrPositionEng")
				sKeyQualifications=objTempRs("expKeyQualificationsEng")
				iExperienceYears=objTempRs("expProfYears")
				
				sAvailability=objTempRs("expAvailabilityEng")
				iShortterm=objTempRs("expShortterm")
				iLongterm=objTempRs("expLongterm")
				
				sUserPhone=objTempRs("Phone")
				sUserEmail=objTempRs("Email")
				
				sRegistrationNumber=objTempRs("expRegNumber")
				sPreferences=objTempRs("expPreferences")

				sCvLanguage = ReplaceIfEmpty(objTempRs("Lng"), sDefaultCvLanguage)
				sCvFolder=objTempRs("KgCvFile")
				bEmailExpertAccountSent = ReplaceIfEmpty(objTempRs("expAccountEmailSent"), 0)
			On Error GoTo 0
		End If 
		objTempRs.Close
		
		If sApplicationName="external" Then
			If sContactDetailsExternally=cNameObfuscated Then
				sFirstName=ObfuscateString(sFirstName)
				sLastName=ObfuscateString(sLastName)
				sMiddleName=ObfuscateString(sMiddleName)
				sUserEmail=ObfuscateEmail(sUserEmail)
				sUserPhone=ObfuscateString(sUserPhone)
				
			End If
			If sContactDetailsExternally=cNameHidden Then
				sFirstName=""
				sLastName=""
				sMiddleName=""
			End If
		End If
	ElseIf Request.Form()>"" Then
		iTitleID=CheckIntegerAndNull(Request.Form("exp_title"))
		sFirstName=Request.Form("exp_firstname")
		sMiddleName=Request.Form("exp_middlename")
		sLastName=Request.Form("exp_familyname")
		
		sBirthPlace=Request.Form("exp_birthplace")
		iBirthDay=CheckIntegerAndNull(Request.Form("exp_dbirth"))
		iBirthMonth=CheckIntegerAndNull(Request.Form("exp_mbirth"))
		iBirthYear=CheckIntegerAndNull(Request.Form("exp_ybirth"))

		iGenderID=Request.Form("exp_gender")
		iMaritalStatusID=CheckIntegerAndNull(Request.Form("exp_marstatus"))

		sProfession=Request.Form("exp_Prof")
		iProfessionalStatusID=CheckIntegerAndNull(Request.Form("exp_prof_status"))
		sCurrPosition=Request.Form("exp_curr_Position")
		sKeyQualifications=Request.Form("exp_key_qualif")
		iExperienceYears=Request.Form("exp_wke_years")
		
		sAvailability=Request.Form("Availability")
		iShortterm=Request.Form("shortterm")
		iLongterm=Request.Form("longterm")
		
		sUserPhone=Request.Form("exp_phone")
		sUserEmail=Request.Form("exp_email")
		
		sRegistrationNumber=Request.Form("exp_registration_number")
		sPreferences=Request.Form("preferences")

		sCvLanguage = ReplaceIfEmpty(Request.Form("exp_language"), sDefaultCvLanguage)

		' Load nationalities
		sNationality=Left(Request.Form("newloc"), 4000)
		arrExpNationalityID=Split(sNationality, ",")
	End If
End Function


Sub LoadExpertAccountDetails(AExpertId)
	If AExpertID>0 Then
		Set objTempRs=GetDataRecordsetSP("usp_ExpAccountDetailsSelect", Array( _
			Array(, adInteger, , AExpertID)))
		If Not objTempRs.Eof Then 
			If sApplicationName<>"expert" Then
				iUserID=objTempRs("id_User")
			End If 
			sUserLogin=objTempRs("UserName")
			sUserPassword=objTempRs("PassWord")
		End If	
	End If
End Sub
	

'  Load expert's nationalities
Function LoadExpertNationality(AExpertID)
	If AExpertID>0 Then
		Set objTempRs=GetDataRecordsetSP("usp_ExpCvvNationalitySelect", Array( _
			Array(, adInteger, , iExpertID)))
		ReDim arrExpNationalityID(1)
		i=1
		While Not objTempRs.Eof 
			ReDim preserve arrExpNationalityID(i)
			arrExpNationalityID(i)=objTempRs("id_Nationality")
			objTempRs.MoveNext
			i=i+1
		WEnd
		objTempRs.Close
	End If
End Function

' Save expert's quick profile
Function SaveExpertShortProfile(AExpertID, byref AFieldSet)
Dim objResult
	If sApplicationName<>"expert" Then
		iUserID=0
	End If

	If IsObject(AFieldSet) Then
		iPersonID=CheckInteger(AFieldSet("id_Person"))
		iTitleID=CheckInteger(AFieldSet("exp_title"))
		sCvLanguage=Left(ReplaceIfEmpty(CheckString(AFieldSet("exp_language")), ""), 3)
		sFirstName=CheckString(AFieldSet("exp_firstname"))
		sLastName=CheckString(AFieldSet("exp_familyname"))
		dBirthDate=ConvertDMYForSQL(CheckString(AFieldSet("exp_ybirth")), CheckString(AFieldSet("exp_mbirth")), CheckString(AFieldSet("exp_dbirth")))

		sUserPhone=Left(CheckString(AFieldSet("exp_phone")),50)
		sUserEmail=Trim(Left(CheckString(AFieldSet("exp_email")),120))
		' Take only the first email
		Dim iPosSpace, iPosSemicolon
		iPosSpace=InStr(sUserEmail, " ")
		iPosSemicolon=InStr(sUserEmail, ";")
		If iPosSemicolon>0 And (iPosSemicolon<iPosSpace Or iPosSpace=0) Then iPosSpace=iPosSemicolon
		If iPosSpace>0 And (iPosSpace<iPosSemicolon Or iPosSemicolon=0) Then iPosSemicolon=iPosSpace
		If iPosSpace>1 Then
			sUserEmail=Left(sUserEmail, iPosSpace-1)
		End If
		
		sAvailability=Left(CheckString(AFieldSet("Availability")),400)
		iShortterm=AFieldSet("shortterm")
		If Not IsNumeric(iShortterm) Then
			iShortterm=Null
		Else
			iShortterm=CInt(iShortterm)
		End If
		iLongterm=AFieldSet("longterm")
		If Not IsNumeric(iLongterm) Then
			iLongterm=Null
		Else
			iLongterm=CInt(iLongterm)
		End If

		objResult=GetDataOutParamsSP("usp_ExpertProfileShortUpdate", _
		Array( _
			Array(, adInteger, , AExpertID), _
			Array(, adInteger, , iUserID), _
			Array(, adVarChar, 3, sCvLanguage), _
			Array(, adTinyInt, , iTitleID), _
			Array(, adVarWChar, 255, sFirstName), _
			Array(, adVarWChar, 255, sLastName), _
			Array(, adVarChar,  16, dBirthDate), _
			Array(, adVarWChar, 255, sUserPhone), _
			Array(, adVarWChar, 255, sUserEmail), _
			Array(, adVarWChar, 4000, sAvailability), _
			Array(, adTinyInt, , iShortterm), _
			Array(, adTinyInt, , iLongterm)), _
		Array( _
			Array(, adInteger), _
			Array(, adInteger), _
			Array(, adInteger),_
			Array(, adVarChar, 255), _
			Array(, adVarChar, 255)))
	End If
SaveExpertShortProfile=objResult
End Function

' Save expert's full profile
Function SaveExpertFullProfile(AExpertID, byref AFieldSet)
Dim objResult, iTempExpertID
Set objResult=Nothing
	If sApplicationName<>"expert" Then
		iUserID=0
	End If

	On Error Resume Next
	If AFieldSet>"" Then
		iPersonID=CheckInteger(AFieldSet("id_Person"))
		sCvLanguage=Left(ReplaceIfEmpty(CheckString(AFieldSet("exp_language")), ""), 3)
		sFirstName=CheckString(AFieldSet("exp_firstname"))
		sMiddleName=CheckString(AFieldSet("exp_middlename"))
		sLastName=CheckString(AFieldSet("exp_familyname"))
		iTitleID=CheckInteger(AFieldSet("exp_title"))
		dBirthDate=ConvertDMYForSQL(CheckString(AFieldSet("exp_ybirth")), CheckString(AFieldSet("exp_mbirth")), CheckString(AFieldSet("exp_dbirth")))
		sBirthPlace=CheckString(AFieldSet("exp_birthplace"))
		iGenderID=CheckInteger(AFieldSet("exp_gender"))
		iMaritalStatusID=CheckInteger(AFieldSet("exp_marstatus"))

		sUserPhone=Left(CheckString(AFieldSet("exp_phone")),50)
		sUserEmail=Trim(Left(CheckString(AFieldSet("exp_email")),120))
		' Take only the first email
		Dim iPosSpace, iPosSemicolon
		iPosSpace=InStr(sUserEmail, " ")
		iPosSemicolon=InStr(sUserEmail, ";")
		If iPosSemicolon>0 And (iPosSemicolon<iPosSpace Or iPosSpace=0) Then iPosSpace=iPosSemicolon
		If iPosSpace>0 And (iPosSpace<iPosSemicolon Or iPosSemicolon=0) Then iPosSemicolon=iPosSpace
		If iPosSpace>1 Then
			sUserEmail=Left(sUserEmail, iPosSpace-1)
		End If
		
		sCurrPosition=CheckString(AFieldSet("exp_curr_Position"))
		sProfession=CheckString(AFieldSet("exp_prof"))
		iProfessionalStatusID=CheckInteger(AFieldSet("exp_prof_status"))
		
		sKeyQualifications=CheckString(AFieldSet("exp_key_qualif"))
		iExperienceYears=CheckInteger(AFieldSet("exp_wke_years"))

		sRegistrationNumber=CheckString(AFieldSet("exp_registration_number"))		
		
		objResult=GetDataOutParamsSP("usp_ExpertProfileFullUpdate", _
		Array( _
			Array(, adInteger, , AExpertID), _
			Array(, adInteger, , iUserID), _
			Array(, adVarChar, 3, sCvLanguage), _
			Array(, adTinyInt, , iTitleID), _
			Array(, adVarWChar, 255, sFirstName), _
			Array(, adVarWChar, 255, sMiddleName), _
			Array(, adVarWChar, 255, sLastName), _
			Array(, adVarChar,  16, dBirthDate), _
			Array(, adVarWChar, 255, sBirthPlace), _
			Array(, adTinyInt, , iGenderID), _
			Array(, adTinyInt, , iMaritalStatusID), _			
			Array(, adVarWChar, 255, sUserPhone), _
			Array(, adVarWChar, 255, sUserEmail), _
			Array(, adVarWChar, 80, sRegistrationNumber), _
			Array(, adInteger, , iProfessionalStatusID), _
			Array(, adVarWChar, 255, sProfession), _
			Array(, adVarWChar, 255, sCurrPosition), _
			Array(, adLongVarWChar, 25000, sKeyQualifications), _
			Array(, adInteger, , iExperienceYears)), _
		Array( _
			Array(, adInteger), _
			Array(, adInteger), _
			Array(, adInteger),_
			Array(, adVarChar, 255), _
			Array(, adVarChar, 255)))
		iTempExpertID = objResult(1)
				
		sNationality=Left(Request.Form("newloc"), 4000)
		objTempRs=UpdateRecordSP("usp_ExpertNationalityUpdate", Array( _
			Array(, adInteger, , iTempExpertID), _
			Array(, adVarChar, 4000, sNationality)))
			
	End If
	On Error GoTo 0
SaveExpertFullProfile=objResult
End Function

Function SaveExpertDocument(AExpertID, ADocumentType, byref AFieldSet)
Dim iResult
iResult=0
Dim objFile
Dim sUploadedFileName, sNewFileName, sFullFileName, sFileExtension
Dim sFieldName
sFieldName="exp_" & ADocumentType

	If AFieldSet(sFieldName).TotalBytes>0 Then
		sUploadedFileName=Trim(AFieldSet(sFieldName).UserFilename)
		sFileExtension=LCase(Right(sUploadedFileName,4))

		If Not (sFileExtension=".doc" or sFileExtension=".txt" or sFileExtension=".rtf" or sFileExtension=".pdf" or sFileExtension=".htm" or sFileExtension="html") Then
			iResult=1
        	ShowMessageStart "error", 580 %>
			Your CV file has unknown extension.</b><br>Please try to upload this file in MS Word document format. Thank you.<br>Click back and try again.
			<% ShowMessageEnd
			SaveExpertDocument=iResult
			Exit Function
		End If
	
		sNewFileName=ADocumentType & "_" & AExpertID & "_" & Replace(ConvertDateForSQL(Now),"/","") & "_" & Mid(sSessionID, 26, 9) & sFileExtension
		sFullFileName=Server.Mappath("/_upload" & sHomePath) & "\" & sNewFileName

		Set objFile=Server.CreateObject("Scripting.FileSystemObject")
		If objFile.FileExists(sFullFileName) Then
			If objFile.FileExists(sFullFileName & "_") Then    
				objFile.DeleteFile sFullFileName & "_"
			End If 
			objFile.MoveFile sFullFileName, sFullFileName & "_"
		End If    
		AFieldSet(sFieldName).SaveAs sFullFileName
		Set objFile=nothing
	End If
SaveExpertDocument=iResult
End Function
	

' Save expert's quick profile
Sub SaveExpertAccountEmailSent(AExpertID, AValue)
	Dim objTempRs
	objTempRs=UpdateRecordSP("usp_ExpertAccountEmailSentUpdate", Array( _
		Array(, adInteger, , AExpertID), _
		Array(, adInteger, , AValue)))
	Set objTempRs = Nothing
End Sub

	
' Save expert's quick profile
Function VerifyExpertProfile(byref AFieldSet)
Dim objResult
Set objResult=Nothing

	If IsObject(AFieldSet) Then
		sCvLanguage=Left(ReplaceIfEmpty(CheckString(AFieldSet("exp_language")), ""), 3)
		sFirstName=CheckString(AFieldSet("exp_firstname"))
		sLastName=CheckString(AFieldSet("exp_familyname"))
		dBirthDate=ConvertDMYForSQL(CheckString(AFieldSet("exp_ybirth")), CheckString(AFieldSet("exp_mbirth")), CheckString(AFieldSet("exp_dbirth")))

		sUserEmail=Left(CheckString(AFieldSet("exp_email")),120)

		Set objResult=GetDataRecordsetSP("usp_AdmExpExistsCheck", Array( _
			Array(, adVarWChar, 255, sFirstName), _
			Array(, adVarWChar, 255, sLastName), _
			Array(, adVarChar, 16, dBirthDate), _ 
			Array(, adVarChar, 255, sUserEmail)))
	End If
Set VerifyExpertProfile=objResult
End Function


Sub SaveExpertLanguage(AExpertID, ACvLanguage)
	Dim objTempRs
	objTempRs=UpdateRecordSP("usp_ExpertProfileLanguageUpdate", Array(_
		Array(, adInteger, , AExpertID), _
		Array(, adVarChar, 3, ACvLanguage)))
	Set objTempRs = Nothing
End Sub

Sub SaveExpertLanguageAndFolder(AExpertID, ACvLanguage, ACvFolder)
	Dim objTempRs
	objTempRs=UpdateRecordSP("usp_ExpertProfileCustomUpdate", Array(_
		Array(, adInteger, , AExpertID), _
		Array(, adVarChar, 3, ACvLanguage), _
		Array(, adVarWChar, 150, ACvFolder)))
	Set objTempRs = Nothing
End Sub

Sub SaveExpertCvLanguageLink(AInitialLanguageCvID, ANewLanguageCvID)
	Dim objTempRs
	objTempRs=UpdateRecordSP("usp_ExpertLanguageLinkUpdate", Array( _
		Array(, adInteger, , AInitialLanguageCvID), _
		Array(, adInteger, , ANewLanguageCvID)))
	Set objTempRs = Nothing
End Sub

Sub CopyExpertCvLanguage(AInitialLanguageCvID, ANewLanguageCvID)
	Dim objTempRs
	objTempRs=UpdateRecordSP("usp_ExpertCopyCvLanguage", Array( _
		Array(, adInteger, , AInitialLanguageCvID), _
		Array(, adInteger, , ANewLanguageCvID)))
	Set objTempRs = Nothing
End Sub
%>