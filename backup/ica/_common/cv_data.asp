<%
If sApplicationName <> "expert" Then
	CheckUserLogin sScriptFullNameAsParams
End If
%>
<!--#include file="_class/document.asp"-->
<!--#include file="_grid/document_list.asp"-->
<!--#include file="keywords_highlight.asp"-->
<%
Dim iCvID, sCvUID, bCvValidForMemberOrExpert
Dim sCVFormat, dflag

'sCvUID = ReplaceIfEmpty(Request.QueryString("uid"), Request.Form("uid"))
sCvUID = Request("uid")
' if a Top Expert is operating - get the UID from the account, to not pass it in the URL:
If bIsMyCV And Len(uUserTopExpertUid) > 5 Then
	sCvUID = uUserTopExpertUid
End If

On Error Resume Next
	Set objTempRs=GetDataRecordsetSP("usp_Ica_ExpertIdSelect", Array( _
		Array(, adVarChar, 40, sCvUID)))

	If Err.Number <> 0 Or objTempRs.Eof Then
		' Redirect top experts to a different page, if their CV ID is not found
		If Len(uUserTopExpertUid) > 5 Then
			Response.Redirect "/backoffice/mycv/notfound.asp"
		Else
			Response.Redirect "/"
		End If
	End If

	iCvID = objTempRs("id_Expert")
	Set objExpertDB = objExpertDBList.Find(objTempRs("id_Database"), "ID")
	
Set objTempRs=Nothing
On Error GoTo 0

Dim objConnCustom
Set objConnCustom = Server.CreateObject("ADODB.Connection")
objConnCustom.Open "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Data Source=MAGNESA;Initial Catalog=" & objExpertDB.DatabasePath & ";"

' Redirect to original CV if expert's CV was updated with another ID (Blacklist=1 & id_ExpertOriginal>0)
objTempRs=GetDataOutParamsSPWithConn(objConnCustom, "usp_ExpCvvOriginalSelect", Array( _
	Array(, adInteger, , iCvID)), _
	Array( Array(, adInteger)))
iExpertOriginalID=objTempRs(0)
If iExpertOriginalID > 0 Then Response.Redirect(Replace(sScriptFileName, "cv_preview", "cv_view") & ReplaceUrlParams(sParams, "id=" & iExpertOriginalID))

' Check is the CV valid for member / expert
bCvValidForMemberOrExpert = IsIcaUserCompanyCvValid(objExpertDB.Database, iCvID, objUserCompanyDB.Database)

' Save log on CV views
LogUserCvView iUserID, objExpertDB.Id, iCvID, bCvValidForMemberOrExpert

'! For affiliate members and members without EDB option show no access page
' And (iMemberAccessExperts = cMemberAccessExpertsOwnOnly) _
If Not (bCvValidForMemberOrExpert = 1 Or bCvValidForMemberOrExpert = 5) _
And ( _
	iUserAccessExpertsOtherDbNoAccess = 1 _
	Or iUserAccessExpertsOtherDbRestricted = 1 _
	Or iMemberAccessExperts = cMemberAccessExpertsOwnOnly _
) _ 
And sScriptFileName <> "cv_noaccess.asp" Then
	If bAssortisSubscriptionEdbActive And objExpertDB.Database = "assortis" Then 
		' Response.Write objExpertDB.Database
	Else
		Response.Redirect "cv_noaccess.asp?uid=" & sCvUID
	End If
End If

'!!!! Redirect to preview if no access rights on the CV
If Not (bCvValidForMemberOrExpert = 1 Or bCvValidForMemberOrExpert = 5) _
And sScriptFileName <> "cv_verify.asp" _
And sScriptFileName <> "cv_copy.asp" _
And sScriptFileName <> "cv_noaccess.asp" _
Then
	If Not InStr(1, sScriptFileName, "cv_preview.asp", vbTextCompare) > 0 Then
		Response.Redirect(Replace(sScriptFullName, sScriptFileName, "cv_preview.asp"))
	End If
End If

' On changing a CV format do redirect
sCvFormat = Request.QueryString("act")
If sCvFormat = "ASR" Or sCvFormat = "ADB" Or sCvFormat = "AFB" Or sCvFormat = "EC" Or sCvFormat = "EP" Or sCvFormat = "WB" Or sCvFormat = "MCC" Or sCvFormat = "USAID" Or sCvFormat = "USA" Then
	If sCvFormat = "ASR" Then sCvFormat = ""
	sUrl = "cv_view" & sCvFormat & ".asp" & ReplaceUrlParams(sParams, "act")
	Response.Redirect(sUrl)
End If

If InStr(sScriptFileName, "adb")>0 Then
	sCvFormat = "ADB"
ElseIf InStr(sScriptFileName, "afb")>0 Then
	sCvFormat = "AFB"
ElseIf InStr(sScriptFileName, "ec")>0 Then
	sCvFormat = "EC"
ElseIf InStr(sScriptFileName, "ep")>0 Then
	sCvFormat = "EP"
ElseIf InStr(sScriptFileName, "wb")>0 Then
	sCvFormat = "WB"
ElseIf InStr(sScriptFileName, "mcc")>0 Then
	sCvFormat = "MCC"
ElseIf InStr(sScriptFileName, "usaid")>0 Then
	sCvFormat = "USAID"
Else
	sCvFormat = ""
End If

Dim iProjectID
iProjectID=CheckIntegerAndZero(Request.QueryString("idproject"))
If iProjectID>0 Then
	sParams=ReplaceUrlParams(sParams, "idproject=" & iProjectID)
End If
sParams=ReplaceUrlParams(sParams, "t")



'Dim sActiveLng, sMasterLng, k, arrRowsValues
Dim sFileType, sFileName, k, arrRowsValues

'sActiveLng="Eng"

Dim sFirstNameEng, sFirstNameFra, sFirstNameSpa, sFirstName
Dim sMiddleNameEng, sMiddleNameFra, sMiddleNameSpa, sMiddleName
Dim sLastNameEng, sLastNameFra, sLastNameSpa, sLastName, sTempLastName
Dim sPhone, sEmail
Dim iTitleID, sTitle, sFullName, sTitleLastName, sFullNameWithSpaces, iGender, sBirthDate, sBirthPlace, iMaritalStatus, iPersonID
Dim sNationality, sOtherLanguages, sTempLanguage
Dim sDonors, sCountries, sSectors, sProcurementTypes
Dim sPermAddress, sPermAddressStreet, sPermAddressPostcode, sPermAddressCity, sPermAddressCountry, sPermAddressPhone, sPermAddressMobile, sPermAddressFax, sPermAddressEmail, sPermAddressWeb, bPermAddress
Dim sCurAddress, sCurAddressStreet, sCurAddressPostcode, sCurAddressCity, sCurAddressCountry, sCurAddressPhone, sCurAddressMobile, sCurAddressFax, sCurAddressEmail, sCurAddressWeb, bCurAddress
Dim objRsExpEdu, objRsExpTrn, objRsExpWke, objRsExpLng, objRsExpLngOther, objRsExpLngNative, objRsExpCou
Dim sEduSubject, sDescription
Dim sComments

Dim iProfYears, sProfession, sMemberships, sOtherSkills, sKeyQualification, sPosition, sPublications
Dim iProfessionalStatusID, sProfessionalStatus
Dim sReferences, sAvailability, bShortterm, bLongterm, sAchievements
Dim sUserPhone, sPreferences

Dim arrStartDateValues, arrEndDateValues, arrPrjTitleValues
Dim bEmailToConfirmCVSent, bEmailToCompleteCVSent, bCVApproved, bCVIbfOnly, bCVHidden, bCVDeleted, bCVRemoved
Dim bEmailExpertAccountSent

Dim bCvAccessValid
bCvAccessValid = 0

If iCvID > 0 Then

Set objTempRs=GetDataRecordsetSPWithConn(objConnCustom, "usp_ExpertSelect", Array( _
	Array(, adInteger, , objExpertDB.ID), _
	Array(, adInteger, , iCvID)))

	If sScriptFileName = "cv_restore.asp" Then
		' Response.Write "OK"
		' Response.End
	ElseIf objTempRs.Eof Then
	' Redirect to the homepage if expert marked as deleted / removed or doesn't exist in the database
		Response.Redirect "/"
		Response.End
	End If
	
If Not objTempRs.Eof Then
	sCvLanguage = ReplaceIfEmpty(objTempRs("Lng"), sDefaultCvLanguage)
End If
%>
<!--#include file="_data/datGender.asp"-->
<!--#include file="_data/datLanguage.asp"-->
<!--#include file="_data/datLngLevel.asp"-->
<!--#include file="_data/datPsnTitle.asp"-->
<!--#include file="_data/datPsnStatus.asp"-->
<!--#include file="_data/datEduSubject.asp"-->
<!--#include file="_data/datEduType.asp"-->
<!--#include file="_data/datMonth.asp"-->
<!--#include file="_data/en/datProfessionalStatus.asp"-->
<%
If Not objTempRs.Eof Then
	On Error Resume Next
	sCvLanguage = ReplaceIfEmpty(objTempRs("Lng"), sDefaultCvLanguage)
	ForceCvLanguage()

	sSearchKeywordsHighlight=Request.QueryString("txt")
	Dim sSearchKeywordsAdd
	sSearchKeywordsAdd=Request.QueryString("srch_queryadd")
	
	iProfYears=objTempRs("expProfYears")
	sProfession=objTempRs("expProfessionEng")
	iProfessionalStatusID=objTempRs("id_ProfessionalStatus")

	If IsArray(arrProfessionalStatusID) Then
	For i=LBound(arrProfessionalStatusID) to UBound(arrProfessionalStatusID) 
		If iProfessionalStatusID=arrProfessionalStatusID(i) Then
			sProfessionalStatus=arrProfessionalStatusTitle(i)
		End If
	Next 
	End If
	
	sOtherSkills = ConvertText(objTempRs("expOtherSkills"))
	sMemberships = ConvertText(objTempRs("expMemberProfEng"))
	sKeyQualification = ConvertText(objTempRs("expKeyQualificationsEng"))
	sPosition = objTempRs("expCurrPositionEng")
	sPublications = ConvertText(objTempRs("expPublicationsEng"))
	sReferences = ConvertText(objTempRs("expReferencesEng"))
	sAvailability = ConvertText(objTempRs("expAvailabilityEng"))
	bShortterm = objTempRs("expShortterm")
	bLongterm = objTempRs("expLongterm")

	sUserLanguage=objTempRs("Lng")
	sEmail=objTempRs("Email")
	sPhone=objTempRs("Phone")
	bEmailToConfirmCVSent=objTempRs("expToConfirmCvEmailSent")
	bEmailToCompleteCVSent=objTempRs("expToCompleteCvEmailSent")
	bCVApproved=objTempRs("expApproved")
	bCVHidden=objTempRs("expHidden")
	bCVDeleted=objTempRs("expDeleted")
	bCVRemoved=objTempRs("expRemoved")

	sPreferences=""
	If bShortterm Then 
		sPreferences=sPreferences & "Short-term missions, "
	End If
	If bLongterm Then 
		sPreferences=sPreferences & "Long-term missions, "
	End If
	If Len(sPreferences)>2 Then
		sPreferences=Left(sPreferences,Len(sPreferences)-2)
	End If
	
	sFirstNameEng=objTempRs("psnFirstNameEng")
	sFirstName=sFirstNameEng
	If sFirstName>"" Then sFirstName=sFirstName & " "

	sMiddleNameEng=objTempRs("psnMiddleNameEng")
	sMiddleName=sMiddleNameEng
	If sMiddleName>"" Then sMiddleName=sMiddleName & " "

	sLastNameEng=objTempRs("psnLastNameEng")
	sLastName=sLastNameEng
	If sLastName>"" Then sLastName=sLastName & " "

	If bCvValidForMemberOrExpert = aClientSecurityCvViewEnabled Or bCvValidForMemberOrExpert = aClientSecurityCvViewAll Then
		iTitleID=objTempRs("id_psnTitle")
		If iTitleID>"" And IsNumeric(iTitleID) Then sTitle=arrPersonTitle(iTitleID) & " "

		sFullName=Trim(sTitle & " " & sFirstName & " " & sMiddleName & " " & sLastName)
		sTitleLastName=Trim(sTitle & " " & sLastName)
		If Len(sFullName) > 80 Then
			sFullName=sTitle & " " & sFirstName & " " & sLastName
		End If		
		If Len(sFullName) < 46 Then
			sFullNameWithSpaces=sFullName+Space(46-Len(sFullName))
		Else
			sFullNameWithSpaces=sFullName+Space(5)
		End If		
		If InStr(sLastName, "&")>0 Then 
			sTempLastName=Left(sLastName, InStr(sLastName, "&")-1) 
		Else 
			sTempLastName=sLastName
		End If
		sFileName="CV_" & Replace(sTempLastName," ","") & "_" & Replace(sFirstName," ","") & "_" & sCvFormat & ConvertDateForText(Now(), "", "DDMMYYYY")
	End If

	sBirthPlace = objTempRs("psnBirthPlaceEng")
	sBirthDate = objTempRs("psnBirthDate")
	iGender=objTempRs("psnGender")
	iMaritalStatus=objTempRs("id_MaritalStatus")

	sComments=objTempRs("expComments")
	
	iPersonID=objTempRs("id_Person")
objTempRs.Close 
On Error GoTo 0
End If

sNationality = GetExpNationalities(sCvLanguage, objExpertDB.Database, iCvID)

Set objRsExpEdu=GetDataRecordsetSPWithConn(objConnCustom, "usp_ExpertEducationSelect", Array( _
	Array(, adInteger, , iCvID), _
	Array(, adInteger, , 1)))

Set objRsExpTrn=GetDataRecordsetSPWithConn(objConnCustom, "usp_ExpertEducationSelect", Array( _
	Array(, adInteger, , iCvID), _
	Array(, adInteger, , 2)))

Set objRsExpWke=GetDataRecordsetSPWithConn(objConnCustom, "usp_ExpCvvExperienceSelect", Array( _
	Array(, adInteger, , iCvID)))


objRsExpLng=GetDataOutParamsSPWithConn(objConnCustom, "usp_GetExpertProfDetails", Array( _
	Array(, adInteger, , iCvID), Array(, adInteger, , 0), Array(, adVarChar, 3, "Eng"), Array(, adInteger, , 0)), Array( _
	Array(, adVarChar, 1), Array(, adVarWChar, 400), Array(, adVarWChar, 50), Array(, adVarWChar, 255), Array(, adVarWChar, 1000), Array(, adVarWChar, 500), Array(, adVarWChar, 1000), Array(, adVarWChar, 500), Array(, adVarWChar, 1000), Array(, adVarWChar, 2000)))
sOtherLanguages=objRsExpLng(5)
Set objRsExpLng=Nothing


Set objRsExpLngNative=GetDataRecordsetSPWithConn(objConnCustom, "usp_ExpCvvLanguageSelect", Array( _
	Array(, adInteger, , iCvID), _
	Array(, adVarChar, 10, "native")))

Set objRsExpLngOther=GetDataRecordsetSPWithConn(objConnCustom, "usp_ExpCvvLanguageSelect", Array( _
	Array(, adInteger, , iCvID), _
	Array(, adVarChar, 10, "other")))


' Get DBs where expert is registered as Top Expert
Set objExpertTopExpertDBList = New CCompanyExpertDBList
objExpertTopExpertDBList.LoadData "usp_Ica_ExpertTopExpertDBSelect", Array( _
		Array(, adVarChar, 50, objExpertDB.Database),_
		Array(, adInteger, , iCvID))

' Get alternative CV owners
Set objExpertDBOtherList = New CCompanyExpertDBList
objExpertDBOtherList.LoadData "usp_Ica_ExpertDBOwnerOtherSelect", Array( _
		Array(, adVarChar, 50, objExpertDB.Database),_
		Array(, adInteger, ,iCvID))

Dim iExpertDBOtherLoop

' For top experts the primary ownership is overwritten by top expert db
If sUserIpAddress = "158.29.157.31" Then
If objExpertTopExpertDBList.Count > 0 Then
	Dim objExpertOriginalDB
	Set objExpertOriginalDB = objExpertDB


	Dim bUserStaffExpertDBOwnerPrimary, _
	bUserStaffExpertDBOwnerAny, _
	bUserStaffExpertDBOwnerTE

	bUserStaffExpertDBOwnerPrimary = 0
	bUserStaffExpertDBOwnerAny = 0
	bUserStaffExpertDBOwnerTE = 0

	If iUserCompanyID = objExpertDB.Company.ID Then
		bUserStaffExpertDBOwnerPrimary = 1
		bUserStaffExpertDBOwnerAny = 1
	End If
	If (bCvValidForMemberOrExpert = 1 Or bCvValidForMemberOrExpert = 5) Then
		bUserStaffExpertDBOwnerAny = 1
	End If
	If bUserStaffExpertDBOwnerAny = 0 Then
		For iExpertDBOtherLoop = 0 To objExpertDBOtherList.Count - 1
			If iUserCompanyID = objExpertTopExpertDBList.Item(0).Company.ID Then
				bUserStaffExpertDBOwnerAny = 1
			End If
		Next
	End If
	If iUserCompanyID = objExpertTopExpertDBList.Item(0).Company.ID Then
		bUserStaffExpertDBOwnerTE = 1
	End If

	Response.Write "bUserStaffExpertDBOwnerPrimary=" & bUserStaffExpertDBOwnerPrimary
	Response.Write "bUserStaffExpertDBOwnerAny=" & bUserStaffExpertDBOwnerAny
	Response.Write "bUserStaffExpertDBOwnerTE=" & bUserStaffExpertDBOwnerTE


	' Overwrite primary ownership by top expert db
	If objExpertDB.ID <> objExpertTopExpertDBList.Item(0).ID Then
		Set objExpertDB = objExpertTopExpertDBList.Item(0)
		objExpertDB.DatabaseCodePrimary = objExpertOriginalDB.DatabaseCodePrimary
		objExpertDBOtherList.ReplaceDB objExpertTopExpertDBList.Item(0), objExpertOriginalDB
	End If

	' Cases 1, 3, 5
	If bUserStaffExpertDBOwnerTE = 0 _
	And bUserStaffExpertDBOwnerAny = 0 _
	And bUserIcaStaff = 0 _
	Then
		objExpertDBOtherList.ClearAll
	End If

	' Cases 2, 4
	If objExpertOriginalDB.Database = "assortis" Then
		' Case 2
		If objExpertTopExpertDBList.Item(0).Database = "ibf" Then
			If (bUserIbfStaff = 1 Or bUserIcaStaff = 1) Then
			ElseIf bUserStaffExpertDBOwnerAny = 1 Then
				'Dim objExpertCompanyDB
				'Set objExpertCompanyDB = New CCompanyExpertDB
				'objExpertCompanyDB = LoadCompanyDatabase(iUserCompanyID, Null, Null)

				'objExpertDBOtherList.ClearAll
				'objExpertDBOtherList.AddDB(objExpertCompanyDB)

			ElseIf bUserStaffExpertDBOwnerAny = 0 Then
				objExpertDBOtherList.ClearAll
			End If
		Else
		' Case 4
			If (bUserIbfStaff = 1 Or bUserIcaStaff = 1) Then

				Dim objExpertIbfAssortisDB
				Set objExpertIbfAssortisDB = New CCompanyExpertDB
				objExpertIbfAssortisDB.LoadData "usp_CompanyDatabaseSelect", Array( _
					Array(, adInteger, , Null), _
					Array(, adInteger, , 100), _
					Array(, adVarChar, 50, Null),_
					Array(, adInteger, , 0), _
					Array(, adInteger, , 1))

				'objExpertDBOtherList.AddDB(objExpertIbfAssortisDB)

				objExpertDBOtherList.ReplaceDB objExpertAssortisDB, objExpertIbfAssortisDB
			ElseIf bUserStaffExpertDBOwnerAny = 0 Then
				objExpertDBOtherList.ClearAll
			End If
		End If
	End If
End If
End If



'If objExpertTopExpertDBList.Count > 0 Then
'	Dim objExpertOriginalDB
'	Set objExpertOriginalDB = objExpertDB
'
'	If objExpertDB.ID <> objExpertTopExpertDBList.Item(0).ID Then
'		Set objExpertDB = objExpertTopExpertDBList.Item(0)
'
'		If objExpertOriginalDB.Database = "assortis" _
'		And objExpertTopExpertDBList.Item(0).Database <> "ibf" _
'		And bUserIbfStaff = 1 Then
'			objExpertDBOtherList.ReplaceDB objExpertTopExpertDBList.Item(0), objExpertOriginalDB
'		Else
'			objExpertDBOtherList.ClearAll
'		End If
'
'	End If
'End If




If bCvValidForMemberOrExpert = aClientSecurityCvViewEnabled Or bCvValidForMemberOrExpert = aClientSecurityCvViewAll Then

' Current address
bCurAddress=False
Set objTempRs=GetDataRecordsetSPWithConn(objConnCustom, "usp_ExpCvvAddressSelect", Array( _
	Array(, adInteger, , iCvID), _
	Array(, adInteger, , 3)))
	
If Not objTempRs.Eof Then 
	sCurAddressStreet=objTempRs("adrStreetEng")
	sCurAddressPostcode=objTempRs("adrPostCode")
	sCurAddressCity=objTempRs("adrCityEng")
	sCurAddressCountry=objTempRs("couNameEng")
	sCurAddressPhone=objTempRs("adrPhone")
	sCurAddressMobile=objTempRs("adrMobile")
	sCurAddressFax=objTempRs("adrFax")
	sCurAddressEmail=objTempRs("adrEmail")
	sCurAddressWeb=objTempRs("adrWeb")

	If CheckLength(sCurAddressStreet) + CheckLength(sCurAddressPostcode) + CheckLength(sCurAddressCity) + CheckLength(sCurAddressCountry) + CheckLength(sCurAddressPhone) + CheckLength(sCurAddressMobile) + CheckLength(sCurAddressFax) + CheckLength(sCurAddressEmail) + CheckLength(sCurAddressWeb)>0 Then
		bCurAddress=True
	End If
End If
objTempRs.Close

' Permanent address
sPermAddress=""
bPermAddress=False
Set objTempRs=GetDataRecordsetSPWithConn(objConnCustom, "usp_ExpCvvAddressSelect", Array( _
	Array(, adInteger, , iCvID), _
	Array(, adInteger, , 1)))
If Not objTempRs.Eof Then 
	sPermAddressStreet=objTempRs("adrStreetEng")
	sPermAddressPostcode=objTempRs("adrPostCode")
	sPermAddressCity=objTempRs("adrCityEng")
	sPermAddressCountry=objTempRs("couNameEng")
	sPermAddressPhone=objTempRs("adrPhone")
	sPermAddressMobile=objTempRs("adrMobile")
	sPermAddressFax=objTempRs("adrFax")
	sPermAddressEmail=objTempRs("adrEmail")
	sPermAddressWeb=objTempRs("adrWeb")

	If CheckLength(sPermAddressStreet) + CheckLength(sPermAddressPostcode) + CheckLength(sPermAddressCity) + CheckLength(sPermAddressCountry) + CheckLength(sPermAddressPhone) + CheckLength(sPermAddressMobile) + CheckLength(sPermAddressFax) + CheckLength(sPermAddressEmail) + CheckLength(sPermAddressWeb)>0 Then
		bPermAddress=True
	End If

	sPermAddress="<p class=""txt"">" & sPermAddressStreet & " " & sPermAddressPostcode & " " & sPermAddressCity & "</p>" 
	If sPermAddressCountry>"" Then
		sPermAddress=sPermAddress & "<p class=""txt"">" & sPermAddressCountry & "</p>"
	End If
	sPermAddressPhone=ReplaceIfEmpty(sPermAddressPhone, sCurAddressPhone)
	If sPermAddressPhone>"" Then
		sPermAddress=sPermAddress & "<p class=""txt"">Phone: " & sPermAddressPhone & "</p>"
	End If
	If sPermAddressFax>"" Then
		sPermAddress=sPermAddress & "<p class=""txt"">Fax: " & sPermAddressFax & "</p>"
	End If
	sPermAddressEmail=ReplaceIfEmpty(sPermAddressEmail, sCurAddressEmail)
	If sPermAddressEmail>"" Then
		sPermAddress=sPermAddress & "<p class=""txt"">E-mail: " & sPermAddressEmail & "</p>"
	End If
End If
objTempRs.Close


If CheckLength(sCurAddressEmail)=0 Then
	If (Not (InStr(sPermAddressEmail, sEmail)>0 Or InStr(sEmail, sPermAddressEmail)>0)) Or CheckLength(sPermAddressEmail)=0 Then
		sCurAddressEmail=sEmail
	End If
	If Len(sCurAddressPhone)<5 And (Not (InStr(sPermAddressPhone, sUserPhone)>0 Or InStr(sUserPhone, sPermAddressPhone)>0)) Or CheckLength(sPermAddressPhone)=0 Then
		sCurAddressPhone=sUserPhone
	End If
End If

If bCurAddress=False And bPermAddress=False And (CheckLength(sEmail)>0 Or CheckLength(sUserPhone)>0) Then
	bCurAddress=True
End If


End If
End If

Function SetECLanguageLevel(iLevel)
Dim iECLevel
  If iLevel>"" And IsNumeric(iLevel) Then
	iECLevel=iLevel
  Else
	iECLevel=""
  End If
SetECLanguageLevel=iECLevel
End Function

Function SetEPLanguageLevel(iLevel)
Dim sEPLevel
  If iLevel>"" And IsNumeric(iLevel) Then
	Select Case iLevel
		Case 5
		sEPLevel="C2"
		Case 4
		sEPLevel="C1"
		Case 3
		sEPLevel="B2"
		Case 2
		sEPLevel="B1"
		Case 1
		sEPLevel="A2"
		Case else
		sEPLevel=""
	End Select  
  Else
	sEPLevel=""
  End If
SetEPLanguageLevel=sEPLevel
End Function


Sub ShowExpCVFeatureBox

	' Verify if expert is already registered in my experts circle
	Dim bIsCompanyCircleExpert
	bIsCompanyCircleExpert = GetExpertCompanyCircleByUid(sCvUID, iUserCompanyID, iUserID)

	Dim iAddedByIcaUserId
	iAddedByIcaUserId = GetAddedByUserId(sCvUID, iUserCompanyID)

	' Verify if expert is already registered in top experts
	Dim iTopExpertStatus
	iTopExpertStatus = GetExpertCompanyTopExpertByUid(sCvUID, iUserCompanyID, iUserID)

	If (bCvValidForMemberOrExpert = 1 Or bCvValidForMemberOrExpert = 5) Then
		ShowFeatureBoxHeader("Expert selection") %>
		<div class="content">
		<% 
		If bIsCompanyCircleExpert Then 
			%><p class="sml"><img src="/image/file_circle.gif" width=18 height=17 border=0 hspace=3 align="left"><strong>My Experts Circle</strong>
			<%
			If (iUserCompanyRoleID = cUserRoleCompanyAdministrator Or _
				iUserCompanyRoleID = cUserRoleGlobalAdministrator Or _
				iUserCompanyRoleID = cUserRoleCVContactPoint Or _
				iUserID = iAddedByIcaUserId) _ 
			Then
				%><br><small><a class="list" href="/backoffice/view/cv_circle_fields.asp?uid=<%=sCvUID%>">Update fields selection</a></small>
				<%
			End If 
			%>
			</p>
			<% If iTopExpertStatus = 1 Then %>
				<p class="sml"><img src="/image/file_top.gif" width=18 height=17 border=0 hspace=3 align="left"><strong>Top Expert</strong></p>
			<% 
			ElseIf iTopExpertStatus = 2 Then %>
				<p class="sml"><img src="/image/file_top.gif" width=18 height=17 border=0 hspace=3 align="left"><strong>Request is pending</strong></p>
			<% 
			' Top experts could be added only if the CV is primary own by the company
			Else
				If objExpertDB.Database = objUserCompanyDB.Database Then %>
					<p class="sml"><a class="list" href="<% =sIcaServerProtocol & sIcaServer %>/Intranet/Dashboard?act=terequest&val=<%=sCvUID%>"><img src="/image/file_top.gif" width=18 height=17 border=0 hspace=3 align="left">Add to Top Experts</a></p>
				<% Else
					' ? message: you cannot add this expert to the Top experts because another company primary own the CV
				End If
			End If %>
		<% Else %>
			<p class="sml"><a class="list" href="javascript:void(0)" onclick="if (confirm('Are you sure you want to add this expert to your circle?')) { location.href = '/backoffice/view/cv_circle_fields.asp?uid=<%=sCvUID%>'; }"><img src="/image/file_circle.gif" width=18 height=17 border=0 hspace=3 align="left">Add to My Experts Circle</a></p>
		<% End If %>
		</div>
		<% ShowFeatureBoxFooter %>
		<br />
	<%
	End If
	%>

	<%
	If (bCvValidForMemberOrExpert=1 Or bCvValidForMemberOrExpert=5) And InStr(sScriptFileName, "cv_view")>0 Then
	%>

		<%
		Dim objCVLanguageList
		Set objCVLanguageList = new CExpertCvLanguageList
		objCVLanguageList.LoadData "usp_Ica_ExpertCvLanguageListSelect", Array( _
			Array(, adInteger, , objExpertDB.ID), _
			Array(, adInteger, , iCvID))

		If objCVLanguageList.Count > 1 Then %>

			<form name="cvlanguage" method="get" action="<%=sApplicationHomePath & "view/cv_view.asp" %>">
			<input type="hidden" name="idproject" value="<% =iProjectID %>">
			<input type="hidden" name="txt" value="<% =sSearchKeywordsHighlight %>">
			<input type="hidden" name="srch_queryadd" value="<% =sSearchKeywordsAdd %>">
			

			<% ShowFeatureBoxHeader "CV in other language" %>
			<div class="content">
			<div align="center">
			<select name="uid" id="uid" size="1"  style="font-face: Arial; font-size:8.5pt;">
			<% objCVLanguageList.ShowSelectItems sCvUID, "uid", "" %>
			</select>

			</select></div><img src="<% =sHomePath %>image/x.gif" width=1 height=1><br />
			<div align="center"><a href="javascript:void(0)" onclick="$(this).closest('form').submit()" class="red-button">Select</a></div>
			</div>
			<% ShowFeatureBoxFooter %>
			</form><br />

		<% End If %>

		<form name="cvformat" method="get" action="<%=sApplicationHomePath & "view/cv_view.asp" %>">
		<input type="hidden" name="uid" value="<% =sCvUID %>">
		<input type="hidden" name="idproject" value="<% =iProjectID %>">
		<input type="hidden" name="txt" value="<% =sSearchKeywordsHighlight %>">
		<input type="hidden" name="srch_queryadd" value="<% =sSearchKeywordsAdd %>">
		
		<% ShowFeatureBoxHeader "Format the CV" %>
		<div class="content">
		<!--<img src="<% =sHomePath %>image/x.gif" width=1 height=3><br />-->
		<div align="center"><select style="font-face: Arial; font-size:8.5pt;" name="act" size=1>
		<%
		If sCvLanguage = cLanguageFrench Then
		%>
			<option value="ASR" <% If sCvFormat = "" Then %>selected<% End If %>>assortis.com</option>
			<option value="EC" <% If sCvFormat = "EC" Then %>selected<% End If %>>European Commission</option>
			<option value="EP" <% If sCvFormat = "EP" Then %>selected<% End If %>>Europass</option>
			<option value="WB" <% If sCvFormat = "WB" Then %>selected<% End If %>>World Bank</option>
		<%
		Else
		%>
			<option value="ASR" <% If sCvFormat = "" Then %>selected<% End If %>>assortis.com</option>
			<option value="ADB" <% If sCvFormat = "ADB" Then %>selected<% End If %>>Asian Development Bank</option>
			<option value="AFB" <% If sCvFormat = "AFB" Then %>selected<% End If %>>African Development Bank</option>
			<option value="EC" <% If sCvFormat = "EC" Then %>selected<% End If %>>European Commission</option>
			<option value="EP" <% If sCvFormat = "EP" Then %>selected<% End If %>>Europass</option>
			<option value="WB" <% If sCvFormat = "WB" Then %>selected<% End If %>>World Bank</option>
		<%
		End If
		%>
		</select></div><br />
		<div align="center"><a href="javascript:void(0)" onclick="$(this).closest('form').submit()" class="red-button">Select</a></div>
		</div>
		<% ShowFeatureBoxFooter %>
		</form><br />

		
		<% ShowFeatureBoxHeader("Document options") %>
		<div class="content">
		<p class="sml"><a class="list" href="<%=Replace(sScriptFileName, "view", "save")%>?uid=<%=sCvUID%>&ftype=doc"><img src="/image/file_doc.gif" width=18 height=17 border=0 hspace=4 align="left">Save&nbsp;document</a></p>
		<p class="sml">This option will save the CV in Microsoft&reg; Word format on your local computer.</p>
		<p class="sml"><a class="list" href="<%=Replace(sScriptFileName, "view", "save")%>?uid=<%=sCvUID%>&ftype=prn"><img src="/image/file_prn.gif" width=18 height=17 border=0 hspace=4 align="left">Print&nbsp;document</a></p>
		<p class="sml">This option will open the CV in Microsoft&reg; Word  and you should click on Print button there.</p>
		</div>
		<% ShowFeatureBoxFooter %>
		<br />

		<% 	If objExpertDBOtherList.Count>0 Then
			ShowFeatureBoxHeader("CV in other DBs") %>
			<div class="content">
			<p>This expert's CV is registered in <% =ShowEntityPlural((objExpertDBOtherList.Count+1), "ICA member's", "ICA members'", " ") %> DB:<br/>
			<b><a href="mailto:<% =objExpertDB.ContactEmail %>"><% =objExpertDB.Company.Name %></a></b>,
			<% 	While iExpertDBOtherLoop < objExpertDBOtherList.Count %>
				<b><a href="mailto:<% =objExpertDBOtherList.Item(iExpertDBOtherLoop).ContactEmail %>"><% =Trim(objExpertDBOtherList.Item(iExpertDBOtherLoop).Company.Name) %></a></b><% 
				If iExpertDBOtherLoop < objExpertDBOtherList.Count - 1 Then Response.Write ","
			iExpertDBOtherLoop = iExpertDBOtherLoop + 1
			WEnd
			%>
			</p>
			</div>
			<% ShowFeatureBoxFooter %>
			<br />
		<% End If %>

		<!--
		<% ShowCvStatisticsFeatureBox %>
		<br />
		-->

		<% ShowReportCvQualityFeatureBox %>
		<br />

	<% End If

	If sScriptFileName="cv_preview.asp" Then
	%>
		<br />
		<% ShowFeatureBoxHeader("Contact details") %>
		<div class="content">
			<% If objExpertDB.Database = "assortis" And iAssortisMemberID > 0 Then %>
				<p>You can download this CV directly from <a class="list" href="https://www.assortis.com/login.asp?token=<% =sAssortisUserToken %>&act=selected&url=/en/members/<% =sIcaServerSqlPrefix %>exp_register.asp&experts=1,<% =iCvID %>"><b><% =objExpertDB.Company.Name %></b></a>:<br /></p>
				<p align="center"><a class="red-button w115" href="https://www.assortis.com/login.asp?token=<% =sAssortisUserToken %>&act=selected&url=/en/members/<% =sIcaServerSqlPrefix %>exp_register.asp&experts=1,<% =iCvID %>">Download CV</a>
				<!---->
			<% Else %>
				<p>In order to get contact details for this expert please contact:</p>
				<p><b><a class="list" href="mailto:<% =objExpertDB.ContactEmail %>"><% =objExpertDB.Company.Name %>: <% =objExpertDB.ContactName %></a></b></p>
			<% End If %>
		
		<%
		iExpertDBOtherLoop = 0
		If objExpertDBOtherList.Count > 0 Then
			%>
			<% If objExpertDB.Database = "assortis" And iAssortisMemberID > 0 Then %>
			<p>Or get contact details for this expert from:
			<% End If %>
			<%
			While iExpertDBOtherLoop < objExpertDBOtherList.Count
			%>
		<br /><p><b><a class="list" href="mailto:<% =objExpertDBOtherList.Item(iExpertDBOtherLoop).ContactEmail %>"><% =objExpertDBOtherList.Item(iExpertDBOtherLoop).Company.Name %>: <% =objExpertDBOtherList.Item(iExpertDBOtherLoop).ContactName %></a></b><br />
		(<a href="mailto:<% =objExpertDBOtherList.Item(iExpertDBOtherLoop).ContactEmail %>"><% =objExpertDBOtherList.Item(iExpertDBOtherLoop).ContactEmail %></a>)</p>
			<%
			iExpertDBOtherLoop = iExpertDBOtherLoop + 1
			WEnd
		End If
		%>
		
		</p>
		</div>
		<% ShowFeatureBoxFooter %>
		<br />
		
		<% ShowReportCvQualityFeatureBox %>
		<br />
	<%
	End If
	
	
	If sScriptFileName="cv_verify.asp" Then
		If bCvValidForMemberOrExpert=1 Then
		%>
			<% ShowFeatureBoxHeader("CV options") %>
			<div class="content">
			<p">Some information from the CV<br />is missing?</p>
			<div align="center"><a href="../register/register.asp?uid=<% =sCvUID %>" class="red-button w150">Update This CV</a></div>

			<p>If this CV is a duplicate and expert has another CV</p>
			<div align="center"><a href="../manage/cv_hide.asp?uid=<% =sCvUID %>" class="red-button w150">Hide this duplicate CV</a></div>

			<p>If expert asked to be removed or it's a fake CV</p>
			<div align="center"><a href="../manage/cv_remove.asp?uid=<% =sCvUID %>" class="red-button w150">Remove this expert</a></div>
			</div>
			<% ShowFeatureBoxFooter %>
			<br />
		<%
		Else
			Dim sVerifyParams
			sVerifyParams = ReplaceUrlParams(sVerifyParams, "act=copy")
			Dim objForm, objField
			Set objForm=Request.QueryString()

			For Each objField In objForm
				If objField<>"cvEng" Then 
					sVerifyParams = sVerifyParams & "&" & objField & "=" & objForm(objField)
				End If
			Next
			%>
		
			<% ShowFeatureBoxHeader("Copy Expert") %>
			<div class="content">
			<p>This expert is the same as we wanted to register.</p>
			<p>We verified all available information and there is no doubt about it.</p>
			<p align="center"><b><a class="mt" href="../manage/cv_copy.asp<% = ReplaceUrlParams(sVerifyParams, "uid=" & sCvUID) %>">Copy expert to <% =objUserCompanyDB.DatabaseTitle %> database</a></b></p>
			<img src="<% =sHomePath %>image/x.gif" width=21 height=1 hspace=0 vspace=25 align="left">
			<p class="sml">* This action will be checked by ICA team in order to avoid any mistake.</p>
			</div>
			<% ShowFeatureBoxFooter %>
		<%
		End If
	End If
	
	If sScriptFileName="register6.asp" Then
	%>
		<form method="post" action="<% =sScriptFullName %>">
		<% ShowFeatureBoxHeader("CV status") %>
		<div class="content">
			<p>Currently the CV status is</p>
			<div align="center"><select name="cv_status" id="cv_status" style="width: 152px;margin-bottom:5px;">
			<option value="0">New CV</option>
			<% 
			Dim objStatusCVList
			Set objStatusCVList = New CStatusCVList
			objStatusCVList.LoadData
			objStatusCVList.ShowSelectItems(objExpertStatusCV.Status.ID)
			%>
			</select></div>
			<div align="center"><input type="submit" class="red-button w150" name="btnUpdateStatus" value="Update status" /></div>
			</form>
			<br/>
			<div class="comment-container">
				<%
				' check for comments:
				Dim objTempRs2
				Set objTempRs2 = GetDataRecordsetSPWithConn(objConnCustom, "usp_ExpertCommentsSelect", Array(Array(, adInteger, , iCvID)))
				Dim sIsMyCommentFound, bCanAddComment
				sIsMyCommentFound = 0
				bCanAddComment = 1
				If objExpertDB.DatabasePath = "assortis2db" And (IsNull(iAssortisUserID) Or iAssortisUserID < 1) Then
					bCanAddComment = 0
				End If
				If Not objTempRs2.Eof Then
					While Not objTempRs2.Eof
						Dim bShowOthesComment
						bShowOthesComment = 1
						' add/edit "My" comment
						If Not IsNull(objTempRs2("CommentorIcaUserId")) Then
							If CStr(iUserID) = CStr(objTempRs2("CommentorIcaUserId")) Then
								' Or iUserID = objTempRs2("id_User_ExpCommentsManager") Then
								If bCanAddComment = 1 Then
									%><div class="comment" data-commentorUserId="<%=objTempRs2("id_User_Commentor") %>" data-expUid="<% =sCvUID %>"><span class="comment-title red">ME:</span><br/><div class="my-comment"><% =objTempRs2("Comment") %></div></div>
									<div align="center"><a href="javascript:void(0)" class="icon-edit-comment red-button w150" 
										data-expUid="<% =sCvUID %>" 
										data-expName="<%=sFirstName & " " & sLastName %>" 
										data-commentorUserId="<%=objTempRs2("id_User_Commentor") %>" 
										data-expId="<%=objExpertDB.ID %>-<%=iCvID %>"
										data-isPublic="0">Edit my comment</a></div>
									<% 
								End If
								sIsMyCommentFound = 1
								bShowOthesComment = 1
							End If
						End If
						
						If bShowOthesComment = 1 Then
							' for now - show other comments only within the member:
							If IsNull(iAssortisMemberID) Then iAssortisMemberID = 0
							If IsNull(iUserCompanyID) Then iUserCompanyID = 0
							If CInt(objTempRs2("CommentorIcaCompanyId")) = cInt(iUserCompanyID) Or CInt(objTempRs2("CommentorAssortisMemberId")) = CInt(iAssortisMemberID) Then
								%><div class="comment"><span class="comment-title"><% =objTempRs2("CommentorUserName") %>:</span><br/><% =objTempRs2("Comment") %></div><%
							End If
						End If 
						
						objTempRs2.MoveNext
					Wend
				End If
				
				If bCanAddComment = 1 And sIsMyCommentFound = 0 Then
					%><div class="comment hidden" data-commentorUserId="<%=iUserID %>" data-expUid="<% =sCvUID %>"><span class="comment-title red">ME:</span><br/><div class="my-comment"></div></div>
					<div id="NotCommented"><p style="font-style:italic">still not commented by me</p></div>
					<div align="center"><a href="javascript:void(0)" class="icon-edit-comment red-button w150" 
							data-expUid="<% =sCvUID %>" 
							data-expName="<%=sFirstName & " " & sLastName %>" 
							data-commentorUserId="<%=iUserID %>" 
							data-expId="<%=objExpertDB.ID %>-<%=iCvID %>" data-isPublic="0">Edit my comment</a></div>
					<% 
				End If
				
				objTempRs2.Close
				Set objTempRs2 = Nothing

				' show these only for IBF, because they came from Assortis DB:
				If sComments > "" And iUserCompanyID = 2 Then
					%><br/><div class="comment"><span class="comment-title">Old comment:</span><br/><% =sComments %></div>
					<%
				End If
				
				
				' OLD comments:
			'	If objUserCompanyDB.ID = objExpertDB.ID Then
			'		% ><p><% If Len(sComments)<1 Then % >Some comments about this CV?<% Else % ><% =sComments % ><% End If % ></p>
			'		<div align="center"><a href="../register/comments.asp?uid=<% =sCvUID % >" class="red-button w150">Edit comments</a></div>
			'		<%
			'	End If 
				%>
			</div>
		</div>
		<% ShowFeatureBoxFooter %>
		<br />
		

		<% ShowFeatureBoxHeader("CV options") %>
		<div class="content">
		<p">Some information from the CV<br />is missing?</p>
		<div align="center"><a href="register.asp?uid=<% =sCvUID %>" class="red-button w150">Update this CV</a></div>

		<p>If this CV is a duplicate and expert has another CV</p>
		<div align="center"><a href="../manage/cv_hide.asp?uid=<% =sCvUID %>" class="red-button w150">Hide this duplicate CV</a></div>

		<p>If expert asked to be removed or it's a fake CV</p>
		<div align="center"><a href="../manage/cv_remove.asp?uid=<% =sCvUID %>" class="red-button w150">Remove this expert</a></div>
		</div>
		<% ShowFeatureBoxFooter %>
		<br />
		
		<% ShowFeatureBoxHeader("CV formats") %>
		<div class="content">
		<p>To view this CV in different formats, to save or to print it</p>
		<div align="center"><a href="<% =sApplicationHomePath %>view/cv_view.asp<% =ReplaceUrlParams(ReplaceUrlParams(sParams, "idexpert"), "id=" & objExpertDB.DatabaseCode & iCvID) %>" class="red-button w150">Format this CV</a></div>
		</div>
		<% ShowFeatureBoxFooter %>
		<br />
		
		<script type="text/javascript">	
		$(function () {
			$('.icon-edit-comment').click(function (e) {
				$("#comment-dialog").show().css('top', (e.pageY + 10)).css('left', (e.pageX - $("#comment-dialog").width()));
				$("#comment-dialog textarea").text($(this).closest('.comment-container').find('.my-comment').text()).focus();
				$("#comment-dialog input[name='uid']").val($(this).attr('data-expUid'));
				$("#comment-dialog input[name='id_Expert']").val($(this).attr('data-expId'));
				$("#comment-dialog input[name='userid']").val($(this).attr('data-commentorUserId'));
				if ($(this).attr('data-isPublic') == '1' || $(this).attr('data-isPublic') == 'true')
				{
					$("#comment-dialog input[name='ispublic']").attr('checked', 'checked');
				}
				$("#comment-dialog .dialog-header").html('Edit your comment for <strong>' + ($(this).attr('data-expName') != '' ? $(this).attr('data-expName') : 'expert') + '</strong>');
			});
			
			$('#comment-dialog .btn-cancel').click(function () {
				$("#comment-dialog form")[0].reset();
				$("#comment-dialog").hide();
			});
			
			$('#commentform').submit(function (e) {
				e.preventDefault();
				
				$.ajax({
					cache: false,
					url: '../register/comments.asp?uid=' + $("#comment-dialog input[name='uid']").val(),
					type: "POST",
					commentorUserId: $("#comment-dialog input[name='userid']").val(),
					expUid: $("#comment-dialog input[name='uid']").val(),
					data: {
						uid: $("#comment-dialog input[name='uid']").val(),
						expertcomment: $("#comment-dialog textarea[name='expertcomment']").val(),
						id_Expert: $("#comment-dialog input[name='id_Expert']").val(),
						userid: $("#comment-dialog input[name='userid']").val(),
						ispublic: $("#comment-dialog input[name='ispublic']").is(':checked') ? 1 : 0,
					},
					success: function (result) {
						if (result == "OK")
						{
							$('.comment-container .comment[data-commentorUserId="' + this.commentorUserId + '"][data-expUid="' + this.expUid + '"]').removeClass('hidden');
							$('.comment-container .comment[data-commentorUserId="' + this.commentorUserId + '"][data-expUid="' + this.expUid + '"] .my-comment').html($("#comment-dialog textarea[name='expertcomment']").val());
							$("#comment-dialog form")[0].reset();
							$("#comment-dialog").hide();
							$('#NotCommented').hide();
						}
						else
						{
							alert("Error while processing your comment.");
						}
					},
					error: function () {
						alert("Error while processing your comment.");
					}
				});
			});
		});

		</script>
		<div id="comment-dialog">
			<div class="dialog-header">Edit your comment for expert</div>
			<div class="dialog-content">
				<form id="commentform" name="commentform" method="post">
					<input type="hidden" name="uid" />
					<input type="hidden" name="id_Expert" />
					<input type="hidden" name="userid" />
					<textarea name="expertcomment"></textarea><br/>
					<%
					' hide the option for now:
					If 1 = 2 Then
						%><input type="checkbox" name="ispublic" value="1" /> Make my comment visible for all ICA members in search results. (if not checked, the comment will be visible only for the users within My Organisation)
						<%
					Else
						%><input type="hidden" name="ispublic" value="0" />
						<%
					End If %>
					<div class="dialog-button-line">
						<input type="submit" class="btn-save red-button floatRight w125" name="btnSubmit" id="btnSubmit" value="Save &amp; Close" />
						<a href="javascript:void(0)" class="btn-cancel grey-button floatLeft">Cancel</a>
						<br class="clear" />
					</div>
				</form>
			</div>
		</div>
	<%
	End If

End Sub

' Right Column for TOP EXPERTS:
Sub ShowTopExpCVFeatureBox
	%>
	<% ShowFeatureBoxHeader("My CV options") %>
	<div class="content">
	<ul class="compact">
		<li><a href="/backoffice/mycv/register.asp"><b>UPDATE MY CV</b><!--<img src="<% =sHomePath %>image/bte_updatecv.gif" vspace="4" border="0">--></a></li>
		<li><a href="/backoffice/mycv/document.asp?document=0"><b>DOCUMENTS</b></a></li>

		<%
		Dim objDocumentList
		Set objDocumentList = New CDocumentList
		objDocumentList.LoadDocumentListByExpertID iCvID, "", NULL
		
		If objDocumentList.Count=0 Then
		%>
			<p class="sml" style="padding: 2px 5px;">There are no documents uploaded yet.</p>
		<% 
		Else
			ShowDocumentListEditTable objDocumentList
		End If
		Set objDocumentList = Nothing
		%>
	</ul>
		<div align="center"><a href="/backoffice/mycv/document.asp?document=0" class="red-button w125">Upload document</a></div>
	</div>
	<% ShowFeatureBoxFooter %>
	<br />

	<form name="cvformat" method="get" action="<%=sApplicationHomePath & "mycv/cv_view.asp" %>">
	<input type="hidden" name="uid" value="<% =sCvUID %>" />
	<input type="hidden" name="idproject" value="<% =iProjectID %>" />
	<% ShowFeatureBoxHeader "Format the CV" %>
	<div class="content">
	<div align="center"><select style="font-face: Arial; font-size:8.5pt;" name="act" size=1>
	<%
	If sCvLanguage = cLanguageFrench Then
	%>
	<option value="ASR" <% If sCvFormat = "" Then %>selected<% End If %>>assortis.com</option>
	<option value="EC" <% If sCvFormat = "EC" Then %>selected<% End If %>>European Commission</option>
	<option value="EP" <% If sCvFormat = "EP" Then %>selected<% End If %>>Europass</option>
	<option value="WB" <% If sCvFormat = "WB" Then %>selected<% End If %>>World Bank</option>
	<%
	Else
	%>
	<option value="ASR" <% If sCvFormat = "" Then %>selected<% End If %>>assortis.com</option>
	<option value="ADB" <% If sCvFormat = "ADB" Then %>selected<% End If %>>Asian Development Bank</option>
	<option value="AFB" <% If sCvFormat = "AFB" Then %>selected<% End If %>>African Development Bank</option>
	<option value="EC" <% If sCvFormat = "EC" Then %>selected<% End If %>>European Commission</option>
	<option value="EP" <% If sCvFormat = "EP" Then %>selected<% End If %>>Europass</option>
	<option value="WB" <% If sCvFormat = "WB" Then %>selected<% End If %>>World Bank</option>
	<%
	End If
	%>
	</select></div>
	<img src="<% =sHomePath %>image/x.gif" width=1 height=6><br />
	<div align="center"><a href="javascript:void(0)" onclick="$(this).closest('form').submit()" class="red-button">Select</a></div>
	</div>
	<% ShowFeatureBoxFooter %>
	</form><br />
	<%
	If Len(Request.QueryString("uid")) > 0 Then
	%>
		<% ShowFeatureBoxHeader("Document options") %>
		<div class="content">
		<p class="sml"><a href="<%=Replace(sScriptFileName, "view", "save")%>?uid=<%=sCvUID%>&ftype=doc"><img src="/image/file_doc.gif" width=18 height=17 border=0 hspace=4 align="left">Save&nbsp;as&nbsp;Word&nbsp;document</a></p>
		<p class="sml">This option will save the CV in Microsoft&reg; Word* format on your local computer.</p>
		<p class="sml"><a href="<%=Replace(sScriptFileName, "view", "save")%>?uid=<%=sCvUID%>&ftype=prn"><img src="/image/file_prn.gif" width=18 height=17 border=0 hspace=4 align="left">Print&nbsp;this&nbsp;document</a></p>
		<p class="sml">This option will open the CV in Microsoft&reg; Word*  and you should click on Print button there.</p>
		</div>
		<% ShowFeatureBoxFooter %>
		<br />
		<%
	End If
End Sub


Sub ShowCvStatisticsFeatureBox
	Exit Sub
	Dim objTempRs

	ShowFeatureBoxHeader("CV statistics") %>
	<div class="content">

	</div>
	<% ShowFeatureBoxFooter
	Set objTempRs=Nothing
End Sub

Sub ShowReportCvQualityFeatureBox
Dim objTempRs
	Set objTempRs=GetDataRecordsetSPWithConn(objConnCustom, "usp_ExpertCvQualityReportSelect", Array( _
		Array(, adInteger, , iCvID)))

	Dim iCvQualityIssue
	iCvQualityIssue=0
	If Not objTempRs.Eof Then
		iCvQualityIssue=CheckIntegerAndZero(objTempRs("expCvQuality"))
	End If	
	
	ShowFeatureBoxHeader("<small>Report CV quality issue</small>") %>
	<div class="content" <% If iCvQualityIssue=1 Then %> style="background: #f0f0f0;"<% End If %>>
		<% If iCvQualityIssue=1 Then %>
			<p class="sml">One of ICA members has reported about the quality of information encoded in this CV.</p>
			<p class="sml">The CV will be reviewed by the company in charge.</p>
		<% Else %>
			<p class="sml">If you believe there is an issue with the quality of information encoded in this CV, <br />please <a class="list" href="../register/quality.asp?uid=<%=sCvUID%>" target="_blank">report it</a>.</p>
		<% End If %>
	</div>
	<% ShowFeatureBoxFooter
Set objTempRs=Nothing
End Sub

If sApplicationName="external" Then
	If sContactDetailsExternally=cNameObfuscated Then
		sFirstName=ObfuscateString(sFirstName)
		sLastName=ObfuscateString(sLastName)
		sMiddleName=ObfuscateString(sMiddleName)
		sPermAddressEmail=ObfuscateEmail(sPermAddressEmail)
		sPermAddressPhone=ObfuscateString(sPermAddressPhone)
		sCurAddressEmail=ObfuscateEmail(sCurAddressEmail)
		sCurAddressPhone=ObfuscateString(sPermAddressPhone)
		
	End If
	If sContactDetailsExternally=cNameHidden Then
		sFirstName=""
		sLastName=""
		sMiddleName=""
	End If
End If

Function RemoveContactDetails(AText)
Dim sResult, objRegEx
	sResult=AText
	
	If IsEmpty(sResult) Or IsNull(sResult) Then Exit Function
	
	Set objRegEx = CreateObject("VBScript.RegExp")
	objRegEx.Global = True
	objRegEx.IgnoreCase = True

	' Remove first name
	'sResult=Replace(sResult, sFirstName, "..", 1, -1, vbTextCompare)
	objRegEx.Pattern = "\b" & sFirstName & "\b"
	If Len(sFirstName)>1 Then sResult = objRegEx.Replace(sResult, "..")

	' Remove last name
	'sResult=Replace(sResult, sLastName, "..", 1, -1, vbTextCompare)
	objRegEx.Pattern = "\b" & sLastName & "\b"
	If Len(sLastName)>1 Then sResult = objRegEx.Replace(sResult, "..")
	
	' Remove emails
	objRegEx.Pattern = "[a-z0-9._%+-]+@[a-z0-9.-]+\.[a-z]{2,4}"
	sResult = objRegEx.Replace(sResult, "..")

	' Remove websites
	objRegEx.Pattern = "(https?:\/\/)?(www\.)?([\da-z\.-]+)\.([a-z\.]{2,4})"
	sResult = objRegEx.Replace(sResult, "..")

	' Remove dates
	objRegEx.Pattern = "(((\d{4,})[-\s\.]+){2,})"
	sResult = objRegEx.Replace(sResult, "..")
	
	' Remove phones
	objRegEx.Pattern = "(((\d{2,})[-\s\.]+){3,4})"
	sResult = objRegEx.Replace(sResult, "..")
	objRegEx.Pattern = "(((\d{5,})[-\s\.]*)+)"
	sResult = objRegEx.Replace(sResult, "..")

	Set objRegEx = Nothing

RemoveContactDetails=sResult
End Function

Function RemoveContactDetailsIfNoAccess(AText)
Dim sResult
	sResult=AText
	If bCvValidForMemberOrExpert <> aClientSecurityCvViewEnabled And bCvValidForMemberOrExpert<>aClientSecurityCvViewAll Then
		sResult=RemoveContactDetails(sResult)
	End If
RemoveContactDetailsIfNoAccess=sResult
End Function
%>