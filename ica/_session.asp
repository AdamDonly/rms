<%
' Time to live in minutes for open sessions
Dim iOpenSessionTTL	
' Time to live in minutes for closed sessions (after browser is closed) 
Dim iClosedSessionTTL
iOpenSessionTTL = 240
iClosedSessionTTL = 0

' Time to live in minutes for empty browser sessions but with saved information in the db. 
' Used to avoid extra sid params when cookies are enabled
Dim iUrlNoCookiesTTL
iUrlNoCookiesTTL = 0

' Type of the active user
Dim sUserType, iUserAdmin

' IDs and some settings of the active user
Dim iUserID, iMemberID, iExpertID, sUserName, sUserLanguage, sUserEmail, sBackOffice
Dim sExpertUID

Dim iAssortisMemberID, _
	iAssortisUserID, _
	sAssortisUserToken

Dim iAssortisSubscriptionEdbStatus, _
	dAssortisSubscriptionEdbExpiryDate, _
	bAssortisSubscriptionEdbActive

Dim iAssortisSubscriptionDtaStatus, _
	dAssortisSubscriptionDtaExpiryDate, _
	bAssortisSubscriptionDtaActive

Dim sUserFullName, _
	sUserCompany, _
	iUserCompanyID, _
	sUserCompanyRole, _
	iUserCompanyRoleID, _
	iUserCompanyMemberType, _
	bUserIbfStaff, _
	bUserIcaStaff, _
	sAsortisLoginToken

Dim bUserAccessMpis, _
	bUserAccessMethodology, _
	bUserAccessMethodologyIbfComments, _
	bUserAccessMethodologyAsrComments, _
	bUserModifyMethodology, _
	iMemberAccessExperts,_
	iUserAccessExpertsTab, _
	iUserAccessExpertsOtherDbNoAccess, _
	iUserAccessExpertsOtherDbRestricted, _
	iUserAccessMaskExperts, _
	iUserAccessMaskProjects

Dim uUserTopExpertUid, _
	bIsUserTopExpert, _
	bIsMyCV

' User browser and ip
Dim sUserAgent, _
	sUserIpAddress, _
	sUserIpBase
	
sUserAgent = Request.ServerVariables("HTTP_USER_AGENT")
sUserIpAddress = Request.ServerVariables("REMOTE_ADDR")

' Session state
Dim sSessionID, _
	sCookiesSessionID, _
	bNewSession, _
	sAccessType

' Session state could be kept either in cookies or in url params
bNewSession=1
sCookiesSessionID=Request.Cookies("SessionID")
sSessionID=Request.QueryString("sid")
If InStr(sSessionID, ",")>0 And IsArray(Split(sSessionID, ",")) Then 
	sSessionID=Split(sSessionID, ",")(0)
	sParams=ReplaceUrlParams(sParams, "sid=" & sSessionID)
End If 

' If user has cookies enabled
If sCookiesSessionID>"" Then
	bNewSession=0
	sSessionID=sCookiesSessionID
	sParams=ReplaceUrlParams(sParams, "sid")
	
' If user has cookies disabled
ElseIf sSessionID>"" And sCookiesSessionID="" Then
	bNewSession=0
	sParams=ReplaceUrlParams(sParams, "sid=" & sSessionID)
End If

' Checking if SessionID is still valid.
If bNewSession=0 Then
	objTempRs=GetDataOutParamsSP("usp_LogSessionValidate", Array( _
		Array(, adVarChar, 40, sSessionID), _ 
		Array(, adVarWChar, 255, sUserAgent), _
		Array(, adVarChar, 16, sUserIpAddress), _ 
		Array(, adInteger, , iOpenSessionTTL)), Array( _ 
		Array(, adSmallInt), _ 
		Array(, adInteger)))

	' if user's SessionID is valid then get UserID
	If objTempRs(0)=1 Then
		iUserID=objTempRs(1)
	Else
		AbandonSession
	End If	
	Set objTempRs=Nothing
End If

' If new user without SessionID
If sSessionID="" And sCookiesSessionID="" Then
	' Creating SessionID
	objTempRs=GetDataOutParamsSP("usp_LogSessionCreate", Array( _
		Array(, adInteger, , 0), _
		Array(, adVarWChar, 255, sUserAgent), _
		Array(, adVarChar, 16, sUserIpAddress), _
		Array(, adInteger, , iUrlNoCookiesTTL)), Array( _ 
		Array(, adVarChar, 40), _
		Array(, adTinyInt)))

	sSessionID=objTempRs(0)
	sCookiesSessionID=sSessionID
	Response.Cookies("SessionID")=sSessionID
	If iClosedSessionTTL>0 Then
		Response.Cookies("SessionID").Expires=DateAdd("n", iClosedSessionTTL, Now())
	End If

	' If cookies are empty but session is open -> cookies are disabled. 
	' Include SessionID in url
	If objTempRs(1)=1 Then
		bNewSession=0
		sParams=ReplaceUrlParams(sParams, "sid=" & sSessionID)
		Response.Redirect sScriptFileName & sParams
	End If
	Set objTempRs=Nothing	
End If

' If user logged in then select his personal data
If iUserID > 0 Then
	Set objTempRs = GetDataRecordsetSP("usp_" & sIcaServerSqlPrefix & "LogSessionUserDataIcaAdvancedSelect", Array( _
		Array(, adInteger, , iUserID)))
	On Error Resume Next
		If Not objTempRs.Eof Then
			iMemberID = objTempRs(0)
			iExpertID = objTempRs(1)
			sUserName = objTempRs(2)
			sUserLanguage = objTempRs(3)
			sUserEmail = objTempRs(4)
			sUserType = objTempRs("UserType")
			sUserFullName = objTempRs("UserFullName")
			sUserCompany = objTempRs("UserCompany")
			iUserCompanyID = objTempRs("UserCompanyID")
			sUserCompanyRole = objTempRs("UserCompanyRole")
			iUserCompanyRoleID = objTempRs("UserCompanyRoleID")
			iUserCompanyMemberType = objTempRs("UserCompanyMemberType")

			iUserAccessExpertsTab = objTempRs("UserAccessExpertsTab")
			iUserAccessExpertsOtherDbNoAccess = CheckIntegerAndZero(objTempRs("UserAccessExpertsOtherDbNoAccess"))
			iUserAccessExpertsOtherDbRestricted = CheckIntegerAndZero(objTempRs("UserAccessExpertsOtherDbLimited"))
			uUserTopExpertUid = objTempRs("UserTopExpertUid")
			If Len(uUserTopExpertUid) >= 36 Then
				bIsUserTopExpert = True
				bIsMyCV = True
			Else
				bIsUserTopExpert = False
				bIsMyCV = False
			End If

			iAssortisMemberID = objTempRs("UserAssortisMemberID")
			iAssortisUserID = objTempRs("UserAssortisUserID")
			sAssortisUserToken = objTempRs("UserAssortisUserToken")

			bAssortisSubscriptionDtaActive = False
			bAssortisSubscriptionEdbActive = False

			If iAssortisMemberID > 0 Then
				iAssortisSubscriptionEdbStatus = objTempRs("EDB_AccountStatus")
				dAssortisSubscriptionEdbExpiryDate = objTempRs("EDB_AccountExpirationDate")
				If iAssortisSubscriptionEdbStatus = 1 And DateDiff("d", Now(), dAssortisSubscriptionEdbExpiryDate) >= 0 Then
					bAssortisSubscriptionEdbActive = True
				End If
			End If

			If iUserCompanyMemberType = cMemberTypeTechnicalAssociate Then
				iMemberAccessExperts = cMemberAccessExpertsOwnOnly
			ElseIf iUserCompanyMemberType = cMemberTypeLocalAssociate Then
				iMemberAccessExperts = cMemberAccessExpertsOwnOnly
			ElseIf iUserAccessExpertsOtherDbNoAccess = 1 Then
				iMemberAccessExperts = cMemberAccessExpertsOwnOnly
			ElseIf iUserAccessExpertsOtherDbRestricted = 1 Then
				iMemberAccessExperts = cMemberAccessExpertsRestricted
			Else			
				iMemberAccessExperts = cMemberAccessExpertsNormal
			End If
			
			bUserAccessMethodology = False
			bUserAccessMethodologyIbfComments = False
			bUserAccessMethodologyAsrComments = False
			bUserModifyMethodology = False

			'22		Frederic
			'135	Vitaly
			'235	Gabriele Bendazzi
			'729	Giulia Albertino
			'1198	Angela Spettoli		108786
			'1245	Vanille Storm
			'1348	Lorenzo Pedretti	113182
			'1457	Ramona Boucas		126447
			'1761	Roberta Aralla 		150601
			'1767	Natasha Ristic 		150853
			'1786	Roberta Resmini  	152127
			'1801	Chiara Giusti 	 	152994
			'1837	Maddalena Lorenzato 	156406
			'1839	Clementina Udine	156752
			'1853	Annalaura Sbrizzi	157431
			'1854	Laure Becuywe
			'1866	Leonardo Gorrieri	158081
			'1923	Manuella Markaj		161041

			' Assortis users
			'730    Giada Rapalino		92778
			'1040	Anastasia Cherepova

			' Please don't change the file and don't add any user without approval of IT Director! Thank you.
			If iUserID = 22 _
			Or iUserID = 135 _
			Or iUserID = 235 _
			Or iUserID = 729 _
			Or iUserID = 1198 _
			Or iUserID = 1245 _
			Or iUserID = 1348 _
			Or iUserID = 1457 _
			Or iUserID = 1761 _
			Or iUserID = 1767 _
			Or iUserID = 1786 _
			Or iUserID = 1801 _
			Or iUserID = 1837 _
			Or iUserID = 1839 _
			Or iUserID = 1853 _
			Or iUserID = 1854 _
			Or iUserID = 1866 _
			Or iUserID = 1923 _
			Then
				bUserAccessMethodology = True
				bUserAccessMethodologyIbfComments = True
				bUserModifyMethodology = True
			End If

			If iUserID = 730 _
			Or iUserID = 1040 _
			Then
				bUserAccessMethodology = True
				bUserAccessMethodologyAsrComments = True
				bUserModifyMethodology = True
			End If
		
			iUserAccessMaskProjects = objTempRs("UserMaskProjects")
			iUserAccessMaskExperts = objTempRs("UserMaskExperts")
			
			bUserAccessMpis = objTempRs("UserMpisAccess")
			' MpisAccess only for IBF users
			If iUserCompanyID = 2 And InStr(1, sUserCompany, "IBF", vbTextCompare) > 0 And bUserAccessMpis = 1 Then
			Else
				bUserAccessMpis = 0
			End If
			
			If iUserCompanyID = 2 And InStr(1, sUserCompany, "IBF", vbTextCompare) > 0 Then
				bUserIbfStaff = 1
			Else
				bUserIbfStaff = 0
			End If
			If iUserCompanyID = 1 And InStr(1, sUserCompany, "ICA", vbTextCompare) > 0 Then
				bUserIcaStaff = 1
			Else
				bUserIcaStaff = 0
			End If
		Else
			iMemberID = 0
			iExpertID =0
		End If
	On Error Goto 0
	objTempRs.Close
	Set objTempRs = Nothing

	' get user's Assortis login token:
	Set objTempRs = GetDataRecordsetSP("[ica].[dbo].[GetAssortisToken]", Array( _
		Array(, adInteger, , iUserID)))
	On Error Resume Next
		If Not objTempRs.Eof Then
			sAsortisLoginToken = objTempRs(0)
		End If
	On Error Goto 0
	objTempRs.Close
	Set objTempRs = Nothing
Else
	iMemberID = 0
	iExpertID = 0
End If

' Writing Session Event Log
objTempRs = InsertRecordSP("usp_LogSessionEvent", Array( _
	Array(, adVarChar, 40, sSessionID), _
	Array(, adVarChar, 500, Left(sScriptBaseName & sScriptFileName & sParams, 500))),"")
Set objTempRs = Nothing	

' Setting up the default language	
If sUserLanguage = "" Then
	sUserLanguage = "Eng"
End If

'--------------------------------------------------------------------
' Procedure for abanding the session
'--------------------------------------------------------------------
Sub AbandonSession
	Session.Abandon
	iUserID = 0
	iMemberID = 0
	iExpertID = 0
	sSessionID = ""
	sCookiesSessionID = ""
	Response.Cookies("SessionID") = ""
End Sub

Sub AbandonWrongApplicationSession
If sUserType <> sApplicationName Then
	Session.Abandon
	iUserID = 0
	iMemberID = 0
	iExpertID = 0
	sSessionID = ""
	sCookiesSessionID = ""
	Response.Cookies("SessionID") = ""
End If
End Sub

'--------------------------------------------------------------------
' Company Member Type
'--------------------------------------------------------------------
Const cMemberTypeAdmin = 1
Const cMemberTypeTechnicalAssociate = 3
Const cMemberTypeLocalAssociate = 5

'--------------------------------------------------------------------
' Company Access EDB
'--------------------------------------------------------------------
Const cMemberAccessExpertsOwnOnly = 0
Const cMemberAccessExpertsRestricted = 1
Const cMemberAccessExpertsForbidden = 2
Const cMemberAccessExpertsNormal = 4

'--------------------------------------------------------------------
' User Role Type
'--------------------------------------------------------------------
Const cUserRoleCompanyUser = 1
Const cUserRoleCompanyDashboardViewer = 2
Const cUserRoleCompanyAdministrator = 3
Const cUserRoleGlobalDashboardViewer = 4
Const cUserRoleGlobalAdministrator = 5
Const cUserRoleLocalICAContactPoint = 6
Const cUserRoleTopExpert = 7
Const cUserRoleCVContactPoint = 8
%>
