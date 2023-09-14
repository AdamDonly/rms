<%
' Time to live in minutes for open sessions
Dim iOpenSessionTTL	
' Time to live in minutes for closed sessions (after browser is closed) 
Dim iClosedSessionTTL
iOpenSessionTTL=240
iClosedSessionTTL=0

' Time to live in minutes for empty browser sessions but with saved information in the db. 
' Used to avoid extra sid params when cookies are enabled
Dim iUrlNoCookiesTTL
iUrlNoCookiesTTL=0

' Type of the active user
Dim sUserType, iUserAdmin

' IDs and some settings of the active user
Dim iUserID, iMemberID, iExpertID, sUserName, sUserLanguage, sUserEmail, sBackOffice
Dim sExpertUID

Dim sUserFullName, sUserCompany, iUserCompanyID, sUserCompanyRole, iUserCompanyRoleID, iUserCompanyMemberType, bUserMpisAccess, bUserIbfStaff

Dim uUserTopExpertUid, bIsUserTopExpert, bIsMyCV
Dim bUserMaskProjects, bUserMaskExperts
Dim iMemberAccessExperts

' User browser and ip
Dim sUserAgent, sUserIpAddress, sUserIpBase
sUserAgent=Request.ServerVariables("HTTP_USER_AGENT")
sUserIpAddress=Request.ServerVariables("REMOTE_ADDR")

' Session state
Dim sSessionID, sCookiesSessionID, bNewSession, sAccessType

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
If iUserID>0 Then
	Set objTempRs=GetDataRecordsetSP("usp_LogSessionUserDataIcaAdvancedSelect", Array( _
		Array(, adInteger, , iUserID)))
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
			uUserTopExpertUid = objTempRs("UserTopExpertUid")
			If Len(uUserTopExpertUid) >= 36 Then
				bIsUserTopExpert = True
				bIsMyCV = True
			Else
				bIsUserTopExpert = False
				bIsMyCV = False
			End If
			
			If iUserCompanyMemberType = cMemberTypeTechnicalAssociate Then
				iMemberAccessExperts = cMemberAccessExpertsOwnOnly
			ElseIf iUserCompanyMemberType = cMemberTypeLocalAssociate Then
				iMemberAccessExperts = cMemberAccessExpertsOwnOnly
			Else
				iMemberAccessExperts = CheckIntegerAndZero(objTempRs("UserCompanyMaskExperts"))
			End If
			
			
			bUserMaskProjects = objTempRs("UserMaskProjects")
			bUserMaskExperts = objTempRs("UserMaskExperts")
			
			bUserMpisAccess = objTempRs("UserMpisAccess")
			' MpisAccess only for IBF users
			If iUserCompanyID = 2 And InStr(1, sUserCompany, "IBF", vbTextCompare) > 0 And bUserMpisAccess = 1 Then
			Else
				bUserMpisAccess = 0
			End If
			
			If iUserCompanyID = 2 And InStr(1, sUserCompany, "IBF", vbTextCompare) > 0 Then
				bUserIbfStaff = 1
			Else
				bUserIbfStaff = 0
			End If
		Else
			iMemberID = 0
			iExpertID = 0
		End If
	objTempRs.Close
	Set objTempRs = Nothing
Else
	iMemberID=0
	iExpertID=0
End If

' Writing Session Event Log
objTempRs=InsertRecordSP("usp_LogSessionEvent", Array( _
	Array(, adVarChar, 40, sSessionID), _
	Array(, adVarChar, 500, Left(sScriptBaseName & sScriptFileName & sParams, 500))),"")
	Set objTempRs=Nothing	

' Setting up the default language	
If sUserLanguage="" Then
	sUserLanguage="Eng"
End If

'--------------------------------------------------------------------
' Procedure for abanding the session
'--------------------------------------------------------------------
Sub AbandonSession
	Session.Abandon
	iUserID=0
	iMemberID=0
	iExpertID=0
	sSessionID=""
	sCookiesSessionID=""
	Response.Cookies("SessionID")=""
End Sub

Sub AbandonWrongApplicationSession
If sUserType<>sApplicationName Then
	Session.Abandon
	iUserID=0
	iMemberID=0
	iExpertID=0
	sSessionID=""
	sCookiesSessionID=""
	Response.Cookies("SessionID")=""
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
