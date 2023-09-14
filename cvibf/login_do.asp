<%@ LANGUAGE="VBSCRIPT" %>
<% Option Explicit 
Response.Expires = 0
Response.ExpiresAbsolute = Now() - 1
Response.CacheControl = "no-cache" %>
<!--#include file="dbc.asp"-->
<!--#include file="fnc.asp"-->
<!--#include file="_encryption.asp"-->
<%
Dim sUserLogin, sUserPassword, sUserPasswordHash, sShortUserType, sIpFilterMatch
Dim bUserLoggedIn

If sAction="online" Then
	sUserLogin = CheckString(Request.QueryString("login_name"))
	sUserPassword = CheckString(Request.QueryString("login_pwd"))
	iUserID = CheckString(Request.QueryString("login_usr"))
	sUrl = ReplaceUrlParams(sUrl, "login_name")
	sUrl = ReplaceUrlParams(sUrl, "login_pwd")
	sUrl = ReplaceUrlParams(sUrl, "login_usr")
	sUrl = ReplaceUrlParams(sUrl, "act")
Else
	sUserLogin = CheckString(Request.Form("login_name"))
	sUserPassword = CheckString(Request.Form("login_pwd"))
End If

If sUserLogin > "" And sUserPassword > "" Then
	iUserID = 0
	iMemberID = 0
	iExpertID = 0

	If Len(sUserPassword) > 0 Then 
		sUserPasswordHash = Encryption.Base64Sha1EncodeWithSalt(sUserPassword)
	End If

	Set objTempRs = GetDataRecordsetSP("usp_UserLoginHashAssortis", Array( _
		Array(, adVarWChar, 50, sUserLogin), _
		Array(, adVarWChar, 50, sUserPassword), _
		Array(, adVarWChar, 50, sUserPasswordHash), _
		Array(, adVarChar, 40, sSessionID), _
		Array(, adVarChar, 16, sUserIpAddress), _
		Array(, adVarChar, 60, NULL)))

	If Not objTempRs.Eof Then
		bUserLoggedIn = objTempRs("usrLoggedIn")
		iUserID = objTempRs("id_User")
		sUserType = objTempRs("usrType")
		'iUserTypeAccessSecurity = GetAccessSecurity(sUserType)
		iMemberID = objTempRs("id_Member")
		iExpertID = objTempRs("id_Expert")
		sUserLanguage = objTempRs("usrLanguage")
		If sUserLanguage = "" Then
			sUserLanguage = "Eng"
		End If
		sUserEmail = objTempRs("usrEmail")
		sIpFilterMatch = objTempRs("usrIpSecurityMatch")

		If (bUserLoggedIn = 0) Or (iUserID = 0) Then
	 		' Access Denied
			sParams = ReplaceUrlParams(sParams, "act=" & sAction)
			sParams = ReplaceUrlParams(sParams, "err=NoMatch")
			Response.Redirect(sApplicationHomePath & "login.asp" & sParams)
		ElseIf sUserType<>"Admin" And sUserType<>"CV Ibf" Then
	 		' Access Denied
			sParams = ReplaceUrlParams(sParams, "act=" & sAction)
			sParams = ReplaceUrlParams(sParams, "err=NoMatch")
			Response.Redirect(sApplicationHomePath & "login.asp" & sParams)
		Else
			' Check limitation on IP address for the user
			If sIpFilterMatch = 0 Then
				Session.Abandon
				iUserID = 0
				iMemberID = 0
				iExpertID = 0
				sSessionID = ""
				sCookiesSessionID=""
				Response.Cookies("SessionID")=""

				Response.Redirect sHomePath & "denied.asp"
			End If
			
			sUrl = ReplaceIfEmpty(sUrl, "/backoffice/")
			sParams = ReplaceUrlParams(sParams, "url")
			Response.Redirect sUrl

			'Response.Write bUserLoggedIn & "<br>"
			'Response.Write iUserID & "<br>"
			'Response.Write sUserType & "<br>"
			'Response.Write iMemberID & "<br>"
			'Response.Write iExpertID & "<br>"
			'Response.Write sUserLanguage & "<br>"
			'Response.Write sUserEmail & "<br>"
			'Response.Write sIpFilterMatch & "<br>"
			'Response.End
		End If
	End If
End If

Response.Redirect sApplicationHomePath
%>
  
