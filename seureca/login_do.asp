<%@ LANGUAGE="VBSCRIPT" %>
<% Option Explicit 
Response.Expires = 0
Response.ExpiresAbsolute = Now() - 1
Response.CacheControl = "no-cache" %>
<!--#include file="dbc.asp"-->
<!--#include file="fnc.asp"-->
<%
Dim sUserLogin, sUserPassword, sShortUserType, sIpFilter
Dim bUserLoggedIn

If sAction="online" Then
	sUserLogin=CheckString(Request.QueryString("login_name"))
	sUserPassword=CheckString(Request.QueryString("login_pwd"))
	iUserID=CheckString(Request.QueryString("login_usr"))
	sUrl=ReplaceUrlParams(sUrl, "login_name")
	sUrl=ReplaceUrlParams(sUrl, "login_pwd")
	sUrl=ReplaceUrlParams(sUrl, "login_usr")
	sUrl=ReplaceUrlParams(sUrl, "act")
Else
	sUserLogin=CheckString(Request.Form("login_name"))
	sUserPassword=CheckString(Request.Form("login_pwd"))
End If

If sUserLogin>"" And sUserPassword>"" Then
	iUserID=0
	iMemberID=0
	iExpertID=0

	objTempRs=GetDataOutParamsSP("usp_UsrLogin", Array( _
		Array(, adVarWChar, 50, sUserLogin), _
		Array(, adVarWChar, 50, sUserPassword), _
		Array(, adVarChar, 40, sSessionID)), Array( _
		Array(, adSmallInt), _
		Array(, adInteger), _
		Array(, adVarChar, 20), _
		Array(, adInteger), _
		Array(, adInteger), _
		Array(, adVarChar, 3), _
		Array(, adVarChar, 50), _
		Array(, adVarChar, 16)))

	bUserLoggedIn=objTempRs(0)
	iUserID=objTempRs(1)
	sUserType=objTempRs(2)
'	iMemberID=objTempRs(3)
	iExpertID=objTempRs(4)
	sUserEmail=objTempRs(6)
	sIpFilter=objTempRs(7)
	Set objTempRs=Nothing	

	If (bUserLoggedIn=0) Or (iUserID=0) Then
 		' Access Denied
		sParams=ReplaceUrlParams(sParams, "act=" & sAction)
		sParams=ReplaceUrlParams(sParams, "err=NoMatch")
		Response.Redirect(sApplicationHomePath & "login.asp" & sParams)
	Else
		
		' Check limitation on IP address for the user
		If (sIpFilter>"" And InStr(sUserIpAddress, sIpFilter)>0) Or (sIpFilter="") Then
		Else
			Session.Abandon
			iUserID=0
			iMemberID=0
			iExpertID=0
			sSessionID=""
			sCookiesSessionID=""
			Response.Cookies("SessionID")=""

			Response.Redirect sHomePath & "denied.asp"
		End If
		
		sParams=ReplaceUrlParams(sParams, "url")
		Response.Redirect sUrl
	End If
Else
	Response.Redirect sApplicationHomePath & "default.asp"
End If
%>
  
