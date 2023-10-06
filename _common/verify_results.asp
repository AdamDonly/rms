<%
'--------------------------------------------------------------------
'
' CV registration.
' Short format.
'
'--------------------------------------------------------------------
Response.Expires = 0
Response.ExpiresAbsolute = Now() - 1
Response.CacheControl = "no-cache"
%>
<!--#include file="_data/datGender.asp"-->
<!--#include file="_data/datPsnTitle.asp"-->
<!--#include file="_data/datMonth.asp"-->
<%
' Remove inactive url params
sParams=ReplaceUrlParams(sParams, "url")

If sApplicationName<>"expert" Then
	CheckUserLogin sScriptFullNameAsParams
End If
%>
<!--#include file="../_common/expProfile.asp"-->
<%
Dim objResult, iResult, objForm, objField
Dim sUserLogin, sUserPassword, sUserPhone
If sAction="push" Then
	Set objForm=Request.QueryString()
	%>
<!--#include file="../_common/verify_push.asp"-->
	<%
ElseIf Request.Form()>"" Then
	Set objForm=Request.Form()

	Set objResult=VerifyExpertProfile(objForm)

	' If there is no expert with similar details
	If objResult.Eof Then
		objResult.Close
		Set objResult=Nothing
		%>
<!--#include file="../_common/verify_push.asp"-->
		<%
	End If
	' If there are experts with similar contact details

	%>
<!--#include file="../_common/verify_list.asp"-->
	<%	
	Set objResult=Nothing
End If
Set objForm=Nothing
%>
