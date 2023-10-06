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
<!--#include virtual="/fnc_file.asp"-->
<%
' Remove inactive url params
sParams=ReplaceUrlParams(sParams, "url")

If sApplicationName<>"expert" Then
	CheckUserLogin sScriptFullNameAsParams
End If
%>
<!--#include virtual="/_common/expProfile.asp"-->
<%
Dim objResult, iResult, objForm, objField
Dim sUserLogin, sUserPassword, sUserPhone
Dim sFileName, sFileExtension, sFullFileName, Msg, myfile
	Set objForm=Request.Form
	Set objResult=VerifyExpertProfile(objForm)
	' If there is no expert with similar details
	If objResult.Eof Then
		objResult.Close
		Set objResult=Nothing
		Response.Redirect "exists_nothing.asp"
	Else
	' If there are experts with similar contact details
	%>
<!--#include virtual="/_common/exists_list.asp"-->
	<%
	End If
	Set objResult=Nothing
%>
