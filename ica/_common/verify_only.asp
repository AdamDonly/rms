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
<!--#include virtual = "/fnc_file.asp"-->
<%
' Remove inactive url params
sParams = ReplaceUrlParams(sParams, "url")

If sApplicationName <> "expert" Then
	CheckUserLogin sScriptFullNameAsParams
End If
%>
<!--#include virtual = "/_common/expProfile.asp"-->
<%
Dim objResult, iResult, objUpload, objForm, objField
Dim sUserLogin, sUserPassword, sUserPhone

	Set objForm = Request.QueryString()

	Set objResult = VerifyExpertProfile(objForm)
	' If there is no expert with similar contact details
	If objResult.Eof Then
		objResult.Close
		Set objResult = Nothing
		%>
		There are no experts matching the provided criteria
		<%
	Else
		' If there are experts with similar contact details
		%>
		<!--#include virtual = "/_common/verify_list.asp"-->
		<%	
	End If
	Set objResult = Nothing
%>
