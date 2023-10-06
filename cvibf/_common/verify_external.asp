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
	'CheckUserLogin sScriptFullNameAsParams
End If
%>
<!--#include virtual="/_common/expProfile.asp"-->
<%
Dim objResult, iResult, objUpload, objForm, objField
Dim sUserLogin, sUserPassword, sUserPhone
Dim sFileName, sFileExtension, sFullFileName, Msg, objFile
Dim bShowResults
bShowResults = Request.QueryString("show_results")

Set objForm=Request.QueryString()
Set objResult=VerifyExpertProfile(objForm)
' If there is no expert with similar contact details
If objResult.Eof Then
	objResult.Close
	Set objResult=Nothing
	Response.Clear
	Response.Write "No"
	Response.End
Else
' If there are experts with similar contact details
	Response.Clear
	Response.Write "Yes"
	Response.Flush
	If bShowResults="1" Then %>
		<!--#include virtual="/_common/verify_list.asp"-->
	<% End If
	Response.End
End If
Set objResult=Nothing
%>