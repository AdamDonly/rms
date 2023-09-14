<% 
Response.Expires = 0
Response.ExpiresAbsolute = Now() - 1
Response.CacheControl = "no-cache" 
%>
<!--#include file="_data/datGender.asp"-->
<!--#include file="_data/datPsnTitle.asp"-->
<!--#include file="_data/datPsnStatus.asp"-->
<!--#include file="_data/datMonth.asp"-->
<!--#include file="_data/datCountry.asp"-->
<%
If sApplicationName<>"expert" Then
	CheckUserLogin sScriptFullNameAsParams
End If
CheckExpertID
%>
<!--#include file="../_common/expProfile.asp"-->
<%
Dim sUserLogin, sUserPassword, sUserPhone

LoadExpertProfile iExpertID

If Not (Len(sUserEmail)>0 And InStr(sUserEmail, "@")>0) Or IsEmpty(sUserEmail) Or IsNull(sUserEmail) Then
	Response.Write "<p>Error: there is no valid email registered for " & sFullName & ".</p>"
	Response.End
End If

LoadExpertAccountDetails iExpertID

If Len(sUserLogin)>2 And Len(sUserPassword)>2 Then
	If bEmailExpertAccountSent = 0 Or bEmailExpertAccountSent = 2 Then
		PrepareEmailTemplate "emlExpertAccountEncoded.htm", ";;sExpertFullName=" & sFullName & ";;sUserLogin=" & sUserLogin & ";;sUserPassword=" & sUserPassword & ";;sSystemUrl=" & "http://cvip.assortis.com" & sHomePath & "apply/login.asp" 
		SendEmail sEmailCvipSystem, sUserEmail, sEmailSubject, sEmailBody, "info"

		SaveExpertAccountEmailSent iExpertID, 1
	End If
End If

Response.Redirect "register6.asp?id=" & iExpertID & "&act=resent"
%>
