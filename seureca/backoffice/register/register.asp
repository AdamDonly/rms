<%@ LANGUAGE="VBSCRIPT" %>
<% Option Explicit %>

<!--#include file="../../dbc.asp"-->
<!--#include file="../../fnc.asp"-->
<!--#include file="../../_forms/frmInterface.asp"-->
<!--#include file="../../../_common/register_seureca.asp"-->

<%
Function AfterCvRegistrationStep1(AResult)
Dim iTempExpertID, iTempUserID, sTempUserLogin, sTempUserPassword
	If IsArray(AResult) Then
		If sApplicationName="backoffice" Then
			If objResult(1)>0 Then
				iTempExpertID=objResult(1)
				iTempUserID=objResult(2)
				sTempUserLogin=objResult(3)
				sTempUserPassword=objResult(4)

				LoadExpertProfile(iTempExpertID)
				
				If bEmailExpertAccountSent = 0 Then
					PrepareEmailTemplate "emlExpertAccountEncoded.htm", ";;sExpertFullName=" & sFullName & ";;sUserLogin=" & sTempUserLogin & ";;sUserPassword=" & sTempUserPassword & ";;sSystemUrl=" & "http://cvip.assortis.com" & sHomePath & "apply/login.asp" 
					'SendEmail sEmailCvipSystem, sUserEmail, sEmailSubject, sEmailBody, "info"
					SendEmail sEmailCvipSystem, sEmailClient, sEmailSubject, sEmailBody, "info"

					SaveExpertAccountEmailSent iTempExpertID, 1
				End If
			End If
		End If
	End If
End Function
%>