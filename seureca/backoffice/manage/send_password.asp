<%@ LANGUAGE="VBSCRIPT" %>
<% Option Explicit %>

<!--#include file="../../dbc.asp"-->
<!--#include file="../../fnc.asp"-->
<!--#include file="../../_forms/frmInterface.asp"-->
<!--#include file="../../../_common/expProfile.asp"-->

<%
Dim objResult
Dim iTempExpertID, iTempUserID, sTempUserLogin, sTempUserPassword

Set objResult=GetDataRecordsetSP("usp_AdmExpAllListExtraModifiedWithPasswordSelect", Array( _
	Array(, adVarChar, 250, Null), _
	Array(, adInteger, , 0), _
	Array(, adInteger, , 0), _
	Array(, adVarChar, 100, NULL), _
	Array(, adVarChar, 255, "no_password"), _
	Array(, adVarChar, 100, NULL), _
	Array(, adVarChar, 16, NULL), _
	Array(, adVarChar, 16, NULL), _
	Array(, adVarChar, 16, NULL), _
	Array(, adVarChar, 16, NULL) _
	))

	While Not objResult.Eof
		If objResult("id_Expert")>0 Then
			iTempExpertID=objResult("id_Expert")
			iTempUserID=objResult("id_User")
			sTempUserLogin=objResult("UserName")
			sTempUserPassword=objResult("Password")

			LoadExpertProfile(iTempExpertID)
			
			sUserEmail=Trim(sUserEmail)
			If Len(sUserEmail)>5 And InStr(sUserEmail, "@")>0 Then
			
				PrepareEmailTemplate "emlExpertAccountEncoded.htm", ";;sExpertFullName=" & sFullName & ";;sUserLogin=" & sTempUserLogin & ";;sUserPassword=" & sTempUserPassword & ";;sSystemUrl=" & "http://cvip.assortis.com" & sHomePath & "apply/login.asp" 
				'SendEmail sEmailCvipSystem, sUserEmail, sEmailSubject, sEmailBody, "info"
				SendEmail sEmailCvipSystem, sEmailClient, sEmailSubject, sEmailBody, "info"

				SaveExpertAccountEmailSent iTempExpertID, 1
			End If
			
			Response.Write sFullName &  ": " & sTempUserLogin & " " & sTempUserPassword & "<br/>"
		End If
		
		objResult.MoveNext
	WEnd
%>