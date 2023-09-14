<% 
'--------------------------------------------------------------------
'
' CV registration.
' Removing expert from the database
'
'--------------------------------------------------------------------
%>
<!--#include file="../expProfile.asp"-->
<% 
CheckUserLogin sScriptFullNameAsParams
CheckExpertID()

sParams=ReplaceUrlParams(sParams, "id")
sParams=ReplaceUrlParams(sParams, "idproject")
sParams=ReplaceUrlParams(sParams, "idexpert")

If iExpertID>0 Then
	ResetPassword(iExpertID)
End If

Function ResetPassword(iExpertID)
Dim objResult
Dim sTempUserLogin, sTempUserPassword
	If sApplicationName="backoffice" Then
	
		objResult=GetDataOutParamsSP("usp_ExpertPasswordReset", _
		Array( _
			Array(, adInteger, , iExpertID)), _
		Array( _
			Array(, adInteger), _
			Array(, adVarChar, 255), _
			Array(, adVarChar, 255)))
	
		If objResult(0)=0 Then
			sTempUserLogin=objResult(1)
			sTempUserPassword=objResult(2)

			LoadExpertProfile(iExpertID)
			
			PrepareEmailTemplate "emlExpertAccount.htm", ";;sExpertFullName=" & sFullName & ";;sUserLogin=" & sTempUserLogin & ";;sUserPassword=" & sTempUserPassword & ";;sSystemUrl=" & "http://cvip.assortis.com" & sHomePath & "apply/" 
			SendEmail sEmailCvipSystem, sUserEmail, sEmailSubject, sEmailBody, "info"
			'SendEmail sEmailCvipSystem, "imc@ibf.be", sEmailSubject, sEmailBody, "info"

		End If
	End If
End Function
%>