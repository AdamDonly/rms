<% 
'--------------------------------------------------------------------
'
' CV registration.
' Removing expert from the database
'
'--------------------------------------------------------------------
%>
<!--#include file="../cv_data.asp"-->
<% 
CheckUserLogin sScriptFullNameAsParams
CheckExpertID()

sParams=ReplaceUrlParams(sParams, "id")
sParams=ReplaceUrlParams(sParams, "idproject")
sParams=ReplaceUrlParams(sParams, "idexpert")

If iExpertID>0 Then
	ShowStandardPageHeader	

	objTempRs=GetDataOutParamsSP("usp_AdmExpRestore", Array( _
		Array(, adInteger, , iExpertID)), Array( _ 
		Array(, adInteger)))
	
	If objTempRs(0)>=1 Then
		Response.Write "<br><br><br><br><p align=""center"">The CV of the expert with ID " & iExpertID & " was successfully restored."
	End If
	%><br><br>
	<a href="<% =sApplicationHomePath %>register/register6.asp<% =ReplaceUrlParams(sParams, "id=" & iExpertID) %>"><img src="<% =sHomePath %>image/bte_continue.gif" border=0></a>
	<%
	ShowStandardPageFooter
	Response.End
End If
%>
