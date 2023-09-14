<%
'--------------------------------------------------------------------
'
' Project registration.
'
'--------------------------------------------------------------------
Response.Expires = 0
Response.ExpiresAbsolute = Now() - 1
Response.CacheControl = "no-cache" 

If sApplicationName<>"expert" Then
	CheckUserLogin sScriptFullNameAsParams
Else 
	Response.Redirect sApplicationHomePath
End If

sParams=ReplaceUrlParams(sParams, "project_type")

' On data submit
If Len(Request.Form())>0 Then
Dim iProjectID, sProjectTitle, sProjectShortName, sProjectReference, sProjectLocation, iProjectStatusID, sProjectDescription, sProjectDeadline


	iProjectID=CheckIntegerAndZero(Request.Form("idproject"))
	sProjectTitle=Request.Form("project_title")
	sProjectShortName=Request.Form("project_shortname")
	sProjectReference=Request.Form("project_reference")
	sProjectLocation=Request.Form("project_country")
	iProjectStatusID=Request.Form("project_status")
	sProjectDescription=Request.Form("project_description")
	sProjectDeadline=ConvertDMYForSql(Request.Form("project_deadline_year"), Request.Form("project_deadline_month"), Request.Form("project_deadline_day"))
	
	objTempRs=GetDataOutParamsSP("usp_ProjectUpdate", Array( _ 
		Array(, adInteger, , iProjectID), _
		Array(, adVarChar, 30, sProjectReference), _
		Array(, adVarChar, 60, sProjectShortName), _
		Array(, adVarChar, 400, sProjectTitle), _
		Array(, adInteger, , iProjectStatusID), _
		Array(, adVarChar, 100, sProjectLocation), _
		Array(, adVarWChar, 20000, sProjectDescription), _
		Array(, adVarChar, 16, sProjectDeadline)), Array( _ 
		Array(, adInteger)))
		
	iProjectID=objTempRs(0)
	Set objTempRs=Nothing
End If

If Len(sUrl)<1 Then
	If iProjectID>0 Then
		sUrl="details.asp" & ReplaceUrlParams(sParams, "idproject=" & iProjectID)
	Else
		sUrl=sApplicationHomePath & sParams
	End If
End If

Response.Redirect sUrl
%>
