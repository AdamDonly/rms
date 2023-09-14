<%
'--------------------------------------------------------------------
'
' Project registration.
'
'--------------------------------------------------------------------
Response.Expires = 0
Response.ExpiresAbsolute = Now() - 1
Response.CacheControl = "no-cache" 
%>
<!--#include virtual="/_common/_class/main.asp"-->
<!--#include virtual="/_common/_class/project.asp"-->
<%
If sApplicationName<>"expert" Then
	CheckUserLogin sScriptFullNameAsParams
Else 
	Response.Redirect sApplicationHomePath
End If

sParams=ReplaceUrlParams(sParams, "project_type")
sParams=ReplaceUrlParams(sParams, "idproject")
Dim iProjectID

Dim objProject
Set objProject = New CProject

' On delete
If sAction="delete"Then
	objProject.ID=CheckIntegerAndZero(Request.QueryString("idproject"))
	objProject.DeleteData
End If
' On data submit
If Len(Request.Form())>0 Then
	objProject.ID=CheckIntegerAndZero(Request.Form("idproject"))
	objProject.Title=Request.Form("project_title")
	objProject.Name=Request.Form("project_shortname")
	objProject.Reference=Request.Form("project_reference")
	objProject.Location=Request.Form("project_country")
	objProject.Status.ID=Request.Form("project_status")
	objProject.Description=Request.Form("project_description")
	objProject.Deadline=ConvertDMYForSql(Request.Form("project_deadline_year"), Request.Form("project_deadline_month"), Request.Form("project_deadline_day"))
	iProjectID=objProject.SaveData
End If

Set objProject = Nothing

If (Len(sUrl)<1) Or (sUrl=sHomePath) Or (sUrl=sApplicationHomePath) Then
	If iProjectID>0 Then
		sUrl="details.asp" & ReplaceUrlParams(sParams, "idproject=" & iProjectID)
	Else
		sUrl=sApplicationHomePath & sParams
	End If
End If

Response.Redirect sUrl
%>
