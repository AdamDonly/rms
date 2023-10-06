<%
'--------------------------------------------------------------------
'
' Link expert.
'
'--------------------------------------------------------------------
Response.Expires = 0
Response.ExpiresAbsolute = Now() - 1
Response.CacheControl = "no-cache" 
%>
<!--#include virtual="/_common/_class/main.asp"-->
<!--#include virtual="/_common/_class/expert.asp"-->
<!--#include virtual="/_common/_class/project.asp"-->
<!--#include virtual="/_common/_class/expert.project.asp"-->
<%
If sApplicationName<>"expert" Then
	CheckUserLogin sScriptFullNameAsParams
Else 
	Response.Redirect sApplicationHomePath
End If

sParams=ReplaceUrlParams(sParams, "idexpert")
sParams=ReplaceUrlParams(sParams, "idproject")
Dim iProjectID

Dim objExpertProject
Set objExpertProject = New CExpertProject

' On delete
If sAction="delete"Then
	objExpertProject.Project.ID=CheckIntegerAndZero(Request.QueryString("idproject"))
	objExpertProject.Expert.ID=CheckIntegerAndZero(Request.QueryString("idexpert"))
	objExpertProject.DeleteData
	iProjectID=objExpertProject.Project.ID
End If
' On data submit
If Len(Request.Form())>0 Then
	objExpertProject.Project.ID=CheckIntegerAndZero(Request.Form("idproject"))
	objExpertProject.Expert.ID=CheckIntegerAndZero(Request.Form("idexpert"))
	objExpertProject.Status.ID=Request.Form("expert_status")
	objExpertProject.Fee.Value=CheckSingleAndNull(Replace(Request.Form("expert_fee"), ",", "."))
	objExpertProject.Fee.CurrencyCode=Request.Form("expert_fee_currency")
	objExpertProject.ProvidedCompany=Request.Form("expert_provided_company")
	objExpertProject.Comments=Request.Form("expert_comments")
	objExpertProject.SaveData

	iProjectID=objExpertProject.Project.ID
End If

Set objExpertProject=Nothing

If (Len(sUrl)<1) Or (sUrl=sHomePath) Or (sUrl=sApplicationHomePath) Then
	If iProjectID>0 Then
		sUrl="details.asp" & ReplaceUrlParams(sParams, "idproject=" & iProjectID)
	Else
		sUrl=sApplicationHomePath & sParams
	End If
End If

Response.Redirect sUrl
%>
