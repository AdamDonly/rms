<%@ LANGUAGE="VBSCRIPT" %>
<% Option Explicit %>
<!--#include file="dbc.asp"-->
<!--#include file="fnc.asp"-->
<!--#include file="fnc_exp.asp"-->
<!--#include file="_forms/frmInterface.asp"-->
<!--#include file="../_common/_class/main.asp"-->
<!--#include file="../_common/_class/project.asp"-->
<%
'Response.Redirect "backoffice/" & sParams
' Check UserID 
CheckUserLogin sScriptFullNameAsParams

If (iUserAccessMaskExperts And aUserAccessMaskView) Then
	Response.Redirect "backoffice/search/exp_search.asp" & sParams
ElseIf (iUserAccessMaskExperts And aUserAccessMaskEdit) Then
	Response.Redirect "backoffice/manage.asp" & sParams
ElseIf (iUserAccessMaskExperts=aUserAccessMaskNoAccess) Then
	Response.Redirect "backoffice/noaccess.asp" & sParams
End If
%>
