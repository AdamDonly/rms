<%
'--------------------------------------------------------------------
'
' Database connector + Application initialization.
' Maintain the application and sessions states.
'
'--------------------------------------------------------------------
%>
<!--#include file="__init.asp"-->
<!--#include file="_ado.asp"-->
<!--#include file="_dal.asp"-->
<% 
Dim objConn, objTempRs, objTempRs2, objTempRs3, objTempRsLog

' Creating ADO connection
Set objConn = Server.CreateObject("ADODB.Connection")
objConn.Open "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Data Source=MAGNESA;Initial Catalog=rms_ica;"

' Procedure for closing all database connections
Sub CloseDBConnection
	Set objTempRs=Nothing
	objconn.Close
	Set objconn=Nothing
End Sub

Dim i
%>
<!--#include file="_urlparams.asp"-->
<!--#include file="_session.asp"-->
<!--#include file="_checks.asp"-->
<%
Dim sApplicationName, sApplicationHomePath
If InStr(LCase(sScriptFullName), "/apply/")>0 Then 
	sApplicationName="expert"
	sApplicationHomePath=sHomePath & "apply/"
ElseIf InStr(LCase(sScriptFullName), "/external/")>0 Then 
	sApplicationName="external"
	sApplicationHomePath=sHomePath & "external/"
ElseIf InStr(LCase(sScriptFullName), "/outsourcing/")>0 Then 
	sApplicationName="outsourcing"
	sApplicationHomePath=sHomePath & "outsourcing/"
ElseIf InStr(LCase(sScriptFullName), "/backoffice/")>0 Or InStr(LCase(sScriptFullName), "/_backoffice/")>0 Then 
	sApplicationName="backoffice"
	sApplicationHomePath=sHomePath & "backoffice/"
Else
	sApplicationName=""
	sApplicationHomePath=sHomePath
End If
%>
<!--#include file="__ica.asp"-->
<!--#include file="_email.asp"-->
<!--#include file="fnc_email.asp"-->

<!--#include file="_cvid.asp"-->
