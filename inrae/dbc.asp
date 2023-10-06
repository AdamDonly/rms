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
Dim objConn, objTempRs, objTempRs2

' Creating ADO connection
Set objConn = Server.CreateObject("ADODB.Connection")
objConn.Open "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Data Source=MAGNESA;Initial Catalog=rms_inrae;"

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
ElseIf InStr(LCase(sScriptFullName), "/expert/")>0 Then 
	sApplicationName="expert"
	sApplicationHomePath=sHomePath & "expert/"
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
<!--#include file="_email.asp"-->
<!--#include file="fnc_email.asp"-->

<% 
	Sub DebugMessageLine(str)
		Response.Write("[Debug]:" & str & "<br>")
	End Sub

	Sub DebugMessageEnd(str)
		Response.Write("[DEBUG]: " & str)
		Response.End
	End Sub
%>

<%
If InStr(sScriptServerName, "cvip2")>0 And InStr(sScriptFileName, "save")=0 Then Response.Write "CVIP2"
%>
<!--#include file="_cvid.asp"-->
