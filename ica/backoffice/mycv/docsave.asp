<%@ LANGUAGE="VBSCRIPT" %>
<% Option Explicit %>

<!--#include file="../../dbc.asp"-->
<!--#include file="../../fnc.asp"-->
<!--#include file="../../fnc_log.asp"-->
<!--#include file="../../_forms/frmInterface.asp"-->
<!--#include file="../../fnc_exp.asp"-->

<%
'--------------------------------------------------------------------
'
' Doc upload.
'
'--------------------------------------------------------------------
Response.Expires = 0
Response.ExpiresAbsolute = Now() - 1
Response.CacheControl = "no-cache"
%>
<!--#include virtual="/fnc_file.asp"-->
<!--#include file="../../_common/cv_data.asp"-->
<%
' Remove inactive url params
sParams=ReplaceUrlParams(sParams, "url")

If sApplicationName<>"expert" Then
	CheckUserLogin sScriptFullNameAsParams
End If

Dim objResult, iResult, objUploadForm, objField
Dim sUserLogin, sUserPassword, sUserPhone
Dim sFileName, sFileExtension, sFullFileName, Msg, myfile

Set objUploadForm = Server.CreateObject("softartisans.fileup")
If objUploadForm.contentdisposition = "form-data" Then
	If objUploadForm.totalbytes >  2800000 Then 
		ShowMessageStart "error", 580 %>
			<p>Your file is too big.</b><br />Please try to keep the size of the file within the allowed 1.5 Mb. Click back and try again.</p>
		<% ShowMessageEnd 
		Msg = 1
	Else
		if objUploadForm.Form("cvEng").TotalBytes>0 then
			sFileName=Trim(objUploadForm.form("cvEng").UserFilename)
			sFileExtension=GetFileExtension(sFileName)

			If sFileExtension="doc" or sFileExtension="txt" or sFileExtension="rtf" or sFileExtension="pdf" or sFileExtension="htm" or sFileExtension="html" or sFileExtension="docx" then 
					Msg=0
			Else
				Msg=1
        		ShowMessageStart "error", 580 %>
				<p>Your CV file has an unknown extension.</b><br />Please try to upload this file in MS Word document format. Thank you.<br />Click back and try again.</p>
				<% ShowMessageEnd 
			End If 
		End If 
	End If
End If

' save the file on the filesystem:
If Msg = 0 then

	Set myfile=Server.CreateObject("scripting.filesystemobject")
	' Saving the file
	sFileName=objUserCompanyDB.Database & "_" & Replace(ConvertDateForSQL(Now),"/","") & "_" & Mid(sSessionID, 26, 9) & "." & sFileExtension
	sFullFileName=Server.Mappath("../../../_upload/ica") & "\" & sFileName

	if myfile.FileExists(sFullFileName) then
		If myfile.FileExists(sFullFileName & "_") then    
			myfile.DeleteFile sFullFileName & "_"
		End If 
		myfile.MoveFile sFullFileName, sFullFileName & "_"
	end if    
	objUploadForm.form("cvEng").SaveAs sFullFileName
		
	' add doc in the DB:

End If
Set objUploadForm=Nothing

If Msg = 0 Then
	Response.Redirect Response.Redirect sScriptFileName
End If
%>

