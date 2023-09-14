<%
'--------------------------------------------------------------------
'
' CV registration.
' Short format.
'
'--------------------------------------------------------------------
Response.Expires = 0
Response.ExpiresAbsolute = Now() - 1
Response.CacheControl = "no-cache"
%>
<!--#include file="_data/datGender.asp"-->
<!--#include file="_data/datPsnTitle.asp"-->
<!--#include file="_data/datMonth.asp"-->
<!--#include virtual="/fnc_file.asp"-->
<%
' Remove inactive url params
sParams=ReplaceUrlParams(sParams, "url")

If sApplicationName<>"expert" Then
	CheckUserLogin sScriptFullNameAsParams
End If
%>
<!--#include virtual="/_common/expProfile.asp"-->
<%
Dim objResult, iResult, objUpload, objForm, objField
Dim sUserLogin, sUserPassword, sUserPhone
Dim sFileName, sFileExtension, sFullFileName, Msg, objFile

If sAction="push" Then
	Set objForm=Request.QueryString()
	%>
<!--#include virtual="/_common/verify_push.asp"-->
	<%
Else
	Set objUpload = Server.CreateObject("softartisans.fileup")
	If objUpload.ContentDisposition = "form-data" Then
		Set objForm=objUpload.Form

		If objUpload.TotalBytes >  2800000 Then 
			Msg = 1
			ShowMessageStart "error", 580 %>
				<p>Your file is too big.</b><br />Please try to keep the size of the file within the allowed 2.5 Mb. Click back and try again.</p>
			<% ShowMessageEnd 
			Response.End
		Else
		
			if objForm("cvEng").TotalBytes>0 then
				sFileName=Trim(objForm("cvEng").UserFilename)
				sFileExtension=GetFileExtension(sFileName)

				If sFileExtension="doc" or sFileExtension="txt" or sFileExtension="rtf" or sFileExtension="pdf" or sFileExtension="htm" or sFileExtension="html" or sFileExtension="docx" then 
					Msg=0
				Else
					Msg=1
					ShowMessageStart "error", 580 %>
						<p>Your CV file has an unknown extension.</b><br />Please try to upload this file in MS Word document format. Thank you.<br />Click back and try again.</p>
					<% ShowMessageEnd 
					Response.End
				End If 
			End If 
		End If

		Set objFile=Server.CreateObject("Scripting.FileSystemObject")
		' Saving the file
		sFileName=objUserCompanyDB.Database & "_" & Replace(ConvertDateForSQL(Now),"/","") & "_" & Mid(sSessionID, 26, 9) & "." & sFileExtension
		sFullFileName=Server.Mappath("../../../_upload/ica") & "\" & sFileName

		If objFile.FileExists(sFullFileName) then
			If objFile.FileExists(sFullFileName & "_") then    
				objFile.DeleteFile sFullFileName & "_"
			End If 
			objFile.MoveFile sFullFileName, sFullFileName & "_"
		End If    
		objForm("cvEng").SaveAs sFullFileName
	End If
		
	Set objResult=VerifyExpertProfile(objForm)
	' If there is no expert with similar contact details
	If objResult.Eof Then
		objResult.Close
		Set objResult=Nothing
		%>
<!--#include virtual="/_common/verify_push.asp"-->
		<%
	End If
	' If there are experts with similar contact details
	%>
<!--#include virtual="/_common/verify_list.asp"-->
	<%	
	Set objResult=Nothing
End If
Set objUpload=Nothing
%>
