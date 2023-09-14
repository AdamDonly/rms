<%
	' Create a new profile
	objResult = SaveExpertShortProfile(objUserCompanyDB.DatabaseCode, 0, objForm)

	If objResult(0)=0 Then
		If objResult(1)>0 Then
			iExpertID=objResult(1)
		End If
	End If
	Set objResult=Nothing

	' If language link is specified - create a record in ExpertsLanguage
	iLanguageLinkID = CheckIntegerAndZero(Request.QueryString("lnglink"))

	If iLanguageLinkID > 0 Then
		SaveExpertCvLanguageLink objExpertDB.DatabaseCode, iLanguageLinkID, iExpertID, Request.QueryString("exp_language") 
		CopyExpertCvLanguage objExpertDB.DatabaseCode, iLanguageLinkID, iExpertID
	End If

	If iExpertID>0 Then
		'On Error Resume Next
			Set objTempRs=GetDataRecordsetSP("usp_Ica_ExpertUidSelect", Array( _
				Array(, adInteger, , objUserCompanyDB.ID), _
				Array(, adInteger, , iExpertID)))

			If Err.Number<>0 Or objTempRs.Eof Then
				'Response.Redirect "/"
			End If
			sExpertUid=objTempRs("uid_Expert")
		Set objTempRs=Nothing
		'On Error GoTo 0
	End If
	
	' Redirect to edit newly created profile
	sUrl=ReplaceUrlParams(sUrl, "act")
	sUrl=ReplaceUrlParams(sUrl, "url")
	For Each objField In objForm 
		sUrl=ReplaceUrlParams(sUrl, objField)
	Next

	sParams=ReplaceUrlParams(sParams, "uid=" & sExpertUid)

	Response.Redirect sUrl & sParams
%>