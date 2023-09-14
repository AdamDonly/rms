<%
	' Create a new profile
	objResult=SaveExpertShortProfile(0, objForm)

	If objResult(0)=0 Then
		If objResult(1)>0 Then
			iExpertID=objResult(1)
		End If
	End If
	Set objResult=Nothing

	' If language link is specified - create a record in ExpertsLanguage
	iLanguageLinkID = CheckIntegerAndZero(Request.QueryString("lnglink"))
	If iLanguageLinkID > 0 Then
		SaveExpertCvLanguageLink iLanguageLinkID, iExpertID
		CopyExpertCvLanguage iLanguageLinkID, iExpertID
	End If

	' Redirect to edit newly created profile
	sUrl=ReplaceUrlParams(sUrl, "act")
	sUrl=ReplaceUrlParams(sUrl, "url")
	For Each objField In objForm 
		sUrl=ReplaceUrlParams(sUrl, objField)
	Next
	sParams=ReplaceUrlParams(sParams, "id=" & iExpertID)
	Response.Redirect sUrl & sParams
%>