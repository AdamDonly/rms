<%
' Get CV ID
Dim iCvUID, iOriginalCvID
iCvUID=Request.QueryString("uid")

GetCvIdDetails(iCvUID)

If iOriginalCvID>0 Then Response.Redirect(sScriptFileName & ReplaceUrlParams(sParams, "uid=" & iOriginalCvID))

Function GetCvIdDetails(ACvUid)
	Dim objRs

	Set objRs=GetDataRecordsetSP("usp_Ica_ExpertUidDetailsSelect", Array( _
		Array(, adVarchar, 40, ACvUid)))

	If Not objRs.EOF Then
		iOriginalCvID = objRs("id_ExpertOriginal")
		sCvLanguage = ReplaceIfEmpty(objRs("Lng"), sDefaultCvLanguage)
		ForceCvLanguage()
	End If

	Set GetCvIdDetails = objRs
End Function
%>