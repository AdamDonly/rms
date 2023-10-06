<%
' Get CV ID
Dim iCvID, iOriginalCvID
iCvID=CheckIntegerAndZero(Request.QueryString("id"))

GetCvIdDetails(iCvID)

If iOriginalCvID>0 Then Response.Redirect(sScriptFileName & ReplaceUrlParams(sParams, "id=" & iOriginalCvID))

Function GetCvIdDetails(ACvID)
	Dim objRs

	Set objRs=GetDataRecordsetSP("usp_ExpertIdDetailsSelect", Array( _
		Array(, adInteger, , ACvID)))

	If Not objRs.EOF Then
		iOriginalCvID = objRs("id_ExpertOriginal")
		sCvLanguage = ReplaceIfEmpty(objRs("Lng"), sDefaultCvLanguage)
		ForceCvLanguage()
	End If

	Set GetCvIdDetails = objRs
End Function
%>