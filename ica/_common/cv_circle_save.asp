<%
If (bCvValidForMemberOrExpert = 1 Or bCvValidForMemberOrExpert = 5) Then

	Dim bIsCompanyCircleExpert
	bIsCompanyCircleExpert = GetExpertCompanyCircleByUid(sCvUID, iUserCompanyID, iUserID)

	If bIsCompanyCircleExpert And sAction = "remove" Then
		RemoveExpertCompanyCircle sCvUID, iUserCompanyID, iUserID
	End If

	If Not bIsCompanyCircleExpert Then
		SaveExpertCompanyCircle sCvUID, iUserCompanyID, iUserID
	End If
End If

Sub SaveExpertCompanyCircle(AExpertUid, ACompanyID, AUserID)
		Dim objTempRs2
		objTempRs2 = UpdateRecordSP("usp_" & sIcaServerSqlPrefix & "ExpertCompanyCircleUpdate", Array( _
			Array(, adVarChar, 40, AExpertUid), _
			Array(, adInteger, , ACompanyID), _
			Array(, adInteger, , AUserID)))
		Set objTempRs2 = Nothing
End Sub

Sub RemoveExpertCompanyCircle(AExpertUid, ACompanyID, AUserID)
		Dim objTempRs2
		objTempRs2 = UpdateRecordSP("usp_" & sIcaServerSqlPrefix & "ExpertCompanyCircleDelete", Array( _
			Array(, adVarChar, 40, AExpertUid), _
			Array(, adInteger, , ACompanyID), _
			Array(, adInteger, , AUserID)))
		Set objTempRs2 = Nothing

		' Response.Redirect "/backoffice/circle.asp"
		Response.Redirect sIcaServerProtocol & sIcaServer & "/Intranet/RemoveTopExpert?uid=" & AExpertUid
End Sub
%>
