<%
Class CExpertStatusCV
	Public Expert
	Public Status
	Public DateModified
	
	' Class Initialize and Terminate --------------------
	Private Sub Class_Initialize()
		Set Expert = New CExpert
		Set Status = New CStatus
	End Sub
	
	Private Sub Class_Terminate()
		If IsObject(Expert) Then
			Set Expert=Nothing
		End If
		If IsObject(Status) Then
			Set Status=Nothing
		End If
	End Sub	
	
	' Class Methods -------------------------------------
	Public Function LoadData
		LoadDataProc "usp_ExpertStatusCVSelect", Array( _
				Array(, adInteger, , Expert.ID))	
	End Function
	
	Public Function LoadDataProc(AProcedure, AParams)
		Dim objTempRs
		Set objTempRs=GetDataRecordsetSP(AProcedure, AParams)
		If Not objTempRs.Eof Then
			Status.ID=CheckInteger(objTempRs("id_Status"))
			Status.Name=objTempRs("stsNameEng")
		End If
		objTempRs.Close
		Set objTempRs=Nothing
	End Function
	
	Public Function SaveData
		SaveDataProc "usp_ExpertStatusCVUpdate", Array( _
				Array(, adInteger, , Expert.ID), _
				Array(, adInteger, , Status.ID), _
				Array(, adVarChar, 16, ConvertDateForSql(Now)))
	End Function
	
	Public Function SaveDataProc(AProcedure, AParams)
		Dim objTempRs
		objTempRs=UpdateRecordSP(AProcedure, AParams)
	End Function
	
End Class


%>
