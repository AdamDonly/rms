<%
Class CProject
	' Class Fields --------------------------------------
	Public ID					' ID
	Public Name					' Internal name
	Public Reference			' Reference
	Public Title				' Official project title
	Public Status				' Status
	Public Location				' Location
	Public FundingAgency		' FundingAgency
	Public Budget
	Public Description
	Public Duration
	Public Deadline
	
	Public PublishedDate
	Public RegistredDate
	Public ValidatedDate
	Public EoiDeadline
	Public EoiSendDeadline
	Public EoiDeadlineEstimated
	Public ShortlistDate
	Public TorReceivedDate
	Public TenderDeadline
	Public TenderSendDeadline
	Public TenderDeadlineEstimated
	Public TenderResultsDate
	Public ContractStartDate
	Public ContractStartDateEstimated
	Public ContractEndDate

	' Class Initialize and Terminate --------------------
	Private Sub Class_Initialize()
		Set Status = New CStatus
		Set Budget = New CFinance
		Set Duration = New CDuration
	End Sub
	
	Private Sub Class_Terminate()
		If IsObject(Status) Then
			Set Status=Nothing
		End If
		If IsObject(Budget) Then
			Set Budget=Nothing
		End If
		If IsObject(Duration) Then
			Set Duration=Nothing
		End If
	End Sub
	
	Public Function LoadData
		If ID>0 Then LoadProjectDetailsById(ID)
	End Function
	
	Public Function LoadProjectDetailsById(AProjectId)
		LoadDataProc "usp_ProjectSelect", Array( _
			Array(, adInteger, , AProjectId), _
			Array(, adVarChar, 30, Null))
	End Function
	
	Public Function LoadProjectDetailsByReference(AReference)
		LoadDataProc "usp_ProjectSelect", Array( _
			Array(, adInteger, , Null), _
			Array(, adVarChar, 30, AReference))
	End Function
	
	Public Function LoadDataProc(AProcedure, AParams)
	Dim objTempRs
		Set objTempRs=GetDataRecordsetSPWithConn(objConn, AProcedure, AParams)
		LoadDataFromRecordset objTempRs

		objTempRs.Close
		Set objTempRs=Nothing
	End Function

	Public Function LoadDataFromRecordset(ARecordSet)
		If IsObject(ARecordSet) Then
			On Error Resume Next
			If Not ARecordSet.Eof Then
				ID=ARecordSet("id_Project")

				Reference=ARecordSet("prjReference")
				Title=ARecordSet("prjTitle")
				Name=ARecordSet("prjShortName")
				Location=ARecordSet("prjLocation")
				Description=ARecordSet("prjDescription")
				Deadline=ARecordSet("prjDeadline")

				Status.ID=CheckInteger(ARecordSet("id_ProjectStatus"))
				Status.Name=ARecordSet("prjStatus")
			End If
			On Error GoTo 0
		End If
	End Function	
	
	Public Function SaveData
		SaveData=SaveDataProcOutParams("usp_ProjectUpdate", Array( _ 
			Array(, adInteger, , ID), _
			Array(, adVarChar, 30, Reference), _
			Array(, adVarChar, 60, Name), _
			Array(, adVarChar, 400, Title), _
			Array(, adInteger, , Status.ID), _
			Array(, adVarChar, 100, Location), _
			Array(, adVarWChar, 20000, Description), _
			Array(, adVarChar, 16, Deadline)), Array( _ 
			Array(, adInteger)))
	End Function

	Public Function SaveDataProcOutParams(AProcedure, AParamsIn, AParamsOut)
	Dim objTempRs, iResult
		objTempRs=GetDataOutParamsSP(AProcedure, AParamsIn, AParamsOut)
		iResult=objTempRs(0)
		Set objTempRs=Nothing

		SaveDataProcOutParams=iResult
	End Function

	Public Function SaveDataProc(AProcedure, AParams)
	Dim objTempRs
		objTempRs=UpdateRecordSP(AProcedure, AParams)
		Set objTempRs=Nothing
	End Function	

	Public Function DeleteData
		SaveDataProc "usp_ProjectDelete", Array( _ 
			Array(, adInteger, , ID))
	End Function
	
End Class


Class CProjectList
	' Class Fields --------------------------------------
	Private FCount
	Public Item()

	' Class Initialize and Terminate --------------------
	Private Sub Class_Initialize()
	End Sub
	
	Private Sub Class_Terminate()
		Call DestoyData()
		If IsObject(Item) Then
			Set Item=Nothing
		End If		
	End Sub	

	' Class Private Functions ---------------------------
	Private Function DestoyData()
		While FCount>0 
			Set Item(FCount-1)=Nothing
			FCount=FCount-1
		WEnd
	End Function

	' Class Properties ----------------------------------
	Public Property Get Count
		Count=FCount
    End Property

	' Class Methods -------------------------------------
	Public Function LoadData
		LoadDataProc "usp_ProjectListSelect", Array( _
			Array(, adVarChar, 100, "111,112,121,122,201,202,203"), _
			Array(, adVarChar, 250, ""), _
			Array(, adVarChar, 50, ""))
	End Function

	Public Function LoadDataByStatusKeywords(AStatusList, AKeywords, AOrderBy)
		LoadDataProc "usp_ProjectListSelect", Array( _
			Array(, adVarChar, 100, AStatusList), _
			Array(, adVarChar, 250, AKeywords), _
			Array(, adVarChar, 50, AOrderBy))
	End Function
	
	Public Function LoadDataProc(AProcedure, AParams)
		Dim i
		FCount=0

		Set objTempRs=GetDataRecordsetSPWithConn(objConn, AProcedure, AParams)
		ReDim Item(objTempRs.RecordCount-1)
		If Not objTempRs.Eof Then
			objTempRs.MoveFirst
			While Not objTempRs.Eof
				Set Item(FCount) = New CProject
				Item(FCount).LoadDataFromRecordset objTempRs
				
				FCount=FCount+1
				objTempRs.MoveNext
			WEnd
		End If
		objTempRs.Close
		Set objTempRs=Nothing
	End Function

	Public Function Clear()
		DestoyData()
	End Function
	
End Class

%>