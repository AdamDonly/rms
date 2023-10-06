<%
Class CExpertProject
	Public Expert
	Public Project
	Public Status
	Public Fee
	Public ProvidedCompany
	Public ProvidedPerson
	Public Comments
	Public ProvidedDate

	
	' Class Initialize and Terminate --------------------
	Private Sub Class_Initialize()
		Set Expert = New CExpert
		Set Project = New CProject
		Set Status = New CStatus
		Set Fee = New CFinance
	End Sub
	
	Private Sub Class_Terminate()
		If IsObject(Expert) Then
			Set Expert=Nothing
		End If
		If IsObject(Project) Then
			Set Project=Nothing
		End If
		If IsObject(Status) Then
			Set Status=Nothing
		End If
		If IsObject(Fee) Then
			Set Fee=Nothing
		End If
	End Sub
	
	Public Function LoadData
		If Expert.ID>0 And Project.ID>0 Then LoadExpertProjectDetailsById Expert.ID, Project.ID
	End Function
	
	Public Function LoadExpertProjectDetailsById(AExpertID, AProjectId)
		LoadDataProc "usp_ExpertProjectSelect", Array( _
			Array(, adInteger, , AExpertId), _
			Array(, adInteger, , AProjectId))
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
				Expert.ID=CheckIntegerAndZero(ARecordSet("id_Expert"))
				Project.ID=CheckIntegerAndZero(ARecordSet("id_Project"))
			
				Status.ID=CheckInteger(ARecordSet("id_ExpertStatus"))
				Status.Name=ARecordSet("exsTitle")

				ProvidedCompany=ARecordSet("epjProvidedCompany")
				ProvidedPerson=ARecordSet("epjProvidedPerson")
				Fee.Value=ARecordSet("epjFee")
				Fee.CurrencyCode=ARecordSet("epjFeeCurrency")
				Comments=ARecordSet("epjComments")
				ProvidedDate=ARecordSet("epjCreateDate")
				
			End If
			On Error GoTo 0
		End If
	End Function
	
	Public Function SaveData
		SaveDataProc "usp_ExpertProjectUpdate", Array( _ 
			Array(, adInteger, , Expert.ID), _
			Array(, adInteger, , Project.ID), _
			Array(, adInteger, , Status.ID), _
			Array(, adVarChar, 400, ProvidedCompany), _
			Array(, adVarChar, 400, ProvidedPerson), _
			Array(, adSingle, , Fee.Value), _
			Array(, adVarChar, 3, Fee.CurrencyCode), _
			Array(, adVarWChar, 20000, Comments))
	End Function

	Public Function SaveDataProc(AProcedure, AParams)
	Dim objTempRs
		objTempRs=UpdateRecordSP(AProcedure, AParams)
		Set objTempRs=Nothing
	End Function

	Public Function DeleteData
		SaveDataProc "usp_ExpertProjectDelete", Array( _ 
			Array(, adInteger, , Expert.ID), _
			Array(, adInteger, , Project.ID))
	End Function
	
End Class

Class CExpertProjectList
	' Class Fields --------------------------------------
	Public Expert
	Private FCount
	Public Item()

	' Class Initialize and Terminate --------------------
	Private Sub Class_Initialize()
		Set Expert = New CExpert
	End Sub
	
	Private Sub Class_Terminate()
		Call DestoyData()
		If IsObject(Item) Then
			Set Item=Nothing
		End If		
		If IsObject(Expert) Then
			Set Expert=Nothing
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
		LoadDataProc "usp_ExpertProjectListSelect", Array( _
			Array(, adInteger, , Expert.ID), _
			Array(, adVarChar, 100, ""), _
			Array(, adVarChar, 250, ""), _
			Array(, adVarChar, 50, ""))
	End Function

	Public Function LoadDataProc(AProcedure, AParams)
		Dim i
		FCount=0

		Set objTempRs=GetDataRecordsetSPWithConn(objConn, AProcedure, AParams)
		ReDim Item(objTempRs.RecordCount-1)
		If Not objTempRs.Eof Then
			objTempRs.MoveFirst
			While Not objTempRs.Eof
				Set Item(FCount) = New CExpertProject
				Item(FCount).LoadDataFromRecordset objTempRs
				
				FCount=FCount+1
				objTempRs.MoveNext
			WEnd
		End If
		objTempRs.Close
		Set objTempRs=Nothing
	End Function

End Class

%>