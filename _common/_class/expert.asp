<%
Class CExpert
	Public ID
	Public FirstName
	Public LastName
	Public FullName

	Public Function LoadData
		If ID>0 Then LoadProjectDetailsById(ID)
	End Function
	
	Public Function LoadProjectDetailsById(AExpertID)
		LoadDataProc "usp_ProjectSelect", Array( _
			Array(, adInteger, , AExpertID))
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
	
End Class


%>
