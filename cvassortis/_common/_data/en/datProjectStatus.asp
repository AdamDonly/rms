<% 
Dim arrProjectStatusTitle, arrProjectStatusID

Sub LoadProjectStatus
Dim objTempRs
Dim iRecordCount, iLoop

	Set objTempRs=GetDataRecordsetSP("usp_ProjectStatusListSelect", Array())

	iRecordCount=objTempRs.RecordCount
	ReDim arrProjectStatusTitle(iRecordCount-1)
	ReDim arrProjectStatusID(iRecordCount-1)

	iLoop=0
	While Not objTempRs.Eof
		arrProjectStatusTitle(iLoop)=objTempRs("prsTitle")
		arrProjectStatusID(iLoop)=objTempRs("id_ProjectStatus")
		
		iLoop=iLoop+1
		objTempRs.MoveNext
	WEnd
	objTempRs.Close
	Set objTempRs=Nothing
End Sub

LoadProjectStatus

%>