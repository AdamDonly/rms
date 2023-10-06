<% 
Dim arrExpertStatusTitle, arrExpertStatusID

Sub LoadProjectStatus
Dim objTempRs
Dim iRecordCount, iLoop

	Set objTempRs=GetDataRecordsetSP("usp_ExpertStatusListSelect", Array())

	iRecordCount=objTempRs.RecordCount
	ReDim arrExpertStatusTitle(iRecordCount-1)
	ReDim arrExpertStatusID(iRecordCount-1)

	iLoop=0
	While Not objTempRs.Eof
		arrExpertStatusTitle(iLoop)=objTempRs("exsTitle")
		arrExpertStatusID(iLoop)=objTempRs("id_ExpertStatus")
		
		iLoop=iLoop+1
		objTempRs.MoveNext
	WEnd
	objTempRs.Close
	Set objTempRs=Nothing
End Sub

LoadProjectStatus

%>
