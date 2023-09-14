<% 
Dim arrProfessionalStatusTitle, arrProfessionalStatusID

Sub LoadProjectStatus
Dim objTempRs
Dim iRecordCount, iLoop

	Set objTempRs=GetDataRecordsetSP("usp_ProfessionalStatusListSelectFra", Array())

	iRecordCount=objTempRs.RecordCount
	ReDim arrProfessionalStatusTitle(iRecordCount-1)
	ReDim arrProfessionalStatusID(iRecordCount-1)

	iLoop=0
	While Not objTempRs.Eof
		arrProfessionalStatusTitle(iLoop)=objTempRs("pfsTitleFra")
		arrProfessionalStatusID(iLoop)=CInt(objTempRs("id_ProfessionalStatus"))
		
		iLoop=iLoop+1
		objTempRs.MoveNext
	WEnd
	objTempRs.Close
	Set objTempRs=Nothing
End Sub

LoadProjectStatus

%>
