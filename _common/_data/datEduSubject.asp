<%
Dim arrEduSubjectID()
Dim arrEduSubjectTitle()

Sub GetEduSubjectListData()
	Dim objRs
	Dim sFieldName
	Dim iEduSubjectCount
	iEduSubjectCount=0

	If sCvLanguage = cLanguageFrench Then
		sFieldName = "edsDescriptionFra"
	ElseIf sCvLanguage = cLanguageSpanish Then
		sFieldName = "edsDescriptionSpa"
	Else
		sFieldName = "edsDescriptionEng"
	End If

	Set objRs=GetDataRecordsetSP("usp_EduSubjectListSelect", Array( _
		Array(, adInteger, , Null), _
		Array(, adVarChar, 255, Null), _
		Array(, adVarChar, 80, sFieldName)))
	i=0
	iEduSubjectCount=objRs.RecordCount
	ReDim arrEduSubjectID(iEduSubjectCount)
	ReDim arrEduSubjectTitle(iEduSubjectCount)

	' On Error Resume Next

	Do Until objRs.EOF 
		arrEduSubjectID(i)=objRs("id_EduSubject")
		arrEduSubjectTitle(i)=objRs(sFieldName)
		
		i=i+1
		objRs.MoveNext
	Loop
	objRs.Close

	Set objRs=Nothing
End Sub

GetEduSubjectListData
%>
