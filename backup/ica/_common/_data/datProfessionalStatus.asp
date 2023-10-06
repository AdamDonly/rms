<%
Dim arrProfessionalStatusID()
Dim arrProfessionalStatusTitle()

Sub GetProfessionalStatusListData()
	Dim objRs
	Dim sFieldName
	Dim iProfessionalStatusCount
	iProfessionalStatusCount=0

	If sCvLanguage = cLanguageFrench Then
		sFieldName = "pfsTitleFra"
	ElseIf sCvLanguage = cLanguageSpanish Then
		sFieldName = "pfsTitleSpa"
	Else
		sFieldName = "pfsTitleEng"
	End If

	Set objRs=GetDataRecordsetSP("usp_ProfessionalStatusListSelect", Array( _
		Array(, adInteger, , Null), _
		Array(, adVarChar, 255, Null), _
		Array(, adVarChar, 80, sFieldName)))
	i=0
	iProfessionalStatusCount=objRs.RecordCount - 1
	ReDim arrProfessionalStatusID(iProfessionalStatusCount)
	ReDim arrProfessionalStatusTitle(iProfessionalStatusCount)

	' On Error Resume Next

	Do Until objRs.EOF 
		arrProfessionalStatusID(i)=objRs("id_ProfessionalStatus")
		arrProfessionalStatusTitle(i)=objRs(sFieldName)
		
		i=i+1
		objRs.MoveNext
	Loop
	objRs.Close

	Set objRs=Nothing
End Sub

GetProfessionalStatusListData
%>
