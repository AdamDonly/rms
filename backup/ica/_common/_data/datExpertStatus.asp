<%
Dim arrExpertStatusID()
Dim arrExpertStatusTitle()

Sub GetExpertStatusListData()
	Dim objRs
	Dim sFieldName
	Dim iExpertStatusCount
	iExpertStatusCount=0

	If sCvLanguage = cLanguageFrench Then
		sFieldName = "exsTitleFra"
	ElseIf sCvLanguage = cLanguageSpanish Then
		sFieldName = "exsTitleSpa"
	Else
		sFieldName = "exsTitleEng"
	End If

	Set objRs=GetDataRecordsetSP("usp_ExpertStatusListSelect", Array( _
		Array(, adInteger, , Null), _
		Array(, adVarChar, 255, Null), _
		Array(, adVarChar, 80, sFieldName)))
	i=0
	iExpertStatusCount=objRs.RecordCount - 1
	ReDim arrExpertStatusID(iExpertStatusCount)
	ReDim arrExpertStatusTitle(iExpertStatusCount)

	' On Error Resume Next

	Do Until objRs.EOF 
		arrExpertStatusID(i)=objRs("id_ExpertStatus")
		arrExpertStatusTitle(i)=objRs(sFieldName)
		
		i=i+1
		objRs.MoveNext
	Loop
	objRs.Close

	Set objRs=Nothing
End Sub

GetExpertStatusListData
%>
