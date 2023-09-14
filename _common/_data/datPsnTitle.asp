<%
Dim arrPersonTitleID()
Dim arrPersonTitle()

Sub GetPsnTitleListData()
	Dim objRs
	Dim sFieldName
	Dim iPsnTitleCount
	iPsnTitleCount=0

	If sCvLanguage = cLanguageFrench Then
		sFieldName = "ptlNameFra"
	ElseIf sCvLanguage = cLanguageSpanish Then
		sFieldName = "ptlNameSpa"
	Else
		sFieldName = "ptlNameEng"
	End If

	Set objRs=GetDataRecordsetSP("usp_PersonTitleListSelect", Array( _
		Array(, adInteger, , Null), _
		Array(, adVarChar, 255, Null), _
		Array(, adVarChar, 80, Null)))
	i=1
	iPsnTitleCount=objRs.RecordCount
	ReDim arrPersonTitleID(iPsnTitleCount)
	ReDim arrPersonTitle(iPsnTitleCount)

	' On Error Resume Next

	Do Until objRs.EOF 
		arrPersonTitleID(i)=objRs("id_psnTitle")
		arrPersonTitle(i)=objRs(sFieldName)
		
		i=i+1
		objRs.MoveNext
	Loop
	objRs.Close

	Set objRs=Nothing
End Sub

GetPsnTitleListData
%>
