<%
Dim iLanguageCount
iLanguageCount=0
Dim arrLanguageID()
Dim arrLanguageTitle()

Sub GetLanguageListData()
	Dim objRs
	Dim sFieldName

	If sCvLanguage = cLanguageFrench Then
		sFieldName = "lngNameFra"
	ElseIf sCvLanguage = cLanguageSpanish Then
		sFieldName = "lngNameSpa"
	Else
		sFieldName = "lngNameEng"
	End If

	Set objRs=GetDataRecordsetSP("usp_LanguageListSelect", Array( _
		Array(, adInteger, , Null), _
		Array(, adVarChar, 255, Null), _
		Array(, adInteger, , Null)))
	i=1
	iLanguageCount=objRs.RecordCount
	ReDim arrLanguageID(iLanguageCount)
	ReDim arrLanguageTitle(iLanguageCount)

	' On Error Resume Next

	Do Until objRs.EOF 
		arrLanguageID(i)=objRs("id_Language")

		arrLanguageTitle(i)=objRs(sFieldName)
		
		i=i+1
		objRs.MoveNext
	Loop
	objRs.Close

	Set objRs=Nothing
End Sub

GetLanguageListData
%>
