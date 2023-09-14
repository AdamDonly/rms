<%
Dim arrCountryID()
Dim arrCountryName()
Dim arrCountryEU()

Sub GetCountryListData()
	Dim objRs
	Dim sFieldName
	Dim iCountryCount
	iCountryCount=0

	If sCvLanguage = cLanguageFrench Then
		sFieldName = "couNameFra"
	ElseIf sCvLanguage = cLanguageSpanish Then
		sFieldName = "couNameSpa"
	Else
		sFieldName = "couNameEng"
	End If

	Set objRs=GetDataRecordsetSP("usp_CountryListSelect", Array( _
		Array(, adInteger, , Null), _
		Array(, adVarChar, 255, Null), _
		Array(, adInteger, , Null), _
		Array(, adVarChar, 255, Null), _
		Array(, adVarChar, 80, sFieldName)))
	i=0
	iCountryCount=objRs.RecordCount
	ReDim arrCountryID(iCountryCount)
	ReDim arrCountryName(iCountryCount)
	ReDim arrCountryEU(iCountryCount)

	' On Error Resume Next

	Do Until objRs.EOF 
		arrCountryID(i)=objRs("id_Country")

		arrCountryName(i)=objRs(sFieldName)
		arrCountryEU(i)=0
		
		i=i+1
		objRs.MoveNext
	Loop
	objRs.Close

	Set objRs=Nothing
End Sub

GetCountryListData
%>
