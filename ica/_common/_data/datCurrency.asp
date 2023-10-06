<%
Dim arrCurrencyID()
Dim arrCurrencyCode()
Dim arrCurrencyTitle()

Sub GetCurrencyListData()
	Dim objRs
	Dim sFieldName
	Dim iCurrencyCount
	iCurrencyCount=0

	If sCvLanguage = cLanguageFrench Then
		sFieldName = "curDescriptionFra"
	ElseIf sCvLanguage = cLanguageSpanish Then
		sFieldName = "curDescriptionSpa"
	Else
		sFieldName = "curDescriptionEng"
	End If

	Set objRs=GetDataRecordsetSP("usp_CurrencyListSelect", Array( _
		Array(, adInteger, , Null), _
		Array(, adVarChar, 255, Null), _
		Array(, adVarChar, 80, sFieldName)))
	i=1
	iCurrencyCount=objRs.RecordCount
	ReDim arrCurrencyID(iCurrencyCount)
	ReDim arrCurrencyCode(iCurrencyCount)
	ReDim arrCurrencyTitle(iCurrencyCount)

	' On Error Resume Next

	Do Until objRs.EOF 
		arrCurrencyID(i)=objRs("id_Currency")
		arrCurrencyCode(i)=objRs("curAbbreviation")
		arrCurrencyTitle(i)=objRs(sFieldName)
		
		i=i+1
		objRs.MoveNext
	Loop
	objRs.Close

	Set objRs=Nothing
End Sub

GetCurrencyListData
%>
