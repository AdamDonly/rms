<%
Dim arrEduTypeID()
Dim arrEduTypeTitle()

Sub GetEduTypeListData()
	Dim objRs
	Dim sFieldName
	Dim iEduTypeCount
	iEduTypeCount=0

	If sCvLanguage = cLanguageFrench Then
		sFieldName = "edtDescriptionFra"
	ElseIf sCvLanguage = cLanguageSpanish Then
		sFieldName = "edtDescriptionSpa"
	Else
		sFieldName = "edtDescriptionEng"
	End If

	Set objRs=GetDataRecordsetSP("usp_EduTypeListSelect", Array( _
		Array(, adInteger, , Null), _
		Array(, adVarChar, 255, Null), _
		Array(, adVarChar, 80, sFieldName)))
	i=0
	iEduTypeCount=objRs.RecordCount - 1
	ReDim arrEduTypeID(iEduTypeCount)
	ReDim arrEduTypeTitle(iEduTypeCount)

	' On Error Resume Next

	Do Until objRs.EOF 
		arrEduTypeID(i)=objRs("id_EduType")
		arrEduTypeTitle(i)=objRs(sFieldName)
		
		i=i+1
		objRs.MoveNext
	Loop
	objRs.Close

	Set objRs=Nothing
End Sub

Function EducationTypeTitleByID(iEduTypeID)
	Dim sResult, iEduTypeLoop
	sResult=""

	If IsArray(arrEduTypeID) And IsArray(arrEduTypeTitle) Then
		If LBound(arrEduTypeID)=LBound(arrEduTypeTitle) And UBound(arrEduTypeID)=UBound(arrEduTypeTitle) Then
			For iEduTypeLoop=LBound(arrEduTypeTitle) To UBound(arrEduTypeTitle)
				If arrEduTypeID(iEduTypeLoop)=iEduTypeID Then
					sResult=arrEduTypeTitle(iEduTypeLoop)
				End If
			Next
		End If
	End If

	EducationTypeTitleByID = sResult
End Function

GetEduTypeListData
%>
