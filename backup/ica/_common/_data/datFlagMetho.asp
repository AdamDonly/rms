<%
Dim arrFlagMethoID()
Dim arrFlagMethoTitle()

Sub GetFlagMethoListData()
	Dim objRs
	Dim sFieldName
	Dim iFlagMethoCount
	iFlagMethoCount = 0

	sFieldName = "flgTitle"

	Set objRs = GetDataRecordsetSP("usp_FlagMethoListSelect", Array( _
		Array(, adInteger, , Null), _
		Array(, adVarChar, 255, Null), _
		Array(, adVarChar, 80, Null)))
	i = 0
	iFlagMethoCount = objRs.RecordCount - 1
	ReDim arrFlagMethoID(iFlagMethoCount)
	ReDim arrFlagMethoTitle(iFlagMethoCount)

	' On Error Resume Next

	Do Until objRs.EOF 
		arrFlagMethoID(i) = objRs("id_FlagMetho")
		arrFlagMethoTitle(i) = objRs(sFieldName)

		i = i + 1
		objRs.MoveNext
	Loop
	objRs.Close

	Set objRs = Nothing
End Sub

Function FlagMethoTitleByID(iFlagMethoID)
	Dim sResult, iFlagMethoLoop
	sResult = ""

	If IsArray(arrFlagMethoID) And IsArray(arrFlagMethoTitle) Then
		If LBound(arrFlagMethoID) = LBound(arrFlagMethoTitle) And UBound(arrFlagMethoID) = UBound(arrFlagMethoTitle) Then
			For iFlagMethoLoop = LBound(arrFlagMethoTitle) To UBound(arrFlagMethoTitle)
				If arrFlagMethoID(iFlagMethoLoop) = iFlagMethoID Then
					sResult = arrFlagMethoTitle(iFlagMethoLoop)
				End If
			Next
		End If
	End If

	FlagMethoTitleByID = sResult
End Function

Sub ShowFlagMethoSelectItems(ASelectedItem, ASelectionElement, AGroupItems)
	Dim iLoop
	Dim sSelectionElement
	If IsArray(arrFlagMethoID) And IsArray(arrFlagMethoTitle) Then
		For iLoop = LBound(arrFlagMethoTitle) To UBound(arrFlagMethoTitle)
			If ASelectionElement = "id" Then
				sSelectionElement = arrFlagMethoID(iLoop)
			ElseIf ASelectionElement = "value" Then
				sSelectionElement = arrFlagMethoID(iLoop)
			Else
				sSelectionElement = arrFlagMethoTitle(iLoop)
			End If
		%>
		<option value="<% =sSelectionElement %>"<% If CStr(ReplaceIfEmpty(ASelectedItem, "")) = CStr(sSelectionElement) Then %> selected<% End If %>><% =arrFlagMethoTitle(iLoop) %></option>
		<% 
		Next
	End If
End Sub

GetFlagMethoListData
%>
