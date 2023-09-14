<%
Class CExpertCvLanguage
	' Class Fields --------------------------------------
	Public Expert
	Public Language
	
	' Class Initialize and Terminate --------------------
	Private Sub Class_Initialize()
		Set Expert = New CBaseItem
		Set Language = New CBaseItem
	End Sub

	Private Sub Class_Terminate()
		If IsObject(Expert) Then
			Set Expert = Nothing
		End If
		If IsObject(Language) Then
			Set Language = Nothing
		End If
	End Sub
	
	' Class Methods -------------------------------------
	Public Function LoadDataFromRecordset(ARecordSet)
		If IsObject(ARecordSet) Then
			On Error Resume Next
			If Not ARecordSet.Eof Then
				Expert.ID = ARecordSet("id_Expert")
				Expert.UID = ARecordSet("uid_Expert")

				Language.ID = ARecordSet("id_Language")
				Language.Name = ARecordSet("lngNameEng")
				Language.Code = ARecordSet("lngAbbreviation")
			End If
			On Error GoTo 0
		End If
	End Function

	Public Function LoadData(AProcedure, AParams)
	Dim objTempRs
		Set objTempRs = GetDataRecordsetSPWithConn(objConnCustom, AProcedure, AParams)
		LoadDataFromRecordset objTempRs

		If objTempRs.State = adStateOpen Then objTempRs.Close
		Set objTempRs=Nothing
	End Function

End Class

Class CExpertCvLanguageList
	' Class Fields --------------------------------------
	Private FCount
	Public Item()

	' Class Initialize and Terminate --------------------
	Private Sub Class_Terminate()
		Call DestoyData()
		If IsObject(Item) Then
			Set Item = Nothing
		End If		
	End Sub	

	' Class Private Functions ---------------------------
	Private Function DestoyData()
		While FCount > 0
			Set Item(FCount - 1) = Nothing
			FCount = FCount - 1
		WEnd
	End Function

	' Class Properties ----------------------------------
	Public Property Get Count
		Count = FCount
    End Property

	' Class Methods -------------------------------------
	Public Function LoadData(AProcedure, AParams)
		Dim objTempRs, i
		FCount = 0
		Set objTempRs = GetDataRecordsetSPWithConn(objConn, AProcedure, AParams)
		ReDim Item(objTempRs.RecordCount - 1)
		If Not objTempRs.Eof Then
			objTempRs.MoveFirst
			While Not objTempRs.Eof
				Set Item(FCount) = New CExpertCvLanguage
				Item(FCount).LoadDataFromRecordset objTempRs
				
				FCount = FCount + 1
				objTempRs.MoveNext
			WEnd
		End If
		objTempRs.Close
		Set objTempRs=Nothing
	End Function
	
	Public Sub ShowSelectItems(ASelectedItem, ASelectionElement, AGroupItems)
		Dim i
		Dim sSelectionElement
		If FCount>0 Then
			For i=0 To FCount-1
				If ASelectionElement = "name" Then
					sSelectionElement = Item(i).Language.Name
				ElseIf ASelectionElement = "code" Then
					sSelectionElement = Item(i).Language.Code
				ElseIf ASelectionElement = "uid" Then
					sSelectionElement = Item(i).Expert.UID
				Else
					sSelectionElement = CStr(ReplaceIfEmpty(Item(i).Expert.ID, ""))
				End If
			%>
			<option value="<% =sSelectionElement %>"<% If CStr(ReplaceIfEmpty(ASelectedItem, "")) = sSelectionElement Then %> selected<% End If %>><% =Item(i).Language.Name %></option>
			<% 
			Next
		End If
	End Sub
End Class
%>
