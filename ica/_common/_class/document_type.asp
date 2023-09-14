<%
Class CDocumentTypeList
	' Class Fields --------------------------------------
	Private FCount
	Public Items()

	' Class Initialize and Terminate --------------------
	Private Sub Class_Initialize()
	End Sub
	
	Private Sub Class_Terminate()
		Call DestoyData()
		If IsObject(Items) Then
			Set Items=Nothing
		End If
	End Sub	

	' Class Private Functions ---------------------------
	Private Function DestoyData()
		While FCount>0 
			If IsObject(Items(FCount-1)) Then
				Set Items(FCount-1)=Nothing
			End If
			FCount=FCount-1
		WEnd
	End Function

	' Class Properties ----------------------------------
	Public Property Get Count
		Count=FCount
    End Property
	
	' Class Methods -------------------------------------
	Public Function LoadData
		LoadDataProc "usp_DocumentTypeSelect", Array()
	End Function
	
	Public Function LoadDataProc(AProcedure, AParams)
		Dim i
		Dim objTempRs
		FCount=0
		Set objTempRs=GetDataRecordsetSP(AProcedure, AParams)
		ReDim Items(objTempRs.RecordCount-1)
		If Not objTempRs.Eof Then
			objTempRs.MoveFirst
			While Not objTempRs.Eof
				Set Items(FCount) = New CStatus
			
				Items(FCount).ID = objTempRs("id_DocType")
				Items(FCount).Name = objTempRs("dtName")
				
				FCount=FCount+1
				objTempRs.MoveNext
			WEnd
		End If
		objTempRs.Close
		Set objTempRs=Nothing
	End Function


	Public Sub ShowSelectItems(ASelectedItem)
		Dim i
		If FCount>0 Then
			For i=0 To FCount-1
			%>
			<option value="<% =Items(i).ID %>"<% If CStr(ASelectedItem) = CStr(Items(i).ID) Then %> selected<% End If %>><% =Items(i).Name %></option>
			<% 
			Next
		End If
	End Sub	

End Class
%>
