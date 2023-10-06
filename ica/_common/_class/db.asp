<!--#include file="../_class/company.asp"-->
<%
Const aCompanyExpertDBAssortisID = 101

Class CCompanyExpertDB
	' Class Fields --------------------------------------
	Public ID
	Public Database
	Public DatabaseCode
	Public DatabaseCodePrimary
	Public DatabaseTitle
	Public DatabasePath
	Public Company
	Public IsVisible

	Public ContactName
	Public ContactEmail
	
	' Class Initialize and Terminate --------------------
	Private Sub Class_Initialize()
		ID = 0
		IsVisible = 0
		Set Company = New CCompany
	End Sub
	
	Private Sub Class_Terminate()
		If IsObject(Company) Then
			Set Company=Nothing
		End If
	End Sub
	
	' Class Methods -------------------------------------
	Public Function LoadDataFromRecordset(ARecordSet)
		If IsObject(ARecordSet) Then
			On Error Resume Next
			If Not ARecordSet.Eof Then
				ID = ARecordSet("id_Database")
				Database = ARecordSet("edbName")
				DatabaseCode = ARecordSet("edbCode")
				DatabaseCodePrimary = DatabaseCode
				DatabaseTitle = ARecordSet("edbTitle")
				DatabasePath = ARecordSet("edbPath")
				IsVisible = CheckInteger(ARecordSet("edbVisible"))

				Company.ID = ARecordSet("id_Company")
				Company.Name = ARecordSet("icaCompanyShortName")
				If Company.Name Is Nothing Or Company.Name = "" Then 
					Company.Name = DatabaseTitle
				End If
				If Company.Name Is Nothing Or Company.Name = "" Then 
					Company.Name = Database
				End If
				
				ContactName = ARecordSet("edbContactName")
				ContactEmail = ARecordSet("edbContactEmail")
			End If
			On Error GoTo 0
		End If
	End Function

	Public Function LoadData(AProcedure, AParams)
	Dim objTempRs
		Set objTempRs=GetDataRecordsetSP(AProcedure, AParams)
		LoadDataFromRecordset objTempRs

		If objTempRs.State = adStateOpen Then objTempRs.Close
		Set objTempRs=Nothing
	End Function

	Public Function LoadCompanyDatabase(ACompanyID, ADatabaseID, ADatabase)
		LoadData "usp_CompanyDatabaseSelect", Array( _
			Array(, adInteger, , ACompanyID), _
			Array(, adInteger, , ADatabaseID), _
			Array(, adVarChar, 50, ADatabase),_
			Array(, adInteger, , 0))
	End Function
	
End Class

Class CCompanyExpertDBList
	' Class Fields --------------------------------------
	Private FCount
	Public Item()

	Public DefaultDatabase
	Private FDefaultDB
	
	' Class Initialize and Terminate --------------------
	Private Sub Class_Terminate()
		If IsObject(FDefaultDB) Then
			Set FDefaultDB = Nothing
		End If

		Call DestoyData()
	End Sub

	' Class Private Functions ---------------------------
	Private Function DestoyData()
		While FCount>0 
			If IsObject(Item(FCount-1).Company) Then
				Set Item(FCount-1).Company=Nothing
			End If
			If IsObject(Item(FCount-1)) Then
				Set Item(FCount-1)=Nothing
			End If
			FCount=FCount-1
		WEnd
		If IsArray(Item) Then
			ReDim Item(-1)
		End If
	End Function

	Public Property Get Count
		Count=FCount
    End Property

	' Class Methods -------------------------------------
	Public Function LoadData(AProcedure, AParams)
		Dim objTempRs
		Dim i
		FCount=0
		Set objTempRs=GetDataRecordsetSPWithConn(objConn, AProcedure, AParams)
		ReDim Item(objTempRs.RecordCount-1)
		If Not objTempRs.Eof Then
			objTempRs.MoveFirst
			While Not objTempRs.Eof
				Set Item(FCount) = New CCompanyExpertDB
				Item(FCount).LoadDataFromRecordset objTempRs
				If Item(FCount).Database = DefaultDatabase Then
					Set FDefaultDB = Item(FCount)
				End If
				
				FCount = FCount + 1
				objTempRs.MoveNext
			WEnd
		End If

		If objTempRs.State = adStateOpen Then objTempRs.Close
		Set objTempRs=Nothing
	End Function
	
	Public Function LoadCompanyDatabaseList(ACompanyID, ADatabaseID, ADatabase, AIncludeAllDB)
		If iMemberAccessExperts = cMemberAccessExpertsOwnOnly Then
			LoadData "usp_CompanyDatabaseSelect", Array( _
				Array(, adInteger, , ACompanyID), _
				Array(, adInteger, , ADatabaseID), _
				Array(, adVarChar, 50, ADatabase),_
				Array(, adInteger, , 0))
		Else
			LoadData "usp_CompanyDatabaseSelect", Array( _
				Array(, adInteger, , ACompanyID), _
				Array(, adInteger, , ADatabaseID), _
				Array(, adVarChar, 50, ADatabase),_
				Array(, adInteger, , AIncludeAllDB))
		End If
	End Function
	
	Public Function Find(ASelectedItem, ASelectionElement)
		Dim Result
		Dim i
		Dim sSelectionElement
		If FCount > 0 Then
			For i = 0 To FCount - 1
				If ASelectionElement="Database" Then
					sSelectionElement=Item(i).Database
				ElseIf ASelectionElement="DatabaseCode" Then
					sSelectionElement=Item(i).DatabaseCode
				ElseIf ASelectionElement="ID" Then
					sSelectionElement=CStr(Item(i).ID)
				ElseIf ASelectionElement="CompanyName" Then
					sSelectionElement=Item(i).Company.Name
				Else
					sSelectionElement=CStr(ReplaceIfEmpty(Item(i).Company.ID, ""))
				End If
				
				If CStr(ReplaceIfEmpty(ASelectedItem, ""))=sSelectionElement Then
					Set Result = Item(i)
				End If
			Next
		End If
		
		If Not IsObject(Result) Then
			Set Result = objNullDB
		End If
		
	Set Find = Result		
	End Function
	
	Public Sub ShowSelectItems(ASelectedItem, ASelectionElement, AGroupItems)
		Dim i
		Dim sSelectionElement
		If FCount>0 Then
			For i=0 To FCount-1
				If ASelectionElement="Database" Then
					sSelectionElement=Item(i).Database
				ElseIf ASelectionElement="DatabaseCode" Then
					sSelectionElement=Item(i).DatabaseCode
				ElseIf ASelectionElement="ID" Then
					sSelectionElement=CStr(Item(i).ID)
				ElseIf ASelectionElement="CompanyName" Then
					sSelectionElement=Item(i).Company.Name
				Else
					sSelectionElement=CStr(ReplaceIfEmpty(Item(i).Company.ID, ""))
				End If
				
				If Item(i).IsVisible = 1 Then %>
					<option value="<% =sSelectionElement %>"<% If CStr(ReplaceIfEmpty(ASelectedItem, ""))=sSelectionElement Then %> selected<% End If %>><% =Item(i).DatabaseTitle %></option>
				<% End If
			Next
		End If
	End Sub
	
	Public Function List(ASelectionElement, ASeparator)
		Dim Result
		Dim i
		Dim sSelectionElement
		If FCount > 0 Then
			For i = 0 To FCount - 1
				On Error Resume Next
				If ASelectionElement="Database" Then
					sSelectionElement=Item(i).Database
				ElseIf ASelectionElement="DatabaseCode" Then
					sSelectionElement=Item(i).DatabaseCode
				ElseIf ASelectionElement="ID" Then
					sSelectionElement=CStr(Item(i).ID)
				ElseIf ASelectionElement="CompanyName" Then
					sSelectionElement=Item(i).Company.Name
				Else
					sSelectionElement=CStr(ReplaceIfEmpty(Item(i).Company.ID, ""))
				End If
				On Error GoTo 0
				
				Result = Result & sSelectionElement
				If i<FCount-1 Then
					Result = Result & ASeparator
				End If
			Next
		End If
	
	List=Result
	End Function

	Public Sub AddDB(DB)
		If Not IsObject(DB) Then Exit Sub

		Dim AlreadyExist
		AlreadyExist = False
		If FCount > 0 Then
			For i = 0 To FCount - 1
				If CStr(DB.ID) = CStr(Item(i).ID) Then 
					AlreadyExist = True
				End If
			Next
		End If
		If AlreadyExist Then Exit Sub

		ReDim Preserve Item(FCount)
		Set Item(FCount) = DB
		FCount = FCount + 1
	End Sub

	Public Sub ReplaceDB(DB, NewDB)
		If Not IsObject(DB) Then Exit Sub
		If Not IsObject(NewDB) Then Exit Sub

		If FCount > 0 Then
			For i = 0 To FCount - 1
				If CStr(DB.ID) = CStr(Item(i).ID) Then 
					Set Item(i) = NewDB
				End If
			Next
		End If
	End Sub

	Public Sub ClearAll
		FCount = 0
		ReDim Item(-1)		
	End Sub

End Class
%>