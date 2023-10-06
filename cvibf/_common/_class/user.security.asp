<%
Const aAccessNone = 0
Const aAccessAdminAll = 1
Const aAccessAdminFinance = 8
Const aAccessAdminIT = 9
Const aAccessAdminAssortis = 10
Const aAccessAdminIbf = 20
Const aAccessAdminIca = 30

Const aAccessAdminCountryHubAssortis = 15
Const aAccessRestrictedCountryHubCountry = 16
Const aAccessReadonlyAssortis = 17
Const aAccessTrialsAssortis = 18
Const aAccessRestrictedViewAddNewAssortis = 19

Function GetAccessSecurity(sUserType)
	Dim iResult
	iResult = aAccessNone

	If sUserType = "Admin" Then
		iResult = aAccessAdminAll
	ElseIf sUserType = "Admin Finance" Then
		iResult = aAccessAdminFinance
	ElseIf sUserType = "Admin IT" Then
		iResult = aAccessAdminIT
	ElseIf sUserType = "Admin Assortis" Then
		iResult = aAccessAdminAssortis
	ElseIf sUserType = "Admin IBF" Then
		iResult = aAccessAdminIbf
	ElseIf sUserType = "Admin ICA" Then
		iResult = aAccessAdminIca
	ElseIf sUserType = "Admin CountryHub Assortis" Then
		iResult = aAccessAdminCountryHubAssortis
	ElseIf sUserType = "Restricted CountryHub Country" Then
		iResult = aAccessRestrictedCountryHubCountry
	ElseIf sUserType = "Readonly Assortis" Then
		iResult = aAccessReadonlyAssortis
	ElseIf sUserType = "Trials Assortis" Then
		iResult = aAccessTrialsAssortis
	ElseIf sUserType = "Restricted View Add New Assortis" Then
		iResult = aAccessRestrictedViewAddNewAssortis
	End If

	GetAccessSecurity = iResult
End Function

Class CIp
	Public Ip4
	Public Ip6
	
	Public Function LoadDataFromRecordset(ARecordSet)
		If IsObject(ARecordSet) Then
			On Error Resume Next
			If Not ARecordSet.Eof Then
				Ip4=ARecordSet("ussIp4Security")
				Ip6=ARecordSet("ussIp6Security")
			End If
			On Error GoTo 0
		End If
	End Function
End Class

Class CIbfSecurity
	Public IbfDB
	Public BrowseDB
	Public MpisUserCode
	
	Public Function LoadDataFromRecordset(ARecordSet)
		If IsObject(ARecordSet) Then
			On Error Resume Next
			If Not ARecordSet.Eof Then
				IbfDB=ARecordSet("usrIbf")
				BrowseDB=ARecordSet("usrBrowseMode")
				MpisUserCode=ARecordSet("usrElinkUserCode")
			End If
			On Error GoTo 0
		End If
	End Function
	
	Public Function LoadData(AProcedure, AParams)
	Dim objTempRs
		Set objTempRs=GetDataRecordsetSPWithConn(objConn, AProcedure, AParams)
		LoadDataFromRecordset objTempRs

		If objTempRs.State = adStateOpen Then objTempRs.Close
		Set objTempRs=Nothing
	End Function
	
	Public Function SaveData(AProcedure, AParams)
	Dim objTempRs
		objTempRs=UpdateRecordSPWithConn(objConn, AProcedure, AParams)
		Set objTempRs=Nothing
	End Function
	
	Public Function Save(AUserID)
		SaveData "usp_UserSecurityIbfUpdate", Array( _
			Array(, adInteger, , AUserID), _
			Array(, adInteger, , IbfDB), _
			Array(, adInteger, , BrowseDB), _
			Array(, adVarChar, 32, MpisUserCode))
	End Function
End Class

Class CIcaSecurity
	Public IcaID
	
	Public Function LoadDataFromRecordset(ARecordSet)
		If IsObject(ARecordSet) Then
			On Error Resume Next
			If Not ARecordSet.Eof Then
				IcaID=ARecordSet("id_IcaUser")
			End If
			On Error GoTo 0
		End If
	End Function
	
	Public Function LoadData(AProcedure, AParams)
	Dim objTempRs
		Set objTempRs=GetDataRecordsetSPWithConn(objConn, AProcedure, AParams)
		LoadDataFromRecordset objTempRs

		If objTempRs.State = adStateOpen Then objTempRs.Close
		Set objTempRs=Nothing
	End Function
	
	Public Function SaveData(AProcedure, AParams)
	Dim objTempRs
		objTempRs=UpdateRecordSPWithConn(objConn, AProcedure, AParams)
		Set objTempRs=Nothing
	End Function
	
	Public Function Save(AUserID)
		SaveData "usp_UserSecurityIcaUpdate", Array( _
			Array(, adInteger, , AUserID), _
			Array(, adInteger, , IcaID))
	End Function
End Class
	
Class CIpList
	' Class Fields --------------------------------------
	Private FCount
	Public Items()
	Public List
	Public ListDelimiter
	Public Ip4Ip6Delimiter
	
	' Class Initialize and Terminate --------------------
	Private Sub Class_Initialize()
		List = ""
		ListDelimiter = ","
		Ip4Ip6Delimiter = "/"
	End Sub
	
	Private Sub Class_Terminate()
		Call DestoyData()
	End Sub	

	' Class Private Functions ---------------------------
	Private Function DestoyData()
		While FCount>0 
			If IsObject(Items(FCount-1)) Then
				Set Items(FCount-1)=Nothing
			End If
			FCount=FCount-1
		WEnd
		If IsArray(Items) Then
			ReDim Items(-1)
		End If
	End Function

	' Class Properties ----------------------------------
	Public Property Get Count
		Count=FCount
    End Property

	' Class Methods -------------------------------------
	Public Function LoadDataFromRecordset(ARecordSet)
		FCount=0
	
		If IsObject(ARecordSet) Then
			On Error Resume Next

			ReDim Items(ARecordSet.RecordCount-1)
			If Not ARecordSet.Eof Then
				ARecordSet.MoveFirst
				While Not ARecordSet.Eof
					Set Items(FCount) = New CIp
					Items(FCount).LoadDataFromRecordset(ARecordSet)
			
					FCount=FCount+1
					ARecordSet.MoveNext
				WEnd
			End If
			
			On Error GoTo 0
		End If
	End Function
	
	Public Function LoadData(AProcedure, AParams)
	Dim objTempRs
	
		Set objTempRs=GetDataRecordsetSPWithConn(objConn, AProcedure, AParams)
		LoadDataFromRecordset objTempRs

		If objTempRs.State = adStateOpen Then objTempRs.Close
		Set objTempRs=Nothing
	End Function
	
	Public Function SaveData(AProcedure, AParams)
	Dim objTempRs
		objTempRs=UpdateRecordSPWithConn(objConn, AProcedure, AParams)
		Set objTempRs=Nothing
	End Function

	Public Function Save(AUserID)
		SaveData "usp_UserSecurityIpUpdate", Array( _
			Array(, adInteger, , AUserID), _
			Array(, adVarChar, 4000, List), _
			Array(, adVarChar, 5, ListDelimiter), _
			Array(, adVarChar, 5, Ip4Ip6Delimiter))
	End Function
End Class


Class CUserSecurity
	Public UserID
	Public IpSecurity
	
	Public AssortisSecurity
	Public IbfSecurity
	Public IcaSecurity
	
	Private Sub Class_Initialize()
		Set IpSecurity = New CIpList
		Set IbfSecurity = New CIbfSecurity
		Set IcaSecurity = New CIcaSecurity
	End Sub
	
	Private Sub Class_Terminate()
		If IsObject(IpSecurity) Then
			Set IpSecurity=Nothing
		End If
		If IsObject(IbfSecurity) Then
			Set IbfSecurity=Nothing
		End If
		If IsObject(IcaSecurity) Then
			Set IcaSecurity=Nothing
		End If
	End Sub	
	
	Public Function LoadUserSecurity()
		If IsNull(UserID) Or IsEmpty(UserID) Then Exit Function

		IpSecurity.LoadData "usp_UserSecurityIpSelect", Array( _
			Array(, adInteger, , UserID))

		IbfSecurity.LoadData "usp_UserSecurityIbfSelect", Array( _
			Array(, adInteger, , UserID))

		IcaSecurity.LoadData "usp_UserSecurityIcaSelect", Array( _
			Array(, adInteger, , UserID))
	End Function
	
	
	Public Function SaveUserSecurity()
		If iUserTypeAccessSecurity > aAccessNone Then
			If iUserTypeAccessSecurity = aAccessAdminAll Then
				objUserSecurity.IpSecurity.Save(UserID)
			End If
		
			If iUserTypeAccessSecurity = aAccessAdminAll Or _
				iUserTypeAccessSecurity = aAccessAdminIbf Then
				objUserSecurity.IbfSecurity.Save(UserID)
			End If

			If iUserTypeAccessSecurity = aAccessAdminAll Or _
				iUserTypeAccessSecurity = aAccessAdminIca Then
				objUserSecurity.IcaSecurity.Save(UserID)
			End If
		End If
	End Function
End Class

%>