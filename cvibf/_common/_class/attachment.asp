<%
Class CAttachment
	' Class Fields --------------------------------------
	Public UID
	Public Title
	Public Path
	Public Length
	Public FileName
	Public FileExtension
	Public MIMEContentType
	Public TempPath
	
	Public Property Get Storage
		Dim objConnCustom
		Set objConnCustom = Server.CreateObject("ADODB.Connection")
		objConnCustom.Open "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Data Source=MAGNESA;Initial Catalog=" & objExpertDB.DatabasePath & ";"
		
		Storage=Null
		Dim objTempRsStorage
		If Len(UID)>16 Then
			Set objTempRsStorage=GetDataRecordsetSPWithConn(objConnCustom, "usp_DocumentBlobByUidSelect", Array( _
				Array(, adVarChar, 40, UID)))
			If Not objTempRsStorage.Eof Then
				Storage=objTempRsStorage("docImage")
			End If
			objTempRsStorage.Close
			Set objTempRsStorage=Nothing
		End If
		
		objConnCustom.Close
		Set objConnCustom = Nothing
    End Property

	Public Function Download()
		Dim objConnCustom
		Set objConnCustom = Server.CreateObject("ADODB.Connection")
		objConnCustom.Open "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Data Source=MAGNESA;Initial Catalog=" & objExpertDB.DatabasePath & ";"

		Response.Buffer = False 
		Server.ScriptTimeout = 30000 
	
		Const adTypeBinary = 1
		Const adModeRead = 1 
		Const adOpenStreamFromRecord = 4 
		Const iChuckSize = 4096 
		
		Dim objTempRsStorage
		
		If MIMEContentType>"" Then
			Response.ContentType = MIMEContentType
		Else
			Response.Write "Error. " & FileExtension & " file could not be opened."
			Response.End
		End If
		
		If Len(UID)>16 Then
			Set objTempRsStorage=GetDataRecordsetSPWithConn(objConnCustom, "usp_DocumentBlobByUidSelect", Array( _
				Array(, adVarChar, 40, UID)))
			If Not objTempRsStorage.Eof Then
				Dim objStream, iStreamSize, iFileSize
				Set objStream = CreateObject("ADODB.Stream") 
				objStream.Open() 
				objStream.Type = adTypeBinary
		
				objStream.Write objTempRsStorage("docImage")
				objStream.Position=0
	 
				iStreamSize = objStream.Size 
				Response.AddHeader "Content-Disposition", "attachment; filename=" & FileName
				Response.AddHeader "Content-Length", iStreamSize 
	 
				For i = 1 To iStreamSize \ iChuckSize 
					If Not Response.IsClientConnected Then Exit For 
					Response.BinaryWrite objStream.Read(iChuckSize) 
				Next 
	 
				If iStreamSize Mod iChuckSize > 0 Then 
					If Response.IsClientConnected Then 
						Response.BinaryWrite objStream.Read(iStreamSize Mod iChuckSize) 
					End If 
				End If 
	 
				objStream.Close 
				Set objStream = Nothing 
	 			End If
			objTempRsStorage.Close
			Set objTempRsStorage=Nothing
		End If

		objConnCustom.Close
		Set objConnCustom = Nothing
	End Function
End Class


Class CAttachmentList
	' Class Fields --------------------------------------
	Private FCount
	Public Item()

	' Class Initialize and Terminate --------------------
	Private Sub Class_Initialize()
	End Sub
	
	Private Sub Class_Terminate()
		Call DestoyData()
	End Sub	

	' Class Private Functions ---------------------------
	Private Function DestoyData()
		While FCount>0 
			If IsObject(Item(FCount-1)) Then
				Set Item(FCount-1)=Nothing
			End If
			FCount=FCount-1
		WEnd
		If IsArray(Item) Then
			ReDim Item(-1)
		End If
	End Function

	' Class Properties ----------------------------------
	Public Property Get Count
		Count=FCount
    End Property
End Class
%>
