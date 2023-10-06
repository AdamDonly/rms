<!--#include file="../_class/attachment.asp"-->
<%
Class CDocument
	' Class Fields --------------------------------------
	Public ID
	Public UID
	Public Title
	Public Type_
	Public Text
	Public Attachment
	Public DateCreated

	' Class Initialize and Terminate --------------------
	Private Sub Class_Initialize()
		ID = NULL
	End Sub
	
	Private Sub Class_Terminate()
		Call DestoyData()
	End Sub
	
	' Class Private Functions ---------------------------
	Private Function DestoyData()
		If IsObject(Attachment) Then
			Set Attachment=Nothing
		End If
	End Function

	Public Function LoadDataFromRecordset(ARecordSet)
		If IsObject(ARecordSet) Then
			On Error Resume Next
			If Not ARecordSet.Eof Then
				ID=ARecordSet("id_Document")
				UID=ARecordSet("uid_Document")
				Title=ARecordSet("docTitle")
				Type_=ARecordSet("docType")
				Text=ARecordSet("docText")
				DateCreated=ARecordSet("docCreated")
				If ARecordSet("docImageSize")>0 Then
					Set Attachment = New CAttachment
					Attachment.Path=ARecordSet("docPath")
					Attachment.FileName=ARecordSet("docPath")
					Attachment.FileExtension=GetFileExtension(Attachment.FileName)
					Attachment.MIMEContentType=GetFileMimeType(Attachment.FileName)
					Attachment.Length=ARecordSet("docImageSize")
					Attachment.UID=UID
				End If
			
			End If
			On Error GoTo 0
		End If
	End Function

	Public Function LoadData(AProcedure, AParams)
		Dim objConnCustom, objTempRs
		Set objConnCustom = Server.CreateObject("ADODB.Connection")
		objConnCustom.Open "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Data Source=MAGNESA;Initial Catalog=" & objExpertDB.DatabasePath & ";"
		
		Set objTempRs=GetDataRecordsetSPWithConn(objConnCustom, AProcedure, AParams)
		LoadDataFromRecordset objTempRs

		If objTempRs.State = adStateOpen Then objTempRs.Close
		Set objTempRs=Nothing
		objConnCustom.Close
		Set objConnCustom = Nothing
	End Function
	
	Public Function LoadDocumentDetailsByUid(ADocumentUID)
		LoadData "usp_DocumentByUidSelect", Array( _
			Array(, adVarChar, 40, ADocumentUID))
	End Function

	Public Function LoadDocumentDetailsById(ADocumentID)
		LoadData "usp_DocumentByIdSelect", Array( _
			Array(, adInteger, , ADocumentID))
	End Function

	Public Function DeleteByUid(ADocumentUID)
		SaveData "usp_DocumentByUidDelete", Array( _
			Array(, adVarChar, 40, ADocumentUID))
	End Function
	
	Public Function SaveData(AProcedure, AParams)
		Dim objConnCustom, Result
		Set objConnCustom = Server.CreateObject("ADODB.Connection")
		objConnCustom.Open "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Data Source=MAGNESA;Initial Catalog=" & objExpertDB.DatabasePath & ";"

		Result = UpdateRecordSPWithConn(objConnCustom, AProcedure, AParams)

		objConnCustom.Close
		Set objConnCustom = Nothing
		SaveData = Result
	End Function
	
	Public Function DownloadAttachment()
		If IsObject(Attachment) Then
			Attachment.Download
		End If
	End Function
	
	Public Function Download()
		DownloadAttachment
	End Function
	
	Public Function SaveForm(AExpertID, byref AFieldSet)
		If IsObject(AFieldSet) Then
			UID=Left(CheckString(AFieldSet("document")), 40)
			Title=Left(CheckString(AFieldSet("document_title")), 255)
			Type_=Left(CheckString(AFieldSet("document_type")), 255)
			Text=NULL
			If CreateFormAttachmentFile(AFieldSet("attachment"), "doc") = 1 Then
				ID=SaveData("usp_ExpertDocumentUpdate", Array( _
					Array(, adInteger, , iExpertID), _
					Array(, adVarChar, 40, UID), _
					Array(, adVarWChar, 255, Title), _
					Array(, adVarWChar, 255, Type_), _
					Array(, adVarWChar, 20000, Text), _
					Array(, adVarChar, 150, Attachment.Path), _
					Array(, adLongVarBinary, Attachment.Length, Attachment.TempPath)))
			Else
				ID=SaveData("usp_ExpertDocumentUpdate", Array( _
					Array(, adInteger, , iExpertID), _
					Array(, adVarChar, 40, UID), _
					Array(, adVarWChar, 255, Title), _
					Array(, adVarWChar, 255, Type_), _
					Array(, adVarWChar, 20000, Text), _
					Array(, adVarChar, 150, Null), _
					Array(, adLongVarBinary, Null, Null)))
			End If
		End If
	End Function
	
	Public Function CreateFormAttachmentFile(AFormField, AFileNameType)
		Dim iResult
		iResult = 0
		ID = 0
		DateCreated = Now()
	
		Dim iFileSize
		iFileSize=AFormField.TotalBytes
		If iFileSize>0 Then
			Set Attachment = New CAttachment
			Attachment.Length = iFileSize
			Attachment.Path = Trim(AFormField.UserFilename)
			Attachment.FileExtension = GetFileExtension(Attachment.Path)
			
			Attachment.Path = AFileNameType & "_" & ConvertDateTimeForFilename(Now) & "_" & Mid(sSessionID, 26, 9) & "." & Attachment.FileExtension
			Attachment.TempPath=Server.Mappath("../../../_upload") & sHomePath & "\" & Attachment.Path

			AFormField.SaveAs Attachment.TempPath
			iResult = 1
		End If
		CreateFormAttachmentFile = iResult
	End Function
End Class


Class CDocumentList
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
			Set Item(FCount-1)=Nothing
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

	' Class Methods -------------------------------------
	Public Function LoadData(AProcedure, AParams)
		Dim objConnCustom, objTempRs
		Set objConnCustom = Server.CreateObject("ADODB.Connection")
		objConnCustom.Open "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Data Source=MAGNESA;Initial Catalog=" & objExpertDB.DatabasePath & ";"

		FCount=0
		Set objTempRs=GetDataRecordsetSPWithConn(objConnCustom, AProcedure, AParams)
		ReDim Item(objTempRs.RecordCount-1)
		If Not objTempRs.Eof Then
			objTempRs.MoveFirst
			While Not objTempRs.Eof
				Set Item(FCount) = New CDocument
				Item(FCount).LoadDataFromRecordset objTempRs
				
				FCount=FCount+1
				objTempRs.MoveNext
			WEnd
		End If

		If objTempRs.State = adStateOpen Then objTempRs.Close
		Set objTempRs=Nothing
		objConnCustom.Close
		Set objConnCustom = Nothing
	End Function
	
	Public Function LoadDocumentListByExpertID(AExpertID, ADocumentType)
		LoadData "usp_ExpertDocumentListSelect", Array( _
			Array(, adInteger, , AExpertID), _
			Array(, adVarchar, 255, ADocumentType))
	End Function
End Class

%>