<!--#include virtual="/_fnc_file.asp"-->
<% 
Sub ShowDocumentListEditTable(AObjDocumentList)
	Dim iDocument
	Dim sFileName, iFileSize, sFileSize, sFileExtension
	If AObjDocumentList.Count>0 Then
		Response.Write "<table width=100% celpadding=0 cellspacing=0 border=0>"
		For iDocument=0 To AObjDocumentList.Count-1
			If IsObject(AObjDocumentList.Item(iDocument).Attachment) Then
				sFileName=AObjDocumentList.Item(iDocument).Attachment.Path
				sFileExtension=GetFileExtension(sFileName)

				iFileSize=AObjDocumentList.Item(iDocument).Attachment.Length
				If iFileSize>1000000 Then
					sFileSize=Round(iFileSize/1048576) & "&nbsp;Mb"
				ElseIf iFileSize>1024 Then
					sFileSize=Round(iFileSize/1024) & "&nbsp;Kb"
				Else
					sFileSize=iFileSize & "&nbsp;b"
				End If

				Dim sDocumentScriptFileName
				sDocumentScriptFileName = "document.asp"
				If sScriptFileName = "exp_methodology.asp" Then
					sDocumentScriptFileName = "doc_methodology.asp"
				End If

				Dim sTempParams
				sTempParams = sParams
				sTempParams = ReplaceUrlParams(sTempParams, "document=" & AObjDocumentList.Item(iDocument).Attachment.UID)
				sTempParams = ReplaceUrlParams(sTempParams, "uid=" & sExpertUID)
				sTempParams = ReplaceUrlParams(sTempParams, "idproject")
				sTempParams = ReplaceUrlParams(sTempParams, "qid")

				If sFileName > "" Then
					Response.Write "<tr><td width=""20"" valign=""top""><a href=""" & sDocumentScriptFileName & sTempParams & """><img src=""" & sHomePath & "image/vn_updt.gif"" hspace=""1"" vspace=""3"" border=""0""></a></td>"
					Response.Write "<td><p class=""sml""><a href=""" & sHomePath & "download/" & sTempParams & """>" & AObjDocumentList.Item(iDocument).Title & "</a> (" & sFileSize & ")</p></td></tr>"
				End If
			End If
		Next
		Response.Write "</table>"
	End If
End Sub

Sub ShowDocumentListViewTable(AObjDocumentList)
	Dim iDocument
	Dim sFileName, iFileSize, sFileSize, sFileExtension
	If AObjDocumentList.Count>0 Then
		Response.Write "<table width=100% celpadding=0 cellspacing=0 border=0>"
		For iDocument=0 To AObjDocumentList.Count-1
			If IsObject(AObjDocumentList.Item(iDocument).Attachment) Then
				sFileName=AObjDocumentList.Item(iDocument).Attachment.Path
				sFileExtension=GetFileExtension(sFileName)

				iFileSize=AObjDocumentList.Item(iDocument).Attachment.Length
				If iFileSize>1000000 Then
					sFileSize=Round(iFileSize/1048576) & "&nbsp;Mb"
				ElseIf iFileSize>1024 Then
					sFileSize=Round(iFileSize/1024) & "&nbsp;Kb"
				Else
					sFileSize=iFileSize & "&nbsp;b"
				End If

				Dim sTempParams
				sTempParams = sParams
				sTempParams = ReplaceUrlParams(sTempParams, "document=" & AObjDocumentList.Item(iDocument).Attachment.UID)
				sTempParams = ReplaceUrlParams(sTempParams, "uid=" & sExpertUID)
				sTempParams = ReplaceUrlParams(sTempParams, "idproject")
				sTempParams = ReplaceUrlParams(sTempParams, "qid")

				If sFileName > "" Then
					Response.Write "<tr><td width=""20"" valign=""top""><a href=""" & sHomePath & "download/" & sTempParams & """><img src=""" & sHomePath & "image/file" & sFileExtension & ".gif"" hspace=""1"" vspace=""3"" border=""0""></a></td>"
					Response.Write "<td><p class=""sml""><a href=""" & sHomePath & "download/" & sTempParams & """>" & AObjDocumentList.Item(iDocument).Title & "</a> (" & sFileSize & ")</p></td></tr>"
				End If
			End If
		Next
		Response.Write "</table>"
	End If
End Sub

Sub ShowDocument(AObjDocument)
	Dim sFileName, sFileExtension
	If IsObject(AObjDocument.Attachment) Then
		sFileName=AObjDocument.Attachment.Path
		sFileExtension=GetFileExtension(sFileName)

		If sFileName>"" Then 
		%>
			<a href="<% =sHomePath %>download/?document=<% =AObjDocument.Attachment.UID %>"><img src="<% =sHomePath %>image/file<% =sFileExtension %>.gif" hspace="0" vspace="0" border="0" align="left"></a>
		<%
		End If
	End If
End Sub
%>