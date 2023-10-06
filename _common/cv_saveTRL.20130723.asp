<%
'--------------------------------------------------------------------
'
' Expert's CV. Save in assortis.com format
'
'--------------------------------------------------------------------
Response.Buffer = True
%>
<!--#include file="cv_data.asp"-->
<%
sFileType=LCase(Request.QueryString("ftype"))
If sFileType="doc" Then
	Response.ContentType = "application/vnd.ms-word"
	Response.AddHeader "Content-Disposition", "attachment; filename=" & sFileName & ".rtf"
End If
If sFileType="prn" Then
	Response.ContentType = "application/vnd.ms-word"
	Response.AddHeader "Content-Disposition", "inline; filename=" & sFileName & ".rtf"
End If
sLastName=Replace(ConvertText2RTF(sLastName), "      ", "")
sFirstName=Replace(ConvertText2RTF(sFirstName), "      ", "")
sFullNameWithSpaces=sFirstName & " " & sLastName
sFullName=Replace(ConvertText2RTF(sFullName), "      ", "")

Dim sHeader
Set objFso=Server.CreateObject("Scripting.FileSystemObject")
Set fInTemplate=objFso.OpenTextFile(Server.MapPath("\_common") & "\cv_trl.vrf", 1)
sHeader = fInTemplate.ReadAll & vbCrLf
sHeader = Replace(sHeader, "<#Name#>", sFullNameWithSpaces)
Response.Write sHeader
Set fInTemplate=Nothing
Set objFSO=Nothing

	WriteTableHeader
	WriteDataRow "Position", "\b1\up1\cf2  \b0\up0\cf1"
	WriteDataRow "Profile", sKeyQualification
	WriteDataRow "Key Skills", ""
	
' Nationality & languages
	Dim sLanguagesList
	sLanguagesList = ""
	If (Not objRsExpLngOther.Eof) Or (Not objRsExpLngNative.Eof) Then
		sLanguagesList = sLanguagesList & GetListStart()
		While Not objRsExpLngNative.Eof
			sLanguagesList = sLanguagesList & GetListItemStart() & objRsExpLngNative("lngNameEng") & " (native)" & GetListItemEnd()
			objRsExpLngNative.MoveNext
		WEnd

		While Not objRsExpLngOther.Eof
			sLanguagesList = sLanguagesList & GetListItemStart() & objRsExpLngOther("lngNameEng") & " (" & LCase(arrLanguageLevelTitle(objRsExpLngOther("exlAverage"))) & ")" & GetListItemEnd()
			objRsExpLngOther.MoveNext
		WEnd
		sLanguagesList = sLanguagesList & GetListEnd()
	End If
	objRsExpLngNative.Close
	Set objRsExpLngNative=Nothing	
	objRsExpLngOther.Close
	Set objRsExpLngOther=Nothing	

	WriteDataRow2 "Nationality", sNationality, "Languages", sLanguagesList
	
' Education	
	WriteDataRow "Qualifications", sKeyQualification
	
' Country Experience
	WriteDataRow "Country Experience", sKeyQualification
	
' Clients
	WriteDataRow "Clients", sKeyQualification
	
	WriteTableFooter

Response.Write("\par }}" & vbCrLf)
%>


