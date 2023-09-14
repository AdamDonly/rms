<%
Function ReadFile(sFileName)
Dim sContent
Dim objFso, objFile

	sContent=""
	Set objFso=Server.CreateObject("Scripting.FileSystemObject")
	Set objFile=objFso.OpenTextFile(sFileName, 1)
	sContent=objFile.ReadAll

	objFile.Close
	Set objFile=Nothing
	Set objFso=Nothing
ReadFile=sContent
End Function


Function WriteFile(sFileName, sContent)
Const ForWriting = 2
Const Create = True 
Dim objFso, objFile

	Set objFso = Server.CreateObject("Scripting.FileSystemObject")
	Set objFile = objFso.OpenTextFile(sFileName, ForWriting, Create)

	objFile.Write sContent

	objFile.Close
	Set objFile = Nothing
	Set objFso = Nothing
End Function


Function GetFileExtension(AFileName)
	Dim sResult, iPos
	
	If IsNull(AFileName) Or IsEmpty(AFileName) Then AFileName=""
	iPos=InStr(StrReverse(AFileName), ".")-1
	If iPos>0 Then
		sResult=LCase(Right(AFileName, iPos))
	Else
		sResult=""
	End If		

GetFileExtension=sResult
End Function


Function GetFileMimeType(AFileName)
	Dim sResult
	sResult=""
	
	Dim sFileExtension
	sFileExtension=GetFileExtension(AFileName)
	
	If sFileExtension="pdf" Then
		sResult = "application/pdf"
	ElseIf sFileExtension="doc" Or sFileExtension="dot" Or sFileExtension="rtf" Or sFileExtension="docx" Then
		sResult = "application/msword"
	ElseIf sFileExtension="xls" Or sFileExtension="xlt" Or sFileExtension="xlw" Or sFileExtension="csv" Or sFileExtension="xlsx" Or sFileExtension="xlsb" Or sFileExtension="xltx" Then
		sResult = "application/vnd.ms-excel"
	ElseIf sFileExtension="ppt" Or sFileExtension="pps" Or sFileExtension="pot" Or sFileExtension="pptx" Or sFileExtension="ppsx" Or sFileExtension="pptm" Or sFileExtension="potm" Then
		sResult = "application/vnd.ms-powerpoint"
	ElseIf sFileExtension="txt" Or sFileExtension="htm" Or sFileExtension="html" Or sFileExtension="xml" Then
		sResult = "application/vnd.text"
	ElseIf sFileExtension="txt" Or sFileExtension="htm" Or sFileExtension="html" Or sFileExtension="xml" Then
		sResult = "application/vnd.text"
	ElseIf sFileExtension="tif" Or sFileExtension="tiff" Then
		sResult = "image/tiff"
	ElseIf sFileExtension="jpg" Or sFileExtension="jpeg"  Or sFileExtension="jpe" Then
		sResult = "image/jpeg"
	ElseIf sFileExtension="gif" Then
		sResult = "image/gif"
	ElseIf sFileExtension="bmp" Then
		sResult = "image/bmp"
	ElseIf sFileExtension="zip" Or sFileExtension="rar" Then
		sResult = "application/zip"
	End If
	
GetFileMimeType=sResult
End Function
%>
