<%
Dim sEmailSubject, sEmailBody
Dim iEmailHeaderSize, iEmailFooterSize

Function ReadEmailTemplate(sTemplateName)
Dim objFso, fileMailTemplate, sTempBody, sTempHeader, sTempFooter, iPos1, iPos2, iResultCode

	sEmailSubject=""
	sEmailBody=""

	Set objFso=Server.CreateObject("Scripting.FileSystemObject")

	' Mail subject & body
	Set fileMailTemplate=objFso.OpenTextFile(Server.MapPath(sHomePath & "_mails") & "\" & sTemplateName, 1)
	sTempBody=fileMailTemplate.ReadAll

	iPos1=InStr(sTempBody,"<subject>")+9
	iPos2=InStr(sTempBody,"</subject>")
	If iPos1>9 And iPos2>0 Then sEmailSubject=Mid(sTempBody, iPos1, iPos2-iPos1)

	iPos1=InStr(sTempBody,"<body>")+6
	iPos2=InStr(sTempBody,"</body>")
	If iPos1>6 And iPos2>0 Then sEmailBody=Mid(sTempBody, iPos1, iPos2-iPos1)
	Set fileMailTemplate=Nothing

	If sEmailBody>"" Then
		iResultCode=1	
			
		' Mail header with assortis logo
		Set fileMailTemplate=objFso.OpenTextFile(Server.MapPath(sHomePath & "_mails") & "\_SystemHeader.htm", 1)
		sTempHeader=fileMailTemplate.ReadAll
		sEmailBody= sTempHeader & sEmailBody
		iEmailHeaderSize=Len(sTempHeader)

		Set fileMailTemplate=Nothing

		' Mail footer 
		Set fileMailTemplate=objFso.OpenTextFile(Server.MapPath(sHomePath & "_mails") & "\_SystemFooter.htm", 1)
		sTempFooter=fileMailTemplate.ReadAll
		sEmailBody= sEmailBody & sTempFooter
		iEmailFooterSize=Len(sTempFooter)
		Set fileMailTemplate=Nothing
	Else
		iResultCode=0
	End If

Set objFso=Nothing
ReadEmailTemplate=iResultCode
End Function


Function PrepareEmailTemplate(sTemplateName, lstParams)
Dim arrParams, sParamName, sParamValue, i, iPos1, iPos2, iResultCode
	If ReadEmailTemplate(sTemplateName)=1 Then
		' Getting the list of params
		arrParams=Split(lstParams, ";;")
		For i=0 To UBound(arrParams)
			iPos1=InStr(arrParams(i), "=")
			If iPos1>0 Then
				' For every param in the list getting name and value
				sParamName=Trim(Left(arrParams(i), iPos1-1))
				sParamValue=Trim(Mid(arrParams(i), iPos1+1, Len(arrParams(i))-iPos1))

				' Replacing param in the template
				sEmailBody=Replace(sEmailBody,"<#" & sParamName & "#>", sParamValue)
			End If
		Next

		' Removing all not replaced params
		iPos1=InStr(sEmailBody, "<#")
		iPos2=InStr(sEmailBody, "#>")
		If iPos1>0 And iPos2>0 Then 
			iResultCode=0
		Else
			iResultCode=1
		End If		

		'While iPos1>0 And iPos2>0
		'	sEmailBody=Replace(sEmailBody, Mid(sEmailBody, iPos1, iPos2-iPos1+2), "")
		'	iPos1=InStr(sEmailBody, "<#")
		'	iPos2=InStr(sEmailBody, "#>")
		'WEnd
		
	End If
End Function

%>