<%
' Script name & params used for maintaining sessions states and tracking users activities in cookieless environment
Dim sScriptServerProtocol, sScriptServerName, sScriptFullName, sScriptFileName, sScriptBaseName
Dim sParams, sScriptFullNameAsParams, sTempParams
sScriptServerProtocol=LCase(Request.ServerVariables("SERVER_PROTOCOL"))
If InStr(sScriptServerProtocol, "http")>0 Then
	If LCase(Request.ServerVariables("HTTPS"))="on" Then
		sScriptServerProtocol="https://"
	Else
		sScriptServerProtocol="http://"
	End If
Else
	sScriptServerProtocol=""
End If

sScriptServerName=LCase(Request.ServerVariables("SERVER_NAME"))

sScriptFullName=LCase(Request.ServerVariables("SCRIPT_NAME"))
sScriptFileName=Right(sScriptFullName, Len(sScriptFullName) - InStrRev(sScriptFullName, "/"))
sScriptBaseName=Replace(sScriptFullName, sScriptFileName, "", 1, -1, 1)
sParams=Request.QueryString

Dim iError, sError, sUrl, sShortUrl, sAction
iError=0
sError=Request.QueryString("err")
sAction=Request.QueryString("act")

If sParams>"" Then 
	sParams=ReplaceUrlParams(sParams, "err")
	sParams=ReplaceUrlParams(sParams, "act")

	sParams=Replace(sParams,"?&","?")
	sParams=Replace(sParams,"&&","&")
	If Right(sParams,1)="&" Then sParams=Left(sParams,Len(sParams)-1)
End If
If sParams>"" Then 
	sParams="?" & sParams
End If

sScriptFullName=sScriptFullName & sParams
sScriptFullNameAsParams="url=" & EncodeUrlParams(sScriptFullName)

sUrl=LCase(Request.QueryString("url"))
If Len(sUrl)<3 Then
	sUrl=sHomePath
Else
	sUrl=sUrl & Replace(sParams, "url=" & sUrl, "")

	If Right(sUrl,1)="?" Then sUrl=Left(sUrl, Len(sUrl)-1)
	If Right(sUrl,1)="&" Then sUrl=Left(sUrl, Len(sUrl)-1)
	sUrl=Replace(sUrl,"?&","?")
	sUrl=Replace(sUrl,"&&","&")
End If

'--------------------------------------------------------------------
' Function for replacing params to the url string
'--------------------------------------------------------------------
Function ReplaceUrlParams(sUrlIn, sParamIn)
Dim sUrlTemp, sParamNameTemp, iPos, iPosUrl1, iPosUrl2
Dim sActualParamValue
	
	iPos=InStr(sParamIn,"=")
	' If selected Params without value then remove it from the sParams list
	If (iPos<1) Or IsNull(iPos) Or iPos=Len(Trim(sParamIn)) Then
		sUrlTemp=sUrlIn
		sActualParamValue=Request.QueryString(sParamIn)
		If sParamIn="sid" And Request.QueryString(sParamIn)="" Then sActualParamValue=sSessionID
		
		' For params with multiple values
		If InStr(sActualParamValue, ",")>0 Then
			Dim arrActualParamValue, loopActualParamValue
			arrActualParamValue=Split(sActualParamValue)
			If IsArray(arrActualParamValue) Then
				For Each loopActualParamValue In arrActualParamValue
					sUrlTemp=RemoveUrlParamValue(sUrlTemp, sParamIn, loopActualParamValue)
				Next
			End If			
		Else
			sUrlTemp=RemoveUrlParamValue(sUrlTemp, sParamIn, sActualParamValue)
		End If
	
		'	sUrlTemp=Replace(sUrlTemp, "&" & sParamIn & "=" & sActualParamValue, "")
		'	sUrlTemp=Replace(sUrlTemp, sParamIn & "=" & sActualParamValue & "&", "")
		'	sUrlTemp=Replace(sUrlTemp, "?" & sParamIn & "=" & sActualParamValue, "?")
		If sUrlTemp=sParamIn & "=" & sActualParamValue Then
			sUrlTemp=""
		End If
	Else
		sParamNameTemp=Left(sParamIn, iPos-1)
		If sUrlIn>"" Then
			sUrlTemp=sUrlIn
			sUrlTemp=Replace(sUrlTemp, "&" & sParamNameTemp & "=" & Request.QueryString(sParamNameTemp), "")
			sUrlTemp=Replace(sUrlTemp, "?" & sParamNameTemp & "=" & Request.QueryString(sParamNameTemp), "?")
			sUrlTemp=sUrlTemp & "&" & sParamIn
		Else
			sUrlTemp=sUrlIn & "?" & sParamIn
		End If
	End If

	sUrlTemp=Replace(sUrlTemp, "&&", "&")
	sUrlTemp=Replace(sUrlTemp, "?&", "?")
	If Right(sUrlTemp,1)="?" Then sUrlTemp=Left(sUrlTemp, Len(sUrlTemp)-1)
	If Right(sUrlTemp,1)="&" Then sUrlTemp=Left(sUrlTemp, Len(sUrlTemp)-1)
ReplaceUrlParams=sUrlTemp
End Function

Function RemoveUrlParamValue(sUrlIn, sParamIn, sValue)
Dim sUrlTemp
	sUrlTemp=sUrlIn
	sUrlTemp=Replace(sUrlTemp, "&" & sParamIn & "=" & sValue, "")
	sUrlTemp=Replace(sUrlTemp, sParamIn & "=" & sValue & "&", "")
	sUrlTemp=Replace(sUrlTemp, "?" & sParamIn & "=" & sValue, "?")
RemoveUrlParamValue=sUrlTemp
End Function

'--------------------------------------------------------------------
' Function for adding params to the url string
'--------------------------------------------------------------------
Function AddUrlParams(sUrlIn, sParamIn)
Dim sUrlTemp
	sUrlTemp=ReplaceUrlParams(sUrlIn, sParamIn)
	AddUrlParams=sUrlTemp
End Function


Function EncodeUrlParams(sParamsIn)
Dim sTempParams
	sTempParams=Replace(sParamsIn, "?", "&")
	sTempParams=ReplaceUrlParams(sTempParams, "url")
	sTempParams=ReplaceUrlParams(sTempParams, "sid")
EncodeUrlParams=sTempParams
End Function



%>