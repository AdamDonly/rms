<%
Dim sIcaServer, sIcaServerType
sIcaServer = Request.ServerVariables("SERVER_NAME")
If InStr(sIcaServer, "test")>0 Then
	sIcaServerType="test."
Else
	sIcaServerType="www."
End If
sIcaServer=sIcaServerType & "icaworld.net"

sTempParams=ReplaceUrlParams(sParams, "url")
sTempParams=ReplaceUrlParams(sTempParams, "id")
sTempParams=ReplaceUrlParams(sTempParams, "idproject")
sTempParams=ReplaceUrlParams(sTempParams, "idexpert")
sTempParams=ReplaceUrlParams(sTempParams, "t")
%>

<div id="header">
<a href="/"><img src="http://www.ibf.be/Resources/images/ibf_nlogo.gif" width="171" height="114" vspace="10"></a><br />
	
