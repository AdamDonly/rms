<%

Dim iLogResult

Function GenerateGuid()
	Dim typeLib
	Set typeLib = Server.CreateObject("Scriptlet.TypeLib")
	GenerateGuid = typeLib.Guid
	Set typeLib = Nothing
End Function

' Activity IDs for ICA CVIP:
' 33 - Add new expert
' 34 - Update expert
' 35 - Download CV
' 36 - Search experts database
' 37 - View CV Anonymously

Function LogActivity(iActivityId, sRequest, sResponse, gGuid)
	Dim iRet
	iRet = 0

	LogActivity = iRet
End Function



%>
