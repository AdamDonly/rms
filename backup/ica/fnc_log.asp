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
	If iUserID > 0 AND iActivityId > 0 Then
		If Not IsObject(gGuid) Then gGuid = GenerateGuid()

		objTempRsLog=GetDataOutParamsSP("usp_LogIcaActivity", Array( _
			Array(, adInteger, , iUserID), _
			Array(, adBoolean, , true), _
			Array(, adTinyInt, , iActivityId), _
			Array(, adVarChar, 4000, sRequest), _
			Array(, adVarChar, 4000, sResponse), _
			Array(, adVarChar, 50, sUserIpAddress), _
			Array(, adGUID, , gGuid), _
			Array(, adChar, 1, "P")), Array( _ 
			Array(, adInteger)))

		Set objTempRsLog = Nothing
		LogActivity = 1
	End If

	LogActivity = iRet
End Function



%>
