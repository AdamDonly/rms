<% 
'--------------------------------------------------------------------
'
' CV registration.
' Save status
'
'--------------------------------------------------------------------
Response.Expires = 0
Response.ExpiresAbsolute = Now() - 1
Response.CacheControl = "no-cache" 

If Request.Form()>"" Then
	Dim iExpertStatusCVID
	iExpertStatusCVID=CheckInteger(Request.Form("cv_status"))

	' Save
	If Not IsNull(iExpertStatusCVID) Then
		objExpertStatusCV.Status.ID=iExpertStatusCVID
		objExpertStatusCV.DateModified=Now()
		objExpertStatusCV.SaveData
	End If
	
	Response.Redirect sScriptFullName
End If
%>
