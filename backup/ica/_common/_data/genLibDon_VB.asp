<%@ LANGUAGE="VBSCRIPT" %>
<% Option Explicit %>

<!--#include file="../../dbc.asp"-->
<!--#include file="../../fnc.asp"-->

<% 
Dim aOrg, strSQL, objrs1, j

Set objRs1=Server.CreateObject("ADODB.Recordset")

	aOrg=0
	Dim aOrgInfo()
	Dim aOrgAbbreviation()
	Dim aOrgCode()
	Dim aOrgMainDonor()
	i=0

	strSQL="SELECT DISTINCT D.id_Organisation, D.orgNameEng, D.orgAbbreviation, D.orgMainDonor, D.orgVisibleDonor FROM tbl_Donors D WHERE D.orgVisibleDonor=1 order by D.orgMainDonor DESC, D.orgAbbreviation, D.orgNameEng"
	objrs1.Open strSQL, objconn, adOpenStatic, adLockReadOnly
	aOrg=objrs1.RecordCount

	ReDim aOrgInfo(aOrg)
	ReDim aOrgAbbreviation(aOrg)
	ReDim aOrgCode(aOrg)
	ReDim aOrgMainDonor(aOrg)
	
	Response.Write "Dim aOrg" & vbCrLf & "<br />"
	Response.Write "aOrg=" & aOrg & vbCrLf & "<br />"
	Response.Write "Dim aOrgInfo(" & aOrg & ")" & "<br />"
	Response.Write "Dim aOrgCode(" & aOrg & ")" & "<br />"
	Response.Write "Dim aOrgMainDonor(" & aOrg & ")" & "<br /><br />"


	Do Until objrs1.EOF 
		aOrgInfo(i)=objrs1("orgNameEng")
		aOrgAbbreviation(i)=objrs1("orgAbbreviation")
		If Len(aOrgInfo(i) & aOrgAbbreviation(i))> 50 Then 
			aOrgInfo(i)=CutStringInMenu(aOrgInfo(i), 50-Len(aOrgAbbreviation(i)), " ", "&lt;br&gt;&amp;nbsp; &amp;nbsp; &amp;nbsp;")
		End If

		aOrgCode(i)=objrs1("id_Organisation")
		aOrgMainDonor(i)=Abs(CInt(objrs1("orgMainDonor")))

		Response.Write "aOrgInfo(" & (i) & ")=""" & aOrgAbbreviation(i) & " - " & aOrgInfo(i) & """" & vbCrLf & "<br />"
		Response.Write "aOrgCode(" & (i) & ")=" & aOrgCode(i) & vbCrLf & "<br />"
		Response.Write "aOrgMainDonor(" & (i) & ")=" & aOrgMainDonor(i) & vbCrLf & "<br />"

		i=i+1
		objrs1.MoveNext
	Loop
	objrs1.Close
	%>

