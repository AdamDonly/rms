<%@ LANGUAGE="VBSCRIPT" %>
<% Option Explicit %>

<!--#include file="../../dbc.asp"-->
<!--#include file="../../fnc.asp"-->

<% 
Dim aOrg, strSQL, objrs1, j

Set objRs1=Server.CreateObject("ADODB.Recordset")

	aOrg=0
	Dim aOrgInfo()
	Dim aOrgCode()
	Dim aOrgMainDonor()
	i=1

	strSQL="SELECT DISTINCT D.id_Organisation, D.orgNameEng, D.orgAbbreviation, D.orgMainDonor, D.orgVisibleDonor FROM tbl_Donors D WHERE D.orgVisibleDonor=1 order by D.orgMainDonor DESC, D.orgAbbreviation, D.orgNameEng"
	objrs1.Open strSQL, objconn, adOpenStatic, adLockReadOnly
	aOrg=objrs1.RecordCount

	ReDim aOrgInfo(aOrg)
	ReDim aOrgCode(aOrg)
	ReDim aOrgMainDonor(aOrg)
	
	Do Until objrs1.EOF 
		aOrgInfo(i)=objrs1("orgNameEng")
		aOrgCode(i)=objrs1("id_Organisation")
		aOrgMainDonor(i)=CInt(objrs1("orgMainDonor"))

		Response.Write "jOrgCode[" & (i) & "]=" & aOrgCode(i) & "; jOrgInt[" & (i) & "]=0; jOrgMain[" & (i) & "]=" & aOrgMainDonor(i) & ";" & vbCrLf & "<br />"

		i=i+1
		objrs1.MoveNext
	Loop
	objrs1.Close
	%>

