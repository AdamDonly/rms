<%@ LANGUAGE="VBSCRIPT" %>
<% Option Explicit %>

<!--#include file="../../dbc.asp"-->
<!--#include file="../../fnc.asp"-->

<% 
Dim aNt, aNtInt, aGZ, strSQL, objrs1, objRs2, j

Set objRs1=Server.CreateObject("ADODB.Recordset")
Set objRs2=Server.CreateObject("ADODB.Recordset")



	aNt=0
	Dim aNtInfo(300)
	Dim aNtCode(300)
	Dim aNtZone(300)

	i=0
	aNtInt=""

Rem Append data in arrays: mGZnInfo - Geo Zones names, mGZnCode - GeoZones codes

	strSQL="SELECT DISTINCT tbl_GeoZone.id_GeoZone, tbl_GeoZone.id_Continent, tbl_GeoZone.Geo_ZoneEng, tbl_GeoZone.db_Scroll FROM (tbl_GeoZone LEFT JOIN tbl_Country ON tbl_GeoZone.id_GeoZone = tbl_Country.id_GeoZone) WHERE db_NotVisible=0 order by 2,3"
	objrs1.Open strSQL,objconn,3,3
	aGZ=objrs1.RecordCount

	Dim aGZnInfo(30)
	Dim aGZnCode(30)
	Dim aGZnContinent(30)
	Dim aGZnScroll(30)
	i=0
	
	Do Until objrs1.EOF 
	aGZnInfo(i)=objrs1("Geo_ZoneEng")
	aGZnCode(i)=objrs1("id_GeoZone")
	aGZnContinent(i)=objrs1("id_Continent")

		objrs2.Open "select tbl_Country.id_Country, tbl_Country.couNameEng, tbl_Country.id_GeoZone FROM tbl_Country WHERE id_GeoZone=" & objrs1("id_GeoZone") & " and id_Country<>524 and id_Country<1000 order by 2;",objconn,3,3
		j=0

		Do Until objrs2.EOF 
		aNtInfo(j+aNt)=objrs2("couNameEng")
		aNtCode(j+aNt)=objrs2("id_Country")
		aNtZone(j+aNt)=objrs2("id_GeoZone")

	Response.Write "jNtCode[" & (j+aNt+1) & "]=" & aNtCode(j+aNt) & "; jNtZone[" & (j+aNt+1) & "]=" & (i+1) & "; jNtInt[" & (j+aNt+1) & "]=0;" & "<br />"


		j=j+1
		objrs2.MoveNext
		Loop
		aNt=aNt+objrs2.Recordcount
		objrs2.Close
		aGZnScroll(i)=j
	Response.Write "<br />"

	i=i+1
	objrs1.MoveNext
	Loop
	objrs1.Close
	%>

