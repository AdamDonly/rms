<%@ LANGUAGE="VBSCRIPT" %>
<% Option Explicit %>

<!--#include file="../../dbc.asp"-->
<!--#include file="../../fnc.asp"-->

<% 
Dim aExTInfo(420)
Dim aExTCode(420)
Dim aExTSrch(420)
Dim aExF, aExT
aExF=22

Dim aExFInfo(22)
Dim aExFCode(22)
Dim aExFShort(22)
Dim aExFScroll(22)
Dim aExFShift(22)

Dim strSQL, objrs1, objRs2, iMainSector, iSector

Set objRs1=Server.CreateObject("ADODB.Recordset")
Set objRs2=Server.CreateObject("ADODB.Recordset")

aExT=0

' Get the list of main sectors

strSQL="SELECT id_MainSector, mnsDescriptionEng, mnsShortEng, mnsAbbreviation, db_Scroll FROM tbl_MainSectors WHERE db_NotVisible=0 ORDER BY mnsDescriptionEng"
objrs1.Open strSQL,objconn,3,3
aExF=objrs1.RecordCount

iMainSector=0

	While Not objrs1.Eof
	aExFInfo(iMainSector)=objrs1("mnsDescriptionEng")
	aExFShort(iMainSector)=objrs1("mnsShortEng")
	aExFCode(iMainSector)=objrs1("id_MainSector")

	Response.Write "<br />"
	'Response.Write "aExFInfo(" & (iMainSector) & ")=""" & aExFInfo(iMainSector) & """" & vbCrLf & "<br />"
	'Response.Write "aExFShort(" & (iMainSector) & ")=""" & aExFShort(iMainSector) & """" & vbCrLf & "<br />"
	'Response.Write "aExFCode(" & (iMainSector) & ")=" & aExFCode(iMainSector) & vbCrLf & "<br />"

	objrs2.Open "SELECT S.id_Sector, S.sctDescriptionEng, S.id_MainSector FROM tbl_Sectors S WHERE (id_Sector<1000 or id_Sector>1021) AND S.id_MainSector=" & objrs1("id_MainSector") & " ORDER BY S.sctDescriptionEng ",objconn,3,3
	iSector=0
	aExFShift(iMainSector)=0

		While Not objrs2.Eof 
		aExTInfo(iSector+aExT)=objrs2("sctDescriptionEng")
		aExTCode(iSector+aExT)=objrs2("id_Sector")
		aExTSrch(iSector+aExT)=objrs2("id_MainSector")

		Response.Write "jExTCode[" & (iSector+aExT+1) & "]=" & aExTCode(iSector+aExT) & "; jExTSrch[" & (iSector+aExT+1) & "]=" & iMainSector+1 & "; jExTInt[" & (iSector+aExT+1) & "]=0;<br />"

		aExFShift(iMainSector)=aExFShift(iMainSector) + (Len(aExTInfo(iSector+aExT)) \ 56)
		
		iSector=iSector+1
		objrs2.MoveNext
		WEnd
		
	Response.Flush
	aExT=aExT + objrs2.RecordCount
	objrs2.Close
	aExFScroll(iMainSector)=iSector
	
	iMainSector=iMainSector+1
	objrs1.MoveNext
	WEnd
	
objrs1.Close
Response.Write "<br />"
	
%>

