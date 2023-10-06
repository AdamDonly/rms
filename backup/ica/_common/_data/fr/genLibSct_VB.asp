<%@ LANGUAGE="VBSCRIPT" %>
<% Option Explicit %>

<!--#include file="../../dbc.asp"-->
<!--#include file="../../fnc.asp"-->

Dim aExTInfo(400)<br />
Dim aExTCode(400)<br />
Dim aExTSrch(400)<br />
Dim aExF, aExT<br />
aExF=23<br /><br />

Dim aExFInfo(23)<br />
Dim aExFCode(23)<br />
Dim aExFShort(23)<br />
Dim aExFScroll(23)<br />
Dim aExFShift(23)<br />

<% 
Dim aExTInfo(400)
Dim aExTCode(400)
Dim aExTSrch(400)
Dim aExF, aExT
aExF=23

Dim aExFInfo(23)
Dim aExFCode(23)
Dim aExFShort(23)
Dim aExFScroll(23)
Dim aExFShift(23)

Dim strSQL, objrs1, objRs2, iMainSector, iSector

Set objRs1=Server.CreateObject("ADODB.Recordset")
Set objRs2=Server.CreateObject("ADODB.Recordset")

aExT=0

' Get the list of main sectors

strSQL="SELECT id_MainSector, mnsDescriptionFra, mnsShortFra, mnsAbbreviation, db_Scroll FROM tbl_MainSectors WHERE db_NotVisible=0 ORDER BY mnsDescriptionEng"
objrs1.Open strSQL,objconn,3,3
aExF=objrs1.RecordCount

iMainSector=0

	While Not objrs1.Eof
	aExFInfo(iMainSector)=objrs1("mnsDescriptionFra")
	aExFShort(iMainSector)=objrs1("mnsShortFra")
	aExFCode(iMainSector)=objrs1("id_MainSector")

	Response.Write "<br />"
	Response.Write "aExFInfo(" & (iMainSector) & ")=""" & aExFInfo(iMainSector) & """" & vbCrLf & "<br />"
	Response.Write "aExFShort(" & (iMainSector) & ")=""" & aExFShort(iMainSector) & """" & vbCrLf & "<br />"
	Response.Write "aExFCode(" & (iMainSector) & ")=" & aExFCode(iMainSector) & vbCrLf & "<br />"

	objrs2.Open "SELECT S.id_Sector, S.sctDescriptionFra, S.id_MainSector FROM tbl_Sectors S WHERE (id_Sector<1000 or id_Sector>1021) AND S.id_MainSector=" & objrs1("id_MainSector") & " ORDER BY S.sctDescriptionEng ",objconn,3,3
	iSector=0
	aExFShift(iMainSector)=0

		While Not objrs2.Eof 
		aExTInfo(iSector+aExT)=objrs2("sctDescriptionFra")
		aExTCode(iSector+aExT)=objrs2("id_Sector")
		aExTSrch(iSector+aExT)=objrs2("id_MainSector")

		Response.Write "aExTInfo(" & (iSector+aExT) & ")=""" & aExTInfo(iSector+aExT) & """" & vbCrLf & "<br />"
		Response.Write "aExTCode(" & (iSector+aExT) & ")=" & aExTCode(iSector+aExT) & vbCrLf & "<br />"
		Response.Write "aExTSrch(" & (iSector+aExT) & ")=" & aExTSrch(iSector+aExT) & vbCrLf & "<br />"

		Response.Write "aExFShift(iMainSector)=" & aExFShift(iMainSector) & "<br />"
		If Len(aExTInfo(iSector+aExT))>0 Then
			aExFShift(iMainSector)=aExFShift(iMainSector) + (Len(aExTInfo(iSector+aExT)) \ 57)
		End If		
		
		iSector=iSector+1
		objrs2.MoveNext
		WEnd
		
	Response.Flush
	aExT=aExT + objrs2.RecordCount
	objrs2.Close
	aExFScroll(iMainSector)=iSector
	Response.Write "aExFScroll(" & (iMainSector) & ")=" & aExFScroll(iMainSector) & vbCrLf & "<br />"
	Response.Write "aExFShift(" & (iMainSector) & ")=" & aExFShift(iMainSector) & vbCrLf & "<br />"
	
	iMainSector=iMainSector+1
	objrs1.MoveNext
	WEnd
	
objrs1.Close
Response.Write "<br />"
Response.Write "aExT=" & aExT  & vbCrLf & "<br />"
	
%>

