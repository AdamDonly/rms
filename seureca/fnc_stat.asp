<% 
'--------------------------------------------------------------------
'
' Functions for selecting statistics on number Busops and Experts
'
' Last update: 10.12.2002
'
'--------------------------------------------------------------------
Dim sNoJSTotalExperts ' using for noscript list of experts by sector

Sub InsertJsStatisticOnBSC
Dim rsStat, sTotalBusops, iLoop
%>
<script language="JavaScript">
<!--
nBusops=new Array()
<%
Set rsStat=objConn.Execute("EXEC usp_UsrBscStatisticSelect")
iLoop=0
While not rsStat.Eof
	sTotalBusops=CutStringNDelete(rsStat("mnsDescriptionEng"),40) & " - <b>" & rsStat("TotalBusops") & "&nbsp;tenders</b>"
	%>
	nBusops[<%=iLoop%>]=new Array()
	nBusops[<%=iLoop%>]["text"]="<%=sTotalBusops%>"
	nBusops[<%=iLoop%>]["link"]="<%=sHomePath & "en/members/bsc_results.asp" & AddUrlParams(sParams, "srch_msectors=" & rsStat("id_MainSector"))%>&act=trial"
	<%
	rsStat.MoveNext
	iLoop=iLoop+1
WEnd 
rsStat.Close
Set rsStat=Nothing
%>
if(bw.bw) onload = fadeInit
//-->
</script>
<%
End Sub


Sub InsertJsStatisticOnEXP
Dim rsStat, sTotalExperts, iLoop, iRndSector
%>
<script language="JavaScript">
<!--
nExperts=new Array()
<%
Set rsStat=objConn.Execute("EXEC usp_UsrExpStatisticSelect")
iLoop=0
Randomize
' iRndSector is using for showing 1 random sector's statistic for browsers (search-engines) without JS support
iRndSector=Round(Rnd(20)*18)+1
While not rsStat.Eof
	sTotalExperts=CutStringNDelete(rsStat("mnsDescriptionEng"),40) & " - <b>" & rsStat("TotalExperts") & "&nbsp;experts</b>"
	' If rsStat("id_MainSector")=iRndSector Then sNoJSTotalExperts=sNoJSTotalExperts & "<p class=""sml""><a href=""" & sHomePath & "en/members/exp_results.asp?srch_cfse=1&srch_msectors=" & rsStat("id_MainSector") & """>" & sTotalExperts & "</a></p>" & vbCrLf
	sNoJSTotalExperts=sNoJSTotalExperts & "<p class=""sml""><a href=""" & sHomePath & "en/members/exp_results.asp?srch_cfse=1&srch_msectors=" & rsStat("id_MainSector") & """>" & sTotalExperts & "</a></p><img src=""" & sHomePath & "image/x.gif"" width=1 height=6><br>" & vbCrLf
	%>
	nExperts[<%=iLoop%>]=new Array()
	nExperts[<%=iLoop%>]["text"]="<%=sTotalExperts%>"
	nExperts[<%=iLoop%>]["link"]="<%=sHomePath & "en/members/exp_results.asp" & AddUrlParams(sParams, "srch_msectors=" & rsStat("id_MainSector"))%>"
	<% 
	rsStat.MoveNext
	iLoop=iLoop+1
WEnd 
sNoJSTotalExperts=sNoJSTotalExperts & "<p class=""sml"" align=""right""><a href=""" & sHomePath & "en/members/exp_results.asp?srch_cfse=1"">See all sectors</a>&nbsp;&nbsp;&nbsp;"
rsStat.Close
Set rsStat=Nothing
%>
<% If InStr(sScriptFileName,"cv_")=0 Then %>
	if(bw.bw) onload = fadeInit
<% End If %>
	//-->
</script>
<%
End Sub
%>
