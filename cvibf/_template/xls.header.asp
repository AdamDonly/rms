<%
Dim sFileName
sFileName = "results.xls"

Response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=" & sFileName
%>
