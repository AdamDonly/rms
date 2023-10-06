<% 
Dim objConnMpis

Set objConnMpis = Server.CreateObject("ADODB.Connection")
'objConnMpis.Open "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Data Source=IBFSQL\DBS;Initial Catalog=pm2db;"
objConnMpis.Open "Provider=SQLOLEDB.1;uid=su;pwd=sysibfassortisavh;Data Source=IBFSQL\DBS;Initial Catalog=pm2db;"

%>
