<html>
<head><meta http-equiv="Content-Type" content="text/html; charset=utf-8"></head>
<body>
<%
' The ASP String component (CkString) can be downloaded from:
' http://www.chilkatsoft.com/download/CkString.zip

' HTML entity decode
set cks = Server.CreateObject("Chilkat.String")
cks.Str = "�Trade and transport of C&#1040; Development�, �Monitoring of transport corridors of C&#1040;�"
cks.HtmlEntityDecode

' The string now contains: <p> e��� e��� </p>
' Prints: e��� e���
Response.Write cks.Str

%>
</body>
</html>