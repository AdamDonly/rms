<%@ LANGUAGE="VBSCRIPT" %>
<% Option Explicit %>
<!--#include virtual="dbc.asp"-->

<%
    Dim sDatabase
    sDatabase = Request.Form("database")

    objConn.Close
    objConn.Open "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Data Source=MAGNESA;Initial Catalog=" & sDatabase & ";"

    Dim Id, sLanguage
    id = Request.Form("expertLanguageId")
    sLanguage = Request.Form("language")

    Dim iResult
    
    Set objTempRs = GetDataRecordsetSP("usp_ExpertRemoveLinkUpdate", Array( _
        Array(, adInteger, , id), _
        Array(, adVarChar, 5, sLanguage)))

    
    Set objTempRs = Nothing
    objConn.Close
%>