<%@ LANGUAGE="VBSCRIPT" %>
<% Option Explicit %>
<!--#include virtual="dbc.asp"-->

<%
    Dim sDatabase
    sDatabase = Request.Form("database")

    objConn.Close
    objConn.Open "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Data Source=MAGNESA;Initial Catalog=" & sDatabase & ";"

    Dim Id, sLanguage
    id = Request.Form("expertId")
    sLanguage = Request.Form("language")

    Dim iResult
    
    Set objTempRs = GetDataRecordsetSP("usp_ExpertCheckIdExistWithLanguageSelect", Array( _
        Array(, adInteger, , id), _
        Array(, adVarChar, 5, sLanguage)))

    If Not objTempRs.Eof Then
        iResult = objTempRs(0)
        Response.Write(iResult)
    Else 
        iResult = 0
        Response.Write(iResult)
    End If 

%>