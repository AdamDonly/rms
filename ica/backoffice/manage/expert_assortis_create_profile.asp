<%@ LANGUAGE="VBSCRIPT" %>
<% Option Explicit %>
<!--#include virtual="dbc.asp"-->

<%
    objConn.Close
    objConn.Open "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Data Source=MAGNESA;Initial Catalog=assortis2db;"
    Dim iPrimaryExpertID, iExistingId, sLang, bInitalCreate, sProc
    iPrimaryExpertID = Request.Form("expertId")

    If Request.Form("existingId") <> "" Then
        iExistingId = Request.Form("existingId")
    Else 
        iExistingId = Null
    End If

    sLang = Request.Form("language")
    bInitalCreate = Request.Form("initalCreate")
    
    Dim iResult

    sProc = "usp_ExpertCreateExpertLanguageInsert"
    Set objTempRs = GetDataRecordsetSP(sProc, Array( _
        Array(, adInteger, , iPrimaryExpertID), _
        Array(, adVarChar, 5, sLang), _
        Array(, adInteger, , iExistingId) ))
    

%>