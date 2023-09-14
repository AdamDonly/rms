<%@ LANGUAGE="VBSCRIPT" %>
<% Option Explicit %>

<!--#include file="../dbc.asp"-->
<!--#include file="../fnc.asp"-->
<!--#include file="../fnc_exp.asp"-->
<%
Dim iNewExpertManagerID, iCircleID

iNewExpertManagerID = CheckIntegerAndNull(Request.QueryString("iuserid"))
iCircleID = CheckIntegerAndNull(Request.QueryString("icircleid"))   

If iNewExpertManagerID > 0 And iCircleID > 0 Then
    UpdateRecordSP "usp_" & sIcaServerSqlPrefix & "UpdateExpertCircleExpertManager", Array( _
        Array(, adInteger, , iNewExpertManagerID), _
		Array(, adInteger, , iCircleID))
    Response.Clear
End If
CloseDBConnection

Response.ContentType = "application/json"
Response.Write("{ ""circleid"": " & iCircleID & ", ""userid"": " & iNewExpertManagerID & "  }")

%>
