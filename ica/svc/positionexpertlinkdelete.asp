<%@ LANGUAGE="VBSCRIPT" %>
<% Option Explicit %>

<!--#include file="../dbc.asp"-->
<!--#include file="../fnc.asp"-->
<!--#include file="../fnc_exp.asp"-->
<%
Dim iPositionId, iExpertDatabaseId, iLinkValue

iPositionId = CheckIntegerAndNull(Request.QueryString("positionid"))
iExpertId = CheckIntegerAndNull(Request.QueryString("expertId"))
sExpertUid = CheckString(Request.QueryString("expertUid"))
iExpertDatabaseId = CheckIntegerAndNull(Request.QueryString("expertDatabaseId"))
iLinkValue = CheckIntegerAndNull(Request.QueryString("linkValue"))

If iPositionId > 0 And iExpertId > 0 And sExpertUid > "" Then
    UpdateRecordSP "usp_" & sIcaServerSqlPrefix & "PositionExpertLinkDelete", Array( _
        Array(, adInteger, , iPositionId), _
        Array(, adInteger, , iExpertId), _
        Array(, adVarChar, 40, sExpertUid))
    Response.Clear
    Response.Write "OK"
End If
CloseDBConnection
Response.End
%>