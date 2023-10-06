<%@ LANGUAGE="VBSCRIPT" %>
<% Option Explicit 
'--------------------------------------------------------------------
'
' Select experts on the search query
'
'--------------------------------------------------------------------
' Don't cache the page
Response.Expires = 0
Response.ExpiresAbsolute = Now() - 1
Response.CacheControl = "no-cache"
%>
<!--#include file="../../dbc.asp"-->
<!--#include file="../../fnc.asp"-->
<!--#include file="../../fnc_exp.asp"-->
<% 
' Checking for user's login 
CheckUserLogin sScriptFullNameAsParams

Dim iSearchQueryID
iSearchQueryID=CheckInt(Request.QueryString("qid"))

Dim iTempCvID
iTempCvID=CheckInt(Request.QueryString("eid"))

sAction=CheckInt(Request.QueryString("act"))

If sAction=0 Or sAction=1 Then
	UpdateMemberExpertQuery iMemberID, iTempCvID, iSearchQueryID, sAction
End If

If sAction=2 Then
	Response.Write ShowMemberExpertQuery(iMemberID, iTempCvID, iSearchQueryID)
End If
%>
