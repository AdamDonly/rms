<!--#include file="dbc.asp"-->
<!--#include file="fnc.asp"-->

<%
Session.Abandon
iUserID=0
iMemberID=0
iExpertID=0
sSessionID=""
sCookiesSessionID=""
Response.Cookies("SessionID")=""

Dim icaLogoutUrl, refid
icaLogoutUrl = Request.QueryString("ica")
refid = Request.QueryString("refid")

If icaLogoutUrl > "" Then
    sUrl = icaLogoutUrl
    If refid > "" Then
        sUrl = sUrl & "?refid=" & refid
    End If
	
    Response.Redirect(sUrl)
Else
    Response.Write "Error. Could not do a clean logout."
End If
Response.End
%>