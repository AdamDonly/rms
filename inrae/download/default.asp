<!--#include file="../dbc.asp"-->
<!--#include file="../fnc.asp"-->
<%
' Check the validity of the session
CheckUserLogin sScriptFullNameAsParams
%>
<!--#include virtual="/_common/_class/document.asp"-->
<!--#include virtual="/_common/_grid/document_list.asp"-->
<%
Response.Buffer = True

Dim uDocumentUID
Dim sFileName, sFileExtension, binFileData, sContentType
uDocumentUID=Request.QueryString("uid")

If Len(uDocumentUID)<36 Or Len(uDocumentUID)>40 Then 
	%>
	<!--#include virtual="/_common/_template/page.close.asp"-->
	<%
	Response.End
End If

Dim objDocument
Set objDocument = New CDocument
objDocument.LoadDocumentDetailsByUid uDocumentUID
objDocument.Download
Set objDocument = Nothing

%>
