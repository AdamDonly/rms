<!--#include virtual="/dbc.asp"-->
<!--#include virtual="/fnc.asp"-->
<%
Response.Buffer = True
' Check the validity of the session
If sApplicationName <> "expert" Then
	CheckUserLogin sScriptFullNameAsParams
End If
CheckExpertID
%>
<!--#include virtual="/_common/_class/document.asp"-->
<!--#include virtual="/_common/_grid/document_list.asp"-->
<%
Dim uDocumentUID
Dim sFileName, sFileExtension, binFileData, sContentType
uDocumentUID = ReplaceIfEmpty(Request.QueryString("document"), Request.QueryString("uid"))

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
