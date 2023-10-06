<%
Response.Expires = 0
Response.ExpiresAbsolute = Now() - 1
Response.CacheControl = "no-cache" 
%>
<!--#include file="../_data/datMonth.asp"-->
<!--#include file="../_class/document.asp"-->
<!--#include file="../_grid/document_list.asp"-->
<%
If sApplicationName<>"expert" Then
	CheckUserLogin sScriptFullNameAsParams
End If
CheckExpertID
%>
<!--#include file="../expProfile.asp"-->
<%
Dim objDocument, sDocumentUid
Set objDocument = New CDocument
sDocumentUid = Request.QueryString("document")

If sAction="delete" Then
	objDocument.DeleteByUid sDocumentUid
	Response.Redirect "register6.asp?id=" & iExpertID
End If

' Save document on submit
Dim objUploadForm
Set objUploadForm = Server.CreateObject("softartisans.fileup")
If objUploadForm.ContentDisposition = "form-data" Then
	If objUploadForm.TotalBytes > 5120000  Then 
		ShowMessageStart "error", 580 %>
			Your file is too big.</b><br>Please try to keep the size of the file within the allowed 5 Mb. Click back and try again.
		<% ShowMessageEnd 
	Else
		objDocument.SaveForm iExpertID, objUploadForm.Form
		Response.Redirect "register6.asp?id=" & iExpertID
	End If
End If

' Get document details
If Len(sDocumentUid)>36 Then
	objDocument.LoadDocumentDetailsByUid(sDocumentUid)
End If

' Get expert personal details
Set objTempRs=GetDataRecordsetSP("usp_ExpCvvExpInfoSelect", Array( _
	Array(, adInteger, , iExpertID)))
If Not objTempRs.Eof Then 
	iPersonID=objTempRs("id_Person")
	sFirstName=objTempRs("psnFirstNameEng")
	sMiddleName=objTempRs("psnMiddleNameEng")
	sLastName=objTempRs("psnLastNameEng")
	iTitleID=objTempRs("id_psnTitle")
End If 
objTempRs.Close
%>  
<html>
<head>
<title>Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" type="text/css" href="<% =sHomePath %>styles.css">
<script type="text/javascript" src="/_scripts/js/main.js"></script>
<script type="text/javascript">
<!-- 
function validateForm() {
<%
Dim sUserSalutation
If sApplicationName="expert" Then
	sUserSalutation="your"
Else
	sUserSalutation="expert's"
End If
%>
	var f=document.forms[0];
	if (!(f)) {
		return false; 
	}
	if (!checkTextFieldValue(f.document_title, "", "Please fill in the document title.", 1)) { return false }
	<% If objDocument.ID>0 Then %>
	<% Else %>
	if (!checkTextFieldValue(f.attachment, "", "Please attach a document.", 1)) { return false }
	<% End If %>
	f.submit();
}

function deleteDocument(expert, document) {
	if (confirm('Are you sure you want to delete this document?')) {
		location.replace('<% =sScriptFileName & AddUrlParams(sParams, "act=delete") %>');
	}
}
-->
</script>
</head>
<body bgcolor="#FFFFFF" topmargin=0 leftmargin=0  marginheight=0 marginwidth=0>
<% ShowTopMenu %>
<% ShowRegistrationProgressBar "CV", 8 %>
<br>

  <!-- Personal information -->
	<% InputFormHeader 580, "UPLOAD DOCUMENT" %>
	<% InputBlockHeader "100%" %>
	<form enctype="multipart/form-data" method="post" action="<% =sScriptFullName %>" onsubmit="validateForm(); return false;">
	<input type="hidden" name="expert" value="<% =iExpertID %>">
	<input type="hidden" name="document" value="<% =objDocument.UID %>">
		<% InputBlockSpace 4 %>
		<% InputBlockElementLeftStart %><p class="ftxt">Expert</p><% InputBlockElementLeftEnd %><% InputBlockElementMiddle %>
		<% InputBlockElementRightStart %>&nbsp;&nbsp;<input type="text" name="exp_name_2" readOnly disabled size=31 style="width:320px;" maxlength=255 value="<% =sLastName & ", " & sFirstName & " " & sMiddleName %> "><% InputBlockElementRightEnd %>
		<% InputBlockElementLeftStart %><p class="ftxt">Document title</p><% InputBlockElementLeftEnd %><% InputBlockElementMiddle %>
		<% InputBlockElementRightStart %>&nbsp;&nbsp;<input type="text" style="width:320px;" name="document_title" value="<% =objDocument.Title %>">&nbsp;&nbsp;<span class="fcmp">*</span><% InputBlockElementRightEnd %>
		<% InputBlockElementLeftStart %><p class="ftxt">Document type</p><% InputBlockElementLeftEnd %><% InputBlockElementMiddle %>
		<% InputBlockElementRightStart %>&nbsp;&nbsp;<input type="text" style="width:320px;" name="document_type" value="<% =objDocument.Type_ %>"><% InputBlockElementRightEnd %>
		<% InputBlockElementLeftStart %><p class="ftxt">Attachment</p><% InputBlockElementLeftEnd %><% InputBlockElementMiddle %>
		<% InputBlockElementRightStart %>
		<p style="margin:2px 12px">
		<% If objDocument.ID>0 Then %>
			<% ShowDocument objDocument %>
			<% If IsDate(objDocument.DateCreated) Then %>
				<small>&nbsp;&nbsp;(Document uploaded on <% =ConvertDateForText(objDocument.DateCreated, "&nbsp;", "DDMonYYYY HHMM") %>)</small>
			<% End If %>
			<br/><img src="../image/x.gif" width="1" height="5"><br/><small>To update this document please attach a new file hereafter:</small><br/>
		<% Else %>
			<small>Please click browse and attach a document hereafter:</small><br/>
		<% End If %>
		</p>&nbsp;&nbsp;<input type="file" style="width:317px;" name="attachment">
		<% If objDocument.ID>0 Then %>
		<% Else %>
		&nbsp;&nbsp;<span class="fcmp">*</span>
		<% End If %>
		<% InputBlockElementRightEnd %>
		<% InputBlockSpace 4 %>
	<% InputBlockFooter %>
	<% InputFormFooter %>
	<% InputFormSpace 12 %>

	<table width=580 cellspacing=0 cellpadding=0 border=0 align="center">
	<tr>
        <td width="300" align="left"><img src="../image/x.gif" width=170 height=1 align="left">
		<% If objDocument.ID>0 Then %>
		<input type="image" src="<% =sHomePath %>image/bte_docupdate.gif" name="Update document" alt="Update document" border=0 align="left">
		<a href="javascript:deleteDocument(<% =iExpertID %>, '<% =objDocument.UID %>')"><img src="<% =sHomePath %>image/bte_docdelete.gif" name="Delete document" alt="Delete document" border=0 hspace=60>
		<% Else %>
		<input type="image" src="<% =sHomePath %>image/bte_docsave.gif" name="Save document" alt="Save document" border=0>
		<% End If %>
		</td>
	</tr>
	</form>
	</table><br>

<% CloseDBConnection %>
</body>
</html>
