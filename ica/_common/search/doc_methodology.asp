<!--#include file="../../_fnc_date.asp"-->
<%
If sApplicationName <> "expert" Then
	CheckUserLogin sScriptFullNameAsParams
End If
CheckExpertID
%>
<%
Dim objDocument, sDocumentUid
Set objDocument = New CDocument
sDocumentUid = Request.QueryString("document")

If sAction="delete" Then
	objDocument.DeleteByUid sDocumentUid
	Response.Redirect "exp_methodology.asp" & ReplaceUrlParams(sParams, "document")
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
		objDocument.SaveForm iCvID, objUploadForm.Form
		Response.Redirect "exp_methodology.asp" & ReplaceUrlParams(sParams, "document")
	End If
End If

' Get document details
If Len(sDocumentUid)>36 Then
	objDocument.LoadDocumentDetailsByUid(sDocumentUid)
End If

%>  
<!--#include virtual="/_template/html.header.start.asp"-->
<script type="text/javascript" src="/_scripts/js/main.js"></script>
<script type="text/javascript">
<!-- 
function validateForm() {
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
<body style="width: 90%; margin-left: 5%;">

	<!-- content -->
	<div id="content" class="searchform">
	<% 
	'ShowRegistrationProgressBar "CV", 8 
	%>

	<form enctype="multipart/form-data" method="post" action="<% =sScriptFullName %>" onsubmit="validateForm(); return false;">
		<input type="hidden" name="expert" value="<% =sExpertUID %>">
		<input type="hidden" name="document" value="<% =objDocument.UID %>">
		<input type="hidden" name="document_typeid" value="<% =ReplaceIfEmpty(objDocument.TypeID, aDocumentTypeIdMethodologySupport) %>">
		<input type="hidden" name="document_type" value="<% =objDocument.Type_ %>">

		<div class="box search blue">
		<h3><% =GetLabel(sCvLanguage, "Upload document") %></h3>
		<table class="search_form" width="100%" cellspacing=0 cellpadding=0 border=0>
		<tr>
		<td class="field splitter"><% = GetLabel(sCvLanguage, "Expert") %></td>
		<td class="value blue"><input type="text" name="exp_name_2" readOnly disabled size=31 style="width:320px;" maxlength=255 value="<% =sLastName & ", " & sFirstName & " " & sMiddleName %> "></td>
		</tr>
		<tr>
		<td class="field splitter"><label for="document_title"><% = GetLabel(sCvLanguage, "Document title") %></label></td>
		<td class="value blue"><input type="text" style="width:320px;" name="document_title" value="<% =objDocument.Title %>">&nbsp;&nbsp;<span class="fcmp">*</span></td>
		</tr>
		<tr>
		<td class="field splitter"><label for="document_title"><% = GetLabel(sCvLanguage, "Attachment") %></label></td>
		<td class="value blue">
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
		</td>
		</tr>
		<tr style="height: 4pt;">
		<td class="field splitter"></td>
		<td class="value blue"></td>
		</tr>
		</table>
		</div>

		<div class="spacebottom">
		<% If objDocument.ID>0 Then %>
		<input type="submit" class="red-button under-right-col w125" src="<% =sHomePath %>image/bte_docupdate.gif" name="Update document" value="Update document" />
		<a href="javascript:deleteDocument(<% =iCvID %>, '<% =objDocument.UID %>')" class="red-button next-btn w125">Delete document</a>
		<% Else %>
		<input type="submit" class="red-button under-right-col w125" src="<% =sHomePath %>image/bte_docsave.gif" name="Save document" value="Save document" />
		<% End If %>
		</div>
		</form>

	</div>

<!-- footer -->
<!--#include virtual="/_template/page.footer.asp"-->

<% CloseDBConnection %>
</body>
<!--#include virtual="/_template/html.footer.asp"-->
