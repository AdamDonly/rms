<%
'--------------------------------------------------------------------
'
' Comments on expert profile
'
'--------------------------------------------------------------------
Response.Expires = 0
Response.ExpiresAbsolute = Now() - 1
Response.CacheControl = "no-cache" 
%>
<!--#include file="../_data/datMonth.asp"-->
<%
If sApplicationName<>"expert" Then
	CheckUserLogin sScriptFullNameAsParams
End If
CheckExpertID
%>
<!--#include file="../expProfile.asp"-->
<%
Dim sComments

If Request.Form()>"" Then
	iExpertID=CheckInteger(Request.Form("id_Expert"))
	sComments=Left(CheckString(Request.Form("exp_comments")), 40000)
	objTempRs=UpdateRecordSP("usp_ExpCvvCommentsUpdate", Array( _
		Array(, adInteger, , iExpertID), _
		Array(, adLongVarWChar, 40000, sComments)))

	Response.Redirect "register6.asp" & sParams
End If                                                       
%>

<%
' Getting personal data from DB
Set objTempRs=GetDataRecordsetSP("usp_ExpCvvExpInfoSelect", Array( _
	Array(, adInteger, , iExpertID)))
If Not objTempRs.Eof Then 

	iPersonID=objTempRs("id_Person")
	sFirstName=objTempRs("psnFirstNameEng")
	sMiddleName=objTempRs("psnMiddleNameEng")
	sLastName=objTempRs("psnLastNameEng")
	iTitleID=objTempRs("id_psnTitle")
	sComments=objTempRs("expComments")

End If 
objTempRs.Close		
%>  
<html>
<head>
<title>CV comments</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" type="text/css" href="<% =sHomePath %>styles.css">
</head>

<body bgcolor="#FFFFFF" topmargin=0 leftmargin=0  marginheight=0 marginwidth=0>
<% ShowTopMenu %>
<% ShowRegistrationProgressBar "CV", 8 %>

  <!-- [i] CV online -->
<% ShowMessageStart "info", 440 %>
Please always specify your name and date before any comment.
<br>
<% ShowMessageEnd %>

  <!-- Personal information -->
	<% InputFormHeader 580, "PERSONAL INFORMATION" %>
	<% InputBlockHeader "100%" %>
	<form method="post" action="<% =sScriptFullName %>">
	<input type="hidden" name="id_Expert" value="<%=iExpertID%>">
		<% InputBlockSpace 4 %>
		<% InputBlockElementLeftStart %><p class="ftxt">Expert</p><% InputBlockElementLeftEnd %><% InputBlockElementMiddle %>
		<% InputBlockElementRightStart %>&nbsp;&nbsp;<input type="text" name="exp_name_2" readOnly disabled size=31 style="width=355px;" maxlength=255 value="<% =sLastName & ", " & sFirstName & " " & sMiddleName %> "><% InputBlockElementRightEnd %>
		<% InputBlockElementLeftStart %><p class="ftxt">Comments</p><% InputBlockElementLeftEnd %><% InputBlockElementMiddle %>
		<% InputBlockElementRightStart %>&nbsp;&nbsp;<textarea cols="34" style="width=355px;" name="exp_comments" rows=12 wrap="yes"><%=sComments%></textarea><% InputBlockElementRightEnd %>
		<% InputBlockSpace 4 %>
	<% InputBlockFooter %>
	<% InputFormFooter %>
	<% InputFormSpace 12 %>

	<table width=580 cellspacing=0 cellpadding=0 border=0 align="center">
	<tr>
        <td width="300" align="right"><img src="../image/x.gif" width=170 height=1><input type="image" src="<% =sHomePath %>image/bte_savecont.gif" name="Save & continue" alt="Save & continue" border=0></td>
	</tr>
	</form>
	</table><br>

<% CloseDBConnection %>
</body>
</html>
