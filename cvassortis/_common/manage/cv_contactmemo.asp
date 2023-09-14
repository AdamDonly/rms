<!--#include file="../../dbc_mpis.asp"-->
<%
CheckUserLogin sScriptFullNameAsParams

Dim objRs1, objRs2, objRs3
Set objRs1=Server.CreateObject("ADODB.Recordset")
Set objRs2=Server.CreateObject("ADODB.Recordset")

Dim iPersonID, sFirstName, sMiddleName, sLastName, iTitleID
Dim iContactID, sContactComments
iContactID=CheckInteger(Request.QueryString("id_contact"))

If Request.Form()>"" Then
	iContactID=CheckInteger(Request.Form("id_contact"))
	sContactComments=Request.Form("cnt_comments")

	objTempRs=UpdateRecordSPWithConn(objConnMpis, "su.CUSTOML_ALL_ContactKgCommentsUpdate", Array( _
		Array(, adInteger, , iContactID), _
		Array(, adVarWChar, 4000, sContactComments)))
%>
	<!--#include file="../../_template/page.close.asp"-->
<%
End If                                                       
%>

<%
' Getting personal data from DB
Set objTempRs=GetDataRecordsetSPWithConn(objConnMpis, "su.CUSTOML_ALL_ContactDetailsSelect", Array( _
	Array(, adInteger, , iContactID), _
	Array(, adInteger, , Null)))
If Not objTempRs.Eof Then 

	iContactID=objTempRs("IDCONTACT")
	sFirstName=objTempRs("FIRSTNAME")
	sLastName=objTempRs("NAME")
	sContactComments=objTempRs("CONTACT_KG_COMMENTS")
	
End If 
objTempRs.Close		
%>  

<!--#include virtual="/_template/html.header.asp"-->
<body>
<div id="holder">
	<!-- header -->
	<!--#include virtual="/_template/page.header.asp"-->

	<!-- content -->
	<div id="content" class="searchform">

		<h2 class="service_title">Curriculum Vitae. <span class="service_slogan">Comments</span></h2>
		<% ShowMessageStart "info", 440 %>
		Please always specify your name and date before any comment.
		<br><br>
		<% ShowMessageEnd %>

	<form method="post" action="<%=sScriptFullName%>">
	<input type="hidden" name="id_contact" value="<%=iContactID%>">
		<div class="box search blue">
		<h3><span class="left">&nbsp;</span><span class="right">&nbsp;</span>Expert details</h3>
		<table class="search_form" width="100%" cellspacing=0 cellpadding=0 border=0>
		<tr>
		<td class="field splitter"><label>Full&nbsp;name</label></td>
		<td class="value blue"><p><% =sLastName & ", " & sFirstName %></p></td>
		</tr>
		</table>
		<table class="search_form" width="100%" cellspacing=0 cellpadding=0 border=0>
		<tr class="last">
		<td class="field splitter"><label for="exp_comments">Comments</label></td>
		<td class="value blue"><textarea cols="34" style="width:355px;" name="cnt_comments" rows=12 wrap="yes"><%=sContactComments%></textarea></td>
		</tr>
		</table>
		</div>

		<div class="spacebottom">
		<input type="image" class="button first" src="/image/bte_savecont.gif" name="btnSubmit" id="btnSubmit" alt="Save & continue" border=0>
		</div>
		</form>

	</div>
</div>
	<!-- footer -->
	<!--#include virtual="/_template/page.footer.asp"-->
<% CloseDBConnection %>
</body>
<!--#include virtual="/_template/html.footer.asp"-->
