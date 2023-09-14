<%
Dim sOrderBy
sOrderBy=UCase(objForm("ord"))
If sOrderBy<>"E" And sOrderBy<>"R" And sOrderBy<>"I" And sOrderBy<>"U" Then sOrderBy="A"

Dim iCurrentRow
%>
<!--#include virtual="/_template/html.header.asp"-->
<body>
<div id="holder">
	<!-- header -->
	<!--#include virtual="/_template/page.header.asp"-->

	<!-- content -->
	<div id="content" class="fullscreen">

		<h2 class="service_title">CV registration. <span class="service_slogan">The expert CV might be already registered</span>
		</h2><br/>

	<% ShowMessageStart "info", 580 %>
	<b>There are some CVs similar to the one you are going to encode.</b><br>
	If the CV you are going to register is in the list click on expert's name to update that CV.
	<% ShowMessageEnd %><br/>

<%
sTempParams = ""
sTempParams = ReplaceUrlParams(sTempParams, "act=push")
sTempParams = ReplaceUrlParams(sTempParams, "url=" & sUrl)

For Each objField In objForm
	If objField<>"cvEng" Then sTempParams = sTempParams & "&" & objField & "=" & objForm(objField)
Next
%>
	
	<div class="box results blue">
	<h3><span class="left">&nbsp;</span><span class="right">&nbsp;</span>Similar CVs</h3>
	<table class="results">
	<tr>
	<th class="number">&nbsp;&nbsp;Expert&nbsp;ID&nbsp;&nbsp;</th>
	<th class="title">Family name, first names (title)</th>
	<th class="date">Date of birth</th>
	<th class="language">CV language(s)</th>
	<th class="title">New CV language</th>
	</tr>

	<%
	Dim sCVTypeColor, sCVTypeText
	iCurrentRow=0
	While Not objResult.Eof

		Set objExpertDB = objExpertDBList.Find(objResult("DB"), "Database")
		%>
		<tr class="tr_results<% If iCurrentRow Mod 2 <> 0 Then %> odd<% End If %>">
		<td class="number"><% =objExpertDB.DatabaseCode & objResult("id_Expert") %></td>
		<td><p class="mt"><a class="mt" href="<% =sApplicationHomePath & "view/cv_verify.asp?uid=" & objResult("uid_Expert") %>" target=_blank><% =objResult("psnLastName") & ", " & objResult("psnFirstName") & " " & objResult("psnMiddleName") & " (" & objResult("ptlName") & ")" %></a></td>
		<td class="date"><% =objResult("psnBirthdate") %></td>
		<td class="language">
			<% 
			Response.Write "<p class=""mt"">" & "<a class=""list"" href=""" & sApplicationHomePath & "view/cv_verify.asp?uid=" & objResult("uid_Expert") & """ target=_blank>" & ReplaceIfEmpty(dictLanguage.Item(Trim(objResult("Lng"))), objResult("Lng")) & "</a></p>"
			Dim bCvLanguageAlreadyRegistered
			bCvLanguageAlreadyRegistered = 0
			If (objResult("Lng") = Trim(objForm("exp_language"))) Then
				bCvLanguageAlreadyRegistered = 1
			End If
			
			Dim objOtherRs
			Set objOtherRs=GetDataRecordsetSP("usp_Ica_ExpertIdDetailsSelect", Array( _
				Array(, adVarChar, 25, objResult("DB")), _
				Array(, adInteger, , objResult("id_Expert"))))

			While Not objOtherRs.EOF
				If (objOtherRs("Lng") = Trim(objForm("exp_language")) Or objOtherRs("Lng2") = Trim(objForm("exp_language"))) Then
					bCvLanguageAlreadyRegistered = 1
				End If
				Response.Write "<p class=""mt"">" & "<a class=""list"" href=""" & sApplicationHomePath & "view/cv_verify.asp?uid=" & objOtherRs("uid_Expert2") & """ target=_blank>" & ReplaceIfEmpty(dictLanguage.Item(Trim(objOtherRs("Lng2"))), objOtherRs("Lng2")) & "</a></p>"
				objOtherRs.MoveNext
			WEnd
			Set objOtherRs = Nothing
			%>
		</td>
		<td><p class="mt">
		<% If bCvLanguageAlreadyRegistered = 0 Then %>
		<a class="list" href="<% =sScriptFileName & ReplaceUrlParams(sTempParams, "lnglink=" & objResult("id_Expert")) %>"><u>Register&nbsp;expert's&nbsp;CV<br />in&nbsp;the&nbsp;new&nbsp;language<% On Error Resume Next %>&nbsp;(<% =dictLanguage.Item(objForm("exp_language")) %>)<% On Error GoTo 0 %></u></a>
		<% Else %>
			CV in the selected language <% On Error Resume Next %>(<% =dictLanguage.Item(objForm("exp_language")) %>)<% On Error GoTo 0 %> is already registered. Click&nbsp;on&nbsp;language&nbsp;to&nbsp;verify&nbsp;or&nbsp;update.
		<% End If %>
		</p></td>

		</tr>
		<%
		iCurrentRow=iCurrentRow+1
		objResult.MoveNext
	WEnd %>
	</table>
	</div><br/>
	
<p align=center><a class="list" href="<%=sScriptFileName & sTempParams %>"><b><u>If the expert is not in the list - continue encoding new CV</u></b></a>
<br>This option will create a new expert in the database - <br>please use it only when you are sure that the expert is not presented in the list above!</p>

	</div>
</div>
	<!-- footer -->
	<!--#include file="../_template/page.footer.asp"-->
<% CloseDBConnection %>
</body>
<!--#include file="../_template/html.footer.asp"-->
