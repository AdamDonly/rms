<%
Dim sOrderBy
sOrderBy=UCase(objForm("ord"))
If sOrderBy<>"E" And sOrderBy<>"R" And sOrderBy<>"I" And sOrderBy<>"U" Then sOrderBy="A"

Dim iCurrentRow
%>
<!--#include virtual="/_template/html.header.asp"-->
<body>
	<!-- header -->
	<!--#include virtual="/_template/page.header.asp"-->

	<!-- content -->
	<div id="content" class="fullscreen">

		<div class="colCCCCCC uprCse f17 spc01 botMrgn10"><span class="service_title">CV registration.</span>The expert CV might be already registered</div>

	<% ShowMessageStart "info", 580 %>
	<b>There are some CVs similar to the one you are going to encode.</b><br>
	If the CV you are going to register is in the list click on expert's name to update that CV.
	<% ShowMessageEnd %><br/>

<%
Dim sPushParams, sVerifyParams
sPushParams = ""
sPushParams = ReplaceUrlParams(sPushParams, "act=push")
sPushParams = ReplaceUrlParams(sPushParams, "url=" & sUrl)

sVerifyParams = ReplaceUrlParams(sVerifyParams, "act=verify")

For Each objField In objForm
	If objField<>"cvEng" Then 
		sPushParams = sPushParams & "&" & objField & "=" & objForm(objField)
		sVerifyParams = sVerifyParams & "&" & objField & "=" & objForm(objField)
	End If
Next
%>
	
	<div class="box results blue" style="max-width:100%">
	<h3>Similar CVs</h3>
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
		<td><p class="mt"><a class="mt" href="<% =sApplicationHomePath & "view/cv_verify.asp" & ReplaceUrlParams(sVerifyParams, "uid=" & objResult("uid_Expert")) %>" target=_blank><% =objResult("psnLastName") & ", " & objResult("psnFirstName") & " " & objResult("psnMiddleName") & " (" & objResult("ptlName") & ")" %></a></td>
		<td class="date"><% =objResult("psnBirthdate") %></td>
		<td class="language">
			<% 
				Response.Write "<p class=""mt"">" & "<a class=""list"" href=""" & sApplicationHomePath & "view/cv_verify.asp" & ReplaceUrlParams(sVerifyParams, "uid=" & objResult("uid_Expert")) & """ target=_blank>" & ReplaceIfEmpty(dictLanguage.Item(Trim(objResult("Lng"))), objResult("Lng")) & "</a></p>"
			
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
					'Response.Write "<p class=""mt"">" & "<a class=""list"" href=""" & sApplicationHomePath & "view/cv_verify.asp?uid=" & objOtherRs("uid_Expert2") & """ target=_blank>" & ReplaceIfEmpty(dictLanguage.Item(Trim(objOtherRs("Lng2"))), objOtherRs("Lng2")) & "</a></p>"
					objOtherRs.MoveNext
				WEnd
				Set objOtherRs = Nothing
			%>

			<% If Not IsNull(objResult("id_Expert_Fra")) Then %>
				<p class="mt">
					<a class="list" href="<%= sApplicationHomePath %>/view/cv_verify.asp?uid=<%= objResult("uid_expert_fra") %>" target="_blank">
						Fran�ais
					</a>
				</p>
			<% End If %>
			<% If Not IsNull(objResult("id_Expert_Spa")) Then %>
				<p class="mt">
					<a class="list" href="<%= sApplicationHomePath %>/view/cv_verify.asp?uid=<%= objResult("uid_expert_spa") %>" target="_blank">
						Spanish
					</a>
				</p>
			<% End If %>
		</td>
		<td>
			<p class="mt">
				<% If bCvLanguageAlreadyRegistered = 0 Then %>
					<a class="list" href="<% =sScriptFileName & ReplaceUrlParams(sPushParams, "lnglink=" & objResult("id_Expert")) %>"><u>Register&nbsp;expert's&nbsp;CV<br />in&nbsp;the&nbsp;new&nbsp;language<% On Error Resume Next %>&nbsp;(<% =dictLanguage.Item(objForm("exp_language")) %>)<% On Error GoTo 0 %></u></a>
				<% Else %>
					CV in the selected language <% On Error Resume Next %>(<% =dictLanguage.Item(objForm("exp_language")) %>)<% On Error GoTo 0 %> is already registered. Click&nbsp;on&nbsp;language&nbsp;to&nbsp;verify&nbsp;or&nbsp;update.
				<% End If %>
			</p>
		</td>

		</tr>
		<%
		iCurrentRow=iCurrentRow+1
		objResult.MoveNext
	WEnd %>
	</table>
	</div><br/>
	
<p align=center><a class="list" href="<%=sScriptFileName & sPushParams %>"><b><u>If the expert is not in the list - continue encoding new CV</u></b></a>
<br>This option will create a new expert in the database - <br>please use it only when you are sure that the expert is not presented in the list above!</p>

	</div>
	<!-- footer -->
	<!--#include file="../_template/page.footer.asp"-->
<% CloseDBConnection %>
</body>
<!--#include file="../_template/html.footer.asp"-->
