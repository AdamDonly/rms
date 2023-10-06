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

		<h2 class="service_title">CV check. <span class="service_slogan">The expert CV might be already registered</span>
		</h2><br/>

	<% ShowMessageStart "info", 580 %>
	<b>There are some CVs similar to the one you are going to encode.</b><br>
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
	<th class="database">Database(s)</th>
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
		<td class="database">
		<%
		' Get alternative CV owners
		Dim objExpertDBOtherList
		Set objExpertDBOtherList = New CCompanyExpertDBList
		objExpertDBOtherList.LoadData "usp_Ica_ExpertDBOwnerOtherSelect", Array( _
				Array(, adVarChar, 50, objExpertDB.Database),_
				Array(, adInteger, ,objResult("id_Expert")))
		Dim sExpertDBOtherList
		sExpertDBOtherList=objExpertDBOtherList.List("Database", ", ")
		Response.Write objExpertDB.Database
		If Len(sExpertDBOtherList)>0 Then 
			Response.Write ", " & sExpertDBOtherList
		End If
		%>
		</td>
		</tr>
		<%
		iCurrentRow=iCurrentRow+1
		objResult.MoveNext
	WEnd %>
	</table>
	</div><br/>
</div>
	<!-- footer -->
	<!--#include file="../_template/page.footer.asp"-->
<% CloseDBConnection %>
</body>
<!--#include file="../_template/html.footer.asp"-->
