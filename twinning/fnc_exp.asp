<% 

Dim iTotalExperts, iExpertCVSelectedByMember
Dim bShowIbfCVs, bShowAssortisCVs, bBrowseMode, sIpFilter

iTotalExperts=0
bShowIbfCVs=0
bShowAssortisCVs=1
bBrowseMode=0
sIpFilter=""

' Check user's security settings
iUserID=ReplaceIfEmpty(iUserID, 0)
Set objTempRs2=GetDataRecordsetSP("usp_UsrSecuritySelect", Array( _
	Array(, adInteger, , iUserID)))


If Not objTempRs2.Eof Then
	sIpFilter=objTempRs2("usrIpSecurity")
	If (sIpFilter>"" And InStr(sUserIpAddress, sIpFilter)>0) Or (sIpFilter="") Then
		bShowIbfCVs=objTempRs2("usrIbf")
		bBrowseMode=objTempRs2("usrBrowseMode")
	End If
End If

Sub CheckExpertRightsOnCVFormat(iExpertID)
Dim iStatusTemp
	iStatusTemp=0
	Set objTempRs2=GetDataRecordsetSP("usp_ExpAccountStatusSelect", Array( _
		Array(, adInteger, , iExpertID), _
		Array(, adInteger, , 31)))

	If Not objTempRs2.Eof Then
		iStatusTemp=objTempRs2("id_AccountStatus")
		If DateDiff("d", Now(), objTempRs2("eacExpDate"))<0 Then iStatusTemp=10
	End If

	If iStatusTemp<>1 Then
		If InStr(sScriptFileName, "adb.asp")>0 Or _
			InStr(sScriptFileName, "afb.asp")>0 Or _
			InStr(sScriptFileName, "ec.asp")>0 Or _
			InStr(sScriptFileName, "wb.asp")>0 Then
			ShowStandardPageHeader
			Response.Write "<br>"
			ShowMessage "You should be subscribed for the Special Info Pack in order to format your CV into the major funding agencies templates.", "error", 600 %>
	
	<br>
	<table width=580 cellpadding=0 cellspacing=0 border=0 align="center">
	<tr><td>
	<p>The Special Info Pack offers a comprehensive service to <b>experts searching for a new mission</b> for one of the major funding agencies*.<br>
	<br>
	<p>During <b>1 year</b>, benefit from:</p>
	<ul>
	<li><p><b>Daily email alerts</b> with the latest information published by all major funding agencies*:
	<ul>	
	<li><p><b>shortlisted companies</b> (incl. contact details) for <b>projects</b> in the <b>tendering</b> phase (incl. project description).</p></li>
	<li><p><b>contracted companies</b> (incl. contact details) for <b>projects</b> that have been <b>awarded</b> (incl. project description).</p></li>
	</ul>
	<li><p>Access to the <b>database of companies in your sectors</b> of interest, with full contact details and company profile (incl. recent contracted and shortlisted projects).</p></li>
	<li><p>Your <b>CV formatted</b> into the major funding agencies templates.</p></li>
	</ul>
	<br>

	<p><b>1 year</b> subscription fee is 175 EUR (excl. VAT).</p>

	<% ' Special offer until 20 Apr. 2003
	If DateDiff("d","2003/04/20",Date())<=0 Then %><p align="center"><span class="rs"><b>Now all at 50 EUR (excl. VAT) until 20th of April 2003!</span></b></p><% End If %>

    	<div align="center"><a href="experts/sc_register.asp<%=sParams%>"><img src="image/bte_registersip.gif" name="Register for the Special Info Pack" alt="Register for the Special Info Pack" width=226 height=18 border=0 vspace=12></a></div>
	<br>

	<p class="sml">* European Commission ( EC ), World Bank ( WB ), Inter-American Development Bank ( IADB ), Asian Development Bank ( ADB ), African Development Bank ( AfDB ), the UK Department for International Development ( DFID ), European Investment Bank (EIB), European Bank for Reconstruction and Development (EBRD).</sup></p>
	</td></tr>
	</table>

			<% 
			ShowStandardPageFooter
			Response.End
		End If
	End If
End Sub

Function IsMemberExpertCvValid(iMemberID, iExpertID, iCvID)
Dim bCvValidTemp, iExpertOriginalID
	bCvValidTemp=0
	If IsNumeric(iCvID) Then
		iCvID=CLng(iCvID)
	End If

	If iMemberID=0 And iExpertID>0 And iExpertID=iCvID Then
		bCvValidTemp=1

		' 20040216 - Check subscription for SIP. Stop showing ADB, AFDB, EC and WB CVs for those who didn't subscribe
		CheckExpertRightsOnCVFormat(iCvID)

	ElseIf iMemberID=0 And iExpertID>0 And iExpertID<>iCvID Then
		objTempRs=GetDataOutParamsSP("usp_ExpCvvOriginalSelect", Array( _
			Array(, adInteger, , iExpertID)), _
			Array( Array(, adInteger)))
		iExpertOriginalID=objTempRs(0)

		If iExpertOriginalID=iCvID Then bCvValidTemp=1

		' 20040216 - Check subscription for SIP. Stop showing ADB, AFDB, EC and WB CVs for those who didn't subscribe
		CheckExpertRightsOnCVFormat(iCvID)

	ElseIf iMemberID>0 And iCvID>0 Then 
		objTempRs=GetDataOutParamsSP("usp_MmbExpDownloadedWithExtraSecurity", Array( _
			Array(, adInteger, , iMemberID), _
			Array(, adInteger, , iCvID), _
			Array(, adInteger, , bShowIbfCVs), _
			Array(, adInteger, , bBrowseMode)), _
			Array( Array(, adInteger)))
		bCvValidTemp=ReplaceIfEmpty(objTempRs(0), 0)

	End If
IsMemberExpertCvValid=bCvValidTemp
End Function


Sub ShowExpertsBlock(bSplitedOnPages, iCurrentPage)
Dim iCurrentRow, iCurrentRecord

	iCurrentRow=0
	If bSplitedOnPages=1 Then
		objTempRs.AbsolutePage=CInt(iCurrentPage)
	Else 
		objTempRs.PageSize=objTempRs.RecordCount
	End If
	While Not objTempRs.Eof And iCurrentRow<objTempRs.PageSize 

		objTempRs2=GetDataOutParamsSP("usp_GetExpertProfDetails", Array( _
			Array(, adInteger, , objTempRs("id_Expert")), Array(, adInteger, , iMemberID), Array(, adVarChar, 3, objTempRs("Lng")), Array(, adInteger, , bShowIbfCVs)), Array( _
			Array(, adTinyInt), Array(, adVarWChar, 400), Array(, adVarWChar, 50), Array(, adVarWChar, 255), Array(, adVarWChar, 1000), Array(, adVarWChar, 500), Array(, adVarWChar, 1000), Array(, adVarWChar, 500), Array(, adVarWChar, 1000), Array(, adVarWChar, 2000)))

		iExpertCVSelectedByMember=ReplaceIfEmpty(objTempRs2(0),0)
		If iExpertCVSelectedByMember=0 Then iExpertCVSelectedByMember=bBrowseMode
	iExpertCVSelectedByMember=1

		If iExpertCVSelectedByMember<>1 And iExpertCVSelectedByMember<>5 Then iExpertCVSelectedByMember=0
		ShowExpertsBlockHeader "100%", "52", "ex" & iExpertCVSelectedByMember, "Expert details", "ex" & iExpertCVSelectedByMember, ((iCurrentPage-1)*10+iCurrentRow+1)
		ShowUserNoticesViewHeader "99%", 180
		ShowUserNoticesViewSpacer 5
		ShowUserNoticesViewText "</b>Nationality", "<b>" & objTempRs2(3) & "</b>"
		'ShowUserNoticesViewText "</b>Education", objTempRs2(4)
		ShowUserNoticesViewText "</b>Profession", "<b>" & objTempRs("Profession") &  "</b>"
		ShowUserNoticesViewText "</b>Languages", objTempRs2(5)
		ShowUserNoticesViewText "</b>Regions experience", objTempRs2(6) 
		ShowUserNoticesViewText "</b>Sectors experience", objTempRs2(8)
		ShowUserNoticesViewText "</b>Seniority", "<b>" & objTempRs("Seniority") & " years</b>"

		ShowUserNoticesViewSpacer 5
		ShowUserNoticesViewFooter
		ShowExpertsBlockFooter "100%", "52", "ex"  & iExpertCVSelectedByMember
		ShowBreakLine
		Set objTempRs2 = Nothing

	iCurrentRow=iCurrentRow+1
	objTempRs.MoveNext
	WEnd

	Response.Flush()
	iTotalExperts=iCurrentRow

End Sub


Sub ShowExpertsBlockHeader(iWidth, iHeight, sBlockType, sBlockDescription, sBlockColor, iCurrentRow)
	SetColorByBlockType(sBlockColor)
	%>	
	<table width=<%=iWidth%> cellspacing=0 cellpadding=0 border=0 align="center">
	<tr height="29">
	<td width="22" bgcolor="<%=sTitleColor%>" background="<%=sHomePath%>image/pmmb_<%=sBlockType%>_t1e_bg.gif" valign="top"><img src="<%=sHomePath%>image/pmmb_<%=sBlockType%>_t1e1.gif" width=22 height=29 hspace=0 vspace=0></td>
  	<td width="99%" bgcolor="<%=sTitleColor%>" valign="center" background="<%=sHomePath%>image/pmmb_<%=sBlockColor%>top_bg11.gif" valign="center"><img src="<%=sHomePath%>image/x.gif" width=100 height=1 vspace=0><br>
		<table width="100%" cellpadding=0 cellspacing=0 border=0><tr>
		<td width="30" valign="top"><img src="../../image/x.gif" width=1 height=1 align="left"><input type="checkbox" class="expert_query_selection" id="expert_<%=objTempRs("id_Expert")%>_query_<% =iSearchQueryID %>_selection"></td>
		<td width="70%" valign="center"><p class="txt"><b><% If Not IsNull(objTempRs2(1)) Then %><%=Replace(objTempRs2(1), " ", "&nbsp;") %><% End If %></b></td>
		<td width="10%" valign="center"><select name="cvformat<%=iCurrentRow%>" size=1>
		<option value="" >- CV format -</option>
		<option value="">assortis.com</option>
		<option value="ADB">Asian Development Bank</option>
		<option value="AFB">African Development Bank</option>
		<option value="EC">European Commission</option>
		<option value="WB">World Bank</option>
		</select></td>
		<td width="20%" valign="center"><a href="javascript:DownloadCV(<%=iCurrentRow &"," & objTempRs("id_Expert") & ",'" & Left(LCase(objTempRs("Lng")),2) &"'" %>);"><img src="<% = sHomePath %>image/bte_download.gif" border=0 width=105 height=18 vspace=0 hspace=8 alt="Download CV"></a></td>
		</tr></table>
	</td>
	<td width="6" bgcolor="<%=sTitleColor%>" background="<%=sHomePath%>image/pmmb_<%=sBlockColor%>top_bg21.gif" valign="top"><img src="<%=sHomePath%>image/pmmb_<%=sBlockColor%>top_21.gif" width=6 height=29 hspace=0 vspace=0></td>
	</tr>
	<tr height="3">
	<td bgcolor="<%=sTitleColor%>" valign="top"><img src="<%=sHomePath%>image/pmmb_<%=sBlockType%>_t1e12.gif" width=22 height=3 hspace=0 vspace=0></td>
  	<td bgcolor="<%=sTitleColor%>" background="<%=sHomePath%>image/pmmb_<%=sBlockColor%>top_bg12.gif" valign="top"><img src="<%=sHomePath%>image/pmmb_<%=sBlockColor%>top_bg12.gif" width=100 height=3 vspace=0></td>
	<td bgcolor="<%=sTitleColor%>" valign="top"><img src="<%=sHomePath%>image/pmmb_<%=sBlockColor%>top_22.gif" width=6 height=3 hspace=0 vspace=0></td>
	</tr>
	</table>

	<table width=<%=iWidth%> cellspacing=0 cellpadding=0 border=0 align="center">
	<tr height="<%=iHeight-37%>">
	<td bgcolor="<%=sTextFrameColor%>" background="<%=sHomePath%>image/pmmb_<%=sBlockColor%>lft_bg1.gif">
<%
End Sub


Function GetExpNationalities(AExpertID, ALanguage)
	Dim sNationalityTemp
	sNationalityTemp=""

	Dim sInfoFieldName
	If ALanguage = cLanguageFrench Then
		sInfoFieldName = "couNameFra"
	ElseIf ALanguage = cLanguageSpanish Then
		sInfoFieldName = "couNameSpa"
	Else
		sInfoFieldName = "couNameEng"
	End If	

	Set objTempRs2=GetDataRecordsetSP("usp_ExpCvvNationalitySelect", Array( _
		Array(, adInteger, , AExpertID)))

	sNationalityTemp = CreateListString(objTempRs2, sInfoFieldName, "", ", ")
	objTempRs2.Close	
	Set objTempRs2=Nothing
	
	GetExpNationalities=sNationalityTemp
End Function

Function UpdateMemberExpertQuery(AMemberID, AExpertID, ASearchQueryID, AAction)
Dim objTempRs5
	objTempRs5=UpdateRecordSP("usp_MmbExpQueryUpdate", Array( _
		Array(, adInteger, , AMemberID), _
		Array(, adInteger, , AExpertID), _
		Array(, adInteger, , ASearchQueryID), _
		Array(, adInteger, , AAction)))
	Set objTempRs5=Nothing
End Function

Function UpdateMemberSearchQuery(AMemberID, ASearchQueryID)
	Dim objTempRs5
	objTempRs5=UpdateRecordSP("usp_MmbExpSearchQueryUpdate", Array( _
		Array(, adInteger, , AMemberID), _
		Array(, adInteger, , ASearchQueryID)))
	Set objTempRs5=Nothing
End Function

Function ShowMemberExpertQuery(AMemberID, AExpertID, ASearchQueryID)
	Dim sResult
	sResult=""
	Dim objTempRs5
	
	Set objTempRs5=GetDataRecordsetSP("usp_MmbExpQuerySelect", Array( _
		Array(, adInteger, , AMemberID), _
		Array(, adInteger, , AExpertID), _
		Array(, adInteger, , ASearchQueryID)))
	While Not objTempRs5.Eof
		sResult=sResult & objTempRs5("id_Expert") & ","
		objTempRs5.MoveNext
	WEnd
	Set objTempRs5=Nothing
	
	If Len(sResult)>1 Then sResult=Left(sResult, Len(sResult)-1)
	ShowMemberExpertQuery=sResult
End Function

Function GetExpertExperienceDonorRs(AExpertID, AExpertExperienceID, AOrderBy)
	Set GetExpertExperienceDonorRs=GetDataRecordsetSP("usp_ExpertExperienceDonorSelect", Array( _
		Array(, adInteger, , AExpertID), _
		Array(, adInteger, , AExpertExperienceID), _
		Array(, adVarChar, 80, AOrderBy)))
End Function

Function GetExpertExperienceCountryRs(AExpertID, AExpertExperienceID, AOrderBy)
	Set GetExpertExperienceCountryRs=GetDataRecordsetSP("usp_ExpertExperienceCountrySelect", Array( _
		Array(, adInteger, , AExpertID), _
		Array(, adInteger, , AExpertExperienceID), _
		Array(, adVarChar, 80, AOrderBy)))
End Function

Function GetExpertExperienceCountryGroupedList(AExpertID, AExpertExperienceID, ALanguage)
	Dim sGroupFieldName
	Dim sInfoFieldName
	If ALanguage = cLanguageFrench Then
		sGroupFieldName = "Geo_ZoneFra"
		sInfoFieldName = "couNameFra"
	ElseIf ALanguage = cLanguageSpanish Then
		sGroupFieldName = "Geo_ZoneSpa"
		sInfoFieldName = "couNameSpa"
	Else
		sGroupFieldName = "Geo_ZoneEng"
		sInfoFieldName = "couNameEng"
	End If
	GetExpertExperienceCountryGroupedList = CreateGroupedListString(GetExpertExperienceCountryRs(AExpertID, AExpertExperienceID, sInfoFieldName), sGroupFieldName, sInfoFieldName, "<p class=""txt""><b>", "</p>", ":</b> ", "", ", ")
End Function

Function GetExpertExperienceCountryList(AExpertID, AExpertExperienceID, ALanguage)
	Dim sGroupFieldName
	Dim sInfoFieldName
	If ALanguage = cLanguageFrench Then
		sGroupFieldName = "Geo_ZoneFra"
		sInfoFieldName = "couNameFra"
	ElseIf ALanguage = cLanguageSpanish Then
		sGroupFieldName = "Geo_ZoneSpa"
		sInfoFieldName = "couNameSpa"
	Else
		sGroupFieldName = "Geo_ZoneEng"
		sInfoFieldName = "couNameEng"
	End If
	Dim sOrderFieldName
	sOrderFieldName = sInfoFieldName & " only"
	GetExpertExperienceCountryList = CreateListString(GetExpertExperienceCountryRs(AExpertID, AExpertExperienceID, sOrderFieldName), sInfoFieldName, "", ", ")
End Function


Function GetExpertExperienceSectorRs(AExpertID, AExpertExperienceID, AOrderBy)
	Set GetExpertExperienceSectorRs=GetDataRecordsetSP("usp_ExpertExperienceSectorSelect", Array( _
		Array(, adInteger, , AExpertID), _
		Array(, adInteger, , AExpertExperienceID), _
		Array(, adVarChar, 80, AOrderBy)))
End Function

Function GetExpertExperienceSectorGroupedList(AExpertID, AExpertExperienceID, ALanguage)
	Dim sGroupFieldName
	Dim sInfoFieldName
	If ALanguage = cLanguageFrench Then
		sGroupFieldName = "mnsDescriptionFra"
		sInfoFieldName = "sctDescriptionFra"
	ElseIf ALanguage = cLanguageSpanish Then
		sGroupFieldName = "mnsDescriptionSpa"
		sInfoFieldName = "sctDescriptionSpa"
	Else
		sGroupFieldName = "mnsDescriptionEng"
		sInfoFieldName = "sctDescriptionEng"
	End If
	GetExpertExperienceSectorGroupedList = CreateGroupedListString(GetExpertExperienceSectorRs(AExpertID, AExpertExperienceID, sInfoFieldName), sGroupFieldName, sInfoFieldName, "<p class=""txt""><b>", "</p>", ":</b><br />", "- ", "<br />")
End Function

%>
