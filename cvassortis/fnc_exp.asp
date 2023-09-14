<% 
Const aClientSecurityCvViewDisabled = 0
Const aClientSecurityCvViewEnabled = 1
Const aClientSecurityCvViewAll = 5
Const aClientSecurityCvViewDenied = 9

Dim iTotalExperts
Dim bShowIbfCVs, bShowAssortisCVs, bBrowseMode, sIpFilter
Dim dLastUpdate

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

Sub LogUserCvView(AUserID, ACvID, ACvValid)
Dim objTempRs5
	objTempRs5=UpdateRecordSP("usp_UserExpertViewUpdate", Array( _
		Array(, adInteger, , AUserID), _
		Array(, adInteger, , ACvID), _
		Array(, adInteger, , ACvValid)))
	Set objTempRs5=Nothing
End Sub


Sub ShowIcaExpertsBlock(bSplitedOnPages, iCurrentPage)
	Dim iCurrentRow, iCurrentRecord, bIsTopExpert

	bIsTopExpert=False
	iCurrentRow=0
	If bSplitedOnPages=1 Then
		objTempRs.AbsolutePage=CInt(iCurrentPage)
	Else 
		objTempRs.PageSize=objTempRs.RecordCount
	End If
	
	While Not objTempRs.Eof And iCurrentRow<objTempRs.PageSize 
		Set objExpertDB = objExpertDBList.Find(objTempRs("DB"), "Database")

		Dim bCvValidForMemberOrExpert
		bCvValidForMemberOrExpert = 1
		
		objTempRs2=GetDataOutParamsSP("usp_Ica_GetExpertProfDetails", Array( _
			Array(, adVarChar, 25, objExpertDB.Database), _
			Array(, adInteger, , objTempRs("id_Expert")), _
			Array(, adInteger, , iMemberID), _
			Array(, adVarChar, 3, objTempRs("Lng")), _
			Array(, adInteger, , bShowIbfCVs)), Array( _
			Array(, adTinyInt), _
			Array(, adVarWChar, 400), _
			Array(, adVarWChar, 50), _
			Array(, adVarWChar, 255), _
			Array(, adVarWChar, 1000), _
			Array(, adVarWChar, 500), _
			Array(, adVarWChar, 1000), _
			Array(, adVarWChar, 500), _
			Array(, adVarWChar, 1000), _
			Array(, adVarWChar, 2000)))

		dLastUpdate=objTempRs2(2)

		' Check if expert is a top expert
		bIsTopExpert = GetExpertTopExpertByUid(objTempRs("uid_Expert"))
		
		Dim sBlockColor, sBlockTopExpert
		sBlockTopExpert=""
		sBlockColor="blue"
		
		If bIsTopExpert Then
			sBlockColor="topexpert"
			sBlockTopExpert="<img src=""/image/topexpert.gif"" align=""right"" width=81 height=19 hspace=5>"
		ElseIf objExpertDB.Database=objUserCompanyDB.Database Then
			sBlockColor="purple"
		End If
		Dim bExpertCvViewedByUser, iPeriodExpertCvViewedByUser
		iPeriodExpertCvViewedByUser = 7
		bExpertCvViewedByUser = GetUserExpertCvViewed(iUserID, objTempRs("id_Expert"), iPeriodExpertCvViewedByUser)
		If bExpertCvViewedByUser = aClientSecurityCvViewDisabled Or bExpertCvViewedByUser = aClientSecurityCvViewEnabled Or bExpertCvViewedByUser = aClientSecurityCvViewAll Then sBlockColor = sBlockColor + " viewed"
		
		ShowIcaExpertsBlockHeader "100%", "52", "ex", "Expert details", sBlockColor, ((iCurrentPage-1)*10+iCurrentRow+1)
		ShowUserNoticesViewHeader "99%", 180

		ShowExpertMpisFlag objExpertDB, objTempRs("id_Expert"), bCvValidForMemberOrExpert

		ShowUserNoticesViewTextWithStyles "</b>Nationality", sBlockTopExpert & "<b>" & objTempRs2(3) & "</b>" & ShowExpertMpisProfile(objExpertDB, objTempRs("id_Expert"), bCvValidForMemberOrExpert), "class=""field splitter""", "class=""value"""
		ShowUserNoticesViewText "</b>Profession", "<b>" & objTempRs("Profession") &  "</b>"
		ShowUserNoticesViewText "</b>Languages", objTempRs2(5)
		ShowUserNoticesViewText "</b>Regions experience", objTempRs2(6) 
		ShowUserNoticesViewText "</b>Sectors experience", objTempRs2(8)
		ShowUserNoticesViewText "</b>Seniority", "<b>" & objTempRs("Seniority") & " years</b>"

		ShowExpertAvailability objExpertDB, objTempRs("id_Expert"), bCvValidForMemberOrExpert

		ShowUserNoticesViewFooter
		ShowExpertsBlockFooter "100%", "52", "ex"
		ShowBreakLine
		Set objTempRs2 = Nothing

	iCurrentRow=iCurrentRow+1
	objTempRs.MoveNext
	WEnd

	Response.Flush()
	iTotalExperts=iCurrentRow
End Sub


Sub ShowIcaExpertsBlockHeader(iWidth, iHeight, sBlockType, sBlockDescription, sBlockColor, iCurrentRow)
	If sBlockColor<>"blue" And sBlockColor<>"purple" And sBlockColor<>"red" And sBlockColor<>"green" And sBlockColor<>"topexpert" Then
		sBlockColor="blue"
	End If
	%>
	<div class="box results <% =sBlockColor %>">
	<table class="expert header" cellpadding="0" cellspacing="0">
	<tr>
	<td width="5"><h3><span class="left">&nbsp;</span></h3></td>
			<td width="55%">
		<%
		Dim bCvValidForMemberOrExpert
		bCvValidForMemberOrExpert=IsIcaUserCompanyCvValid(objExpertDB.Database, objTempRs("id_Expert"), objUserCompanyDB.Database)
		
		If bCvValidForMemberOrExpert=1 Or bCvValidForMemberOrExpert=5 Then
		%>
			<h3><% If Not IsNull(objTempRs2(1)) Then %><%=objTempRs2(1) %><% End If %></h3></td>
		<% Else %>
			<h3>ID: <% =objExpertDB.DatabaseCode %><% =objTempRs("id_Expert") %></h3></td>
		<% End If %>
		<%
		' Get alternative CV owners
		Dim objExpertDBOtherList
		Set objExpertDBOtherList = New CCompanyExpertDBList
		objExpertDBOtherList.LoadData "usp_Ica_ExpertDBOwnerOtherSelect", Array( _
				Array(, adVarChar, 50, objExpertDB.Database),_
				Array(, adInteger, ,objTempRs("id_Expert")))
		Dim sExpertDBOtherList
		sExpertDBOtherList=objExpertDBOtherList.List("Database", ", ")
		If Len(sExpertDBOtherList)>1 Then 
			sExpertDBOtherList = "&nbsp;+&nbsp;<div class=""tooltip"">" & UCase(objExpertDB.Database) & ", " & UCase(sExpertDBOtherList) & "&nbsp;&nbsp;</div>"
		Else
			sExpertDBOtherList = ""
		End If
		%>		
			<td width="10%"><h3><small><div class="hover"><% =UCase(objExpertDB.Database) & sExpertDBOtherList %>&nbsp;&nbsp;</div></small></h3></td>
			<td width="10%"><h3><select class="cvformat" name="cvformat<%=iCurrentRow%>" size=1>
			<option value="" >- CV format -</option>
			<option value="">assortis.com</option>
			<option value="ADB">Asian Development Bank</option>
			<option value="AFB">African Development Bank</option>
			<option value="EC">European Commission</option>
			<option value="EP">Europass</option>
			<option value="WB">World Bank</option>
			</select></h3></td>
			<td width="25%"><h3><a href="javascript:DownloadCV(<%=iCurrentRow &", '" & objTempRs("uid_Expert") & "', '" & Left(LCase(objTempRs("Lng")),2) & "'" %>);"><img src="/image/bte_viewcv.gif" style="margin: 8px 6px 8px 8px;" border="0" width="72" height="18" alt="View CV"></a></h3></td>
		<td width="60"><h3><% If dLastUpdate>"" And IsDate(dLastUpdate) Then %><small>Update:&nbsp;<% =ConvertDateForText(dLastUpdate, "&nbsp;", "MMYYYY") %>&nbsp;&nbsp;</small><% End If %></h3></td>
	<td width="5"><h3><span class="right">&nbsp;</span></h3></td>
	</tr>
	</table>
<%
End Sub

Function GetExpNationalities(ALanguage, ADatabase, iExpertID)
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

	Set objTempRs2=GetDataRecordsetSPWithConn(objConnCustom, "usp_ExpCvvNationalitySelect", Array( _
		Array(, adInteger, , iExpertID)))

	sNationalityTemp = CreateListString(objTempRs2, sInfoFieldName, "", ", ")
	objTempRs2.Close	
	Set objTempRs2=Nothing
	
	GetExpNationalities=sNationalityTemp
End Function



Function GetExpertExperienceDonorRs(AExpertID, AExpertExperienceID, AOrderBy)
	Set GetExpertExperienceDonorRs=GetDataRecordsetSPWithConn(objConnCustom, "usp_ExpertExperienceDonorSelect", Array( _
		Array(, adInteger, , AExpertID), _
		Array(, adInteger, , AExpertExperienceID), _
		Array(, adVarChar, 80, AOrderBy)))
End Function

Function GetExpertExperienceCountryRs(AExpertID, AExpertExperienceID, AOrderBy)
	Set GetExpertExperienceCountryRs=GetDataRecordsetSPWithConn(objConnCustom, "usp_ExpertExperienceCountrySelect", Array( _
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
	GetExpertExperienceCountryGroupedList = CreateGroupedListString(GetExpertExperienceCountryRs(AExpertID, AExpertExperienceID, sInfoFieldName), sGroupFieldName, sInfoFieldName, "<p><b>", "</p>", ":</b> ", "", ", ")
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
	Set GetExpertExperienceSectorRs=GetDataRecordsetSPWithConn(objConnCustom, "usp_ExpertExperienceSectorSelect", Array( _
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
	GetExpertExperienceSectorGroupedList = CreateGroupedListString(GetExpertExperienceSectorRs(AExpertID, AExpertExperienceID, sInfoFieldName), sGroupFieldName, sInfoFieldName, "<p><b>", "</p>", ":</b><br />", "- ", "<br />")
End Function

Function GetExpCount(ADatabase)
Dim iResult
iResult = 0
	Dim objTempRs2
	Set objTempRs2=GetDataRecordsetSP("usp_Ica_ExpertsCountSelect", Array( _
		Array(, adInteger, , Null), _
		Array(, adVarChar, 25, ADatabase)))

	If Not objTempRs2.Eof Then
		iResult=objTempRs2(0)
		objTempRs2.MoveNext
	End If
	objTempRs2.Close	
	Set objTempRs2=Nothing
	
GetExpCount=iResult
End Function

Function GetUserExpertCvViewed(AUserID, AExpertID, APeriodExpertCvViewedByMember)
Dim objTempRs5
Dim bResult
	bResult=-1

	Set objTempRs5=GetDataRecordsetSP("usp_UserExpertViewSelect", Array( _
		Array(, adInteger, , AUserID), _
		Array(, adInteger, , AExpertID), _
		Array(, adInteger, , APeriodExpertCvViewedByMember)))
	If Not objTempRs5.Eof Then
		bResult=CheckIntegerAndZero(objTempRs5("Active"))
	End If
	Set objTempRs5=Nothing
	
	GetUserExpertCvViewed=bResult
End Function

Function GetExpertTopExpertByUid(AExpertUid)
Dim bResult
bResult = False
	Dim objTempRs2
	Set objTempRs2=GetDataRecordsetSP("usp_Ica_ExpertTopExpertByUidSelect", Array( _
		Array(, adVarChar, 40, AExpertUid)))

	If Not objTempRs2.Eof Then
		If CheckIntegerAndZero(objTempRs2(0))>0 Then
			bResult=True
		End If
	End If
	objTempRs2.Close	
	Set objTempRs2=Nothing
	
GetExpertTopExpertByUid=bResult
End Function


Sub ShowExpertMpisFlag(AExpertDB, AExpertID, ACvValidForMemberOrExpert)
Dim objTempRs3
	If bUserMpisAccess=1 Then
	On Error Resume Next
		' Extract status from MPIS
		Set objTempRs3=GetDataRecordsetSP("usp_Ica_ExpertMpisFlagByIDSelect", Array( _
			Array(, adInteger, , AExpertDB.ID), _
			Array(, adInteger, , AExpertID)))
		If Not objTempRs3.Eof Then
			If Len(objTempRs3("CONTACT_RELATION_STATUS_FLAG"))>2 Then
				%>
				<tr>
				<td class="field splitter"><p>&nbsp;</p></td>
				<td class="value flag_<% =objTempRs3("CONTACT_RELATION_STATUS_FLAG") %>_bg"><p>
				<img src="/image/flag_<% =objTempRs3("CONTACT_RELATION_STATUS_FLAG") %>.gif" alt="<% =objTempRs3("CONTACT_RELATION_STATUS_VALUE") %>" vspace="3" width="7" height="12" border="0" align="left">
				&nbsp;<b><% =objTempRs3("CONTACT_RELATION_STATUS_VALUE") %></b>
				<% If Len(objTempRs3("CONTACT_RELATION_COMMENTS"))>2 Then %>
					(<% =objTempRs3("CONTACT_RELATION_COMMENTS") %>)
				<% End If %>
				</p></td>
				</tr>
				<%
			End If
		End If
		objTempRs3.Close
		Set objTempRs3 = Nothing
	On Error GoTo 0	
	End If
End Sub


Function ShowExpertMpisProfile(AExpertDB, AExpertID, ACvValidForMemberOrExpert)
Dim sResult
Dim objTempRs3
Dim iMpisContactID
	iMpisContactID = 0

	If bUserMpisAccess=1 Then
	On Error Resume Next
		Set objTempRs3=GetDataRecordsetSP("usp_Ica_ExpertMpisContactByIDSelect	", Array( _
			Array(, adInteger, , AExpertDB.ID), _
			Array(, adInteger, , AExpertID)))
		If Not objTempRs3.Eof Then
			iMpisContactID = objTempRs3("IDCONTACT")
		End If
		objTempRs3.Close
		Set objTempRs3 = Nothing
			
		If iMpisContactID>0 Then
			' Extract status from MPIS
			Set objTempRs3=GetDataRecordsetSP("usp_ExpertMpisProjectsByIDSelect", Array( _
				Array(, adInteger, , AExpertDB.ID), _
				Array(, adInteger, , AExpertID)))

			If Not objTempRs3.Eof Then
				sResult = "<a href='http://mpis.ibf.be/elinkapps/elink.dll/consult?PAGE=ContConsultFrameSet.htm&TARGET=CONT&DETAIL=DEFAULT&KEY=" & iMpisContactID & "' target=_blank><img src='/image/mpis_profile.gif' hspace='10' vspace='2' align='right'></a>"
			End If
			objTempRs3.Close
			Set objTempRs3 = Nothing
		End If
	On Error GoTo 0
	End If
ShowExpertMpisProfile = sResult
End Function


Sub ShowExpertAvailability(AExpertDB, AExpertID, ACvValidForMemberOrExpert)
Dim objTempRs3

	If ACvValidForMemberOrExpert=1 Or ACvValidForMemberOrExpert=5 Then
		Set objTempRs3=GetDataRecordsetSP("usp_Ica_ExpertSelect", Array( _
			Array(, adVarChar, 25, AExpertDB.Id), _
			Array(, adInteger, , AExpertID)))
		If Not objTempRs3.Eof Then
			If Len(objTempRs3("expAvailabilityEng"))>2 Then
				%>
				<tr>
				<td class="field splitter"><p>Availability</p></td>
				<td class="value"><p><b>
				(<% =objTempRs3("expAvailabilityEng") %>)
				</b></p></td>
				</tr>
				<%
			End If
		End If
		objTempRs3.Close
		Set objTempRs3 = Nothing
	End If
End Sub
%>
