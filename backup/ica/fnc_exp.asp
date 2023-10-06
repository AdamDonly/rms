<% 
Const aClientSecurityCvViewDisabled = 0
Const aClientSecurityCvViewEnabled = 1
Const aClientSecurityCvViewAll = 5
Const aClientSecurityCvViewDenied = 9

Dim iTotalExperts
Dim bBrowseMode, sIpFilter
Dim dLastUpdate

iTotalExperts=0
bBrowseMode=0
sIpFilter=""

iUserID=ReplaceIfEmpty(iUserID, 0)

Dim objExpertTopExpertDBList
Dim objExpertDBOtherList


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
		objTempRs=GetDataOutParamsSPWithConn(objConnCustom, "usp_ExpCvvOriginalSelect", Array( _
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
			Array(, adInteger, , 0), _
			Array(, adInteger, , bBrowseMode)), _
			Array( Array(, adInteger)))
		bCvValidTemp=ReplaceIfEmpty(objTempRs(0), 0)

	End If
IsMemberExpertCvValid=bCvValidTemp
End Function

Sub LogUserCvView(AUserID, ADatabaseID, ACvID, ACvValid)
	Dim objTempRs5
	objTempRs5 = UpdateRecordSP("usp_UserExpertViewUpdate", Array( _
		Array(, adInteger, , AUserID), _
		Array(, adInteger, , ADatabaseID), _
		Array(, adInteger, , ACvID), _
		Array(, adInteger, , ACvValid)))
	Set objTempRs5 = Nothing
End Sub


Sub ShowIcaExpertsBlock(bSplitedOnPages, iCurrentPage)
	Dim iCurrentRow, iCurrentRecord

	iCurrentRow = 0
	If bSplitedOnPages = 1 Then
		objTempRs.AbsolutePage = CInt(iCurrentPage)
	Else 
		objTempRs.PageSize = objTempRs.RecordCount
	End If
	
	While Not objTempRs.Eof And iCurrentRow < objTempRs.PageSize 
		Set objExpertDB = objExpertDBList.Find(objTempRs("DB"), "Database")

		' Get DBs where expert is registered as Top Expert
		Set objExpertTopExpertDBList = New CCompanyExpertDBList
		objExpertTopExpertDBList.LoadData "usp_Ica_ExpertTopExpertDBSelect", Array( _
				Array(, adVarChar, 50, objExpertDB.Database),_
				Array(, adInteger, ,objTempRs("id_Expert")))

		' Get alternative CV owners
		Set objExpertDBOtherList = New CCompanyExpertDBList
		objExpertDBOtherList.LoadData "usp_Ica_ExpertDBOwnerOtherSelect", Array( _
				Array(, adVarChar, 50, objExpertDB.Database),_
				Array(, adInteger, ,objTempRs("id_Expert")))

		Dim bCvValidForMemberOrExpert
		bCvValidForMemberOrExpert = IsIcaUserCompanyCvValid(objExpertDB.Database, objTempRs("id_Expert"), objUserCompanyDB.Database)
		
		objTempRs2 = GetDataOutParamsSP("usp_Ica_GetExpertProfDetails", Array( _
			Array(, adVarChar, 25, objExpertDB.Database), _
			Array(, adInteger, , objTempRs("id_Expert")), _
			Array(, adInteger, , iMemberID), _
			Array(, adVarChar, 3, objTempRs("Lng")), _
			Array(, adInteger, , 0)), Array( _
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

		dLastUpdate = objTempRs2(2)

		Dim iTopExpertStatus
		Dim bIsCompanyCircleExpert
		bIsCompanyCircleExpert = False

		If bCvValidForMemberOrExpert = 1 Or bCvValidForMemberOrExpert = 5 Then
			' Verify if expert is already registered in my experts circle
			bIsCompanyCircleExpert = GetExpertCompanyCircleByUid(objTempRs("uid_Expert"), iUserCompanyID, iUserID)

			If bIsCompanyCircleExpert Then
				' Verify if expert is already registered in top experts
				iTopExpertStatus = GetExpertCompanyTopExpertByUid(objTempRs("uid_Expert"), iUserCompanyID, iUserID)
			End If
		End If


		
				Dim dictTopExpertOfOtherMember
				' check if is a top expert of another member:
				If Not bIsCompanyCircleExpert Then
					Set dictTopExpertOfOtherMember = GetExpertIsTopExpertByUid(objTempRs("uid_Expert"), iUserID)
				End If
		

		Dim sBlockColor, sBlockTopExtra
		sBlockTopExtra = ""
		sBlockColor = "blue"
		
		If bIsCompanyCircleExpert Then
			If iTopExpertStatus = 1 Then
				sBlockTopExtra = sBlockTopExtra & "<img src=""/image/file_top_full.gif"" align=""right"" width=86 height=18 hspace=5>"
			End If
			If iTopExpertStatus = 2 Then
				sBlockTopExtra = sBlockTopExtra & "<img src=""/image/file_top_pending.gif"" align=""right"" width=86 height=18 hspace=5>"
			End If

			sBlockTopExtra = sBlockTopExtra & "<img src=""/image/file_circle_full.gif"" align=""right"" width=125 height=18 hspace=5>"
		ElseIf Not IsNull(dictTopExpertOfOtherMember) Then
			Dim tmpMembersForTopExpert, mmbKey
			tmpMembersForTopExpert = ""
			mmbKey = ""
			
			For Each mmbKey In dictTopExpertOfOtherMember.Keys
				If Len(tmpMembersForTopExpert) > 0 Then 
					tmpMembersForTopExpert = tmpMembersForTopExpert & ", "
				End If
				' only for the approved top experts:
				If CDbl(dictTopExpertOfOtherMember.Item(mmbKey)) = 1 Then
					tmpMembersForTopExpert = tmpMembersForTopExpert & CStr(mmbKey)
				End If
			Next
			If Len(tmpMembersForTopExpert) > 0 Then
				sBlockTopExtra = sBlockTopExtra & "<div class=""topExpert_othMember""><img src=""/image/file_top_full.gif"" width=86 height=18/> of " & tmpMembersForTopExpert & "</div>"
			End If
		End If

		Dim bExpertCvViewedByUser, iPeriodExpertCvViewedByUser
		iPeriodExpertCvViewedByUser = 7
		bExpertCvViewedByUser = GetUserExpertCvViewed(iUserID, objExpertDB.Id, objTempRs("id_Expert"), iPeriodExpertCvViewedByUser)
		If bExpertCvViewedByUser = aClientSecurityCvViewDisabled Or bExpertCvViewedByUser = aClientSecurityCvViewEnabled Or bExpertCvViewedByUser = aClientSecurityCvViewAll Then sBlockColor = sBlockColor + " viewed"
		
		ShowIcaExpertsBlockHeader "100%", "52", "ex", "Expert details", sBlockColor, ((iCurrentPage-1)*10+iCurrentRow+1)
		ShowUserNoticesViewHeader "99%", 180

		If bCvValidForMemberOrExpert = 1 Or bCvValidForMemberOrExpert = 5 Then
			ShowExpertMethodologyFlag objTempRs("uid_Expert"), objExpertDB, objTempRs("id_Expert"), bCvValidForMemberOrExpert
			ShowExpertMpisFlag objExpertDB, objTempRs("id_Expert"), bCvValidForMemberOrExpert
		End If

		ShowUserNoticesViewTextWithStyles "</b>Nationality", sBlockTopExtra & "<b>" & objTempRs2(3) & "</b>" & ShowExpertMpisProfile(objExpertDB, objTempRs("id_Expert"), bCvValidForMemberOrExpert), "class=""field splitter""", "class=""value"""
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

Function IsRestrictedRowVisible(ARow)
	Dim bResult
	bResult = False
	If ARow Mod 7 = 0 And ARow < 200 Then
		If ARow <> 28 And _
			ARow <> 70 And _
			ARow <> 98 And _
			ARow <> 112 And _
			ARow <> 140 And _
			ARow <> 168 And _
			ARow <> 182 Then
			bResult = True
		End If
	End If
	IsRestrictedRowVisible = bResult
End Function


Sub ShowIcaExpertsBlockHeader(iWidth, iHeight, sBlockType, sBlockDescription, sBlockColor, iCurrentRow)
	Dim bCvValidForMemberOrExpert
	bCvValidForMemberOrExpert = IsIcaUserCompanyCvValid(objExpertDB.Database, objTempRs("id_Expert"), objUserCompanyDB.Database)
	Dim bCvViewRestricted
	If bCvValidForMemberOrExpert = 1 Or bCvValidForMemberOrExpert = 5 Then
		bCvViewRestricted = False
	ElseIf bAssortisSubscriptionEdbActive And objExpertDB.Database = "assortis" Then
		bCvViewRestricted = False
	ElseIf iMemberAccessExperts = cMemberAccessExpertsOwnOnly Then
		bCvViewRestricted = True
	ElseIf iMemberAccessExperts = cMemberAccessExpertsRestricted And Not IsRestrictedRowVisible(iCurrentRow) Then 
		bCvViewRestricted = True
	Else
		bCvViewRestricted = False
	End If

	If bCvViewRestricted Then
		sBlockColor = "grey"
	End If
	%>
	<div class="box results <% =sBlockColor %>">
	<table class="expert header" cellpadding="0" cellspacing="0">
	<tr>
		<%
		' For top experts the primary ownership is overwritten by top expert db
		If objExpertTopExpertDBList.Count > 0 Then
			Dim objExpertOriginalDB
			Set objExpertOriginalDB = objExpertDB

			If objExpertDB.ID <> objExpertTopExpertDBList.Item(0).ID Then
				Set objExpertDB = objExpertTopExpertDBList.Item(0)

				If objExpertOriginalDB.Database = "assortis" _
				And objExpertTopExpertDBList.Item(0).Database <> "ibf" _
				And bUserIbfStaff = 1 Then
					objExpertDBOtherList.ReplaceDB objExpertTopExpertDBList.Item(0), objExpertOriginalDB
				Else
					objExpertDBOtherList.ClearAll
				End If

			End If
		End If
		%>

		<%
		If bCvValidForMemberOrExpert = 1 Or bCvValidForMemberOrExpert = 5 Then
		%>
			<td width="57%"><h3><% If Not IsNull(objTempRs2(1)) Then %><%=objTempRs2(1) %><% End If %></h3></td>
		<% ElseIf bCvViewRestricted And objExpertDB.DatabaseCode = 101 And bAssortisSubscriptionEdbActive = True Then %>
			<td width="57%"><h3>ID: <% =objExpertDB.DatabaseCode %><% =objTempRs("id_Expert") %></h3></td>
		<% ElseIf bCvViewRestricted Then %>
			<td width="100%"><h3>ID: xxx-xxxxx</h3></td>
		<% Else %>
			<td width="57%"><h3>ID: <% =objExpertDB.DatabaseCode %><% =objTempRs("id_Expert") %></h3></td>
		<% End If %>
		<%

		Dim sExpertDBOtherList
		sExpertDBOtherList = objExpertDBOtherList.List("CompanyName", ", ")
		If Len(sExpertDBOtherList) > 1 Then 
			sExpertDBOtherList = "&nbsp;+&nbsp;<div class=""tooltip"">" & UCase(objExpertDB.Company.Name) & ", " & UCase(sExpertDBOtherList) & "&nbsp;&nbsp;</div>"
		Else
			sExpertDBOtherList = ""
		End If

		If Not bCvViewRestricted Then %>
			<td width="10%"><h3><small><div class="hover"><% If objExpertDB.Company Is Nothing Then %><%=objExpertDB.DatabaseTitle %><% Else %><%=UCase(objExpertDB.Company.Name) %><% End If %><%=sExpertDBOtherList %>&nbsp;&nbsp;</div></small></h3></td>
			<td width="10%"><h3><select class="cvformat" name="cvformat<%=iCurrentRow%>" size=1>
			<option value="" >- CV format - &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;</option>
			<option value="">assortis</option>
			<option value="ADB">ADB</option>
			<option value="AFB">AFDB</option>
			<option value="EC">EC</option>
			<option value="EP">Europass</option>
			<option value="WB">WB</option>
			</select></h3></td>
			<td width="25%"><h3><a href="javascript:DownloadCV(<%=iCurrentRow &", '" & objTempRs("uid_Expert") & "', '" & Left(LCase(objTempRs("Lng")),2) & "'" %>);" class="red-button white-shadow" style="margin:5px 0 0;vertical-align:top;">View CV</a></h3></td>
			<td width="60"><h3><% If dLastUpdate>"" And IsDate(dLastUpdate) Then %><small>Update:&nbsp;<% =ConvertDateForText(dLastUpdate, "&nbsp;", "MMYYYY") %>&nbsp;&nbsp;</small><% End If %></h3></td>
		<% End If %>
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
	Set objTempRs2 = GetDataRecordsetSP("usp_Ica_ExpertsCountSelect", Array( _
		Array(, adInteger, , Null), _
		Array(, adVarChar, 25, ADatabase)))

	If Not objTempRs2.Eof Then
		iResult = objTempRs2(0)
		objTempRs2.MoveNext
	End If
	objTempRs2.Close	
	Set objTempRs2 = Nothing
	
GetExpCount = iResult
End Function

Function GetUserExpertCvViewed(AUserID, ADatabaseID, AExpertID, APeriodExpertCvViewedByMember)
	Dim objTempRs5
	Dim bResult
	bResult = -1

	Set objTempRs5 = GetDataRecordsetSP("usp_UserExpertViewSelect", Array( _
		Array(, adInteger, , AUserID), _
		Array(, adInteger, , ADatabaseID), _
		Array(, adInteger, , AExpertID), _
		Array(, adInteger, , APeriodExpertCvViewedByMember)))
	If Not objTempRs5.Eof Then
		bResult = CheckIntegerAndZero(objTempRs5("Active"))
	End If
	Set objTempRs5 = Nothing
	
	GetUserExpertCvViewed = bResult
End Function

Function GetExpertTopExpertByUid(AExpertUid)
	Dim bResult
	bResult = False

	Dim objTempRs2
	Set objTempRs2 = GetDataRecordsetSP("usp_" & sIcaServerSqlPrefix & "ExpertTopExpertByUidSelect", Array( _
		Array(, adVarChar, 40, AExpertUid)))

	If Not objTempRs2.Eof Then
		If CheckIntegerAndZero(objTempRs2(0))>0 Then
			bResult = True
		End If
	End If
	objTempRs2.Close	
	Set objTempRs2 = Nothing
	
	GetExpertTopExpertByUid = bResult
End Function

Function GetExpertCompanyTopExpertByUid(AExpertUid, ACompanyID, AUserID)
	Dim iResult
	iResult = 0

	Dim objTempRs2
	Set objTempRs2 = GetDataRecordsetSP("usp_" & sIcaServerSqlPrefix & "ExpertCompanyTopExpertByUidSelect", Array( _
		Array(, adVarChar, 40, AExpertUid), _
		Array(, adInteger, , ACompanyID), _
		Array(, adInteger, , AUserID)))

	If Not objTempRs2.Eof Then
		iResult = CheckIntegerAndZero(objTempRs2(0))
	End If
	objTempRs2.Close	
	Set objTempRs2 = Nothing
	
	GetExpertCompanyTopExpertByUid = iResult
End Function

Function GetExpertIsTopExpertByUid(AExpertUid, AUserID)
	Dim dictOtherMember, objTempRs2, dKey, dValue
	Set dictOtherMember = Server.CreateObject("Scripting.Dictionary")
	Set objTempRs2 = GetDataRecordsetSP("usp_" & sIcaServerSqlPrefix & "ExpertIsTopExpertByUidSelect", Array( _
		Array(, adVarChar, 40, AExpertUid), _
		Array(, adInteger, , AUserID)))

	If Not objTempRs2.Eof Then
		While Not objTempRs2.Eof
			dKey = objTempRs2(1)
			dValue = CheckIntegerAndZero(objTempRs2(0))
			If dKey > "" And dValue > 0 Then
				If dictOtherMember.Exists(dKey) = false Then
					dictOtherMember.Add dKey, dValue
				End If
			End If
			objTempRs2.MoveNext
		Wend
	End If
	objTempRs2.Close	
	Set objTempRs2 = Nothing
	
	Set GetExpertIsTopExpertByUid = dictOtherMember
End Function

Function GetExpertCompanyCircleByUid(AExpertUid, ACompanyID, AUserID)
	Dim bResult
	bResult = False

	Dim objTempRs2
	Set objTempRs2 = GetDataRecordsetSP("usp_" & sIcaServerSqlPrefix & "ExpertCompanyCircleByUidSelect", Array( _
		Array(, adVarChar, 40, AExpertUid), _
		Array(, adInteger, , ACompanyID), _
		Array(, adInteger, , AUserID)))

	If Not objTempRs2.Eof Then
		If CheckIntegerAndZero(objTempRs2(0)) > 0 Then
			bResult = True
		End If
	End If
	objTempRs2.Close	
	Set objTempRs2 = Nothing
	
	GetExpertCompanyCircleByUid = bResult
End Function


Sub ShowExpertMethodologyFlag(AExpertUID, AExpertDB, AExpertID, ACvValidForMemberOrExpert)
	Dim objTempRs3, bExpertMethodologyVisible
	bExpertMethodologyVisible = False

	If bUserAccessMethodology Then
	On Error Resume Next
		' Extract status from ExpertsMetho
		Set objTempRs3 = GetDataRecordsetSP("usp_Ica_ExpertMethoFlagByIDSelect", Array( _
			Array(, adInteger, , AExpertDB.ID), _
			Array(, adInteger, , AExpertID)))
		If Not objTempRs3.Eof Then
			If Len(objTempRs3("expMethoSelectedFlag")) > 1 Then
				bExpertMethodologyVisible = True
				Dim sFlagMetho
				sFlagMetho = LCase(FlagMethoTitleByID(objTempRs3("expMethoSelectedFlag")))
				%>
				<tr>
				<td class="field splitter"><p>&nbsp;</p></td>
				<td class="value flag_<% =sFlagMetho %>_bg"><p>
				<% If bUserAccessModifyMethodology = True Then %>
				<a href="javascript:editMethodology('<% =AExpertUID %>', '<% =AExpertDB.ID %>', '<% =AExpertID %>');"><img src="/image/modify.gif" hspace="12" align="right"></a>
				<% End If %>
				<img src="/image/flag_<% =sFlagMetho %>.gif" alt="<% =sFlagMetho %>" vspace="2" hspace="3" width="7" height="12" border="0" align="left">
				<b>Methodology writer</b>
				<%
				Dim sMethoFields, sMethoLanguages
				sMethoFields = ""
				If objTempRs3("expMethoTA") = 1 Then 
					sMethoFields = sMethoFields & "TA"
				End If
				If objTempRs3("expMethoFWC") = 1 Then 
					If Len(sMethoFields) > 0 Then sMethoFields = sMethoFields & ", "
					sMethoFields = sMethoFields & "FWC"
				End If
				If Len(sMethoFields) > 0 Then sMethoFields = " <b>" & sMethoFields & "</b>"
				Response.Write sMethoFields & "</p>"
				
				If objTempRs3("expMethoShowAll") = "1" Then
					sMethoFields = ""
					If objTempRs3("expMethoCountTa") >= 1 _
					Or objTempRs3("expMethoCountFwc") >= 1 _
					Or objTempRs3("expMethoCountGrant") >= 1 _
					Or objTempRs3("expMethoCountType") >= 1 _
					Then
						sMethoFields = sMethoFields & "<p class=""txt"">Methodologies #<br>"
					End If

					If objTempRs3("expMethoCountType") = 1 Then 
						sMethoFields = sMethoFields & "&nbsp; &nbsp; between 1 and 4"
					End If
					If objTempRs3("expMethoCountType") = 2 Then 
						sMethoFields = sMethoFields & "&nbsp; &nbsp; between 5 and 8"
					End If
					If objTempRs3("expMethoCountType") = 3 Then 
						sMethoFields = sMethoFields & "&nbsp; &nbsp; more than 8"
					End If
					If objTempRs3("expMethoCountType") >= 1 Then
						sMethoFields = sMethoFields & "<br>"
					End If

					If objTempRs3("expMethoCountTa") = 1 Then 
						sMethoFields = sMethoFields & "&nbsp; &nbsp; TA: between 1 and 4"
					End If
					If objTempRs3("expMethoCountTa") = 2 Then 
						sMethoFields = sMethoFields & "&nbsp; &nbsp; TA: between 5 and 8"
					End If
					If objTempRs3("expMethoCountTa") = 3 Then 
						sMethoFields = sMethoFields & "&nbsp; &nbsp; TA: more than 8"
					End If
					If objTempRs3("expMethoCountTa") >= 1 Then
						sMethoFields = sMethoFields & "<br>"
					End If

					If objTempRs3("expMethoCountFwc") = 1 Then 
						sMethoFields = sMethoFields & "&nbsp; &nbsp; FWC: between 1 and 4"
					End If
					If objTempRs3("expMethoCountFwc") = 2 Then 
						sMethoFields = sMethoFields & "&nbsp; &nbsp; FWC: between 5 and 8"
					End If
					If objTempRs3("expMethoCountFwc") = 3 Then 
						sMethoFields = sMethoFields & "&nbsp; &nbsp; FWC: more than 8"
					End If
					If objTempRs3("expMethoCountFwc") >= 1 Then
						sMethoFields = sMethoFields & "<br>"
					End If

					If objTempRs3("expMethoCountGrant") = 1 Then 
						sMethoFields = sMethoFields & "&nbsp; &nbsp; Grants: between 1 and 4"
					End If
					If objTempRs3("expMethoCountGrant") = 2 Then 
						sMethoFields = sMethoFields & "&nbsp; &nbsp; Grants: between 5 and 8"
					End If
					If objTempRs3("expMethoCountGrant") = 3 Then 
						sMethoFields = sMethoFields & "&nbsp; &nbsp; Grants: more than 8"
					End If
					If objTempRs3("expMethoCountGrant") >= 1 Then
						sMethoFields = sMethoFields & "<br>"
					End If

					Response.Write sMethoFields & "</p>"
				End If

				sMethoFields = ""
				If objTempRs3("expMethoContribRev") = 1 Then 
					sMethoFields = sMethoFields & "Review"
				End If
				If objTempRs3("expMethoContribTech") = 1 Then 
					If Len(sMethoFields) > 0 Then sMethoFields = sMethoFields & ", "
					sMethoFields = sMethoFields & "Technical inputs"
				End If
				If objTempRs3("expMethoContribFull") = 1 Then 
					If Len(sMethoFields) > 0 Then sMethoFields = sMethoFields & ", "
					sMethoFields = sMethoFields & "Full methodology"
				End If
				If objTempRs3("expMethoContribRead") = 1 Then 
					If Len(sMethoFields) > 0 Then sMethoFields = sMethoFields & ", "
					sMethoFields = sMethoFields & "Proofreading and editing"
				End If

				sMethoLanguages = ""
				If objTempRs3("expMethoEN") = 1 Then 
					sMethoLanguages = sMethoLanguages & "English, "
				End If
				If objTempRs3("expMethoFR") = 1 Then 
					sMethoLanguages = sMethoLanguages & "French, "
				End If
				If objTempRs3("expMethoSP") = 1 Then 
					sMethoLanguages = sMethoLanguages & "Spanish, "
				End If
				If objTempRs3("expMethoPT") = 1 Then 
					sMethoLanguages = sMethoLanguages & "Portuguese, "
				End If
				If objTempRs3("expMethoDE") = 1 Then 
					sMethoLanguages = sMethoLanguages & "German, "
				End If
				If Len(sMethoLanguages) > 2 Then 
					sMethoLanguages = Left(sMethoLanguages, Len(sMethoLanguages) - 2)
					sMethoLanguages = "&nbsp;(" & sMethoLanguages & ")"
				End If

				If Len(sMethoFields) > 0 Then sMethoFields = "<p class=""txt"">" & sMethoFields & sMethoLanguages & "</p>"
				Response.Write sMethoFields
				%>

				<% If Len(objTempRs3("expMethoComments")) > 2 Then %>
					<p class="txt">Comments: <% =HtmlEncode(objTempRs3("expMethoComments")) %></p>
				<% End If %>

				<% If objTempRs3("expMethoShowAll") = "1" Then %>
					<% If Len(objTempRs3("expMethoExpComments")) > 2 Then %>
						<p class="txt"><% If Len(objTempRs3("expMethoComments")) > 2 Then %>&nbsp; &nbsp; Expert<% Else %>Expert's comments<% End If %>:
						<% = HtmlEncode(objTempRs3("expMethoExpComments")) %></p>
					<% End If %>

					<% If Len(objTempRs3("expMethoKeywords")) > 2 Then %>
						<p class="txt">Writing experience: <% =HtmlEncode(objTempRs3("expMethoKeywords")) %></p>
					<% End If %>
					<% If Len(objTempRs3("expMethoExpKeywords")) > 2 Then %>
						<p class="txt"><% If Len(objTempRs3("expMethoKeywords")) > 2 Then %>&nbsp; &nbsp; Expert<% Else %>Expert's writing experience<% End If %>:
						<% =HtmlEncode(objTempRs3("expMethoExpKeywords")) %></p>
					<% End If %>


					<% If Len(objTempRs3("expMethoDonors")) > 2 Then %>
						<p class="txt">Donors: <% = HtmlEncode(objTempRs3("expMethoDonors")) %></p>
					<% End If %>
				<% End If %>

				</td>
				</tr>
				<%
			End If
		End If
		If Not bExpertMethodologyVisible Then
			%>
				<tr>
				<td class="field splitter"><p>&nbsp;</p></td>
				<td class="value"><p>
				<% If bUserAccessModifyMethodology = True Then %>
					<a href="javascript:editMethodology('<% =AExpertUID %>', '<% =AExpertDB.ID %>', '<% =AExpertID %>');"><img src="/image/add_metho.gif" hspace="12" align="right"></a>
				<% End If %>
				</p></td>
				</tr>
			<%
		End If
		objTempRs3.Close
		Set objTempRs3 = Nothing
	On Error GoTo 0	
	End If
End Sub


Sub ShowExpertMpisFlag(AExpertDB, AExpertID, ACvValidForMemberOrExpert)
Dim objTempRs3
	If bUserAccessMpis=1 Then
	On Error Resume Next
		' Extract status from MPIS
		Set objTempRs3 = GetDataRecordsetSP("usp_Ica_ExpertMpisFlagByIDSelect", Array( _
			Array(, adInteger, , AExpertDB.ID), _
			Array(, adInteger, , AExpertID)))
		If Not objTempRs3.Eof Then
			If Len(objTempRs3("CONTACT_RELATION_STATUS_FLAG"))>2 Then
				%>
				<tr>
				<td class="field splitter"><p>&nbsp;</p></td>
				<td class="value flag_<% =objTempRs3("CONTACT_RELATION_STATUS_FLAG") %>_bg"><p>
				<img src="/image/flag_<% =objTempRs3("CONTACT_RELATION_STATUS_FLAG") %>.gif" alt="<% =objTempRs3("CONTACT_RELATION_STATUS_VALUE") %>"  vspace="2" hspace="3" width="7" height="12" border="0" align="left">
				<b><% =objTempRs3("CONTACT_RELATION_STATUS_VALUE") %></b>
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

	If bUserAccessMpis=1 Then
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

Function GetAddedByUserId(AExpertUid, ACompanyID)
	Dim iResult
	iResult = 0

	Dim objTempRs2
	Set objTempRs2 = GetDataRecordsetSP("usp_" & sIcaServerSqlPrefix & "ExpertGetAddedByUserIdForCompany", Array( _
		Array(, adVarChar, 40, AExpertUid), _
		Array(, adInteger, , ACompanyID)))

	If Not objTempRs2.Eof Then
		iResult = CheckIntegerAndZero(objTempRs2(0))
	End If
	objTempRs2.Close	
	Set objTempRs2 = Nothing
	
	GetAddedByUserId = iResult
End Function
%>
