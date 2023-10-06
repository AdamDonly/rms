<%
'--------------------------------------------------------------------
'
' Common functions for all users.
'
'--------------------------------------------------------------------
Dim iFlag
Dim sSearchKeywordsHighlight, arrSearchKeywordsHighlight, sKeyword

Function HighlightKeywords(sText, sSearchKeywordsHighlight)
Dim sResult, i
sResult=sText
i=1
	If Len(sSearchKeywordsHighlight)>2 Then 
	For Each sKeyword in arrSearchKeywordsHighlight
                If Len(sResult)>2 Then sResult=Replace(sResult, sKeyword & " ", "<span class=""marktext" & i & """>" & sKeyword & "</span> ", 1, -1, 1)
                If Len(sResult)>2 Then sResult=Replace(sResult, sKeyword & ".", "<span class=""marktext" & i & """>" & sKeyword & "</span>.", 1, -1, 1)
                If Len(sResult)>2 Then sResult=Replace(sResult, sKeyword & ",", "<span class=""marktext" & i & """>" & sKeyword & "</span>,", 1, -1, 1)
                If Len(sResult)>2 Then sResult=Replace(sResult, sKeyword & ")", "<span class=""marktext" & i & """>" & sKeyword & "</span>)", 1, -1, 1)
                If Len(sResult)>2 Then sResult=Replace(sResult, sKeyword & "/", "<span class=""marktext" & i & """>" & sKeyword & "</span>/", 1, -1, 1)
                If Len(sResult)>2 Then sResult=Replace(sResult, sKeyword & """", "<span class=""marktext" & i & """>" & sKeyword & "</span>""", 1, -1, 1)
                If Len(sResult)>2 Then sResult=Replace(sResult, "/" & sKeyword, "/<span class=""marktext" & i & """>" & sKeyword & "</span>", 1, -1, 1)
		If i<=6 Then i=i+1
	Next
	End If
HighlightKeywords=sResult
End Function

'--------------------------------------------------------------------
' Check user's access rights for the actual application
'--------------------------------------------------------------------
Function CheckUserLogin(sScriptNameIn)
	If (Not iUserID>0) _
		Or (sApplicationName="expert" 		And iExpertID=0) _
		Or (sApplicationName="backoffice" 	And sUserType<>sApplicationName And sUserType<>"Admin" And sUserType<>"CV Assortis") _
		Or (sApplicationName="outsourcing" 	And sUserType<>sApplicationName And sUserType<>"backoffice") _
		Or (sApplicationName="external" 	And sUserType<>sApplicationName And sUserType<>"backoffice") _
	Then
		If Len(sScriptNameIn)>0 Then sScriptNameIn="?" & sScriptNameIn
		
		Dim sMainServerName
		If InStr(sScriptServerName, "test.")>0 Then
			sMainServerName=Replace(sScriptServerName, "experts.", "")
		Else
			sMainServerName=Replace(sScriptServerName, "experts.", "www.")
		End If
		
		sUrl=sScriptServerProtocol & sMainServerName & "/login.asp"
		Response.Redirect sUrl
		
	End If
End Function 


'--------------------------------------------------------------------
' Check 
'--------------------------------------------------------------------
Function IsIcaUserCompanyCvValid(ADatabase, iCvID, AUserCompanyDatabase)
Dim bCvValidTemp, objTempRs
	bCvValidTemp=0
	If IsNumeric(iCvID) Then
		iCvID=CLng(iCvID)
	End If

	Set objTempRs=GetDataRecordsetSP("usp_Ica_ExpertDBAccessValidSelect", Array( _
		Array(, adVarChar, 25, ADatabase), _
		Array(, adInteger, , iCvID), _
		Array(, adVarChar, 25, AUserCompanyDatabase)))

	If Not objTempRs.Eof Then
	On Error Resume Next
		bCvValidTemp=objTempRs("ExpertDbAccessValid")
	On Error GoTo 0
	End If
objTempRs.Close 

IsIcaUserCompanyCvValid=bCvValidTemp
End Function


Function CheckExpertID()

	' If it is not an expert application - get iExpertID from QueryString
	If sApplicationName<>"expert" Then
		iExpertID=0
		sExpertUid=Request.QueryString("uid")
		' if a Top Expert is operating - get the UID from the account, to not pass it in the URL:
		If bIsMyCV And Len(uUserTopExpertUid) > 5 Then
			sExpertUid = uUserTopExpertUid
		End If
		If Len(sExpertUid)>0 Then
		On Error Resume Next
			Set objTempRs=GetDataRecordsetSP("usp_Ica_ExpertIdSelect", Array( _
				Array(, adVarChar, 40, sExpertUid)))

			If Err.Number<>0 Or objTempRs.Eof Then
				Response.Redirect "/"
			End If

			iExpertID = objTempRs("id_Expert")
			Set objExpertDB = objExpertDBList.Find(objTempRs("id_Database"), "ID")
			
		Set objTempRs=Nothing
		On Error GoTo 0
		End If
	End If

	' For backoffice and outsourcing teams - verify that this expert does not exist in the database already
	' for external registration team it should be the same if names of experts are visible for them - depends on the agreement with particular client
	If sApplicationName="backoffice" Or sApplicationName="outsourcing" _
		Or (sApplicationName="external" And sContactDetailsExternally=cNameVisible) _
		Then
		If iExpertID=0 Then
			Response.Redirect "verify.asp" & ReplaceUrlParams("", sScriptFullNameAsParams)
		End If
	End If

	If InStr(sScriptFileName, "cv_copy.asp")<=0 And IsIcaUserCompanyCvValid(objExpertDB.Database, iExpertID, objUserCompanyDB.Database)<>1 Then
		Response.Redirect "/"
	End If

	CheckExpertOriginalID()
End Function


Dim iExpertOriginalID
Function CheckExpertOriginalID()
	' Redirect to original CV if expert's CV was updated with another ID (Blacklist=1 & id_ExpertOriginal>0)
	objTempRs=GetDataOutParamsSP("usp_ExpCvvOriginalSelect", Array( _
		Array(, adInteger, , iExpertID)), _
		Array( Array(, adInteger)))
	iExpertOriginalID=objTempRs(0)

	If iExpertOriginalID>0 Then 
		iExpertID=iExpertOriginalID
	End If
End Function


Function ObfuscateString(sInputString)
Dim sTempString, sObfuscateChar
sTempString=sInputString
sObfuscateChar="x"
	If Len(sTempString)>1 Then
		sTempString=Left(sTempString, 1) & String(Len(sTempString)-1, sObfuscateChar)
	End If
ObfuscateString=sTempString
End Function

Function ObfuscateEmail(sInputString)
Dim sTempString, iPosition
sTempString=sInputString
iPosition=InStr(sTempString, "@")
	If iPosition>1 Then
		sTempString=ObfuscateString(Left(sTempString, iPosition-1)) & Mid(sTempString, iPosition, Len(sTempString)-iPosition+1)
	End If
ObfuscateEmail=sTempString
End Function



'--------------------------------------------------------------------
' Procedure for creating user's login and password from email 
'--------------------------------------------------------------------
Sub CreateLoginAndPassword(sUserEmail)
	i=InStr(sUserEmail, "@")
	If i>5 And iFlag=0 Then
		sUserLogin=LCase(Left(sUserEmail, i-1))
	ElseIf i<3 Then
		sUserLogin= sUserEmail & Int(Rnd(99)*100)
	ElseIf i>3 And iFlag>0 Then
		sUserLogin=LCase(Left(sUserEmail, i-1) & Int(Rnd(99)*100))
	Else
		sUserLogin=LCase(Left(sUserEmail, i-1) & Second(Now) & Int(Rnd(99)*100) & Mid(sSessionID, 3, i-3))
	End If
	If Len(sUserLogin)>15 And iFlag<3 Then
		sUserLogin=Left(sUserLogin, 15)
	End If
	sUserPassword=LCase(Left(sUserLogin,1) & Second(Now) & Mid(sSessionID, 28, 5))
End Sub

'''''''''''''''''''''
Function GetTotalCVs()
Dim iTotalExperts
	objTempRs=GetDataOutParamsSP("usp_AdmExpTotalSelect", Array(), _
		Array( Array(, adInteger)))
	iTotalExperts=objTempRs(0)
	Set objTempRs=Nothing	
GetTotalCVs = iTotalExperts
End function   


' ----------------------------------------------------------------------------

Sub ShowNavigationPages(iActivePage, iTotalPages, sParamsIn)
If iTotalPages<2 Or sAccessType="trial" Then 
	Response.Write("<img src=""" & sHomePath & "image/x.gif"" width=1 height=5><br>") 
	Exit Sub
End If
Dim iPage, PagesPhrase, iStartPage, iEndPage, iPrevPage, iNextPage, iMaxPagesOnScreen
Dim sTmpParams
Dim sImagesFolder

sImagesFolder=sHomePath & "image/"
iMaxPagesOnScreen=15
iStartPage=iMaxPagesOnScreen * ((iActivePage-1) \ iMaxPagesOnScreen) + 1
iEndPage=iMaxPagesOnScreen * ((iActivePage-1) \ iMaxPagesOnScreen + 1)
iPrevPage=iStartPage-1
iNextPage=iEndPage+1

If iEndPage > iTotalPages Then
	iEndPage=iTotalPages
End If

Response.Write("<table width=""100%"" cellpadding=0 cellspacing=0 border=0><tr><td width=""50%""><img src=""" & sImagesFolder & "x.gif"" width=42 height=1 vspace=25></td><td><p class=""txt"">Pages:&nbsp;&nbsp;</td>")

If iPrevPage>0 Then
	sTmpParams=AddUrlParams(sParams, "page=" & iPrevPage)
	Response.Write "<td align=""center""><p class=""sml""><a href = """ & sScriptFileName & sTmpParams & """><img src=""" & sImagesFolder & "mprev.gif"" width=7 height=15 hspace=1 vspace=4 border=0 alt=""Previous pages""></a><br>&nbsp;&nbsp;<a href = """ & sScriptFileName & sTmpParams & """>Prev</a>&nbsp;&nbsp;</td>"
End If 

For iPage=iStartPage To iEndPage
If iPage=iActivePage Then
	Response.Write "<td align=""center""><p class=""sml""><img src=""" & sImagesFolder & "act_mpage.gif"" width=12 height=15 hspace=2 vspace=3 border=0 alt=""Page " & iPage & " - Active""><br>" & iPage & "</td>"
Else
	sTmpParams=AddUrlParams(sParams, "page=" & iPage)
	On Error Resume Next
		sTmpParams=AddUrlParams(sTmpParams, "act=" & sAction)
		sTmpParams=AddUrlParams(sTmpParams, "ord=" & sOrderBy)
	On Error Goto 0
	Response.Write "<td align=""center""><p class=""sml""><a href=""" & sScriptFileName & sTmpParams & """><img src=""" & sImagesFolder & "mpage.gif"" width=12 height=15 hspace=2 vspace=3 border=0 alt=""Page " & iPage & """><br>" & iPage & "</a></td>"
End If
Next

If iNextPage < iTotalPages Then
	sTmpParams=AddUrlParams(sParams, "page=" & iNextPage)
	Response.Write "<td align=""center""><p class=""sml""><a href = """ & sScriptFileName & sTmpParams & """><img src=""" & sImagesFolder & "mnext.gif"" width=7 height=15 hspace=1 vspace=3 border=0 alt=""Next pages""></a><br>&nbsp;&nbsp;<a href = """ & sScriptFileName & sTmpParams & """>Next</a>&nbsp;&nbsp;</td>"
End If 

Response.Write("<td width=""50%"">&nbsp;</td></tr></table>")
End Sub


' ----------------------------------------------------------------------------

Sub ShowOrgBSCInfo(iMemberID, iSearchQueryID, iOrganisationID, sBSCType, sListType, sRange)

Dim objTempRs2
Dim i, iShowMaxRecords, sBSCTypeFullName, sBSCTypeShortName, sBSCStyle, sBSCDate, sPath
Dim sClass
If sBSCType="CA" Then
	sBSCTypeFullName="contracts awarded"
	sBSCTypeShortName="contract"
Else
	sBSCTypeFullName="times shortlisted"
	sBSCTypeShortName="shortlist"
End If

If sRange="show all" Then
	iShowMaxRecords=30000
	sPath=""
ElseIf sRange="show top 2" Then
	iShowMaxRecords=2
	sPath=""
ElseIf sRange="show top 0" Then
	iShowMaxRecords=0
	sPath=""
End If

Dim sStoredProcedureName


If sUserType="expert" Then
	sStoredProcedureName="sp_ExpPrfSearchBSCSelect"
Else
	sStoredProcedureName="usp_MmbPrfOrgBusopsSelect"
End If
Set objTempRs2=GetDataRecordsetSP(sStoredProcedureName, Array( _
	Array(, adInteger, , iMemberID), _
	Array(, adInteger, , iSearchQueryID),_
	Array(, adInteger, , iOrganisationID),_
	Array(, adVarChar, 3 , sBSCType),_
	Array(, adVarChar, 12 , sListType)))

If sListType="other" Then
	sBSCStyle="class=""bl"""
	sListType="other available"
Else
	sBSCStyle=""
End If

	i=0
	If not objTempRs2.Eof Then
		If sRange<>"show all" Then
			If objTempRs2.Recordcount=1 Then
				Response.Write("<p class=""txt"">" & sBSCTypeFullName & ": once <br>")
			ElseIf objTempRs2.Recordcount>1 Then
				Response.Write("<p class=""txt"">" & sBSCTypeFullName & ": " & objTempRs2.Recordcount & " times <br>")
			End If
		Else
			If iSearchQueryID>0 And objTempRs2("CountMatchSearchCriteria")>0 And objTempRs2("CountMatchSearchCriteria")<objTempRs2.Recordcount Then
				sBSCTypeFullName=sBSCTypeFullName & " (" & ShowEntityPlural(objTempRs2("CountMatchSearchCriteria"), sBSCTypeShortName, sBSCTypeShortName & "s", " ") & " matching the search criteria)"
			End If
			Response.Write("<p class=""txt""><b>" & objTempRs2.Recordcount & " " & sListType & " " & sBSCTypeFullName & "</b></p></a>")
       	End If
	End If
	While (Not objTempRs2.Eof) And (i<iShowMaxRecords)
		If IsDate(objTempRs2(3)) Then
			sBSCDate=ConvertDateForText(objTempRs2(3), "&nbsp;", "DayMonthYear")
		Else
			sBSCDate=""
		End If

		If objTempRs2("MatchSearchCriteria")=0 Then sBSCStyle="class=""nomatch""" Else sBSCStyle=""
		
		If sBSCType="CA" Then
			Response.Write("<p class=""sml""><img src=""../../image/b.gif"" width=5 height=5 vspace=5 hspace=6 align=""left""><a " & sBSCStyle & " href=""" & sPath & "bsc_view.asp?id=" & objTempRs2("id_Busop") & "&DataType=" & sBSCTypeShortName & "&caid=" & objTempRs2("id_Contract") & """>" & objTempRs2("bspTitle") & "</a><i> (" & objTempRs2("DonorCode") & " - " & objTempRs2("CountryCode") & " - " & sBSCDate & ")</i></p>")
		Else
			Response.Write("<p class=""sml""><img src=""../../image/b.gif"" width=5 height=5 vspace=5 hspace=6 align=""left""><a " & sBSCStyle & " href=""" & sPath & "bsc_view.asp?id=" & objTempRs2("id_Busop") & "&DataType=" & sBSCTypeShortName & """>" & objTempRs2("bspTitle") & "</a><i> (" & objTempRs2("DonorCode") & " - " & objTempRs2("CountryCode") & " - " & sBSCDate & ")</i></p>")
		End If
	objTempRs2.MoveNext
	i=i+1
	WEnd
	If i>0 Then
		Response.Write("<img src=""../image/x.gif"" width=250 height=8><br>")
	End If
objTempRs2.Close 
Set objTempRs2= Nothing
End Sub


Function CreateSearchString(sSearchString, sSearchStringType)
Dim sProjectKeywords, sProjectKeywordsType, arrKeywords, i
If Not IsNull(sSearchString) Then
	If Len(sSearchString)=1 Then sSearchString=""
End If
sProjectKeywords=Trim(sSearchString)
sProjectKeywordsType=sSearchStringType

	If sProjectKeywords>"" Then
		sProjectKeywords=Replace(sProjectKeywords, ",", " ")
		sProjectKeywords=Replace(sProjectKeywords, ".", " ")
		sProjectKeywords=Replace(sProjectKeywords, ":", " ")
		sProjectKeywords=Replace(sProjectKeywords, ";", " ")
		sProjectKeywords=Replace(sProjectKeywords, "-", " ")
		sProjectKeywords=Replace(sProjectKeywords, "+", " ")
		sProjectKeywords=Replace(sProjectKeywords, "*", " ")
		sProjectKeywords=Replace(sProjectKeywords, "/", " ")
		sProjectKeywords=Replace(sProjectKeywords, "<", " ")
		sProjectKeywords=Replace(sProjectKeywords, ">", " ")
		sProjectKeywords=Replace(sProjectKeywords, "%", " ")

		sProjectKeywords=Replace(sProjectKeywords, "'", "")
		sProjectKeywords=Replace(sProjectKeywords, """", "")
		sProjectKeywords=Replace(sProjectKeywords, "   ", " ")
		sProjectKeywords=Replace(sProjectKeywords, "  ", " ")

                sProjectKeywords=" " & sProjectKeywords & " "

		// Removing from keyword all single letters
		arrKeywords=Split(sProjectKeywords, " ")
		For i=1 To UBound(arrKeywords)
			If Len(arrKeywords(i))=1 Then sProjectKeywords=Replace(sProjectKeywords, " " & arrKeywords(i) & " ", " ")
		Next

		' ?? sSearchFullText="""" & sSearchFullText & sProjectKeywords & """ AND "

	' All of the words
		If sProjectKeywordsType="all of the words from" Then
			sProjectKeywords=Replace(sProjectKeywords, " AND ", " ",1,1000,1)
			sProjectKeywords=Replace(sProjectKeywords, " OR ", " ",1,1000,1)
			sProjectKeywords=Replace(sProjectKeywords, " NOT ", " ",1,1000,1)
			sProjectKeywords=Replace(sProjectKeywords, " NEAR ", " ",1,1000,1)
			While InStr(sProjectKeywords, "  ")>0
				sProjectKeywords=Replace(sProjectKeywords, "  ", " ")
			WEnd
			sProjectKeywords=LTrim(RTrim(sProjectKeywords))
			sProjectKeywords="""" & Replace(sProjectKeywords, " ", """ AND """) & """"

	' Any of the words
		ElseIf sProjectKeywordsType="any of the words from" Then
			sProjectKeywords=Replace(sProjectKeywords, " AND ", " ",1,1000,1)
			sProjectKeywords=Replace(sProjectKeywords, " OR ", " ",1,1000,1)
			sProjectKeywords=Replace(sProjectKeywords, " NOT ", " ",1,1000,1)
			sProjectKeywords=Replace(sProjectKeywords, " NEAR ", " ",1,1000,1)
			While InStr(sProjectKeywords, "  ")>0
				sProjectKeywords=Replace(sProjectKeywords, "  ", " ")
			WEnd
			sProjectKeywords="""" & Replace(sProjectKeywords, " ", """ OR """) & """"

	' The exact phrase
		ElseIf sProjectKeywordsType="the exact phrase" Then
			sProjectKeywords="""" & sProjectKeywords & """"

	' Boolean expression
		Else

		' Replacing all the variants of the word "and" to "AND" and then placing quotes
		sProjectKeywords=Replace(sProjectKeywords, " AND ", " AND ",1,1000,1)
		sProjectKeywords=Replace(sProjectKeywords, " AND ", """ AND """)

		' Replacing all the variants of the word "or" to "OR" and then placing quotes
		sProjectKeywords=Replace(sProjectKeywords, " OR ", " OR ",1,1000,1)
		sProjectKeywords=Replace(sProjectKeywords, " OR ", """ OR """)

		sProjectKeywords=Replace(sProjectKeywords, " NEAR ", " NEAR ",1,1000,1)
		sProjectKeywords=Replace(sProjectKeywords, " NEAR ", """ NEAR """)
		sProjectKeywords="""" & sProjectKeywords & """"
		End If
	End If

CreateSearchString=sProjectKeywords
End Function



Function CreateAccountReferenceNumber(iServiceID, iMemberID, iExpertID)
Dim sReferenceTemp
	If iMemberID>0 Then
		sReferenceTemp="M" & iMemberID
	ElseIf iExpertID>0 Then
		sReferenceTemp="E" & iExpertID
	Else
		sReferenceTemp="I"
	End If
	sReferenceTemp=iServiceID & sReferenceTemp &"_" & Hour(Now()) & Minute(Now()) & Second(Now())
CreateAccountReferenceNumber=sReferenceTemp
End Function



Function RoundTo5(iAmount)
Dim iAddAmount, iRoundAmount, iLastDigit
iRoundAmount=Round(iAmount)
iLastDigit=Right(iRoundAmount,1)
Select Case iLastDigit
	Case 0 iAddAmount=0
	Case 1 iAddAmount=4
	Case 2 iAddAmount=3
	Case 3 iAddAmount=2
	Case 4 iAddAmount=1
	Case 5 iAddAmount=0
	Case 6 iAddAmount=4
	Case 7 iAddAmount=3
	Case 8 iAddAmount=2
	Case 9 iAddAmount=1
End Select
RoundTo5=iRoundAmount+iAddAmount
End Function

Function ShowActLngName(Field,Lng,MasterLng)
Dim tmpField
	If Eval(Field & Lng)>"" Then
		tmpField=Eval(Field & Lng)
	ElseIf Eval(Field & MasterLng)>"" Then
		tmpField=Eval(Field & MasterLng)
	End If
ShowActLngName = tmpField 
End function   



'--------------------------------------------------------------------
'
' Functions for data types checking and converting 
'
'--------------------------------------------------------------------

' Function for converting date to format YYYY/MM/DD from year, month and day values.
' It is used everywhere to transfer data to sql db
'
Function ConvertDMYForSql(sYear, sMonth, sDay)
Dim dTempDate
dTempDate=Null

If IsNumeric(sYear) And IsNumeric(sMonth) And IsNumeric(sDay) Then
	If Not (sDay>=1 And sDay<=31) Then sDay=1
	If Not (sMonth>=1 And sMonth<=12) Then sMonth=1

	If (sMonth=2) And (sDay>28) Then
		If sYear Mod 4 =0 Then
			sDay=29
		Else
			sDay=28
		End If
	End If
	If (sMonth=4 Or sMonth=6 Or sMonth=9 Or sMonth=11) And (sDay>30) Then sDay=30

	If (sYear>=1800 And sYear<=2100) Then
		dTempDate=sYear & "/" & Left("00", 2-Len(sMonth)) & sMonth & "/" & Left("00", 2-Len(sDay)) & sDay
	End If
End If
ConvertDMYForSql=dTempDate
End Function


' Function for converting date to format YYYY/MM/DD from date value.
' It is used everywhere to transfer data to sql db
'
Function ConvertDateForSql(sDate)
Dim dTempDate, sYear, sMonth, sDay
	If IsDate(sDate) Then
		sYear=Year(sDate)
		sMonth=Month(sDate)
		sDay=Day(sDate)
		dTempDate=sYear & "/" & Left("00", 2-Len(sMonth)) & sMonth & "/" & Left("00", 2-Len(sDay)) & sDay
	Else
		dTempDate=Null
	End If
ConvertDateForSql=dTempDate
End Function

Function ConvertDateTimeForSql(sDate)
Dim dTempDate, sYear, sMonth, sDay, sHour, sMinute, sSecond
	If IsDate(sDate) Then
		sYear=Year(sDate)
		sMonth=Month(sDate)
		sDay=Day(sDate)
		sHour=Hour(sDate)
		sMinute=Minute(sDate)
		sSecond=Second(sDate)

		dTempDate=sYear & "/" & Left("00", 2-Len(sMonth)) & sMonth & "/" & Left("00", 2-Len(sDay)) & sDay & "_" & Left("00", 2-Len(sHour)) & sHour & "" & Left("00", 2-Len(sMinute)) & sMinute & "" & Left("00", 2-Len(sSecond)) & sSecond
	Else
		dTempDate=Null
	End If
ConvertDateTimeForSql=dTempDate
End Function


' Function for showing date in text format
' MM - first 3 characters of the month name
'
Function ConvertDateForText(sDate, sDelimeter, sFormat)
Dim dDateTemp
	If IsDate(sDate) Then
		If sFormat="MMYYYY" Then
			dDateTemp=Left(arrMonthName(Month(sDate)),3) & sDelimeter & Year(sDate)
		ElseIf sFormat="DayMonthYear" Then
			dDateTemp=Day(sDate) & sDelimeter & arrMonthName(Month(sDate)) & sDelimeter & Year(sDate)
		ElseIf sFormat="MonthYear" Then
			dDateTemp=arrMonthName(Month(sDate)) & sDelimeter & Year(sDate)
	        ElseIf sFormat="DMY" Then 
			dDateTemp=Day(sDate) & sDelimeter & Month(sDate) & sDelimeter & Year(sDate)
	        Else ' DDMMYYYY
			dDateTemp=Day(sDate) & sDelimeter & Left(arrMonthName(Month(sDate)),3) & sDelimeter & Year(sDate)
		End If
	Else
		dDateTemp=""
	End If  
ConvertDateForText=dDateTemp
End Function



Function ConvertSQLDateToText(sSqlDate, sDelimeter, sFormat)
Dim sYearTemp, sMonthTemp, sDayTemp, sDateTemp
sDateTemp=""
	If Len(sSqlDate)=8 Then
		sYearTemp=Left(sSqlDate,4)
		sMonthTemp=Mid(sSqlDate,5,2)
		sDayTemp=Right(sSqlDate,2)

		sDateTemp= sYearTemp & "/" & sMonthTemp & "/" & sDayTemp
		If IsDate(sDateTemp) Then
			sDateTemp=ConvertDateForText(sDateTemp, sDelimeter, sFormat)
		Else
			sDateTemp=""
		End If
	End If

ConvertSQLDateToText=sDateTemp
End Function

' Function for changing empty values
'
Function ReplaceIfEmpty(sTextValue, sReplaceValue)
Dim sTextTemp
	sTextTemp=sTextValue
	If IsNull(sTextValue) Or Trim(sTextValue)="" Then 
		sTextTemp=sReplaceValue
	End If
ReplaceIfEmpty=sTextTemp
End Function


Function CheckLength(sTextInput)
Dim iLengthTemp
	If Not IsNull(sTextInput) Then
		iLengthTemp=Len(Trim(sTextInput))
	Else
		iLengthTemp=0
	End If
	CheckLength=iLengthTemp
End Function


' Function for converting text delemiters to html ones
' It replaces end of string delimiter to <br> tag
'
Function ConvertText(sTextInput)
Dim sTextTemp
	sTextTemp = sTextInput
	If Not IsNull(sTextTemp) Then
		sTextTemp=Replace(sTextTemp,CHR(13)+CHR(10),"<br>")
		sTextTemp=Replace(sTextTemp,"<br><br>","<br>")
		sTextTemp=Replace(sTextTemp,"&#61550;","*")
		
		sTextTemp=ReplaceHtmlSpecialCodes(sTextTemp, "* ")
		
		If Len(sTextTemp)>"4" Then
		If Right(sTextTemp, 4)="<br>" Then sTextTemp=Trim(Left(sTextTemp, Len(sTextTemp)-4))
		End If
	End If
ConvertText=sTextTemp
End Function 

Function ConvertTextForEmail(sTextInput)
Dim sTextTemp
	sTextTemp = sTextInput
	If Not IsNull(sTextTemp) Then
		sTextTemp=Replace(sTextTemp,CHR(13)+CHR(10),"<br>")
		sTextTemp=Replace(sTextTemp,"&#61550;","*")
		sTextTemp=Replace(sTextTemp,"&#61623;","*")
		If Len(sTextTemp)>"4" Then
		If Right(sTextTemp, 4)="<br>" Then sTextTemp=Trim(Left(sTextTemp, Len(sTextTemp)-4))
		End If
	End If
ConvertTextForEmail=sTextTemp
End Function 

Function ReadTextFile(sFileName)
Dim objFso, objFile, sTemp

	Set objFso=Server.CreateObject("Scripting.FileSystemObject")
	Set objFile=objFso.OpenTextFile(Server.MapPath(sHomePath) & sFileName, 1)
	sTemp=objFile.ReadAll
	objFile.Close
	Set objFile=Nothing
	Set objFso=Nothing
	
ReadTextFile=sTemp
End Function

' Function for converting amount to euro type
'
Function ConvertEuro(sAmount)
	ConvertEuro=ConvertMoneyToText(sAmount, "EUR", 2)
End Function 


Function ConvertMoneyToText(sAmount, sCurrency, iNumberOfSymbolsAfterDecimal)
Dim sAmountTemp, sDecimalPart, sDecimalSymbol, sGroupingSymbol
	If sCurrency="USD" Then
		sDecimalSymbol="."
		sGroupingSymbol=","
	Else 
		sDecimalSymbol="."
		sGroupingSymbol=" "
		sCurrency="EUR"
	End If

	If Not IsNumeric(sAmount) Then Exit Function
	
	' getting the integer and decimal parts of the amount
	If sAmount<>Int(sAmount) Then
		sDecimalPart=Abs(sAmount - Int(sAmount))
		sAmount=sAmount - sDecimalPart
		sDecimalPart=Round(10^iNumberOfSymbolsAfterDecimal * sDecimalPart)
	End If

	If iNumberOfSymbolsAfterDecimal=0 Then
               	sDecimalPart=""
	Else 
		If sDecimalPart="" Then sDecimalPart=Left("0000000000", iNumberOfSymbolsAfterDecimal)
	End If


	' inserting group delitemer symbol in integer part of the amount
	sAmountTemp=""
	If Len(sAmount)>3 Then
		While Len(sAmount)>=3
			sAmountTemp=Right(sAmount,3) & sGroupingSymbol & sAmountTemp
			sAmount=Left(sAmount, Len(sAmount)-3)
		WEnd
		sAmountTemp=sAmount & sGroupingSymbol & sAmountTemp
	Else
		sAmountTemp=sAmount	
	End If

	If sDecimalPart>"" Then
		sAmountTemp=sAmountTemp & sDecimalSymbol & sDecimalPart
	End If

ConvertMoneyToText=sAmountTemp
End Function 


' Function wrapping strings
'
Function CutString(strInput, cnum)
Dim strTemp, cp
	strTemp = strInput
	If Len(strTemp) > cnum Then
	   cp=InStrRev(LEFT(strTemp,cnum), "/")
	   If cp>0 Then 
		strTemp=LEFT(strTemp, cp) & "<br> &nbsp; &nbsp; &nbsp;" & RIGHT(strTemp, LEN(strTemp)-cp-1)
	   Else 
		   If cp<=0 Then cp=InStrRev(LEFT(strTemp,cnum), "@")
		   If cp<=0 Then cp=InStrRev(LEFT(strTemp,cnum), ".")
		   If cp<=0 Then cp=InStrRev(LEFT(strTemp,cnum), "-")
		   If cp>0 Then strTemp=LEFT(strTemp, cp) & "<br>" & RIGHT(strTemp, LEN(strTemp)-cp)
	   End If		
	End If
	CutString = strTemp
End Function 


' Function wrapping strings in menus
'
Function CutStringInMenu(sInput, iMaxLength, sCutBySymbol, sInsertAferCut)
Dim sTemp, iPos
	sTemp = sInput
	If Len(sTemp) > iMaxLength Then
		iPos=InStrRev(Left(sTemp,iMaxLength), sCutBySymbol)
		If iPos>0 Then 
			sTemp=Left(sTemp, iPos) & sInsertAferCut & Right(sTemp, Len(sTemp)-iPos)
		End If		
	End If
CutStringInMenu = sTemp
End Function 


' Function wrapping strings in menus
'
Function CutStringInMenu2(strInput, cnum)
Dim strTemp
	strTemp = strInput
	If Len(strTemp) > cnum Then
	   cp=InStrRev(LEFT(strTemp,cnum), " ")
	   If cp>0 Then
		strTemp=LEFT(strTemp, cp) & "<br>&nbsp;&nbsp;&nbsp;" & RIGHT(strTemp, LEN(strTemp)-cp)
	   End If	
	End If
	CutStringInMenu = strTemp
End Function 


' Function cutting strings
' 
Function CutStringNDelete(strInput, cnum)
Dim strTemp, cp
	strTemp = strInput
	If Len(strTemp) > cnum Then
	   cp=InStrRev(LEFT(strTemp,cnum), "/")
	   If cp>0 Then
		strTemp=LEFT(strTemp, cp-2) 
	   End If	
	End If
	CutStringNDelete = strTemp
End Function 

' Function cutting strings with spaces
' 
Function CutStringNDeleteAfterSpace(strInput, cnum)
Dim strTemp, cp, cp1, cp2, cp3
	strTemp = strInput
	If Len(strTemp) > cnum Then
	   cp=InStrRev(LEFT(strTemp,cnum), " ")
	   cp1=InStrRev(LEFT(strTemp,cnum), ".")
	   cp2=InStrRev(LEFT(strTemp,cnum), "/")
	   cp3=InStrRev(LEFT(strTemp,cnum), ",")

		If cp1>cp Then cp=cp1
		If cp2>cp Then cp=cp2
		If cp3>cp Then cp=cp3
	
	   If cp>0 Then
		strTemp=LEFT(strTemp, cp-1) 
	   Else
		strTemp=LEFT(strTemp, cnum) 
	   End If	
	End If
	CutStringNDeleteAfterSpace = strTemp
End Function 


Function GetMemberTrialValidation()
Dim bRequestValid
	objTempRs2=GetDataOutParamsSP("usp_AdmUrsValidityCheck", Array( _
		Array(, adInteger, , iMemberID), _
		Array(, adInteger, , iExpertID), _
		Array(, adVarChar, 255, sUserIpAddress)), _
		Array( Array(, adInteger)))
	bRequestValid=objTempRs2(0)

	If iMemberID=0 And iExpertID=0 Then
		If bRequestValid<30 Then
			bRequestValid=1
		Else
			bRequestValid=0
		End If
	Else
		If bRequestValid<500 Or Left(sUserIpAddress,10)="158.29.157" Then
			bRequestValid=1
		Else
			bRequestValid=0
			iUserID=0
			iMemberID=0
			iExpertID=0
			sSessionID=""
			sCookiesSessionID=""
			Response.Cookies("SessionID")=""
		End If
	End If
	Set objTempRs2=Nothing
GetMemberTrialValidation=bRequestValid
End Function


Function ShowActLngField(Field,Lng,MasterLng)
Dim tmpField
 If Eval(Field & Lng)>"" Then
	tmpField=Eval(Field & Lng)
 ElseIf  Eval(Field & MasterLng)>"" Then
	tmpField="[" & LCase(Left(MasterLng,2)) & "] " & Eval(Field & MasterLng)
 End If
ShowActLngField = tmpField 
End function   
'''''''''''''''''''''''''''
''''''''''''''''''''''''''
Function ShowActLngName(Field,Lng,MasterLng)
Dim tmpField
 If Eval(Field & Lng)>"" Then
	tmpField=Eval(Field & Lng)
 ElseIf  Eval(Field & MasterLng)>"" Then
	tmpField=Eval(Field & MasterLng)
 End If
ShowActLngName = tmpField 
End function   

Function SaveActLngField(Field)
 If Left(Field,1)="[" and Mid(Field,4,1)="]" Then
	Field=Mid(Field,6,Len(Field))
 End If
SaveActLngField = Field 
End function   

Function InsertRecord(TableName,Parameters,ParameterValues)
		strSQL = "Insert into "& TableName &" ("&Parameters&") VALUES (" & ParameterValues & " )" 
   '    	Response.Write strSQL
   objconn.Execute(strSQL)
End Function

Function ShowEntityPlural(iEntitiesNumber, sEntitySingular, sEntityPlural, sDelimiter)
Dim sTemp
	If Abs(iEntitiesNumber)>1 Then sTemp=sEntityPlural Else sTemp=sEntitySingular
ShowEntityPlural=iEntitiesNumber & sDelimiter & sTemp
End Function


Function HighlightHttpLinks(AText)
Dim sResult

	Dim iPositionLinkStart, iPositionLinkEnd, iLinkLength, sLink, iPositionLinkPreviousEnd
	iPositionLinkStart=InStr(1, AText, "http", 1)
	iPositionLinkPreviousEnd=0
	
	If iPositionLinkStart=0 Then
		sResult=AText
	Else
		While iPositionLinkStart>0
			iPositionLinkEnd=iPositionLinkStart + FindHttpLinkEnd(Mid(AText, iPositionLinkStart, 512) & " ")
			If iPositionLinkStart>0 And iPositionLinkEnd>iPositionLinkStart Then
				iLinkLength=iPositionLinkEnd-iPositionLinkStart
				sLink=Mid(AText, iPositionLinkStart, iLinkLength)
				
				sResult=sResult & Mid(AText, iPositionLinkPreviousEnd+1, iPositionLinkStart-iPositionLinkPreviousEnd-1) + "<a href=""" & sLink & """ target=""blank"">" & sLink & "</a>"
				iPositionLinkPreviousEnd=iPositionLinkEnd
			End If
			
			iPositionLinkStart=InStr(iPositionLinkEnd+1, AText, "http", 1)
		WEnd
		If iPositionLinkPreviousEnd>0 And Len(AText)>iPositionLinkPreviousEnd+1 Then
			sResult=sResult & Mid(AText, iPositionLinkPreviousEnd, Len(AText)-iPositionLinkPreviousEnd+1)
		End if
	End If

HighlightHttpLinks=sResult
End Function

Function FindHttpLinkEnd(AText)
Dim iResult
iResult=0

Dim iLoop, iCharCode
	For iLoop=1 To Len(AText)
		iCharCode=Asc(Mid(AText, iLoop, 1))
		' 38/26 &, 58/3A :, 47/2F /,
		
		If Not (iCharCode=37 Or iCharCode=38 Or iCharCode=45 Or iCharCode=46 Or iCharCode=47 Or iCharCode=58 Or iCharCode=59 Or iCharCode=61 Or iCharCode=63 Or iCharCode=64 Or iCharCode=95 _
				Or (iCharCode>=48 And iCharCode<=57) Or (iCharCode>=65 And iCharCode<=90) Or (iCharCode>=97 And iCharCode<=122)) Then
			iResult=iLoop
			Exit For
		End If
	Next
	
FindHttpLinkEnd=iResult
End Function

Function ReplaceHtmlSpecialCodes(AText, AReplaceWith)
Dim sResult

	Dim iPositionCodeStart, iPositionCodeEnd, iCodeLength, sCode, iPositionCodePreviousEnd
	iPositionCodeStart=InStr(1, AText, "&#", 1)
	iPositionCodePreviousEnd=0
	
	If iPositionCodeStart=0 Then
		sResult=AText
	Else
		While iPositionCodeStart>0
			iPositionCodeEnd=iPositionCodeStart + FindHtmlCodeEnd(Mid(AText, iPositionCodeStart, 10) & " ")
			If iPositionCodeStart>0 And iPositionCodeEnd>iPositionCodeStart Then
				iCodeLength=iPositionCodeEnd-iPositionCodeStart
				sCode=Mid(AText, iPositionCodeStart, iCodeLength)
				
				sResult=sResult & Mid(AText, iPositionCodePreviousEnd+1, iPositionCodeStart-iPositionCodePreviousEnd-1) & AReplaceWith
				iPositionCodePreviousEnd=iPositionCodeEnd-1
			End If
			
			iPositionCodeStart=InStr(iPositionCodeEnd+1, AText, "&#", 1)
		WEnd
		If iPositionCodePreviousEnd>0 And Len(AText)>iPositionCodePreviousEnd+1 Then
			sResult=sResult & Mid(AText, iPositionCodePreviousEnd, Len(AText)-iPositionCodePreviousEnd+1)
		End if
	End If

ReplaceHtmlSpecialCodes=sResult
End Function

Function FindHtmlCodeEnd(AText)
Dim iResult
iResult=0

Dim iLoop, iCharCode
	For iLoop=1 To Len(AText)
		iCharCode=Asc(Mid(AText, iLoop, 1))
		
		If Not (iCharCode=35 Or iCharCode=38 Or iCharCode=59 Or _
			(iCharCode>=48 And iCharCode<=57)) Then
			iResult=iLoop
			Exit For
		End If
	Next
	
FindHtmlCodeEnd=iResult
End Function


Function CreateGroupedListString(ARecordset, AGroupFieldName, AInfoFieldName, AGroupDelimiterStart, AGroupDelimiterEnd, AGroupInfoDelimiter, AInfoDelimiterStart, AInfoDelimiterEnd)
	Dim sResult
	sResult = ""

	Dim sGroupTemp
	sGroupTemp = ""
	If IsObject(ARecordset) Then 
		While Not ARecordset.Eof
			Dim sFieldValue
			sFieldValue = ARecordset(AInfoFieldName)
			On Error Resume Next
				' If field value is empty or null take the english version
				sFieldValue = ReplaceIfEmpty(sFieldValue, ARecordset(Left(AInfoFieldName, Len(AInfoFieldName)-3) & "Eng"))
			On Error GoTo 0

			If sGroupTemp<>Trim(ARecordset(AGroupFieldName)) Then
				If sGroupTemp<>"" Then 
					sResult = sResult & AGroupDelimiterEnd
				End If
				sResult = sResult & AGroupDelimiterStart & ARecordset(AGroupFieldName) & AGroupInfoDelimiter & AInfoDelimiterStart & sFieldValue 'ARecordset(AInfoFieldName)
				sGroupTemp = Trim(ARecordset(AGroupFieldName))
			Else
				sResult = sResult & AInfoDelimiterEnd & AInfoDelimiterStart & sFieldValue 'ARecordset(AInfoFieldName)
			End If

			ARecordset.MoveNext
		WEnd
	End If

	CreateGroupedListString = sResult
End Function


Function CreateListString(ARecordset, AInfoFieldName, AInfoDelimiterStart, AInfoDelimiterEnd)
	Dim sResult
	sResult = ""

	If IsObject(ARecordset) Then 
		While Not ARecordset.Eof
			Dim sFieldValue
			sFieldValue = ARecordset(AInfoFieldName)
			On Error Resume Next
				' If field value is empty or null take the english version
				sFieldValue = ReplaceIfEmpty(sFieldValue, ARecordset(Left(AInfoFieldName, Len(AInfoFieldName)-3) & "Eng"))
			On Error GoTo 0

			sResult = sResult & AInfoDelimiterStart & sFieldValue

			ARecordset.MoveNext
			If Not ARecordset.Eof Then 
				sResult = sResult & AInfoDelimiterEnd
			End If
		WEnd
	End If

	CreateListString = sResult
End Function

Function CheckGuidAndNull(sTextInput)
	If IsNull(sTextInput) Then
		CheckGuidAndNull = Null
		Exit Function
	End If
	
	Dim regEx
	Set regEx = New RegExp
	regEx.Pattern = "^({|\()?[A-Fa-f0-9]{8}-([A-Fa-f0-9]{4}-){3}[A-Fa-f0-9]{12}(}|\))?$"
	If regEx.Test(sTextInput) Then
		CheckGuidAndNull = UCase(sTextInput)
	Else
		CheckGuidAndNull = Null
	End If
	Set RegEx = Nothing
End Function

Function IsGuid(sTextInput)
	If IsNull(sTextInput) Then
		IsGuid = False
		Exit Function
	End If
	
	Dim regEx
	Set regEx = New RegExp
	regEx.Pattern = "^({|\()?[A-Fa-f0-9]{8}-([A-Fa-f0-9]{4}-){3}[A-Fa-f0-9]{12}(}|\))?$"
	If regEx.Test(sTextInput) Then
		IsGuid = True
	Else
		IsGuid = False
	End If
	Set RegEx = Nothing
End Function

Function CutStringNSplit(AInput, AMaxLineLength, ANewLineStart)
Dim strTemp, cp
	strTemp = AInput
	If Len(strTemp) > AMaxLineLength Then
		cp = InStrRev(Left(strTemp, AMaxLineLength), "/")
		If cp > 0 Then
			strTemp = Left(strTemp, cp) & ANewLineStart & Mid(strTemp, cp + 1, Len(strTemp))
		End If
	End If
	CutStringNSplit = strTemp
End Function
%>
