<%
If sApplicationName<>"expert" Then
	CheckUserLogin sScriptFullNameAsParams
End If
%>
<!--#include virtual="/_common/_class/document.asp"-->
<!--#include virtual="/_common/_grid/document_list.asp"-->
<%
Dim bCvValidForMemberOrExpert
Dim sCVFormat, dflag

' Check is the CV valid for member / expert
bCvValidForMemberOrExpert=1

' On changing a CV format do redirect
sCvFormat=Request.QueryString("act")
If sCvFormat="ASR" Or sCvFormat="ADB" Or sCvFormat="AFB" Or sCvFormat="EC" Or sCvFormat="EP" Or sCvFormat="WB" Or sCvFormat="TRL" Or sCvFormat="MRL" Then
	If sHomePath="/tripleline/" And sCvFormat="" Then sCvFormat="TRL"
	If sHomePath="/merlin/" And sCvFormat="" Then sCvFormat="MRL"
	If sCvFormat="ASR" Then sCvFormat=""
	sUrl="cv_view" & sCvFormat & ".asp" & ReplaceUrlParams(sParams, "act")
	Response.Redirect(sUrl)
End If

If InStr(sScriptFileName, "adb")>0 Then
	sCVFormat="ADB"
ElseIf InStr(sScriptFileName, "afb")>0 Then
	sCVFormat="AFB"
ElseIf InStr(sScriptFileName, "ec")>0 Then
	sCVFormat="EC"
ElseIf InStr(sScriptFileName, "ep")>0 Then
	sCVFormat="EP"
ElseIf InStr(sScriptFileName, "wb")>0 Then
	sCVFormat="WB"
ElseIf InStr(sScriptFileName, "trl")>0 Then
	sCVFormat="TRL"
ElseIf InStr(sScriptFileName, "mrl")>0 Then
	sCVFormat="MRL"
Else
	sCVFormat=""
End If

Dim iProjectID
iProjectID=CheckIntegerAndZero(Request.QueryString("idproject"))
If iProjectID>0 Then
	sParams=ReplaceUrlParams(sParams, "idproject=" & iProjectID)
End If
sParams=ReplaceUrlParams(sParams, "t")

'Dim sActiveLng, sMasterLng, k, arrRowsValues
Dim sFileType, sFileName, k, arrRowsValues

'sActiveLng="Eng"

Dim sFirstNameEng, sFirstNameFra, sFirstNameSpa, sFirstName
Dim sMiddleNameEng, sMiddleNameFra, sMiddleNameSpa, sMiddleName
Dim sLastNameEng, sLastNameFra, sLastNameSpa, sLastName, sTempLastName
Dim iTitleID, sTitle, sFullName, sTitleLastName, sFullNameWithSpaces, iGender, sBirthDate, sBirthPlace, iMaritalStatus, iPersonID
Dim sNationality, sOtherLanguages, sTempLanguage
Dim sDonors, sCountries, sSectors
Dim sPermAddress, sPermAddressStreet, sPermAddressPostcode, sPermAddressCity, sPermAddressCountry, sPermAddressPhone, sPermAddressMobile, sPermAddressFax, sPermAddressEmail, sPermAddressWeb, bPermAddress
Dim sCurAddress, sCurAddressStreet, sCurAddressPostcode, sCurAddressCity, sCurAddressCountry, sCurAddressPhone, sCurAddressMobile, sCurAddressFax, sCurAddressEmail, sCurAddressWeb, bCurAddress
Dim objRsExpEdu, objRsExpTrn, objRsExpWke, objRsExpLng, objRsExpLngOther, objRsExpLngNative, objRsExpCou
Dim sEduSubject, sDescription
Dim sComments

Dim iProfYears, sProfession, sMemberships, sOtherSkills, sKeyQualification, sPosition, sPublications
Dim iProfessionalStatusID, sProfessionalStatus
Dim sReferences, sAvailability, bShortterm, bLongterm, sAchievements
Dim sUserPhone, sPreferences

Dim arrStartDateValues, arrEndDateValues, arrPrjTitleValues
Dim bEmailToConfirmCVSent, bEmailToCompleteCVSent, bCVApproved, bCVIbfOnly, bCVHidden, bCVDeleted, bCVRemoved
Dim bEmailExpertAccountSent

If iCvID>0 Then 
Set objTempRs=GetDataRecordsetSP("usp_ExpCvvExpInfoSelect", Array( _
	Array(, adInteger, , iCvID)))

If Not objTempRs.Eof Then
	sCvLanguage = ReplaceIfEmpty(sForceCvLanguage, ReplaceIfEmpty(objTempRs("Lng"), sDefaultCvLanguage))
End If
%>
<!--#include file="_data/datGender.asp"-->
<!--#include file="_data/datLanguage.asp"-->
<!--#include file="_data/datLngLevel.asp"-->
<!--#include file="_data/datPsnTitle.asp"-->
<!--#include file="_data/datPsnStatus.asp"-->
<!--#include file="_data/datEduSubject.asp"-->
<!--#include file="_data/datEduType.asp"-->
<!--#include file="_data/datMonth.asp"-->
<!--#include file="_data/en/datProfessionalStatus.asp"-->
<%
If Not objTempRs.Eof Then
	On Error Resume Next
	sCvLanguage = ReplaceIfEmpty(objTempRs("Lng"), sDefaultCvLanguage)
	ForceCvLanguage()

	iProfYears=objTempRs("expProfYears")
	sProfession=objTempRs("expProfessionEng")
	iProfessionalStatusID=objTempRs("id_ProfessionalStatus")

	If IsArray(arrProfessionalStatusID) Then
	For i=LBound(arrProfessionalStatusID) to UBound(arrProfessionalStatusID) 
		If iProfessionalStatusID=arrProfessionalStatusID(i) Then
			sProfessionalStatus=arrProfessionalStatusTitle(i)
		End If
	Next 
	End If
	
	sOtherSkills=ConvertText(objTempRs("expOtherSkills"))
	sMemberships=ConvertText(objTempRs("expMemberProfEng"))
	sKeyQualification=ConvertText(objTempRs("expKeyQualificationsEng"))
	sPosition=objTempRs("expCurrPositionEng")
	sPublications=ConvertText(objTempRs("expPublicationsEng"))
	sReferences=ConvertText(objTempRs("expReferencesEng"))
	sAvailability=ConvertText(objTempRs("expAvailabilityEng"))
	bShortterm=objTempRs("expShortterm")
	bLongterm=objTempRs("expLongterm")
	sUserLanguage=objTempRs("Lng")
	sUserEmail=objTempRs("Email")
	sUserPhone=objTempRs("Phone")
	bEmailToConfirmCVSent=objTempRs("expToConfirmCvEmailSent")
	bEmailToCompleteCVSent=objTempRs("expToCompleteCvEmailSent")
	bCVApproved=objTempRs("expApproved")
	bCVHidden=objTempRs("expHidden")
	bCVDeleted=objTempRs("expDeleted")
	bCVRemoved=objTempRs("expRemoved")

	sPreferences=""
	If bShortterm Then 
		sPreferences=sPreferences & "Short-term missions, "
	End If
	If bLongterm Then 
		sPreferences=sPreferences & "Long-term missions, "
	End If
	If Len(sPreferences)>2 Then
		sPreferences=Left(sPreferences,Len(sPreferences)-2)
	End If
	

	If bCvValidForMemberOrExpert=1 Then
	sFirstNameEng=objTempRs("psnFirstNameEng")
	sFirstName=sFirstNameEng
	If sFirstName>"" Then sFirstName=sFirstName & " "

	sMiddleNameEng=objTempRs("psnMiddleNameEng")
	sMiddleName=sMiddleNameEng
	If sMiddleName>"" Then sMiddleName=sMiddleName & " "

	sLastNameEng=objTempRs("psnLastNameEng")
	sLastName=sLastNameEng
	If sLastName>"" Then sLastName=sLastName & " "

	iTitleID=objTempRs("id_psnTitle")

	If iTitleID>"" And IsNumeric(iTitleID) Then sTitle=arrPersonTitle(iTitleID) & " "

	sFullName=sTitle & sFirstName & sMiddleName & sLastName
	sTitleLastName=sTitle & sLastName
	If Len(sFullName) > 80 Then
		sFullName=sTitle & sFirstName & sLastName
	End If		
	If Len(sFullName) < 46 Then
		sFullNameWithSpaces=sFullName+Space(46-Len(sFullName))
	Else
		sFullNameWithSpaces=sFullName+Space(5)
	End If		
	If InStr(sLastName, "&")>0 Then 
		sTempLastName=Left(sLastName, InStr(sLastName, "&")-1) 
	Else 
		sTempLastName=sLastName
	End If
	sFileName = "CV_" & Replace(sTempLastName," ","") & "_" & Replace(sFirstName," ","") & "_" & sCvFormat & ConvertDateForText(Now(), "", "DDMMYYYY")
	sFileName = Replace(sFileName, ",", "_")
	sFileName = Replace(sFileName, ":", "_")

	sBirthDate = objTempRs("psnBirthDate")
	sBirthPlace = objTempRs("psnBirthPlaceEng")
	End If
	iGender=objTempRs("psnGender")
	iMaritalStatus=objTempRs("id_MaritalStatus")

	sComments=objTempRs("expComments")
	
	iPersonID=objTempRs("id_Person")
	bEmailExpertAccountSent = ReplaceIfEmpty(objTempRs("expAccountEmailSent"), 0)
	On Error GoTo 0
objTempRs.Close   
End If

sNationality=GetExpNationalities(iCvID, sCvLanguage)

Set objRsExpEdu=GetDataRecordsetSP("usp_ExpCvvEducationSelect", Array( _
	Array(, adInteger, , iCvID), Array(, adInteger, , 1)))

Set objRsExpTrn=GetDataRecordsetSP("usp_ExpCvvEducationSelect", Array( _
	Array(, adInteger, , iCvID), Array(, adInteger, , 2)))

Set objRsExpWke=GetDataRecordsetSP("usp_ExpCvvExperienceSelect", Array( _
	Array(, adInteger, , iCvID)))


objRsExpLng=GetDataOutParamsSP("usp_GetExpertProfDetails", Array( _
	Array(, adInteger, , iCvID), Array(, adInteger, , 0), Array(, adVarChar, 3, "Eng"), Array(, adInteger, , 0)), Array( _
	Array(, adVarChar, 1), Array(, adVarWChar, 400), Array(, adVarWChar, 50), Array(, adVarWChar, 255), Array(, adVarWChar, 1000), Array(, adVarWChar, 500), Array(, adVarWChar, 1000), Array(, adVarWChar, 500), Array(, adVarWChar, 1000), Array(, adVarWChar, 2000)))
sOtherLanguages=objRsExpLng(5)
Set objRsExpLng=Nothing


Set objRsExpLngNative=GetDataRecordsetSP("usp_ExpCvvLanguageSelect", Array( _
	Array(, adInteger, , iCvID), Array(, adVarChar, 10, "native")))

Set objRsExpLngOther=GetDataRecordsetSP("usp_ExpCvvLanguageSelect", Array( _
	Array(, adInteger, , iCvID), Array(, adVarChar, 10, "other")))


If bCvValidForMemberOrExpert=1 Then

' Current address
bCurAddress=False
Set objTempRs=GetDataRecordsetSP("usp_ExpCvvAddressSelect", Array( _
	Array(, adInteger, , iCvID), Array(, adInteger, , 3)))
If Not objTempRs.Eof Then 
	sCurAddressStreet=objTempRs("adrStreetEng")
	sCurAddressPostcode=objTempRs("adrPostCode")
	sCurAddressCity=objTempRs("adrCityEng")
	sCurAddressCountry=objTempRs("couNameEng")
	sCurAddressPhone=objTempRs("adrPhone")
	sCurAddressMobile=objTempRs("adrMobile")
	sCurAddressFax=objTempRs("adrFax")
	sCurAddressEmail=objTempRs("adrEmail")
	sCurAddressWeb=objTempRs("adrWeb")

	If CheckLength(sCurAddressStreet) + CheckLength(sCurAddressPostcode) + CheckLength(sCurAddressCity) + CheckLength(sCurAddressCountry) + CheckLength(sCurAddressPhone) + CheckLength(sCurAddressMobile) + CheckLength(sCurAddressFax) + CheckLength(sCurAddressEmail) + CheckLength(sCurAddressWeb)>0 Then
		bCurAddress=True
	End If
End If
objTempRs.Close

' Permanent address
sPermAddress=""
bPermAddress=False
Set objTempRs=GetDataRecordsetSP("usp_ExpCvvAddressSelect", Array( _
	Array(, adInteger, , iCvID), Array(, adInteger, , 1)))
If Not objTempRs.Eof Then 
	sPermAddressStreet=objTempRs("adrStreetEng")
	sPermAddressPostcode=objTempRs("adrPostCode")
	sPermAddressCity=objTempRs("adrCityEng")
	sPermAddressCountry=objTempRs("couNameEng")
	sPermAddressPhone=objTempRs("adrPhone")
	sPermAddressMobile=objTempRs("adrMobile")
	sPermAddressFax=objTempRs("adrFax")
	sPermAddressEmail=objTempRs("adrEmail")
	sPermAddressWeb=objTempRs("adrWeb")

	If CheckLength(sPermAddressStreet) + CheckLength(sPermAddressPostcode) + CheckLength(sPermAddressCity) + CheckLength(sPermAddressCountry) + CheckLength(sPermAddressPhone) + CheckLength(sPermAddressMobile) + CheckLength(sPermAddressFax) + CheckLength(sPermAddressEmail) + CheckLength(sPermAddressWeb)>0 Then
		bPermAddress=True
	End If

	sPermAddress="<p class=""txt"">" & sPermAddressStreet & " " & sPermAddressPostcode & " " & sPermAddressCity & "</p>" 
	If sPermAddressCountry>"" Then
		sPermAddress=sPermAddress & "<p class=""txt"">" & sPermAddressCountry & "</p>"
	End If
	sPermAddressPhone=ReplaceIfEmpty(sPermAddressPhone, sCurAddressPhone)
	If sPermAddressPhone>"" Then
		sPermAddress=sPermAddress & "<p class=""txt"">Phone: " & sPermAddressPhone & "</p>"
	End If
	If sPermAddressFax>"" Then
		sPermAddress=sPermAddress & "<p class=""txt"">Fax: " & sPermAddressFax & "</p>"
	End If
	sPermAddressEmail=ReplaceIfEmpty(sPermAddressEmail, sCurAddressEmail)
	If sPermAddressEmail>"" Then
		sPermAddress=sPermAddress & "<p class=""txt"">E-mail: " & sPermAddressEmail & "</p>"
	End If
End If
objTempRs.Close


If CheckLength(sCurAddressEmail)=0 Then
	If (Not (InStr(sPermAddressEmail, sUserEmail)>0 Or InStr(sUserEmail, sPermAddressEmail)>0)) Or CheckLength(sPermAddressEmail)=0 Then
		sCurAddressEmail=sUserEmail
	End If
	If Len(sCurAddressPhone)<5 And (Not (InStr(sPermAddressPhone, sUserPhone)>0 Or InStr(sUserPhone, sPermAddressPhone)>0)) Or CheckLength(sPermAddressPhone)=0 Then
		sCurAddressPhone=sUserPhone
	End If
End If

If bCurAddress=False And bPermAddress=False And (CheckLength(sUserEmail)>0 Or CheckLength(sUserPhone)>0) Then
	bCurAddress=True
End If


End If
End If

Function SetECLanguageLevel(iLevel)
Dim iECLevel
  If iLevel>"" And IsNumeric(iLevel) Then
	iECLevel=iLevel
  Else
	iECLevel=""
  End If
SetECLanguageLevel=iECLevel
End Function

Function SetEPLanguageLevel(iLevel)
Dim sEPLevel
  If iLevel>"" And IsNumeric(iLevel) Then
	Select Case iLevel
		Case 5
		sEPLevel="C2"
		Case 4
		sEPLevel="C1"
		Case 3
		sEPLevel="B2"
		Case 2
		sEPLevel="B1"
		Case 1
		sEPLevel="A2"
		Case else
		sEPLevel=""
	End Select  
  Else
	sEPLevel=""
  End If
SetEPLanguageLevel=sEPLevel
End Function


Sub ShowExpCVFeatureBox
Response.Write "<img src=""" & sHomePath & "image/x.gif"" width=1 height=5><br><br>"
If bCvValidForMemberOrExpert=1 And sScriptFileName<>"cv_preview.asp" Then %>

	<form name="cvformat" method="get" action="<%=sApplicationHomePath & "view/cv_view.asp" %>">
	<input type="hidden" name="id" value="<% =iCvID %>">
	<input type="hidden" name="idproject" value="<% =iProjectID %>">
	<% ShowFeatureBoxHeader "Format the CV" %>
	<p class="sml" align="center">Select a format to get the CV in&nbsp;</p>
	<img src="<% =sHomePath %>image/x.gif" width=1 height=3><br>
	<div align="center"><select style="font-face: Arial; font-size:8.5pt;" name="act" size=1>
	<%
	If sHomePath="/tripleline/" Then
	%>
	<option value="TRL" <% If sCVFormat="TRL" Then %>selected<% End If %>>Tripleline</option>
	<option value="ASR" <% If sCVFormat="" Then %>selected<% End If %>>assortis.com</option>
	<option value="ADB" <% If sCVFormat="ADB" Then %>selected<% End If %>>Asian Development Bank</option>
	<option value="AFB" <% If sCVFormat="AFB" Then %>selected<% End If %>>African Development Bank</option>
	<option value="EC" <% If sCVFormat="EC" Then %>selected<% End If %>>European Commission</option>
	<option value="EP" <% If sCVFormat="EP" Then %>selected<% End If %>>Europass</option>
	<option value="WB" <% If sCVFormat="WB" Then %>selected<% End If %>>World Bank</option>
	<%
	ElseIf sHomePath="/merlin/" Then
	%>
	<option value="MRL" <% If sCVFormat="MRL" Then %>selected<% End If %>>Cabinet Merlin</option>
	<option value="ASR" <% If sCVFormat="" Then %>selected<% End If %>>assortis.com</option>
	<option value="ADB" <% If sCVFormat="ADB" Then %>selected<% End If %>>Asian Development Bank</option>
	<option value="AFB" <% If sCVFormat="AFB" Then %>selected<% End If %>>African Development Bank</option>
	<option value="EC" <% If sCVFormat="EC" Then %>selected<% End If %>>European Commission</option>
	<option value="EP" <% If sCVFormat="EP" Then %>selected<% End If %>>Europass</option>
	<option value="WB" <% If sCVFormat="WB" Then %>selected<% End If %>>World Bank</option>
	<%
	ElseIf sCvLanguage = cLanguageFrench Then
	%>
	<option value="ASR" <% If sCVFormat="" Then %>selected<% End If %>>assortis.com</option>
	<option value="AFB" <% If sCVFormat="AFB" Then %>selected<% End If %>>African Development Bank</option>
	<option value="EC" <% If sCVFormat="EC" Then %>selected<% End If %>>European Commission</option>
	<option value="EP" <% If sCVFormat="EP" Then %>selected<% End If %>>Europass</option>
	<option value="WB" <% If sCVFormat="WB" Then %>selected<% End If %>>World Bank</option>
	<%
	Else
	%>
	<option value="ASR" <% If sCVFormat="" Then %>selected<% End If %>>assortis.com</option>
	<option value="ADB" <% If sCVFormat="ADB" Then %>selected<% End If %>>Asian Development Bank</option>
	<option value="AFB" <% If sCVFormat="AFB" Then %>selected<% End If %>>African Development Bank</option>
	<option value="EC" <% If sCVFormat="EC" Then %>selected<% End If %>>European Commission</option>
	<option value="EP" <% If sCVFormat="EP" Then %>selected<% End If %>>Europass</option>
	<option value="WB" <% If sCVFormat="WB" Then %>selected<% End If %>>World Bank</option>
	<%
	End If
	%>
	</select></div>
	<img src="<% =sHomePath %>image/x.gif" width=1 height=6><br>
	<div align="center"><input type="image" src="<% =sHomePath %>image/bte_select.gif" alt="Select" height=18 vspace=0 border=0></a></div>
	<% ShowFeatureBoxFooter %>
	</form>

<% If InStr(sScriptFileName, "register")=0 Then %>
	<% ShowFeatureBoxHeader("Document options") %>
	<p class="sml"><a href="<%=Replace(sScriptFileName, "view", "save")%>?id=<%=iCvID%>&ftype=doc"><img src="<% =sHomePath %>image/file_doc.gif" width=18 height=17 border=0 hspace=4 align="left">Save&nbsp;as&nbsp;Word&nbsp;document</a></p>
	<img src="<% =sHomePath %>image/x.gif" width=1 height=3><br><img src="<% =sHomePath %>image/x.gif" width=31 height=1 hspace=0 vspace=25 align="left">
	<p class="sml">This option will save the CV in Microsoft&reg; Word* format on your local computer.</p>
	<img src="<% =sHomePath %>image/x.gif" width=1 height=8><br>
	<p class="sml"><a href="<%=Replace(sScriptFileName, "view", "save")%>?id=<%=iCvID%>&ftype=prn"><img src="<% =sHomePath %>image/file_prn.gif" width=18 height=17 border=0 hspace=4 align="left">Print&nbsp;this&nbsp;document</a></p>
	<img src="<% =sHomePath %>image/x.gif" width=1 height=3><br><img src="<% =sHomePath %>image/x.gif" width=31 height=1 hspace=0 vspace=25 align="left">
	<p class="sml">This option will open the CV in Microsoft&reg; Word*  and you should click on Print button there.</p>
	<% ShowFeatureBoxFooter %>
	<br>
<% End If %>

<% If sUserType="backoffice" And  _
bCvDocumentActive = cCvDocumentEnabled Then %>
	<% ShowFeatureBoxHeader("Supporting documents") %>
	<p class="sml" style="padding: 2px 5px;">
	<%
	Dim objDocumentList
	Set objDocumentList = New CDocumentList
	objDocumentList.LoadDocumentListByExpertID iCvID, ""
	If objDocumentList.Count=0 Then
	%>
		<p class="sml" style="padding: 2px 5px;">There are no documents uploaded for this expert yet.</p>
	<% 
	Else
		ShowDocumentListViewTable objDocumentList
	End If
	Set objDocumentList = Nothing
	%>
	</p>
	<% ShowFeatureBoxFooterWithFormFooter %>
	<br>	
<% End If %>

<% If sApplicationName="backoffice" Then %>
<% On Error Resume Next %>
	<% ShowFeatureBoxHeader("Place on project") %>
	<p class="sml" style="padding: 2px 5px;">Expert is suitable for a project?</p>
	<div align="center"><a href="../project/link_expert.asp<% =ReplaceUrlParams(ReplaceUrlParams(sParams, "id"), "idexpert=" & iCvID) %>"><img src="<% =sHomePath %>image/bte_addexpert152.gif" vspace="4" border="0"></a></div>
	<% ShowFeatureBoxFooter %>
	<br>
	
	<% ShowFeatureBoxHeader("Manage CV") %>
	<p class="sml" style="padding: 2px 5px;">To update or delete this CV, modify status or register any comments</p>
	<div align="center"><a href="<% =sApplicationHomePath %>register/register6.asp<% =ReplaceUrlParams(ReplaceUrlParams(sParams, "idexpert"), "id=" & iCvID) %>"><img src="<% =sHomePath %>image/bte_cvmanagement152.gif" vspace="4" border="0"></a></div>
	<% ShowFeatureBoxFooter %>
	<br>
<% On Error GoTo 0 %>
<% ElseIf sApplicationName="external" Then %>
	<% ShowFeatureBoxHeader("Manage CV") %>
	<p class="sml" style="padding: 2px 5px;">To update or delete this CV, modify status or register any comments</p>
	<div align="center"><a href="<% =sApplicationHomePath %>register/register6.asp<% =ReplaceUrlParams(ReplaceUrlParams(sParams, "idexpert"), "id=" & iCvID) %>"><img src="<% =sHomePath %>image/bte_cvmanagement152.gif" vspace="4" border="0"></a></div>
	<% ShowFeatureBoxFooter %>
	<br>
<% End If%>

<%
End If
End Sub
%>


<%
	If sApplicationName="external" Then
		If sContactDetailsExternally=cNameObfuscated Then
			sFirstName=ObfuscateString(sFirstName)
			sLastName=ObfuscateString(sLastName)
			sMiddleName=ObfuscateString(sMiddleName)
			sPermAddressEmail=ObfuscateEmail(sPermAddressEmail)
			sPermAddressPhone=ObfuscateString(sPermAddressPhone)
			sCurAddressEmail=ObfuscateEmail(sCurAddressEmail)
			sCurAddressPhone=ObfuscateString(sPermAddressPhone)
			
		End If
		If sContactDetailsExternally=cNameHidden Then
			sFirstName=""
			sLastName=""
			sMiddleName=""
		End If
	End If
%>