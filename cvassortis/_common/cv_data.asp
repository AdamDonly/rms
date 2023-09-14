<%
If sApplicationName<>"expert" Then
	CheckUserLogin sScriptFullNameAsParams
End If
%>
<!--#include file="_class/document.asp"-->
<!--#include file="_grid/document_list.asp"-->
<%
Dim iCvID, sCvUID, bCvValidForMemberOrExpert
Dim sCVFormat, dflag

sCvUID=Request.QueryString("uid")
On Error Resume Next
	Set objTempRs=GetDataRecordsetSP("usp_Ica_ExpertIdSelect", Array( _
		Array(, adVarChar, 40, sCvUID)))

	If Err.Number<>0 Or objTempRs.Eof Then
		' Redirect top experts to a different page, if their CV ID is not found
		Response.Redirect "/"
	End If

	iCvID=objTempRs("id_Expert")
	Set objExpertDB = objExpertDBList.Find(objTempRs("id_Database"), "ID")
	
Set objTempRs=Nothing
On Error GoTo 0

Dim objConnCustom
Set objConnCustom = Server.CreateObject("ADODB.Connection")
objConnCustom.Open "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Data Source=MAGNESA;Initial Catalog=" & objExpertDB.DatabasePath & ";"

' Redirect to original CV if expert's CV was updated with another ID (Blacklist=1 & id_ExpertOriginal>0)
objTempRs=GetDataOutParamsSPWithConn(objConnCustom, "usp_ExpCvvOriginalSelect", Array( _
	Array(, adInteger, , iCvID)), _
	Array( Array(, adInteger)))
iExpertOriginalID=objTempRs(0)
If iExpertOriginalID>0 Then Response.Redirect(Replace(sScriptFileName, "cv_preview", "cv_view") & ReplaceUrlParams(sParams, "id=" & iExpertOriginalID))


' Check is the CV valid for member / expert
bCvValidForMemberOrExpert=IsIcaUserCompanyCvValid(objExpertDB.Database, iCvID, objUserCompanyDB.Database)

' On changing a CV format do redirect
sCvFormat=Request.QueryString("act")
If sCvFormat="ASR" Or sCvFormat="ADB" Or sCvFormat="AFB" Or sCvFormat="EC" Or sCvFormat="EP" Or sCvFormat="WB" Then
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
Dim sPhone, sEmail
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

Dim bCvAccessValid
bCvAccessValid=0

If iCvID>0 Then 

Set objTempRs=GetDataRecordsetSPWithConn(objConnCustom, "usp_ExpertSelect", Array( _
	Array(, adInteger, , objExpertDB.ID), _
	Array(, adInteger, , iCvID)))

If Not objTempRs.Eof Then
	sCvLanguage = ReplaceIfEmpty(objTempRs("Lng"), sDefaultCvLanguage)
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
	sEmail=objTempRs("Email")
	sPhone=objTempRs("Phone")
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
	sFileName="CV_" & Replace(sTempLastName," ","") & "_" & Replace(sFirstName," ","") & "_" & sCvFormat & ConvertDateForText(Now(), "", "DDMMYYYY")

	sBirthDate = objTempRs("psnBirthDate")
	sBirthPlace = objTempRs("psnBirthPlaceEng")

	iGender=objTempRs("psnGender")
	iMaritalStatus=objTempRs("id_MaritalStatus")

	sComments=objTempRs("expComments")
	
	iPersonID=objTempRs("id_Person")
objTempRs.Close 
On Error GoTo 0
End If

sNationality=GetExpNationalities(sCvLanguage, objExpertDB.Database, iCvID)

Set objRsExpEdu=GetDataRecordsetSPWithConn(objConnCustom, "usp_ExpertEducationSelect", Array( _
	Array(, adInteger, , iCvID), _
	Array(, adInteger, , 1)))

Set objRsExpTrn=GetDataRecordsetSPWithConn(objConnCustom, "usp_ExpertEducationSelect", Array( _
	Array(, adInteger, , iCvID), _
	Array(, adInteger, , 2)))

Set objRsExpWke=GetDataRecordsetSPWithConn(objConnCustom, "usp_ExpCvvExperienceSelect", Array( _
	Array(, adInteger, , iCvID)))


objRsExpLng=GetDataOutParamsSPWithConn(objConnCustom, "usp_GetExpertProfDetails", Array( _
	Array(, adInteger, , iCvID), Array(, adInteger, , 0), Array(, adVarChar, 3, "Eng"), Array(, adInteger, , 0)), Array( _
	Array(, adVarChar, 1), Array(, adVarWChar, 400), Array(, adVarWChar, 50), Array(, adVarWChar, 255), Array(, adVarWChar, 1000), Array(, adVarWChar, 500), Array(, adVarWChar, 1000), Array(, adVarWChar, 500), Array(, adVarWChar, 1000), Array(, adVarWChar, 2000)))
sOtherLanguages=objRsExpLng(5)
Set objRsExpLng=Nothing


Set objRsExpLngNative=GetDataRecordsetSPWithConn(objConnCustom, "usp_ExpCvvLanguageSelect", Array( _
	Array(, adInteger, , iCvID), _
	Array(, adVarChar, 10, "native")))

Set objRsExpLngOther=GetDataRecordsetSPWithConn(objConnCustom, "usp_ExpCvvLanguageSelect", Array( _
	Array(, adInteger, , iCvID), _
	Array(, adVarChar, 10, "other")))


If bCvValidForMemberOrExpert=1 Or bCvValidForMemberOrExpert=5 Then

' Current address
bCurAddress=False
Set objTempRs=GetDataRecordsetSPWithConn(objConnCustom, "usp_ExpCvvAddressSelect", Array( _
	Array(, adInteger, , iCvID), _
	Array(, adInteger, , 3)))
	
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
Set objTempRs=GetDataRecordsetSPWithConn(objConnCustom, "usp_ExpCvvAddressSelect", Array( _
	Array(, adInteger, , iCvID), _
	Array(, adInteger, , 1)))
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
	If (Not (InStr(sPermAddressEmail, sEmail)>0 Or InStr(sEmail, sPermAddressEmail)>0)) Or CheckLength(sPermAddressEmail)=0 Then
		sCurAddressEmail=sEmail
	End If
	If Len(sCurAddressPhone)<5 And (Not (InStr(sPermAddressPhone, sUserPhone)>0 Or InStr(sUserPhone, sPermAddressPhone)>0)) Or CheckLength(sPermAddressPhone)=0 Then
		sCurAddressPhone=sUserPhone
	End If
End If

If bCurAddress=False And bPermAddress=False And (CheckLength(sEmail)>0 Or CheckLength(sUserPhone)>0) Then
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

	If (bCvValidForMemberOrExpert=1 Or bCvValidForMemberOrExpert=5) And InStr(sScriptFileName, "cv_view")>0 Then %>

	<form name="cvformat" method="get" action="<%=sApplicationHomePath & "view/cv_view.asp" %>">
	<input type="hidden" name="uid" value="<% =sCvUID %>">
	<input type="hidden" name="idproject" value="<% =iProjectID %>">
	<% ShowFeatureBoxHeader "Format the CV" %>
	<div class="content">
	<p class="sml" align="center">Select a format for CV</p>
	<img src="<% =sHomePath %>image/x.gif" width=1 height=3><br>
	<div align="center"><select style="font-face: Arial; font-size:8.5pt;" name="act" size=1>
	<%
	If sCvLanguage = cLanguageFrench Then
	%>
	<option value="ASR" <% If sCVFormat="" Then %>selected<% End If %>>assortis.com</option>
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
	</div>
	<% ShowFeatureBoxFooter %>
	</form><br/>
	
	<% ShowFeatureBoxHeader("Document options") %>
	<div class="content">
	<p class="sml"><a href="<%=Replace(sScriptFileName, "view", "save")%>?uid=<%=sCvUID%>&ftype=doc"><img src="/image/file_doc.gif" width=18 height=17 border=0 hspace=4 align="left">Save&nbsp;as&nbsp;Word&nbsp;document</a></p>
	<p class="sml">This option will save the CV in Microsoft&reg; Word* format on your local computer.</p>
	<p class="sml"><a href="<%=Replace(sScriptFileName, "view", "save")%>?uid=<%=sCvUID%>&ftype=prn"><img src="/image/file_prn.gif" width=18 height=17 border=0 hspace=4 align="left">Print&nbsp;this&nbsp;document</a></p>
	<p class="sml">This option will open the CV in Microsoft&reg; Word*  and you should click on Print button there.</p>
	</div>
	<% ShowFeatureBoxFooter %>
	<br />

	<% End If

	If sScriptFileName="cv_preview.asp" Then
	%>
	<br/>
	<% ShowFeatureBoxHeader("Contact details") %>
	<div class="content">
	<p>In order to get contact details for this expert please contact:</p>
	<p><b><a class="list" href="mailto:<% =objExpertDB.ContactEmail %>"><% =objExpertDB.Company.Name %>: <% =objExpertDB.ContactName %></a></b><br />
	(<a href="mailto:<% =objExpertDB.ContactEmail %>"><% =objExpertDB.ContactEmail %></a>)</p>
	
	<%
	' Get alternative CV owners
	Dim objExpertDBOtherList
	Set objExpertDBOtherList = New CCompanyExpertDBList
	objExpertDBOtherList.LoadData "usp_Ica_ExpertDBOwnerOtherSelect", Array( _
			Array(, adVarChar, 50, objExpertDB.Database),_
			Array(, adInteger, ,iCvID))
			
	Dim iExpertDBOtherLoop
	iExpertDBOtherLoop=0
	If objExpertDBOtherList.Count>0 Then
		While iExpertDBOtherLoop<objExpertDBOtherList.Count
		%>
	<br /><p><b><a class="list" href="mailto:<% =objExpertDBOtherList.Item(iExpertDBOtherLoop).ContactEmail %>"><% =objExpertDBOtherList.Item(iExpertDBOtherLoop).Company.Name %>: <% =objExpertDBOtherList.Item(iExpertDBOtherLoop).ContactName %></a></b><br />
	(<a href="mailto:<% =objExpertDBOtherList.Item(iExpertDBOtherLoop).ContactEmail %>"><% =objExpertDBOtherList.Item(iExpertDBOtherLoop).ContactEmail %></a>)</p>
		<%
		iExpertDBOtherLoop = iExpertDBOtherLoop + 1
		WEnd
	End If
	%>
	
	</p>
	</div>
	<% ShowFeatureBoxFooter %>
	<%
	End If
	
	
	If sScriptFileName="cv_verify.asp" Then
		If bCvValidForMemberOrExpert=1 Then
		%>
			<% ShowFeatureBoxHeader("CV options") %>
			<div class="content">
			<p">Some information from the CV<br>is missing?</p>
			<div align="center"><a href="../register/register.asp?uid=<% =sCvUID %>"><img src="<% =sHomePath %>image/bte_updatethiscv152.gif" vspace="4" border="0"></a></div>

			<p>If this CV is a duplicate and expert has another CV</p>
			<div align="center"><a href="../manage/cv_hide.asp?uid=<% =sCvUID %>"><img src="<% =sHomePath %>image/bte_hideexpert.gif" vspace="4" border="0"></a></div>

			<p>If expert asked to be removed or it's a fake CV</p>
			<div align="center"><a href="../manage/cv_remove.asp?uid=<% =sCvUID %>"><img src="<% =sHomePath %>image/bte_removeexpert.gif" vspace="4" border="0"></a></div>
			</div>
			<% ShowFeatureBoxFooter %>
			<br />
		<%
		Else
		%>
			<% ShowFeatureBoxHeader("Copy Expert") %>
			<div class="content">
			<p>This expert is the same as we wanted to register.</p>
			<p>We verified all available information and there is no doubt about it.</p>
			<p align="center"><b><a class="mt" href="../manage/cv_copy.asp?uid=<% =sCvUID %>">Copy expert to <% =objUserCompanyDB.DatabaseTitle %> database</a></b></p>
			<img src="<% =sHomePath %>image/x.gif" width=21 height=1 hspace=0 vspace=25 align="left">
			<p class="sml">* This action will be checked by ICA team in order to avoid any mistake.</p>
			</div>
			<% ShowFeatureBoxFooter %>
		<%
		End If
	End If
	
	If sScriptFileName="register6.asp" Then
	%>
		<form method="post" action="<% =sScriptFullName %>">
		<% ShowFeatureBoxHeader("CV status") %>
		<div class="content">
		<p>Currently the CV status is</p>
		<div align="center"><select name="cv_status" id="cv_status" style="width: 152px;">
		<option value="0">New CV</option>
		<% 
		Dim objStatusCVList
		Set objStatusCVList = New CStatusCVList
		objStatusCVList.LoadData
		objStatusCVList.ShowSelectItems(objExpertStatusCV.Status.ID)
		%>
		</select></div>
		<div align="center"><input type="image" src="<% =sHomePath %>image/bte_updatestatus.gif" vspace="4" border="0" alt="Update status"></a></div>

		<p><% If Len(sComments)<1 Then %>Some comments about this CV?<% Else %><% =sComments %><% End If %></p>
		<div align="center"><a href="../register/comments.asp?uid=<% =sCvUID %>"><img src="<% =sHomePath %>image/bte_editcomments.gif" vspace="4" border="0" alt="Edit comments"></a></div>
		</form>
		</div>
		<% ShowFeatureBoxFooter %>
		<br/>
		
		<% ShowFeatureBoxHeader("CV options") %>
		<div class="content">
		<p">Some information from the CV<br>is missing?</p>
		<div align="center"><a href="register.asp?uid=<% =sCvUID %>"><img src="<% =sHomePath %>image/bte_updatethiscv152.gif" vspace="4" border="0"></a></div>

		<p>If this CV is a duplicate and expert has another CV</p>
		<div align="center"><a href="../manage/cv_hide.asp?uid=<% =sCvUID %>"><img src="<% =sHomePath %>image/bte_hideexpert.gif" vspace="4" border="0"></a></div>

		<p>If expert asked to be removed or it's a fake CV</p>
		<div align="center"><a href="../manage/cv_remove.asp?uid=<% =sCvUID %>"><img src="<% =sHomePath %>image/bte_removeexpert.gif" vspace="4" border="0"></a></div>
		</div>
		<% ShowFeatureBoxFooter %>
		<br />
		
		
		<% ShowFeatureBoxHeader("CV formats") %>
		<div class="content">
		<p>To view this CV in different formats, to save or to print it</p>
		<div align="center"><a href="<% =sApplicationHomePath %>view/cv_view.asp<% =ReplaceUrlParams(ReplaceUrlParams(sParams, "idexpert"), "id=" & objExpertDB.DatabaseCode & iExpertID) %>"><img src="<% =sHomePath %>image/bte_formatcv152.gif" vspace="4" border="0"></a></div>
		</div>
		<% ShowFeatureBoxFooter %>
		<br />
	<%
	End If

End Sub

' Right Column for TOP EXPERTS:
Sub ShowTopExpCVFeatureBox
	%>
	<% ShowFeatureBoxHeader("My CV options") %>
	<div class="content">
	<ul class="compact">
		<li><a href="/backoffice/mycv/register.asp"><b>UPDATE MY CV</b><!--<img src="<% =sHomePath %>image/bte_updatecv.gif" vspace="4" border="0">--></a></li>
		<li><a href="/backoffice/mycv/document.asp?document=0"><b>DOCUMENTS</b></a></li>

		<%
		Dim objDocumentList
		Set objDocumentList = New CDocumentList
		objDocumentList.LoadDocumentListByExpertID iExpertID, ""
		
		If objDocumentList.Count=0 Then
		%>
			<p class="sml" style="padding: 2px 5px;">There are no documents uploaded yet.</p>
		<% 
		Else
			ShowDocumentListEditTable objDocumentList
		End If
		Set objDocumentList = Nothing
		%>
	</ul>
		<div align="center"><a href="/backoffice/mycv/document.asp?document=0"><img src="<% =sHomePath %>image/bte_upload.gif" vspace="4" border="0" alt="Upload document"></a></div>
	</div>
	<% ShowFeatureBoxFooter %>
	<br/>

	<form name="cvformat" method="get" action="<%=sApplicationHomePath & "mycv/cv_view.asp" %>">
	<input type="hidden" name="uid" value="<% =sCvUID %>" />
	<input type="hidden" name="idproject" value="<% =iProjectID %>" />
	<% ShowFeatureBoxHeader "Format the CV" %>
	<div class="content">
	<p class="sml" align="center">Select a format for CV&nbsp;</p>
	<div align="center"><select style="font-face: Arial; font-size:8.5pt;" name="act" size=1>
	<%
	If sCvLanguage = cLanguageFrench Then
	%>
	<option value="ASR" <% If sCVFormat="" Then %>selected<% End If %>>assortis.com</option>
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
	<img src="<% =sHomePath %>image/x.gif" width=1 height=6><br />
	<div align="center"><input type="image" src="<% =sHomePath %>image/bte_select.gif" alt="Select" height=18 vspace=0 border=0 /></a></div>
	</div>
	<% ShowFeatureBoxFooter %>
	</form><br/>
	<%
	If Len(Request.QueryString("uid")) > 0 Then
	%>
		<% ShowFeatureBoxHeader("Document options") %>
		<div class="content">
		<p class="sml"><a href="<%=Replace(sScriptFileName, "view", "save")%>?uid=<%=sCvUID%>&ftype=doc"><img src="/image/file_doc.gif" width=18 height=17 border=0 hspace=4 align="left">Save&nbsp;as&nbsp;Word&nbsp;document</a></p>
		<p class="sml">This option will save the CV in Microsoft&reg; Word* format on your local computer.</p>
		<p class="sml"><a href="<%=Replace(sScriptFileName, "view", "save")%>?uid=<%=sCvUID%>&ftype=prn"><img src="/image/file_prn.gif" width=18 height=17 border=0 hspace=4 align="left">Print&nbsp;this&nbsp;document</a></p>
		<p class="sml">This option will open the CV in Microsoft&reg; Word*  and you should click on Print button there.</p>
		</div>
		<% ShowFeatureBoxFooter %>
		<br />
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