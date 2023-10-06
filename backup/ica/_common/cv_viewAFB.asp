<%
'--------------------------------------------------------------------
'
' Expert's CV. View in AFB format
'
'--------------------------------------------------------------------
%>
<!--#include file="cv_data.asp"-->
<%
' Log:
If bCvValidForMemberOrExpert > 0 Then
	' 38 - View CV Full
	iLogResult = LogActivity(38, "CVID=" & Cstr(iCvID) & " Format: " & sCvFormat, "", "")
End If
%>

<!--#include virtual="/_common/_class/main.asp"-->
<!--#include virtual="/_common/_class/project.asp"-->
<!--#include virtual="/_common/_class/expert.asp"-->
<!--#include virtual="/_common/_class/expert.project.asp"-->
<!--#include virtual="/_common/_class/expert.cv.language.asp"-->

<!--#include virtual="/_template/html.header.asp"-->
<body class="cv-view">
	<!-- header -->
	<!--#include virtual="/_template/page.header.asp"-->

	<!-- content -->
	<div id="content" class="workscreen">
	<% 
	If Not bIsMyCV Then 
		%><div id="hdrUpdatedList" class="colCCCCCC uprCse f17 spc01 botMrgn10"><span class="service_title">Curriculum Vitae.</span> Expert ID: <% =objExpertDB.DatabaseCode %><%=iCvID%>
		<br /><span style="font-size:0.7em;text-transform:none;color:#999">African Development Bank format</span>
		<% If bCvValidForMemberOrExpert=1 Or bCvValidForMemberOrExpert=5 Then %>
		<% Else %>
		<br /><span style="font-size:0.7em;text-transform:none">Contact details free version</span>
		<% End If %>
		</div>
		<% 
	Else
		%><div class="colCCCCCC uprCse f17 spc01 botMrgn10"><span class="service_title">Curriculum Vitae</span></div>
		<%
	End If

' Personal information
	WriteTableHeader
	WriteDataRow "Proposed Position:", "<font color=""#C0C0C0"">(fill in with your data)</font>"
	WriteDataRow "Name of Firm:", "<font color=""#C0C0C0"">(fill in with your data)</font>"
	WriteDataRow "Name of Staff:", sFullName
	WriteDataRow "Profession:", sProfession
	WriteDataRow "Date of birth:", ConvertDateForText(sBirthDate, "&nbsp;", "DDMMYYYY")
	WriteSpaceRow
	WriteSpaceRow
	WriteDataRow "Address:<br><br>(phone / e-mail)", sPermAddress
	WriteSpaceRow

	WriteDataRow "Years with Firm/Entity:", "<font color=""#C0C0C0"">(fill in with your data)</font>"
	WriteDataRow "Nationality:", sNationality 
	WriteDataRow "Membership&nbsp;in<br>Professional&nbsp;Societies:", "<br>" & sMemberships
	WriteDataRow "Detailed Tasks Assigned:", "<font color=""#C0C0C0"">(fill in with your data)</font><br><br>"

	WriteDataRow "<b>Key&nbsp;qualifications:</b>", ConvertText(sKeyQualification)

	WriteTableFooter

' Education
	WriteTableHeader
	WriteDataRow "<b>Education:</b>", " "
	WriteTableFooter
	WriteGridTableHeader
	WriteGridDataRow 4, "25%#|#13%#|#13%#|#50%", "<b>Institution</b>#|#<b>Start&nbsp;date</b>#|#<b>End&nbsp;date</b>#|#<b>Degree&nbsp;/&nbsp;Diploma&nbsp;obtained</b>"

	If objRsExpEdu.Eof Then
		WriteGridDataRow 4, "", "-#|#-#|#-#|#-"
	Else
        i=1
	While Not objRsExpEdu.Eof
		WriteGridDataRow 4, "", objRsExpEdu("InstNameEng") & "#|#" & ConvertDateForText(objRsExpEdu("eduStartDate"), "&nbsp;", "MMYYYY") & "#|#" & ConvertDateForText(objRsExpEdu("eduEndDate"), "&nbsp;", "MMYYYY") & "#|#" & Trim(objRsExpEdu("edtDescriptionEng") & " " & objRsExpEdu("eduDiploma1Eng")) & "<br>" & Trim(objRsExpEdu("edsDescriptionEng") & " " & objRsExpEdu("id_EduSubject1Eng"))
		i=i+1
		objRsExpEdu.MoveNext
	WEnd
	End If 
	objRsExpEdu.Close  
	Set objRsExpEdu=Nothing
        i=1
	While Not objRsExpTrn.eof
		WriteGridDataRow 4, "", objRsExpTrn("eduDiploma1Eng") & "#|#" & ConvertDateForText(objRsExpTrn("eduStartDate"), "&nbsp;", "MMYYYY") & "#|#" & ConvertDateForText(objRsExpTrn("eduEndDate"), "&nbsp;", "MMYYYY") & "#|#" & objRsExpTrn("edtDescriptionEng")
		i=i+1
		objRsExpTrn.MoveNext
	WEnd
	objRsExpTrn.Close  
	Set objRsExpTrn=Nothing
	WriteGridTableFooter


' Employment records
	WriteTableHeader
	WriteDataRow "<b>Employment&nbsp;records:</b>", " "
	WriteTableFooter

	While Not objRsExpWke.Eof
		
		WriteGridTableHeader
		WriteGridDataRow 2, "25%#|#75%", "<b>Date</b>" & "#|#" & ConvertDateForText(objRsExpWke("wkeStartDate"), "&nbsp;", "MMYYYY") & " - " & ConvertDateForText(objRsExpWke("wkeEndDate"), "&nbsp;", "MMYYYY")
		WriteGridDataRow 2, "", "<b>Employer</b>"  & "#|#" &  objRsExpWke("wkeOrgNameEng")
		WriteGridDataRow 2, "", "<b>Position held</b>" & "#|#" & objRsExpWke("wkePositionEng")
		dflag=0
		sDescription=ConvertText(objRsExpWke("wkeDescriptionEng"))
		WriteGridDataRow 2, "", "<b>Description of duties</b>" & "#|#" & sDescription
		WriteGridTableFooter

	objRsExpWke.MoveNext
	WEnd
	objRsExpWke.Close
	Set objRsExpWke=Nothing


' Languages
	WriteTableHeader
	WriteDataRow "<b>Languages:</b>", " "
	WriteTableFooter

	If (Not objRsExpLngOther.Eof) Or (Not objRsExpLngNative.Eof) Then
		WriteGridTableHeader
		WriteGridDataRow 4, "25%#|#25%#|#25%#|#25%", "<b>Language</b>#|#<b>Reading</b>#|#<b>Speaking</b>#|#<b>Writing</b>"

		While Not objRsExpLngNative.Eof
			WriteGridDataRow 4, "", objRsExpLngNative("lngNameEng") & "#|#" & arrLanguageLevelTitle(objRsExpLngNative("exlReading")) & "#|#" & arrLanguageLevelTitle(objRsExpLngNative("exlSpeaking")) & "#|#" & arrLanguageLevelTitle(objRsExpLngNative("exlWriting"))
			objRsExpLngNative.MoveNext
		WEnd

		While Not objRsExpLngOther.Eof
			WriteGridDataRow 4, "", objRsExpLngOther("lngNameEng") & "#|#" & arrLanguageLevelTitle(objRsExpLngOther("exlReading")) & "#|#" & arrLanguageLevelTitle(objRsExpLngOther("exlSpeaking")) & "#|#" & arrLanguageLevelTitle(objRsExpLngOther("exlWriting"))
			objRsExpLngOther.MoveNext
		WEnd
		WriteGridTableFooter
	End If
	objRsExpLngNative.Close
	Set objRsExpLngNative=Nothing	
	objRsExpLngOther.Close
	Set objRsExpLngOther=Nothing	
	%>

	<table cellspacing=0 cellpadding=0 align="center" width="96%">
	<tr><td width="100%" bgcolor="#FFFFFF"><p class="txt"><b>Certification:</b> <br><br>I, the undersigned, certify that to the best of my knowledge and belief, these data correctly describe me, my qualifications, and my experience.<br><br>____________________________________________________<br>Signature of staff member and authorized representative of the firm &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; Date&nbsp;of&nbsp;signing&nbsp;<font color="#C0C0C0">(&nbsp;Day&nbsp;/&nbsp;Month&nbsp;/&nbsp;Year&nbsp;)</font><br></td></tr>
	</table><br>
	
	</div>
	<div id="rightspace">
	<!-- feature boxes -->
	<% 
	If Not bIsMyCV Then
		ShowExpCVFeatureBox 
	Else
		ShowTopExpCVFeatureBox
	End If 
	%>
	</div>

	<!-- footer -->
	<!--#include file="../_template/page.footer.asp"-->
<% CloseDBConnection %>
</body>
<!--#include file="../_template/html.footer.asp"-->
