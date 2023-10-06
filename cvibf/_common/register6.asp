<% 
'--------------------------------------------------------------------
'
' CV registration.
' Full format. Check CV & send an email
'
'--------------------------------------------------------------------
%>
<!--#include virtual="/_template/asp.header.nocache.asp"-->
<!--#include file="cv_data.asp"-->
<%
' Check user's access rights
If sApplicationName<>"expert" Then
	CheckUserLogin sScriptFullNameAsParams
Else
	Response.Redirect sHomePath & "apply/"
End If
CheckExpertID()
Set objConnCustom = Server.CreateObject("ADODB.Connection")
objConnCustom.Open "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Data Source=MAGNESA;Initial Catalog=" & objExpertDB.DatabasePath & ";"

' The page is accessible only if expert id is valid
If iExpertID<=0 Then Response.Redirect(sApplicationHomePath & "register/register.asp" & sParams)
%>
<!--#include virtual="/_common/_class/main.asp"-->
<!--#include virtual="/_common/_class/project.asp"-->
<!--#include virtual="/_common/_class/expert.asp"-->
<!--#include virtual="/_common/_class/expert.project.asp"-->
<!--#include virtual="/_common/_class/status_cv.asp"-->
<!--#include virtual="/_common/_class/expert.status_cv.asp"-->
<%
Dim objExpertStatusCV
Set objExpertStatusCV = New CExpertStatusCV
objExpertStatusCV.Expert.ID=iExpertID
objExpertStatusCV.LoadData
%>
<!--#include virtual="/_common/register/update_status.asp"-->
<!--#include virtual="/_template/html.header.asp"-->
<body>
<div id="holder">
	<!-- header -->
	<!--#include virtual="/_template/page.header.asp"-->

	<!-- content -->
	<div id="content" class="workscreen">
	<% 
	If Not bIsMyCV Then 
		%><h2 class="service_title">Curriculum Vitae. <span class="service_slogan">Expert ID: <% =objExpertDB.DatabaseCode %><%=iCvID%>
		<% If bCvValidForMemberOrExpert=1 Or bCvValidForMemberOrExpert=5 Then %>
		<% Else %>
		<br />Contact details free version
		<% End If %>
		</span>
		</h2><br/>
		<% 
	End If %>

	<!--#include file="register6_data.asp"-->
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
</div>
	<!-- footer -->
	<!--#include file="../_template/page.footer.asp"-->
<% CloseDBConnection %>
</body>
<!--#include file="../_template/html.footer.asp"-->
