<!--#include file="../cv_data.asp"-->
<%

CheckUserLogin sScriptFullNameAsParams
CheckExpertID()

sParams = ReplaceUrlParams(sParams, "id")
sParams = ReplaceUrlParams(sParams, "uid")
sParams = ReplaceUrlParams(sParams, "idproject")
sParams = ReplaceUrlParams(sParams, "idexpert")
%>

<!--#include virtual="/_template/html.header.asp"-->
<body>
<div id="holder">
	<!-- header -->
	<!--#include virtual="/_template/page.header.asp"-->

	<!-- content -->
	<div id="content" class="searchform">

	<%
	If iExpertID > 0 Then
		objTempRs=GetDataOutParamsSP("usp_Ica_ExpertRestore", Array( _
			Array(, adInteger, , objExpertDB.ID), _
			Array(, adInteger, , iExpertID)), Array( _ 
			Array(, adInteger)))
		
		If objTempRs(0) >= 1 Then
			Response.Write "<br><br><br><br><h3><span class=""service_slogan"">The CV of the expert " & objExpertDB.DatabaseCode & iExpertID & " was successfully restored.</span></h3>"
		End If
		%><br><br>
		<a href="<% =sApplicationHomePath %>register/register6.asp<% =ReplaceUrlParams(sParams, "uid=" & sCvUID) %>"><img src="<% =sHomePath %>image/bte_continue.gif" border=0></a>
		<%
		ShowStandardPageFooter
		Response.End
	End If
	%>
	
	</div>
</div>
	<!-- footer -->
	<!--#include virtual="/_template/page.footer.asp"-->
<% CloseDBConnection %>
</body>
<!--#include virtual="/_template/html.footer.asp"-->
