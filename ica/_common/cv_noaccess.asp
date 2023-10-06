<%
'--------------------------------------------------------------------
'
' Expert's CV. View in assortis.com format
' With or without contact details
'
'--------------------------------------------------------------------
%>
<!--#include file="cv_data.asp"-->

<!--#include virtual="/_common/_class/main.asp"-->
<!--#include virtual="/_common/_class/project.asp"-->
<!--#include virtual="/_common/_class/expert.asp"-->
<!--#include virtual="/_common/_class/expert.project.asp"-->
<!--#include virtual="/_template/html.header.asp"-->
<body>
<div id="holder">
	<!-- header -->
	<!--#include virtual="/_template/page.header.asp"-->

	<!-- content -->
	<div id="content" class="workscreen">
	<% 
	If Not bIsMyCV Then 
		%><h2 class="service_title">Your access to this section of the ICA Platform is restricted.<br>
		</h2><br/>
		<p>If you wish to get full access and be able to see all <% =GetExpCount("all") %> experts in the ICA Common Database of Experts please contact your <a href="mailto:info@icaworld.net">ICA Team</a>.</p>
	<%
	End If
	%>
	</div>
	<div id="rightspace">
	</div>
</div>
	<!-- footer -->
	<!--#include file="../_template/page.footer.asp"-->
<% CloseDBConnection %>
</body>
<!--#include file="../_template/html.footer.asp"-->
