<!--#include file="html.header.asp"-->
<body>
<div id="holder">
	<!-- header -->
	<!--#include file="page.header.asp"-->

	<!-- content -->
	<div id="content" class="searchform">
		<h2 class="service_title">Daily Tender Alerts & Companies: <span class="service_slogan">Downloading blocked</span></h2>
		<% ShowMessageStart "error", 580 %>
			<p>You have reached your <b>daily data download limit</b><br/>and your account is blocked from downloading any more companies profiles or project details today.</p>
			<p>If you require further assistance, please do not hesitate to contact one of our Support Team
			<br/> who will be happy to assist you with your enquiry.</b>
		<% ShowMessageEnd %>
	</div>
</div>
	<!-- footer -->
	<!--#include file="page.footer.asp"-->
<% CloseDBConnection %>
</body>
<!--#include file="html.footer.asp"-->
<% Response.End %>
