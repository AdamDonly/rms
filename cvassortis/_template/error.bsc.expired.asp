<!--#include file="html.header.asp"-->
<body>
<div id="holder">
	<!-- header -->
	<!--#include file="page.header.asp"-->

	<!-- content -->
	<div id="content" class="searchform">
		<h2 class="service_title">Daily Tender Alerts & Companies: <span class="service_slogan">Subscription is not activated</span></h2>
		<% ShowMessageStart "error", 580 %>
			<p>Your subscription for this service is not activated because we have not received confirmation of your payment.</p>
			<p>If you believe you have sent your payment in time, please write to <a href="mailto:info@assortis.com">info@assortis.com</a><br />
			We hope to resolve this situation and look forward to again providing you valuable, profiled tender information.</p>
		<% ShowMessageEnd
		Response.End
		%>
	</div>
</div>
	<!-- footer -->
	<!--#include file="page.footer.asp"-->
<% CloseDBConnection %>
</body>
<!--#include file="html.footer.asp"-->

