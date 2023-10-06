<!--#include virtual="/_template/html.header.asp"-->
<body>
<div id="holder">
	<!-- header -->
	<!--#include virtual="/_template/page.header.asp"-->

	<!-- content -->
	<div id="content" class="searchform">

		<h2 class="service_title">CV check. <span class="service_slogan">The expert is not found.</span>
		</h2><br/>

	<% ShowMessageStart "info", 580 %>
	<b>There are no CVs similar to the one you are going to encode.</b><br>
	<% ShowMessageEnd %><br/>

	</div>
</div>
	<!-- footer -->
	<!--#include file="../_template/page.footer.asp"-->
<% CloseDBConnection %>
</body>
<!--#include file="../_template/html.footer.asp"-->
