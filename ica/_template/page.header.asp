<%
sTempParams=ReplaceUrlParams(sParams, "url")
sTempParams=ReplaceUrlParams(sTempParams, "id")
sTempParams=ReplaceUrlParams(sTempParams, "idproject")
sTempParams=ReplaceUrlParams(sTempParams, "idexpert")
sTempParams=ReplaceUrlParams(sTempParams, "t")

Dim assortisButtonUrl, assortisLogoutUrl, icaLogoutUrl, icaExpertsLogoutUrl
assortisButtonUrl = "http://www.assortis.com/login.asp?url=/en/members/bsc_list.asp&token=" & sAsortisLoginToken
assortisLogoutUrl = "http://www.assortis.com/logout4ica.asp"
icaLogoutUrl = "http://" & sIcaServer & "/intranet/logout"
icaExpertsLogoutUrl = "http://experts.icaworld.net/logout.asp"

' render header:
%>
<h1 class="off_jws">ICANET</h1>
<hr/>

<script>
	$(function () {
		$('#mainMenu li').mouseover(function () {
			if ($(this).children('ul').length > 0) {
				//$('#mainMenu ul').hide();
				$(this).children('ul').show();
			}
		});

		$('#mainMenu li').mouseout(function () {
			if ($(this).children('ul').length > 0) {
				$(this).children('ul').hide();
			}
		});
	});
</script>

<div id="header">

	<div class="headerlink" onclick="location.href='http://<% =sIcaServer %>/Intranet/News'">&nbsp;</div>

	<div style="display: block;background: #00a8ac url(/Resources/images/topmenu-edge.png) no-repeat left top;color: #FFF;
    height: 36px;padding: 0 0 0 25px;text-align: right;line-height: 26px;font-size: 12px;font-weight: 800;letter-spacing: -0.12px;float:right;width:970px;">

		<div class="user-quick-topmenu floatRight lgn alignRight">
			<div style="display: inline-block;position:relative;"><a href="https://twitter.com/ICAnetWorld" target="_blank" title="@ICAnetWorld on Twitter" class="linkedin-link hdr-icon" style="left:-30px;"></a></div>
			<div style="display: inline-block;position:relative;"><a href="https://www.linkedin.com/company/international-consulting-alliance-ica-" target="_blank" title="ICA on LinkedIn" class="twitter-link hdr-icon" style="left:-70px;"></a></div>
			<div class="hdr-info_item"><div class="sep"></div><a href="http://<% =sIcaServer %>/Intranet/Contact" class="hdr-nav_a">Contact</a></div>
			<div class="hdr-info_item hdr-user">
                <div class="sep"></div>
                <% =sUserFullName %> [<% =sUserCompany %>]
                <div class="hdr-userContainer">
                    <% If iUserCompanyRoleID = 7 Then %>
                        <div>
                            <a href="http://<% =sIcaServer %>/Intranet/TopExpertProfile?id=0" class="hdr-nav_a">
                                My Profile
                            </a>
                        </div>
                        <div>
                            <a href="http://<% =sIcaServer %>/Intranet/ChangePassword" class="hdr-nav_a">
                                Change Password
                            </a>
                        </div>
                    <% End If %>
                    <div>
                        <a id="IcaLogOutBtn" href="javascript:void(0)" class="hdr-nav_a">logout</a>
                    </div>
                </div>
            </div> 
            <% If iUserCompanyRoleID <> 7 Then %>
                <div class="hdr-info_item"><div class="sep"></div><a href="http://<% =sIcaServer %>/Intranet/MyAccount" class="hdr-nav_a">my profile</a></div>
            <% End If %>
			<!-- <div class="hdr-info_item"><div class="sep"></div><a id="IcaLogOutBtn" href="javascript:void(0)" class="hdr-nav_a">logout</a></div>			 -->
			<%If (iUserCompanyRoleID=2 Or iUserCompanyRoleID=3 Or iUserCompanyRoleID=4 Or iUserCompanyRoleID=5) Then %>
			    <div class="hdr-info_item" style="margin-left:30px;"><a href="http://<% =sIcaServer %>/Intranet/Dashboard" style="text-decoration:none;padding: 5px 10px;background:#FFF;color:#00a8ac;" title="Dashboard">ADMIN PANEL</a></div>
			<% End If %>
		</div>
	</div>
	<br class="clear"/>
	<script type="text/javascript">
		$(function () {
			$('#IcaLogOutBtn').click(function () {
				document.location = "<%=assortisLogoutUrl%>?exp=<%=Server.URLEncode(icaExpertsLogoutUrl) %>&ica=<%=Server.URLEncode(icaLogoutUrl) %>";
			});
		});
	</script>

	
		<div class="topMenu floatRight">
			<p class="off_jws"><strong>Main menu</strong></p>
			<ul id="mainMenu" class="floatRight">
                <li class="first"><a href="http://<% =sIcaServer %>/Intranet/News" title="NEWS">NEWS</a></li>
                <li class="hdr-navitem">
					<a href="http://<% =sIcaServer %>/Intranet/my_projects" title="PROJECTS">NETWORK</a>
					<ul class="hdr-navitem_subNav">
						<li>
							<a href="http://<% =sIcaServer %>/Intranet/ICACommunity" >MEMBERS</a>
						</li>
						<li>
							<a href="http://<% =sIcaServer %>/Intranet/ICAShowcase" >ICA WORLD MAP</a>
						</li>
					</ul>
				</li>
				
                <li class="hdr-navitem">
                    <% If iUserCompanyRoleID <> 7 Then  %>
                        <a href="#" title="Projects">PROJECTS</a>
                        <ul class="hdr-navitem_subNav">
                            <li>
                                <a href="http://<% =sIcaServer %>/Intranet/my_projects?av=my" class="hdr-contents-links ">
                                    My Projects
                                </a>
                            </li>
                            <li>
                                <a href="http://<% =sIcaServer %>/Intranet/my_projects?av=our" class="hdr-contents-links ">
                                    Our Projects
                                </a>
                            </li>
                            <li>
                                <a href="http://<% =sIcaServer %>/Intranet/my_projects?av=ica" class="hdr-contents-links ">
                                    ICA Projects
                                </a>
                            </li>
                            <li>
                                <a href="http://<% =sIcaServer %>/Intranet/References?Length=8" class="hdr-contents-links ">Reference Database</a>
                            </li>
                        </ul>
                    <% Else %>
                        <a href="http://<% =sIcaServer %>/Intranet/my_projects?av=my" title="NEWS">MY PROJECTS</a>
                    <% End If %>
                </li>
				<li class="hdr-navitem">
					<a href="http://<% =sIcaServer %>/Intranet/MyVacancies" title="VACANCIES">VACANCIES</a>
					<ul class="hdr-navitem_subNav">
						<li>
							<a href="http://<% =sIcaServer %>/Intranet/MyVacancies?av=my" >MY VACANCIES</a>
                        </li>
                        <% If iUserCompanyRoleID <> 7 Then  %>
                            <li>
                                <a href="http://<% =sIcaServer %>/Intranet/OurVacancies?av=our" >OUR VACANCIES</a>
                            </li>
                            <li>
                                <a href="http://<% =sIcaServer %>/Intranet/icavacancies?av=ica" >ICA VACANCIES</a>
                            </li>
                        <% End If %>
					</ul>
                </li>
                <% If iUserCompanyRoleID <> 7 Then %> 
                    <li class="hdr-navitem">
                        <a href="http://<% If sIcaServerType<>"www." Then Response.Write sIcaServerType %>experts.icaworld.net/" title="Experts">EXPERTS</a>
                        <ul class="hdr-navitem_subNav">
                            <li>
                                <a href="http://<% If sIcaServerType<>"www." Then Response.Write sIcaServerType %>experts.icaworld.net/backoffice/search/exp_search.asp">SEARCH</a>
                            </li>    
                            <li>
                                <a href="http://<% If sIcaServerType<>"www." Then Response.Write sIcaServerType %>experts.icaworld.net/backoffice/register/register.asp" >REGISTER EXPERT</a>
                            </li>    
                            <li>
                                <a href="http://<% If sIcaServerType<>"www." Then Response.Write sIcaServerType %>experts.icaworld.net/backoffice/manage.asp" >MANAGE OUR DATABASE</a>
                            </li>    
                            <li>
                                <a href="http://<% If sIcaServerType<>"www." Then Response.Write sIcaServerType %>experts.icaworld.net/backoffice/circle.asp" >MY EXPERTS CIRCLE</a>
                            </li>
							<li>
                                <a href="http://www.icaworld.net/Intranet/OurTopExpert" >OUR TOP EXPERTS</a>
                            </li>
						</ul>
                    </li>
                <% Else %>
                    <li class="hdr-navitem">
                        <a href="http://<% =sIcaServer %>/Intranet/TopExpertProfile?expertid=0" title="VACANCIES">MY PROFILE</a>
                        <ul class="hdr-navitem_subNav">
                            <li>
                                <a href="http://<% If sIcaServerType<>"www." Then Response.Write sIcaServerType %>experts.icaworld.net/" title="Edit my CV">MY CV</a>
                            </li>
                        </ul>
                    </li>
                <% End If %>
				<li class="hdr-navitem">
					<a href="http://<% =sIcaServer %>/Hubs" title="HUBS">HUBS</a> 
				</li>
				
                <li class="hdr-navitem">
					<a href="javascript:void(0);" title="TOOLS">TOOLS</a>
					<ul class="hdr-navitem_subNav">
						<li>
							<a href="http://<% =sIcaServer %>/Intranet/icatools" >ICA TOOLS</a>
						</li>
						<li>
							<a href="http://<% =sIcaServer %>/Intranet/projecttools" >PROJECT TOOLS</a>
						</li>
					</ul>
				</li>

                <% If (iUserCompanyRoleID <> 7) Then  %>
                    <li><a href="<% =assortisButtonUrl %>" title="ASSORTIS" class="btn-assortis"></a></li>
                <% End If %>
			</ul>
		</div>
		<br class="clr h0" />
		<div class="topSubMenu floatRight">
			' 
		</div>
<hr/>
<div class="clr h0"><!--  --></div>
</div>

<a id="paramsHolder" name="" style="display:none;"></a>

<div id="pageContent"><p class="off_jws"><strong>Page content</strong></p>