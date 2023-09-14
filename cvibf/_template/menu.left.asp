<%
Dim sActiveMenuSection, sActiveMenuItem
If InStr(sScriptFullName, "/members/")>0 Then 
	sActiveMenuSection = "members"
ElseIf InStr(sScriptFullName, "/experts/")>0 Then 
	sActiveMenuSection = "experts"
Else
	sActiveMenuSection = ""
End If
%>
	<div id="leftmenu">
		<div class="box menu <% If sActiveMenuSection="experts" Then %>grey<% Else %>blue<% End If %>">
		<h3><span class="left">&nbsp;</span><span class="right">&nbsp;</span>Services for&nbsp;companies</h3>
		<div class="content">
		<ul>
		<li class="submenu <% If InStr(sScriptFileName, "recruitment_solutions")=1 Or _
		InStr(sScriptFileName, "experts_database")=1 Or _
		InStr(sScriptFileName, "job_posting")=1 Or _
		InStr(sScriptFileName, "cvip")=1 Or _
		InStr(sScriptFileName, "experts_recruitment")=1 Or _
		InStr(sScriptFileName, "inhouse_recruitment")=1 _
		Then %>selected<% End If %>"><a href="#">Recruitment Solutions&nbsp;<img src="<% =sHomePath %>image/box_extra.gif" hspace="2" vspace="2"></a><br /><span>Find consultants and in-house managers, acquire an efficient database structure</span>
					<ul id="rcsmenu" class="submenu">
						<li class="<% If InStr(sScriptFileName, "experts_database")=1 Then %>selected<% End If %> mfirst"><a href="<% =sHomePath %>en/members/experts_database.asp<% =sParams %>">Consultants Database</a><br /><span>Search consultants for your projects</span></li>
						<li class="<% If InStr(sScriptFileName, "experts_recruitment")=1 Then %>selected<% End If %>"><img src="<% =sHomePath %>image/new.gif" align="right" hspace="0" vspace="4" alt="New service"><a href="<% =sHomePath %>en/members/experts_recruitment.asp">Recruitment of Consultants</a><br /><span>We search & select the best consultants for you</span></li>
						<li class="<% If InStr(sScriptFileName, "job_posting")=1 Then %>selected<% End If %>"><a href="<% =sHomePath %>en/members/job_posting.asp<% =sParams %>">Job Posting Board</a><br /><span>Advertise job opportunities</span></li>
						<li class="<% If InStr(sScriptFileName, "inhouse_recruitment")=1 Then %>selected<% End If %>"><img src="<% =sHomePath %>image/new.gif" align="right" hspace="0" vspace="4" alt="New service"><a href="<% =sHomePath %>en/members/inhouse_recruitment.asp">International Project Managers Recruitment</a><br /><span>We assist you in internal recruitments</span></li>
						<li class="<% If InStr(sScriptFileName, "cvip")=1 Then %>selected<% End If %> last"><a href="<% =sHomePath %>en/members/cvip.asp">CViP</a><br /><span>Manage your pool of consultants</span></li>
					</ul>
		</li>
		<li class="<% If InStr(sScriptFileName, "tender_alerts")=1 Then %>selected<% End If %>"><a href="<% =sHomePath %>en/members/tender_alerts.asp<% =sParams %>">Daily Tender Alerts,<br/>Tenders &amp; Companies Databases</a><br /><span>Receive projects of interest,<br />view awarded companies' profiles</span></li>
		<li class="<% If InStr(sScriptFileName, "partners_search")=1 Then %>selected<% End If %>"><img src="<% =sHomePath %>image/new.gif" align="right" hspace="0" vspace="4" alt="New service"><a href="<% =sHomePath %>en/members/partners_search.asp<% =sParams %>">Consortium &amp; Local Partners</a><br /><span>We find consortium leaders and partners</span></li>
		<li class="<% If InStr(sScriptFileName, "trends")=1 Then %>selected<% End If %>"><img src="<% =sHomePath %>image/new.gif" align="right" hspace="0" vspace="4" alt="New service"><a href="<% =sHomePath %>en/members/trends.asp">Trends</a><br /><span>Receive tailored reports about development aid evolution</span></li>
		<li class="<% If InStr(sScriptFileName, "investments")=1 Then %>selected<% End If %> last"><img src="<% =sHomePath %>image/new.gif" align="right" hspace="0" vspace="4" alt="New service"><a href="<% =sHomePath %>en/members/investments.asp<% =sParams %>">Investments</a><br /><span>We advise you in investing in, buying or selling a company</span></li>
		</ul>
		</div>
		<h5><span class="left">&nbsp;</span><span class="right">&nbsp;</span></h5>
		</div>

		<div class="box menu <% If sActiveMenuSection="members" Then %>grey<% Else %>blue<% End If %>">
		<h3><span class="left">&nbsp;</span><span class="right">&nbsp;</span>Services for&nbsp;consultants</h3>
		<div class="content">
		<ul>
		<li class="<% If InStr(sScriptFileName, "cv_")=1 Then %>selected<% End If %>"><a href="<% =sHomePath %>en/experts/cv_about.asp<% =sParams %>">CV Registration</a> &nbsp;/&nbsp; <a href="<% =sHomePath %>en/experts/cv_update.asp<% =sParams %>">CV Update</a><br /><span>We make your CV available to hundreds of companies</span></li>
		<li class="<% If InStr(sScriptFileName, "sc_")=1 Then %>selected<% End If %>"><a href="<% =sHomePath %>en/experts/sc_about.asp<% =sParams %>">Special Info Pack</a><br /><span>Be the first to contact awarded & shortlisted companies / view their profiles</span></li>
		<li class="<% If InStr(sScriptFileName, "jbp")=1 Then %>selected<% End If %> last"><a href="<% =sHomePath %>en/experts/jbp_list.asp<% =sParams %>">Open Job Opportunities</a><br /><span>Latest announcement published by consulting companies</span></li>
		</ul>
		</div>
		<h5><span class="left">&nbsp;</span><span class="right">&nbsp;</span></h5>
		</div>
	</div>
	<script>
	$(document).ready(function(){
		$("li.submenu a").click(
			function(){
				$(this).parent().children("ul").show(0);
			}
		);	

		$("li.submenu a").hover(
			function(){
				$(this).parent().children("ul").show(0);
			},
			function(){
			}
		);	

		$("li.submenu").hover(
			function(){
			},
			function(){
				$(this).children("ul").hide(0);
			}
		);	
	});
	</script>