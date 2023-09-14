		<ul class="headermenu location">
		<% If InStr(sScriptFileName, "default")=1 _
			Then %>
			<li>Home</li>
		<% Else %>
			<li><a href="<% =sHomePath %>">Home</a>&nbsp;/</li>
		<% End If %>
		<% If InStr(sScriptFileName, "tender_alerts")=1 _
			Then %>
			<li>Daily Tender Alerts &amp; Companies Database</li>
		<% End If %>
		<% If InStr(sScriptFileName, "bsc_")=1 _
			Then %>
			<li><a href="<% =sHomePath %>en/members/tender_alerts.asp">Daily Tender Alerts &amp; Companies Database</a>&nbsp;/</li>
		<% End If %>
		<% If InStr(sScriptFileName, "recruitment_solutions")=1 _
			Then %>
			<li>Recruitment Solutions</li>
		<% End If %>
		<% If InStr(sScriptFileName, "experts_database")=1 Or _
			InStr(sScriptFileName, "job_posting")=1 Or _
			InStr(sScriptFileName, "cvip")=1 Or _
			InStr(sScriptFileName, "experts_recruitment")=1 Or _
			InStr(sScriptFileName, "inhouse_recruitment")=1 Or _
			InStr(sScriptFileName, "exp_")=1 Or _ 
			InStr(sScriptFileName, "jbp_")=1 Or _ 
			InStr(sScriptFileName, "cv_view")=1 Or _
			InStr(sScriptFileName, "cv_preview")=1 _
			Then %>
			<li><a href="<% =sHomePath %>en/members/recruitment_solutions.asp">Recruitment Solutions</a>&nbsp;/</li>
		<% End If %>
		<% If InStr(sScriptFileName, "experts_database")=1 _
			Then %>
			<li>Consultants Database</li>
		<% End If %>
		<% If InStr(sScriptFileName, "exp_")=1 Or _
		InStr(sScriptFileName, "cv_view")=1 Or _
		InStr(sScriptFileName, "cv_preview")=1 _
			Then %>
			<li><a href="<% =sHomePath %>en/members/experts_database.asp">Consultants Database</a>&nbsp;/</li>
		<% End If %>
		<% If Request.QueryString("qid")>"" And _
		(InStr(sScriptFileName, "exp_results")=1) _
			Then %>
			<li>Search results</li>
		<% End If %>
		<% If Request.QueryString("qid")>"" And _
		(InStr(sScriptFileName, "cv_view")=1 Or _
		InStr(sScriptFileName, "cv_preview")=1) _
			Then %>
			<li><a href="<% =sHomePath %>en/members/exp_results.asp<% =AddUrlParams(sParams, "qid=" & Request.QueryString("qid")) %>">Search results</a>&nbsp;/</li>
		<% End If %>
		<% If InStr(sScriptFileName, "cv_view")=1 Or _
		InStr(sScriptFileName, "cv_preview")=1 _
			Then %>
			<li>CV review</li>
		<% End If %>
		
		<% If InStr(sScriptFileName, "job_posting")=1 _
			Then %>
			<li>Job Posting Board</li>
		<% End If %>
		<% If InStr(sScriptFileName, "jbp_")=1 _
			Then %>
			<li><a href="<% =sHomePath %>en/members/job_posting.asp">Job Posting Board</a>&nbsp;/</li>
		<% End If %>
		<% If InStr(sScriptFileName, "cvip")=1 _
			Then %>
			<li>CViP</li>
		<% End If %>
		<% If InStr(sScriptFileName, "experts_recruitment")=1 _
			Then %>
			<li>Recruitment of Consultants</li>
		<% End If %>
		<% If InStr(sScriptFileName, "inhouse_recruitment")=1 _
			Then %>
			<li>International Project Managers Recruitment</li>
		<% End If %>
		<% If InStr(sScriptFileName, "partners_search")=1 _
			Then %>
			<li>Consortium & Local Partners</li>
		<% End If %>
		<% If InStr(sScriptFileName, "investments")=1 _
			Then %>
			<li>Investments</li>
		<% End If %>
		<% If InStr(sScriptFileName, "trends")=1 _
			Then %>
			<li>Trends</li>
		<% End If %>
		<% If InStr(sScriptFileName, "mmb_page")=1 _
			Then %>
			<li>My account</li>
		<% End If %>
		<% If InStr(sScriptFileName, "mmb_login")=1 _
			Then %>
			<li>Login</li>
		<% End If %>
		<% If InStr(sScriptFileName, "_price")>0 _
			Then %>
			<li>Price list</li>
		<% End If %>
		<% If InStr(sScriptFileName, "bsc_list")>0 _
			Then %>
			<li>Daily notices</li>
		<% End If %>
		<% If InStr(sScriptFileName, "_register")>0 _
			Then %>
			<li>Registration</li>
		<% End If %>
		<% If InStr(sScriptFileName, "_search")>0 _
			And InStr(sScriptFileName, "_search")<5 _
			And sAccessType="trial" _
			Then %>
			<li>Sample search</li>
		<% ElseIf InStr(sScriptFileName, "_search")>0 _
			And InStr(sScriptFileName, "_search")<5 _
			Then %>
			<li>Search</li>
		<% End If %>
		<% If InStr(sScriptFileName, "about")=1 _
			Then %>
			<li>About assortis</li>
		<% End If %>
		<% If InStr(sScriptFileName, "terms")=1 _
			Then %>
			<li>Terms of use</li>
		<% End If %>
		<% If InStr(sScriptFileName, "contact")=1 _
			Then %>
			<li>Contact us</li>
		<% End If %>
		</ul>

		
		