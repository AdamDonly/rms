<%@ LANGUAGE="VBSCRIPT" %>
<% Option Explicit %>

<!--#include file="../dbc.asp"-->
<!--#include file="../fnc.asp"-->
<!--#include file="../_forms/frmInterface.asp"-->

<html>
<head>
<title><% =GetLabel(sCvLanguage, "CV registration") %>. <% =GetLabel(sCvLanguage, "Education") %></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" type="text/css" href="<% =sHomePath %>styles.css">
<script type="text/javascript" src="/_scripts/js/main.js"></script>
<script language="JavaScript">
<!--
function validateForm() { 
	var f=document.RegForm;
<% If Len(sBackOffice)<3 Then %>
	if (f.exp_Inst_name.value!="" || f.exp_edu_subj.selectedIndex>0 || f.exp_edu_syear.selectedIndex>0) { 
		AddEducation(1); 
	} else 
<% End If %>
	{ f.submit(); }
}

function AddEducation(cont_next) {  
	var f=document.RegForm;
	if (cont_next!=1) {
		f.exp_edu_continue.value="0"; 
	}
<% If sApplicationName="external" Then %>
	f.submit();
	return;
<% End If %>

<% If Len(sBackOffice)<3 Then %>
	if (f.exp_Inst_name.value=="") {
		alert("<% =GetLabel(sCvLanguage, "Please fill in the institution name") %>"); document.RegForm.exp_Inst_name.focus(); return;
	}
	var start_month=parseInt(f.exp_edu_smonth.options[f.exp_edu_smonth.selectedIndex].value);
	var start_year=parseInt(f.exp_edu_syear.options[f.exp_edu_syear.selectedIndex].value);
	var end_month=parseInt(f.exp_edu_emonth.options[f.exp_edu_emonth.selectedIndex].value);
	var end_year=parseInt(f.exp_edu_eyear.options[f.exp_edu_eyear.selectedIndex].value);
	if (end_month==0 || end_year==0) {
		alert("<% =GetLabel(sCvLanguage, "Please fill in the education end date") %>"); return;
	}
	if ((start_year>end_year) || (start_year>0 && start_month>0 && start_year==end_year && start_month>end_month)) {
		alert("<% =GetLabel(sCvLanguage, "Please fill in the education dates properly") %>"); return;
	}
	if (f.exp_edu_diploma.selectedIndex==0 && f.exp_edu_diploma1.value=="") {
		alert("<% =GetLabel(sCvLanguage, "Please specify a type of diploma or degree obtained") %>"); return;
	}
	if (f.exp_edu_subj.selectedIndex==0) {
		alert("<% =GetLabel(sCvLanguage, "Please specify the education subject") %>"); return;
	}
<% End If %>  

	if (cont_next!=1) {
		f.exp_edu_continue.value="0"; 
	}
	f.submit();
}
-->
</script>

<% 
	If Not iExpertID > 0 Then
		Response.Redirect sApplicationHomePath & "login.asp"
 	End If 
%>

</head>
	<body bgcolor="#FFFFFF" topmargin=0 leftmargin=0  marginheight=0 marginwidth=0>
		<% ShowTopMenu %>

		<table width="100%" cellspacing="0" cellpadding="0" border="0" align="center">
			<tbody><tr height="2"><td bgcolor="#97CAFB"></td></tr>
			</tbody>
		</table>

		<div style="margin-left: 10%">
			<h3>Experts Section:</h3>
			<div align="left" style="margin-top: 20px">
				<a href="document.asp?id=<%= iExpertID %>&amp;document=0"
					style="color: #ffffff;
					background: #cc0000;
					padding: 3px 25px;
					text-decoration: none;
					border-radius: 3px;
					border: 1px solid black;" target="_blank">
						Upload document
				</a>
			</div>
			
			<div align="left" style="margin-top: 20px">
				<a href="register.asp?id=<%= iExpertID %>"
					style="color: #ffffff;
					background: #cc0000;
					padding: 3px 25px;
					text-decoration: none;
					border-radius: 3px;
					border: 1px solid black;" target="_blank">
						Update your CV
				</a>
			</div>

			<div align="left" style="margin-top: 20px">
				<a href="view/cv_view.asp?id=<%= iExpertID %>"
					style="color: #ffffff;
					background: #cc0000;
					padding: 3px 25px;
					text-decoration: none;
					border-radius: 3px;
					border: 1px solid black;" target="_blank">
						View and Format your CV
				</a>
			</div>


	</body>
</html>