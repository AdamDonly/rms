<%@ LANGUAGE="VBSCRIPT" %>
<% Option Explicit %>

<!--#include file="../dbc.asp"-->
<!--#include file="../fnc.asp"-->
<!--#include file="../_forms/frmInterface.asp"-->
<!--#include file="../../_common/_data/en/lib.asp"-->
<!--#include file="../_forms/frmScrollBoxNew.asp"-->
<!--#include file="../../_common/register3.asp"-->

<%
Sub BeforeClientValidationCvRegistrationStep3
End Sub


Sub AfterClientValidationCvRegistrationStep3
%>
	if (f.proj_title && f.proj_title.value) {
		if (!checkTextFieldValue(f.exp_ref_firstname, "", "Please specify reference contact person first name.", 1)) { 
			return 
		}
		if (!checkTextFieldValue(f.exp_ref_lastname, "", "Please specify reference contact person last name.", 1)) { 
			return 
		}
		if (!checkTextFieldValue(f.exp_ref_phone, "", "Please specify reference contact person phone.", 1)) { 
			return 
		}
		if (!checkTextFieldValue(f.exp_ref_email, "", "Please specify reference contact person email.", 1)) { 
			return 
		}
	}
<%
End Sub
%>