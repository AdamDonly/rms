<!--#include file="cntColors.asp"-->

<%
Dim sTextFrameColor, sTitleColor


Sub ShowTopMenu
End Sub


Sub ShowMmbLeftMenu(bShowSubItems, iActiveItem)
	Dim sMmbMenuItem, sMmbMenuUrl, sMmbMenuBgImage
	Dim iMmbMenuNumItems, iLoop
	iMmbMenuNumItems=4

	ReDim sMmbMenuItem(iMmbMenuNumItems)
	ReDim sMmbMenuUrl(iMmbMenuNumItems)
	ReDim sMmbMenuBgImage(iMmbMenuNumItems)

	sMmbMenuItem(1)="Daily Tenders Alerts,<br />Contracted & Shortlisted Companies"
	sMmbMenuItem(2)="Consultants Database"
	sMmbMenuItem(3)="Job Posting Board"
	'sMmbMenuItem(4)="assortis Navigator"
	sMmbMenuItem(4)="assortis CVIP"
	           
	sMmbMenuUrl(1)=sHomePath & "en/members/register.asp" & AddUrlParams(sParams,"act=BSC")
	sMmbMenuUrl(2)=sHomePath & "en/members/register.asp" & AddUrlParams(sParams,"act=EXP")
	sMmbMenuUrl(3)=sHomePath & "en/members/register.asp" & AddUrlParams(sParams,"act=JBP")
	'sMmbMenuUrl(4)=sHomePath & "en/members/navigator.asp" & sParams
	sMmbMenuUrl(4)=sHomePath & "en/members/rms.asp" & sParams

	sMmbMenuBgImage(1)=sHomePath & "image/lmnu_mmb_bg230.gif"
	sMmbMenuBgImage(2)=sHomePath & "image/lmnu_mmb_bg230.gif"
	sMmbMenuBgImage(3)=sHomePath & "image/lmnu_mmb_bg230.gif"
	'sMmbMenuBgImage(4)=sHomePath & "image/lmnu_mmb_bg230.gif"
	sMmbMenuBgImage(4)=sHomePath & "image/lmnu_mmb_bg230.gif"
	sMmbMenuBgImage(iActiveItem)=sHomePath & "image/lmnu_mmb_sel230.gif"
%>

  	<table width="230" border="0" cellpadding="0" cellspacing="0">
	<tr><td width="100%" bgcolor="#3366CC"><a href="<%=sHomePath & "en/members/register.asp" & sParams%>"><img src="<%=sHomePath%>image/lmnu_mmb_top230.gif" width="230" height="30" border="0" alt="Services for companies"></a></td></tr>
	<% If bShowSubItems=1 Then %>
	<tr heigth=1><td width="100%" bgcolor="#FFFFFF"><img src="<%=sHomePath%>image/lmnu_mmb_ip_s.gif" width="230" height="1" border="0" alt=""></a></td></tr>
	<tr height=14><td width="100%" bgcolor="#3366CC"><img src="<%=sHomePath%>image/lmnu_mmb_iptab.gif" width="230" height="14" border="0" alt="Information Portal"></a></td></tr>
	<% For iLoop=1 To 3 %>	
	<% If iLoop<>1 Then %><tr><td width="100%"><img src="<%=sHomePath%>image/lmnu_mmb_sp230.gif" width="230" height="<% If iLoop=0 Then %>1<% Else %>3<% End If %>"></td></tr><% End If %>
		<tr><td width="100%" bgcolor="#EBF3FE" background="<%=sMmbMenuBgImage(iLoop)%>"><% If iLoop=1 Then %><a href="<%=sMmbMenuUrl(iLoop)%>"><img src="<%=sHomePath%>image/freetrial/freetrial<% If iActiveItem=1 Then %>b<% End If %>.gif" width=65 height=15 vspace=0 hspace=0 border=0 alt="Free trial" align="right"></a><% End If %><p><img src="<%=sHomePath%>image/x.gif" width=5 height=1 vspace=<% If iLoop=1 Then %>14<% Else %>4<% End If %> align="left"><a class="mmbmenu" href="<%=sMmbMenuUrl(iLoop)%>"><%=sMmbMenuItem(iLoop)%></a></p></td></tr>
	<% Next %>
	<tr><td width="100%"><img src="<%=sHomePath%>image/lmnu_mmb_sp230.gif" width="230" height="<% If iLoop=0 Then %>1<% Else %>3<% End If %>"></td></tr>
	<tr height=14><td width="100%" bgcolor="#3366CC"><img src="<%=sHomePath%>image/lmnu_mmb_mttab.gif" width="230" height="14" border="0" alt="Information Portal"></a></td></tr>
	<% For iLoop=4 To iMmbMenuNumItems %>	
	<% If iLoop<>4 Then %><tr><td width="100%"><img src="<%=sHomePath%>image/lmnu_mmb_sp230.gif" width="230" height="<% If iLoop=0 Then %>1<% Else %>3<% End If %>"></td></tr><% End If %>
		<tr><td width="100%" bgcolor="#EBF3FE" background="<%=sMmbMenuBgImage(iLoop)%>"><p><img src="<%=sHomePath%>image/x.gif" width=5 height=1 vspace=<% If iLoop=1 Then %>14<% Else %>4<% End If %> align="left"><a class="mmbmenu" href="<%=sMmbMenuUrl(iLoop)%>"><%=sMmbMenuItem(iLoop)%></a></p></td></tr>
		<% If iLoop=4 And iActiveItem=5 Then %>
		<tr><td width="100%" bgcolor="#EBF3FE" background="<%=sMmbMenuBgImage(iLoop-1)%>">
			<p class="sml"><img src="<%=sHomePath%>image/x.gif" width=5 height=1 vspace=16 hspace=6 align="left"><a class="bl" href="#" onclick="window.open('elink/elinkscreens/company2.htm', 'company', 'toolbar=1,scrollbars=1,location=0,directories=0,status=1,menubar=1,height=350,resizable=1,width=670')">Company Directory</a>
		</p></td></tr>
		<tr><td width="100%" bgcolor="#EBF3FE" background="<%=sMmbMenuBgImage(iLoop-1)%>">
			<p class="sml"><img src="<%=sHomePath%>image/x.gif" width=5 height=1 vspace=16 hspace=6 align="left"><a class="bl" href="#" onclick="window.open('elink/elinkscreens/contact2.htm', 'contact', 'toolbar=1,scrollbars=1,location=0,directories=0,status=1,menubar=1,height=350,resizable=1,width=670')">Contact Directory</a>
		</p></td></tr>
		<tr><td width="100%" bgcolor="#EBF3FE" background="<%=sMmbMenuBgImage(iLoop-1)%>">
			<p class="sml"><img src="<%=sHomePath%>image/x.gif" width=5 height=1 vspace=16 hspace=6 align="left"><a class="bl" href="#" onclick="window.open('elink/elinkscreens/project2.htm', 'project', 'toolbar=1,scrollbars=1,location=0,directories=0,status=1,menubar=1,height=350,resizable=1,width=670')">Project Manager</a>
		</p></td></tr>
		<tr><td width="100%" bgcolor="#EBF3FE" background="<%=sMmbMenuBgImage(iLoop-1)%>">
			<p class="sml"><img src="<%=sHomePath%>image/x.gif" width=5 height=1 vspace=16 hspace=6 align="left"><a class="bl" href="#" onclick="window.open('elink/elinkscreens/agenda2.htm', 'agenda', 'toolbar=1,scrollbars=1,location=0,directories=0,status=1,menubar=1,height=350,resizable=1,width=670')">Personal/Group Agenda</a>
		</p></td></tr>
		<tr><td width="100%" bgcolor="#EBF3FE" background="<%=sMmbMenuBgImage(iLoop-1)%>">
			<p class="sml"><img src="<%=sHomePath%>image/x.gif" width=5 height=1 vspace=16 hspace=6 align="left"><a class="bl" href="#" onclick="window.open('elink/elinkscreens/todolist2.htm', 'todolist', 'toolbar=1,scrollbars=1,location=0,directories=0,status=1,menubar=1,height=350,resizable=1,width=670')">Personal/Group To-Do List</a>
		</p></td></tr>
		<tr><td width="100%" bgcolor="#EBF3FE" background="<%=sMmbMenuBgImage(iLoop-1)%>">
			<p class="sml"><img src="<%=sHomePath%>image/x.gif" width=5 height=1 vspace=16 hspace=6 align="left"><a class="bl" href="#" onclick="window.open('elink/elinkscreens/document2.htm', 'document', 'toolbar=1,scrollbars=1,location=0,directories=0,status=1,menubar=1,height=350,resizable=1,width=670')">Document Library</a>
		</p></td></tr>
		<tr><td width="100%" bgcolor="#EBF3FE" background="<%=sMmbMenuBgImage(iLoop-1)%>">
			<p class="sml"><img src="<%=sHomePath%>image/x.gif" width=5 height=1 vspace=16 hspace=6 align="left"><a class="bl" href="#" onclick="window.open('elink/elinkscreens/mail.htm', 'mail', 'toolbar=1,scrollbars=1,location=0,directories=0,status=1,menubar=1,height=350,resizable=1,width=670')">Personal Mailbox</a>
		</p></td></tr>
		<% End If %>
	<% Next %>

	<% End If %>
	<tr><td width="100%"><img src="<%=sHomePath%>image/lmnu_mmb_btm230.gif" width="230" height="4"></td></tr>
	</table>
<%
End Sub


Sub ShowExpLeftMenu(bShowSubItems, iActiveItem)
	Dim sExpMenuItem, sExpMenuUrl, sExpMenuBgImage
	Dim iExpMenuNumItems, iLoop
	iExpMenuNumItems=4

	ReDim sExpMenuItem(iExpMenuNumItems)
	ReDim sExpMenuUrl(iExpMenuNumItems)
	ReDim sExpMenuBgImage(iExpMenuNumItems)

	sExpMenuItem(1)="New CV Registration"
	sExpMenuItem(2)="CV Update"
	sExpMenuItem(3)="Special Info Pack"
	sExpMenuItem(4)="Latest Job Offers"

	sExpMenuUrl(1)=sHomePath & "en/experts/register.asp" & AddUrlParams(sParams,"act=RCV")
	sExpMenuUrl(2)=sHomePath & "en/experts/cv_register1.asp" & AddUrlParams(sParams,"act=UCV")
	sExpMenuUrl(3)=sHomePath & "en/experts/register.asp" & AddUrlParams(sParams,"act=SIP")
	sExpMenuUrl(4)=sHomePath & "en/experts/jbp_list.asp" & sParams

	sExpMenuBgImage(1)=sHomePath & "image/lmnu_exp_bg230.gif"
	sExpMenuBgImage(2)=sHomePath & "image/lmnu_exp_bg230.gif"
	sExpMenuBgImage(3)=sHomePath & "image/lmnu_exp_bg230.gif"
	sExpMenuBgImage(4)=sHomePath & "image/lmnu_exp_bg230.gif"
	sExpMenuBgImage(iActiveItem)=sHomePath & "image/lmnu_exp_sel230.gif"
%>
  	<table width="200" border="0" cellpadding="0" cellspacing="0">
	<tr><td width="100%" bgcolor="#CC0000"><a href="<%=sHomePath & "en/experts/register.asp" & sParams%>"><img src="<%=sHomePath%>image/lmnu_exp_top230.gif" width="230" height="30" border="0" alt="Services for experts"></a></td></tr>
	<% If bShowSubItems=1 Then %>
	<tr heigth=1><td width="100%" bgcolor="#FFFFFF"><img src="<%=sHomePath%>image/lmnu_exp_cv_s.gif" width="230" height="1" border="0" alt=""></a></td></tr>
	<tr height=14><td width="100%" bgcolor="#CC0000"><img src="<%=sHomePath%>image/lmnu_exp_cvtab.gif" width="230" height="14" border="0" alt="Curriculum Vitae"></a></td></tr>
	<% For iLoop=1 To 2 %>	
	<% If iLoop<>1 Then %><tr><td width="100%"><img src="<%=sHomePath%>image/lmnu_exp_sp230.gif" width="230" height="<% If iLoop=0 Then %>1<% Else %>3<% End If %>"></td></tr><% End If %>
		<tr><td width="100%" bgcolor="#FFF0E9" background="<%=sExpMenuBgImage(iLoop)%>"><p><img src="<%=sHomePath%>image/x.gif" width=5 height=1 vspace=4 align="left"><a class="expmenu" href="<%=sExpMenuUrl(iLoop)%>"><%=sExpMenuItem(iLoop)%></a></p></td></tr>
	<% Next %>
	<tr><td width="100%"><img src="<%=sHomePath%>image/lmnu_exp_sp230.gif" width="230" height="<% If iLoop=0 Then %>1<% Else %>3<% End If %>"></td></tr>
	<tr height=14><td width="100%" bgcolor="#CC0000"><img src="<%=sHomePath%>image/lmnu_exp_jstab.gif" width="230" height="14" border="0" alt="Jobs Search"></a></td></tr>
	<% For iLoop=3 To iExpMenuNumItems %>	
	<% If iLoop<>3 Then %><tr><td width="100%"><img src="<%=sHomePath%>image/lmnu_exp_sp230.gif" width="230" height="<% If iLoop=0 Then %>1<% Else %>3<% End If %>"></td></tr><% End If %>
		<tr><td width="100%" bgcolor="#FFF0E9" background="<%=sExpMenuBgImage(iLoop)%>"><p><img src="<%=sHomePath%>image/x.gif" width=5 height=1 vspace=4 align="left"><a class="expmenu" href="<%=sExpMenuUrl(iLoop)%>"><%=sExpMenuItem(iLoop)%></a></p></td></tr>
	<% Next %>
	<% End If %>
	<tr><td width="100%"><img src="<%=sHomePath%>image/lmnu_exp_btm230.gif" width="230" height="4"></td></tr>
	</table>

<%
End Sub


Sub ShowBreakLine
%>
	<img src="<%=sHomePath%>image/x.gif" width=1 height=10><br />
<%
End Sub


Sub ShowInputFormHeader(iFormWidth, sFormTitle)
%>
	<div class="box search blue">
	<h3><span class="left">&nbsp;</span><span class="right">&nbsp;</span><% =sFormTitle %></h3>
	<table class="search_form" width="100%" cellspacing=0 cellpadding=0 border=0>
<%
End Sub

Sub ShowInputFormElement(iFormWidth, sElementText, sElementType, sElementName, sElementValue, iTextMaxLength, iWidthPx, bCompulsory, sOtherOptions)
Dim iLeftColumnWidth, iWidthSmb
iLeftColumnWidth=Round(iFormWidth*0.34)
iWidthSmb=Round(iWidthPx/10*1.18)
%>
		<tr class="<% If InStr(sOtherOptions, "last")>0 Then %>last<% End If %>">
		<td class="field splitter"><label for="<% =sElementName %>"><% =sElementText %></label></td>
		<td class="value blue"><input type="<% =sElementType %>" maxlength="<% =iTextMaxLength %>" id="<% =sElementName %>" name="<% =sElementName %>" size="<%=iWidthSmb%>" style="width: <% =iWidthPx %>px;" <% If sElementValue>"" Then %>value="<% =sElementValue %>"<% End If %> <% =sOtherOptions %>>&nbsp;<% If bCompulsory=1 Then %><span class="rs">*</span><% End If %></td>
		</tr>
<%
End Sub

Sub ShowInputFormSpacer(iFormWidth, iSpaceHeight)
Dim iLeftColumnWidth
iLeftColumnWidth=Round(iFormWidth*0.34)
%>
		<tr style="height: <%=iSpaceHeight%>px">
		<td class="field splitter"></td>
		<td class="value blue"></td>
		</tr>
<%
End Sub


Sub ShowInputFormButton(iFormWidth, sButtonTitle, sButtonImagePath)
Dim iLeftColumnWidth
iLeftColumnWidth=Round(iFormWidth*0.34)
%>
	<input type="image" class="button first" src="<% =sButtonImagePath %>" name="<% =sButtonTitle %>" alt="<% =sButtonTitle %>">
<%
End Sub


Sub ShowInputFormFooter(iFormWidth)
%>
	</table>
	</div>
<%
End Sub


Sub SetColorByBlockType(sBlockColor)
	If sBlockColor="ca" Then
		sTextFrameColor="#EEEEEE"
		sTitleColor="#666666"
	ElseIf sBlockColor="sl" Then
		sTextFrameColor="#FFF0E9"
		sTitleColor="#CC0000"
	ElseIf sBlockColor="ex0" Then
		sTextFrameColor="#EBF3FE"
		sTitleColor="#85B6FE"
	ElseIf sBlockColor="ex1" Then
		sTextFrameColor="#FFF0E9"
		sTitleColor="#FFD0BA"
	ElseIf sBlockColor="ex5" Then
		sTextFrameColor="#FCECFF"
		sTitleColor="#F2BEFE"
	Else
		sTextFrameColor="#EBF3FE"
		sTitleColor="#3366CC"
	End If
End Sub



Sub ShowExpertsBlockSubTitle(iWidth, iHeight, sBlockColor)
%>
	<table class="cv" cellpadding="0" cellspacing="0">
<%
End Sub

Sub ShowExpertsBlockFooter(iWidth, iHeight, sBlockColor)
%>	
	</div><br/>
<%
End Sub


Sub ShowUserNoticesBlockHeader(iWidth, iHeight, sBlockType, sBlockDescription, sBlockColor)
	SetColorByBlockType(sBlockColor)
%>
	<div class="box blue notice">
	<h3><span class="left">&nbsp;</span><span class="right">&nbsp;</span><% =sBlockDescription %></h3>
<%
End Sub

Sub ShowUserNoticesBlockFooter(iWidth, iHeight, sBlockColor)
%>	
	</div>
<%
End Sub

Sub ShowUserNoticesViewHeader(iWidthTotal, iWidthLeft)
%>
	<table class="notice" cellpadding="0" cellspacing="0">
<%
End Sub

Sub ShowUserNoticesViewFooter()
%>
	</table>
<%
End Sub

Sub ShowUserNoticesViewText(sTitle, sText)
	If Len(Trim(sText))>0 And sText<>"<b></b>" Then
	%>
	<tr>
	<td class="field splitter"><p><% =sTitle %></p></td>
	<td class="value"><p><% =sText %></p></td>
	</tr>
	<%
	End If
End Sub

Sub ShowUserNoticesViewTextWithStyles(sTitle, sText, sTitleStyle, sTextStyle)
	If Len(Trim(sText))>0 And sText<>"<b></b>" Then
	%>
	<tr>
	<td <% =sTitleStyle %>><p><% =sTitle %></p></td>
	<td <% =sTextStyle %>><p><% =sText %></p></td>
	</tr>
	<%
	End If
End Sub

Sub ShowUserNoticesViewDescription(sText)
	If Len(Trim(sText))>0 And sText<>"<b></b>" Then
	%>
	<tr>
	<td colspan="2" class="value"><p><% =sText %></p></td>
	</tr>
	<%
	End If
End Sub


Sub ShowUserNoticesViewSpacer(iSpaceHeight)
%>
	<tr style="height: <% =iSpaceHeight %>px">
	<td class="field splitter"><img src="<% =sHomePath %>image/x.gif" width="1" height="<% =iSpaceHeight %>"></td>
	<td class="value"></td>
	</tr>
<%
End Sub

Sub ShowUserNoticesViewDelimiter()
%>
<%
End Sub


Sub ShowTableTrSpacer(AStyleClass, AColumns, AHeight)
%>
	<tr class="<% =AStyleClass %>" style="height: <% =AHeight %>px">
	<td class="empty" colspan="<% =AColumns %>"></td>
	</tr>
<%
End Sub

Sub ShowTableTrText(AText, AStyleClass, AColumns, AHeight)
%>
	<tr class="<% =AStyleClass %>" style="height: <% =AHeight %>px">
	<td class="empty" colspan="<% =AColumns %>"><% =AText %></td>
	</tr>
<%
End Sub


Sub ShowMessage(sMessageText, sMessageType, iWidth)
	ShowMessageStart sMessageType, iWidth
	Response.Write sMessageText
	ShowMessageEnd
End Sub

Sub ShowMessageStart(sMessageType, iWidth)
	%>
	<div class="information"<% If iWidth>0 Then %> style="width: <% =iWidth %>px"<% End If %>>
	<%
End Sub

Sub ShowMessageEnd()
%>
	</div>
<%
End Sub


Sub ShowFeatureBoxHeader(sFBoxTitle)
%>
	<div class="box grey gadget">
	<h3><span class="left">&nbsp;</span><span class="right">&nbsp;</span><% =sFBoxTitle %></h3>
<%
End Sub


Sub ShowFeatureBoxDelimiter
%>
<%
End Sub

Sub ShowFeatureBoxFooter
%>
	</div>
<%
End Sub


Sub ShowProgressBar(iTotalBarsNumber, iActiveBarsNumber)
	If IsNumeric(iTotalBarsNumber) And IsNumeric(iActiveBarsNumber) Then
	Dim iBarWidth, iBarHeight, iB
	iBarWidth=iTotalBarsNumber*4-1
	%>
	<!--
	<table width="<%=iBarWidth+2%>" cellpadding=0 cellspacing=0 border=0 align="center"><tr>
	<td width=1 bgcolor="#B5B5B5"><img src="<%=sHomePath%>image/pgb_frame.gif" width=1 height=6></td>
	<td width=<%=iBarWidth%> bgcolor="#EAEAEA" background="<%=sHomePath%>image/pgb_grey.gif">
	<% For iB=1 To iActiveBarsNumber %>
	<img src="<%=sHomePath%>image/pgb_blue.gif" width=3 height=6 align="left" vspace=0 hspace=0><% If iB<iTotalBarsNumber Then %><img src="<%=sHomePath%>image/pgb_grey.gif" width=1 height=6 align="left" vspace=0 hspace=0><% End If %><% Next %>
	</td>
	<td width=1 bgcolor="#B5B5B5"><img src="<%=sHomePath%>image/pgb_frame.gif" width=1 height=6></td>
	</tr></table>
	-->
	<%	
	End If
End Sub


Sub ShowLoginBoxNew
	If (Not Session("UserID")>0) Or sScriptFileName="default.asp" Then
	%>

  	<table width="176" border="0" cellpadding="0" cellspacing="0">
	<tr><td width="100%"><img src="<%=sHomePath%>image/fbox_ttltop.gif" width="176" height="4"></td></tr>
	<tr><td width="100%" bgcolor="#EAEAEA" background="<%=sHomePath%>image/fbox_ttlbg.gif"><p class="fbox" align="center"><b>assortis<span class="rs">.com</span></b> users entry</p></td></tr>
	<tr><td width="100%"><img src="<%=sHomePath%>image/fbox_ttlbtm.gif" width="176" height="4"></td></tr>
	<tr><td width="100%"><img src="<%=sHomePath%>image/fbox_bg.gif" width="176" height="4"></td></tr>
	<tr><td width="100%" background="<%=sHomePath%>image/fbox_bg.gif">
		<table width="170" background="" cellspacing=0 cellpadding=0 border=0>
		<tr><form method="post" action="<%=sHomePath%>login.asp?<%=sScriptFullNameAsParams%>" name="LoginForm" onSubmit="Login();return false;">
		<td width="70"><p class="sml">&nbsp;&nbsp;Login&nbsp;name&nbsp;</td>
		<td width="100" align="right"><input type="text" style="width=90px;" name="login_name" size="11" value="<%=Session("LoginName")%>">&nbsp;&nbsp;</td></tr>
		<tr><td width="70" ><p class="sml">&nbsp;&nbsp;Password&nbsp;</td>
		<td align="right"><input type="password" name="login_pwd" style="width=90px;" size="11">&nbsp;&nbsp;</td></tr>
		<tr><td colspan=2 align="right"><input type="image" src="<%=sHomePath%>image/bte_login.gif" name="Login" border=0 alt="Login" vspace=4 hspace=10></td></tr>
		</form>
		</table>
		<p class="sml" align="right"><a class="bl" href="<%=sHomePath%>en/log_fpwd.asp?<%=sScriptFullNameAsParams%>"><img src="<%=sHomePath%>image/x.gif" width=7 height=1 hspace=1 border=0>Forgot your password ?</a>&nbsp;&nbsp;&nbsp;</p>
	</td></tr>
	<tr><td width="100%"><img src="<%=sHomePath%>image/fbox_bg.gif" width="176" height="4"></td></tr>
	<tr><td width="100%"><img src="<%=sHomePath%>image/fbox_btm.gif" width="176" height="7"></td></tr>
	</table>
	<%
	End If
End Sub

Sub ShowLoginBox
	If Session("UserID")>0 Then
	%>
  	<table width="176" border="0" cellpadding="0" cellspacing="0">
	<tr><td width="100%"><img src="<%=sHomePath%>image/fbox_ttltop.gif" width="176" height="4"></td></tr>
	<tr><td width="100%" bgcolor="#EAEAEA" background="<%=sHomePath%>image/fbox_ttlbg.gif"><p class="fbox" align="center">Welcome&nbsp;<b><%=Session("UserLogin")%></b></p></td></tr>
	<tr><td width="100%"><img src="<%=sHomePath%>image/fbox_ttlbtm.gif" width="176" height="4"></td></tr>
	<tr><td width="100%"><img src="<%=sHomePath%>image/fbox_bg.gif" width="176" height="4"></td></tr>
	<tr><td width="100%" background="<%=sHomePath%>image/fbox_bg.gif">
		<p class="sml">&nbsp;<a href="<%=sHomePath & Session("UserURL")%>">My&nbsp;<b>assortis<span class="rs">.com</span></b>&nbsp;account</a></p>
		<img src="<%=sHomePath%>image/x.gif" width=50 height=5><br />
		<p class="sml">&nbsp;<a class="bl" href="<%=sHomePath & "logout.asp" & sParams%>">Logout</a></p>
	</td></tr>
	<tr><td width="100%"><img src="<%=sHomePath%>image/fbox_bg.gif" width="176" height="4"></td></tr>
	<tr><td width="100%"><img src="<%=sHomePath%>image/fbox_btm.gif" width="176" height="7"></td></tr>
	</table>

	<% Else	%>

  	<table width="176" border="0" cellpadding="0" cellspacing="0">
	<tr><td width="100%"><img src="<%=sHomePath%>image/fbox_ttltop.gif" width="176" height="4"></td></tr>
	<tr><td width="100%" bgcolor="#EAEAEA" background="<%=sHomePath%>image/fbox_ttlbg.gif"><p class="fbox" align="center">Client / Expert Login</p></td></tr>
	<tr><td width="100%"><img src="<%=sHomePath%>image/fbox_ttlbtm.gif" width="176" height="4"></td></tr>
	<tr><td width="100%"><img src="<%=sHomePath%>image/fbox_bg.gif" width="176" height="4"></td></tr>
	<tr><td width="100%" background="<%=sHomePath%>image/fbox_bg.gif">
		<table width="170" background="" cellspacing=0 cellpadding=0 border=0>
		<tr><form method="post" action="<%=sHomePath%>login.asp?<%=sScriptFullNameAsParams%>" name="LoginForm" onSubmit="Login();return false;">
		<td width="70"><p class="sml">&nbsp;&nbsp;User&nbsp;name&nbsp;</td>
		<td width="100" align="left"><input type="text" style="width=90px; margin-top:1px; margin-bottom:1px;" name="login_name" size="11" value="<%=Session("LoginName")%>"></td></tr>
		<tr><td><p class="sml">&nbsp;&nbsp;Password&nbsp;</td>
		<td align="left"><input type="password" name="login_pwd" style="width=90px; margin-top:1px; margin-bottom:1px;" size="11"></td></tr>
		<tr><td colspan=2 align="right"><input type="image" src="<%=sHomePath%>image/bte_login.gif" name="Login" border=0 alt="Login" vspace=4 hspace=10></td></tr>
		</form>
		</table>
		<p class="sml" align="right"><a class="bl" href="<%=sHomePath%>en/log_fpwd.asp?<%=sScriptFullNameAsParams%>"><img src="<%=sHomePath%>image/x.gif" width=7 height=1 hspace=1 border=0>Retrieve your password</a>&nbsp;&nbsp;&nbsp;</p>
	</td></tr>
	<tr><td width="100%"><img src="<%=sHomePath%>image/fbox_bg.gif" width="176" height="4"></td></tr>
	<tr><td width="100%"><img src="<%=sHomePath%>image/fbox_btm.gif" width="176" height="7"></td></tr>
	</table>
	<%
	End If
End Sub


Sub ShowRegistrationProgressBar(sServiceTypeIn, iStep)
Dim arrStepTitle(), arrStepLink(), iLoop, iStepsNumber, sStepStyle, sStepImage, sServiceDescription, sServiceType
sServiceType=Left(sServiceTypeIn, 3)

	If sServiceType="CV" Then
		If sBackOffice>"" Then 
			iStepsNumber=7
		Else
			iStepsNumber=6
		End If

		ReDim arrStepTitle(iStepsNumber)
		arrStepTitle(1)="&nbsp;Personal&nbsp;<br />&nbsp;information&nbsp;"
		arrStepTitle(2)="&nbsp;Education&nbsp;"
		arrStepTitle(3)="&nbsp;Training&nbsp;"
		arrStepTitle(4)="&nbsp;Professional&nbsp;<br />&nbsp;experience&nbsp;"
		arrStepTitle(5)="&nbsp;Languages&nbsp;"
		arrStepTitle(6)="&nbsp;Contact&nbsp;details&nbsp;<br />&nbsp;&amp;&nbsp;availability&nbsp;"

		ReDim arrStepLink(iStepsNumber)
		arrStepLink(1)="<a href=""register.asp" & sParams & """>"
		arrStepLink(2)="<a href=""register2.asp" & sParams & """>"
		arrStepLink(3)="<a href=""register21.asp" & sParams & """>"
		arrStepLink(4)="<a href=""register3.asp" & sParams & """>"
		arrStepLink(5)="<a href=""register4.asp" & sParams & """>"
		arrStepLink(6)="<a href=""register5.asp" & sParams & """>"
		sServiceDescription="Curriculum Vitae online registration"
		If sBackOffice>"" Then 
			arrStepTitle(7)="&nbsp;Review&nbsp;<br />&nbsp;&amp;&nbsp;manage&nbsp;CV&nbsp;"
			arrStepLink(7)="<a href=""cv_register6.asp" & sParams & """>"
		End If
	End If %>

	<!-- The registration progress bar -->

	<% If iStep>0 And sServiceType<>"RCV" And sServiceType<>"UCV" Then %>
	<table width=600 cellspacing=0 cellpadding=0 border=0 <% If Not bIsMyCV Then %> align="center"<% End If %> >
	<tr>
	<% For iLoop=1 To iStepsNumber
		If iStep=iLoop Then 
			sStepStyle="class=""rs"""
			sStepImage="progs_" & LCase(sServiceType) & CStr(iLoop) & "r" %>
		<% Else
			sStepImage="progs_" & LCase(sServiceType) & CStr(iLoop) & "b"
		End If %>
	<% If sServiceType<>"CV" Or iStep=iLoop Then %>
	<td><img src="<%=sHomePath%>image/<% =sStepImage %>.gif" hspace=5 alt="<% =arrStepTitle(iLoop) %>"></td>
	<% End If %>
	<td style="text-align:center; vertical-align:middle;"><p>
		<% If iLoop<2 Or iExpertID>0 Or sBackOffice>"" Then %><% =arrStepLink(iLoop) %><% Else %><u><% End If %>
		<% =arrStepTitle(iLoop) %>
		<% If iLoop<2 Or iExpertID>0 Or sBackOffice>"" Then %></a><% Else %></u><% End If %>
	</p></td>
		<% If iLoop < iStepsNumber Then %>
			<td><img src="<%=sHomePath%>image/progs_arw.gif" width=10 height=42 hspace=5></td>
		<% End If %>
	<% sStepStyle=""
	Next %>
	<td width="5%">&nbsp;</td>
	</tr>
	</table><br /> 

	<% End If %>
<%
End Sub


Sub ShowWaitMessage()
If 1=2 Then
%>
<script language="JavaScript">
document.writeln('<div name="divWait" id="DivWait" align="center"><img src="<%=sHomePath%>image/wait.gif" width=101 height=35></div>');
</script>
<%
End If
End Sub


Sub HideWaitMessage()
If 1=2 Then
%>
<script language="JavaScript">
var is_mzl
   if (is_nav)
	{
	layerStyleRef="layer.";
	layerRef="document.layers";
	styleSwitch="";
	eval(layerRef+'["divWait"]'+styleSwitch+'.visibility="hidden"');
	}
   else
	{if (is_mzl)
	{
	layerRef="document.getElementById";
	styleSwitch=".style";
	eval(layerRef+'("divWait")'+styleSwitch+'.visibility="hidden"');
	}
	else
	{
	layerStyleRef="layer.style.";
	layerRef="document.all";
	styleSwitch=".style";
	eval(layerRef+'["divWait"]'+styleSwitch+'.visibility="hidden"');
	}}
</script>
<%
End If
End Sub


Sub ShowNoticesCalendar(dActiveDate, sUserType)
Dim iActiveYear, iActiveMonth, iActiveDay, dFirstDayOfMonth, dLastDayOfMonth, iFWDayOfMonth, iLDayOfMonth, dPrevMonthLastDay, dNextMonthFirstDay
Dim wd, wk, d, sCalendarUrl
Dim sTempParams
	sCalendarUrl=sScriptFileName
	sTempParams=sParams
	sTempParams=ReplaceUrlParams(sTempParams, "md")

	If sUserType="expert"  Then
		sTempParams=ReplaceUrlParams(sTempParams, "ei=" & iExpertID)
	Else
		sTempParams=ReplaceUrlParams(sTempParams, "mi=" & iMemberID)
		Set objTempRs=GetMemberSubscription(iMemberID, 21)
		If objTempRs("id_AccountStatus")=1 And objTempRs("id_PaymentType")=1 Then
			Dim dTrialStartDate, dTrialEndDate
			dTrialStartDate=DateAdd("d", -7, objTempRs("macStartDate"))
			dTrialEndDate=objTempRs("macExpDate")
			Response.Write("<div class=""content""><p align=""center"">You can view notices from " & ConvertDateForText(dTrialStartDate, "&nbsp;", "DDMMYYYY") & "</p></div>")

		End If
	End If

	If IsDate(dActiveDate) Then
		iActiveDay = Day(dActiveDate) 
		iActiveMonth = Month(dActiveDate) 
		iActiveYear = Year(dActiveDate) 
	Else
		iActiveDay = Day(Date()) 
		iActiveMonth = Month(Date()) 
		iActiveYear = Year(Date()) 
	End If
	dFirstDayOfMonth=DateAdd("d", (-1)*(iActiveDay-1), dActiveDate)
	dLastDayOfMonth=DateAdd("d", -1, DateAdd("m", 1, dFirstDayOfMonth))
	iFWDayOfMonth = Weekday(dFirstDayOfMonth, 2)
	iLDayOfMonth = Day(dLastDayOfMonth)
	dPrevMonthLastDay=DateAdd("d", -1, dFirstDayOfMonth)
	dNextMonthFirstDay=DateAdd("d", 1, dLastDayOfMonth)
%>

<table class="calendar" >
<tr><td valign="top" align="center"><a href="<%=sCalendarUrl & ReplaceUrlParams(sTempParams, "md=" & dPrevMonthLastDay) %>"><img src="<%=sHomePath%>image/ico_prv.gif" width=15 height=13 hspace=0 vspace=2 border=0 alt="Previous month" align="center"></a></td><td valign="center" align="center" <% If sDuration="m1" Then %>bgcolor="#C9DAFC"<% End If %> colspan=5><b><a href="<%=sCalendarUrl & ReplaceUrlParams(sTempParams, "md=" & LCase(Left(MonthName(Month(dMailDate)),3)) & Year(dMailDate)) %>"><%=ConvertDateForText(dMailDate, "&nbsp;", "MonthYear")%></a></b></td><td valign="top" align="center"><a href="<%=sCalendarUrl & ReplaceUrlParams(sTempParams, "md=" & dNextMonthFirstDay )%>"><img src="<%=sHomePath%>image/ico_nxt.gif" width=15 height=13 border=0 hspace=0 vspace=2 alt="Next month" align="center"></td></tr>

<tr>
<th width="14%" class="wd">M</th>
<th width="14%" class="wd">T</th>
<th width="14%" class="wd">W</th>
<th width="14%" class="wd">T</th>
<th width="14%" class="wd">F</th>
<th width="14%" class="wd">S</th>
<th width="14%" class="wd">S</th>
</tr>

<% wd = 1
dim bg(7)
bg(1)="#EAEAEA"
bg(2)="#EAEAEA"
bg(3)="#EAEAEA"
bg(4)="#EAEAEA"
bg(5)="#EAEAEA"
bg(6)="#F2F2F2"
bg(7)="#F2F2F2"
%>
<% For wk=1 To 6 %>
<% If wk<6 Or (wk=6 And wd<=iLDayOfMonth) Then %>
<tr>
<% For d=1 To 7 %>
<td <% If (CInt(iActiveDay)=wd) And (iFWDayOfMonth <=(wk-1)*7+d) And sDuration="d1" Then %>bgcolor="#C9DAFC"<% Else %>bgcolor="<%=bg(d)%>"<% End If %> align="center"><% If (iFWDayOfMonth <=(wk-1)*7+d) And wd<= iLDayOfMonth Then %><a class="darkblue" href="<%=sCalendarUrl & ReplaceUrlParams(sTempParams, "md=" & DateAdd("d", wd-1, dFirstDayOfMonth))%>">
<% 
	If iActiveYear=Year(Now()) And iActiveMonth=Month(Now()) And wd=Day(Now()) Then Response.Write "<b>"

%>
<%=wd%></b>
<%wd=wd+1%></a><% Else %>&nbsp;<% End If %></td>
<% Next %>
</tr>
<% End If %>
<% Next %>
</table>
<%
End Sub


Sub ShowStandardPageHeader() 
%>
<html>
<head>
<title>Assortis.com.</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" type="text/css" href="<%=sHomePath%>styles.css">
</head>

<body bgcolor="#FFFFFF" topmargin=0 leftmargin=0 marginheight=0 marginwidth=0>
<% 
ShowTopMenu
End Sub 


Sub ShowStandardPageFooter() 
%>
</body></html>
<%
End Sub 



Sub ShowInfoBlockHeader(iWidthMax, iWidthMin, iHeight, sTitle, sColorScheme) 
Dim sTitleColor
sTitleColor="#85B6FE"
%>
	<table width=<%=iWidthMax%> cellspacing=0 cellpadding=0 border=0>
	<tr height="3">
	<td width="22" bgcolor="#FFFFFF"><img src="<%=sHomePath%>image/bx_<%=sColorScheme%>_01.gif" width=22 height=3 hspace=0 vspace=0 alt=""></td>
  	<td width="100%" bgcolor="#FFFFFF" background="<%=sHomePath%>image/bx_<%=sColorScheme%>_02.gif"><img src="<%=sHomePath%>image/bx_<%=sColorScheme%>_02.gif" width=<%=iWidthMin%> height=3></td>
	<td width="5" bgcolor="#FFFFFF" background="<%=sHomePath%>image/bx_<%=sColorScheme%>_03.gif"><img src="<%=sHomePath%>image/x.gif" width=5 height=3 hspace=0 vspace=0></td>
	</tr>
	<tr height="26">
	<td width="22" bgcolor="#FFFFFF" background="<%=sHomePath%>image/bx_<%=sColorScheme%>_11.gif"><img src="<%=sHomePath%>image/x.gif" width=22 height=26 hspace=0 vspace=0 alt=""></td>
  	<td width="100%" bgcolor="<%=sTitleColor%>" valign="center"><p class="nltxt0"><b><%=sTitle%></p></td>
	<td width="5" bgcolor="#FFFFFF" background="<%=sHomePath%>image/bx_<%=sColorScheme%>_13.gif"><img src="<%=sHomePath%>image/x.gif" width=5 height=26 hspace=0 vspace=0></td>
	</tr>
	<tr height="3">
	<td width="22" bgcolor="#FFFFFF" background="<%=sHomePath%>image/bx_<%=sColorScheme%>_21.gif"><img src="<%=sHomePath%>image/x.gif" width=22 height=3 hspace=0 vspace=0 alt=""></td>
  	<td width="100%" bgcolor="#FFFFFF" background="<%=sHomePath%>image/bx_<%=sColorScheme%>_22.gif" valign="center"><img src="<%=sHomePath%>image/x.gif" width=2 height=3 align="left"></td>
	<td width="5" bgcolor="#FFFFFF" background="<%=sHomePath%>image/bx_<%=sColorScheme%>_23.gif"><img src="<%=sHomePath%>image/x.gif" width=5 height=3 hspace=0 vspace=0></td>
	</tr>
	</table>

	<table width=<%=iWidthMax%> cellspacing=0 cellpadding=0 border=0>
	<tr height="3">
	<td width="2" bgcolor="#FFFFFF" background="<%=sHomePath%>image/bx_<%=sColorScheme%>_31.gif"><img src="<%=sHomePath%>image/x.gif" width=2 height=3 hspace=0 vspace=0 alt=""></td>
  	<td width="100%" bgcolor="#FFFFFF" valign="center"><img src="<%=sHomePath%>image/x.gif" width=2 height=3 align="left">

<%
End Sub


Sub ShowInfoBlockFooter(iWidthMax, sColorScheme)
%>
	</td>
	<td width="5" bgcolor="#FFFFFF" background="<%=sHomePath%>image/bx_<%=sColorScheme%>_33.gif"><img src="<%=sHomePath%>image/x.gif" width=5 height=3 hspace=0 vspace=0></td>
	</tr>
	</table>

	<table width=<%=iWidthMax%> cellspacing=0 cellpadding=0 border=0>
	<tr height="4">
	<td width="22" bgcolor="#FFFFFF" background="<%=sHomePath%>image/bx_<%=sColorScheme%>_41.gif"><img src="<%=sHomePath%>image/x.gif" width=22 height=4 hspace=0 vspace=0></td>
  	<td width="99%" bgcolor="#FFFFFF" background="<%=sHomePath%>image/bx_<%=sColorScheme%>_42.gif" valign="center"><img src="<%=sHomePath%>image/x.gif" width=2 height=4></td>
	<td width="5" bgcolor="#FFFFFF" background="<%=sHomePath%>image/bx_<%=sColorScheme%>_43.gif"><img src="<%=sHomePath%>image/x.gif" width=5 height=4 hspace=0 vspace=0></td>
	</tr>
	</table>
<%
End Sub


Function GetInformerBox(sText, sColorType)
	GetInformerBox="<table align=""right"" border=0 width=210 cellpadding=0 cellspacing=0><tr><td width=""100%"" bgcolor=""#3366CC""><table width=210 cellpadding=2 cellspacing=1 border=0><tr><td bgcolor=""#EBF3FE""><p align=""center"">" & sText & "</p></td></tr></table></td></tr></table>"
End Function
%>


