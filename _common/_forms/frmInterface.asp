<!--#include file="cntColors.asp"-->
<%
Dim sTextFrameColor, sTitleColor

Dim sFieldCompulsoryMark
sFieldCompulsoryMark="&nbsp;&nbsp;<span class=""fcmp"">*</span>&nbsp;"

Sub InputFormHeader(iFormWidth, sFormTitle)
%>
	<table width="<%=iFormWidth%>" cellspacing="0" cellpadding="0" border="0" align="center">
	<tr height=1><td width=1 bgcolor="#97CAFB"><img src="<%=sHomePath%>image/x.gif" width=1 height=1></td>
	<td width="<%=iFormWidth-2%>" bgcolor="#97CAFB"><img src="<%=sHomePath%>image/x.gif" width="<%=iFormWidth-2%>" height=1></td>
	<td width=1 bgcolor="#153B80"><img src="<%=sHomePath%>image/x.gif" width=1 height=1></td></tr>
	<tr height=18><td width=1 bgcolor="#97CAFB"><img src="<%=sHomePath%>image/x.gif" width=1 height=18></td>
	<td width="<%=iFormWidth-2%>" bgcolor="#3366CC"><img src="<%=sHomePath%>image/x.gif" width="<%=iFormWidth-2%>" height=1><br><p class="fttl"><img src="<%=sHomePath%>image/<% =imgFormBullet %>" width=7 height=7 align=left vspace=3 hspace=8><%=sFormTitle%></p></td>
	<td width=1 bgcolor="#153B80"><img src="<%=sHomePath%>image/x.gif" width=1 height=18></td></tr>
	<%
	InputFormDualLine()
	InputFormBeforeBlock(iFormWidth)
End Sub

Sub InputFormFooter()
	InputFormAfterBlock
	InputFormDualLine
	%>
	</table>
<%
End Sub
	
Sub InputFormBeforeBlock(iFormWidth)
%>
	<tr><td width=1 bgcolor="<%=colFormBodyText%>"><img src="<%=sHomePath%>image/x.gif" width=1 height=1></td>
	<td width="<%=iFormWidth-2%>" bgcolor="<%=colFormHeaderTop%>" valign="top">
<%
End Sub

Sub InputFormAfterBlock()
%>
	</td>
	<td width=1 bgcolor="<% =colFormBodyRight %>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td></tr>
<%
End Sub

Sub InputFormDualLine()
%>
	<tr height=1><td width=1 bgcolor="<% =colFormBodyLeft %>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td>
	<td width=578 bgcolor="<% =colFormHeaderMain %>"><img src="<% =sHomePath %>image/x.gif" width=498 height=1></td>
	<td width=1 bgcolor="<% =colFormBodyRight %>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td></tr>
	<tr height=1><td width=1 bgcolor="<% =colFormHeaderSplitter %>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td>
	<td width=578 bgcolor="<% =colFormHeaderSplitter %>"><img src="<% =sHomePath %>image/x.gif" width=498 height=1></td>
	<td width=1 bgcolor="<% =colFormHeaderSplitter %>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td></tr>
<%
End Sub

Sub InputFormSpace(iSpaceHeight)
%>
	<img src="<%=sHomePath%>image/x.gif" width="1" height="<% =iSpaceHeight %>"><br>
<%
End Sub

Sub InputBlockHeader(iBlockWidth)
%>
	<table width="100%" cellspacing="0" cellpadding="0" border="0">
<%
End Sub

Sub InputBlockFooter
%>
	</table>
<%
End Sub

Sub InputBlockSpace(iSpaceHeight)
%>
		<tr><td width=170></td>
		<td bgcolor="<% =colFormBodyRight %>" width="1"><img src="<%=sHomePath%>image/x.gif" width="1" height="<% =iSpaceHeight %>"></td>
		<td bgcolor="<% =colFormBodyText %>" width=407></td></tr>
<%
End Sub

Sub InputBlockElementLeftStart
%><tr><td width="170" valign="top"><%
End Sub

Sub InputBlockElementLeftEnd
%></td><%
End Sub

Sub InputBlockElementMiddle
%><td bgcolor="<% =colFormBodyRight %>" width="1"><img src="x.gif" width="1" height="24"></td><%
End Sub

Sub InputBlockElementRightStart
%><td bgcolor="<% =colFormBodyText %>" width="407"><%
End Sub

Sub InputBlockElementRightEnd
%></td></tr><%
End Sub

Sub InputBlockElement(iFormWidth, sElementText, sElementType, sElementName, sElementValue, iTextMaxLength, iWidthPx, bCompulsory, sOtherOptions)
Dim iLeftColumnWidth, iWidthSmb
iLeftColumnWidth=Round(iFormWidth*0.34)
iWidthSmb=Round(iWidthPx/10*1.18)
%>
		<tr><td width="<%=iLeftColumnWidth%>"><p class="ftxt"><%=sElementText%></p></td>
		<td bgcolor="<% =colFormBodyRight %>" width="1"><img src="<%=sHomePath%>image/x.gif" width="1" height="24"></td>
		<td bgcolor="<% =colFormBodyText %>" width="<%=iFormWidth-iLeftColumnWidth%>">&nbsp;&nbsp;<input type="<%=sElementType%>" maxlength="<%=iTextMaxLength%>" name="<%=sElementName%>" size="<%=iWidthSmb%>" style="width:<%=iWidthPx%>px;" <% If sElementValue>"" Then %>value="<%=sElementValue%>"<% End If %> <%=sOtherOptions%>>&nbsp;<% If bCompulsory=1 Then %><span class="rs">*</span><% End If %></td></tr>
<%
End Sub


Sub ShowBreakLine
%>
	<img src="<%=sHomePath%>image/x.gif" width=1 height=10><br>
<%
End Sub


Sub ShowInputFormHeader(iFormWidth, sFormTitle)
%>
	<table width="<%=iFormWidth%>" cellspacing="0" cellpadding="0" border="0" align="center">
	<tr height=1><td width=1 bgcolor="#97CAFB"><img src="<%=sHomePath%>image/x.gif" width=1 height=1></td>
	<td width="<%=iFormWidth-2%>" bgcolor="#97CAFB"><img src="<%=sHomePath%>image/x.gif" width="<%=iFormWidth-2%>" height=1></td>
	<td width=1 bgcolor="#153B80"><img src="<%=sHomePath%>image/x.gif" width=1 height=1></td></tr>

	<tr height=18><td width=1 bgcolor="#97CAFB"><img src="<%=sHomePath%>image/x.gif" width=1 height=18></td>
	<td width="<%=iFormWidth-2%>" bgcolor="#3366CC"><img src="<%=sHomePath%>image/x.gif" width="<%=iFormWidth-2%>" height=1><br><p class="fttl"><img src="<%=sHomePath%>image/<% =imgFormBullet %>" width=7 height=7 align=left vspace=3 hspace=8><%=sFormTitle%></p></td>
	<td width=1 bgcolor="#153B80"><img src="<%=sHomePath%>image/x.gif" width=1 height=18></td></tr>

	<tr height=1><td width=1 bgcolor="#97CAFB"><img src="<%=sHomePath%>image/x.gif" width=1 height=1></td>
	<td width="<%=iFormWidth-2%>" bgcolor="#153B80"><img src="<%=sHomePath%>image/x.gif" width="<%=iFormWidth-2%>" height=1></td>
	<td width=1 bgcolor="#153B80"><img src="<%=sHomePath%>image/x.gif" width=1 height=1></td></tr>

	<tr height=1><td width=1 bgcolor="#FFFFFF"><img src="<%=sHomePath%>image/x.gif" width=1 height=1></td>
	<td width="<%=iFormWidth-2%>" bgcolor="#FFFFFF"><img src="<%=sHomePath%>image/x.gif" width="<%=iFormWidth-2%>" height=1></td>
	<td width=1 bgcolor="#FFFFFF"><img src="<%=sHomePath%>image/x.gif" width=1 height=1></td></tr>

	<tr><td width=1 bgcolor="#D1E7F7"><img src="<%=sHomePath%>image/x.gif" width=1 height=1></td>
	<td width="<%=iFormWidth-2%>" bgcolor="#97CAFB" valign="top">
		<table width="100%" cellspacing=0 cellpadding=0 border=0>
<%
End Sub


Sub ShowInputFormElement(iFormWidth, sElementText, sElementType, sElementName, sElementValue, iTextMaxLength, iWidthPx, bCompulsory, sOtherOptions)
Dim iLeftColumnWidth, iWidthSmb
iLeftColumnWidth=Round(iFormWidth*0.34)
iWidthSmb=Round(iWidthPx/10*1.18)
%>
		<tr><td width="<%=iLeftColumnWidth%>"><p class="ftxt"><%=sElementText%></p></td>
		<td bgcolor="#4694E1" width=1><img src="<%=sHomePath%>image/x.gif" width=1 height=26></td>
		<td bgcolor="#E0F3FF" width=<%=iFormWidth-iLeftColumnWidth%>>&nbsp;&nbsp;<input type="<%=sElementType%>" maxlength="<%=iTextMaxLength%>" name="<%=sElementName%>" size="<%=iWidthSmb%>" style="width:<%=iWidthPx%>px;" <% If sElementValue>"" Then %>value="<%=sElementValue%>"<% End If %> <%=sOtherOptions%>>&nbsp;<% If bCompulsory=1 Then %><span class="rs">*</span><% End If %></td></tr>
<%
End Sub


Sub ShowInputFormSpacer(iFormWidth, iSpaceHeight)
Dim iLeftColumnWidth
iLeftColumnWidth=Round(iFormWidth*0.34)
%>
		<tr><td width="<%=iLeftColumnWidth%>"></td>
		<td bgcolor="#4694E1" width=1><img src="<%=sHomePath%>image/x.gif" width=1 height=<%=iSpaceHeight%>></td>
		<td bgcolor="#E0F3FF" width=<%=iFormWidth-iLeftColumnWidth%>></td></tr>
<%
End Sub


Sub ShowInputFormButton(iFormWidth, sButtonTitle, sButtonImagePath)
Dim iLeftColumnWidth
iLeftColumnWidth=Round(iFormWidth*0.34)
%>
		</table>
	</td>
	<td width=1 bgcolor="#4694E1"><img src="<%=sHomePath%>image/x.gif" width=1 height=1></td></tr>

	<tr height=1><td width=1 bgcolor="#D1E7F7"><img src="<%=sHomePath%>image/x.gif" width=1 height=1></td>
	<td width=<%=iFormWidth-2%> bgcolor="#3366CC"><img src="<%=sHomePath%>image/x.gif" width=<%=iFormWidth-2%> height=1></td>
	<td width=1 bgcolor="#4694E1"><img src="<%=sHomePath%>image/x.gif" width=1 height=1></td></tr>
	<tr height=1><td width=1 bgcolor="#FFFFFF"><img src="<%=sHomePath%>image/x.gif" width=1 height=1></td>
	<td width=<%=iFormWidth-2%> bgcolor="#FFFFFF"><img src="<%=sHomePath%>image/x.gif" width=<%=iFormWidth-2%> height=1></td>
	<td width=1 bgcolor="#FFFFFF"><img src="<%=sHomePath%>image/x.gif" width=1 height=1></td></tr>

	<tr><td width=1 bgcolor="#D1E7F7"><img src="<%=sHomePath%>image/x.gif" width=1 height=1></td>
	<td width=<%=iFormWidth-2%> bgcolor="#97CAFB" valign="top">
		<table border=0 cellspacing=0 cellpadding=0 width="100%">
		<tr height=25><td width=<%=iLeftColumnWidth%>>&nbsp;</td>
		<td width=<%=iFormWidth-iLeftColumnWidth%> valign="center"><input type="image" src="<%=sButtonImagePath%>" name="<%=sButtonTitle%>" border=0 alt="<%=sButtonTitle%>" vspace=4 align="left"></td>
		</tr>
<%
End Sub


Sub ShowInputFormFooter(iFormWidth)
%>
		</table>
	</td>
	<td width=1 bgcolor="#4694E1"><img src="<%=sHomePath%>image/x.gif" width=1 height=1></td></tr>

	<tr height=1><td width=1 bgcolor="#D1E7F7"><img src="<%=sHomePath%>image/x.gif" width=1 height=1></td>
	<td width=348 bgcolor="#3366CC"><img src="<%=sHomePath%>image/x.gif" width=348 height=1></td>
	<td width=1 bgcolor="#4694E1"><img src="<%=sHomePath%>image/x.gif" width=1 height=1></td></tr>
	<tr height=1><td width=1 bgcolor="#FFFFFF"><img src="<%=sHomePath%>image/x.gif" width=1 height=1></td>
	<td width=348 bgcolor="#FFFFFF"><img src="<%=sHomePath%>image/x.gif" width=348 height=1></td>
	<td width=1 bgcolor="#FFFFFF"><img src="<%=sHomePath%>image/x.gif" width=1 height=1></td></tr>
	</table>
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
	SetColorByBlockType(sBlockColor)
	%>	

	<table width=<%=iWidth%> cellspacing=0 cellpadding=0 border=0 align="center">
	<tr height="2">
	<td width="6"><img src="<%=sHomePath%>image/pmmb_<%=sBlockColor%>dlm_bg0.gif" width="6" height="2" hspace=0 vspace=0></td>
	<td width="100%" background="<%=sHomePath%>image/pmmb_<%=sBlockColor%>dlm_bg1.gif"><img src="<%=sHomePath%>image/x.gif" width="2" height="2"></td>
	<td width="6"><img src="<%=sHomePath%>image/pmmb_<%=sBlockColor%>dlm_bg2.gif" width="6" height="2" hspace=0 vspace=0></td>
	</tr>
	</table>

	<table width=<%=iWidth%> cellspacing=0 cellpadding=0 border=0 align="center">
	<tr height="<%=iHeight-37%>">
	<td width="1" bgcolor="<%=sTextFrameColor%>" background="<%=sHomePath%>image/pmmb_<%=sBlockColor%>lft_bg1.gif"><img src="<%=sHomePath%>image/x.gif" width="1" height=<%=iHeight-37%> hspace="1" vspace="0">
	<td bgcolor="<%=sTextFrameColor%>" >
<%
End Sub

Sub ShowExpertsBlockFooter(iWidth, iHeight, sBlockColor)
%>	
	</td>
	<td width="3" bgcolor="#EBF3FE" background="<%=sHomePath%>image/pmmb_<%=sBlockColor%>rht_bg1.gif"><img src="<%=sHomePath%>image/pmmb_<%=sBlockColor%>rht_bg1.gif" width="3" height="<%=iHeight-37%>" hspace="0" vspace="0"></td>
	</tr>
	</table>

	<table width=<%=iWidth%> cellspacing=0 cellpadding=0 border=0 align="center">
	<tr height="4">
	<td width="6"><img src="<%=sHomePath%>image/pmmb_<%=sBlockColor%>btm_bg0.gif" width="6" height="4" hspace=0 vspace=0></td>
	<td width="100%" background="<%=sHomePath%>image/pmmb_<%=sBlockColor%>btm_bg1.gif"><img src="<%=sHomePath%>image/x.gif" width="2" height="4"></td>
	<td width="6"><img src="<%=sHomePath%>image/pmmb_<%=sBlockColor%>btm_bg2.gif" width="6" height="4" hspace=0 vspace=0></td>
	</tr>
	</table>
	<%
	ShowBreakLine
End Sub


Sub ShowUserNoticesBlockHeader(iWidth, iHeight, sBlockType, sBlockDescription, sBlockColor)
	SetColorByBlockType(sBlockColor)
	%>	
	<table width=<%=iWidth%> cellspacing=0 cellpadding=0 border=0>
	<tr height="32">
	<td width="254" bgcolor="<%=sTitleColor%>"><img src="<%=sHomePath%>image/pmmb_<%=sBlockType%>_t1e.gif" width=254 height=32 hspace=0 vspace=0 alt="<%=sBlockDescription%>"></td>
  	<td width="99%" bgcolor="<%=sTitleColor%>" background="<%=sHomePath%>image/pmmb_<%=sBlockColor%>top_bg1.gif" valign="center"><img src="<%=sHomePath%>image/x.gif" width=2 height=32 align="left"><% If sBlockType="ca" Or sBlockType="sl" Then %><img src="<%=sHomePath%>image/x.gif" width=1 height=8><br><p class="sml" align="right"><a href="#top" class="fttl">Top</a>&nbsp;&nbsp;</p><% End If %></td>
	<td width="6" bgcolor="<%=sTitleColor%>"><img src="<%=sHomePath%>image/pmmb_<%=sBlockColor%>top_bg2.gif" width=6 height=32 hspace=0 vspace=0></td>
	</tr>
	</table>

	<table width=<%=iWidth%> cellspacing=0 cellpadding=0 border=0>
	<tr height="<%=iHeight-37%>">
	<td bgcolor="<%=sTextFrameColor%>" background="<%=sHomePath%>image/pmmb_<%=sBlockColor%>lft_bg1.gif"><img src="<%=sHomePath%>image/x.gif" width="1" height=<%=iHeight-37%> hspace=0 vspace=0 align="left">

<%
End Sub


Sub ShowUserNoticesBlockSecondHeader(iWidth, iHeight, sBlockColor)
%>	

	<table width=<%=iWidth%> cellspacing=0 cellpadding=0 border=0>
	<tr height="3">
	<td width="6"><img src="<%=sHomePath%>image/pmmb_<%=sBlockColor%>dlm_bg0.gif" width="6" height="3" hspace=0 vspace=0></td>
	<td width="100%" background="<%=sHomePath%>image/pmmb_<%=sBlockColor%>dlm_bg1.gif"><img src="<%=sHomePath%>image/x.gif" width="2" height="3"></td>
	<td width="6"><img src="<%=sHomePath%>image/pmmb_<%=sBlockColor%>dlm_bg2.gif" width="6" height="3" hspace=0 vspace=0></td>
	</tr>
	</table>

	<table width=<%=iWidth%> cellspacing=0 cellpadding=0 border=0>
	<tr height="<%=iHeight-37%>">
	<td bgcolor="#FFFFFF" background="<%=sHomePath%>image/pmmb_<%=sBlockColor%>lft_bg1.gif"><img src="<%=sHomePath%>image/x.gif" width="2" height=<%=iHeight-37%> hspace=0 vspace=0 align="left">

<%
End Sub


Sub ShowUserNoticesBlockFooter(iWidth, iHeight, sBlockColor)
%>	
	</td>
	<td width="3" bgcolor="#EBF3FE" background="<%=sHomePath%>image/pmmb_<%=sBlockColor%>rht_bg1.gif"><img src="<%=sHomePath%>image/pmmb_<%=sBlockColor%>rht_bg1.gif" width="3" height="<%=iHeight-37%>" hspace="0" vspace="0"></td>
	</tr>
	</table>

	<table width=<%=iWidth%> cellspacing=0 cellpadding=0 border=0>
	<tr height="4">
	<td width="254"><img src="<%=sHomePath%>image/pmmb_<%=sBlockColor%>btm_bg0.gif" width="254" height="4" hspace=0 vspace=0></td>
	<td width="99%" background="<%=sHomePath%>image/pmmb_<%=sBlockColor%>btm_bg1.gif"><img src="<%=sHomePath%>image/x.gif" width="2" height="4"></td>
	<td width="6"><img src="<%=sHomePath%>image/pmmb_<%=sBlockColor%>btm_bg2.gif" width="6" height="4" hspace=0 vspace=0></td>
	</tr>
	</table>
	<%
	ShowBreakLine
End Sub


Sub ShowUserNoticesViewHeader(iWidthTotal, iWidthLeft)
%>
	<table cellspacing=0 cellpadding=0 align="center" width="<%=iWidthTotal%>" border=0>
	<tr>
	<td bgcolor="<%=sTextFrameColor%>" valign="top"><img src="image/x.gif" width="<%=iWidthLeft%>" height="1"><br></td>
	<td width=1 bgcolor="<%=sTitleColor%>"><img src="image/x.gif" width="1" height="1"><br></td>
	<td bgcolor="#FFFFFF" align="left" valign="top"><img src="image/x.gif" width="1" height="1"><br></td>         
	</tr>
<%
End Sub


Sub ShowUserNoticesViewText(sTitle, sText)
	If Len(Trim(sText))>0 And sText<>"<b></b>" Then
	%>
	<tr>
	<td bgcolor="<%=sTextFrameColor%>" valign="top"><p class="txt"><%=sTitle%></p></td>
	<td width=1 bgcolor="<%=sTitleColor%>"><img src="image/x.gif" width=1 height=1><br></td>
	<td width="85%" bgcolor="#FFFFFF" align="left" valign="top"><p class="txt"><%=sText%></p></td>
	</tr>
	<%
	End If
End Sub


Sub ShowUserNoticesViewSpacer(iSpaceHeight)
%>
	<tr>
	<td bgcolor="<%=sTextFrameColor%>" valign="top"><img src="image/x.gif" width="1" height="<%=iSpaceHeight%>"><br></td>
	<td width=1 bgcolor="<%=sTitleColor%>"><img src="image/x.gif" width="1" height="<%=iSpaceHeight%>"><br></td>
	<td bgcolor="#FFFFFF" align="left" valign="top"><img src="image/x.gif" width="1" height="<%=iSpaceHeight%>"><br></td>         
	</tr>
<%
End Sub


Sub ShowUserNoticesViewDelimiter()
%>
	<tr>
	<td bgcolor="<%=sTitleColor%>" valign="top"><img src="image/x.gif" width="1" height="1"><br></td>
	<td width=1 bgcolor="<%=sTitleColor%>"><img src="image/x.gif" width="1" height="1"><br></td>
	<td bgcolor="<%=sTitleColor%>" align="left" valign="top"><img src="image/x.gif" width="1" height="1"><br></td>         
	</tr>
<%
End Sub


Sub ShowUserNoticesViewFooter()
%>
	</table>
<%
End Sub


Sub ShowMessage(sMessageText, sMessageType, iWidth)

	If sMessageType="error" Then
		HideWaitMessage
	End If
	%>
	<img src="<%=sHomePath%>image/x.gif" width=1 height=7><br>
	<table width=<%=iWidth%> cellspacing=0 cellpadding=0 border=0 align="center">
	<tr><td width=30 valign="top">
	<img src="<%=sHomePath%>image/<%=Left(sMessageType,1)%>.gif" width=18 height=18 hspace=10 vspace=3 alt="!" align="left">
	</td>
	<td width=<%=iWidth-30%> valign="center"><p class="txt">
	<% If sMessageType="info" Then %>
		<% Response.Write sMessageText %>
	<% ElseIf sMessageType="error" Then %>
		<span class="rs"><% Response.Write sMessageText %></span>
	<% End If %>
	</p></td>
	</tr></table>
	<img src="<%=sHomePath%>image/x.gif" width=1 height=7><br>
<%
End Sub


Sub ShowMessageStart(sMessageType, iWidth)

	If sMessageType="error" Then
		HideWaitMessage
	End If
	%>
	<img src="<%=sHomePath%>image/x.gif" width=1 height=7><br>
	<table width=<%=iWidth%> cellspacing=0 cellpadding=0 border=0 align="center">
	<tr><td width=30 valign="top">
	<img src="<%=sHomePath%>image/<%=Left(sMessageType,1)%>.gif" width=18 height=18 hspace=10 vspace=3 alt="!" align="left">
	</td>
	<td width=<%=iWidth-30%> valign="center"><p class="txt">
	<% If sMessageType="error" Then %>
	<span class="rs">
	<%
	End If
End Sub


Sub ShowMessageEnd()
%>
	</span>
	</p></td>
	</tr></table>
	<img src="<%=sHomePath%>image/x.gif" width=1 height=7><br>
<%
End Sub


Sub ShowFeatureBoxHeader(sFBoxTitle)
%>
  	<table width="176" border="0" cellpadding="0" cellspacing="0">
	<tr><td width="100%"><img src="<%=sHomePath%>image/fbox_ttltop.gif" width="176" height="4"></td></tr>
	<tr><td width="100%" bgcolor="#EAEAEA" background="<%=sHomePath%>image/fbox_ttlbg.gif"><p class="fbox" align="center"><% Response.Write sFBoxTitle %></p></td></tr>
	<tr><td width="100%"><img src="<%=sHomePath%>image/fbox_ttlbtm.gif" width="176" height="4"></td></tr>
	<tr><td width="100%"><img src="<%=sHomePath%>image/fbox_bg.gif" width="176" height="4"></td></tr>
	<tr><td width="100%" background="<%=sHomePath%>image/fbox_bg.gif">
<%
End Sub


Sub ShowFeatureBoxDelimiter
%>
	</td></tr>
	<tr><td width="100%"><img src="<%=sHomePath%>image/fbox_bg.gif" width="176" height="3"></td></tr>
	<tr><td width="100%"><img src="<%=sHomePath%>image/fbox_dlm.gif" width="176" height="5"></td></tr>
	<tr><td width="100%"><img src="<%=sHomePath%>image/fbox_bg.gif" width="176" height="3"></td></tr>
	<tr><td width="100%" background="<%=sHomePath%>image/fbox_bg.gif">
<%
End Sub


Sub ShowFeatureBoxFooter
%>
	</td></tr>
	<tr><td width="100%"><img src="<%=sHomePath%>image/fbox_bg.gif" width="176" height="4"></td></tr>
	<tr><td width="100%"><img src="<%=sHomePath%>image/fbox_btm.gif" width="176" height="7"></td></tr>
	</table>
<%
End Sub

Sub ShowFeatureBoxFooterWithFormFooter
%>
	</td></tr>
	<tr><td width="100%"><img src="<%=sHomePath%>image/fbox_bg.gif" width="176" height="4"></td></tr>
	<tr><td width="100%"><img src="<%=sHomePath%>image/fbox_btm.gif" width="176" height="7"></td></tr>
	</form>
	</table>
<%
End Sub

Sub ShowProgressBar(iTotalBarsNumber, iActiveBarsNumber)
	If IsNumeric(iTotalBarsNumber) And IsNumeric(iActiveBarsNumber) Then
	Dim iBarWidth, iBarHeight, iB
	iBarWidth=iTotalBarsNumber*4-1
	%>
	<img src="<%=sHomePath%>image/x.gif" width=1 height=3 vspace=0 hspace=0><br>
	<img src="<%=sHomePath%>image/x.gif" width=1 height=1 vspace=0 hspace=0 align="right">
	<table width="<%=iBarWidth+2%>" cellpadding=0 cellspacing=0 border=0 align="center"><tr>
	<td width=1 bgcolor="#B5B5B5"><img src="<%=sHomePath%>image/pgb_frame.gif" width=1 height=6></td>
	<td width=<%=iBarWidth%> bgcolor="#EAEAEA" background="<%=sHomePath%>image/pgb_grey.gif">
	<% For iB=1 To iActiveBarsNumber %>
	<img src="<%=sHomePath%>image/pgb_blue.gif" width=3 height=6 align="left" vspace=0 hspace=0><% If iB<iTotalBarsNumber Then %><img src="<%=sHomePath%>image/pgb_grey.gif" width=1 height=6 align="left" vspace=0 hspace=0><% End If %><% Next %>
	</td>
	<td width=1 bgcolor="#B5B5B5"><img src="<%=sHomePath%>image/pgb_frame.gif" width=1 height=6></td>
	</tr></table>
	<img src="<%=sHomePath%>image/x.gif" width=1 height=1 vspace=0 hspace=0><br>
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
		<td width="100" align="right"><input type="text" style="width:90px;" name="login_name" size="11" value="<%=Session("LoginName")%>">&nbsp;&nbsp;</td></tr>
		<tr><td width="70" ><p class="sml">&nbsp;&nbsp;Password&nbsp;</td>
		<td align="right"><input type="password" name="login_pwd" style="width:90px;" size="11">&nbsp;&nbsp;</td></tr>
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
	<tr><td width="100%" bgcolor="#EAEAEA" background="<%=sHomePath%>image/fbox_ttlbg.gif"><p class="fbox" align="center">Welcome&nbsp;<b><%=Session("UserName")%></b></p></td></tr>
	<tr><td width="100%"><img src="<%=sHomePath%>image/fbox_ttlbtm.gif" width="176" height="4"></td></tr>
	<tr><td width="100%"><img src="<%=sHomePath%>image/fbox_bg.gif" width="176" height="4"></td></tr>
	<tr><td width="100%" background="<%=sHomePath%>image/fbox_bg.gif">
		<p class="sml">&nbsp;<a href="<%=sHomePath & Session("UserURL")%>">My&nbsp;<b>assortis<span class="rs">.com</span></b>&nbsp;account</a></p>
		<img src="<%=sHomePath%>image/x.gif" width=50 height=5><br>
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
		<td width="70"><p class="sml">&nbsp;&nbsp;Username&nbsp;</td>
		<td width="100" align="left"><input type="text" style="width:90px; margin-top:1px; margin-bottom:1px;" name="login_name" size="11" value="<%=Session("LoginName")%>"></td></tr>
		<tr><td><p class="sml">&nbsp;&nbsp;Password&nbsp;</td>
		<td align="left"><input type="password" name="login_pwd" style="width:90px; margin-top:1px; margin-bottom:1px;" size="11"></td></tr>
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


Sub ShowLoginBoxBefore2007Dec
	If Session("UserID")>0 Then
	%>
  	<table width="176" border="0" cellpadding="0" cellspacing="0">
	<tr><td width="100%"><img src="<%=sHomePath%>image/fbox_ttltop.gif" width="176" height="4"></td></tr>
	<tr><td width="100%" bgcolor="#EAEAEA" background="<%=sHomePath%>image/fbox_ttlbg.gif"><p class="fbox" align="center">Welcome&nbsp;<b><%=Session("UserName")%></b></p></td></tr>
	<tr><td width="100%"><img src="<%=sHomePath%>image/fbox_ttlbtm.gif" width="176" height="4"></td></tr>
	<tr><td width="100%"><img src="<%=sHomePath%>image/fbox_bg.gif" width="176" height="4"></td></tr>
	<tr><td width="100%" background="<%=sHomePath%>image/fbox_bg.gif">
		<p class="sml">&nbsp;<a href="<%=sHomePath & Session("UserURL")%>">My&nbsp;<b>assortis<span class="rs">.com</span></b>&nbsp;account</a></p>
		<img src="<%=sHomePath%>image/x.gif" width=50 height=5><br>
		<p class="sml">&nbsp;<a class="bl" href="<%=sHomePath & "logout.asp" & sParams%>">Logout</a></p>
	</td></tr>
	<tr><td width="100%"><img src="<%=sHomePath%>image/fbox_bg.gif" width="176" height="4"></td></tr>
	<tr><td width="100%"><img src="<%=sHomePath%>image/fbox_btm.gif" width="176" height="7"></td></tr>
	</table>

	<% Else	%>

  	<table width="176" border="0" cellpadding="0" cellspacing="0">
	<tr><td width="100%"><img src="<%=sHomePath%>image/fbox_ttltop.gif" width="176" height="4"></td></tr>
	<tr><td width="100%" bgcolor="#EAEAEA" background="<%=sHomePath%>image/fbox_ttlbg.gif"><p class="fbox" align="center"><b>assortis<span class="rs">.com</span></b> users entry</p></td></tr>
	<tr><td width="100%"><img src="<%=sHomePath%>image/fbox_ttlbtm.gif" width="176" height="4"></td></tr>
	<tr><td width="100%"><img src="<%=sHomePath%>image/fbox_bg.gif" width="176" height="4"></td></tr>
	<tr><td width="100%" background="<%=sHomePath%>image/fbox_bg.gif">
		<table width="170" background="" cellspacing=0 cellpadding=0 border=0>
		<tr><form method="post" action="<%=sHomePath%>login.asp?<%=sScriptFullNameAsParams%>" name="LoginForm" onSubmit="Login();return false;">
		<td width="70"><p class="sml">&nbsp;&nbsp;Login&nbsp;name&nbsp;</td>
		<td width="100" align="left"><input type="text" style="width:90px; margin-top:1px; margin-bottom:1px;" name="login_name" size="11" value="<%=Session("LoginName")%>"></td></tr>
		<tr><td><p class="sml">&nbsp;&nbsp;Password&nbsp;</td>
		<td align="left"><input type="password" name="login_pwd" style="width:90px; margin-top:1px; margin-bottom:1px;" size="11"></td></tr>
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


Sub ShowRegistrationProgressBar(sServiceTypeIn, iStep)
If Not (iUserAdmin=1 Or sApplicationName="expert" Or sApplicationName="external") Then Exit Sub

Dim arrStepTitle(), arrStepLink(), iLoop, iStepsNumber, sStepStyle, sStepImage, sServiceDescription, sServiceType
sServiceType=Left(sServiceTypeIn, 3)

	If sServiceType="CV" Then
		If sBackOffice>"" Then 
			iStepsNumber=7
		Else
			iStepsNumber=6
		End If

		ReDim arrStepTitle(iStepsNumber)
		arrStepTitle(1)=GetLabel(sCvLanguage, "Personal information")
		arrStepTitle(2)=GetLabel(sCvLanguage, "Education")
		arrStepTitle(3)=GetLabel(sCvLanguage, "Training")
		arrStepTitle(4)=GetLabel(sCvLanguage, "Professional experience")
		arrStepTitle(5)=GetLabel(sCvLanguage, "Languages")
		arrStepTitle(6)=GetLabel(sCvLanguage, "Contact details & availability")

		ReDim arrStepLink(iStepsNumber)
		arrStepLink(1)="<a href=""register.asp" & sParams & """>"
		arrStepLink(2)="<a href=""register2.asp" & sParams & """>"
		arrStepLink(3)="<a href=""register21.asp" & sParams & """>"
		arrStepLink(4)="<a href=""register3.asp" & sParams & """>"
		arrStepLink(5)="<a href=""register4.asp" & sParams & """>"
		arrStepLink(6)="<a href=""register5.asp" & sParams & """>"
		sServiceDescription="Curriculum Vitae online registration"
		If sBackOffice>"" Then 
			arrStepTitle(7)="&nbsp;Review&nbsp;<br>&nbsp;&amp;&nbsp;manage&nbsp;CV&nbsp;"
			arrStepLink(7)="<a href=""register6.asp" & sParams & """>"
		End If
	End If %>

	<!-- The registration progress bar -->

	<% If iStep>0 And sServiceType<>"RCV" And sServiceType<>"UCV" Then %>
	<table width=500 cellspacing=0 cellpadding=0 border=0 align="center">
	<tr><td width="10%">&nbsp;</td>
	<% For iLoop=1 To iStepsNumber
		If iStep=iLoop Then 
			sStepStyle="class=""rs"""
			sStepImage="progs_" & LCase(sServiceType) & CStr(iLoop) & "r" %>
<!--			<td><img src="<% =sHomePath %>image/progs_line.gif" width=2 height=42 hspace=5></td> -->
		<% Else
			sStepImage="progs_" & LCase(sServiceType) & CStr(iLoop) & "b"
		End If %>
	<% If sServiceType<>"CV" Or iStep=iLoop Then %>
	<td><img src="<% =sHomePath %>image/<% =sStepImage %>.gif" hspace=5 alt="<% =arrStepTitle(iLoop) %>"></td>
	<% End If %>
	<td align="center"><p><span <% =sStepStyle %>>
		<% If iLoop<2 Or iExpertID>0 Or sBackOffice>"" Then %><% =arrStepLink(iLoop) %><% Else %><u><% End If %>
		<% =arrStepTitle(iLoop) %>
		<% If iLoop<2 Or iExpertID>0 Or sBackOffice>"" Then %></a><% Else %></u><% End If %>
		</span></p></td>
		<% If iStep=iLoop Then %>
<!--			<td><img src="<% =sHomePath %>image/progs_line.gif" width=2 height=42 hspace=5></td> -->
		<% End If %>
		<% If iLoop < iStepsNumber Then %>
			<td><img src="<% =sHomePath %>image/progs_arw.gif" width=10 height=42 hspace=5></td>
		<% End If %>
	<% sStepStyle=""
	Next %>

	<td width="10%">&nbsp;</td></tr>
	</table><br> 

	<!-- Blue horisontal line -->
	<table width=100% cellspacing=0 cellpadding=0 border=0 align="center">
	<tr height=2><td bgcolor="#97CAFB"><img src="<% =sHomePath %>image/x.gif" width=600 height=2></td></tr>
	</table>
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

	If sUserType="expert"  Then
		sCalendarUrl=sScriptFileName & ReplaceUrlParams(sParams,"ei=" & iExpertID)
	Else
		sCalendarUrl=sScriptFileName & ReplaceUrlParams(sParams,"mi=" & iMemberID)
		Set objTempRs=GetMemberSubscription(iMemberID, 21)
		If objTempRs("id_AccountStatus")=1 And objTempRs("id_PaymentType")=1 Then
			Dim dTrialStartDate, dTrialEndDate
			dTrialStartDate=DateAdd("d", -7, objTempRs("macStartDate"))
			dTrialEndDate=objTempRs("macExpDate")
			Response.Write("<p class=""sml"" align=""center"">You can view notices since " & ConvertDateForText(dTrialStartDate, "&nbsp;", "DDMMYYYY") & "</p><img src=""../../image/fbox_bg.gif"" width=176 height=5><br>")

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

<table width="96%" cellpadding=1 cellspacing=1 border=1 bgcolor="#FFFFFF" align="center">
<tr><td valign="top" align="center"><a href="<%=sCalendarUrl & "&md=" & dPrevMonthLastDay %>"><img src="<%=sHomePath%>image/ico_prv.gif" width=15 height=13 hspace=0 vspace=2 border=0 alt="Previous month" align="center"></a></td><td valign="bottom" <% If sDuration="m1" Then %>bgcolor="#C9DAFC"<% End If %> colspan=5><p class="sml" align="center"><a class="bl" href="<%=sCalendarUrl & "&md=" & LCase(Left(MonthName(Month(dMailDate)),3)) & Year(dMailDate) %>"><%=ConvertDateForText(dMailDate, "&nbsp;", "MonthYear")%></p></td><td valign="top" align="center"><a href="<%=sCalendarUrl & "&md=" & dNextMonthFirstDay %>"><img src="<%=sHomePath%>image/ico_nxt.gif" width=15 height=13 border=0 hspace=0 vspace=2 alt="Next month" align="center"></td></tr>

<tr>
<td width="14%" bgcolor="#B3CBF9" align="center"><p class="sml2">M</p></td>
<td width="14%" bgcolor="#B3CBF9" align="center"><p class="sml2">T</p></td>
<td width="14%" bgcolor="#B3CBF9" align="center"><p class="sml2">W</p></td>
<td width="14%" bgcolor="#B3CBF9" align="center"><p class="sml2">T</p></td>
<td width="14%" bgcolor="#B3CBF9" align="center"><p class="sml2">F</p></td>
<td width="14%" bgcolor="#B3CBF9" align="center"><p class="sml2">S</p></td>
<td width="14%" bgcolor="#B3CBF9" align="center"><p class="sml2">S</p></td>
</tr>
<tr><td bgcolor="#B3CBF9" align="center"><p class="sml2"></p></td><td bgcolor="#B3CBF9" align="center"><p class="sml2"></p></td><td bgcolor="#B3CBF9" align="center"><p class="sml2"></p></td><td bgcolor="#B3CBF9" align="center"><p class="sml2"></p></td><td bgcolor="#B3CBF9" align="center"><p class="sml2"></p></td><td bgcolor="#B9CFFA" align="center"><p class="sml2"></p></td><td bgcolor="#B9CFFA" align="center"><p class="sml2"></p></td></tr>

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
<td <% If (CInt(iActiveDay)=wd) And (iFWDayOfMonth <=(wk-1)*7+d) And sDuration="d1" Then %>bgcolor="#C9DAFC"<% Else %>bgcolor="<%=bg(d)%>"<% End If %> align="center"><p class="sml2"><% If (iFWDayOfMonth <=(wk-1)*7+d) And wd<= iLDayOfMonth Then %><a class="bl" href="<%=sCalendarUrl & "&md=" & DateAdd("d", wd-1, dFirstDayOfMonth)%>">
<% 
	' To highlight the day with any notices
	'objTempRs=GetDataOutParamsSP("usp_MmbBscDailyNoticesNumber", Array( _
	'	Array(, adInteger, , iMemberID), _
	'	Array(, adVarChar, 16, ConvertDMYForSQL(iActiveYear, iActiveMonth, wd))), _
	'	Array( Array(, adInteger)))
	'If objTempRs(0)>0 Then Response.Write "<b>"
	'Set objTempRs=Nothing

	If iActiveYear=Year(Now()) And iActiveMonth=Month(Now()) And wd=Day(Now()) Then Response.Write "<b>"

%>
<%=wd%></b>
<%wd=wd+1%></a><% Else %>&nbsp;<% End If %></p></td>
<% Next %>
</tr>
<% End If %>
<% Next %>

<tr><td bgcolor="#B3CBF9" align="center"><p class="sml2"></p></td><td bgcolor="#B3CBF9" align="center"><p class="sml2"></p></td><td bgcolor="#B3CBF9" align="center"><p class="sml2"></p></td><td bgcolor="#B3CBF9" align="center"><p class="sml2"></p></td><td bgcolor="#B3CBF9" align="center"><p class="sml2"></p></td><td bgcolor="#B9CFFA" align="center"><p class="sml2"></p></td><td bgcolor="#B9CFFA" align="center"><p class="sml2"></p></td></tr>
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
