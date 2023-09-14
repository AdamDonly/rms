<%
sSearchKeywordsHighlight=Trim(Request.QueryString("txt") + " " + Request.QueryString("srch_queryadd"))
sSearchKeywordsHighlight=Replace(sSearchKeywordsHighlight, " AND ", " ")
sSearchKeywordsHighlight=Replace(sSearchKeywordsHighlight, " OR ", " ")
sSearchKeywordsHighlight=Replace(sSearchKeywordsHighlight, " NOT ", " ")
sSearchKeywordsHighlight=Replace(sSearchKeywordsHighlight, " NEAR ", " ")
sSearchKeywordsHighlight=Replace(sSearchKeywordsHighlight, "  ", " ")

If Len(sSearchKeywordsHighlight)>2 Then
	arrSearchKeywordsHighlight=Split(sSearchKeywordsHighlight, " ")
End If
%>