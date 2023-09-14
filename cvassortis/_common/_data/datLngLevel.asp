<%
Dim arrLanguageLevelID(5)
Dim arrLanguageLevelTitle(5)

If sCvLanguage = cLanguageFrench Then %>
<!--#include file="fr/datLngLevel.asp"-->
<% Else %>
<!--#include file="en/datLngLevel.asp"-->
<% End If %>