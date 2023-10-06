<%
Dim arrMaritalStatusID(2), arrMaritalStatusTitle(2)

If sCvLanguage = cLanguageFrench Then %>
<!--#include file="fr/datPsnStatus.asp"-->
<% ElseIf sCvLanguage = cLanguageSpanish Then %>
<!--#include file="sp/datPsnStatus.asp"-->
<% Else %>
<!--#include file="en/datPsnStatus.asp"-->
<% End If %>