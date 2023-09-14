<%
Dim arrMonthID(12)
Dim arrMonthName(12)

If sCvLanguage = cLanguageFrench Then %>
<!--#include file="fr/datMonth.asp"-->
<% ElseIf sCvLanguage = cLanguageSpanish Then %>
<!--#include file="sp/datMonth.asp"-->
<% Else %>
<!--#include file="en/datMonth.asp"-->
<% End If %>