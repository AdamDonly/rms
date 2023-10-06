<%
Dim arrGenderID(2), arrGenderTitle(2)

If sCvLanguage = cLanguageFrench Then %>
<!--#include file="fr/datGender.asp"-->
<% ElseIf sCvLanguage = cLanguageSpanish Then %>
<!--#include file="sp/datGender.asp"-->
<% Else %>
<!--#include file="en/datGender.asp"-->
<% End If %>