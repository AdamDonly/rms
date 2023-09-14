<%
' Available languages
Const cLanguageEnglish = "Eng"
Const cLanguageFrench = "Fra"
Const cLanguageSpanish = "Spa"

' Activation of CVIP in several languages
Const cMultiLanguageDisabled = 0
Const cMultiLanguageEnabled = 1

Dim bCvMultiLanguageActive
bCvMultiLanguageActive = cMultiLanguageEnabled

Dim bInterfaceMultiLanguage
bInterfaceMultiLanguage = cMultiLanguageDisabled

' List of the languages enabled on the CVIP
Dim dictLanguage
Set dictLanguage = CreateObject("Scripting.Dictionary")

dictLanguage.Add cLanguageEnglish, "English"
dictLanguage.Add cLanguageFrench, "Français"
dictLanguage.Add cLanguageSpanish, "Español"

' Language variables
Dim sCvLanguage
Dim sDefaultCvLanguage
sDefaultCvLanguage = cLanguageEnglish

Dim sInterfaceLanguage
sForceCvLanguage = ReplaceIfEmpty(Request.QueryString("l"), Request.QueryString("lang"))
Dim sForceCvLanguage
sForceCvLanguage = Request.QueryString("lng")

Dim iLanguageLinkID

Function ForceCvLanguage()
	If (sForceCvLanguage = cLanguageEnglish) _
	Or (sForceCvLanguage = cLanguageFrench) _
	Or (sForceCvLanguage = cLanguageSpanish) Then
		sCvLanguage = sForceCvLanguage
	End If
End Function
%>