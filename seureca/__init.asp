<%
Response.Redirect "https://www.assortis.com/"

' Path of the main home directory
Dim sHomePath
sHomePath="/seureca/"

Dim sApplicationTitle
sApplicationTitle="Seureca CVs database"

' Email addresses of a client
Dim sEmailClient, sEmailClientCopy
sEmailClient="stephanie.ahossi@veolia.com"
'sEmailClientCopy="PierreASCENCIO@veic.fr"
sEmailClientCopy=""

' Email address of a client CVIP system
Dim sEmailCvipSystem
sEmailCvipSystem="seureca@assortis.com"

' Color scheme used for the application forms (RED, BLUE, ? GREEN, ? GRAY)
Dim sColorScheme
sColorScheme="BLUE"

' Type of expert registration used by default (QUICK, FULL)
Dim sDefaultRegistration
sDefaultRegistration="FULL"

' Type of visibility for external registrar (VISIBLE, OBFUSCATED, HIDDEN)
Const cNameVisible = 0
Const cNameObfuscated = 1
Const cNameHidden = 2

Dim sContactDetailsExternally
sContactDetailsExternally=cNameVisible
%>
<!--#include file="__init_language.asp"-->
<!--#include file="__init_document.asp"-->
<!--#include file="__init_type.asp"-->
<!--#include virtual="/_common/_data/datLabel.asp"-->