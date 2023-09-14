<%
' Path of the main home directory
Dim sHomePath
sHomePath="/"

Dim sApplicationTitle
sApplicationTitle="ICA CVs database"

' Email addresses of a client
Dim sEmailClient, sEmailClientCopy
sEmailClient="imc@ibf.be"
sEmailClientCopy="grunina@icaworld.net"

' Email address of a client CVIP system
Dim sEmailCvipSystem
sEmailCvipSystem="imc@ibf.be"

' Color scheme used for the application forms (RED, BLUE, ? GREEN, ? GRAY)
Dim sColorScheme
sColorScheme="BLUE"

' Type of expert registration used by default (QUICK, FULL)
Dim sDefaultRegistration
sDefaultRegistration="FULL"

' Type of view of the search for experts results page
Const cViewList = 0
Const cViewTable = 1
Const cViewTable2 = 2

Dim sDefaultViewSearchResults
sDefaultViewSearchResults=cViewTable2

' Type of visibility for external registrar (VISIBLE, OBFUSCATED, HIDDEN)
Const cNameVisible = 0
Const cNameObfuscated = 1
Const cNameHidden = 2

Dim sContactDetailsExternally
sContactDetailsExternally=cNameVisible

Dim sPageTitle
%>
<!--#include file="__init_language.asp"-->
<!--#include file="__init_document.asp"-->
<!--#include file="__init_type.asp"-->
<!--#include virtual="/_common/_data/datLabel.asp"-->