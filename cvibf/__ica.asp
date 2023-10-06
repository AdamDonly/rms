<!--#include virtual="/_common/_class/db.asp"-->
<%
' Active user company database
Dim objUserCompanyDB
Set objUserCompanyDB = New CCompanyExpertDB
objUserCompanyDB.LoadCompanyDatabase iUserCompanyID, NULL, NULL

' List of all available ICA databases
Dim objExpertDBList
Set objExpertDBList = New CCompanyExpertDBList
objExpertDBList.DefaultDatabase=objUserCompanyDB.Database
objExpertDBList.LoadCompanyDatabaseList NULL, NULL, NULL, 1

' Selected expert database
Dim objExpertDB

Const aUserMaskNoAccess = 0
Const aUserMaskView = 1
Const aUserMaskAdd = 2
Const aUserMaskEdit = 4
%>