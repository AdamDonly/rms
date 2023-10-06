<!--#include virtual="/_common/_class/db.asp"-->
<%
' Active user company database
Dim objUserCompanyDB
Set objUserCompanyDB = New CCompanyExpertDB
objUserCompanyDB.LoadCompanyDatabase iUserCompanyID, NULL, NULL

Dim objExpertAssortisDB
Set objExpertAssortisDB = New CCompanyExpertDB
objExpertAssortisDB.LoadCompanyDatabase Null, Null, "assortis"


' Fake db for all experts from members that already left ICA or from a database to which a current user don't have any access
Dim objNullDB
Set objNullDB = New CCompanyExpertDB
objNullDB.ID = 0
objNullDB.DatabaseCode = "000-"

' List of all available ICA databases
Dim objExpertDBList
Set objExpertDBList = New CCompanyExpertDBList
objExpertDBList.DefaultDatabase = objUserCompanyDB.Database
If bAssortisSubscriptionEdbActive Then
	' Force including assortis db to the list of experts dbs available for a member
	objExpertDBList.LoadCompanyDatabaseList iUserCompanyID, aCompanyExpertDBAssortisID, NULL, 1
Else
	objExpertDBList.LoadCompanyDatabaseList iUserCompanyID, NULL, NULL, 1
End If 

' Selected expert database
Dim objExpertDB

Const aUserAccessMaskNoAccess = 0
Const aUserAccessMaskView = 1
Const aUserAccessMaskAdd = 2
Const aUserAccessMaskEdit = 4
%>