<!--#include file="../dbc.asp"-->
<!--#include file="../fnc.asp"-->
<!--#include file="../fnc_exp.asp"-->
<%
CheckUserLogin sScriptFullName
%>
<!--#include file="cv_data.asp"-->
<%
' If member subscribing for Busops service  
If Request.Form() > "" Then

	sCountries = CheckString(Request.Form("mmb_cou_hid"))
	sDonors = CheckString(Request.Form("mmb_don_hid"))
	sSectors = CheckString(Request.Form("mmb_sct_hid"))
	sProcurementTypes = CheckString(Request.Form("mmb_proctypes"))

	' Removing the number of total selected items in every field
	sCountries = Mid(sCountries, InStr(sCountries,",")+1,Len(sCountries))
	sDonors = Mid(sDonors, InStr(sDonors,",")+1,Len(sDonors))
	sSectors = Mid(sSectors, InStr(sSectors,",")+1,Len(sSectors))

	' Saving countries of interest
	objTempRs = DeleteRecordSP("usp_" & sIcaServerSqlPrefix & "ExpertCompanyExpertCircleCountryDelete", Array( _ 
		Array(, adVarChar, 40, sCvUID), _
		Array(, adInteger, , iUserCompanyID)))
	Set objTempRs = Nothing
	objTempRs = GetDataOutParamsSP("usp_" & sIcaServerSqlPrefix & "ExpertCompanyExpertCircleCountryInsert", Array( _
		Array(, adVarChar, 2000, sCountries), _
		Array(, adVarChar, 40, sCvUID), _
		Array(, adInteger, , iUserCompanyID)), _
		Array( Array(, adInteger)))
	iTotalCountries=objTempRs(0)
	Set objTempRs=Nothing

	' Saving donors of interest - no need to save donors, as by default all funding agencies are selected:
'	objTempRs=DeleteRecordSP("usp_" & sIcaServerSqlPrefix & "ExpertCompanyExpertCircleDonorDelete", Array( _ 
'		Array(, adVarChar, 40, sCvUID), _
'		Array(, adInteger, , iUserCompanyID)))
'	Set objTempRs=Nothing
'	objTempRs=GetDataOutParamsSP("usp_" & sIcaServerSqlPrefix & "ExpertCompanyExpertCircleDonorInsert", Array( _ 
'		Array(, adVarChar, 2000, sDonors), _
'		Array(, adVarChar, 40, sCvUID), _
'		Array(, adInteger, , iUserCompanyID)), _
'		Array( Array(, adInteger)))
'	iTotalDonors=objTempRs(0)
'	Set objTempRs=Nothing

	' Saving sectors of interest (for non activated account)
	objTempRs=DeleteRecordSP("usp_" & sIcaServerSqlPrefix & "ExpertCompanyExpertCircleSectorDelete", Array( _ 
		Array(, adVarChar, 40, sCvUID), _
		Array(, adInteger, , iUserCompanyID)))
	Set objTempRs=Nothing
	objTempRs=GetDataOutParamsSP("usp_" & sIcaServerSqlPrefix & "ExpertCompanyExpertCircleSectorInsert", Array( _ 
		Array(, adVarChar, 2000, sSectors), _
		Array(, adVarChar, 40, sCvUID), _
		Array(, adInteger, , iUserCompanyID)), _
		Array( Array(, adInteger)))
	iTotalSectors=objTempRs(0)
	Set objTempRs=Nothing
	
	' Saving procurement types of interest:
	objTempRs=DeleteRecordSP("usp_" & sIcaServerSqlPrefix & "ExpertCompanyExpertCircleProcurementTypeDelete", Array( _ 
		Array(, adVarChar, 40, sCvUID), _
		Array(, adInteger, , iUserCompanyID)))
	Set objTempRs=Nothing
	objTempRs=GetDataOutParamsSP("usp_" & sIcaServerSqlPrefix & "ExpertCompanyExpertCircleProcurementTypeInsert", Array( _ 
		Array(, adVarChar, 20, sProcurementTypes), _
		Array(, adVarChar, 40, sCvUID), _
		Array(, adInteger, , iUserCompanyID)), _
		Array( Array(, adInteger)))
	iTotalProcurementTypes=objTempRs(0)
	Set objTempRs=Nothing

	Response.Redirect "/backoffice/register/register6.asp?uid=" & sCvUID
		
End If
%>	

<% CloseDBConnection %>
