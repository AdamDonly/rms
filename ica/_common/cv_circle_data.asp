<% 
' Appending data in arrays: mNtInfo - Countries names, mNtCode - Countries codes
' Appending data in arrays: mGZnInfo - Geo Zones names, mGZnCode - GeoZones codes

i = 0
mNt = 0
mGZ = 0
iFlag = 0

Dim mNtInfo(270), mNtCode(270), mNtZone(270)
Dim mGZnInfo(17), mGZnCode(17), mGZnContinent(17), mGZnScroll(17)

If sAction = "fromcv" Then
Set objTempRs = GetDataRecordsetSPWithConn(objConnCustom, "usp_ExpertExperienceCountrySelect", Array( _
	Array(, adInteger, , iCvID), _
	Array(, adInteger, , Null), _
	Array(, adVarChar, 80, Null)))
Else
Set objTempRs = GetDataRecordsetSP("usp_" & sIcaServerSqlPrefix & "ExpertCompanyExpertCircleCountrySelect", Array( _
	Array(, adVarChar, 40, sCvUID), _
	Array(, adInteger, , iUserCompanyID)))
End If
While Not objTempRs.Eof
	mNtInfo(mNt) = objTempRs("couNameEng")
	mNtCode(mNt) = objTempRs("id_Country")
	mNtZone(mNt) = objTempRs("id_GeoZone")

	If iFlag<>objTempRs("id_GeoZone") Then
		mGZ=mGZ+1
		mGZnInfo(mGZ) = objTempRs("Geo_ZoneEng")
		mGZnCode(mGZ) = objTempRs("id_GeoZone")
		mGZnContinent(mGZ) = objTempRs("id_Continent")
		iFlag=mGZnCode(mGz)
		mGZnScroll(mGz)=i
		i = 0
	End If

	mNt = mNt + 1
	i = i + 1
objTempRs.MoveNext
WEnd
objTempRs.Close
Set objTempRs = Nothing

i = 0
mOrg = 0
Dim mOrgInfo(150), mOrgCode(150)

' Appending data in arrays: mOrgInfo - Organisations names, mOrgCode - Organisations ids
If sAction = "fromcv" Then
Set objTempRs = GetDataRecordsetSPWithConn(objConnCustom, "usp_ExpertExperienceDonorSelect", Array( _
	Array(, adInteger, , iCvID), _
	Array(, adInteger, , Null), _
	Array(, adVarChar, 80, Null)))
Else
Set objTempRs = GetDataRecordsetSP("usp_" & sIcaServerSqlPrefix & "ExpertCompanyExpertCircleDonorSelect", Array( _
	Array(, adVarChar, 40, sCvUID), _
	Array(, adInteger, , iUserCompanyID)))
End If
While Not objTempRs.Eof
	mOrgInfo(mOrg) = objTempRs("orgNameEng")
	mOrgCode(mOrg) = objTempRs("id_Organisation")
	mOrg = mOrg + 1
objTempRs.MoveNext
WEnd
objTempRs.Close
Set objTempRs = Nothing


i = 0
mExT = 0
mExF = 0
iFlag = 0

Dim mExTInfo(400), mExTCode(400), mExTSrch(400)
Dim mExFInfo(21), mExFCode(21), mExFShort(21), mExFScroll(21), mExFShift(21)

' Appending data in arrays: mExFInfo - Expertise fields names, mExFCode - fields codes

If sAction = "fromcv" Then
Set objTempRs = GetDataRecordsetSPWithConn(objConnCustom, "usp_ExpertExperienceSectorSelect", Array( _
	Array(, adInteger, , iCvID), _
	Array(, adInteger, , Null), _
	Array(, adVarChar, 80, Null)))
Else
Set objTempRs = GetDataRecordsetSP("usp_" & sIcaServerSqlPrefix & "ExpertCompanyExpertCircleSectorSelect", Array( _
	Array(, adVarChar, 40, sCvUID), _
	Array(, adInteger, , iUserCompanyID)))
End If
While Not objTempRs.Eof
	mExTInfo(mExT) = objTempRs("sctDescriptionEng")
	mExTCode(mExT) = objTempRs("id_Sector")
	mExTSrch(mExT) = objTempRs("id_MainSector")

	If iFlag<>objTempRs("id_MainSector") Then
		mExF = mExF + 1
		mExFInfo(mExF) = objTempRs("mnsDescriptionEng")
		mExFCode(mExF) = objTempRs("id_MainSector")
		iFlag = mExFCode(mExF)
		mExFScroll(mExF) = i
		i = 0
	End If

	mExT = mExT + 1
	i = i + 1
objTempRs.MoveNext
WEnd
objTempRs.Close
Set objTempRs = Nothing


i = 0
mPT = 0
Dim mProcTypeNames(4), mProcTypeIDs(4)
' Appending data in arrays: mProcTypeNames - Procurement type names, mOrgCode - Procurement type ids
Set objTempRs = GetDataRecordsetSP("usp_" & sIcaServerSqlPrefix & "ExpertCompanyExpertCircleProcurementTypeSelect", Array( _
	Array(, adVarChar, 40, sCvUID), _
	Array(, adInteger, , iUserCompanyID)))
While Not objTempRs.Eof
	mProcTypeNames(mPT) = objTempRs("PROCTYPE_NAME")
	mProcTypeIDs(mPT) = objTempRs("IDPROCUREMENTTYPE")
	mPT = mPT + 1
objTempRs.MoveNext
WEnd
objTempRs.Close
Set objTempRs = Nothing
%>
