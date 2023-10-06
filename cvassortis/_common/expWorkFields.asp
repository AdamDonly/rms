<% 
' Appending data in arrays: mNtInfo - Countries names, mNtCode - Countries codes
' Appending data in arrays: mGZnInfo - Geo Zones names, mGZnCode - GeoZones codes

i=0
mNt=0
mGZ=0
iFlag=0

Dim mNtInfo(196), mNtCode(196), mNtZone(196)
Dim mGZnInfo(17), mGZnCode(17), mGZnContinent(17), mGZnScroll(17)

Set objTempRs=GetDataRecordsetPlusOutSPWithConn(objConnCustom, "usp_ExpCvvExperienceCouSelect", Array( _
	Array(, adInteger, , iExpertID), _
	Array(, adInteger, , iExpPrjID), _
	Array(, adVarChar, 10, "rs")), Array( _
	Array(, adVarChar, 4000)))
While Not objTempRs.Eof

	' mNtInfo(mNt)=objTempRs("couNameEng")
	mNtCode(mNt)=objTempRs("id_Country")
	mNtZone(mNt)=objTempRs("id_GeoZone")

	If iFlag<>objTempRs("id_GeoZone") Then
		mGZ=mGZ+1
		mGZnInfo(mGZ)=objTempRs("Geo_ZoneEng")
		mGZnCode(mGZ)=objTempRs("id_GeoZone")
		' mGZnContinent(mGZ)=objTempRs("id_Continent")
		iFlag=mGZnCode(mGz)
		mGZnScroll(mGz)=i
		i=0
	End If

	mNt=mNt+1
	i=i+1
objTempRs.MoveNext
WEnd
objTempRs.Close
Set objTempRs=Nothing

i=0
mOrg=0
Dim mOrgInfo(20), mOrgCode(20)

' Appending data in arrays: mOrgInfo - Organisations names, mOrgCode - Organisations ids
Set objTempRs=GetDataRecordsetPlusOutSPWithConn(objConnCustom, "usp_ExpCvvExperienceDonSelect", Array( _
	Array(, adInteger, , iExpertID), _
	Array(, adInteger, , iExpPrjID), _
	Array(, adVarChar, 10, "rs")), Array( _
	Array(, adVarChar, 4000)))
While Not objTempRs.Eof
	' mOrgInfo(mOrg)=objTempRs("orgNameEng")
	mOrgCode(mOrg)=objTempRs("id_Organisation")
	mOrg=mOrg+1
objTempRs.MoveNext
WEnd
objTempRs.Close
Set objTempRs=Nothing


i=0
mExT=0
mExF=0
iFlag=0

Dim mExTInfo(400), mExTCode(400), mExTSrch(400)
Dim mExFInfo(20), mExFCode(20), mExFShort(20), mExFScroll(20), mExFShift(20)

' Appending data in arrays: mExFInfo - Expertise fields names, mExFCode - fields codes

Set objTempRs=GetDataRecordsetPlusOutSPWithConn(objConnCustom, "usp_ExpCvvExperienceSctSelect", Array( _
	Array(, adInteger, , iExpertID), _
	Array(, adInteger, , iExpPrjID), _
	Array(, adVarChar, 10, "rs")), Array( _
	Array(, adVarChar, 4000)))

While Not objTempRs.Eof

	' mExTInfo(mExT)=objTempRs("sctDescriptionEng")
	mExTCode(mExT)=objTempRs("id_Sector")
	mExTSrch(mExT)=objTempRs("id_MainSector")

	If iFlag<>objTempRs("id_MainSector") Then
		mExF=mExF+1
		mExFInfo(mExF)=objTempRs("mnsDescriptionEng")
		mExFCode(mExF)=objTempRs("id_MainSector")
		iFlag=mExFCode(mExF)
		mExFScroll(mExF)=i
		i=0
	End If

	mExT=mExT+1
	i=i+1
objTempRs.MoveNext
WEnd
objTempRs.Close
Set objTempRs=Nothing

%>

