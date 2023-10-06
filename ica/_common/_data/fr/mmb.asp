<% 
' Appending data in arrays: mNtInfo - Countries names, mNtCode - Countries codes




    if Request("Active")<>"" then
    Active=Request("Active")
    else
    Active=1
    end if
    ''''''''''''''''''''''''''
	mNt=0
	Dim mNtInfo(300)
	Dim mNtCode(300)
	Dim mNtZone(300)

	i=0
	mNtInt=""

Rem Append data in arrays: mGZnInfo - Geo Zones names, mGZnCode - GeoZones codes

	strSQL="SELECT DISTINCT tbl_GeoZone.id_GeoZone, tbl_GeoZone.id_Continent, tbl_GeoZone.Geo_ZoneEng, tbl_GeoZone.db_Scroll FROM (tbl_GeoZone LEFT JOIN tbl_Country ON tbl_GeoZone.id_GeoZone = tbl_Country.id_GeoZone) LEFT JOIN lnkMmb_Cou_BSC ON tbl_Country.id_Country = lnkMmb_Cou_BSC.id_Country "
	strSQL=strSQL & "WHERE lnkMmb_Cou_BSC.id_Member="& MemberId& "  and lnkMmb_Cou_BSC.Active="&Active&"  AND db_NotVisible=0 AND id_Country<>524 order by 2,3;"
	objrs1.Open strSQL,objconn,3,3
	mGZ=objrs1.RecordCount

	Dim mGZnInfo(30)
	Dim mGZnCode(30)
	Dim mGZnContinent(30)
	Dim mGZnScroll(30)
	i=0
	
	Do Until objrs1.EOF 
	mGZnInfo(i)=objrs1("Geo_ZoneEng")
	mGZnCode(i)=objrs1("id_GeoZone")
	mGZnContinent(i)=objrs1("id_Continent")

		objrs2.Open "select tbl_Country.id_Country, tbl_Country.couNameEng, tbl_Country.id_GeoZone FROM tbl_Country LEFT JOIN lnkMmb_Cou_BSC ON tbl_Country.id_Country=lnkMmb_Cou_BSC.id_Country WHERE id_GeoZone=" & objrs1("id_GeoZone") &" AND lnkMmb_Cou_BSC.id_Member="& MemberId &"  and lnkMmb_Cou_BSC.Active="&Active&"  order by 2;",objconn,3,3
		j=0

		Do Until objrs2.EOF 
		mNtInfo(j+mNt)=objrs2("couNameEng")
		mNtCode(j+mNt)=objrs2("id_Country")
		mNtZone(j+mNt)=objrs2("id_GeoZone")
		j=j+1
		objrs2.MoveNext
		Loop
		mNt=mNt+objrs2.Recordcount
		objrs2.Close
		mGZnScroll(i)=j

	i=i+1
	objrs1.MoveNext
	Loop
	objrs1.Close
	%>

<%
 Rem Append data in arrays: mOrgInfo - Organisations names, mOrgCode - Organisations ids
	objrs1.Open "select * from (tbl_Donors left join lnkMmb_Don_BSC on tbl_Donors.id_Organisation=lnkMmb_Don_BSC.id_Organisation) where ( lnkMmb_Don_BSC.id_Member=" & MemberId & " and lnkMmb_Don_BSC.Active="&Active&") order by orgMainDonor DESC, orgNameEng;",objconn,3,3
	mOrg=objrs1.RecordCount
	Dim mOrgInfo(200)
	Dim mOrgCode(200)
	mOrgInt=""
	i=0

	Do Until objrs1.EOF 
	mOrgInfo(i)=objrs1("orgNameEng")
	mOrgCode(i)=objrs1("id_Organisation")
	i=i+1
	objrs1.MoveNext
	Loop
	objrs1.Close
	%>

<%
	Dim mExTInfo(400)
	Dim mExTCode(400)
	Dim mExTSrch(400)
	mExT=0

Rem Append data in arrays: mExFInfo - Expertise fields names, mExFCode - fields codes

	strSQL="SELECT DISTINCT tbl_MainSectors.id_MainSector, tbl_MainSectors.mnsDescriptionEng, tbl_MainSectors.mnsDescriptionFra, tbl_MainSectors.mnsDescriptionSpa, tbl_MainSectors.mnsShortEng, tbl_MainSectors.db_Scroll "
	strSQL=strSQL & "FROM (tbl_MainSectors LEFT JOIN tbl_Sectors ON tbl_MainSectors.id_MainSector = tbl_Sectors.id_MainSector) LEFT JOIN lnkMmb_Sct_BSC ON tbl_Sectors.id_Sector = lnkMmb_Sct_BSC.id_Sector "
	strSQL=strSQL & "WHERE lnkMmb_Sct_BSC.id_Member="& MemberId &" and lnkMmb_Sct_BSC.Active="&Active&"  order by mnsDescriptionEng;"
	objrs1.Open strSQL,objconn,3,3
	mExF=objrs1.RecordCount
	Dim mExFInfo(50)
	Dim mExFCode(50)
	Dim mExFShort(50)
	Dim mExFScroll(50)
	Dim mExFShift(50)
	i=0

	Do Until objrs1.EOF 
	mExFInfo(i)=objrs1("mnsDescriptionEng")
	mExFCode(i)=objrs1("id_MainSector")
	mExFShort(i)=objrs1("mnsShortEng")

		objrs2.Open "SELECT * from tbl_Sectors LEFT JOIN lnkMmb_Sct_BSC ON tbl_Sectors.id_Sector=lnkMmb_Sct_BSC.id_Sector WHERE tbl_Sectors.id_Sector<1000 AND id_MainSector=" & objrs1("id_MainSector") &" AND lnkMmb_Sct_BSC.id_Member="& MemberId &" and lnkMmb_Sct_BSC.Active="&Active&"  order by 2;",objconn,3,3
		j=0
		sh=0
		Do Until objrs2.EOF 
		mExTInfo(j+mExT)=objrs2("sctDescriptionEng")
		mExTCode(j+mExT)=objrs2("id_Sector")
		mExTSrch(j+mExT)=objrs2("id_MainSector")
		If Len(mExTInfo(j+mExT))>56 Then
		  sh=sh+1
		End If

		j=j+1
		objrs2.MoveNext
		Loop
		mExT=mExT+objrs2.RecordCount
		objrs2.Close
		mExFScroll(i)=j
		mExFShift(i)=sh

	i=i+1
	objrs1.MoveNext
	Loop
	objrs1.Close

%>  
