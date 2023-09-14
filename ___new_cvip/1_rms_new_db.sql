/* 
Script to generate empty CVIP database for assortis clients
Rename inrae -> real name of a client
*/

USE [master]
GO

/* Create database */
CREATE DATABASE [rms_inrae] ON  PRIMARY 
( NAME = N'rms_inrae_data', FILENAME = N'E:\Program Files\Microsoft SQL Server\MSSQL11.MSSQLSERVER\MSSQL\Data\rms_inrae.mdf' , SIZE = 6144KB , MAXSIZE = UNLIMITED, FILEGROWTH = 10%)
 LOG ON 
( NAME = N'rms_inrae_log', FILENAME = N'E:\Program Files\Microsoft SQL Server\MSSQL11.MSSQLSERVER\MSSQL\Data\rms_inrae_log.ldf' , SIZE = 2048KB , MAXSIZE = UNLIMITED, FILEGROWTH = 10%)
COLLATE SQL_Latin1_General_CP1_CI_AI
GO

USE [rms_inrae]
GO

/* Enable fulltext index */
if (select DATABASEPROPERTY(DB_NAME(), N'IsFullTextEnabled')) <> 1 
exec sp_fulltext_database N'enable' 
GO

/* Create users */
EXEC dbo.sp_grantdbaccess @loginame = N'BRUSSELS\ASSORTIS BE', @name_in_db = N'assortis2'
GO
EXEC dbo.sp_grantdbaccess @loginame = N'BRUSSELS\rms', @name_in_db = N'rms'
GO

if not exists (select * from dbo.sysusers where name = N'assortis' and uid > 16399)
	EXEC sp_addrole N'assortis'
GO

exec sp_addrolemember N'assortis', N'assortis2'
GO


GO

if not exists (select * from dbo.sysfulltextcatalogs where name = N'Experts')
exec sp_fulltext_catalog N'Experts', N'create' 

GO

CREATE TABLE [dbo].[lnkExp_Edu] (
	[id_ExpEdu] [int] IDENTITY (1, 1) NOT NULL ,
	[id_Expert] [int] NOT NULL ,
	[id_EduSubject] [int] NULL ,
	[id_EduSubject1Eng] [nvarchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[id_EduSubject1Fra] [nvarchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[id_EduSubject1Spa] [nvarchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[id_EduType] [smallint] NULL ,
	[eduStartDate] [smalldatetime] NULL ,
	[eduEndDate] [smalldatetime] NULL ,
	[eduDiploma] [smallint] NULL ,
	[eduDiploma1Eng] [nvarchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[eduDiploma1Fra] [nvarchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[eduDiploma1Spa] [nvarchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[eduDescriptionEng] [ntext] COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[eduDescriptionFra] [ntext] COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[eduDescriptionSpa] [ntext] COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[InstNameEng] [nvarchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[InstNameFra] [nvarchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[InstNameSpa] [nvarchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[InstLocationEng] [nvarchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[InstLocationFra] [nvarchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[InstLocationSpa] [nvarchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[eduOtherEng] [nvarchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[eduOtherFra] [nvarchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[eduOtherSpa] [nvarchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[LinkId] [int] NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[lnkExp_Lan] (
	[id_ExpLan] [int] IDENTITY (1, 1) NOT NULL ,
	[id_Expert] [int] NOT NULL ,
	[id_Language] [smallint] NOT NULL ,
	[exlSpeaking] [smallint] NULL ,
	[exlWriting] [smallint] NULL ,
	[exlReading] [smallint] NULL ,
	[exlAverage] AS (round((convert(numeric(5,2),([exlReading] + [exlSpeaking] + [exlWriting])) / 3),0)) ,
	[exlLevel] AS ([exlReading] + [exlSpeaking] + [exlWriting]) ,
	[Language1] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AI NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[lnkExp_Pos] (
	[id_Expert] [int] NOT NULL ,
	[id_Position] [int] NOT NULL ,
	[id_Status] [tinyint] NULL ,
	[epsProvidedCompany] [varchar] (400) COLLATE Latin1_General_CI_AS NULL ,
	[epsProvidedPerson] [varchar] (400) COLLATE Latin1_General_CI_AS NULL ,
	[epsFee] [smallint] NULL ,
	[epsFeeCurrency] [smallint] NULL ,
	[epsComments] [text] COLLATE Latin1_General_CI_AS NULL ,
	[epsCreateDate] [smalldatetime] NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[lnkExp_Prj] (
	[id_Expert] [int] NOT NULL ,
	[id_Project] [int] NOT NULL ,
	[id_ExpertStatus] [smallint] NULL ,
	[epjProvidedCompany] [varchar] (400) COLLATE Latin1_General_CI_AS NULL ,
	[epjProvidedPerson] [varchar] (400) COLLATE Latin1_General_CI_AS NULL ,
	[epjFee] [numeric](9, 2) NULL ,
	[epjFeeCurrency] [char] (3) COLLATE Latin1_General_CI_AS NULL ,
	[epjComments] [text] COLLATE Latin1_General_CI_AS NULL ,
	[epjCreateDate] [smalldatetime] NULL ,
	[epjModifyDate] [smalldatetime] NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[lnkExp_RankCou] (
	[id_ExpRankCou] [int] IDENTITY (1, 1) NOT NULL ,
	[id_Expert] [int] NOT NULL ,
	[id_Country] [smallint] NULL ,
	[rnkCouValue] [real] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[lnkExp_RankSct] (
	[id_ExpRankSct] [int] IDENTITY (1, 1) NOT NULL ,
	[id_Expert] [int] NOT NULL ,
	[id_Sector] [int] NULL ,
	[rnkSctValue] [real] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[lnkExp_StatusCV] (
	[id_Expert] [int] NOT NULL ,
	[id_Status] [tinyint] NOT NULL ,
	[estModifyDate] [smalldatetime] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[lnkExp_Wke] (
	[id_ExpWke] [int] IDENTITY (1, 1) NOT NULL ,
	[id_Expert] [int] NOT NULL ,
	[wkeStartDate] [smalldatetime] NULL ,
	[wkeEndDate] [smalldatetime] NULL ,
	[wkeEndDateOpen] [tinyint] NOT NULL ,
	[wkePeriod] AS (datediff(month,[wkeStartDate],[wkeEndDate])) ,
	[wkeOrgNameEng] [nvarchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[wkeOrgNameFra] [nvarchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[wkeOrgNameSpa] [nvarchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[wkeBnfNameEng] [nvarchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[wkeBnfNameFra] [nvarchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[wkeBnfNameSpa] [nvarchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[wkePrjTitleEng] [nvarchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[wkePrjTitleFra] [nvarchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[wkePrjTitleSpa] [nvarchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[wkePositionEng] [nvarchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[wkePositionFra] [nvarchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[wkePositionSpa] [nvarchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[wkeDescriptionEng] [ntext] COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[wkeDescriptionFra] [ntext] COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[wkeDescriptionSpa] [ntext] COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[wkeClientRefEng] [nvarchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[wkeClientRefFra] [nvarchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[wkeClientRefSpa] [nvarchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[TypeofWke] [tinyint] NULL ,
	[wkeLocationEng] [nvarchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[wkeLocationFra] [nvarchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[wkeLocationSpa] [nvarchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[wkeDonorEng] [nvarchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[wkeRefFirstName] [nvarchar] (255) COLLATE Latin1_General_CI_AS NULL ,
	[wkeRefLastName] [nvarchar] (255) COLLATE Latin1_General_CI_AS NULL,
	[wkeRefEmail] [nvarchar] (150) COLLATE Latin1_General_CI_AS NULL ,
	[wkeRefPhone] [varchar] (150) COLLATE Latin1_General_CI_AS NULL ,
	[wkeRefExtended] [tinyint] NULL,
	[id_ExpWkeOld] [int] NULL,
	[wkeProjectDescription] [ntext] COLLATE Latin1_General_CI_AS NULL ,
	[wkeRefName] [nvarchar] (255) COLLATE Latin1_General_CI_AS NULL ,
	[wkeRefPosition] [nvarchar] (255) COLLATE Latin1_General_CI_AS NULL ,
	[wkeInfoGroup] [nvarchar] (255) COLLATE Latin1_General_CI_AS NULL,
	[id_ExpWkeOriginal] [int] NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[lnkMmb_Exp_Select] (
	[id_MmbExpSelect] [int] IDENTITY (1, 1) NOT NULL ,
	[id_Member] [int] NOT NULL ,
	[id_Expert] [int] NOT NULL ,
	[Active] [tinyint] NULL ,
	[DownloadDate] [smalldatetime] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[lnkPrj_Pos] (
	[id_Project] [int] NOT NULL ,
	[id_Position] [int] NOT NULL ,
	[ppsCreateDate] [smalldatetime] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[lnkWke_Cou] (
	[id_WkeCou] [int] IDENTITY (1, 1) NOT NULL ,
	[id_ExpWke] [int] NOT NULL ,
	[id_Country] [smallint] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[lnkWke_Don] (
	[id_WkeDon] [int] IDENTITY (1, 1) NOT NULL ,
	[id_ExpWke] [int] NOT NULL ,
	[id_Organisation] [int] NULL ,
	[wkd_OtherNameEng] [nvarchar] (150) COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[wkd_OtherNameFra] [nvarchar] (150) COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[wkd_OtherNameSpa] [nvarchar] (150) COLLATE SQL_Latin1_General_CP1_CI_AI NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[lnkWke_Sct] (
	[id_WkeSct] [int] IDENTITY (1, 1) NOT NULL ,
	[id_ExpWke] [int] NOT NULL ,
	[id_Sector] [int] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[lnkWke_Srv] (
	[id_WkeSty] [int] NULL ,
	[id_ServiceType] [int] NULL ,
	[id_WrkExp] [int] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[lnk_Exp_Nationality] (
	[id_expNationality] [int] IDENTITY (1, 1) NOT NULL ,
	[id_Expert] [int] NOT NULL ,
	[id_Nationality] [smallint] NOT NULL ,
	[exnCreateDate] [smalldatetime] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[log_MmbExpSearch] (
	[id_SQLQuery] [int] IDENTITY (1, 1) NOT NULL ,
	[id_Member] [int] NULL ,
	[srchKeywords] [nvarchar] (400) COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[srchNationality] [varchar] (1400) COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[srchEducation] [varchar] (350) COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[srchNativeLng] [varchar] (700) COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[srchOtherLng] [varchar] (700) COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[srchCountries] [varchar] (1000) COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[srchRegions] [varchar] (150) COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[srchSectors] [varchar] (2000) COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[srchMainSectors] [varchar] (250) COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[srchDonors] [varchar] (150) COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[srchDB] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[srchDate] [smalldatetime] NULL ,
	[srchSQLQuery] [ntext] COLLATE SQL_Latin1_General_CP1_CI_AI NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[log_Session] (
	[id_UserSession] [bigint] IDENTITY (1, 1) NOT NULL ,
	[id_Session] [uniqueidentifier] NULL ,
	[id_AspSession] [bigint] NULL ,
	[id_User] [int] NULL ,
	[ussUserAgent] [nvarchar] (512) COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[ussIpAddress] [varchar] (16) COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[ussIpBase] AS (substring([ussIpAddress],1,charindex('.',[ussIpAddress],1))) ,
	[ussIpBase2] AS (substring([ussIpAddress],1,charindex('.',[ussIpAddress],5))) ,
	[ussCreateDate] [datetime] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[log_SessionEvent] (
	[id_SessionLog] [bigint] IDENTITY (1, 1) NOT NULL ,
	[id_UserSession] [bigint] NOT NULL ,
	[slgUrl] [varchar] (1000) COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[slgDate] [datetime] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[log_SessionEventOld] (
	[id_SessionLog] [bigint] NOT NULL ,
	[id_UserSession] [bigint] NOT NULL ,
	[slgUrl] [nvarchar] (1000) COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[slgDate] [datetime] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[log_SessionOld] (
	[id_UserSession] [bigint] NOT NULL ,
	[id_Session] [uniqueidentifier] NULL ,
	[id_AspSession] [bigint] NULL ,
	[id_User] [int] NULL ,
	[ussUserAgent] [nvarchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[ussIpAddress] [varchar] (16) COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[ussIpBase] AS (substring([ussIpAddress],1,charindex('.',[ussIpAddress],5))) ,
	[ussCreateDate] [smalldatetime] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[tbl_Continent] (
	[id_Continent] [int] NOT NULL ,
	[conDescriptionEng] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[db_NotVisible] [bit] NULL ,
	[conDescriptionSpa] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[conDescriptionFra] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[conDescription] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AI NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[tbl_Country] (
	[id_Country] [smallint] NOT NULL ,
	[id_GeoZone] [int] NOT NULL ,
	[couAbbreviation] [nvarchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[couNameEng] [nvarchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[couNameFra] [nvarchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[couNameSpa] [nvarchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[couNationalityEng] [nvarchar] (30) COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[couNationalityFra] [nvarchar] (30) COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[couNationalitySpa] [nvarchar] (30) COLLATE SQL_Latin1_General_CP1_CI_AI NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[tbl_CountryDev] (
	[id_Country] [int] NOT NULL ,
	[couDev] [tinyint] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[tbl_CountryGDP] (
	[id_Country] [int] NOT NULL ,
	[couGDP] [tinyint] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[tbl_Currency] (
	[id_Currency] [smallint] NOT NULL ,
	[curAbbreviation] [nvarchar] (4) COLLATE Latin1_General_CI_AS NULL ,
	[curDescriptionEng] [nvarchar] (255) COLLATE Latin1_General_CI_AS NULL ,
	[curDescriptionSpa] [nvarchar] (255) COLLATE Latin1_General_CI_AS NULL ,
	[curDescriptionFra] [nvarchar] (255) COLLATE Latin1_General_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[tbl_DocType] (
	[id_DocType] [int] NULL ,
	[dtName] [nvarchar] (50) COLLATE Latin1_General_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[tbl_Donors] (
	[id_Organisation] [int] NOT NULL ,
	[orgNameEng] [nvarchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[orgNameFra] [nvarchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[orgNameSpa] [nvarchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[orgAbbreviation] [nvarchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[orgMainDonor] [bit] NULL ,
	[orgVisibleDonor] [bit] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[tbl_EduSubjects] (
	[id_EduSubject] [int] NOT NULL ,
	[edsDescriptionEng] [nvarchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[edsDescriptionFra] [nvarchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[edsDescriptionSpa] [nvarchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[db_Order] [tinyint] NULL
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[tbl_EducationType] (
	[id_EduType] [int] NOT NULL ,
	[edtDescriptionEng] [nvarchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[edtDescriptionFra] [nvarchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[edtDescriptionSpa] [nvarchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[db_Order] [tinyint] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[tbl_Exp_Address] (
	[id_Address] [int] IDENTITY (1, 1) NOT NULL ,
	[id_AddressType] [smallint] NULL ,
	[id_Expert] [int] NULL ,
	[id_Country] [smallint] NULL ,
	[adrPhone] [nvarchar] (250) COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[adrFax] [nvarchar] (250) COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[adrEmail] [nvarchar] (250) COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[adrMobile] [nvarchar] (250) COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[adrStreetEng] [nvarchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[adrStreetFra] [nvarchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[adrStreetSpa] [nvarchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[adrPostCode] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[adrCityEng] [nvarchar] (150) COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[adrCityFra] [nvarchar] (150) COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[adrCitySpa] [nvarchar] (150) COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[adrWeb] [nvarchar] (250) COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[adrCreated] [smalldatetime] NULL ,
	[adrModified] [smalldatetime] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[tbl_ExpertStatus](
	[id_ExpertStatus] [smallint] NOT NULL,
	[exsTitleEng] [nvarchar](50) NULL,
	[exsTitleFra] [nvarchar](50) NULL,
	[exsTitleSpa] [nvarchar](50) NULL,
	[exsVisible] [tinyint] NULL,
	[exsCreateDate] [smalldatetime] NULL,
) ON [PRIMARY]
GO


CREATE TABLE [dbo].[tbl_Experts] (
	[id_Expert] [int] IDENTITY (100, 1) NOT NULL ,
	[id_ExpertOriginal] [int] NOT NULL ,
	[uid_Expert] [uniqueidentifier] NULL ,
	[id_User] [int] NOT NULL ,
	[id_ProfessionalStatus] [smallint] NULL ,
	[expProfYears] [tinyint] NULL ,
	[expKeyQualificationsEng] [ntext] COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[expKeyQualificationsFra] [ntext] COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[expKeyQualificationsSpa] [ntext] COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[expKeyQualificationsBrief] AS (substring([expKeyQualificationsEng],1,250)) ,
	[expCurrPositionEng] [nvarchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[expCurrPositionFra] [nvarchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[expCurrPositionSpa] [nvarchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[expProfessionEng] [nvarchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[expProfessionFra] [nvarchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[expProfessionSpa] [nvarchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[expMemberProfEng] [ntext] COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[expMemberProfFra] [ntext] COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[expMemberProfSpa] [ntext] COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[expPublicationsEng] [ntext] COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[expPublicationsFra] [ntext] COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[expPublicationsSpa] [ntext] COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[expReferencesEng] [ntext] COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[expReferencesFra] [ntext] COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[expReferencesSpa] [ntext] COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[expAvailabilityEng] [nvarchar] (400) COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[expAvailabilityFra] [nvarchar] (400) COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[expAvailabilitySpa] [nvarchar] (400) COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[expShortterm] [tinyint] NULL ,
	[expLongterm] [tinyint] NULL ,
	[Lng] [char] (3) COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[Email] [nvarchar] (150) COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[Phone] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[expIbfOnly] [bit] NULL ,
	[expIbfOnlyComments] [text] COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[expHidden] [bit] NOT NULL ,
	[expHiddenComments] [text] COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[expIncompleteCV] [bit] NULL ,
	[expToCompleteCVEmailSent] [bit] NULL ,
	[expToCompleteCVEmailDate] [smalldatetime] NULL ,
	[expToCompleteCVEmailText] [text] COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[expToConfirmCvEmailSent] [bit] NULL ,
	[expToConfirmCvEmailDate] [smalldatetime] NULL ,
	[expApproved] [bit] NOT NULL ,
	[expApprovedDate] [smalldatetime] NULL ,
	[expApprovedComments] [text] COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[expRemoved] [bit] NOT NULL ,
	[expRemovedDate] [smalldatetime] NULL ,
	[expRemovedComments] [text] COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[expDeleted] [bit] NOT NULL ,
	[expDeletedDate] [smalldatetime] NULL ,
	[expDeletedComments] [text] COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[BlackList] [tinyint] NULL ,
	[KgEncoded] [tinyint] NULL ,
	[KgCVFile] [nvarchar] (150) COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[expCreateDate] [smalldatetime] NULL ,
	[expLastUpdate] [smalldatetime] NULL ,
	[expRanking] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[expComments] [ntext] COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[BlackListMailSent] [bit] NULL ,
	[BlackListMailDate] [smalldatetime] NULL ,
	[BlackListMe] [bit] NULL ,
	[Subscribe] [bit] NULL ,
	[UnsubscribeDate] [smalldatetime] NULL ,
	[expFullTextIndex] [timestamp] NULL ,
	[expAccountEmailSent] [tinyint] NULL ,
	[expRegNumber] [nvarchar] (80) COLLATE Latin1_General_CI_AS NULL ,
	[expPreferences] [ntext] COLLATE Latin1_General_CI_AS NULL ,
	[expOtherSkills] [ntext] COLLATE Latin1_General_CI_AS NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[tbl_Gender] (
	[id_Gender] [tinyint] NOT NULL ,
	[genderNameEng] [nvarchar] (50) COLLATE Latin1_General_CI_AS NULL ,
	[genderNameFra] [nvarchar] (50) COLLATE Latin1_General_CI_AS NULL ,
	[genderNameSpa] [nvarchar] (50) COLLATE Latin1_General_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[tbl_GeoZone] (
	[id_GeoZone] [int] NOT NULL ,
	[id_Continent] [int] NOT NULL ,
	[Geo_Zone] [nvarchar] (255) COLLATE Latin1_General_CI_AS NULL ,
	[db_NotVisible] [bit] NOT NULL ,
	[db_Scroll] [smallint] NULL ,
	[Geo_ZoneEng] [nvarchar] (50) COLLATE Latin1_General_CI_AS NULL ,
	[Geo_ZoneFra] [nvarchar] (50) COLLATE Latin1_General_CI_AS NULL ,
	[Geo_ZoneSpa] [nvarchar] (50) COLLATE Latin1_General_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[tbl_LangLevel] (
	[id_LangLevel] [int] NOT NULL ,
	[lnlDescriptionEng] [nvarchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[lnlDescriptionFra] [nvarchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[lnlDescriptionSpa] [nvarchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AI NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[tbl_Languages] (
	[id_Language] [int] NOT NULL ,
	[lngNameEng] [nvarchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[lngNameFra] [nvarchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[lngNameSpa] [nvarchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[db_NotVisible] [bit] NULL ,
	[db_Order] [tinyint] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[tbl_LegalStatus] (
	[id_LegalStatus] [int] NOT NULL ,
	[LegalStatusName] [nvarchar] (50) COLLATE Latin1_General_CI_AS NULL ,
	[LegalStatusNameEng] [nvarchar] (50) COLLATE Latin1_General_CI_AS NULL ,
	[LegalStatusNameFra] [nvarchar] (50) COLLATE Latin1_General_CI_AS NULL ,
	[LegalStatusNameSpa] [nvarchar] (50) COLLATE Latin1_General_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[tbl_MainSectors] (
	[id_MainSector] [int] NOT NULL ,
	[mnsDescriptionEng] [nvarchar] (80) COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[mnsDescriptionFra] [nvarchar] (80) COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[mnsDescriptionSpa] [nvarchar] (80) COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[mnsShortEng] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[mnsAbbreviation] [nvarchar] (4) COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[db_Scroll] [smallint] NULL ,
	[db_NotVisible] [tinyint] NULL ,
	[mnsShortFra] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[mnsShortSpa] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AI NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[tbl_MaritalStatus] (
	[id_MaritalStatus] [tinyint] NOT NULL ,
	[mstNameEng] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[mstNameFra] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[mstNameSpa] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AI NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[tbl_Members] (
	[id_Member] [int] IDENTITY (1, 1) NOT NULL ,
	[id_User] [int] NULL ,
	[id_MemberType] [int] NULL ,
	[Email] [nvarchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[mmbEmailExtra] [varchar] (4000) COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[BlackList] [tinyint] NULL ,
	[Lng] [char] (3) COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[subscribe] [bit] NULL ,
	[mmbDevbusiness] [tinyint] NULL ,
	[Comments] [text] COLLATE SQL_Latin1_General_CP1_CI_AI NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[tbl_Month] (
	[id_Month] [tinyint] NOT NULL ,
	[MonthNameEng] [nvarchar] (50) COLLATE Latin1_General_CI_AS NULL ,
	[MonthNameFra] [nvarchar] (50) COLLATE Latin1_General_CI_AS NULL ,
	[MonthNameSpa] [nvarchar] (50) COLLATE Latin1_General_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[tbl_Native_Lng] (
	[id_Native] [int] IDENTITY (1, 1) NOT NULL ,
	[id_Language] [int] NULL ,
	[id_Expert] [int] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[tbl_PersonTitles] (
	[id_psnTitle] [tinyint] NOT NULL ,
	[ptlNameEng] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[ptlNameFra] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[ptlNameSpa] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AI NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[tbl_Persons] (
	[id_Person] [int] IDENTITY (100, 1) NOT NULL ,
	[id_Expert] [int] NULL ,
	[id_psnTitle] [tinyint] NULL ,
	[psnFirstNameEng] [nvarchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[psnFirstNameFra] [nvarchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[psnFirstNameSpa] [nvarchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[psnMiddleNameEng] [nvarchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[psnMiddleNameFra] [nvarchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[psnMiddleNameSpa] [nvarchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[psnLastNameEng] [nvarchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[psnLastNameFra] [nvarchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[psnLastNameSpa] [nvarchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[psnGender] [tinyint] NULL ,
	[psnBirthPlaceEng] [nvarchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[psnBirthPlaceFra] [nvarchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[psnBirthPlaceSpa] [nvarchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[psnBirthDate] [smalldatetime] NULL ,
	[id_MaritalStatus] [smallint] NULL ,
	[psnCreationDate] [smalldatetime] NULL ,
	[psLastUpdate] [smalldatetime] NULL ,
	[psnEncodedBy] [tinyint] NULL ,
	[psnComments] [nvarchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AI NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[tbl_Position] (
	[id_Position] [int] IDENTITY (1000, 1) NOT NULL ,
	[posTitle] [varchar] (400) COLLATE Latin1_General_CI_AS NULL ,
	[posIndex] [smallint] NULL ,
	[posAnnex] [varchar] (5) COLLATE Latin1_General_CI_AS NULL ,
	[id_Status] [tinyint] NULL ,
	[posWorkDays] [smallint] NULL ,
	[posDuration] [smallint] NULL ,
	[posDurationMeasure] [char] (1) COLLATE Latin1_General_CI_AS NULL ,
	[posCategory] [char] (1) COLLATE Latin1_General_CI_AS NULL ,
	[posSeniority] [varchar] (40) COLLATE Latin1_General_CI_AS NULL ,
	[posType] [varchar] (40) COLLATE Latin1_General_CI_AS NULL ,
	[posRequirements] [text] COLLATE Latin1_General_CI_AS NULL ,
	[posNumberExperts] [smallint] NULL ,
	[posDeadline] [smalldatetime] NULL ,
	[posStartDate] [smalldatetime] NULL ,
	[posEndDate] [smalldatetime] NULL ,
	[posCreateDate] [smalldatetime] NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[tbl_ProfessionalStatus] (
	[id_ProfessionalStatus] [smallint] NOT NULL ,
	[pfsTitle] [varchar] (100) COLLATE Latin1_General_CI_AS NULL ,
	[pfsCreateDate] [smalldatetime] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[tbl_Project] (
	[id_Project] [int] IDENTITY (1000, 1) NOT NULL ,
	[prjReference] [varchar] (30) COLLATE Latin1_General_CI_AS NULL ,
	[prjShortName] [varchar] (60) COLLATE Latin1_General_CI_AS NULL ,
	[prjTitle] [varchar] (400) COLLATE Latin1_General_CI_AS NULL ,
	[id_ProjectStatus] [smallint] NULL ,
	[prjLocation] [varchar] (100) COLLATE Latin1_General_CI_AS NULL ,
	[prjDescription] [text] COLLATE Latin1_General_CI_AS NULL ,
	[prjDeadline] [smalldatetime] NULL ,
	[prjCreateDate] [smalldatetime] NULL ,
	[prjModifyDate] [smalldatetime] NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[tbl_ProjectStatus] (
	[id_ProjectStatus] [smallint] NOT NULL ,
	[prsTitle] [varchar] (100) COLLATE Latin1_General_CI_AS NULL ,
	[prsCreateDate] [smalldatetime] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[tbl_Sectors] (
	[id_Sector] [smallint] NOT NULL ,
	[sctDescriptionEng] [nvarchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[sctDescriptionFra] [nvarchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[sctDescriptionSpa] [nvarchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[id_MainSector] [smallint] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[tbl_ServiceType] (
	[id_ServiceType] [int] NULL ,
	[styDescriptionEng] [nvarchar] (255) COLLATE Latin1_General_CI_AS NULL ,
	[styDescriptionFra] [nvarchar] (255) COLLATE Latin1_General_CI_AS NULL ,
	[styDescriptionSpa] [nvarchar] (255) COLLATE Latin1_General_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[tbl_StatusCV] (
	[id_Status] [tinyint] NOT NULL ,
	[stsNameEng] [varchar] (50) COLLATE Latin1_General_CI_AS NULL ,
	[stsNameFra] [varchar] (50) COLLATE Latin1_General_CI_AS NULL ,
	[stsNameSpa] [varchar] (50) COLLATE Latin1_General_CI_AS NULL ,
	[stsCreateDate] [smalldatetime] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[tbl_UserType] (
	[id_UserType] [smallint] NOT NULL ,
	[ustName] [varchar] (20) COLLATE Latin1_General_CI_AS NULL ,
	[ustDescription] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AI NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[tbl_Users] (
	[id_User] [int] IDENTITY (200, 1) NOT NULL ,
	[id_UserType] [smallint] NULL ,
	[UserName] [nvarchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[PassWord] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[FullName] [nvarchar] (100) COLLATE Latin1_General_CI_AS NULL ,
	[CreateDate] [smalldatetime] NULL ,
	[usrIpSecurity] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[usrFastLoginID] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[usrComments] [varchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AI NULL ,
	[usrCreateDate] [smalldatetime] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[tmp_Numbers] (
	[Number] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [tbl_Documents] (
	[id_Document] [int] IDENTITY (1000, 1) NOT NULL ,
	[uid_Document] [uniqueidentifier] NULL CONSTRAINT [DF_tbl_Documents_uid_Document] DEFAULT (newid()),
	[docTitle] [nvarchar] (255) COLLATE Latin1_General_CI_AS NULL ,
	[id_DocType] [smallint] NULL ,
	[docType] [nvarchar] (255) COLLATE Latin1_General_CI_AS NULL ,
	[docText] [ntext] COLLATE Latin1_General_CI_AS NULL ,
	[docComments] [ntext] COLLATE Latin1_General_CI_AS NULL ,
	[docPath] [nvarchar] (255) COLLATE Latin1_General_CI_AS NULL ,
	[docImage] [image] NULL ,
	[docImageSize] [int] NULL ,
	[docCreated] [smalldatetime] NULL ,
	[docUpdated] [smalldatetime] NULL ,
	CONSTRAINT [PK_tbl_Documents] PRIMARY KEY  CLUSTERED 
	(
		[id_Document]
	)  ON [PRIMARY] 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [lnkExp_Doc] (
	[id_Expert] [int] NOT NULL ,
	[id_Document] [int] NOT NULL ,
	[edcCreateDate] [smalldatetime] NULL CONSTRAINT [DF_lnkExp_Doc_edcCreateDate] DEFAULT (getdate()),
	CONSTRAINT [PK_lnkExp_Doc] PRIMARY KEY  CLUSTERED 
	(
		[id_Expert],
		[id_Document]
	)  ON [PRIMARY] 
) ON [PRIMARY]
GO

CREATE TABLE [lnkMmb_Exp_Query] (
	[id_Member] [int] NOT NULL ,
	[id_Expert] [int] NOT NULL ,
	[id_Query] [int] NOT NULL ,
	[SelectedDate] [smalldatetime] NULL CONSTRAINT [DF_lnkMmb_Exp_Query_SelectedDate] DEFAULT (getdate()),
	CONSTRAINT [PK_lnkMmb_Exp_Query_Select] PRIMARY KEY  CLUSTERED 
	(
		[id_Member],
		[id_Expert],
		[id_Query]
	)  ON [PRIMARY] 
) ON [PRIMARY]
GO


CREATE TABLE [tbl_ExpertsLanguage] (
	[id_Expert] [int] NOT NULL ,
	[id_Expert2] [int] NOT NULL ,
	[exlCreateDate] [smalldatetime] NULL ,
	[exlComment] [text] COLLATE Latin1_General_CI_AS NULL ,
	CONSTRAINT [PK_ExpertsLanguage] PRIMARY KEY  CLUSTERED 
	(
		[id_Expert],
		[id_Expert2]
	)  ON [PRIMARY] 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO


ALTER TABLE [dbo].[lnkExp_Pos] WITH NOCHECK ADD 
	CONSTRAINT [PK_lnkExp_Pos] PRIMARY KEY  CLUSTERED 
	(
		[id_Expert],
		[id_Position]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[lnkExp_Prj] WITH NOCHECK ADD 
	CONSTRAINT [PK_lnkExp_Prj] PRIMARY KEY  CLUSTERED 
	(
		[id_Expert],
		[id_Project]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[lnkExp_StatusCV] WITH NOCHECK ADD 
	CONSTRAINT [PK_tbl_ExpertStatus] PRIMARY KEY  CLUSTERED 
	(
		[id_Expert],
		[id_Status]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[lnkPrj_Pos] WITH NOCHECK ADD 
	CONSTRAINT [PK_lnkPrj_Pos] PRIMARY KEY  CLUSTERED 
	(
		[id_Project],
		[id_Position]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[log_MmbExpSearch] WITH NOCHECK ADD 
	CONSTRAINT [PK_log_MmbExpSearch] PRIMARY KEY  CLUSTERED 
	(
		[id_SQLQuery]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[log_Session] WITH NOCHECK ADD 
	CONSTRAINT [PK_log_Session] PRIMARY KEY  CLUSTERED 
	(
		[id_UserSession]
	) WITH  FILLFACTOR = 75  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[log_SessionEvent] WITH NOCHECK ADD 
	CONSTRAINT [PK_log_SessionEvent0] PRIMARY KEY  CLUSTERED 
	(
		[id_SessionLog]
	) WITH  FILLFACTOR = 75  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[log_SessionEventOld] WITH NOCHECK ADD 
	CONSTRAINT [PK_log_SessionEventOld] PRIMARY KEY  CLUSTERED 
	(
		[id_SessionLog]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[log_SessionOld] WITH NOCHECK ADD 
	CONSTRAINT [PK_log_SessionOld] PRIMARY KEY  CLUSTERED 
	(
		[id_UserSession]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tbl_Continent] WITH NOCHECK ADD 
	CONSTRAINT [PK_tbl_Continent] PRIMARY KEY  CLUSTERED 
	(
		[id_Continent]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tbl_Country] WITH NOCHECK ADD 
	CONSTRAINT [PK_tbl_Country1] PRIMARY KEY  CLUSTERED 
	(
		[id_Country]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tbl_Currency] WITH NOCHECK ADD 
	CONSTRAINT [PK_tbl_Currency] PRIMARY KEY  CLUSTERED 
	(
		[id_Currency]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tbl_Donors] WITH NOCHECK ADD 
	CONSTRAINT [PK_tbl_Donors] PRIMARY KEY  CLUSTERED 
	(
		[id_Organisation]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tbl_EduSubjects] WITH NOCHECK ADD 
	CONSTRAINT [PK_tbl_EduSubjects] PRIMARY KEY  CLUSTERED 
	(
		[id_EduSubject]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tbl_EducationType] WITH NOCHECK ADD 
	CONSTRAINT [PK_tbl_EducationType] PRIMARY KEY  CLUSTERED 
	(
		[id_EduType]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tbl_ExpertStatus] WITH NOCHECK ADD 
	CONSTRAINT [PK_tbl_ExpertStatus_1] PRIMARY KEY  CLUSTERED 
	(
		[id_ExpertStatus]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tbl_Experts] WITH NOCHECK ADD 
	CONSTRAINT [PK_tbl_Experts1] PRIMARY KEY  CLUSTERED 
	(
		[id_Expert]
	)  ON [PRIMARY] 
GO

if (select DATABASEPROPERTY(DB_NAME(), N'IsFullTextEnabled')) <> 1 
exec sp_fulltext_database N'enable' 

GO

if not exists (select * from dbo.sysfulltextcatalogs where name = N'Experts')
exec sp_fulltext_catalog N'Experts', N'create' 

GO

exec sp_fulltext_table N'[dbo].[tbl_Experts]', N'create', N'Experts', N'PK_tbl_Experts1'
GO

exec sp_fulltext_column N'[dbo].[tbl_Experts]', N'expKeyQualificationsEng', N'add', 1033  
GO

exec sp_fulltext_column N'[dbo].[tbl_Experts]', N'expCurrPositionEng', N'add', 1033  
GO

exec sp_fulltext_column N'[dbo].[tbl_Experts]', N'expProfessionEng', N'add', 1033  
GO

exec sp_fulltext_column N'[dbo].[tbl_Experts]', N'expMemberProfEng', N'add', 1033  
GO

exec sp_fulltext_column N'[dbo].[tbl_Experts]', N'expPublicationsEng', N'add', 1033  
GO

exec sp_fulltext_column N'[dbo].[tbl_Experts]', N'expReferencesEng', N'add', 1033  
GO

exec sp_fulltext_column N'[dbo].[tbl_Experts]', N'expAvailabilityEng', N'add', 1033  
GO

exec sp_fulltext_table N'[dbo].[tbl_Experts]', N'activate'  
GO

ALTER TABLE [dbo].[tbl_LangLevel] WITH NOCHECK ADD 
	CONSTRAINT [PK_tbl_LangLevel] PRIMARY KEY  CLUSTERED 
	(
		[id_LangLevel]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tbl_Languages] WITH NOCHECK ADD 
	CONSTRAINT [PK_tbl_Languages] PRIMARY KEY  CLUSTERED 
	(
		[id_Language]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tbl_MainSectors] WITH NOCHECK ADD 
	CONSTRAINT [PK_tbl_MainSectors] PRIMARY KEY  CLUSTERED 
	(
		[id_MainSector]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tbl_MaritalStatus] WITH NOCHECK ADD 
	CONSTRAINT [PK_tbl_MaritalStatus] PRIMARY KEY  CLUSTERED 
	(
		[id_MaritalStatus]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tbl_Members] WITH NOCHECK ADD 
	CONSTRAINT [PK_tbl_Members1] PRIMARY KEY  CLUSTERED 
	(
		[id_Member]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tbl_PersonTitles] WITH NOCHECK ADD 
	CONSTRAINT [PK_tbl_PersonTitles] PRIMARY KEY  CLUSTERED 
	(
		[id_psnTitle]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tbl_Position] WITH NOCHECK ADD 
	CONSTRAINT [PK_tbl_Position] PRIMARY KEY  CLUSTERED 
	(
		[id_Position]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tbl_ProfessionalStatus] WITH NOCHECK ADD 
	CONSTRAINT [PK_tbl_ProfessionalStatus] PRIMARY KEY  CLUSTERED 
	(
		[id_ProfessionalStatus]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tbl_Project] WITH NOCHECK ADD 
	CONSTRAINT [PK_tbl_Project] PRIMARY KEY  CLUSTERED 
	(
		[id_Project]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tbl_ProjectStatus] WITH NOCHECK ADD 
	CONSTRAINT [PK_tbl_ProjectStatus] PRIMARY KEY  CLUSTERED 
	(
		[id_ProjectStatus]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tbl_Sectors] WITH NOCHECK ADD 
	CONSTRAINT [PK_tbl_Sectors1] PRIMARY KEY  CLUSTERED 
	(
		[id_Sector]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tbl_StatusCV] WITH NOCHECK ADD 
	CONSTRAINT [PK_tbl_Status_Expert] PRIMARY KEY  CLUSTERED 
	(
		[id_Status]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tbl_UserType] WITH NOCHECK ADD 
	CONSTRAINT [PK_tbl_UserType] PRIMARY KEY  CLUSTERED 
	(
		[id_UserType]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tbl_Users] WITH NOCHECK ADD 
	CONSTRAINT [PK_tbl_Users] PRIMARY KEY  CLUSTERED 
	(
		[id_User]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tmp_Numbers] WITH NOCHECK ADD 
	CONSTRAINT [PK_tmp_Numbers] PRIMARY KEY  CLUSTERED 
	(
		[Number]
	)  ON [PRIMARY] 
GO

 CREATE  CLUSTERED  INDEX [IX_lnkExp_Edu1] ON [dbo].[lnkExp_Edu]([id_Expert]) ON [PRIMARY]
GO

 CREATE  CLUSTERED  INDEX [IX_lnkExp_Lan1] ON [dbo].[lnkExp_Lan]([id_Expert]) ON [PRIMARY]
GO

 CREATE  CLUSTERED  INDEX [IX_lnkExp_RankCou1] ON [dbo].[lnkExp_RankCou]([id_Expert]) ON [PRIMARY]
GO

 CREATE  CLUSTERED  INDEX [IX_lnkExp_RankSct1] ON [dbo].[lnkExp_RankSct]([id_Expert]) ON [PRIMARY]
GO

 CREATE  CLUSTERED  INDEX [IX_lnkExp_Wke1] ON [dbo].[lnkExp_Wke]([id_Expert]) ON [PRIMARY]
GO

 CREATE  CLUSTERED  INDEX [IX_lnkWke_Cou1] ON [dbo].[lnkWke_Cou]([id_ExpWke]) ON [PRIMARY]
GO

 CREATE  CLUSTERED  INDEX [IX_lnkWke_Don1] ON [dbo].[lnkWke_Don]([id_ExpWke]) ON [PRIMARY]
GO

 CREATE  CLUSTERED  INDEX [IX_lnkWke_Sct1] ON [dbo].[lnkWke_Sct]([id_ExpWke]) ON [PRIMARY]
GO

 CREATE  CLUSTERED  INDEX [IX_lnk_Exp_Nationality1] ON [dbo].[lnk_Exp_Nationality]([id_Expert]) ON [PRIMARY]
GO

 CREATE  CLUSTERED  INDEX [IX_tbl_Exp_Address1] ON [dbo].[tbl_Exp_Address]([id_Expert]) ON [PRIMARY]
GO

 CREATE  CLUSTERED  INDEX [IX_tbl_Native_Lng1] ON [dbo].[tbl_Native_Lng]([id_Expert]) ON [PRIMARY]
GO

 CREATE  CLUSTERED  INDEX [IX_tbl_Persons1] ON [dbo].[tbl_Persons]([id_Expert]) ON [PRIMARY]
GO

ALTER TABLE [dbo].[lnkExp_Edu] WITH NOCHECK ADD 
	CONSTRAINT [PK_lnkExp_Edu0] PRIMARY KEY  NONCLUSTERED 
	(
		[id_ExpEdu]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[lnkExp_Lan] WITH NOCHECK ADD 
	CONSTRAINT [PK_lnkExp_Lan0] PRIMARY KEY  NONCLUSTERED 
	(
		[id_ExpLan]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[lnkExp_Pos] WITH NOCHECK ADD 
	CONSTRAINT [DF_lnkExp_Pos_epsCreateDate] DEFAULT (getdate()) FOR [epsCreateDate]
GO

ALTER TABLE [dbo].[lnkExp_Prj] WITH NOCHECK ADD 
	CONSTRAINT [DF_lnkExp_Prj_CreateDate] DEFAULT (getdate()) FOR [epjCreateDate]
GO

ALTER TABLE [dbo].[lnkExp_RankCou] WITH NOCHECK ADD 
	CONSTRAINT [PK_lnkExp_RankCou0] PRIMARY KEY  NONCLUSTERED 
	(
		[id_ExpRankCou]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[lnkExp_RankSct] WITH NOCHECK ADD 
	CONSTRAINT [PK_lnkExp_RankSct0] PRIMARY KEY  NONCLUSTERED 
	(
		[id_ExpRankSct]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[lnkExp_Wke] WITH NOCHECK ADD 
	CONSTRAINT [DF_lnkExp_Wke_wkeEndDateOpen] DEFAULT (0) FOR [wkeEndDateOpen],
	CONSTRAINT [DF_lnkExp_Wke_wkeRefExtended] DEFAULT (0) FOR [wkeRefExtended],
	CONSTRAINT [PK_lnkExp_Wke0] PRIMARY KEY  NONCLUSTERED 
	(
		[id_ExpWke]
	)  ON [PRIMARY] 
GO

if (select DATABASEPROPERTY(DB_NAME(), N'IsFullTextEnabled')) <> 1 
exec sp_fulltext_database N'enable' 

GO

if not exists (select * from dbo.sysfulltextcatalogs where name = N'Experts')
exec sp_fulltext_catalog N'Experts', N'create' 

GO

exec sp_fulltext_table N'[dbo].[lnkExp_Wke]', N'create', N'Experts', N'PK_lnkExp_Wke0'
GO

exec sp_fulltext_column N'[dbo].[lnkExp_Wke]', N'wkeOrgNameEng', N'add', 1033  
GO

exec sp_fulltext_column N'[dbo].[lnkExp_Wke]', N'wkeBnfNameEng', N'add', 1033  
GO

exec sp_fulltext_column N'[dbo].[lnkExp_Wke]', N'wkePrjTitleEng', N'add', 1033  
GO

exec sp_fulltext_column N'[dbo].[lnkExp_Wke]', N'wkePositionEng', N'add', 1033  
GO

exec sp_fulltext_column N'[dbo].[lnkExp_Wke]', N'wkeDescriptionEng', N'add', 1033  
GO

exec sp_fulltext_column N'[dbo].[lnkExp_Wke]', N'wkeClientRefEng', N'add', 1033  
GO

exec sp_fulltext_column N'[dbo].[lnkExp_Wke]', N'wkeLocationEng', N'add', 1033  
GO

exec sp_fulltext_column N'[dbo].[lnkExp_Wke]', N'wkeDonorEng', N'add', 1033  
GO

exec sp_fulltext_table N'[dbo].[lnkExp_Wke]', N'activate'  
GO


if (select DATABASEPROPERTY(DB_NAME(), N'IsFullTextEnabled')) <> 1 
exec sp_fulltext_database N'enable' 

GO

if not exists (select * from dbo.sysfulltextcatalogs where name = N'Experts')
exec sp_fulltext_catalog N'Experts', N'create' 

GO

exec sp_fulltext_table N'[dbo].[lnkExp_Edu]', N'create', N'Experts', N'PK_lnkExp_Edu0'
GO

exec sp_fulltext_column N'[dbo].[lnkExp_Edu]', N'id_EduSubject1Eng', N'add', 1033  
GO

exec sp_fulltext_column N'[dbo].[lnkExp_Edu]', N'eduDiploma1Eng', N'add', 1033  
GO

exec sp_fulltext_column N'[dbo].[lnkExp_Edu]', N'InstNameEng', N'add', 1033  
GO

exec sp_fulltext_column N'[dbo].[lnkExp_Edu]', N'InstLocationEng', N'add', 1033  
GO

exec sp_fulltext_column N'[dbo].[lnkExp_Edu]', N'eduOtherEng', N'add', 1033  
GO

exec sp_fulltext_table N'[dbo].[lnkExp_Edu]', N'activate'  
GO

ALTER TABLE [dbo].[lnkMmb_Exp_Select] WITH NOCHECK ADD 
	CONSTRAINT [DF_lnkMmb_Exp_Select_Active] DEFAULT (0) FOR [Active],
	CONSTRAINT [DF_lnkMmb_Exp_Select_DownloadDate] DEFAULT (getdate()) FOR [DownloadDate],
	CONSTRAINT [PK_lnkMmb_Exp_Select0] PRIMARY KEY  NONCLUSTERED 
	(
		[id_MmbExpSelect]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[lnkPrj_Pos] WITH NOCHECK ADD 
	CONSTRAINT [DF_lnkPrj_Pos_ppsCreateDate] DEFAULT (getdate()) FOR [ppsCreateDate]
GO

ALTER TABLE [dbo].[lnkWke_Cou] WITH NOCHECK ADD 
	CONSTRAINT [PK_lnkWke_Cou0] PRIMARY KEY  NONCLUSTERED 
	(
		[id_WkeCou]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[lnkWke_Don] WITH NOCHECK ADD 
	CONSTRAINT [PK_lnkWke_Don0] PRIMARY KEY  NONCLUSTERED 
	(
		[id_WkeDon]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[lnkWke_Sct] WITH NOCHECK ADD 
	CONSTRAINT [PK_lnkWke_Sct0] PRIMARY KEY  NONCLUSTERED 
	(
		[id_WkeSct]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[lnk_Exp_Nationality] WITH NOCHECK ADD 
	CONSTRAINT [DF_lnk_Exp_Nationality_exnCreateDate] DEFAULT (getdate()) FOR [exnCreateDate],
	CONSTRAINT [PK_lnk_Exp_Nationality0] PRIMARY KEY  NONCLUSTERED 
	(
		[id_expNationality]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[log_MmbExpSearch] WITH NOCHECK ADD 
	CONSTRAINT [DF_log_MmbExpSearch_srchDate] DEFAULT (getdate()) FOR [srchDate]
GO

ALTER TABLE [dbo].[log_Session] WITH NOCHECK ADD 
	CONSTRAINT [DF_log_Session_ussCreateDate] DEFAULT (getdate()) FOR [ussCreateDate]
GO

ALTER TABLE [dbo].[log_SessionEvent] WITH NOCHECK ADD 
	CONSTRAINT [DF_log_SessionEvent_slgDate] DEFAULT (getdate()) FOR [slgDate]
GO

ALTER TABLE [dbo].[log_SessionEventOld] WITH NOCHECK ADD 
	CONSTRAINT [DF_log_SessionEventOld_slgDate] DEFAULT (getdate()) FOR [slgDate]
GO

ALTER TABLE [dbo].[log_SessionOld] WITH NOCHECK ADD 
	CONSTRAINT [DF_log_SessionOld_ussCreateDate] DEFAULT (getdate()) FOR [ussCreateDate]
GO

ALTER TABLE [dbo].[tbl_Exp_Address] WITH NOCHECK ADD 
	CONSTRAINT [DF_tbl_Exp_Address_adrCreated] DEFAULT (getdate()) FOR [adrCreated],
	CONSTRAINT [PK_tbl_Exp_Address0] PRIMARY KEY  NONCLUSTERED 
	(
		[id_Address]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tbl_ExpertStatus] WITH NOCHECK ADD 
	CONSTRAINT [DF_tbl_ExpertStatus_exsCreateDate] DEFAULT (getdate()) FOR [exsCreateDate]
GO

ALTER TABLE [dbo].[tbl_Experts] WITH NOCHECK ADD 
	CONSTRAINT [DF_tbl_Experts_id_ExpertUpdated] DEFAULT (0) FOR [id_ExpertOriginal],
	CONSTRAINT [DF_tbl_Experts_uid_Expert] DEFAULT (newid()) FOR [uid_Expert],
	CONSTRAINT [DF_tbl_Experts_expShortterm] DEFAULT (1) FOR [expShortterm],
	CONSTRAINT [DF_tbl_Experts_expLongterm] DEFAULT (1) FOR [expLongterm],
	CONSTRAINT [DF_tbl_Experts_expIbfOnly] DEFAULT (0) FOR [expIbfOnly],
	CONSTRAINT [DF_tbl_Experts_expHidden] DEFAULT (0) FOR [expHidden],
	CONSTRAINT [DF_tbl_Experts_expToConfirmEmailSent] DEFAULT (0) FOR [expToConfirmCvEmailSent],
	CONSTRAINT [DF_tbl_Experts_expApproved] DEFAULT (0) FOR [expApproved],
	CONSTRAINT [DF_tbl_Experts_expRemoved] DEFAULT (0) FOR [expRemoved],
	CONSTRAINT [DF_tbl_Experts_expDeleted] DEFAULT (0) FOR [expDeleted],
	CONSTRAINT [DF_tbl_Experts_BlackList] DEFAULT (1) FOR [BlackList],
	CONSTRAINT [DF_tbl_Experts_KgEncoded] DEFAULT (0) FOR [KgEncoded],
	CONSTRAINT [DF_tbl_Experts_expCreateDate] DEFAULT (getdate()) FOR [expCreateDate],
	CONSTRAINT [DF_tbl_Experts_BlackListMailSent] DEFAULT (0) FOR [BlackListMailSent],
	CONSTRAINT [DF_tbl_Experts_BlackListMe] DEFAULT (0) FOR [BlackListMe],
	CONSTRAINT [DF_tbl_Experts_Subscribe] DEFAULT (0) FOR [Subscribe],
	CONSTRAINT [DF_tbl_Experts_expAccountEmailSent] DEFAULT (0) FOR [expAccountEmailSent]
GO

ALTER TABLE [dbo].[tbl_Languages] WITH NOCHECK ADD 
	CONSTRAINT [DF_tbl_Languages_db_Order] DEFAULT (100) FOR [db_Order]
GO

ALTER TABLE [dbo].[tbl_Members] WITH NOCHECK ADD 
	CONSTRAINT [DF_tbl_Members_BlackList] DEFAULT (0) FOR [BlackList],
	CONSTRAINT [DF_tbl_Members_Lng] DEFAULT ('Eng') FOR [Lng],
	CONSTRAINT [DF_tbl_Members_subscribe] DEFAULT (0) FOR [subscribe],
	CONSTRAINT [DF_tbl_Members_mmbDevbusiness] DEFAULT (1) FOR [mmbDevbusiness]
GO

ALTER TABLE [dbo].[tbl_Native_Lng] WITH NOCHECK ADD 
	CONSTRAINT [PK_tbl_Native_Lng0] PRIMARY KEY  NONCLUSTERED 
	(
		[id_Native]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tbl_Persons] WITH NOCHECK ADD 
	CONSTRAINT [DF_tbl_Persons_psnCreationDate] DEFAULT (getdate()) FOR [psnCreationDate],
	CONSTRAINT [PK_tbl_Persons0] PRIMARY KEY  NONCLUSTERED 
	(
		[id_Person]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tbl_Position] WITH NOCHECK ADD 
	CONSTRAINT [DF_tbl_Position_posCreateDate] DEFAULT (getdate()) FOR [posCreateDate]
GO

ALTER TABLE [dbo].[tbl_ProfessionalStatus] WITH NOCHECK ADD 
	CONSTRAINT [DF_tbl_ProfessionalStatus_pfsCreateDate] DEFAULT (getdate()) FOR [pfsCreateDate]
GO

ALTER TABLE [dbo].[tbl_Project] WITH NOCHECK ADD 
	CONSTRAINT [DF_tbl_Project_prjCreateDate] DEFAULT (getdate()) FOR [prjCreateDate]
GO

ALTER TABLE [dbo].[tbl_ProjectStatus] WITH NOCHECK ADD 
	CONSTRAINT [DF_tbl_ProjectStatus_prsCreateDate] DEFAULT (getdate()) FOR [prsCreateDate]
GO

ALTER TABLE [dbo].[tbl_StatusCV] WITH NOCHECK ADD 
	CONSTRAINT [DF_tbl_Status_Expert_stsCreateDate] DEFAULT (getdate()) FOR [stsCreateDate]
GO

ALTER TABLE [dbo].[tbl_Users] WITH NOCHECK ADD 
	CONSTRAINT [DF_tbl_Users_CreateDate] DEFAULT (getdate()) FOR [CreateDate],
	CONSTRAINT [DF_tbl_Users_usrCreateDate] DEFAULT (getdate()) FOR [usrCreateDate]
GO

 CREATE  INDEX [IX_lnkExp_Edu2] ON [dbo].[lnkExp_Edu]([id_EduSubject]) ON [PRIMARY]
GO

 CREATE  INDEX [IX_lnkExp_Lan2] ON [dbo].[lnkExp_Lan]([id_Language]) ON [PRIMARY]
GO

 CREATE  INDEX [IX_lnkExp_RankCou2] ON [dbo].[lnkExp_RankCou]([id_Country]) ON [PRIMARY]
GO

 CREATE  INDEX [IX_lnkExp_RankSct2] ON [dbo].[lnkExp_RankSct]([id_Sector]) ON [PRIMARY]
GO

 CREATE  INDEX [IX_lnkWke_Cou2] ON [dbo].[lnkWke_Cou]([id_Country]) ON [PRIMARY]
GO

 CREATE  INDEX [IX_lnkWke_Don2] ON [dbo].[lnkWke_Don]([id_Organisation]) ON [PRIMARY]
GO

 CREATE  INDEX [IX_lnkWke_Sct0] ON [dbo].[lnkWke_Sct]([id_Sector]) ON [PRIMARY]
GO

 CREATE  INDEX [IX_lnk_Exp_Nationality2] ON [dbo].[lnk_Exp_Nationality]([id_Nationality]) ON [PRIMARY]
GO

 CREATE  INDEX [IX_log_Session_SessionId] ON [dbo].[log_Session]([id_Session]) WITH  FILLFACTOR = 75 ON [PRIMARY]
GO

 CREATE  INDEX [IX_log_Session_UserId] ON [dbo].[log_Session]([id_User]) ON [PRIMARY]
GO

 CREATE  INDEX [log_SessionEvent1] ON [dbo].[log_SessionEvent]([id_UserSession]) WITH  FILLFACTOR = 75 ON [PRIMARY]
GO

 CREATE  INDEX [IX_tbl_Country2] ON [dbo].[tbl_Country]([id_GeoZone]) ON [PRIMARY]
GO

 CREATE  INDEX [IX_tbl_Exp_Address] ON [dbo].[tbl_Exp_Address]([id_Country]) ON [PRIMARY]
GO

 CREATE  INDEX [IX_tbl_Experts2] ON [dbo].[tbl_Experts]([id_User]) ON [PRIMARY]
GO

 CREATE  INDEX [IX_tbl_Native_Lng2] ON [dbo].[tbl_Native_Lng]([id_Language]) ON [PRIMARY]
GO

 CREATE  INDEX [IX_tbl_Sectors2] ON [dbo].[tbl_Sectors]([id_MainSector]) ON [PRIMARY]
GO

ALTER TABLE [dbo].[lnk_Exp_Nationality] ADD 
	CONSTRAINT [FK_lnk_Exp_Nationality_tbl_Experts] FOREIGN KEY 
	(
		[id_Expert]
	) REFERENCES [dbo].[tbl_Experts] (
		[id_Expert]
	) ON DELETE CASCADE 
GO

ALTER TABLE [dbo].[log_SessionEvent] ADD 
	CONSTRAINT [FK_log_SessionEvent_log_Session] FOREIGN KEY 
	(
		[id_UserSession]
	) REFERENCES [dbo].[log_Session] (
		[id_UserSession]
	) ON DELETE CASCADE  NOT FOR REPLICATION 
GO

ALTER TABLE [dbo].[tbl_Persons] ADD 
	CONSTRAINT [FK_tbl_Persons_tbl_Experts] FOREIGN KEY 
	(
		[id_Expert]
	) REFERENCES [dbo].[tbl_Experts] (
		[id_Expert]
	)
GO

alter table [dbo].[tbl_Persons] nocheck constraint [FK_tbl_Persons_tbl_Experts]
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE VIEW dbo.uvw_Country
AS
SELECT 
C.id_Country,
C.couAbbreviation, 
C.couNameEng,
C.couNameFra,
C.couNameSpa,
C.id_GeoZone id_Region,
C.id_GeoZone
FROM tbl_Country C
WHERE C.id_Country<>524
AND (C.id_Country<1000 OR C.id_Country>1025)

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE VIEW dbo.uvw_Experts
AS
SELECT 
P.id_Person, 
P.id_psnTitle, 
CASE WHEN E.Lng='Spa' THEN ISNULL(PT.ptlNameSpa, ISNULL(PT.ptlNameEng, '')) WHEN E.Lng='Fra' THEN ISNULL(PT.ptlNameFra, ISNULL(PT.ptlNameEng, '')) ELSE ISNULL(PT.ptlNameEng, '')  END AS ptlName,
CASE WHEN E.Lng='Spa' THEN ISNULL(P.psnFirstNameSpa, ISNULL(P.psnFirstNameEng, '')) WHEN E.Lng='Fra' THEN ISNULL(P.psnFirstNameFra, ISNULL(P.psnFirstNameEng,'')) ELSE ISNULL(P.psnFirstNameEng,'') END AS psnFirstName,
CASE WHEN E.Lng='Spa' THEN ISNULL(P.psnMiddleNameSpa, ISNULL(P.psnMiddleNameEng, '')) WHEN E.Lng='Fra' THEN ISNULL(P.psnMiddleNameFra, ISNULL(P.psnMiddleNameEng,'')) ELSE ISNULL(P.psnMiddleNameEng,'') END AS psnMiddleName,
CASE WHEN E.Lng='Spa' THEN ISNULL(P.psnLastNameSpa, ISNULL(P.psnLastNameEng, '')) WHEN E.Lng='Fra' THEN ISNULL(P.psnLastNameFra, ISNULL(P.psnLastNameEng,'')) ELSE ISNULL(P.psnLastNameEng,'') END AS psnLastName,
ISNULL(P.psnBirthPlaceEng, '') psnBirthPlace, 
P.psnBirthDate, 
P.psnGender,
P.id_MaritalStatus, 
E.id_Expert, 
E.uid_Expert, 
E.id_ExpertOriginal, 
E.id_User, 
E.Lng, 
E.Email, 
E.Phone, 
E.KgEncoded, 
E.KgCVFile, 
E.expProfessionEng expProfession,
E.expAvailabilityEng expAvailability,
E.expProfYears,
E.expShortterm,
E.expLongterm,
E.expHidden, 
E.expIncompleteCV, 
E.expToCompleteCVEmailSent, 
E.expToConfirmCvEmailSent, 
E.expApproved, 
E.expRemoved, 
E.expDeleted, 
NULL expInHouseAgreed, 
NULL expInHouseSent
FROM dbo.tbl_Experts E 
INNER JOIN dbo.tbl_Persons P ON E.id_Expert = P.id_Expert 
LEFT OUTER JOIN dbo.tbl_PersonTitles PT ON P.id_psnTitle = PT.id_psnTitle


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE VIEW dbo.uvw_Region
AS
SELECT 
R.id_GeoZone id_Region,
R.Geo_ZoneEng regNameEng,
R.id_Continent
FROM tbl_GeoZone R
WHERE db_NotVisible = 0 


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE FUNCTION udf_TitleCase (@InputString varchar(4000) )
RETURNS VARCHAR(4000)
AS
BEGIN
DECLARE @Index          INT
DECLARE @Char           CHAR(1)
DECLARE @OutputString   VARCHAR(255)
SET @OutputString = LOWER(@InputString)
SET @Index = 2
SET @OutputString =
   STUFF(@OutputString, 1, 1, UPPER(SUBSTRING(@InputString,1,1)))
WHILE @Index <= LEN(@InputString)
BEGIN
SET @Char = SUBSTRING(@InputString, @Index, 1)
IF @Char IN (' ', ';', ':', '!', '?', ',', '.', '_', '-', '/', '&', '(', '''', CHAR(39))
IF @Index + 1 <= LEN(@InputString)
BEGIN
--IF @Char != '''' OR
--UPPER(SUBSTRING(@InputString, @Index + 1, 1)) != 'S'
SET @OutputString =
   STUFF(@OutputString, @Index + 1, 1,UPPER(SUBSTRING(@InputString, @Index + 1, 1)))
END
SET @Index = @Index + 1
END
RETURN ISNULL(@OutputString,'')
END


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE FUNCTION udf_UserRegisteredExpert (
@iExpertID int
)
RETURNS int
AS
BEGIN
DECLARE @iResult int

DECLARE @sExpertID varchar(8)
SET @sExpertID=CAST(@iExpertID AS varchar(8))

DECLARE @tUserDate TABLE (
	id_User int,
	slgDate smalldatetime
	)

INSERT INTO @tUserDate (
id_User,
slgDate
) 
SELECT S.id_User, SE.slgDate
FROM log_Session S
INNER JOIN log_SessionEvent SE ON S.id_UserSession=SE.id_UserSession
WHERE SE.slgUrl = '/fei/external/register/register.asp?id=' + @sExpertID
OR SE.slgUrl LIKE '/fei/external/register/register.asp?id=' + @sExpertID + '&%'

SELECT TOP 1 @iResult=S1.id_User
FROM @tUserDate S1
INNER JOIN 
	(
	SELECT MIN(slgDate) slgDate
	FROM @tUserDate
) S2 ON S1.slgDate=S2.slgDate

RETURN @iResult
END

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE FUNCTION udf_ExpertEmailAll (
@iExpertID int
) 
RETURNS varchar(500)
AS
BEGIN
DECLARE @sResult varchar(500)
SET @sResult=''

SELECT @sResult=@sResult + Email + '; '
FROM (	
	SELECT Email
	FROM tbl_Experts
	WHERE id_Expert=@iExpertID
	AND Email IS NOT NULL
	UNION 
	SELECT adrEmail Email
	FROM tbl_Exp_Address
	WHERE id_Expert=@iExpertID
	AND adrEmail IS NOT NULL
	) T1

RETURN @sResult
END


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE FUNCTION udf_ExpertWebsite (
@iExpertID int
) 
RETURNS varchar(500)
AS
BEGIN
DECLARE @sResult varchar(500)
SET @sResult=''

SELECT @sResult=@sResult + adrWeb + '; '
FROM tbl_Exp_Address
WHERE id_Expert=@iExpertID
AND adrWeb IS NOT NULL


RETURN @sResult
END



GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE FUNCTION udf_ExpertExperienceLastDate (
@iExpertID int
)
RETURNS smalldatetime
AS
BEGIN
DECLARE @dResult smalldatetime

SELECT @dResult=MAX(wkeEndDate) 
FROM lnkExp_Wke 
WHERE id_Expert=@iExpertID

RETURN @dResult
END


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO



SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO


CREATE PROCEDURE usp_SrvRemoveNonTextSymbols
@sInputText nvarchar(2000), @sOutputText nvarchar(2000) OUTPUT
AS
SET @sOutputText=@sInputText
DECLARE @iPosition int, @iUnicode int
SET @iPosition = 1
WHILE @iPosition <= DATALENGTH(@sInputText)
BEGIN
	SET @iUnicode=UNICODE(SUBSTRING(@sInputText, @iPosition, 1))
	IF (@iUnicode<48 AND @iUnicode<>32 AND @iUnicode<>34) Or (@iUnicode>57 AND @iUnicode<65) Or (@iUnicode>90 And @iUnicode<97) Or (@iUnicode>123 And @iUnicode<191)
		SET @sOutputText=REPLACE(@sOutputText,SUBSTRING(@sInputText, @iPosition, 1), '')
	SET @iPosition = @iPosition + 1
END
SET @sOutputText=REPLACE(@sOutputText, '"" AND ', '')
SET @sOutputText=REPLACE(@sOutputText, '"" OR ', '')
SET @sOutputText=REPLACE(@sOutputText, '"" NEAR ', '')
SET @sOutputText=REPLACE(@sOutputText, '"" NOT ', '')
SET @sOutputText=REPLACE(@sOutputText, ' AND ""', '')
SET @sOutputText=REPLACE(@sOutputText, ' OR ""', '')
SET @sOutputText=REPLACE(@sOutputText, ' NEAR ""', '')
SET @sOutputText=REPLACE(@sOutputText, ' NOT ""', '')



GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO


SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE usp_AdmCreateList1FromRecordset @@sListType varchar(10), @@sFieldsList nvarchar(4000) OUTPUT
AS
DECLARE @iMainFieldID int, @sMainFieldText nvarchar(255), @iFieldID int, @sFieldText nvarchar(255), @sFlag nvarchar(255)
DECLARE cursorFieldsTable CURSOR LOCAL FAST_FORWARD FOR 
SELECT * FROM #tblFieldsTable
OPEN cursorFieldsTable
FETCH NEXT FROM cursorFieldsTable INTO @iFieldID, @sFieldText 
SET @@sFieldsList=''
SET @sFlag=''
WHILE @@FETCH_STATUS = 0
BEGIN
	
	IF @@sListType='row'
		SET @@sFieldsList = @@sFieldsList + REPLACE(@sFieldText,' ','&nbsp;') + ',   '
	ELSE
		SET @@sFieldsList = @@sFieldsList + '-&nbsp;' + REPLACE(@sFieldText,' ','&nbsp;') + '<br>'
	FETCH NEXT FROM cursorFieldsTable INTO @iFieldID, @sFieldText 
END
SET @@sFieldsList=@@sFieldsList+'.'
IF LEN(@@sFieldsList+'.')>5
	SET @@sFieldsList=SUBSTRING(@@sFieldsList, 1, LEN(@@sFieldsList)-5)
IF @@sFieldsList='.'
	SET @@sFieldsList=''
CLOSE cursorFieldsTable
DEALLOCATE cursorFieldsTable

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE usp_AdmCreateList2FromRecordset @@sListType varchar(10), @@sFieldsList nvarchar(4000) OUTPUT
AS
DECLARE @iMainFieldID int, @sMainFieldText nvarchar(255), @iFieldID int, @sFieldText nvarchar(255), @sFlag nvarchar(255)
DECLARE cursorFieldsTable CURSOR LOCAL FAST_FORWARD FOR 
SELECT * FROM #tblFieldsTable
OPEN cursorFieldsTable
FETCH NEXT FROM cursorFieldsTable INTO @iMainFieldID, @sMainFieldText, @iFieldID, @sFieldText 
SET @@sFieldsList=''
SET @sFlag=''
WHILE @@FETCH_STATUS = 0
BEGIN
	IF @sFlag<>@sMainFieldText
		BEGIN
		IF @@sListType='row'
 			SET @@sFieldsList = @@sFieldsList + '</p><p><b>' + @sMainFieldText + ':</b>'
		ELSE
 			SET @@sFieldsList = @@sFieldsList + '</p><p><b>' + @sMainFieldText + ':</b><br>'
		SET @sFlag=@sMainFieldText
		END
	
	IF @@sListType='row'
		SET @@sFieldsList = @@sFieldsList + '&nbsp;' + REPLACE(@sFieldText,' ','&nbsp;') + ',   '
	ELSE
		SET @@sFieldsList = @@sFieldsList + '&nbsp;-&nbsp;' + REPLACE(@sFieldText,' ','&nbsp;') + '<br>'
	FETCH NEXT FROM cursorFieldsTable INTO @iMainFieldID, @sMainFieldText, @iFieldID, @sFieldText
END
IF LEN(@@sFieldsList)>8
SET @@sFieldsList=SUBSTRING(@@sFieldsList, 8, LEN(@@sFieldsList)-8)
--IF LEN(@@sFieldsList)>4
--SET @@sFieldsList=SUBSTRING(@@sFieldsList, 1, LEN(@@sFieldsList)-1)
SET @@sFieldsList=REPLACE(@@sFieldsList, ',   </p>', '</p>')
CLOSE cursorFieldsTable
DEALLOCATE cursorFieldsTable

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE usp_AdmCreateList3FromRecordset @@sListType varchar(10), @@sFieldsList nvarchar(4000) OUTPUT
AS
DECLARE @iMainFieldID int, @sMainFieldText nvarchar(255), @iFieldID int, @sFieldText nvarchar(255), @sFlag nvarchar(255)
DECLARE cursorFieldsTable CURSOR LOCAL FAST_FORWARD FOR 
SELECT * FROM #tblFieldsTable
OPEN cursorFieldsTable
FETCH NEXT FROM cursorFieldsTable INTO @iMainFieldID, @sMainFieldText, @iFieldID, @sFieldText 
SET @@sFieldsList='-'
SET @sFlag=''
WHILE @@FETCH_STATUS = 0
BEGIN
	IF @@sListType='row'
		SET @@sFieldsList = @@sFieldsList + '&nbsp;' + REPLACE(@sFieldText,' ','&nbsp;') + ',   '
	ELSE
		SET @@sFieldsList = @@sFieldsList + '&nbsp;-&nbsp;' + REPLACE(@sFieldText,' ','&nbsp;') + '<br>'
	FETCH NEXT FROM cursorFieldsTable INTO @iMainFieldID, @sMainFieldText, @iFieldID, @sFieldText
END
IF LEN(@@sFieldsList)>8
SET @@sFieldsList=SUBSTRING(@@sFieldsList, 8, LEN(@@sFieldsList)-8)
--IF LEN(@@sFieldsList)>4
--SET @@sFieldsList=SUBSTRING(@@sFieldsList, 1, LEN(@@sFieldsList)-1)
SET @@sFieldsList=REPLACE(@@sFieldsList, ',   </p>', '</p>')
CLOSE cursorFieldsTable
DEALLOCATE cursorFieldsTable

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE usp_AdmCreateRecordset
@@sRecordsetValues varchar(8000)
AS
IF @@sRecordsetValues>'' 
	BEGIN
	SET @@sRecordsetValues=REPLACE(@@sRecordsetValues,',',' UNION ALL SELECT ')
	EXEC ('SELECT  '+@@sRecordsetValues )
	END
ELSE
	SELECT null

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

/*
CVIP usp_AdmExpAllListExtraModifiedSelect
*/
CREATE PROCEDURE usp_AdmExpAllListExtraModifiedSelect
@sStatusList varchar(250),
@bShowHiddenExperts bit, 
@bShowRemovedExperts bit, 
@sCondition varchar(100)=NULL, 
@sSearchString varchar(255)=NULL, 
@sOrderBy varchar(100)='A',
@sLastExperienceFromDate varchar(16), 
@sLastExperienceToDate varchar(16),
@sCvModifiedFromDate varchar(16), 
@sCvModifiedToDate varchar(16)
AS
SET NOCOUNT OFF

-- 1. Convert dates
SET @sSearchString=ISNULL(@sSearchString,'')
-- 
DECLARE  @dLastExperienceFromDate smalldatetime, @dLastExperienceToDate smalldatetime

IF ISDATE(@sLastExperienceFromDate)=1
	SET @dLastExperienceFromDate=CONVERT(smalldatetime, @sLastExperienceFromDate)
ELSE
	SET @dLastExperienceFromDate=NULL

IF ISDATE(@sLastExperienceToDate)=1
	SET @dLastExperienceToDate=CONVERT(smalldatetime, @sLastExperienceToDate)
ELSE
	SET @dLastExperienceToDate=NULL
--
DECLARE  @dCvModifiedFromDate smalldatetime, @dCvModifiedToDate smalldatetime

IF ISDATE(@sCvModifiedFromDate)=1
	SET @dCvModifiedFromDate=CONVERT(smalldatetime, @sCvModifiedFromDate)
ELSE
	SET @dCvModifiedFromDate=NULL

IF ISDATE(@sCvModifiedToDate)=1
	SET @dCvModifiedToDate=CONVERT(smalldatetime, @sCvModifiedToDate)
ELSE
	SET @dCvModifiedToDate=NULL


-- 2.1 If the list of all experts is requested then reset flags for hidden and removed
IF @sCondition='all'
	SELECT @bShowHiddenExperts=1, @bShowRemovedExperts=1

IF @sCondition='hidden'
	SELECT @bShowHiddenExperts=1, @bShowRemovedExperts=0

IF @sCondition='deleted'
	SELECT @bShowHiddenExperts=1, @bShowRemovedExperts=1

-- 3. If the list of experts with CVs not updated over a certain period is requested then identify the period
DECLARE @iNumberMonth int
IF ISNUMERIC(REPLACE(@sCondition, 'updatedover', ''))>0
	SET @iNumberMonth=CONVERT(int, REPLACE(@sCondition, 'updatedover', ''))
ELSE
	SET @iNumberMonth=0

-- 4. Select the list of experts
SELECT E2.*,
E.Lng,
E.KgCvFile
FROM udf_AdmExpAllListExtraModifiedSelect(@sStatusList, @bShowHiddenExperts, @bShowRemovedExperts, @dLastExperienceFromDate, @dLastExperienceToDate, @dCvModifiedFromDate, @dCvModifiedToDate) E2
INNER JOIN tbl_Experts E ON E2.id_Expert=E.id_Expert
WHERE 1 = CASE 
	WHEN @sCondition='noemail' 
		AND NOT ((E2.Email IS NULL) OR (E2.Email NOT LIKE '%@%')) THEN 0
	WHEN @sCondition='noaddress' 
		AND NOT ((E2.Email IS NULL) OR (E2.Email not like '%@%')) AND ((E2.Phone IS NULL) OR (LEN(E2.Phone)<5)) AND E2.id_Expert not in (SELECT id_Expert FROM tbl_Exp_Address WHERE adrEmail like '%@%') THEN 0
	WHEN CHARINDEX('updatedover', @sCondition)>0 AND @iNumberMonth>0 
		AND NOT (DATEDIFF(mm, dbo.udf_ExpertExperienceLastDate(E2.id_Expert), GETDATE())>@iNumberMonth) THEN 0
	WHEN @sCondition='updatedrange3to6' 
		AND NOT (DATEDIFF(mm, dbo.udf_ExpertExperienceLastDate(E2.id_Expert), GETDATE())>3 AND DATEDIFF(mm, dbo.udf_ExpertExperienceLastDate(E2.id_Expert), GETDATE())<=6) THEN 0
	WHEN @sCondition='updatedrange6to12' 
		AND NOT (DATEDIFF(mm, dbo.udf_ExpertExperienceLastDate(E2.id_Expert), GETDATE())>6 AND DATEDIFF(mm, dbo.udf_ExpertExperienceLastDate(E2.id_Expert), GETDATE())<=12) THEN 0
	WHEN @sCondition='updatedrange12to24' 
		AND NOT (DATEDIFF(mm, dbo.udf_ExpertExperienceLastDate(E2.id_Expert), GETDATE())>12 AND DATEDIFF(mm, dbo.udf_ExpertExperienceLastDate(E2.id_Expert), GETDATE())<=24) THEN 0
	WHEN CHARINDEX('registeredweek', @sCondition)>0 AND NOT (DATEPART(wk, E2.expCreateDate)=DATEPART(wk, GETDATE()) AND YEAR(E2.expCreateDate)=YEAR(GETDATE())) THEN 0
	WHEN CHARINDEX('registeredmonth', @sCondition)>0 AND NOT (DATEPART(mm, E2.expCreateDate)=DATEPART(mm, GETDATE()) AND DATEDIFF(mm, E2.expCreateDate, GETDATE())=0) THEN 0
	WHEN @sCondition='deleted' 
		AND E2.expRemoved=0 AND E2.expDeleted=0 THEN 0
	ELSE 1 
	END
AND
(
E2.id_Expert = CASE WHEN ISNUMERIC(@sSearchString)>0 THEN CONVERT(varchar, @sSearchString) ELSE 0 END
OR
(E2.EmailAll LIKE '%' + @sSearchString + '%') OR (E2.psnLastName LIKE '%' + @sSearchString + '%') OR (E2.psnFirstName LIKE @sSearchString + '%') OR (E2.psnMiddleName LIKE @sSearchString + '%') OR (E2.expComments LIKE @sSearchString + '%')
)
ORDER BY 
CASE WHEN @sOrderBy='I' THEN E2.id_Expert ELSE NULL END DESC,
CASE WHEN @sOrderBy='A' THEN E2.psnLastName ELSE NULL END,
CASE WHEN @sOrderBy='R' THEN E2.expCreateDate ELSE NULL END DESC,
CASE WHEN @sOrderBy='U' THEN E2.expLastUpdate ELSE NULL END DESC,
CASE WHEN @sOrderBy='E' THEN E2.wkeEndDate ELSE NULL END,
CASE WHEN @sOrderBy='B' THEN MONTH(E2.psnBirthDate)*100 + DAY(E2.psnBirthDate) ELSE NULL END

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

/*
CVIP usp_AdmExpAllListExtraModifiedSelect
*/
CREATE PROCEDURE usp_AdmExpAllListExtraModifiedWithPasswordSelect
@sStatusList varchar(250),
@bShowHiddenExperts bit, 
@bShowRemovedExperts bit, 
@sCondition varchar(100)=NULL, 
@sSearchString varchar(255)=NULL, 
@sOrderBy varchar(100)='A',
@sLastExperienceFromDate varchar(16), 
@sLastExperienceToDate varchar(16),
@sCvModifiedFromDate varchar(16), 
@sCvModifiedToDate varchar(16)
AS
SET NOCOUNT OFF

-- 1. Convert dates
SET @sSearchString=ISNULL(@sSearchString,'')
-- 
DECLARE  @dLastExperienceFromDate smalldatetime, @dLastExperienceToDate smalldatetime

IF ISDATE(@sLastExperienceFromDate)=1
	SET @dLastExperienceFromDate=CONVERT(smalldatetime, @sLastExperienceFromDate)
ELSE
	SET @dLastExperienceFromDate=NULL

IF ISDATE(@sLastExperienceToDate)=1
	SET @dLastExperienceToDate=CONVERT(smalldatetime, @sLastExperienceToDate)
ELSE
	SET @dLastExperienceToDate=NULL
--
DECLARE  @dCvModifiedFromDate smalldatetime, @dCvModifiedToDate smalldatetime

IF ISDATE(@sCvModifiedFromDate)=1
	SET @dCvModifiedFromDate=CONVERT(smalldatetime, @sCvModifiedFromDate)
ELSE
	SET @dCvModifiedFromDate=NULL

IF ISDATE(@sCvModifiedToDate)=1
	SET @dCvModifiedToDate=CONVERT(smalldatetime, @sCvModifiedToDate)
ELSE
	SET @dCvModifiedToDate=NULL


-- 2.1 If the list of all experts is requested then reset flags for hidden and removed
IF @sCondition='all'
	SELECT @bShowHiddenExperts=1, @bShowRemovedExperts=1

IF @sCondition='hidden'
	SELECT @bShowHiddenExperts=1, @bShowRemovedExperts=0

IF @sCondition='deleted'
	SELECT @bShowHiddenExperts=1, @bShowRemovedExperts=1

-- 3. If the list of experts with CVs not updated over a certain period is requested then identify the period
DECLARE @iNumberMonth int
IF ISNUMERIC(REPLACE(@sCondition, 'updatedover', ''))>0
	SET @iNumberMonth=CONVERT(int, REPLACE(@sCondition, 'updatedover', ''))
ELSE
	SET @iNumberMonth=0

-- 4. Select the list of experts
SELECT ES.*, 
U.id_User,
U.[UserName],
U.[Password]
FROM udf_AdmExpAllListExtraModifiedSelect(@sStatusList, @bShowHiddenExperts, @bShowRemovedExperts, @dLastExperienceFromDate, @dLastExperienceToDate, @dCvModifiedFromDate, @dCvModifiedToDate) ES
INNER JOIN tbl_Experts E ON ES.id_Expert=E.id_Expert
INNER JOIN tbl_Users U ON E.id_User=U.id_User
WHERE 1 = CASE 
	WHEN @sCondition='noemail' 
		AND NOT ((ES.Email IS NULL) OR (ES.Email NOT LIKE '%@%')) THEN 0
	WHEN @sCondition='noaddress' 
		AND NOT ((ES.Email IS NULL) OR (ES.Email not like '%@%')) AND ((ES.Phone IS NULL) OR (LEN(ES.Phone)<5)) AND ES.id_Expert not in (SELECT id_Expert FROM tbl_Exp_Address WHERE adrEmail like '%@%') THEN 0
	WHEN CHARINDEX('updatedover', @sCondition)>0 AND @iNumberMonth>0 
		AND NOT (DATEDIFF(mm, dbo.udf_ExpertExperienceLastDate(ES.id_Expert), GETDATE())>@iNumberMonth) THEN 0
	WHEN @sCondition='updatedrange3to6' 
		AND NOT (DATEDIFF(mm, dbo.udf_ExpertExperienceLastDate(ES.id_Expert), GETDATE())>3 AND DATEDIFF(mm, dbo.udf_ExpertExperienceLastDate(ES.id_Expert), GETDATE())<=6) THEN 0
	WHEN @sCondition='updatedrange6to12' 
		AND NOT (DATEDIFF(mm, dbo.udf_ExpertExperienceLastDate(ES.id_Expert), GETDATE())>6 AND DATEDIFF(mm, dbo.udf_ExpertExperienceLastDate(ES.id_Expert), GETDATE())<=12) THEN 0
	WHEN @sCondition='updatedrange12to24' 
		AND NOT (DATEDIFF(mm, dbo.udf_ExpertExperienceLastDate(ES.id_Expert), GETDATE())>12 AND DATEDIFF(mm, dbo.udf_ExpertExperienceLastDate(ES.id_Expert), GETDATE())<=24) THEN 0
	WHEN CHARINDEX('registeredweek', @sCondition)>0 AND NOT (DATEPART(wk, ES.expCreateDate)=DATEPART(wk, GETDATE()) AND YEAR(ES.expCreateDate)=YEAR(GETDATE())) THEN 0
	WHEN CHARINDEX('registeredmonth', @sCondition)>0 AND NOT (DATEPART(mm, ES.expCreateDate)=DATEPART(mm, GETDATE()) AND DATEDIFF(mm, ES.expCreateDate, GETDATE())=0) THEN 0
	WHEN @sCondition='deleted' 
		AND ES.expRemoved=0 AND ES.expDeleted=0 THEN 0
	ELSE 1 
	END
AND
(
ES.id_Expert = CASE WHEN ISNUMERIC(@sSearchString)>0 THEN CONVERT(varchar, @sSearchString) ELSE 0 END
OR
(ES.EmailAll LIKE '%' + @sSearchString + '%') OR (ES.psnLastName LIKE '%' + @sSearchString + '%') OR (ES.psnFirstName LIKE @sSearchString + '%') OR (ES.psnMiddleName LIKE @sSearchString + '%') OR (ES.expComments LIKE @sSearchString + '%')
)
ORDER BY 
CASE WHEN @sOrderBy='I' THEN ES.id_Expert ELSE NULL END DESC,
CASE WHEN @sOrderBy='A' THEN ES.psnLastName ELSE NULL END,
CASE WHEN @sOrderBy='R' THEN ES.expCreateDate ELSE NULL END DESC,
CASE WHEN @sOrderBy='U' THEN ES.expLastUpdate ELSE NULL END DESC,
CASE WHEN @sOrderBy='E' THEN ES.wkeEndDate ELSE NULL END,
CASE WHEN @sOrderBy='B' THEN MONTH(ES.psnBirthDate)*100 + DAY(ES.psnBirthDate) ELSE NULL END

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE usp_AdmExpAllListExtraSelect
@sStatusList varchar(250), 
@bShowHiddenExperts bit, 
@bShowRemovedExperts bit, 
@sCondition varchar(100)=NULL, 
@sSearchString varchar(255)=NULL, 
@sOrderBy varchar(100)='A',
@sLastExperienceFromDate varchar(16), 
@sLastExperienceToDate varchar(16)
AS
SET NOCOUNT OFF

-- 1. Convert dates
SET @sSearchString=ISNULL(@sSearchString,'')
DECLARE  @dLastExperienceFromDate smalldatetime, @dLastExperienceToDate smalldatetime

IF ISDATE(@sLastExperienceFromDate)=1
	SET @dLastExperienceFromDate=CONVERT(smalldatetime, @sLastExperienceFromDate)
ELSE
	SET @dLastExperienceFromDate=NULL

IF ISDATE(@sLastExperienceToDate)=1
	SET @dLastExperienceToDate=CONVERT(smalldatetime, @sLastExperienceToDate)
ELSE
	SET @dLastExperienceToDate=NULL

-- 2.1 If the list of all experts is requested then reset flags for hidden and removed
IF @sCondition='all'
	SELECT @bShowHiddenExperts=1, @bShowRemovedExperts=1

IF @sCondition='hidden'
	SELECT @bShowHiddenExperts=1, @bShowRemovedExperts=0


-- 3. If the list of experts with CVs not updated over a certain period is requested then identify the period
DECLARE @iNumberMonth int
IF ISNUMERIC(REPLACE(@sCondition, 'updatedover', ''))>0
	SET @iNumberMonth=CONVERT(int, REPLACE(@sCondition, 'updatedover', ''))
ELSE
	SET @iNumberMonth=0

-- 4. Select the list of experts
SELECT * FROM udf_AdmExpAllListExtraSelect(@sStatusList, @bShowHiddenExperts, @bShowRemovedExperts, @dLastExperienceFromDate, @dLastExperienceToDate)
WHERE 1 = CASE 
	WHEN @sCondition='noemail' 
		AND NOT ((Email IS NULL) OR (Email NOT LIKE '%@%')) THEN 0
	WHEN @sCondition='bademail' 
		AND NOT (Email in (SELECT Email FROM eml_bad_emails_all)) THEN 0
	WHEN @sCondition='noaddress' 
		AND NOT ((Email IS NULL) OR (Email not like '%@%')) AND ((Phone IS NULL) OR (LEN(Phone)<5)) AND id_Expert not in (SELECT id_Expert FROM tbl_Exp_Address WHERE adrEmail like '%@%') THEN 0
	WHEN CHARINDEX('updatedover', @sCondition)>0 AND @iNumberMonth>0 
		AND NOT (DATEDIFF(mm, dbo.udf_ExpertExperienceLastDate(id_Expert), GETDATE())>@iNumberMonth) THEN 0
	ELSE 1 
	END
AND
(
id_Expert = CASE WHEN ISNUMERIC(@sSearchString)>0 THEN CONVERT(varchar, @sSearchString) ELSE 0 END
OR
(EmailAll LIKE '%' + @sSearchString + '%') OR (psnLastName LIKE '%' + @sSearchString + '%') OR (psnFirstName LIKE @sSearchString + '%')
)
ORDER BY 
CASE WHEN @sOrderBy='I' THEN id_Expert ELSE NULL END DESC,
CASE WHEN @sOrderBy='A' THEN psnLastName ELSE NULL END,
CASE WHEN @sOrderBy='R' THEN expCreateDate ELSE NULL END DESC,
CASE WHEN @sOrderBy='U' THEN expLastUpdate ELSE NULL END DESC,
CASE WHEN @sOrderBy='E' THEN wkeEndDate ELSE NULL END,
CASE WHEN @sOrderBy='B' THEN MONTH(psnBirthDate)*100 + DAY(psnBirthDate) ELSE NULL END



GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

/*
CVIP usp_AdmExpAllListSelect
*/
CREATE PROCEDURE usp_AdmExpAllListSelect
@sStatusList varchar(250),
@bShowHiddenExperts bit, 
@bShowRemovedExperts bit, 
@sCondition varchar(100)=NULL, 
@sSearchString varchar(255)=NULL, 
@sOrderBy varchar(100)='A'
AS
SET NOCOUNT ON

SET @sSearchString=ISNULL(@sSearchString,'')

-- 2.1 If the list of all experts is requested then reset flags for hidden and removed
IF @sCondition='all'
	SELECT @bShowHiddenExperts=1, @bShowRemovedExperts=1

IF @sCondition='hidden'
	SELECT @bShowHiddenExperts=1, @bShowRemovedExperts=0

IF @sCondition='deleted'
	SELECT @bShowHiddenExperts=1, @bShowRemovedExperts=1

-- 3. If the list of experts with CVs not updated over a certain period is requested then identify the period
DECLARE @iNumberMonth int
IF ISNUMERIC(REPLACE(@sCondition, 'updatedover', ''))>0
	SET @iNumberMonth=CONVERT(int, REPLACE(@sCondition, 'updatedover', ''))
ELSE
	SET @iNumberMonth=0

SET NOCOUNT OFF

-- 4. Select the list of experts
SELECT * FROM udf_AdmExpAllListExtraSelect(@sStatusList, @bShowHiddenExperts, @bShowRemovedExperts, NULL, NULL)
WHERE 1 = CASE 
	WHEN @sCondition='noemail' 
		AND NOT ((Email IS NULL) OR (Email NOT LIKE '%@%')) THEN 0
	WHEN @sCondition='noaddress' 
		AND NOT ((Email IS NULL) OR (Email not like '%@%')) AND ((Phone IS NULL) OR (LEN(Phone)<5)) AND id_Expert not in (SELECT id_Expert FROM tbl_Exp_Address WHERE adrEmail like '%@%') THEN 0
	WHEN CHARINDEX('updatedover', @sCondition)>0 AND @iNumberMonth>0 
		AND NOT (DATEDIFF(mm, dbo.udf_ExpertExperienceLastDate(id_Expert), GETDATE())>@iNumberMonth) THEN 0
	WHEN CHARINDEX('registeredweek', @sCondition)>0 AND NOT DATEPART(wk, expCreateDate)=DATEPART(wk, GETDATE()) THEN 0
	WHEN CHARINDEX('registeredmonth', @sCondition)>0 AND NOT (DATEPART(mm, expCreateDate)=DATEPART(mm, GETDATE()) AND DATEDIFF(mm, expCreateDate, GETDATE())=0) THEN 0
	WHEN @sCondition='deleted' 
		AND expRemoved=0 AND expDeleted=0 THEN 0
	ELSE 1 
	END
AND
(
id_Expert = CASE WHEN ISNUMERIC(@sSearchString)>0 THEN CONVERT(varchar, @sSearchString) ELSE 0 END
OR
(Email LIKE @sSearchString + '%') OR (psnLastName LIKE '%' + @sSearchString + '%') OR (psnFirstName LIKE @sSearchString + '%')
)
ORDER BY 
CASE WHEN @sOrderBy='I' THEN id_Expert ELSE NULL END DESC,
CASE WHEN @sOrderBy='A' THEN psnLastName ELSE NULL END,
CASE WHEN @sOrderBy='R' THEN expCreateDate ELSE NULL END DESC,
CASE WHEN @sOrderBy='U' THEN expLastUpdate ELSE NULL END DESC,
CASE WHEN @sOrderBy='E' THEN wkeEndDate ELSE NULL END,
CASE WHEN @sOrderBy='B' THEN MONTH(psnBirthDate)*100 + DAY(psnBirthDate) ELSE NULL END


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE usp_AdmExpDuplicateHide
@iExpertID int, @iOriginalExpertID int, @sComments text, @iResults int output
AS
SET NOCOUNT ON
SET @iResults=0

IF EXISTS(SELECT id_Expert 
	FROM tbl_Experts 
	WHERE id_Expert=@iOriginalExpertID 
	AND @iExpertID<>@iOriginalExpertID 
	AND expDeleted=0 
	AND expRemoved=0
	)
	BEGIN
	UPDATE tbl_Experts
	SET id_ExpertOriginal=@iOriginalExpertID, expDeleted=1, expDeletedDate=GETDATE(), expDeletedComments=@sComments
	WHERE id_Expert=@iExpertID
	SELECT @iResults=@@ROWCOUNT
	END


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE usp_AdmExpExistsCheck
@sFirstName nvarchar(255), 
@sFamilyName nvarchar(255), 
@dBirthDate smalldatetime=Null, 
@sEmail nvarchar(255)=Null
AS
SET NOCOUNT OFF
DECLARE @sSearchFor nvarchar(500)
SET @sSearchFor=@sFamilyName + ' OR ' + @sFirstName
SET @sFirstName=LTRIM(RTRIM(@sFirstName))
SET @sFamilyName=LTRIM(RTRIM(@sFamilyName))

SELECT E.ptlName, E.psnFirstName, E.psnMiddleName, E.psnLastName, E.psnBirthDate, E.id_Expert, E.Lng, E.Email, E.Phone, E.KgEncoded, E.KgCVFile, 
E.expHidden, E.expIncompleteCV, E.expToCompleteCVEmailSent, E.expToConfirmCvEmailSent, E.expApproved, E.expRemoved, E.expDeleted,
E.rnkMatch,
dbo.udf_ExpertExperienceLastDate(E.id_Expert) wkeEndDate
FROM
(
	SELECT ptlName, psnFirstName, psnMiddleName, psnLastName, psnBirthDate, id_Expert, Lng, Email, Phone, KgEncoded, KgCVFile, 
	expHidden, expIncompleteCV, expToCompleteCVEmailSent, expToConfirmCvEmailSent, expApproved, expRemoved, expDeleted,
	MIN(rnkMatch) AS rnkMatch
	FROM
	(
	-- 1 - Full names and dates of birth are the same
	SELECT ptlName, psnFirstName, psnMiddleName, psnLastName, psnBirthDate, id_Expert, Lng, Email, Phone, KgEncoded, KgCVFile, 
		expHidden, expIncompleteCV, expToCompleteCVEmailSent, expToConfirmCvEmailSent, expApproved, expRemoved, expDeleted,
		1 AS rnkMatch FROM uvw_Experts 
		WHERE (psnLastName like '%' + @sFamilyName + '%' AND psnFirstName like '%' + @sFirstName + '%')
		AND @dBirthDate=psnBirthDate
	UNION
	-- 2 - Full names are the same
	SELECT ptlName, psnFirstName, psnMiddleName, psnLastName, psnBirthDate, id_Expert, Lng, Email, Phone, KgEncoded, KgCVFile, 
		expHidden, expIncompleteCV, expToCompleteCVEmailSent, expToConfirmCvEmailSent, expApproved, expRemoved, expDeleted,
		2 AS rnkMatch FROM uvw_Experts 
		WHERE (psnLastName like '%' + @sFamilyName + '%' AND psnFirstName like '%' + @sFirstName + '%')
	UNION
	-- 3 - Emails are the same
	SELECT ptlName, psnFirstName, psnMiddleName, psnLastName, psnBirthDate, id_Expert, Lng, Email, Phone, KgEncoded, KgCVFile, 
		expHidden, expIncompleteCV, expToCompleteCVEmailSent, expToConfirmCvEmailSent, expApproved, expRemoved, expDeleted,
		3 AS rnkMatch FROM uvw_Experts 
		WHERE Email like CASE WHEN LEN(Email)>0 THEN '%' + @sEmail + '%' ELSE '/*/*/' END 
	UNION
	-- 4 - Soundexes are the same
	SELECT ptlName, psnFirstName, psnMiddleName, psnLastName, psnBirthDate, id_Expert, Lng, Email, Phone, KgEncoded, KgCVFile, 
		expHidden, expIncompleteCV, expToCompleteCVEmailSent, expToConfirmCvEmailSent, expApproved, expRemoved, expDeleted,
		4 AS rnkMatch FROM uvw_Experts 
		WHERE (psnLastName like '%' + @sFamilyName + '%' AND psnFirstName like '%' + @sFirstName + '%')
		OR (DIFFERENCE(psnLastName, @sFamilyName)=4 AND CASE WHEN LEN(@sFirstName)>0 THEN DIFFERENCE(psnFirstName, @sFirstName) ELSE 4 END>3)
	UNION
	-- 5 - Firstname<->Surname
	SELECT ptlName, psnFirstName, psnMiddleName, psnLastName, psnBirthDate, id_Expert, Lng, Email, Phone, KgEncoded, KgCVFile, 
		expHidden, expIncompleteCV, expToCompleteCVEmailSent, expToConfirmCvEmailSent, expApproved, expRemoved, expDeleted,
		5 AS rnkMatch FROM uvw_Experts 
		WHERE 
		(psnLastName like '%' + @sFirstName + '%' AND psnFirstName like '%' + @sFamilyName + '%')
	) AS T1
	GROUP BY ptlName, psnFirstName, psnMiddleName, psnLastName, psnBirthDate, id_Expert, Lng, Email, Phone, KgEncoded, KgCVFile, 
	expHidden, expIncompleteCV, expToCompleteCVEmailSent, expToConfirmCvEmailSent, expApproved, expRemoved, expDeleted
) E
-- SHOW EXPERTS IN 1 LANGUAGE ONLY
LEFT OUTER JOIN tbl_ExpertsLanguage EL ON E.id_Expert=EL.id_Expert2
WHERE EL.id_Expert2 IS NULL
ORDER BY rnkMatch, psnLastName


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE usp_AdmExpRemove
@iExpertID int, @iReason int, @sComments text, @iResults int output
AS
SET NOCOUNT ON
SET @iResults=0

IF @iReason=1 
	BEGIN
	UPDATE tbl_Experts
	SET expDeleted=1, expDeletedDate=GETDATE(), expDeletedComments=@sComments
	WHERE id_Expert=@iExpertID
	SELECT @iResults=@@ROWCOUNT
	END
ELSE
IF @iReason=2
	BEGIN
	UPDATE tbl_Experts
	SET expRemoved=1, expRemovedDate=GETDATE(), expRemovedComments=@sComments
	WHERE id_Expert=@iExpertID
	SELECT @iResults=@@ROWCOUNT
	END


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE usp_AdmExpRestore
@iExpertID int, @iResults int output
AS
SET NOCOUNT ON
SET @iResults=0

	BEGIN
	UPDATE tbl_Experts
	SET expDeleted=0, expRemoved=0
	WHERE id_Expert=@iExpertID
	SELECT @iResults=@@ROWCOUNT
	END




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE usp_CurrencyListSelect
@iCurrencyID int = NULL,
@sCurrencyName nvarchar(255) = NULL,
@sOrderBy varchar(80) = NULL
AS
SELECT C.*
FROM tbl_Currency C
WHERE 
(@iCurrencyID IS NULL OR C.id_Currency = @iCurrencyID)
AND
(@sCurrencyName IS NULL OR C.curAbbreviation LIKE @sCurrencyName + '%')
ORDER BY id_Currency


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE usp_DatContinentSelect
AS
SELECT id_Continent, conDescriptionEng FROM tbl_Continent WHERE db_NotVisible=0 ORDER BY id_Continent;

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE usp_ExpAccountDetailsSelect
@iExpertID int
AS
BEGIN

SELECT U.*
FROM tbl_Users U
INNER JOIN tbl_Experts E ON U.id_User=E.id_User
WHERE E.id_Expert=@iExpertID

END

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE usp_ExpCvvADBCouSelect
@@iExpertID int  
AS  
SET NOCOUNT ON  
  
DECLARE @sExpCou nvarchar(255), @sExpStartYear varchar(4), @sExpEndYear varchar(4)  
DECLARE @sExpCouFlag nvarchar(255), @sExpPeriod varchar(2000)  
  
DECLARE ExpertExperienceCursor CURSOR FAST_FORWARD FOR  
SELECT DISTINCT tbl_Country.couNameEng, Year(lnkExp_Wke.wkeStartDate) As wkeStartYear, Year(lnkExp_Wke.wkeEndDate) As wkeEndYear   
FROM lnkExp_Wke  
INNER join lnkWke_Cou ON lnkExp_Wke.id_ExpWke = lnkWke_Cou.id_ExpWke INNER JOIN tbl_Country ON lnkWke_Cou.id_Country = tbl_Country.id_Country   
WHERE lnkExp_Wke.id_Expert=@@iExpertID  
ORDER BY tbl_Country.couNameEng, wkeEndYear DESC
  
CREATE TABLE #tmp_ExpADBCou  
(  
tmpExpCou nvarchar(255),  
tmpExpPeriod varchar(2000)  
)  
  
OPEN ExpertExperienceCursor  
FETCH NEXT FROM ExpertExperienceCursor  
INTO @sExpCou, @sExpStartYear, @sExpEndYear  
  
SET @sExpPeriod=''  
SET @sExpCouFlag=@sExpCou  
  
WHILE @@FETCH_STATUS = 0  
BEGIN  
 IF @sExpCouFlag=@sExpCou  
 BEGIN  
  IF @sExpStartYear=@sExpEndYear  
   SET @sExpPeriod=@sExpPeriod + @sExpStartYear + ', '  
  ELSE  
   SET @sExpPeriod=@sExpPeriod + @sExpStartYear + ' - ' + @sExpEndYear + ', '  
 END  
 ELSE  
 BEGIN  
  IF LEN(@sExpPeriod)>2   
   SET @sExpPeriod=LEFT(@sExpPeriod, LEN(@sExpPeriod)-1)  
  INSERT INTO #tmp_ExpADBCou VALUES (@sExpCouFlag, @sExpPeriod)  
  SET @sExpCouFlag=@sExpCou  
  SET @sExpPeriod=''  
  IF @sExpStartYear=@sExpEndYear  
   SET @sExpPeriod=@sExpPeriod + @sExpStartYear + ', '  
  ELSE  
   SET @sExpPeriod=@sExpPeriod + @sExpStartYear + ' - ' + @sExpEndYear + ', '  
 END  
    
 FETCH NEXT FROM ExpertExperienceCursor  
 INTO @sExpCou, @sExpStartYear, @sExpEndYear  
END  
IF LEN(@sExpPeriod)>2   
SET @sExpPeriod=LEFT(@sExpPeriod, LEN(@sExpPeriod)-1)  
INSERT INTO #tmp_ExpADBCou VALUES (@sExpCouFlag, @sExpPeriod)  
  
CLOSE ExpertExperienceCursor  
DEALLOCATE ExpertExperienceCursor  
  
SET NOCOUNT OFF  
  
SELECT tmpExpCou, tmpExpPeriod FROM #tmp_ExpADBCou ORDER BY tmpExpPeriod DESC
  
SET NOCOUNT ON  
DROP TABLE #tmp_ExpADBCou  
SET NOCOUNT OFF

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO



SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE usp_ExpCvvAddressInsert
@@sLanguage varchar(3), @@iExpertID int, @@iAddressTypeID int, @@sStreet nvarchar(255), @@sPostcode nvarchar(50), @@sCity nvarchar(150), @@iCoountryID int, @@sPhone nvarchar(150), @@sMobile nvarchar(150), @@sFax nvarchar(150), @@sEmail nvarchar(150), @@sWebsite nvarchar(255)
AS
SET NOCOUNT ON
IF EXISTS(SELECT id_Address FROM tbl_Exp_Address WHERE id_Expert=@@iExpertID AND id_AddressType=@@iAddressTypeID)
	BEGIN
	UPDATE tbl_Exp_Address SET adrStreetEng=@@sStreet, adrPostCode=@@sPostcode, adrCityEng=@@sCity, id_Country=@@iCoountryID, adrPhone=@@sPhone, adrMobile=@@sMobile, adrFax=@@sFax, adrEmail=@@sEmail, adrWeb=@@sWebsite
	WHERE id_Expert=@@iExpertID AND id_AddressType=@@iAddressTypeID
	END
ELSE
	BEGIN
	INSERT INTO tbl_Exp_Address (id_Expert, id_AddressType, adrStreetEng, adrPostCode, adrCityEng, id_Country, adrPhone, adrMobile, adrFax, adrEmail, adrWeb)
	VALUES (@@iExpertID, @@iAddressTypeID, @@sStreet, @@sPostcode, @@sCity, @@iCoountryID, @@sPhone, @@sMobile, @@sFax, @@sEmail, @@sWebsite)
	END

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE usp_ExpCvvAddressSelect
@@iExpertID int, @@iAddressTypeID int
AS
SELECT DISTINCT EA.id_Address, C.couNameEng, C.couNameFra, C.couNameSpa, EA.id_Country, EA.adrStreetEng, EA.adrPostCode, EA.adrCityEng, EA.adrPhone, EA.adrMobile, EA.adrFax, EA.adrEmail, EA.adrWeb
FROM tbl_Exp_Address EA LEFT OUTER JOIN tbl_Country C on EA.id_Country=C.id_Country 
WHERE EA.id_Expert=@@iExpertID AND id_AddressType=@@iAddressTypeID
ORDER BY EA.id_Address DESC

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE usp_ExpCvvAddressUpdate
@sLanguage varchar(3), 
@iExpertID int, 
@iAddressTypeID int, 
@iExpAddressID int, 
@sStreet nvarchar(255), 
@sPostcode nvarchar(50), 
@sCity nvarchar(150), 
@iCountryID int, 
@sPhone nvarchar(150), 
@sMobile nvarchar(150), 
@sFax nvarchar(150), 
@sEmail nvarchar(150), 
@sWebsite nvarchar(255)
AS
SET NOCOUNT ON

BEGIN TRAN

-- Updating address
UPDATE tbl_Exp_Address 
SET adrStreetEng=@sStreet, adrPostCode=@sPostcode, adrCityEng=@sCity, id_Country=@iCountryID, adrPhone=@sPhone, adrMobile=@sMobile, adrFax=@sFax, adrEmail=@sEmail, adrWeb=@sWebsite
WHERE id_Expert=@iExpertID 
AND id_AddressType=@iAddressTypeID 
AND id_Address=@iExpAddressID

-- Checking and updating primary email
IF LEN(@sEmail)>5 AND CHARINDEX('@', @sEmail)>0 
	AND (NOT EXISTS(SELECT id_Expert 
			FROM tbl_Experts 
			WHERE id_Expert=@iExpertID
			 AND Email=@sEmail)) 
	AND (NOT EXISTS(SELECT E.id_Expert 
			FROM tbl_Experts E 
			INNER JOIN tbl_Exp_Address EA ON E.Email=EA.adrEmail AND E.id_Expert=EA.id_Expert 
			WHERE E.id_Expert=@iExpertID 
			AND EA.id_AddressType<4))
	BEGIN
	UPDATE tbl_Experts 
	SET Email=@sEmail
	WHERE id_Expert=@iExpertID
	END

COMMIT TRAN


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE usp_ExpCvvAvailabilityUpdate
@sLanguage varchar(3), 
@iExpertID int, 
@sAvailability nvarchar(400), 
@iShortterm tinyint, 
@iLongterm tinyint, 
@sPreferences ntext, 
@sOtherSkills ntext, 
@sMembership ntext, 
@sPublications ntext, 
@sReferences ntext
AS
UPDATE tbl_Experts 
SET expAvailabilityEng=@sAvailability, 
expShortterm=@iShortterm, 
expLongterm=@iLongterm, 
expPreferences=@sPreferences, 
expOtherSkills=@sOtherSkills,
expMemberProfEng=@sMembership, 
expPublicationsEng=@sPublications, 
expReferencesEng=@sReferences
WHERE id_Expert=@iExpertID


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE usp_ExpCvvBring2Assortis
@iExpertID int, @bIsUpdated int output
AS
IF EXISTS(SELECT id_Expert FROM tbl_Experts WHERE id_Expert=@iExpertID AND expIbfOnly=1)
	BEGIN
	UPDATE tbl_Experts 
	SET expIbfOnly=0, expDeleted=0, expRemoved=0
	WHERE id_Expert=@iExpertID 
	SET @bIsUpdated=@@ROWCOUNT
	END

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE usp_ExpCvvCommentsUpdate
@iExpertID int,
@sComments ntext
AS
SET NOCOUNT ON

UPDATE tbl_Experts
SET expComments=@sComments
WHERE id_Expert=@iExpertID


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE usp_ExpCvvECCouSelect
@@iExpertID int    
AS    
SET NOCOUNT ON    
    
DECLARE @sExpCou nvarchar(255), @sExpStartDate varchar(10), @sExpEndDate varchar(12), @sExpPrjTitle nvarchar(255)  
DECLARE @sExpStartDate1 varchar(200), @sExpEndDate1 varchar(200), @sExpPrjTitle1 nvarchar(3500)  
DECLARE @sExpCouFlag nvarchar(255)  
    
DECLARE ExpertExperienceCursor CURSOR FAST_FORWARD FOR    
SELECT C.couNameEng, ISNULL(dbo.CONVERTDATE(EW.wkeStartDate), ' ') AS wkeStartYear, ISNULL(dbo.CONVERTDATE(EW.wkeEndDate), ' ') AS wkeEndYear, ISNULL(EW.wkePrjTitleEng, ISNULL(EW.wkeOrgNameEng, ISNULL(EW.wkePositionEng, ' '))) AS wkePrjTitleEng
FROM lnkExp_Wke EW INNER join lnkWke_Cou WC ON EW.id_ExpWke = WC.id_ExpWke INNER JOIN tbl_Country C ON WC.id_Country = C.id_Country  
INNER JOIN tbl_CountryDev CG ON C.id_Country=CG.id_Country  
WHERE CG.couDev=1 AND ((EW.wkeStartDate IS NOT NULL) OR (EW.wkeEndDate IS NOT NULL)) AND EW.id_Expert=@@iExpertID
ORDER BY C.couNameEng, EW.wkeStartYear DESC  
    
CREATE TABLE #tmp_ExpECCou    
(    
tmpExpCou nvarchar(255),    
tmpExpStartDate varchar(200),  
tmpExpEndDate varchar(200),  
tmpExpPrjTitle nvarchar(3500)  
)    
    
OPEN ExpertExperienceCursor    
FETCH NEXT FROM ExpertExperienceCursor    
INTO @sExpCou, @sExpStartDate, @sExpEndDate, @sExpPrjTitle  
    
SET @sExpCouFlag=@sExpCou    
SET @sExpStartDate1 = ''  
SET @sExpEndDate1 = ''  
SET @sExpPrjTitle1 = ''  
    
WHILE @@FETCH_STATUS = 0    
BEGIN    
 IF @sExpCouFlag=@sExpCou    
 BEGIN    
   -- Add #-# symbols to split data grouped by country name  
   SET @sExpStartDate1 = @sExpStartDate1 + IsNull(@sExpStartDate,'') + '#-#'  
   SET @sExpEndDate1 = @sExpEndDate1 + IsNull(@sExpEndDate,'') + '#-#'  
   SET @sExpPrjTitle1 = @sExpPrjTitle1 + IsNull(@sExpPrjTitle,'') + '#-#'  
 END    
 ELSE    
 BEGIN    
  
  IF RIGHT(@sExpPrjTitle1,3)='#-#'     
 BEGIN  
 SET @sExpStartDate1=LEFT(@sExpStartDate1, LEN(@sExpStartDate1)-3)    
 SET @sExpEndDate1=LEFT(@sExpEndDate1, LEN(@sExpEndDate1)-3)  
 SET @sExpPrjTitle1=LEFT(@sExpPrjTitle1, LEN(@sExpPrjTitle1)-3)  
 END  
  INSERT INTO #tmp_ExpECCou VALUES (@sExpCouFlag, @sExpStartDate1, @sExpEndDate1, @sExpPrjTitle1)  
  
  SET @sExpCouFlag=@sExpCou    
  SET @sExpStartDate1 = @sExpStartDate + '#-#'  
  SET @sExpEndDate1 = @sExpEndDate + '#-#'  
  SET @sExpPrjTitle1 = @sExpPrjTitle + '#-#'  
 END    
      
 FETCH NEXT FROM ExpertExperienceCursor    
 INTO @sExpCou, @sExpStartDate, @sExpEndDate, @sExpPrjTitle  
END    
  
  
IF RIGHT(@sExpPrjTitle1,3)='#-#'     
 BEGIN  
 SET @sExpStartDate1=LEFT(@sExpStartDate1, LEN(@sExpStartDate1)-3)    
 SET @sExpEndDate1=LEFT(@sExpEndDate1, LEN(@sExpEndDate1)-3)    
 SET @sExpPrjTitle1=LEFT(@sExpPrjTitle1, LEN(@sExpPrjTitle1)-3)    
 END  
INSERT INTO #tmp_ExpECCou VALUES (@sExpCouFlag, @sExpStartDate1, @sExpEndDate1, @sExpPrjTitle1)  
  
    
CLOSE ExpertExperienceCursor    
DEALLOCATE ExpertExperienceCursor    
    
SET NOCOUNT OFF    
    
SELECT tmpExpCou, tmpExpStartDate, tmpExpEndDate, tmpExpPrjTitle FROM #tmp_ExpECCou ORDER BY tmpExpStartDate DESC  
    
SET NOCOUNT ON    
DROP TABLE #tmp_ExpECCou    
SET NOCOUNT OFF

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE usp_ExpCvvEducationDelete
@@iExpertID int, @@iExpEducationID int
AS
SET NOCOUNT ON
DECLARE @iError int, @iNumberRecordsDeleted int
SET @iError=0
SET @iNumberRecordsDeleted=0
	DELETE FROM lnkExp_Edu WHERE id_Expert=@@iExpertID AND id_ExpEdu=@@iExpEducationID
	SET @iNumberRecordsDeleted=@@ROWCOUNT
	SET @iError=@iError + @@ERROR
RETURN @iNumberRecordsDeleted

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE usp_ExpCvvEducationInfoSelect
@iExpertID int, 
@iExpEduID int, 
@iEducationTypeID int
AS
SELECT ES.edsDescriptionEng, 
ET.edtDescriptionEng, 
EE.InstNameEng, 
EE.eduOtherEng, 
EE.eduStartDate, 
EE.eduEndDate, 
EE.eduDiploma1Eng, 
EE.id_eduSubject, 
EE.id_eduSubject1Eng, 
EE.eduDiploma, 
EE.InstLocationEng, 
eduDiploma1Eng, 
id_ExpEdu, 
eduDescriptionEng
FROM lnkExp_Edu EE 
LEFT OUTER JOIN tbl_EduSubjects ES ON EE.id_EduSubject=ES.id_EduSubject 
LEFT OUTER JOIN tbl_EducationType ET ON EE.eduDiploma=ET.id_EduType 
WHERE EE.id_EduType=@iEducationTypeID AND EE.id_Expert=@iExpertID AND EE.id_ExpEdu=@iExpEduID
ORDER BY eduEndDate DESC


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE usp_ExpCvvEducationInsert
@@iExpertID int, @@iEduType int, @@iEduSubjectID smallint, @@iEduDiplomaID smallint, @@sEduSubject1 nvarchar(150), @@sEduDiploma1 nvarchar(150), @@sUserLanguage varchar(3)
AS
SET NOCOUNT ON
IF @@sUserLanguage='Eng' 
	BEGIN
	INSERT INTO lnkExp_Edu (id_Expert, id_EduType, id_EduSubject, eduDiploma, id_EduSubject1Eng, eduDiploma1Eng) 
	VALUES (@@iExpertID, @@iEduType, @@iEduSubjectID, @@iEduDiplomaID, @@sEduSubject1, @@sEduDiploma1)
	END

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE usp_ExpCvvEducationInsertNew
@@iExpertID int, @@iExpEduTypeID smallint, @@iEduSubjectID smallint, @@iEduDiplomaID smallint, @@sEduSubject1 nvarchar(150), @@sEduDiploma1 nvarchar(150), @@sUserLanguage varchar(3), @@sEduInstitution nvarchar(255), @@sEduLocaition nvarchar(255), @@sEduStartDate varchar(16), @@sEduEndDate varchar(16)
AS
SET NOCOUNT ON
IF @@sUserLanguage='Eng' 
	BEGIN
	INSERT INTO lnkExp_Edu (id_Expert, id_EduType, id_EduSubject, eduDiploma, id_EduSubject1Eng, eduDiploma1Eng, InstNameEng, InstLocationEng, eduStartDate, eduEndDate) 
	VALUES (@@iExpertID, @@iExpEduTypeID, @@iEduSubjectID, @@iEduDiplomaID, @@sEduSubject1, @@sEduDiploma1, @@sEduInstitution, @@sEduLocaition, @@sEduStartDate, @@sEduEndDate)
	END

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE usp_ExpCvvEducationSelect
@@iExpertID int, @@iEducationTypeID int
AS
SELECT ES.edsDescriptionEng, ET.edtDescriptionEng, EE.InstNameEng, EE.eduOtherEng, EE.eduStartDate, EE.eduEndDate, EE.eduDiploma1Eng, EE.id_eduSubject1Eng, EE.eduDiploma, EE.InstLocationEng, eduDiploma1Eng, id_ExpEdu, eduDescriptionEng
FROM lnkExp_Edu EE LEFT OUTER JOIN tbl_EduSubjects ES ON EE.id_EduSubject=ES.id_EduSubject LEFT OUTER JOIN tbl_EducationType ET on EE.eduDiploma=ET.id_EduType 
WHERE EE.id_EduType=@@iEducationTypeID and EE.id_Expert=@@iExpertID 
ORDER BY ISNULL(eduStartDate,eduEndDate) DESC, ISNULL(eduEndDate, eduStartDate) DESC

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE usp_ExpCvvEducationUpdate
@@iExpertID int, @@iExpEduID int, @@iEduType int, @@iEduSubjectID smallint, @@iEduDiplomaID smallint, @@sEduSubject1 nvarchar(150), @@sEduDiploma1 nvarchar(150), @@sUserLanguage varchar(3), @@sEduInstitution nvarchar(255), @@sEduLocaition nvarchar(255), @@sEduStartDate varchar(16), @@sEduEndDate varchar(16)
AS
SET NOCOUNT ON
IF @@sUserLanguage='Eng' 
	BEGIN
	UPDATE lnkExp_Edu SET id_EduSubject=@@iEduSubjectID, eduDiploma=@@iEduDiplomaID, id_EduSubject1Eng=@@sEduSubject1, eduDiploma1Eng=@@sEduDiploma1, 
	InstNameEng=@@sEduInstitution, InstLocationEng=@@sEduLocaition, eduStartDate=@@sEduStartDate, eduEndDate=@@sEduEndDate
	WHERE id_Expert=@@iExpertID AND id_ExpEdu=@@iExpEduID AND id_EduType=@@iEduType
	END

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE usp_ExpCvvExpInfoSelect
@iExpertID int
AS
SELECT P.id_person, 
P.psnLastNameEng, P.psnLastNameFra, P.psnLastNameSpa, 
P.psnFirstNameEng, P.psnFirstNameFra, P.psnFirstNameSpa, 
P.psnMiddleNameEng, P.psnMiddleNameFra, P.psnMiddleNameSpa, 
P.id_psnTitle, 
P.psnGender, 
P.psnBirthDate, 
P.psnBirthPlaceEng, 
P.id_MaritalStatus, 
E.id_Expert, 
E.expRegNumber,
E.id_ProfessionalStatus,
E.expProfessionEng, 
E.expProfYears, 
E.Lng, 
E.expCurrPositionEng, 
E.expKeyQualificationsEng, 
E.expMemberProfEng, 
E.expPublicationsEng, 
E.expReferencesEng, 
E.expAvailabilityEng, 
E.expShortterm, 
E.expLongterm, 
E.expPreferences,
E.expOtherSkills,
E.Phone, 
E.Email,
E.expComments,
E.expHidden, 
E.expIncompleteCV, 
E.expToConfirmCvEmailSent, 
E.expToConfirmCvEmailDate, 
E.expToCompleteCvEmailSent, 
E.expToCompleteCvEmailDate, 
E.expApproved, 
E.expApprovedDate, 
E.expRemoved, 
E.expRemovedDate, 
E.expDeleted, 
E.expDeletedDate,
E.expAccountEmailSent
FROM tbl_Experts E 
INNER JOIN tbl_Persons P ON E.id_Expert=P.id_Expert 
WHERE E.id_Expert=@iExpertID
AND expDeleted=0
AND expRemoved=0


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE usp_ExpCvvExpInfoUpdate
@@iExpTitleID tinyint, @@sExpFirstName nvarchar(255), @@sExpMiddleName nvarchar(255), @@sExpLastName nvarchar(255), 
@@dExpBirthDate smalldatetime, @@sExpBirthPlace nvarchar(255), @@iExpGenderID tinyint, @@iExpMaritalStatusID tinyint, 
@@sExpCurPosition nvarchar(255), @@sExpProfession nvarchar(255), @@sExpKeyQualifications ntext, @@iExpProfYears int,
@@iExpertID int, @@sUserLanguage varchar(3)
AS
SET NOCOUNT ON
IF @@sUserLanguage='Eng'
	BEGIN
	TRANSACTION BEGIN
	UPDATE tbl_Persons 
	SET id_psnTitle=@@iExpTitleID, psnFirstNameEng=@@sExpFirstName, psnMiddleNameEng=@@sExpMiddleName, psnLastNameEng=@@sExpLastName, 
	psnBirthDate=@@dExpBirthDate, psnBirthPlaceEng=@@sExpBirthPlace, psnGender=@@iExpGenderID, id_MaritalStatus=@@iExpMaritalStatusID
	WHERE id_Expert=@@iExpertID
	UPDATE tbl_Experts 
	SET expCurrPositionEng=@@sExpCurPosition, expProfessionEng=@@sExpProfession, expKeyQualificationsEng=@@sExpKeyQualifications, expProfYears=@@iExpProfYears
	WHERE id_Expert=@@iExpertID
	COMMIT
	END

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE usp_ExpCvvExperienceBriefInsert
@@id_Expert int, @@expCurrPosition nvarchar(255), @@expMSct varchar(400), @@expReg varchar(400)
AS
SET NOCOUNT ON
DECLARE @WKEID int
DECLARE @arrMSct varchar(400), @arrReg varchar(400)
SET @arrMSct=@@expMSct
SET @arrReg=@@expReg
DECLARE @curMSct int, @curReg int, @ci int
DECLARE @id_Sector int, @id_Country int
INSERT INTO lnkExp_Wke
(id_Expert, wkePositionEng) 
VALUES (@@id_Expert,@@expCurrPosition)
SELECT @WKEID=id_ExpWke FROM lnkExp_Wke
WHILE LEN(@arrMSct)>0
BEGIN
	SET @ci=CHARINDEX(',',@arrMSct)
	IF @ci>0
	BEGIN
		SET @curMSct=CONVERT(int, LEFT(@arrMSct, @ci-1))
		SET @arrMSct=SUBSTRING(@arrMSct, @ci+1, LEN(@arrMSct))
	END
	ELSE
	BEGIN
		SET @curMSct=CONVERT(int, @arrMSct)
		SET @arrMSct=''
	END
	-- Extracting subsectors for selected main sectors and insert them in WKE
	DECLARE tbl_Sectors_cursor CURSOR FOR
	SELECT DISTINCT id_Sector FROM tbl_Sectors WHERE id_Sector>1000 AND id_MainSector = @curMSct
	OPEN tbl_Sectors_cursor
		
	FETCH NEXT FROM tbl_Sectors_cursor
	INTO @id_Sector
	WHILE @@FETCH_STATUS = 0
	BEGIN
		INSERT INTO lnkWke_Sct (id_ExpWke, id_Sector)
		VALUES(@WKEID, @id_Sector)
		FETCH NEXT FROM tbl_Sectors_cursor
		INTO @id_Sector
	END
	CLOSE tbl_Sectors_cursor
	DEALLOCATE tbl_Sectors_cursor
END
WHILE LEN(@arrReg)>0
BEGIN
	SET @ci=CHARINDEX(',',@arrReg)
	IF @ci>0
	BEGIN
		SET @curReg=CONVERT(int, LEFT(@arrReg, @ci-1))
		SET @arrReg=SUBSTRING(@arrReg, @ci+1, LEN(@arrReg))
	END
	ELSE
	BEGIN
		SET @curReg=CONVERT(int, @arrReg)
		SET @arrReg=''
	END
	-- Extracting countries for selected regions and insert them in WKE
	DECLARE tbl_Country_cursor CURSOR FOR
	SELECT DISTINCT id_Country FROM tbl_Country WHERE id_Country>1000 AND id_GeoZone=@curReg
	OPEN tbl_Country_cursor
		
	FETCH NEXT FROM tbl_Country_cursor
	INTO @id_Country
	WHILE @@FETCH_STATUS = 0
	BEGIN
		INSERT INTO lnkWke_Cou (id_ExpWke, id_Country)
		VALUES(@WKEID, @id_Country)
		FETCH NEXT FROM tbl_Country_cursor
		INTO @id_Country
	END
	CLOSE tbl_Country_cursor
	DEALLOCATE tbl_Country_cursor
END

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE usp_ExpCvvExperienceCouDelete
@@iExpertID int, @@iExpWkeID int
AS
SET NOCOUNT ON
IF EXISTS(SELECT id_ExpWke FROM lnkExp_Wke WHERE id_Expert=@@iExpertID AND id_ExpWke=@@iExpWkeID)
	BEGIN
	DELETE FROM lnkWke_Cou WHERE id_ExpWke=@@iExpWkeID 
	RETURN @@ROWCOUNT
	END

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE usp_ExpCvvExperienceCouInsert  
@@sCountries varchar(2000), @@iExpertID int, @@iExpWkeID int, @@iTotalCountries int output
AS  
DECLARE @iCountryID int, @iPosition int  
  
SET @@sCountries=REPLACE(@@sCountries, ' 0,', '')  
SET @@sCountries=REPLACE(@@sCountries, ' ', '')  
SET @@iTotalCountries=0
  
WHILE LEN(@@sCountries)>0  
BEGIN  
	SET @iCountryID=0  
	SET @iPosition=CHARINDEX(',', @@sCountries)  
	IF @iPosition>0   
		BEGIN  
		SET @iCountryID=CONVERT(int, LEFT(@@sCountries, @iPosition-1))  
		SET @@sCountries=RIGHT(@@sCountries, LEN(@@sCountries)-@iPosition)  
		END  
	ELSE  
		BEGIN  
		SET @iCountryID=CONVERT(int, @@sCountries)  
		SET @@sCountries=''  
		END  
	IF @iCountryID>0 
		BEGIN
		INSERT INTO lnkWke_Cou (id_Country, id_ExpWke) VALUES (@iCountryID, @@iExpWkeID)
		IF @@ROWCOUNT>0 SET @@iTotalCountries = @@iTotalCountries + 1
		END
END  

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE usp_ExpCvvExperienceCouSelect
@@iExpertID int, @@iExpWkeID int, @@sResultsType varchar(10), @@sCountriesList nvarchar(4000) OUTPUT
AS
SET NOCOUNT ON
IF EXISTS (SELECT id FROM tempdb.dbo.sysobjects WHERE NAME = '#tblFieldsTable')
DROP TABLE #tblFieldsTable
IF @@sResultsType='rs'
	BEGIN
	SET NOCOUNT OFF
	IF @@iExpWkeID>0
		BEGIN
		SELECT DISTINCT R.id_GeoZone, R.Geo_ZoneEng, C.id_Country, C.couNameEng
		FROM lnkWke_Cou WC INNER JOIN lnkExp_Wke EW ON WC.id_ExpWke=EW.id_ExpWke INNER JOIN tbl_Country C ON WC.id_Country=C.id_Country INNER JOIN tbl_GeoZone R ON C.id_GeoZone=R.id_GeoZone
		WHERE EW.id_Expert=@@iExpertID AND EW.id_ExpWke=@@iExpWkeID
		ORDER BY R.id_GeoZone
		END
	ELSE
		BEGIN
		-- union is used to replace coutries with id>1000 (all countries from the geozone) on list of counties
		SELECT CVC.id_Country, TGZ.id_GeoZone FROM
		(SELECT C.id_Country
		FROM lnkWke_Cou WC INNER JOIN lnkExp_Wke EW ON WC.id_ExpWke=EW.id_ExpWke INNER JOIN tbl_GeoZone GZ ON WC.id_Country=GZ.id_GeoZone+1000 INNER JOIN tbl_Country C ON GZ.id_GeoZone=C.id_GeoZone
		WHERE EW.id_Expert=@@iExpertID AND C.id_Country<1000
		UNION
		SELECT WC.id_Country
		FROM lnkWke_Cou WC INNER JOIN lnkExp_Wke EW ON WC.id_ExpWke=EW.id_ExpWke
		WHERE EW.id_Expert=@@iExpertID AND WC.id_Country<1000) AS CVC
		INNER JOIN tbl_Country TC ON CVC.id_Country=TC.id_Country INNER JOIN tbl_GeoZone TGZ ON TC.id_GeoZone=TGZ.id_GeoZone 
		ORDER BY TGZ.id_GeoZone
		END
	END
ELSE
	BEGIN
	SET NOCOUNT ON
	SELECT DISTINCT R.id_GeoZone, R.Geo_ZoneEng, C.id_Country, C.couNameEng
	INTO #tblFieldsTable
	FROM lnkWke_Cou WC INNER JOIN lnkExp_Wke EW ON WC.id_ExpWke=EW.id_ExpWke INNER JOIN tbl_Country C ON WC.id_Country=C.id_Country INNER JOIN tbl_GeoZone R ON C.id_GeoZone=R.id_GeoZone
	WHERE EW.id_Expert=@@iExpertID AND EW.id_ExpWke=@@iExpWkeID
	ORDER BY R.id_GeoZone
	IF @@sResultsType='listshort'
		EXEC usp_AdmCreateList3FromRecordset 'row', @@sCountriesList OUTPUT
	ELSE
		EXEC usp_AdmCreateList2FromRecordset 'row', @@sCountriesList OUTPUT
	IF EXISTS (SELECT id FROM tempdb.dbo.sysobjects WHERE NAME = '#tblFieldsTable')
	DROP TABLE #tblFieldsTable
	END

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE usp_ExpCvvExperienceDelete
@@iExpertID int, @@iExpWkeID int
AS
SET NOCOUNT ON
DECLARE @iError int, @iNumberRecordsDeleted int
SET @iError=0
SET @iNumberRecordsDeleted=0
BEGIN TRANSACTION
	DELETE FROM lnkExp_Wke WHERE id_Expert=@@iExpertID AND id_ExpWke=@@iExpWkeID
	SET @iNumberRecordsDeleted=@@ROWCOUNT
	SET @iError=@iError + @@ERROR
	DELETE FROM lnkWke_Cou WHERE id_ExpWke=@@iExpWkeID
	SET @iError=@iError + @@ERROR
	DELETE FROM lnkWke_Don WHERE id_ExpWke=@@iExpWkeID
	SET @iError=@iError + @@ERROR
	DELETE FROM lnkWke_Sct WHERE id_ExpWke=@@iExpWkeID
	SET @iError=@iError + @@ERROR
IF @iError=0
	COMMIT TRANSACTION
ELSE
	ROLLBACK TRANSACTION
RETURN @iNumberRecordsDeleted

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE usp_ExpCvvExperienceDonDelete
@@iExpertID int, @@iExpWkeID int
AS
SET NOCOUNT ON
IF EXISTS(SELECT id_ExpWke FROM lnkExp_Wke WHERE id_Expert=@@iExpertID AND id_ExpWke=@@iExpWkeID)
	BEGIN
	DELETE FROM lnkWke_Don WHERE id_ExpWke=@@iExpWkeID 
	RETURN @@ROWCOUNT
	END

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE usp_ExpCvvExperienceDonInsert
@@sDonors varchar(2000), @@iExpertID int, @@iExpWkeID int, 
@@iTotalDonors int output
AS
DECLARE @iDonorID int, @iPosition int
SET @@sDonors=REPLACE(@@sDonors, ' 0,', '')
SET @@sDonors=REPLACE(@@sDonors, ' ', '')
SET @@iTotalDonors=0
WHILE LEN(@@sDonors)>0
BEGIN
	SET @iDonorID=0
	SET @iPosition=CHARINDEX(',', @@sDonors)
	IF @iPosition>0 
		BEGIN
		SET @iDonorID=CONVERT(int, LEFT(@@sDonors, @iPosition-1))
		SET @@sDonors=RIGHT(@@sDonors, LEN(@@sDonors)-@iPosition)
		END
	ELSE
		BEGIN
		SET @iDonorID=CONVERT(int, @@sDonors)
		SET @@sDonors=''
		END
	IF @iDonorID>0 
		BEGIN
		INSERT INTO lnkWke_Don (id_Organisation, id_ExpWke) VALUES (@iDonorID, @@iExpWkeID)
		IF @@ROWCOUNT>0 SET @@iTotalDonors = @@iTotalDonors + 1
		END
END

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE usp_ExpCvvExperienceDonSelect
@@iExpertID int, @@iExpWkeID int, @@sResultsType varchar(10), @@sDonorsList nvarchar(4000) OUTPUT
AS
SET NOCOUNT ON
IF EXISTS (SELECT id FROM tempdb.dbo.sysobjects WHERE NAME = '#tblFieldsTable')
DROP TABLE #tblFieldsTable
IF @@sResultsType='rs'
	BEGIN
	SET NOCOUNT OFF
	SELECT DISTINCT id_Organisation
	FROM lnkWke_Don WD INNER JOIN lnkExp_Wke EW ON WD.id_ExpWke=EW.id_ExpWke 
	WHERE EW.id_Expert=@@iExpertID AND EW.id_ExpWke=@@iExpWkeID
	END
ELSE
	BEGIN
	SET NOCOUNT ON
	SELECT DISTINCT D.id_Organisation, D.orgNameEng
	INTO #tblFieldsTable
	FROM lnkWke_Don WD INNER JOIN lnkExp_Wke EW ON WD.id_ExpWke=EW.id_ExpWke INNER JOIN tbl_Donors D ON WD.id_Organisation=D.id_Organisation 
	WHERE EW.id_Expert=@@iExpertID AND EW.id_ExpWke=@@iExpWkeID AND D.orgNameEng IS NOT NULL
	UNION 
	SELECT DISTINCT WD.id_Organisation, WD.wkd_OtherNameEng
	FROM lnkWke_Don WD INNER JOIN lnkExp_Wke EW ON WD.id_ExpWke=EW.id_ExpWke
	WHERE EW.id_Expert=@@iExpertID AND EW.id_ExpWke=@@iExpWkeID and WD.wkd_OtherNameEng IS NOT NULL
	UNION 
	SELECT DISTINCT 0, EW.wkeDonorEng
	FROM lnkExp_Wke EW 
	WHERE EW.id_Expert=@@iExpertID AND EW.id_ExpWke=@@iExpWkeID AND EW.wkeDonorEng IS NOT NULL
	EXEC usp_AdmCreateList1FromRecordset 'row', @@sDonorsList OUTPUT
	IF EXISTS (SELECT id FROM tempdb.dbo.sysobjects WHERE NAME = '#tblFieldsTable')
	DROP TABLE #tblFieldsTable
	END

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE usp_ExpCvvExperienceInfoSelect
@iExpertID int, 
@iExpWkeID int
AS
SELECT id_ExpWke, 
wkeStartDate, 
wkeEndDate, 
wkeEndDateOpen,
wkePeriod, 
wkeOrgNameEng, 
wkeOrgNameFra, 
wkeOrgNameSpa, 
wkeBnfNameEng, 
wkeBnfNameFra, 
wkeBnfNameSpa, 
wkePrjTitleEng, 
wkePrjTitleFra, 
wkePrjTitleSpa, 
wkePositionEng, 
wkePositionFra, 
wkePositionSpa, 
wkeDescriptionEng, 
wkeDescriptionFra, 
wkeDescriptionSpa, 
wkeClientRefEng, 
wkeClientRefFra, 
wkeClientRefSpa,
wkeRefName,
wkeRefPosition,
wkeRefEmail,
wkeRefPhone,
wkeRefExtended,
TypeofWke, 
wkeLocationEng, 
wkeLocationFra, 
wkeLocationSpa, 
wkeDonorEng,
wkeProjectDescription
FROM lnkExp_Wke 
WHERE id_Expert=@iExpertID 
AND id_ExpWke=@iExpWkeID


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE usp_ExpCvvExperienceInsert
@sLanguage varchar(3), 
@iExpertID int, 
@sProjectTitle nvarchar(255), 
@sOrganisation nvarchar(200), 
@sPosition nvarchar(255), 
@sBeneficiary nvarchar(200), 
@sReferences nvarchar(255), 
@sRefName nvarchar(255),
@sRefPosition nvarchar(255),
@sRefPhone varchar(150),
@sRefEmail nvarchar(150),
@sRefExtended tinyint,
@sDescription ntext, 
@sProjectDescription ntext,
@sDonor nvarchar(255), 
@sStartDate varchar(16), 
@sEndDate varchar(16),
@bOngoing tinyint,
@iType tinyint = 1,
@iExpWkeID int output 
AS
INSERT INTO lnkExp_Wke (
id_Expert, 
wkePrjTitleEng, 
wkeOrgNameEng, 
wkePositionEng, 
wkeBnfNameEng, 
wkeClientRefEng, 
wkeRefName,
wkeRefPosition,
wkeRefPhone,
wkeRefEmail,
wkeRefExtended,
wkeDescriptionEng, 
wkeDonorEng, 
wkeStartDate, 
wkeEndDate, 
wkeEndDateOpen,
wkeProjectDescription,
TypeofWke
) VALUES (
@iExpertID, 
@sProjectTitle, 
@sOrganisation, 
@sPosition, 
@sBeneficiary, 
@sReferences, 
@sRefName,
@sRefPosition,
@sRefPhone,
@sRefEmail,
@sRefExtended,
@sDescription, 
@sDonor, 
@sStartDate, 
@sEndDate, 
@bOngoing,
@sProjectDescription,
@iType
)

SELECT @iExpWkeID=@@IDENTITY FROM lnkExp_Wke

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE usp_ExpCvvExperienceSctDelete
@@iExpertID int, @@iExpWkeID int
AS
SET NOCOUNT ON
IF EXISTS(SELECT id_ExpWke FROM lnkExp_Wke WHERE id_Expert=@@iExpertID AND id_ExpWke=@@iExpWkeID)
	BEGIN
	DELETE FROM lnkWke_Sct WHERE id_ExpWke=@@iExpWkeID 
	RETURN @@ROWCOUNT
	END

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE usp_ExpCvvExperienceSctInsert
@@sSectors varchar(2000), @@iExpertID int, @@iExpWkeID int,
@@iTotalSectors int output
AS
DECLARE @iSectorID int, @iPosition int
SET @@sSectors=REPLACE(@@sSectors, ' 0,', '')
SET @@sSectors=REPLACE(@@sSectors, ' ', '')
SET @@iTotalSectors=0
WHILE LEN(@@sSectors)>0
BEGIN
	SET @iSectorID=0
	SET @iPosition=CHARINDEX(',', @@sSectors)
	IF @iPosition>0 
		BEGIN
		SET @iSectorID=CONVERT(int, LEFT(@@sSectors, @iPosition-1))
		SET @@sSectors=RIGHT(@@sSectors, LEN(@@sSectors)-@iPosition)
		END
	ELSE
		BEGIN
		SET @iSectorID=CONVERT(int, @@sSectors)
		SET @@sSectors=''
		END
	IF @iSectorID>0 
		BEGIN
		INSERT INTO lnkWke_Sct (id_Sector, id_ExpWke) VALUES (@iSectorID, @@iExpWkeID)
		IF @@ROWCOUNT>0 SET @@iTotalSectors = @@iTotalSectors + 1
		END
END

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE usp_ExpCvvExperienceSctSelect
@@iExpertID int, @@iExpWkeID int, @@sResultsType varchar(10), @@sSectorsList nvarchar(4000) OUTPUT
AS
SET NOCOUNT ON
IF EXISTS (SELECT id FROM tempdb.dbo.sysobjects WHERE NAME = '#tblFieldsTable')
DROP TABLE #tblFieldsTable
IF @@sResultsType='rs'
	BEGIN
	SET NOCOUNT OFF
	IF @@iExpWkeID>0
		BEGIN
		SELECT DISTINCT MS.id_MainSector, MS.mnsDescriptionEng, S.id_Sector, S.sctDescriptionEng
		FROM lnkWke_Sct WS INNER JOIN lnkExp_Wke EW ON WS.id_ExpWke=EW.id_ExpWke INNER JOIN tbl_Sectors S ON WS.id_Sector=S.id_Sector INNER JOIN tbl_MainSectors MS ON S.id_MainSector=MS.id_MainSector
		WHERE EW.id_Expert=@@iExpertID AND EW.id_ExpWke=@@iExpWkeID
		END
	ELSE
		BEGIN
		-- union is used to replace sectors with id>1000 (all sectors from the mainsector) on list of sectors
		SELECT CVS.id_Sector, TMS.id_MainSector FROM
		(SELECT S.id_Sector
		FROM lnkWke_Sct WS INNER JOIN lnkExp_Wke EW ON WS.id_ExpWke=EW.id_ExpWke INNER JOIN tbl_MainSectors MS ON WS.id_Sector=MS.id_MainSector+1000 INNER JOIN tbl_Sectors S ON MS.id_MainSector=S.id_MainSector
		WHERE EW.id_Expert=@@iExpertID AND S.id_Sector<1000
		UNION
		SELECT WS.id_Sector
		FROM lnkWke_Sct WS INNER JOIN lnkExp_Wke EW ON WS.id_ExpWke=EW.id_ExpWke
		WHERE EW.id_Expert=@@iExpertID AND WS.id_Sector<1000) AS CVS
		INNER JOIN tbl_Sectors TS ON CVS.id_Sector=TS.id_Sector INNER JOIN tbl_MainSectors TMS ON TS.id_MainSector=TMS.id_MainSector
		END
	END
ELSE
	BEGIN
	SET NOCOUNT ON
	SELECT DISTINCT MS.id_MainSector, MS.mnsDescriptionEng, S.id_Sector, S.sctDescriptionEng
	INTO #tblFieldsTable
	FROM lnkWke_Sct WS INNER JOIN lnkExp_Wke EW ON WS.id_ExpWke=EW.id_ExpWke INNER JOIN tbl_Sectors S ON WS.id_Sector=S.id_Sector INNER JOIN tbl_MainSectors MS ON S.id_MainSector=MS.id_MainSector
	WHERE EW.id_Expert=@@iExpertID AND EW.id_ExpWke=@@iExpWkeID
	EXEC usp_AdmCreateList2FromRecordset 'column', @@sSectorsList OUTPUT
	IF EXISTS (SELECT id FROM tempdb.dbo.sysobjects WHERE NAME = '#tblFieldsTable')
	DROP TABLE #tblFieldsTable
	END

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE usp_ExpCvvExperienceSelect
@iExpertID int
AS
SET NOCOUNT OFF
SELECT id_ExpWke, 
wkeStartDate, 
wkeEndDate, 
wkeEndDateOpen,
wkePeriod, 
wkeOrgNameEng, 
wkeOrgNameFra, 
wkeOrgNameSpa, 
wkeBnfNameEng, 
wkeBnfNameFra, 
wkeBnfNameSpa, 
wkePrjTitleEng, 
wkePrjTitleFra, 
wkePrjTitleSpa, 
wkePositionEng, 
wkePositionFra, 
wkePositionSpa, 
wkeDescriptionEng, 
wkeDescriptionFra, 
wkeDescriptionSpa, 
wkeClientRefEng, 
wkeClientRefFra, 
wkeClientRefSpa,
wkeRefFirstName,
wkeRefLastName,
wkeRefName,
wkeRefPosition,
wkeRefEmail,
wkeRefPhone,
wkeRefExtended,
TypeofWke, 
wkeLocationEng, 
wkeLocationFra, 
wkeLocationSpa,
wkeProjectDescription,
wkeInfoGroup
FROM lnkExp_Wke 
WHERE id_Expert=@iExpertID 
ORDER BY 
wkeInfoGroup,
wkeEndDateOpen DESC,
CASE WHEN wkeEndDate IS NULL AND wkeStartDate IS NULL THEN 1 ELSE 2 END,
ISNULL(wkeEndDate, DATEADD(y, 1, wkeStartDate)) DESC, 
wkeStartDate DESC,
id_ExpWke


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE usp_ExpCvvExperienceUpdate
@sLanguage varchar(3), 
@iExpertID int, 
@iExpWkeID int, 
@sProjectTitle nvarchar(255), 
@sOrganisation nvarchar(200), 
@sPosition nvarchar(255), 
@sBeneficiary nvarchar(200), 
@sReferences nvarchar(255), 
@sRefName nvarchar(255),
@sRefPosition nvarchar(255),
@sRefPhone varchar(150),
@sRefEmail nvarchar(150),
@sRefExtended tinyint,
@sDescription ntext, 
@sProjectDescription ntext,
@sDonor nvarchar(255), 
@sStartDate varchar(16), 
@sEndDate varchar(16),
@bOngoing tinyint,
@iType tinyint = 1
AS
UPDATE lnkExp_Wke 
SET wkePrjTitleEng=@sProjectTitle, 
wkeOrgNameEng=@sOrganisation, 
wkePositionEng=@sPosition,
wkeBnfNameEng=@sBeneficiary, 
wkeClientRefEng=@sReferences, 
wkeRefName=@sRefName,
wkeRefPosition=@sRefPosition,
wkeRefPhone=@sRefPhone,
wkeRefEmail=@sRefEmail,
wkeRefExtended=@sRefExtended,
wkeDescriptionEng=@sDescription, 
wkeDonorEng=@sDonor, 
wkeStartDate=@sStartDate, 
wkeEndDate=@sEndDate,
wkeEndDateOpen=@bOngoing,
wkeProjectDescription=@sProjectDescription,
TypeofWke=@iType
WHERE id_Expert=@iExpertID 
AND id_ExpWke=@iExpWkeID



GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE usp_ExpCvvLanguageInfoSelect
@@iExpertID int, @@iExpLanguageID int
AS
SET NOCOUNT OFF
SELECT L.lngNameEng, EL.id_Language, EL.exlReading, EL.exlSpeaking, EL.exlWriting, EL.exlReading+EL.exlSpeaking + EL.exlWriting  AS exlSummary 
FROM lnkExp_Lan EL INNER JOIN tbl_Languages L ON EL.id_Language=L.id_Language WHERE EL.id_Expert=@@iExpertID AND EL.id_ExpLan=@@iExpLanguageID

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE usp_ExpCvvLanguageInsert
@@id_Expert int, @@id_Language smallint, @@id_Read smallint, @@id_Speak smallint, @@id_Write smallint
AS
SET NOCOUNT ON
INSERT INTO lnkExp_Lan (id_Expert, id_Language, exlReading, exlSpeaking, exlWriting) 
VALUES (@@id_Expert, @@id_Language, @@id_Read, @@id_Speak, @@id_Write)

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE usp_ExpCvvLanguageNativeDelete
@@iExpertID int, @@iExpLngID int
AS
SET NOCOUNT ON
IF @@iExpLngID>0 
	DELETE FROM tbl_Native_Lng WHERE id_Expert=@@iExpertID AND id_Native=@@iExpLngID
ELSE
	DELETE FROM tbl_Native_Lng WHERE id_Expert=@@iExpertID
RETURN @@ROWCOUNT

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE usp_ExpCvvLanguageNativeInsert  
@@sLanguages varchar(2000), @@iExpertID int, @@iTotalLanguages int output
AS  
DECLARE @iLanguageID int, @iPosition int  
  
SET @@sLanguages=REPLACE(@@sLanguages, ' 0,', '')  
SET @@sLanguages=REPLACE(@@sLanguages, ' ', '')  
SET @@iTotalLanguages=0
  
WHILE LEN(@@sLanguages)>0  
BEGIN  
	SET @iLanguageID=0  
	SET @iPosition=CHARINDEX(',', @@sLanguages)  
	IF @iPosition>0   
		BEGIN  
		SET @iLanguageID=CONVERT(int, LEFT(@@sLanguages, @iPosition-1))  
		SET @@sLanguages=RIGHT(@@sLanguages, LEN(@@sLanguages)-@iPosition)  
		END  
	ELSE  
		BEGIN  
		SET @iLanguageID=CONVERT(int, @@sLanguages)  
		SET @@sLanguages=''  
		END  
	IF @iLanguageID>0 
		BEGIN
		INSERT INTO tbl_Native_Lng (id_Language, id_Expert) VALUES (@iLanguageID, @@iExpertID)
		IF @@ROWCOUNT>0 SET @@iTotalLanguages = @@iTotalLanguages + 1
		END
END  

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE usp_ExpCvvLanguageOtherDelete
@@iExpertID int, @@iExpLanguageID int
AS
SET NOCOUNT ON
DECLARE @iError int, @iNumberRecordsDeleted int
SET @iError=0
SET @iNumberRecordsDeleted=0
	DELETE FROM lnkExp_Lan WHERE id_Expert=@@iExpertID AND id_ExpLan=@@iExpLanguageID
	SET @iNumberRecordsDeleted=@@ROWCOUNT
	SET @iError=@iError + @@ERROR
RETURN @iNumberRecordsDeleted

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE usp_ExpCvvLanguageOtherInsert
@@iExpertID int, @@iLanguageID smallint, @@iReadingLevel smallint, @@iSpeakingLevel smallint, @@iWritingLevel smallint
AS
SET NOCOUNT ON
IF EXISTS(SELECT id_ExpLan FROM lnkExp_Lan WHERE id_Expert=@@iExpertID AND id_Language=@@iLanguageID)
	BEGIN
	UPDATE lnkExp_Lan SET exlReading=@@iReadingLevel, exlSpeaking=@@iSpeakingLevel, exlWriting=@@iWritingLevel 
	WHERE id_Expert=@@iExpertID AND id_Language=@@iLanguageID
	END
ELSE
	BEGIN
	INSERT INTO lnkExp_Lan (id_Expert, id_Language, exlReading, exlSpeaking, exlWriting) 
	VALUES (@@iExpertID, @@iLanguageID, @@iReadingLevel, @@iSpeakingLevel, @@iWritingLevel)
	END

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE usp_ExpCvvLanguageOtherUpdate
@@iExpertID int, @@iExpLngID int, @@iLanguageID smallint, @@iReadingLevel smallint, @@iSpeakingLevel smallint, @@iWritingLevel smallint
AS
SET NOCOUNT ON
UPDATE lnkExp_Lan SET id_Language=@@iLanguageID, exlReading=@@iReadingLevel, exlSpeaking=@@iSpeakingLevel, exlWriting=@@iWritingLevel 
WHERE id_Expert=@@iExpertID AND id_ExpLan=@@iExpLngID

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE usp_ExpCvvLanguageSelect
@iExpertID int, 
@sLanguageType varchar(10)
AS
SET NOCOUNT OFF
IF @sLanguageType='native' 
BEGIN
	SELECT NL.id_Language, 
	L.lngNameEng,
	L.lngNameFra,
	L.lngNameSpa,
	1 As exlReading, 
	1 As exlSpeaking, 
	1 As exlWriting, 
	1 AS exlSummary
	FROM tbl_Native_Lng NL 
	INNER JOIN tbl_Languages L ON NL.id_Language=L.id_Language 
	WHERE NL.id_Language NOT IN (SELECT EL.id_Language FROM lnkExp_Lan EL WHERE EL.id_Expert=@iExpertID) 
	AND NL.id_Expert=@iExpertID
END
ELSE IF @sLanguageType='other' 
BEGIN
	SELECT EL.id_ExpLan, 
	EL.id_Language, 
	L.lngNameEng, 
	L.lngNameFra,
	L.lngNameSpa,
	EL.exlReading, 
	EL.exlSpeaking, 
	EL.exlWriting, 
	EL.exlReading+EL.exlSpeaking + EL.exlWriting AS exlSummary 
	FROM lnkExp_Lan EL 
	INNER JOIN tbl_Languages L ON EL.id_Language=L.id_Language 
	WHERE EL.id_Expert=@iExpertID 
	ORDER BY exlSummary
END
ELSE
BEGIN
	SELECT L.lngNameEng, 
	L.lngNameFra,
	L.lngNameSpa,
	1 As exlReading, 
	1 As exlSpeaking, 
	1 As exlWriting, 
	1 AS exlSummary
	FROM tbl_Native_Lng NL 
	INNER JOIN tbl_Languages L ON NL.id_Language=L.id_Language 
	WHERE NL.id_Language NOT IN (SELECT EL.id_Language FROM lnkExp_Lan EL WHERE EL.id_Expert=@iExpertID) 
	AND NL.id_Expert=@iExpertID
	UNION
	SELECT L.lngNameEng, 
	L.lngNameFra,
	L.lngNameSpa,
	EL.exlReading, 
	EL.exlSpeaking, 
	EL.exlWriting, 
	EL.exlReading+EL.exlSpeaking + EL.exlWriting  AS exlSummary 
	FROM lnkExp_Lan EL 
	INNER JOIN tbl_Languages L ON EL.id_Language=L.id_Language 
	WHERE EL.id_Expert=@iExpertID 
	ORDER BY exlSummary
END

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE usp_ExpCvvNationalityDelete
@@id_Expert int
AS
SET NOCOUNT ON
DELETE FROM lnk_Exp_Nationality 
WHERE id_Expert=@@id_Expert

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE usp_ExpCvvNationalityInsert
@@id_Expert int, @@id_Nationality smallint
AS
SET NOCOUNT ON
INSERT INTO lnk_Exp_Nationality (id_Expert, id_Nationality) 
VALUES (@@id_Expert, @@id_Nationality)

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE usp_ExpCvvNationalitySelect
@@iExpertID int
AS
SELECT DISTINCT C.couNameEng, C.couNameFra, C.couNameSpa, EN.id_Nationality
FROM lnk_Exp_Nationality EN INNER JOIN tbl_Country C on EN.id_Nationality=C.id_Country 
WHERE EN.id_Expert=@@iExpertID

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE usp_ExpCvvOriginalSelect
@@iExpertID int, @@iExpertOriginalID int output
AS
SET NOCOUNT ON
IF EXISTS(SELECT id_ExpertOriginal FROM tbl_Experts WHERE id_Expert=@@iExpertID AND id_ExpertOriginal>0 AND (Blacklist=1 OR expDeleted=1))
	SELECT @@iExpertOriginalID=id_ExpertOriginal FROM tbl_Experts WHERE id_Expert=@@iExpertID AND id_ExpertOriginal>0 AND (Blacklist=1 OR expDeleted=1)
ELSE 
	SET @@iExpertOriginalID=0

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE usp_ExpCvvPsnInfoInsert
@@psnTitleID tinyint, @@psnFirstName nvarchar(255),  @@psnMiddleName nvarchar(255),  
@@psnLastName nvarchar(255), @@psnBirthDate smalldatetime, 
@@psnBirthPlace nvarchar(255), @@GenderID  tinyint, @@MaritalStatusID  tinyint, @@ExpertID int, @@sUserLanguage varchar(3),
@@ID int OUTPUT
AS
SET NOCOUNT ON
IF @@sUserLanguage='Eng'
	BEGIN
	INSERT INTO tbl_Persons (id_psnTitle, psnFirstNameEng, psnMiddleNameEng, psnLastNameEng, psnBirthDate, psnBirthPlaceEng, psnGender, id_MaritalStatus, id_Expert) 
	VALUES (@@psnTitleID, @@psnFirstName, @@psnMiddleName, @@psnLastName, @@psnBirthDate, @@psnBirthPlace, @@GenderID, @@MaritalStatusID, @@ExpertID)
	SELECT @@ID=@@IDENTITY FROM tbl_Persons
	END

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE usp_ExpCvvPsnInfoSelect
@@iExpertID int, @@sActiveLng varchar(3)
AS
/*
DECLARE @iPersonID int, @iTitleID smallint, @iGenderID smallint
DECLARE @sFirstNameEng nvarchar(255), @sFirstNameFra nvarchar(255), @sFirstNameSpa nvarchar(255), @sFirstName nvarchar(255)
DECLARE @sMiddleNameEng nvarchar(255), @sMiddleNameFra nvarchar(255), @sMiddleNameSpa nvarchar(255), @sMiddleName nvarchar(255)
DECLARE @sLastNameEng nvarchar(255), @sLastNameFra nvarchar(255), @sLastNameSpa nvarchar(255), @sLastName  nvarchar(255)
DECLARE @iExpertID int, @iShortterm tinyint, @iLongterm tinyint, @sMasterLng varchar(3), @sEmail nvarchar(50), @sPhone nvarchar(50)
DECLARE @sKeyQualificationsEng nvarchar, @sKeyQualificationsFra nvarchar, @sKeyQualificationsSpa nvarchar, @sKeyQualifications nvarchar
DECLARE @sCurrPositionEng nvarchar(255), @sCurrPositionFra nvarchar(255), @sCurrPositionSpa nvarchar(255), @sCurrPosition nvarchar(255)
SELECT @iExpertID=E.id_Expert, @sMasterLng=E.Lng, 
@sKeyQualificationsEng=E.expKeyQualificationsEng, @sKeyQualificationsFra=E.expKeyQualificationsFra, @sKeyQualificationsSpa=E.expKeyQualificationsSpa, @sCurrPositionEng=E.expCurrPositionEng, @sCurrPositionFra=E.expCurrPositionFra, @sCurrPositionSpa=E.expCurrPositionSpa, 
@iPersonID=P.id_Person, @iTitleID=P.id_psnTitle, @iGenderID=P.psnGender, P.psnBirthPlaceEng, P.psnBirthDate, P.id_MaritalStatus,
@sFirstNameEng=P.psnFirstNameEng, @sFirstNameFra=P.psnFirstNameFra, @sFirstNameSpa=P.psnFirstNameSpa, @sMiddleNameEng=P.psnMiddleNameEng, @sMiddleNameFra=P.psnMiddleNameFra, @sMiddleNameSpa=P.psnMiddleNameSpa, @sLastNameEng=P.psnLastNameEng, @sLastNameFra=P.psnLastNameFra, @sLastNameSpa=P.psnLastNameSpa
FROM tbl_Experts E INNER JOIN tbl_Persons P ON E.id_Expert = P.id_Expert
WHERE E.id_Expert=@@iExpertID
*/
SELECT E.id_Expert, E.Lng, E.expKeyQualificationsEng, E.expKeyQualificationsFra, E.expKeyQualificationsSpa, E.expCurrPositionEng, E.expCurrPositionFra, E.expCurrPositionSpa, 
P.id_Person, P.id_psnTitle, P.psnGender, P.psnBirthPlaceEng, P.psnBirthPlaceFra, P.psnBirthPlaceSpa, P.psnBirthDate, P.id_MaritalStatus,
P.psnFirstNameEng, P.psnFirstNameFra, P.psnFirstNameSpa, P.psnMiddleNameEng, P.psnMiddleNameFra, P.psnMiddleNameSpa, P.psnLastNameEng, P.psnLastNameFra, P.psnLastNameSpa
FROM tbl_Experts E RIGHT JOIN tbl_Persons P ON E.id_Expert = P.id_Expert
WHERE E.id_Expert=@@iExpertID

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE usp_ExpCvvPsnInfoUpdate
@@psnTitleID tinyint, @@psnFirstName nvarchar(255),  @@psnMiddleName nvarchar(255),  
@@psnLastName nvarchar(255), @@psnBirthDate smalldatetime, 
@@psnBirthPlace nvarchar(255), @@GenderID  tinyint, @@MaritalStatusID  tinyint, @@ExpertID int, @@sUserLanguage varchar(3)
AS
SET NOCOUNT ON
IF @@sUserLanguage='Eng'
	BEGIN
	UPDATE tbl_Persons 
	SET id_psnTitle=@@psnTitleID, psnFirstNameEng=@@psnFirstName, psnMiddleNameEng=@@psnMiddleName, psnLastNameEng=@@psnLastName, 
	psnBirthDate=@@psnBirthDate, psnBirthPlaceEng=@@psnBirthPlace, psnGender=@@GenderID, id_MaritalStatus=@@MaritalStatusID
	END

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE usp_ExpCvvTrainingInsert
@@iExpertID int, @@iEduType int, @@sEduOther nvarchar(255), @@sEduTitle nvarchar(255), @@sEduAchievements nvarchar(255), @@sEduStartDate varchar(16), @@sEduEndDate varchar(16), @@sUserLanguage varchar(3)
AS
SET NOCOUNT ON
IF @@sUserLanguage='Eng' 
	BEGIN
	INSERT INTO lnkExp_Edu (id_Expert, id_EduType, eduOtherEng, eduDescriptionEng, eduDiploma1Eng, eduStartDate, eduEndDate)
	VALUES (@@iExpertID, @@iEduType, @@sEduOther, @@sEduTitle, @@sEduAchievements, @@sEduStartDate, @@sEduEndDate)
	END

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE usp_ExpCvvTrainingUpdate
@@iExpertID int, @@iExpEduID int, @@iEduType int, @@sEduOther nvarchar(255), @@sEduTitle nvarchar(255), @@sEduAchievements nvarchar(255), @@sEduStartDate varchar(16), @@sEduEndDate varchar(16), @@sUserLanguage varchar(3)
AS
SET NOCOUNT ON
IF @@sUserLanguage='Eng' 
	BEGIN
	UPDATE lnkExp_Edu SET eduOtherEng=@@sEduOther, eduDescriptionEng=@@sEduTitle, eduDiploma1Eng=@@sEduAchievements, eduStartDate=@@sEduStartDate, eduEndDate=@@sEduEndDate
	WHERE id_Expert=@@iExpertID AND id_ExpEdu=@@iExpEduID AND id_EduType=@@iEduType
	END

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE usp_ExpertAccountEmailSentUpdate
@iExpertID int,
@iValue tinyint
AS
UPDATE tbl_Experts
SET expAccountEmailSent=@iValue
WHERE id_Expert=@iExpertID

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE usp_ExpertNationalityUpdate
@iExpertID int, 
@sNationalityList varchar(4000)
AS
SET NOCOUNT ON 

SET @sNationalityList=REPLACE(@sNationalityList, ' ', '')
-- 1. Transform the list into a table
DECLARE @t_country TABLE(id_Country varchar(4) COLLATE Latin1_General_CI_AI, GLOBAL_ORDER int IDENTITY(1,1))
INSERT INTO @t_country
SELECT VALUE 
FROM dbo.udf_AdmSelectFromList(@sNationalityList, ',')
WHERE VALUE IS NOT NULL
AND ISNUMERIC(VALUE)=1

-- 2. Delete the previous data
DELETE FROM lnk_Exp_Nationality
WHERE id_Expert=@iExpertID

-- 3. Insert the new data
INSERT INTO lnk_Exp_Nationality
(id_Expert, id_Nationality)
SELECT @iExpertID, id_Country
FROM @t_country
ORDER BY GLOBAL_ORDER



GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE usp_ExpertProfileFullUpdate
@iExpertID int, 
@iUserID int, 
@sUserLanguage varchar(3),
@iTitleID tinyint, 
@sFirstName nvarchar(255), 
@sMiddleName nvarchar(255),  
@sLastName nvarchar(255), 
@sBirthDate smalldatetime, 
@sBirthPlace nvarchar(255), 
@iGenderID  tinyint, 
@iMaritalStatusID tinyint, 
@sUserPhone nvarchar(255), 
@sUserEmail nvarchar(255), 
@sRegistrationNumber nvarchar(80),
@iProfessionalStatusID smallint, 
@sProfession nvarchar(255), 
@sCurrentPosition nvarchar(255), 
@sKeyQualifications ntext, 
@iSeniority tinyint, 
--@sMembership ntext, 
--@sPublications ntext, 
--@sReferences ntext,
--@sAvailability nvarchar(4000), 
--@bShortterm  tinyint, 
--@bLongterm  tinyint, 
@iResult int output,
@iNewExpertID int output,
@iNewUserID int output, 
@sUserLogin varchar(255) output,
@sUserPassword varchar(255) output
AS  
SET NOCOUNT ON 
BEGIN TRANSACTION
SET @iResult=1
-- 0. Try to get UserID from ExpertID
IF ISNULL(@iUserID, 0)=0
BEGIN
	SELECT TOP 1 @iUserID=ISNULL(id_User, 0)
	FROM tbl_Experts
	WHERE id_Expert=@iExpertID
END


-- 1. Generate user login details if UserID = 0
IF ISNULL(@iUserID, 0)=0
BEGIN
	SET @sUserLogin=@sUserEmail
	SET @sUserPassword=LOWER(LEFT(@sUserLogin, 2) + CAST(DATEPART(dd, GETDATE()) AS varchar(2)) + CAST(DATEPART(ss, GETDATE()) AS varchar(2)) + SUBSTRING(CONVERT(varchar(64), NEWID()), 28, 2))
	INSERT INTO tbl_Users (id_UserType, UserName, [Password], usrFastLoginID) 
	VALUES (101, @sUserLogin, @sUserPassword, REPLACE(CONVERT(varchar(64), NEWID()), '-', ''))  
	SELECT @iNewUserID=@@IDENTITY FROM tbl_Users  
END
ELSE
BEGIN
	SET @iNewUserID=@iUserID
	SELECT @sUserLogin=UserName, @sUserPassword=[Password]
	FROM tbl_Users
	WHERE id_User=@iUserID
END

-- 2. Insert expert profile
IF ISNULL(@iExpertID, 0)=0
BEGIN
	-- 2.1 tbl_Expers
	INSERT INTO tbl_Experts (
	id_User, 
	expProfYears, 
	expKeyQualificationsEng, 
	expCurrPositionEng, 
	expProfessionEng, 
	id_ProfessionalStatus,
	expRegNumber,
	--expMemberProfEng, 
	--expPublicationsEng, 
	--expReferencesEng,
	--expAvailabilityEng, 
	--expShortterm, 
	--expLongterm, 
	Phone, 
	Email, 
	Lng, 
	expHidden,
	expCreateDate,
	expLastUpdate
	) VALUES (
	@iNewUserID, 
	@iSeniority, 
	@sKeyQualifications, 
	@sCurrentPosition, 
	@sProfession, 
	@iProfessionalStatusID,
	@sRegistrationNumber,
	--@sMembership, 
	--@sPublications, 
	--@sReferences,
	--@sAvailability, 
	--@bShortterm, 
	--@bLongterm, 
	@sUserPhone, 
	@sUserEmail, 
	@sUserLanguage, 
	0,
	GETDATE(),
	GETDATE())
	SELECT @iNewExpertID=@@IDENTITY FROM tbl_Experts
	-- 2.2 tbl_Persons
	INSERT INTO tbl_Persons (
	id_psnTitle, 
	psnFirstNameEng, 
	psnMiddleNameEng, 
	psnLastNameEng, 
	psnBirthDate, 
	psnBirthPlaceEng, 
	psnGender, 
	id_MaritalStatus, 
	id_Expert
	) VALUES (
	@iTitleID, 
	@sFirstName, 
	@sMiddleName, 
	@sLastName, 
	@sBirthDate, 
	@sBirthPlace, 
	@iGenderID, 
	@iMaritalStatusID, 
	@iNewExpertID)
	--2.3 tbl_Exp_Address
	INSERT INTO tbl_Exp_Address (
	id_AddressType,
	id_Expert,
	adrPhone,
	adrEmail,
	adrCreated
	) VALUES (
	1,
	@iNewExpertID, 
	@sUserPhone, 
	@sUserEmail, 
	GETDATE())
END  
ELSE
-- 3. Update expert profile
BEGIN
	-- 2.1 tbl_Expers
	UPDATE tbl_Experts
	SET
	expProfYears=@iSeniority, 
	expKeyQualificationsEng=@sKeyQualifications, 
	expCurrPositionEng=@sCurrentPosition, 
	expProfessionEng=@sProfession, 
	id_ProfessionalStatus=@iProfessionalStatusID,
	expRegNumber=@sRegistrationNumber,
	--expMemberProfEng=@sMembership, 
	--expPublicationsEng=@sPublications, 
	--expReferencesEng=@sReferences,
	--expAvailabilityEng=@sAvailability, 
	--expShortterm=@bShortterm, 
	--expLongterm=@bLongterm, 
	Phone=@sUserPhone, 
	Email=@sUserEmail, 
	Lng=@sUserLanguage, 
	expLastUpdate=GETDATE()
	WHERE id_Expert=@iExpertID

	SET @iNewExpertID=@iExpertID

	-- 2.2 tbl_Persons
	UPDATE tbl_Persons
	SET
	id_psnTitle=@iTitleID, 
	psnFirstNameEng=@sFirstName, 
	psnMiddleNameEng=@sMiddleName, 
	psnLastNameEng=@sLastName, 
	psnBirthDate=@sBirthDate,
	psnBirthPlaceEng=@sBirthPlace, 
	psnGender=@iGenderID, 
	id_MaritalStatus=@iMaritalStatusID
	WHERE id_Expert=@iExpertID
	--2.3 tbl_Exp_Address
	UPDATE tbl_Exp_Address
	SET
	adrPhone=@sUserPhone,
	adrEmail=@sUserEmail
	WHERE id_Expert=@iExpertID
	AND id_AddressType=1
END  
SET @iResult=0
COMMIT TRANSACTION


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE usp_ExpertProfileShortUpdate
@iExpertID int, 
@iUserID int,
@sUserLanguage varchar(3),
@iTitleID tinyint, 
@sFirstName nvarchar(255), 
--@sMiddleName nvarchar(255),  
@sLastName nvarchar(255), 
@sBirthDate smalldatetime, 
--@sBirthPlace nvarchar(255), 
--@GenderID  tinyint, 
--@iMaritalStatusID  tinyint, 
@sUserPhone nvarchar(255), 
@sUserEmail nvarchar(255), 
--@sProfession nvarchar(255), 
--@sCurrentPosition nvarchar(255), 
--@sKeyQualifications ntext, 
--@iSeniority tinyint, 
--@sMembership ntext, 
--@sPublications ntext, 
--@sReferences ntext,
@sAvailability nvarchar(400), 
@bShortterm  tinyint, 
@bLongterm  tinyint, 
@iResult int output,
@iNewExpertID int output,
@iNewUserID int output, 
@sUserLogin varchar(255) output,
@sUserPassword varchar(255) output
AS  
SET NOCOUNT ON 
BEGIN TRANSACTION
SET @iResult=1
-- 0. Try to get UserID from ExpertID
IF ISNULL(@iUserID, 0)=0
BEGIN
	SELECT TOP 1 @iUserID=ISNULL(id_User, 0)
	FROM tbl_Experts
	WHERE id_Expert=@iExpertID
END


-- 1. Generate user login details if UserID = 0
IF ISNULL(@iUserID, 0)=0 -- OR ISNULL(@iExpertID, 0)=0
BEGIN
	SET @sUserLogin=@sUserEmail
	IF LEN(ISNULL(@sUserLogin, ''))<3
	BEGIN
		SET @sUserLogin=ISNULL(@sLastName, @sFirstName)
	END

	SET @sUserPassword=LOWER(LEFT(@sUserLogin, 2) + CAST(DATEPART(dd, GETDATE()) AS varchar(2)) + CAST(DATEPART(ss, GETDATE()) AS varchar(2)) + SUBSTRING(CONVERT(varchar(64), NEWID()), 28, 2))
	INSERT INTO tbl_Users (id_UserType, UserName, [Password], usrFastLoginID) 
	VALUES (101, @sUserLogin, @sUserPassword, REPLACE(CONVERT(varchar(64), NEWID()), '-', ''))  
	SELECT @iNewUserID=@@IDENTITY FROM tbl_Users  
END
ELSE
BEGIN
	SET @iNewUserID=@iUserID
	SELECT @sUserLogin=UserName, @sUserPassword=[Password]
	FROM tbl_Users
	WHERE id_User=@iUserID
END


-- 2. Insert expert profile
IF ISNULL(@iExpertID, 0)=0
BEGIN
	-- 2.1 tbl_Expers
	INSERT INTO tbl_Experts (
	id_User, 
	--expProfYears, 
	--expKeyQualificationsEng, 
	--expCurrPositionEng, 
	--expProfessionEng, 
	--expMemberProfEng, 
	--expPublicationsEng, 
	--expReferencesEng,
	expAvailabilityEng, 
	expShortterm, 
	expLongterm, 
	Phone, 
	Email, 
	Lng, 
	expHidden,
	expCreateDate,
	expLastUpdate
	) VALUES (
	@iNewUserID, 
	--@iSeniority, 
	--@sKeyQualifications, 
	--@sCurrentPosition, 
	--@sProfession, 
	--@sMembership, 
	--@sPublications, 
	--@sReferences,
	@sAvailability, 
	@bShortterm, 
	@bLongterm, 
	@sUserPhone, 
	@sUserEmail, 
	@sUserLanguage, 
	0,
	GETDATE(),
	GETDATE())
	SELECT @iNewExpertID=@@IDENTITY FROM tbl_Experts
	-- 2.2 tbl_Persons
	INSERT INTO tbl_Persons (
	id_psnTitle, 
	psnFirstNameEng, 
	--psnMiddleNameEng, 
	psnLastNameEng, 
	psnBirthDate, 
	--psnBirthPlaceEng, 
	--psnGender, 
	--id_MaritalStatus, 
	id_Expert
	) VALUES (
	@iTitleID, 
	@sFirstName, 
	--@sMiddleName, 
	@sLastName, 
	@sBirthDate, 
	--@sBirthPlace, 
	--@GenderID, 
	--@iMaritalStatusID, 
	@iNewExpertID)
	--2.3 tbl_Exp_Address
	INSERT INTO tbl_Exp_Address (
	id_AddressType,
	id_Expert,
	adrPhone,
	adrEmail,
	adrCreated
	) VALUES (
	1,
	@iNewExpertID, 
	@sUserPhone, 
	@sUserEmail, 
	GETDATE())
END  
ELSE
-- 3. Update expert profile
BEGIN
	-- 2.1 tbl_Expers
	UPDATE tbl_Experts
	SET
	--expProfYears=@iSeniority, 
	--expKeyQualificationsEng=@sKeyQualifications, 
	--expCurrPositionEng=@sCurrentPosition, 
	--expProfessionEng=@sProfession, 
	--expMemberProfEng=@sMembership, 
	--expPublicationsEng=@sPublications, 
	--expReferencesEng=@sReferences,
	expAvailabilityEng=@sAvailability, 
	expShortterm=@bShortterm, 
	expLongterm=@bLongterm, 
	Phone=@sUserPhone, 
	Email=@sUserEmail, 
	Lng=@sUserLanguage, 
	expLastUpdate=GETDATE()
	WHERE id_Expert=@iExpertID
	SET @iNewExpertID=NULL
	-- 2.2 tbl_Persons
	UPDATE tbl_Persons
	SET
	id_psnTitle=@iTitleID, 
	psnFirstNameEng=@sFirstName, 
	--psnMiddleNameEng=@sMiddleName, 
	psnLastNameEng=@sLastName, 
	psnBirthDate=@sBirthDate
	--psnBirthPlaceEng=@sBirthPlace, 
	--psnGender=@GenderID, 
	--id_MaritalStatus=@iMaritalStatusID
	WHERE id_Expert=@iExpertID
	--2.3 tbl_Exp_Address
	UPDATE tbl_Exp_Address
	SET
	adrPhone=@sUserPhone,
	adrEmail=@sUserEmail
	WHERE id_Expert=@iExpertID
	AND id_AddressType=1
END  
SET @iResult=0
COMMIT TRANSACTION


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE usp_ExpertProjectDelete
@iExpertID int,
@iProjectID int
AS

DELETE FROM lnkExp_Prj
WHERE id_Expert=@iExpertID
AND id_Project=@iProjectID


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE usp_ExpertProjectListSelect
@iExpertID int,
@sProjectStatusList varchar(100),	-- List of statuses splitted by commas
@sKeyword varchar(250),			-- Specific reference / Keywords 
@sOrderBy varchar(50)
AS
SELECT EP.*,
EPS.exsTitleEng exsTitle, 
EPS.exsTitleEng, 
EPS.exsTitleFra, 
EPS.exsTitleSpa 
FROM lnkExp_Prj EP
INNER JOIN dbo.udf_ProjectListSelect(@sProjectStatusList, @sKeyword, @sOrderBy) P ON EP.id_Project=P.id_Project
LEFT OUTER JOIN tbl_ExpertStatus EPS ON EP.id_ExpertStatus=EPS.id_ExpertStatus
WHERE EP.id_Expert=@iExpertID




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE usp_ExpertProjectSelect
@iExpertID int,
@iProjectID int
AS
SELECT EP.*,
EPS.exsTitleEng exsTitle,
EPS.exsTitleEng,
EPS.exsTitleFra,
EPS.exsTitleSpa
FROM lnkExp_Prj EP
LEFT OUTER JOIN tbl_ExpertStatus EPS ON EP.id_ExpertStatus=EPS.id_ExpertStatus
WHERE EP.id_Expert=@iExpertID
AND EP.id_Project=@iProjectID


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE usp_ExpertProjectUpdate
@iExpertID int,
@iProjectID int,
@iExpertStatus smallint,
@sProvidedCompany varchar(400),
@sProvidedPerson varchar(400),
@fFee numeric(9,2),
@sFeeCurrency char(3),
@sComments text
AS
IF EXISTS(SELECT id_Expert 
	FROM lnkExp_Prj
	WHERE id_Expert=@iExpertID
	AND id_Project=@iProjectID
	)
	BEGIN
	UPDATE lnkExp_Prj
	SET id_ExpertStatus=@iExpertStatus,
	epjProvidedCompany=@sProvidedCompany,
	epjProvidedPerson=@sProvidedPerson,
	epjFee=@fFee,
	epjFeeCurrency=@sFeeCurrency,
	epjComments=@sComments,
	epjModifyDate=GETDATE()
	WHERE id_Expert=@iExpertID
	AND id_Project=@iProjectID
	END
ELSE
	BEGIN
	INSERT INTO lnkExp_Prj (
	id_Expert,
	id_Project,
	id_ExpertStatus,
	epjProvidedCompany,
	epjProvidedPerson,
	epjFee,
	epjFeeCurrency,
	epjComments,
	epjCreateDate
	) VALUES (
	@iExpertID,
	@iProjectID,
	@iExpertStatus,
	@sProvidedCompany,
	@sProvidedPerson,
	@fFee,
	@sFeeCurrency,
	@sComments,
	GETDATE()
	)
	END


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE usp_ExpertStatusCVSelect
@iExpertID int
AS
BEGIN
SELECT ES.id_Status, 
S.stsNameEng,
S.stsNameFra,
S.stsNameSpa,
ES.estModifyDate
FROM lnkExp_StatusCV ES
INNER JOIN tbl_StatusCV S ON ES.id_Status=S.id_Status
WHERE ES.id_Expert=@iExpertID

END


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE usp_ExpertStatusCVUpdate
@iExpertID int,
@iStatusID int,
@sDateModified varchar(16)
AS
BEGIN
SET NOCOUNT ON

DECLARE @dDateModified smalldatetime
IF ISDATE(@sDateModified)=1
	SET @dDateModified=CONVERT(smalldatetime, @sDateModified)

IF EXISTS(SELECT ES.id_Status
	FROM lnkExp_StatusCV ES
	WHERE ES.id_Expert=@iExpertID
	)
	BEGIN
	UPDATE lnkExp_StatusCV
	SET id_Status=@iStatusID,
	estModifyDate=@dDateModified
	WHERE id_Expert=@iExpertID
	END
ELSE
	BEGIN
	INSERT INTO lnkExp_StatusCV (
	id_Expert,
	id_Status,
	estModifyDate
	) VALUES (
	@iExpertID,
	@iStatusID,
	@sDateModified
	)	
	END

SET NOCOUNT OFF
END


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE usp_ExpertStatusListSelect
@iExpertStatusID int = NULL,
@sExpertStatusName nvarchar(255) = NULL,
@sOrderBy varchar(80) = NULL
AS
SELECT 
ES.id_ExpertStatus,
ES.exsTitleEng exsTitle,
ES.exsTitleEng,
ES.exsTitleFra,
ES.exsTitleSpa
FROM tbl_ExpertStatus ES
WHERE
(@iExpertStatusID IS NULL OR ES.id_ExpertStatus = @iExpertStatusID)
AND
(@sExpertStatusName IS NULL OR ES.exsTitleEng LIKE @sExpertStatusName + '%' OR ES.exsTitleFra LIKE @sExpertStatusName + '%' OR ES.exsTitleSpa LIKE @sExpertStatusName + '%')
ORDER BY ES.id_ExpertStatus


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE usp_GetExpertProfDetails
@id_Expert int, 
@id_Member int, 
@Lng varchar(3), 
@bReserved int, 
@ActiveExpert tinyint OUTPUT, @strFullName nvarchar(255) OUTPUT, @strSelectedDate nvarchar(50) OUTPUT, @strNationality nvarchar(400) OUTPUT, @strEducations nvarchar(1000) OUTPUT, @strLanguages nvarchar(500) OUTPUT, @strRegions nvarchar(1000) OUTPUT, @strFAgencies nvarchar(500) OUTPUT, @strMSectors nvarchar(1000) OUTPUT, @strKeyQualifications nvarchar(2000) OUTPUT
AS
SET NOCOUNT ON
DECLARE @curLanguage nvarchar(50), @Level nvarchar(50), @intLevel int
DECLARE @curRegion nvarchar(500), @curFAgency nvarchar(300), @curMSector nvarchar(500) 
DECLARE @strPeriod nvarchar(50), @Period int
SET @ActiveExpert=Null

SELECT @ActiveExpert=Active, @strSelectedDate=DATENAME(day, DownloadDate)+'&nbsp;'+DATENAME(month, DownloadDate)+'&nbsp;'+DATENAME(year, DownloadDate) 
FROM lnkMmb_Exp_Select 
WHERE id_Member=@id_Member 
AND id_Expert=@id_Expert

------------------------------------------------------------------------------------- 
-- Full name of expert
-- Updated on 20/08/2003
-------------------------------------------------------------------------------------
SELECT TOP 1 @strFullName=CASE WHEN E.Lng='Spa' THEN ISNULL(PT.ptlNameSpa,'') + '&nbsp;'+ ISNULL(P.psnFirstNameSpa,'') + '&nbsp;' + ISNULL(P.psnLastNameSpa,'')
				WHEN E.Lng='Fra' THEN ISNULL(PT.ptlNameFra,'') + '&nbsp;'+ ISNULL(P.psnFirstNameFra,'') + '&nbsp;' + ISNULL(P.psnLastNameFra,'')
				ELSE ISNULL(PT.ptlNameEng,'') + '&nbsp;'+ ISNULL(P.psnFirstNameEng,'') + '&nbsp;' + ISNULL(P.psnLastNameEng,'') END,
@strKeyQualifications=CASE WHEN E.Lng='Spa' THEN SUBSTRING(expKeyQualificationsSpa,1,2000)
			WHEN E.Lng='Fra' THEN SUBSTRING(expKeyQualificationsFra,1,2000)
			ELSE SUBSTRING(expKeyQualificationsEng,1,2000) END
FROM tbl_Persons P LEFT OUTER JOIN tbl_PersonTitles PT ON P.id_psnTitle=PT.id_psnTitle
INNER JOIN tbl_Experts E ON P.id_Expert=E.id_Expert
WHERE P.id_Expert=@id_Expert

------------------------------------------------------------------------------------- 
-- Nationalities
-- Updated on 20/08/2003
-- Using nested query to select only distinct countries if expert has registered the same nationality several times
-------------------------------------------------------------------------------------
SET @strNationality=''
SELECT @strNationality=@strNationality+couNameEng+', '
FROM
(
SELECT DISTINCT C.couNameEng 
FROM lnk_Exp_Nationality EN INNER JOIN tbl_Country C ON EN.id_Nationality=C.id_Country 
WHERE EN.id_Expert=@id_Expert
) AS ENL ORDER BY couNameEng

IF LEN(@strNationality)>1
	SET @strNationality=LTRIM(LEFT(@strNationality, LEN(@strNationality)-1))

------------------------------------------------------------------------------------- 
-- Education
-- Updated
-------------------------------------------------------------------------------------
SET @strEducations=''

SELECT @strEducations=@strEducations + CASE WHEN ESUBJ.edsDescriptionEng='Other' THEN EE.id_EduSubject1Eng ELSE ESUBJ.edsDescriptionEng END + 
		' (' + ETYPE.edtDescriptionEng + '), ' + CONVERT(varchar, DATEPART(YEAR,EE.eduEndDate)) + '<br> '
FROM lnkExp_Edu EE INNER JOIN tbl_EduSubjects ESUBJ ON EE.id_EduSubject = ESUBJ.id_EduSubject 
INNER JOIN tbl_EducationType ETYPE ON EE.eduDiploma = ETYPE.id_EduType 
WHERE EE.id_Expert=@id_Expert ORDER BY EE.eduEndDate DESC

------------------------------------------------------------------------------------- 
-- Languages
-- Updated
-------------------------------------------------------------------------------------
-- Native
SET @strLanguages=''

SELECT @strLanguages=@strLanguages+lngNameEng+'&nbsp;(Native), '
FROM
(
SELECT DISTINCT L.lngNameEng FROM
tbl_Languages L INNER JOIN tbl_Native_Lng NL on L.id_Language=NL.id_Language 
WHERE NL.id_Expert=@id_Expert
) AS T1


-- Other languages
SELECT @strLanguages=@strLanguages + lngNameEng + '&nbsp;(' + lnlDescriptionEng + '), ' FROM
(
SELECT DISTINCT L.lngNameEng, LL.lnlDescriptionEng, EL.exlAverage FROM
tbl_Languages L INNER JOIN lnkExp_Lan EL on L.id_Language=EL.id_Language 
INNER JOIN tbl_LangLevel LL ON EL.exlAverage=LL.id_LangLevel
WHERE EL.id_Expert=@id_Expert 
AND L.id_Language NOT IN (SELECT id_Language FROM tbl_Native_Lng WHERE id_Expert=@id_Expert AND id_Language IS NOT NULL)
) AS T1
ORDER BY exlAverage

IF LEN(@strLanguages)>1
	SET @strLanguages=LTRIM(LEFT(@strLanguages, LEN(@strLanguages)-1))

------------------------------------------------------------------------------------- 
-- Regions of working experience
-------------------------------------------------------------------------------------
/*
DECLARE @strRegions varchar(400)
SET @strRegions=''

SELECT @strRegions=@strRegions + Geo_ZoneEng + ' ' + CONVERT(varchar, wkePeriodByReg) + ', '
FROM udf_ExpCvvExperienceRegSelect(2715, 0)
--ORDER BY wkePeriodByReg DESC
PRINT @strRegions
*/

DECLARE cursorWkeRegions CURSOR FAST_FORWARD FOR 
SELECT tbl_GeoZone.Geo_ZoneEng, MAX(tbl_Res.Period) as PeriodGeoZone 
FROM tbl_GeoZone INNER join tbl_Country on tbl_GeoZone.id_GeoZone=tbl_Country.id_GeoZone inner join
(SELECT DISTINCT lnkWke_Cou.id_Country, SUM(wkePeriod) AS Period FROM lnkExp_Wke INNER JOIN 
lnkWke_Cou ON lnkExp_Wke.id_ExpWke = lnkWke_Cou.id_ExpWke 
where lnkExp_Wke.id_Expert=@id_Expert GROUP BY lnkWke_Cou.id_Country) As tbl_Res 
ON tbl_Country.id_Country=tbl_Res.id_Country GROUP BY tbl_GeoZone.Geo_ZoneEng ORDER BY PeriodGeoZone DESC
OPEN cursorWkeRegions
FETCH NEXT FROM cursorWkeRegions INTO @curRegion, @Period
SET @strRegions=''
WHILE @@FETCH_STATUS = 0
BEGIN
	IF @Period>240
	SET @strPeriod='&nbsp;(<b>over&nbsp;20&nbsp;years</b>)'
	ELSE
	IF @Period>=24
	SET @strPeriod='&nbsp;(<b>'+LTRIM(STR(ROUND(@Period/12,0))) + '&nbsp;years</b>)'
	ELSE
	IF @Period=1
	SET @strPeriod='&nbsp;(<b>1&nbsp;month</b>)'
	ELSE
	IF @Period>0
	SET @strPeriod='&nbsp;(<b>'+LTRIM(STR(ROUND(@Period,0))) + '&nbsp;months</b>)'
	ELSE
	SET @strPeriod=''
	
	SET @strRegions = @strRegions + ' ' + REPLACE(@curRegion,' ','&nbsp;') + @strPeriod + ', '
	FETCH NEXT FROM cursorWkeRegions INTO @curRegion, @Period
END
IF LEN(@strRegions)>1
SET @strRegions=LTRIM(LEFT(@strRegions, LEN(@strRegions)-1))
CLOSE cursorWkeRegions
DEALLOCATE cursorWkeRegions

------------------------------------------------------------------------------------- 
-- Major sectors of working experience
-------------------------------------------------------------------------------------
DECLARE cursorWkeMSectors CURSOR FAST_FORWARD FOR 
SELECT tbl_MainSectors.mnsDescriptionEng, MAX(tbl_Res.Period) as PeriodMSector 
FROM tbl_MainSectors INNER join tbl_Sectors on tbl_MainSectors.id_MainSector=tbl_Sectors.id_MainSector inner join
(SELECT DISTINCT lnkWke_Sct.id_Sector, SUM(wkePeriod) AS Period FROM lnkExp_Wke INNER JOIN 
lnkWke_Sct ON lnkExp_Wke.id_ExpWke = lnkWke_Sct.id_ExpWke 
where lnkExp_Wke.id_Expert=@id_Expert GROUP BY lnkWke_Sct.id_Sector) As tbl_Res 
ON tbl_Sectors.id_Sector=tbl_Res.id_Sector GROUP BY tbl_MainSectors.mnsDescriptionEng ORDER BY PeriodMSector DESC
OPEN cursorWkeMSectors
FETCH NEXT FROM cursorWkeMSectors INTO @curMSector, @Period
SET @strMSectors=''
WHILE @@FETCH_STATUS = 0
BEGIN
	IF @Period>240
	SET @strPeriod='&nbsp;(<b>over&nbsp;20&nbsp;years</b>)'
	ELSE
	IF @Period>=24
	SET @strPeriod='&nbsp;(<b>'+LTRIM(STR(ROUND(@Period/12,0))) + '&nbsp;years</b>)'
	ELSE
	IF @Period=1
	SET @strPeriod='&nbsp;(<b>1&nbsp;month</b>)'
	ELSE
	IF @Period>0
	SET @strPeriod='&nbsp;(<b>'+LTRIM(STR(ROUND(@Period,0))) + '&nbsp;months</b>)'
	ELSE
	SET @strPeriod=''
	
	SET @strMSectors = @strMSectors + ' ' + REPLACE(@curMSector,' ','&nbsp;') + @strPeriod + ', '
	FETCH NEXT FROM cursorWkeMSectors INTO @curMSector, @Period
END
IF LEN(@strMSectors)>1
SET @strMSectors=LTRIM(LEFT(@strMSectors, LEN(@strMSectors)-1))
CLOSE cursorWkeMSectors
DEALLOCATE cursorWkeMSectors


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE usp_LogErrorAdd
@@uSessionID uniqueidentifier, @@iErrorNumber int, @@sErrorCategory nvarchar(255), @@sErrorDescription nvarchar(1000), @@sErrorUrl nvarchar(255), @@sErrLocation nvarchar(100), @@sErrorMethod varchar(50), @@sErrorPostData ntext=Null,
@@iErrorID int output
AS
DECLARE @iUserSessionID bigint
SET NOCOUNT ON
SELECT @iUserSessionID=id_UserSession FROM log_Session WHERE id_Session=@@uSessionID
IF @iUserSessionID>0 
INSERT INTO log_Error (id_UserSession, errNumber, errCategory, errDescription, errUrl, errLocation, errMethod, errPostData)
	VALUES (@iUserSessionID, @@iErrorNumber, @@sErrorCategory, @@sErrorDescription, @@sErrorUrl, @@sErrLocation, @@sErrorMethod, @@sErrorPostData)
SELECT @@iErrorID=@@IDENTITY FROM log_Error

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE usp_LogSessionCreate
@iUserID int, @sUserAgent nvarchar(255), @sIpAddress varchar(16), @iUrlNoCookiesTTL int,
@uSessionID uniqueidentifier output, @bAddUrlSession tinyint output
AS
SET NOCOUNT ON
SET @bAddUrlSession=0

-- 1. TRY TO FIND OPEN SESSION. 
-- USED WHEN COOKIES ARE DISABLED
SELECT @uSessionID=S.id_Session
FROM log_Session S 
INNER JOIN log_SessionEvent SE ON S.id_UserSession=SE.id_UserSession
WHERE ussIpAddress=@sIpAddress
AND ussUserAgent=@sUserAgent
GROUP BY S.id_Session
HAVING DATEDIFF(n, MAX(S.ussCreateDate), GETDATE())<@iUrlNoCookiesTTL


-- 2. GENERATE NEW SESSION IF NEEDED
IF @uSessionID IS NULL
	BEGIN
	SET @uSessionID=NEWID()

	INSERT INTO log_Session (id_Session, id_User, ussUserAgent, ussIpAddress)
	VALUES (@uSessionID, @iUserID, @sUserAgent, @sIpAddress)
	END
ELSE
	BEGIN
	SET @bAddUrlSession=1
	END


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE usp_LogSessionEvent
@@uSessionID uniqueidentifier, @@sUrlRequested nvarchar(255)
AS
DECLARE @iUserSessionID bigint
SET NOCOUNT ON
SELECT @iUserSessionID=id_UserSession FROM log_Session WHERE id_Session=@@uSessionID
IF @iUserSessionID>0 
INSERT INTO log_SessionEvent (id_UserSession, slgUrl) VALUES (@iUserSessionID, @@sUrlRequested)

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE usp_LogSessionUser
@@uSessionID uniqueidentifier, @@iUserID int
AS
SET NOCOUNT ON
UPDATE log_Session SET id_User=@@iUserID WHERE id_Session=@@uSessionID

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE usp_LogSessionUserDataSelect
@iUserID int
AS
SELECT ISNULL(M.id_Member, 0) id_Member,
ISNULL(E.id_Expert, 0) id_Expert, 
U.UserName, 
ISNULL(E.Lng, M.Lng) Lng,
ISNULL(E.Email, M.Email) Email,
UT.ustName UserType,
UT.ustDescription UserTypeDescription
FROM tbl_Users U
LEFT OUTER JOIN tbl_UserType UT ON U.id_UserType=UT.id_UserType
LEFT OUTER JOIN tbl_Members M ON U.id_User=M.id_User
LEFT OUTER JOIN tbl_Experts E ON U.id_User=E.id_User
WHERE U.id_User=@iUserID


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE usp_LogSessionValidate
@uSessionID uniqueidentifier, @sUserAgent nvarchar(255), @sIpAddress varchar(16), @iOpenSessionTTL int,
@bValidSession tinyint output, @iUserID int output
AS
SET NOCOUNT ON

DECLARE @sIpBase varchar(16)

-- Checking for session only by first two octets of IpAddress
SET @sIpBase=SUBSTRING(@sIpAddress, 1, CHARINDEX('.', @sIpAddress, 1))

DECLARE @dLastSessionActivity smalldatetime
SET @dLastSessionActivity='20010101'
--SET @iUserID=0

SELECT @iUserID=S.id_User, @dLastSessionActivity=MAX(SE.slgDate) 
FROM log_Session S 
INNER JOIN log_SessionEvent SE ON S.id_UserSession=SE.id_UserSession
WHERE S.id_Session=@uSessionID 
AND ussUserAgent=@sUserAgent
GROUP BY S.id_User

IF DATEDIFF(n, @dLastSessionActivity, GETDATE())>@iOpenSessionTTL
	SET @bValidSession=0
ELSE
	SET @bValidSession=1



GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO


CREATE PROCEDURE usp_MmbExpDownloadedCleanup
@iMemberID int,
@sExpertIds varchar(4000),
@sResultExpertIds varchar(4000) output
AS
SET @sResultExpertIds=@sExpertIds

SELECT @sResultExpertIds=REPLACE(@sResultExpertIds, ',' + CAST(id_Expert AS varchar(16)), '')
FROM lnkMmb_Exp_Select
WHERE id_Member=@iMemberID
AND Active=1




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE usp_MmbExpSearchFirstSelect
@iMemberID int, 
@sSearchExperts varchar(150), 
@sSearchFirstname nvarchar(150), 
@sSearchSurname nvarchar(150), 
@sSearchKeywords nvarchar(1000), 
@sSearchNationality varchar(1400), 
@sSearchEduSubject varchar(350), 
@sSearchNativeLng varchar(700), 
@sSearchOtherLng varchar(700), 
@sSearchSeniority varchar(50), 
@sSearchCountries varchar(1000), 
@sSearchRegions varchar(150), 
@sSearchSectors varchar(1600), 
@sSearchMainSectors varchar (150), 
@sSearchDonors varchar(200), 
@sSearchDB varchar(10), 
@iSearchCurrentlyIn int, 
@iSearchPastYears int,
@iSearchPastProjects int,
@sCvLanguage varchar(3), 
@sCvType varchar(150), 
@bShowHiddenExperts bit, 
@bShowRemovedExperts bit, 
@bSaveSearchLog tinyint,
@iSearchQueryID int OUTPUT
AS
DECLARE @sSQLQuery nvarchar(4000), @sSQLRankQuery nvarchar(100)
DECLARE @iExpertID int
SET @iExpertID=0
-- If search done by expert id
IF @sSearchSurname>'' AND ISNUMERIC(@sSearchSurname)=1 AND CHARINDEX('.', @sSearchSurname)=0 AND CHARINDEX(',', @sSearchSurname)=0
	SET @iExpertID=CONVERT(int, @sSearchSurname)

EXEC usp_SrvRemoveNonTextSymbols @sSearchKeywords, @sSearchKeywords OUTPUT
SET @sSearchKeywords=REPLACE(@sSearchKeywords,'  ',' ')
SET @sSearchKeywords=REPLACE(@sSearchKeywords,'" " AND ','')
SET @sSearchKeywords=REPLACE(@sSearchKeywords,'AND " "','')

-- Search for Serbia or Montenegro should include the old country "Serbia and Montenergro"
SET @sSearchCountries=REPLACE(@sSearchCountries, '711', '711, 524')
SET @sSearchCountries=REPLACE(@sSearchCountries, '712', '712, 524')


-- Writing queries log
SET NOCOUNT ON
IF @bSaveSearchLog=1 
	BEGIN
	INSERT INTO log_MmbExpSearch (id_Member, srchKeywords, srchNationality, srchEducation, srchCountries, srchRegions, srchSectors, srchMainSectors, srchDonors, srchDB, srchSQLQuery) VALUES (@iMemberID, LEFT(ISNULL(@sSearchKeywords, ''), 400), SUBSTRING(@sSearchNationality,1,250), @sSearchEduSubject, @sSearchCountries, @sSearchRegions, @sSearchSectors, @sSearchMainSectors, @sSearchDonors, @sSearchDB, @sSQLQuery)
	SELECT @iSearchQueryID = @@IDENTITY FROM log_MmbExpSearch
	END

-- For EU nationalities
SET @sSearchNationality=REPLACE(@sSearchNationality, '1100', '548, 550, 554, 565, 563, 555, 571, 570, 558, 551, 574, 560, 569, 564, 559, 515, 517, 520, 521, 522, 525, 529, 530, 556, 568, 516, 519')
-- TACIS
SET @sSearchNationality=REPLACE(@sSearchNationality, '1112', '543, 513, 544, 546, 508, 510, 541, 578, 545, 514, 512, 542, 511')
-- CARDS 
SET @sSearchNationality=REPLACE(@sSearchNationality, '1113', '526, 528, 518, 523, 524')
-- MEDA
SET @sSearchNationality=REPLACE(@sSearchNationality, '1114', '601, 556, 597, 596, 591, 594, 568, 598, 583, 602, 584, 704')
-- ALA
SET @sSearchNationality=REPLACE(@sSearchNationality, '1115', '509, 654, 665, 671, 661, 655, 685')


-- Building SQL query
SET @sSQLQuery=''
SET @sSQLQuery=@sSQLQuery+'SELECT DISTINCT E.id_Expert as id_Expert, E.expProfessionEng AS Profession, E.expProfYears AS Seniority, ISNULL(E.expProfessionEng, ''zzz'') AS ProfessionOrder, (100-ISNULL(E.expProfYears, 0)) As SeniorityOrder, E.Lng, expAvailabilityEng, expAvailabilityFra, expAvailabilitySpa, expShortterm, expLongterm, psnLastNameEng '
IF (@sSearchMainSectors>'') Or (@sSearchSectors>'')
	BEGIN
	SET @sSQLQuery=@sSQLQuery + ', MAX(lnkExp_RankSct.rnkSctValue) AS SctRank '
	SET @sSQLRankQuery=', (14*MAX(lnkExp_RankSct.rnkSctValue)'
	END
ELSE
	BEGIN
	SET @sSQLQuery=@sSQLQuery + ', 1 as SctRank '
	SET @sSQLRankQuery=', (1'
	END

IF (@sSearchRegions>'') Or (@sSearchCountries>'')
	BEGIN
	SET @sSQLQuery=@sSQLQuery + ', MAX(lnkExp_RankCou.rnkCouValue) AS CouRank '
	SET @sSQLRankQuery=@sSQLRankQuery + '+MAX(lnkExp_RankCou.rnkCouValue)) AS FieldsRank '
	END
ELSE
	BEGIN
	SET @sSQLQuery=@sSQLQuery + ', 1 as CouRank '
	SET @sSQLRankQuery=@sSQLRankQuery + '+1) AS FieldsRank '
	END
SET @sSQLQuery=@sSQLQuery + @sSQLRankQuery
IF (@sSearchDB)>''
	SET @sSQLQuery=@sSQLQuery + ', lnkMmb_Exp_Select.DownloadDate '
IF @sSearchKeywords>''
	BEGIN
	SET @sSQLQuery=@sSQLQuery + ', ISNULL(E_KWD.[RANK],0) + ISNULL(EE_KWD.[RANK],0) + ISNULL(EW_KWD.[RANK],0) AS TextRank '
	END

SET @sSQLQuery=@sSQLQuery + ' FROM tbl_Experts E '
-- FILTER ON EXPERT STATUS
SET @sSQLQuery=@sSQLQuery + ' LEFT OUTER JOIN lnkExp_StatusCV E2 ON E.id_Expert=E2.id_Expert '

IF (@sSearchMainSectors>'') Or (@sSearchSectors>'') Or (@sSearchRegions>'') Or (@sSearchCountries>'') Or (@sSearchDonors>'') Or (@sSearchKeywords>'') Or (ISNULL(@iSearchCurrentlyIn, 0)>0) Or (ISNULL(@iSearchPastYears, 0)>0) Or (ISNULL(@iSearchPastProjects, 0)>0)
	SET @sSQLQuery=@sSQLQuery + ' INNER JOIN lnkExp_Wke EW ON E.id_Expert=EW.id_Expert '
IF @sSearchKeywords>''
	BEGIN
	SET @sSQLQuery=@sSQLQuery + ' LEFT JOIN CONTAINSTABLE(tbl_Experts, *, ''' + @sSearchKeywords + ''') AS E_KWD ON E.id_Expert=E_KWD.[KEY] '
	SET @sSQLQuery=@sSQLQuery + ' LEFT JOIN (SELECT DISTINCT EW.id_Expert, SUM([RANK]) AS Rank 
						FROM lnkExp_Wke EW LEFT JOIN CONTAINSTABLE(lnkExp_Wke, *, ''' + @sSearchKeywords + ''') AS EW_KWDTMP ON EW.id_ExpWke=EW_KWDTMP.[KEY] 
						WHERE Rank IS NOT NULL GROUP BY EW.id_Expert) AS EW_KWD ON E.id_Expert=EW_KWD.id_Expert '
	SET @sSQLQuery=@sSQLQuery + ' LEFT JOIN (SELECT DISTINCT EE.id_Expert, SUM([RANK]) AS Rank 
						FROM lnkExp_Edu EE LEFT JOIN CONTAINSTABLE(lnkExp_Edu, *, ''' + @sSearchKeywords + ''') AS EE_KWDTMP ON EE.id_ExpEdu=EE_KWDTMP.[KEY] 
						WHERE Rank IS NOT NULL GROUP BY EE.id_Expert) AS EE_KWD ON E.id_Expert=EE_KWD.id_Expert '
	END

IF (@sSearchMainSectors>'') Or (@sSearchSectors>'')
BEGIN
	SET @sSQLQuery=@sSQLQuery + ' INNER JOIN lnkWke_Sct ON EW.id_ExpWke=lnkWke_Sct.id_ExpWke 
				LEFT JOIN lnkExp_RankSct ON EW.id_Expert=lnkExp_RankSct.id_Expert AND lnkWke_Sct.id_Sector=lnkExp_RankSct.id_Sector '
	IF (@sSearchMainSectors>'') 
		SET @sSQLQuery=@sSQLQuery + ' INNER JOIN tbl_Sectors ON lnkWke_Sct.id_Sector=tbl_Sectors.id_Sector '
END
IF (@sSearchRegions>'') Or (@sSearchCountries>'') Or (@iSearchCurrentlyIn>0)
BEGIN
	SET @sSQLQuery=@sSQLQuery + ' INNER JOIN lnkWke_Cou ON EW.id_ExpWke=lnkWke_Cou.id_ExpWke INNER JOIN tbl_Country ON lnkWke_Cou.id_Country=tbl_Country.id_Country 
				LEFT JOIN lnkExp_RankCou ON EW.id_Expert=lnkExp_RankCou.id_Expert AND lnkWke_Cou.id_Country=lnkExp_RankCou.id_Country '
	IF (@sSearchRegions>'') 
		SET @sSQLQuery=@sSQLQuery + ' INNER JOIN tbl_Country ON lnkWke_Cou.id_Country=tbl_Country.id_Country '
END
IF (@iSearchCurrentlyIn>0)
	BEGIN
	SET @sSQLQuery=@sSQLQuery + ' INNER JOIN (SELECT id_Expert, MAX(wkeEndDate) AS wkeLastDate FROM lnkExp_Wke GROUP BY id_Expert) T1 ON EW.id_Expert=T1.id_Expert AND EW.wkeEndDate=T1.wkeLastDate '
	SET @sSQLQuery=@sSQLQuery + ' INNER JOIN tbl_Exp_Address EADR ON E.id_Expert=EADR.id_Expert '
	END

IF @sSearchDonors>''
	SET @sSQLQuery=@sSQLQuery + ' INNER JOIN lnkWke_Don ON EW.id_ExpWke=lnkWke_Don.id_ExpWke '
IF @sSearchNationality>'' 
	SET @sSQLQuery=@sSQLQuery + ' INNER JOIN lnk_Exp_Nationality ON E.id_Expert=lnk_Exp_Nationality.id_Expert '
--IF @sSearchNativeLng>''
--	SET @sSQLQuery=@sSQLQuery + ' INNER JOIN tbl_Native_Lng ON E.id_Expert=tbl_Native_Lng.id_Expert '
--IF @sSearchOtherLng>''
--	SET @sSQLQuery=@sSQLQuery + ' INNER JOIN lnkExp_Lan ON E.id_Expert=lnkExp_Lan.id_Expert '

IF @sSearchNativeLng>''
	SET @sSQLQuery=@sSQLQuery + ' INNER JOIN (SELECT id_Language, id_Expert FROM tbl_Native_Lng WHERE id_Language in (' + @sSearchNativeLng + ') UNION SELECT id_Language, id_Expert FROM lnkExp_Lan WHERE id_Language in (' + @sSearchNativeLng + ')) AS lnkExp_Lng ON E.id_Expert=lnkExp_Lng.id_Expert '
	--SET @sSQLQuery=@sSQLQuery + ' INNER JOIN (SELECT id_Expert FROM tbl_Native_Lng WHERE id_Language in (' + @sSearchNativeLng + ') UNION SELECT id_Expert FROM lnkExp_Lan WHERE id_Language in (' + @sSearchNativeLng + ') AND exlLevel<9) AS lnkExp_Lng ON E.id_Expert=lnkExp_Lng.id_Expert '

IF @sSearchEduSubject>'' 
	SET @sSQLQuery=@sSQLQuery + ' INNER JOIN lnkExp_Edu ON E.id_Expert=lnkExp_Edu.id_Expert '
IF @sSearchDB>''
	SET @sSQLQuery=@sSQLQuery + ' INNER JOIN lnkMmb_Exp_Select ON lnkMmb_Exp_Select.id_Expert=E.id_Expert '
--IF @sSearchFirstname>'' OR @sSearchSurname>''
	SET @sSQLQuery=@sSQLQuery + ' INNER JOIN tbl_Persons ON E.id_Expert=tbl_Persons.id_Expert '

SET @sSQLQuery=@sSQLQuery + ' WHERE id_ExpertOriginal=0 AND expDeleted=0 '

-- FILTER ON EXPERT STATUS
SET @sSQLQuery=@sSQLQuery + ' AND (E2.id_Status IS NULL OR E2.id_Status NOT IN (19, 90)) '

-- IF NO SEARCH CRITERIA DEFINED - SHOW NO RESULTS
IF LEN(ISNULL(@sSearchExperts, ''))<2 
	AND LEN(ISNULL(@sSearchFirstname, ''))<2 
	AND LEN(ISNULL(@sSearchSurname, ''))<2 
	AND LEN(ISNULL(@sSearchKeywords, ''))<2 
	AND LEN(ISNULL(@sSearchNationality, ''))<2 
	AND LEN(ISNULL(@sSearchEduSubject, ''))<2 
	AND LEN(ISNULL(@sSearchNativeLng, ''))<2 
	AND LEN(ISNULL(@sSearchOtherLng, ''))<2 
	AND LEN(ISNULL(@sSearchSeniority, ''))<2 
	AND LEN(ISNULL(@sSearchCountries, ''))<2 
	AND LEN(ISNULL(@sSearchRegions, ''))<2 
	AND LEN(ISNULL(@sSearchSectors, ''))<2 
	AND LEN(ISNULL(@sSearchMainSectors, ''))<2 
	AND LEN(ISNULL(@sSearchDonors, ''))<2 
	AND ISNULL(@iSearchCurrentlyIn, 0)<1
	AND ISNULL(@iSearchPastYears, 0)<1
	AND ISNULL(@iSearchPastProjects, 0)<1
	AND LEN(ISNULL(@sCvLanguage, ''))<2 
	AND LEN(ISNULL(@sCvType, ''))<2 
	BEGIN
	SET @sSQLQuery=@sSQLQuery + ' AND 1=0 '
	END

IF @bShowHiddenExperts=0
	SET @sSQLQuery=@sSQLQuery + ' AND E.expHidden=0 '
IF @bShowRemovedExperts=0
	SET @sSQLQuery=@sSQLQuery + ' AND E.expRemoved=0 '
IF (@iSearchCurrentlyIn>0)
	BEGIN
	SET @sSQLQuery=@sSQLQuery + ' AND (DATEDIFF(m, ISNULL(wkeEndDate, DATEADD(y, 1, wkeStartDate)), GETDATE())<=6 OR EADR.id_Country in (' + CONVERT(varchar(4), @iSearchCurrentlyIn) + ')) '
	SET @sSQLQuery=@sSQLQuery + ' AND lnkWke_Cou.id_Country in (' + CONVERT(varchar(4), @iSearchCurrentlyIn) + ')'
	END
IF (ISNULL(@iSearchPastYears, 0)>0)
	BEGIN
	SET @sSQLQuery=@sSQLQuery + ' AND (DATEDIFF(yy, ISNULL(wkeEndDate, DATEADD(y, 1, wkeStartDate)), GETDATE())<= ' + CAST(@iSearchPastYears AS varchar(2)) + ' ) '
	END
IF @iExpertID>0
	BEGIN
		SET @sSQLQuery=@sSQLQuery + ' AND E.id_Expert=' + @sSearchSurname + ' '
	END
ELSE
	BEGIN
	IF @sSearchFirstname>''
		SET @sSQLQuery=@sSQLQuery + ' AND (tbl_Persons.psnFirstNameEng=''' + @sSearchFirstname + ''' OR tbl_Persons.psnFirstNameFra=''' + @sSearchFirstname + ''' OR tbl_Persons.psnFirstNameSpa=''' + @sSearchFirstname + ''') '
	IF @sSearchSurname>''
		SET @sSQLQuery=@sSQLQuery + ' AND (tbl_Persons.psnLastNameEng LIKE ''%' + @sSearchSurname + '%'') '
	END

IF @sSearchExperts>''
	SET @sSQLQuery=@sSQLQuery + ' AND E.id_Expert in ('  + @sSearchExperts + ') '
IF @sSearchSeniority>''
	SET @sSQLQuery=@sSQLQuery + ' AND (E.expProfYears between ' + @sSearchSeniority + ')'
IF @sSearchDB>''
	SET @sSQLQuery=@sSQLQuery + ' AND lnkMmb_Exp_Select.Active=1 AND lnkMmb_Exp_Select.id_Member=' +  CONVERT(varchar, @iMemberID)
IF @sSearchNationality>''
	SET @sSQLQuery=@sSQLQuery + ' AND lnk_Exp_Nationality.id_Nationality in (' + @sSearchNationality + ')'
IF @sSearchEduSubject>'' 
	SET @sSQLQuery=@sSQLQuery + ' AND lnkExp_Edu.id_EduSubject in (' + @sSearchEduSubject + ')'

--IF @sSearchNativeLng>''
--	BEGIN
--	SET @sSQLQuery=@sSQLQuery + ' AND lnkExp_Lng.id_Language in (' + @sSearchNativeLng + ')'
--	END

IF @sSearchSectors>''
	BEGIN
--	SET @sSQLQuery=@sSQLQuery + ' AND tbl_Sectors.id_Sector in (' + @sSearchSectors + ')'
	SET @sSQLQuery=@sSQLQuery + ' AND lnkWke_Sct.id_Sector in (' + @sSearchSectors + ') ' --AND lnkExp_RankSct.rnkSctValue>=1
	END
IF @sSearchMainSectors>''
	SET @sSQLQuery=@sSQLQuery + ' AND tbl_Sectors.id_MainSector in (' + @sSearchMainSectors + ')'

If @sSearchCountries>''
	BEGIN
--	SET @sSQLQuery=@sSQLQuery + ' AND tbl_Country.id_Country in (' + @sSearchCountries + ')'
	SET @sSQLQuery=@sSQLQuery + ' AND lnkWke_Cou.id_Country in (' + @sSearchCountries + ') ' -- AND lnkExp_RankCou.rnkCouValue>=1
	END
If @sSearchRegions>''
	SET @sSQLQuery=@sSQLQuery + ' AND tbl_Country.id_GeoZone in (' + @sSearchRegions + ')'
IF @sSearchDonors>''
	SET @sSQLQuery=@sSQLQuery + ' AND lnkWke_Don.id_Organisation in (' + @sSearchDonors + ')'
IF @sSearchKeywords>''
	SET @sSQLQuery=@sSQLQuery + ' AND (E_KWD.[RANK]>0 OR EE_KWD.[RANK]>0 OR EW_KWD.[RANK]>0)'

IF @sCvLanguage>''
	SET @sSQLQuery=@sSQLQuery + ' AND (E.Lng=''' + @sCvLanguage + ''')'
IF @sCvType>''
	SET @sSQLQuery=@sSQLQuery + ' AND (E.KgCVFile=''' + @sCvType + ''')'

SET @sSQLQuery=@sSQLQuery + ' GROUP BY E.id_Expert, E.expProfessionEng, E.expProfYears, E.Lng, expAvailabilityEng, expAvailabilityFra, expAvailabilitySpa, expShortterm, expLongterm, psnLastNameEng '
If @sSearchDB>''
	SET @sSQLQuery=@sSQLQuery + ', lnkMmb_Exp_Select.DownloadDate '
IF @sSearchKeywords>''
	BEGIN
	SET @sSQLQuery=@sSQLQuery + ', ISNULL(E_KWD.[RANK],0) + ISNULL(EE_KWD.[RANK],0) + ISNULL(EW_KWD.[RANK],0) '
	END

IF (ISNULL(@iSearchPastProjects, 0)>0)
	BEGIN
	SET @sSQLQuery=@sSQLQuery + ' HAVING COUNT(EW.id_Expert)> ' + CAST(@iSearchPastProjects AS varchar(2)) + ' '
	END

SET @sSQLQuery=@sSQLQuery + ' ORDER BY '
IF @sSearchDB>'' 
		SET @sSQLQuery=@sSQLQuery + ' lnkMmb_Exp_Select.DownloadDate DESC, '
IF @sSearchKeywords>''
	SET @sSQLQuery=@sSQLQuery + ' TextRank DESC, FieldsRank DESC '
ELSE
	SET @sSQLQuery=@sSQLQuery + ' FieldsRank DESC '


IF @bSaveSearchLog=1 
	UPDATE log_MmbExpSearch SET srchSQLQuery=@sSQLQuery WHERE id_SQLQuery=@iSearchQueryID

SET NOCOUNT OFF

--PRINT @sSQLQuery
--EXEC sp_executesql @sSQLQuery



GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO


CREATE PROCEDURE usp_MmbExpSearchRepeatSelect
@@iMemberID int, @@iSearchQueryID int, @@sKeywordsAddtitional nvarchar(255), @@sOrderBy varchar(50)
AS
SET NOCOUNT ON
DECLARE @sSQLQuery nvarchar(4000), @iFlag int
SELECT @sSQLQuery=srchSQLQuery FROM log_MmbExpSearch WHERE (id_Member=@@iMemberID or id_Member=0) AND id_SQLQuery=@@iSearchQueryID

IF LEN(@@sKeywordsAddtitional)>1 
	BEGIN
	SET @sSQLQuery=REPLACE(@sSQLQuery, 'GROUP BY E.id_Expert', 'AND CONTAINS (E.*, ''"' + @@sKeywordsAddtitional + '"'') GROUP BY E.id_Expert')
	END

IF @@sOrderBy='expRank'
	SET @sSQLQuery=@sSQLQuery

IF @@sOrderBy='expProfession'
	SET @sSQLQuery=REPLACE(@sSQLQuery, 'ORDER BY', 'ORDER BY ProfessionOrder, ')

IF @@sOrderBy='expSeniority'
	SET @sSQLQuery=REPLACE(@sSQLQuery, 'ORDER BY', 'ORDER BY SeniorityOrder, ')

IF @@sOrderBy='expLastName' 
	BEGIN
	SET @sSQLQuery=REPLACE(@sSQLQuery, 'ORDER BY', 'ORDER BY psnLastNameEng, ')
	END
SET NOCOUNT OFF

IF @sSQLQuery>''
	EXEC sp_executesql @sSQLQuery
ELSE
	SELECT null




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE usp_ProfessionalStatusListSelect
AS
SELECT id_ProfessionalStatus, pfsTitle
FROM tbl_ProfessionalStatus
ORDER BY id_ProfessionalStatus


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE usp_ProjectDelete
@iProjectID int
AS

DELETE FROM tbl_Project
WHERE id_Project=@iProjectID



GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE usp_ProjectExpertListSelect (
@iProjectID int,
@sAction varchar(100)=NULL, 
@sSearchString varchar(255)=NULL, 
@sOrderBy varchar(100)='S'
)
AS
SELECT E.*,
PE.*,
PES.exsTitleEng exsTitle,
PES.exsTitleEng,
PES.exsTitleFra,
PES.exsTitleSpa
FROM lnkExp_Prj PE
LEFT OUTER JOIN tbl_ExpertStatus PES ON PE.id_ExpertStatus=PES.id_ExpertStatus
INNER JOIN uvw_Experts E ON PE.id_Expert=E.id_Expert
WHERE PE.id_Project=@iProjectID
ORDER BY CASE WHEN @sOrderBy='S' THEN PE.id_ExpertStatus ELSE NULL END,
CASE WHEN @sOrderBy='A' THEN E.psnLastName ELSE NULL END,
CASE WHEN @sOrderBy='R' THEN PE.epjCreateDate ELSE NULL END



GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE usp_ProjectListSelect (
@sProjectStatusList varchar(100),	-- List of statuses splitted by commas
@sKeyword varchar(250),			-- Specific reference / Keywords 
@sOrderBy varchar(50)
)
AS
SET NOCOUNT ON

-- 1. Transform the list of statuses into a table
DECLARE @t_Status TABLE(id_Status varchar(3) COLLATE Latin1_General_CI_AI)
INSERT INTO @t_Status
SELECT VALUE FROM dbo.udf_AdmSelectFromList(@sProjectStatusList, ',')
WHERE VALUE IS NOT NULL 
AND ISNUMERIC(VALUE)=1

SET NOCOUNT OFF

SELECT *
FROM tbl_Project P
WHERE 1 = CASE 
	WHEN LEN(ISNULL(@sProjectStatusList, ''))=0 THEN 1
	WHEN P.id_ProjectStatus IN (SELECT id_Status FROM @t_Status) THEN 1 
	ELSE 0
	END
AND 1 = CASE
	WHEN LEN(ISNULL(@sKeyword, ''))<=1 THEN 1
	WHEN P.prjTitle LIKE '%' + @sKeyword + '%' THEN 1
	ELSE 0
	END
ORDER BY P.id_ProjectStatus, P.prjTitle



GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE usp_ProjectSelect (
@iProjectID int,
@sReference varchar(30)
)
AS
SELECT P.*,
PS.prsTitle prjStatus
FROM tbl_Project P
LEFT OUTER JOIN tbl_ProjectStatus PS ON P.id_ProjectStatus=PS.id_ProjectStatus
WHERE 1 = CASE 
	WHEN ISNULL(@iProjectID, 0)=0 AND LEN(ISNULL(@sReference, ''))<=1 THEN 0
	WHEN P.id_Project=@iProjectID THEN 1 
	WHEN P.prjReference=@sReference THEN 1 
	ELSE 0
	END


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE usp_ProjectStatusListSelect
AS
SELECT id_ProjectStatus, prsTitle
FROM tbl_ProjectStatus
ORDER BY id_ProjectStatus

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE usp_ProjectUpdate
@iProjectID int,
@sProjectReference varchar(30),
@sProjectShortName varchar(60),
@sProjectTitle varchar(400),
@iProjectStatusID smallint,
@sProjectLocation varchar(100),
@sProjectDescription text,
@sProjectDeadline varchar(16),
@iNewProjectID int output
AS
DECLARE @dProjectDeadline smalldatetime

IF ISDATE(@sProjectDeadline)=1
	SET @dProjectDeadline=CAST(@sProjectDeadline AS smalldatetime)
ELSE
	SET @dProjectDeadline=NULL

IF ISNULL(@iProjectID, 0)=0 
	BEGIN
	INSERT INTO tbl_Project (
	prjReference,
	prjShortName,
	prjTitle,
	id_ProjectStatus,
	prjLocation,
	prjDescription,
	prjDeadline,
	prjCreateDate
	) VALUES (
	@sProjectReference,
	@sProjectShortName,
	@sProjectTitle,
	@iProjectStatusID,
	@sProjectLocation,
	@sProjectDescription,
	@dProjectDeadline,
	GETDATE()
	)

	SELECT @iNewProjectID=@@IDENTITY FROM tbl_Project
	END
ELSE
	BEGIN
	UPDATE tbl_Project
	SET prjReference=@sProjectReference,
	prjShortName=@sProjectShortName,
	prjTitle=@sProjectTitle,
	id_ProjectStatus=@iProjectStatusID,
	prjLocation=@sProjectLocation,
	prjDescription=@sProjectDescription,
	prjDeadline=@dProjectDeadline,
	prjModifyDate=GETDATE()
	WHERE id_Project=@iProjectID

	SET @iNewProjectID=@iProjectID
	END

RETURN


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO


SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE usp_StatusCVSelect
AS
BEGIN
SELECT S.id_Status, 
S.stsNameEng,
S.stsNameFra,
S.stsNameSpa
FROM tbl_StatusCV S
ORDER BY S.id_Status

END


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO







CREATE PROCEDURE usp_UsrChangeEmail
@@sUserLogin nvarchar(50), @@sUserPassword nvarchar(50), @@sUserEmailNew nvarchar(50), 
@@bResults int OUTPUT
AS
DECLARE @iUserID int, @sUserType varchar(20)
SET @iUserID=0
SET @@bResults=0
SET NOCOUNT ON

SELECT @iUserID=id_User, @sUserType=ustDescription 
FROM tbl_Users U INNER JOIN tbl_UserType UT ON U.id_UserType=UT.id_UserType
WHERE U.UserName=@@sUserLogin AND U.[PassWord]=@@sUserPassword

IF @iUserID>0 
BEGIN
	IF @sUserType='Member'
		BEGIN
		UPDATE tbl_Members SET Email=@@sUserEmailNew WHERE id_User=@iUserID
		SET @@bResults=@@ROWCOUNT
		END
	ELSE
	IF @sUserType='Expert'
		BEGIN
		UPDATE tbl_Experts SET Email=@@sUserEmailNew WHERE id_User=@iUserID
		SET @@bResults=@@ROWCOUNT
		END
END










GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO







CREATE PROCEDURE usp_UsrChangePassword
@@sUserLogin nvarchar(50), @@sUserPassword nvarchar(50), @@sUserPasswordNew nvarchar(50), 
@@bResults int OUTPUT
AS
SET @@bResults=0
SET NOCOUNT ON

UPDATE tbl_Users SET [PassWord]=@@sUserPasswordNew
WHERE UserName=@@sUserLogin AND [PassWord]=@@sUserPassword
SET @@bResults=@@ROWCOUNT








GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE usp_UsrLogin
@sUserLogin nvarchar(50), @sUserPassword nvarchar(50), @uSessionID uniqueidentifier,
@bUserLoggedIn tinyint OUTPUT, @iUserID int OUTPUT, @sUserType varchar(20) OUTPUT, @iMemberID int OUTPUT, @iExpertID int OUTPUT, 
@sUserLanguage varchar(3) OUTPUT, @sUserEmail varchar(50) OUTPUT, @sUserIpSecurity varchar(16) OUTPUT
AS
SET @iUserID=0
SET NOCOUNT ON

SELECT TOP 1 @iUserID=id_User, @sUserType=ustDescription, @sUserIpSecurity=ISNULL(usrIpSecurity, '')
FROM tbl_Users U 
INNER JOIN tbl_UserType UT ON U.id_UserType=UT.id_UserType
WHERE U.UserName=@sUserLogin AND U.[PassWord]=@sUserPassword

IF @iUserID>0 
BEGIN
	IF @sUserType='expert'
	BEGIN
		DECLARE @iExpertOriginalID int

		SELECT @iExpertID=E.id_Expert, @iExpertOriginalID=id_ExpertOriginal, @sUserLanguage=E.Lng, @sUserEmail=E.Email 
		FROM tbl_Experts E 
		WHERE E.id_User=@iUserID
		AND (E.expDeleted=0 or E.id_ExpertOriginal>0)
		AND E.expRemoved=0

		IF @iExpertID=0 AND (NOT EXISTS(SELECT id_Expert 
			FROM tbl_Experts 
			WHERE id_Expert=@iExpertOriginalID 
			AND expDeleted=0
			AND expRemoved=0))
			BEGIN
			SET @iExpertID=0			
			END
	
		IF @iExpertID>0 
			SET @bUserLoggedIn=1 
		ELSE 
			SET @bUserLoggedIn=0
	END
	ELSE
	BEGIN
		SET @bUserLoggedIn=1 
	END
	

	IF @bUserLoggedIn=1 
		EXEC usp_LogSessionUser @uSessionID, @iUserID
END


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO


CREATE PROCEDURE usp_UsrPasswordSelect
@@sUserEmail nvarchar(255)
AS
SELECT [UserName], [Password], id_UserType, CreateDate FROM
(
SELECT U.id_User, U.UserName, U.[Password], U.id_UserType, U.CreateDate FROM tbl_Users U INNER JOIN tbl_Members M ON U.id_User=M.id_User WHERE M.Email=@@sUserEmail
UNION 
SELECT U.id_User, U.UserName, U.[Password], U.id_UserType, U.CreateDate FROM tbl_Users U INNER JOIN tbl_Experts E ON U.id_User=E.id_User WHERE E.Email=@@sUserEmail AND E.id_ExpertOriginal=0
) As UME
ORDER BY UME.id_UserType DESC, UME.id_User DESC



GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE usp_UsrSecuritySelect
@iUserID int
AS
SELECT id_UserType, 0 AS usrIbf, ISNULL(usrIpSecurity, '') AS usrIpSecurity, 0 AS usrBrowseMode, '' AS usrElinkUserCode, ISNULL(usrFastLoginID, '') AS usrFastLoginID
FROM tbl_Users
WHERE id_User=@iUserID

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE usp_DocumentByUidDelete
@sDocumentUID varchar(40)
AS
SET NOCOUNT ON

DELETE
FROM tbl_Documents 
WHERE uid_Document=@sDocumentUID
GO

CREATE PROCEDURE usp_DocumentBlobByUidSelect
@sDocumentUID varchar(40)
AS
SET NOCOUNT ON

DECLARE @uDocumentUID uniqueidentifier
SET @uDocumentUID=CAST(@sDocumentUID AS uniqueidentifier)

IF @uDocumentUID IS NULL 
	RETURN 1

SELECT 
D.docImage,
D.docImageSize
FROM tbl_Documents D
WHERE D.uid_Document=@sDocumentUID
GO

CREATE PROCEDURE usp_DocumentByUidSelect
@sDocumentUID varchar(40)
AS
SET NOCOUNT ON

DECLARE @uDocumentUID uniqueidentifier
SET @uDocumentUID=CAST(@sDocumentUID AS uniqueidentifier)

IF @uDocumentUID IS NULL 
	RETURN 1

SELECT D.id_Document,
D.uid_Document,
D.docTitle,
D.docType,
D.docText,
D.docCreated,
D.docPath,
--D.docImage,
D.docImageSize
FROM tbl_Documents D
WHERE D.uid_Document=@sDocumentUID

GO

CREATE PROCEDURE usp_ExpertDocumentListSelect
@iExpertID int, 
@sType nvarchar(255)
AS

SELECT D.id_Document,
D.uid_Document,
D.docTitle,
D.docType,
D.docText,
D.docCreated,
D.docPath,
D.docImageSize
FROM tbl_Documents D
INNER JOIN lnkExp_Doc ED ON D.id_Document=ED.id_Document
WHERE ED.id_Expert=@iExpertID
ORDER BY D.docTitle

GO

CREATE PROCEDURE usp_ExpertDocumentUpdate
@iExpertID int, 
@sDocumentUID varchar(40),
@sTitle nvarchar(255),
@sType nvarchar(255),
@sText ntext,
-- attachment
@sFileName varchar(150), 
@binFileData image
AS

IF ISNULL(@iExpertID, 0)=0
	RETURN NULL

DECLARE @uDocumentUID uniqueidentifier
SET @uDocumentUID=CAST(@sDocumentUID AS uniqueidentifier)

DECLARE @iDocumentID int
SELECT @iDocumentID=id_Document
FROM tbl_Documents
WHERE uid_Document=@uDocumentUID

IF ISNULL(@iDocumentID, 0)=0
BEGIN
	BEGIN TRAN
	INSERT INTO tbl_Documents (
	docTitle,
	docType,
	docText,
	docCreated
	) VALUES (
	@sTitle,
	@sType,
	@sText,
	GETDATE()
	)

	SELECT @iDocumentID=@@IDENTITY FROM tbl_Documents

	INSERT INTO lnkExp_Doc (
	id_Expert,
	id_Document,
	edcCreateDate
	) VALUES (
	@iExpertID, 
	@iDocumentID, 
	GETDATE()
	)

	COMMIT TRAN
END
ELSE
BEGIN
	UPDATE tbl_Documents
	SET
	docTitle=@sTitle,
	docType=@sType,
	docText=@sText,
	docUpdated=GETDATE()
	FROM tbl_Documents D, 
	lnkExp_Doc ED
	WHERE D.id_Document=@iDocumentID
	AND D.id_Document=ED.id_Document
	AND ED.id_Expert=@iExpertID
END

-- Save attachment
IF DATALENGTH(@binFileData)>1 
BEGIN
	UPDATE tbl_Documents
	SET
	docPath=@sFileName,
	docImage=@binFileData,
	docImageSize=DATALENGTH(@binFileData)
	WHERE id_Document=@iDocumentID
END

RETURN @iDocumentID

GO

CREATE PROCEDURE usp_ExpertProfileCustomUpdate
@iExpertID int, 
@sCvLanguage varchar(3),
@sCvFolder nvarchar(150)
AS
SET NOCOUNT ON

UPDATE tbl_Experts
SET
Lng=@sCvLanguage,
KgCVFile=@sCvFolder
WHERE id_Expert=@iExpertID

GO

CREATE PROCEDURE usp_ExpertProfileLanguageUpdate
@iExpertID int, 
@sCvLanguage varchar(3)
AS
SET NOCOUNT ON

UPDATE tbl_Experts
SET
Lng=@sCvLanguage
WHERE id_Expert=@iExpertID

GO

CREATE PROCEDURE usp_MmbExpSearchQueryUpdate
@iMemberID int, 
@iSearchQueryID int
AS
SET NOCOUNT ON
DECLARE @sSQLQuery nvarchar(4000), @iFlag int
SELECT @sSQLQuery=srchSQLQuery FROM log_MmbExpSearch WHERE (id_Member=@iMemberID or id_Member=0) AND id_SQLQuery=@iSearchQueryID

SET @sSQLQuery='SELECT id_Expert FROM (' + SUBSTRING(@sSQLQuery, 1, CHARINDEX(' ORDER BY', @sSQLQuery)) + ') AS T'

PRINT @sSQLQuery

SET NOCOUNT OFF

IF @sSQLQuery>''
	BEGIN
	CREATE TABLE #t_Experts (id_Expert int)

	INSERT INTO #t_Experts (
	id_Expert
	)
	EXEC sp_executesql @sSQLQuery

	INSERT INTO lnkMmb_Exp_Query (
	id_Member,
	id_Expert,
	id_Query,
	SelectedDate
	) 
	SELECT
	@iMemberID,
	id_Expert,
	@iSearchQueryID,
	GETDATE()
	FROM #t_Experts
	WHERE id_Expert NOT IN (
		SELECT id_Expert
		FROM lnkMmb_Exp_Query
		WHERE id_Member=@iMemberID
		AND id_Query=@iSearchQueryID
	)	

	DROP TABLE #t_Experts
	END

GO

CREATE PROCEDURE usp_MmbExpListQuerySelect
@iMemberID int,
@iExpertID int,
@iSearchQueryID int
AS

SELECT P.id_Person, 
CASE WHEN E.Lng='Spa' THEN ISNULL(PT.ptlNameSpa,'') WHEN E.Lng='Fra' THEN ISNULL(PT.ptlNameFra,'') ELSE ISNULL(PT.ptlNameEng,'')  END AS ptlName,
CASE WHEN E.Lng='Spa' THEN ISNULL(P.psnFirstNameSpa,ISNULL(P.psnFirstNameEng,'')) WHEN E.Lng='Fra' THEN ISNULL(P.psnFirstNameFra,'') ELSE ISNULL(P.psnFirstNameEng,'') END AS psnFirstName,
CASE WHEN E.Lng='Spa' THEN ISNULL(P.psnMiddleNameSpa,ISNULL(P.psnMiddleNameEng,'')) WHEN E.Lng='Fra' THEN ISNULL(P.psnMiddleNameFra,'') ELSE ISNULL(P.psnMiddleNameEng,'') END AS psnMiddleName,
CASE WHEN E.Lng='Spa' THEN ISNULL(P.psnLastNameSpa,ISNULL(P.psnLastNameEng,'')) WHEN E.Lng='Fra' THEN ISNULL(P.psnLastNameFra,ISNULL(P.psnLastNameEng,'')) ELSE ISNULL(P.psnLastNameEng,'') END AS psnLastName,
dbo.udf_TitleCase(P.psnLastNameEng) psnLastNameCase,
E.id_Expert, 
E.Email, 
dbo.udf_ExpertEmailAll(E.id_Expert) EmailAll,
dbo.udf_ExpertWebsite(E.id_Expert) expWebsite,
E.Phone, 
E.KgCVFile, 
P.psnBirthDate, 
E.expCreateDate, 
NULL expEmailBad,
E.expLastUpdate, 
dbo.udf_ExpertExperienceLastDate(E.id_Expert) wkeEndDate,
CAST(expComments AS varchar(8000)) expComments,
expHidden, 
expIncompleteCV, 
expApproved, 
expApprovedDate, 
expRemoved, 
expRemovedDate,
expRemovedComments,
expDeleted, 
expDeletedDate,
expDeletedComments, 
expToCompleteCVEmailSent, 
expToCompleteCVEmailDate, 
expToConfirmCvEmailSent, 
expToConfirmCvEmailDate
FROM lnkMmb_Exp_Query ME 
INNER JOIN tbl_Experts E ON ME.id_Expert=E.id_Expert
INNER JOIN tbl_Persons P ON E.id_Expert=P.id_Expert
LEFT OUTER JOIN tbl_PersonTitles PT ON P.id_psnTitle=PT.id_psnTitle
WHERE ME.id_Member=@iMemberID
AND 1 = CASE 
	WHEN ISNULL(@iExpertID, 0)>0 AND ME.id_Expert=@iExpertID THEN 1
	WHEN ISNULL(@iExpertID, 0)=0 THEN 1
	ELSE 0
	END
AND ME.id_Query=@iSearchQueryID

GO

CREATE PROCEDURE usp_MmbExpQueryUpdate
@iMemberID int,
@iExpertID int,
@iSearchQueryID int,
@iAction tinyint
AS

-- Delete selection
IF @iAction=0
	BEGIN
	DELETE FROM lnkMmb_Exp_Query 
	WHERE id_Member=@iMemberID
	AND 1 = CASE 
		WHEN ISNULL(@iExpertID, 0)>0 AND id_Expert=@iExpertID THEN 1
		WHEN ISNULL(@iExpertID, 0)=0 THEN 1
		ELSE 0
		END
	AND id_Query=@iSearchQueryID
	END

-- Insert selection
IF @iAction=1 AND 
NOT EXISTS(SELECT id_Expert
FROM lnkMmb_Exp_Query
WHERE id_Member=@iMemberID
AND id_Expert=@iExpertID
AND id_Query=@iSearchQueryID
)
	BEGIN
	INSERT INTO lnkMmb_Exp_Query (
	id_Member,
	id_Expert,
	id_Query,
	SelectedDate
	) VALUES (
	@iMemberID,
	@iExpertID,
	@iSearchQueryID,
	GETDATE()
	)
	END

GO

CREATE PROCEDURE usp_CountryListSelect 
@iCountryID int = NULL,
@sCountryName nvarchar(255) = NULL,
@iRegionID int = NULL,
@sRegionName nvarchar(255) = NULL,
@sOrderBy varchar(80) = NULL
AS
SELECT 
C.id_Country,
C.couAbbreviation, 
C.couNameEng,
C.couNameFra,
C.couNameSpa,
C.id_GeoZone id_Region,
C.id_GeoZone
FROM uvw_Country C
WHERE 
(@iCountryID IS NULL OR C.id_Country = @iCountryID)
AND
(@sCountryName IS NULL OR C.couAbbreviation = @sCountryName OR C.couNameEng LIKE @sCountryName + '%' OR C.couNameFra LIKE @sCountryName + '%' OR C.couNameSpa LIKE @sCountryName + '%')
AND
(@iRegionID IS NULL OR C.id_GeoZone = @iRegionID)
ORDER BY 
CASE WHEN @sOrderBy = 'couNameEng' THEN couNameEng ELSE NULL END,
CASE WHEN @sOrderBy = 'couNameFra' THEN couNameFra ELSE NULL END,
CASE WHEN @sOrderBy = 'couNameSpa' THEN couNameSpa ELSE NULL END,
id_Country
GO


CREATE PROCEDURE usp_LanguageListSelect 
@iLanguageID int = NULL,
@sLanguageName nvarchar(255) = NULL,
@sOrderBy varchar(80) = NULL
AS
SELECT 
L.id_Language,
L.lngNameEng,
L.lngNameFra,
L.lngNameSpa,
L.db_Order
FROM tbl_Languages L
WHERE L.db_NotVisible = 0
AND
(@iLanguageID IS NULL OR L.id_Language = @iLanguageID)
AND
(@sLanguageName IS NULL OR L.lngNameEng LIKE @sLanguageName + '%' OR L.lngNameFra LIKE @sLanguageName + '%' OR L.lngNameSpa LIKE @sLanguageName + '%')
ORDER BY L.db_Order
GO


CREATE PROCEDURE usp_EduSubjectListSelect 
@iEduSubjectID int = NULL,
@sEduSubjectName nvarchar(255) = NULL,
@sOrderBy varchar(80) = NULL
AS
SELECT 
ES.id_EduSubject,
ES.edsDescriptionEng,
ES.edsDescriptionFra,
ES.edsDescriptionSpa
FROM tbl_EduSubjects ES
WHERE 
(@iEduSubjectID IS NULL OR ES.id_EduSubject = @iEduSubjectID)
AND
(@sEduSubjectName IS NULL OR ES.edsDescriptionEng LIKE @sEduSubjectName + '%' OR ES.edsDescriptionFra LIKE @sEduSubjectName + '%' OR ES.edsDescriptionSpa LIKE @sEduSubjectName + '%')
ORDER BY 
CASE WHEN @sOrderBy = 'edsDescriptionEng' THEN edsDescriptionEng ELSE NULL END,
CASE WHEN @sOrderBy = 'edsDescriptionFra' THEN edsDescriptionFra ELSE NULL END,
CASE WHEN @sOrderBy = 'edsDescriptionSpa' THEN edsDescriptionSpa ELSE NULL END,
id_EduSubject
GO


CREATE PROCEDURE usp_EduTypeListSelect 
@iEduTypeID int = NULL,
@sEduTypeName nvarchar(255) = NULL,
@sOrderBy varchar(80) = NULL
AS
SELECT 
ET.id_EduType,
ET.edtDescriptionEng,
ET.edtDescriptionFra,
ET.edtDescriptionSpa
FROM tbl_EducationType ET
WHERE 
(@iEduTypeID IS NULL OR ET.id_EduType = @iEduTypeID)
AND
(@sEduTypeName IS NULL OR ET.edtDescriptionEng LIKE @sEduTypeName + '%' OR ET.edtDescriptionFra LIKE @sEduTypeName + '%' OR ET.edtDescriptionSpa LIKE @sEduTypeName + '%')
AND 1 = CASE
	WHEN @sOrderBy = 'edtDescriptionFra' AND edtDescriptionFra IS NOT NULL THEN 1
	WHEN @sOrderBy = 'edtDescriptionSpa' AND edtDescriptionSpa IS NOT NULL THEN 1
	WHEN edtDescriptionEng IS NOT NULL THEN 1
	ELSE 0
END
ORDER BY 
db_Order,
CASE WHEN @sOrderBy = 'edtDescriptionEng' THEN edtDescriptionEng ELSE NULL END,
CASE WHEN @sOrderBy = 'edtDescriptionFra' THEN edtDescriptionFra ELSE NULL END,
CASE WHEN @sOrderBy = 'edtDescriptionSpa' THEN edtDescriptionSpa ELSE NULL END,
id_EduType
GO


CREATE PROCEDURE usp_PersonTitleListSelect 
@sPersonTitleID int = NULL,
@sPersonTitle nvarchar(255) = NULL,
@sOrderBy varchar(80) = NULL
AS
SELECT 
PT.id_psnTitle,
PT.ptlNameEng,
PT.ptlNameFra,
PT.ptlNameSpa
FROM tbl_PersonTitles PT
WHERE 
(@sPersonTitleID IS NULL OR PT.id_psnTitle = @sPersonTitleID)
AND
(@sPersonTitle IS NULL OR PT.ptlNameEng LIKE @sPersonTitle + '%' OR PT.ptlNameFra LIKE @sPersonTitle + '%' OR PT.ptlNameSpa LIKE @sPersonTitle + '%')
ORDER BY 
CASE WHEN @sOrderBy = 'ptlNameEng' THEN ptlNameEng ELSE NULL END,
CASE WHEN @sOrderBy = 'ptlNameFra' THEN ptlNameFra ELSE NULL END,
CASE WHEN @sOrderBy = 'ptlNameSpa' THEN ptlNameSpa ELSE NULL END,
id_psnTitle
GO


CREATE PROCEDURE usp_MmbExpQuerySelect
@iMemberID int,
@iExpertID int,
@iSearchQueryID int
AS

-- Return selection
SELECT id_Member, id_Expert, id_Query, SelectedDate
FROM lnkMmb_Exp_Query
WHERE id_Member=@iMemberID
AND 1 = CASE 
	WHEN ISNULL(@iExpertID, 0)>0 AND id_Expert=@iExpertID THEN 1
	WHEN ISNULL(@iExpertID, 0)=0 THEN 1
	ELSE 0
	END
AND id_Query=@iSearchQueryID

GO


CREATE PROCEDURE usp_ExpertLanguageLinkUpdate (
@iInitialLanguageCvID int,
@iNewLanguageCvID int
)
AS
IF NOT EXISTS (
	SELECT id_Expert
	FROM tbl_ExpertsLanguage
	WHERE id_Expert = @iInitialLanguageCvID
	AND id_Expert2 = @iNewLanguageCvID
	)
BEGIN
	INSERT INTO tbl_ExpertsLanguage (
	id_Expert,
	id_Expert2,
	exlCreateDate
	) VALUES (
	@iInitialLanguageCvID,
	@iNewLanguageCvID,
	GETDATE()
	)
END
GO


CREATE FUNCTION dbo.CONVERTDATE (@Date smalldatetime)  
RETURNS varchar(10) AS  
BEGIN 
DECLARE @sYear varchar(4), @sMonth varchar(2), @sDay varchar(2)
SET @sYear=CONVERT(varchar(4),YEAR(@Date))
SET @sMonth=CONVERT(varchar(2),MONTH(@Date))
SET @sDay=CONVERT(varchar(2),DAY(@Date))
RETURN(@sYear + LEFT('00',2-LEN(@sMonth))+ @sMonth + LEFT('00',2-LEN(@sDay))+ @sDay)
END


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

/*
CVIP udf_AdmExpAllListExtraSelect
*/
CREATE FUNCTION udf_AdmExpAllListExtraModifiedSelect (
@sStatusList varchar(250),
@bShowHiddenExperts bit, 
@bShowRemovedExperts bit, 
@dLastExperienceFromDate smalldatetime, 
@dLastExperienceToDate smalldatetime,
@dCvModifiedFromDate smalldatetime, 
@dCvModifiedToDate smalldatetime
)
RETURNS TABLE
AS RETURN
(SELECT P.id_Person, 
CASE WHEN E.Lng='Spa' THEN ISNULL(PT.ptlNameSpa, '') WHEN E.Lng='Fra' THEN ISNULL(PT.ptlNameFra,'') ELSE ISNULL(PT.ptlNameEng,'')  END AS ptlName,
CASE WHEN E.Lng='Spa' THEN ISNULL(P.psnFirstNameSpa, ISNULL(P.psnFirstNameEng,'')) WHEN E.Lng='Fra' THEN ISNULL(P.psnFirstNameFra, ISNULL(P.psnFirstNameEng,'')) ELSE ISNULL(P.psnFirstNameEng,'') END AS psnFirstName,
CASE WHEN E.Lng='Spa' THEN ISNULL(P.psnMiddleNameSpa, ISNULL(P.psnMiddleNameEng,'')) WHEN E.Lng='Fra' THEN ISNULL(P.psnMiddleNameFra, ISNULL(P.psnMiddleNameEng,'')) ELSE ISNULL(P.psnMiddleNameEng,'') END AS psnMiddleName,
CASE WHEN E.Lng='Spa' THEN ISNULL(P.psnLastNameSpa, ISNULL(P.psnLastNameEng,'')) WHEN E.Lng='Fra' THEN ISNULL(P.psnLastNameFra, ISNULL(P.psnLastNameEng,'')) ELSE ISNULL(P.psnLastNameEng,'') END AS psnLastName,
dbo.udf_TitleCase(P.psnLastNameEng) psnLastNameCase,
E.id_Expert, 
E.Email, 
dbo.udf_ExpertEmailAll(E.id_Expert) EmailAll,
dbo.udf_ExpertWebsite(E.id_Expert) expWebsite,
E.Phone, 
E.KgCVFile, 
P.psnBirthDate, 
E.expCreateDate, 
NULL expEmailBad,
E.expLastUpdate, 
dbo.udf_ExpertExperienceLastDate(E.id_Expert) wkeEndDate,
CAST(expComments AS varchar(8000)) expComments,
expHidden, 
expIncompleteCV, 
expApproved, 
expApprovedDate, 
expRemoved, 
expRemovedDate,
expRemovedComments,
expDeleted, 
expDeletedDate,
expDeletedComments, 
expToCompleteCVEmailSent, 
expToCompleteCVEmailDate, 
expToConfirmCvEmailSent, 
expToConfirmCvEmailDate
FROM tbl_Experts E 
INNER JOIN tbl_Persons P ON E.id_Expert=P.id_Expert
LEFT OUTER JOIN tbl_PersonTitles PT ON P.id_psnTitle=PT.id_psnTitle
WHERE 
1 = CASE
	WHEN @bShowHiddenExperts=0 AND expHidden=0 THEN 1
	WHEN @bShowHiddenExperts=1 AND @bShowRemovedExperts=0 AND expHidden=1 THEN 1
	WHEN @bShowHiddenExperts=1 AND @bShowRemovedExperts=1 THEN 1
	ELSE 0
	END 
AND
0 = CASE 
	WHEN @bShowRemovedExperts=1 THEN 0 
	ELSE CAST(expRemoved AS tinyint) + CAST(expDeleted AS tinyint) 
	END
AND id_ExpertOriginal=0
AND 0 <= CASE 
	WHEN (@dLastExperienceFromDate IS NULL) THEN 0 
	ELSE DATEDIFF(m, @dLastExperienceFromDate, dbo.udf_ExpertExperienceLastDate(E.id_Expert)) 
	END
AND 0 <= CASE 
	WHEN (@dLastExperienceToDate IS NULL) THEN 0 
	ELSE DATEDIFF(m, dbo.udf_ExpertExperienceLastDate(E.id_Expert), @dLastExperienceToDate) 
	END
AND 0 <= CASE 
	WHEN (@dCvModifiedFromDate IS NULL) THEN 0 
	ELSE DATEDIFF(m, @dCvModifiedFromDate, E.expLastUpdate) 
	END
AND 0 <= CASE 
	WHEN (@dCvModifiedToDate IS NULL) THEN 0 
	ELSE DATEDIFF(m, E.expLastUpdate, @dCvModifiedToDate) 
	END
)

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

/*
CVIP udf_AdmExpAllListExtraSelect
*/
CREATE FUNCTION udf_AdmExpAllListExtraSelect(
@sStatusList varchar(250),
@bShowHiddenExperts bit, 
@bShowRemovedExperts bit, 
@dLastExperienceFromDate smalldatetime, 
@dLastExperienceToDate smalldatetime
)
RETURNS TABLE
AS RETURN
(SELECT P.id_Person, 
CASE WHEN E.Lng='Spa' THEN ISNULL(PT.ptlNameSpa,'') WHEN E.Lng='Fra' THEN ISNULL(PT.ptlNameFra,'') ELSE ISNULL(PT.ptlNameEng,'')  END AS ptlName,
CASE WHEN E.Lng='Spa' THEN ISNULL(P.psnFirstNameSpa,ISNULL(P.psnFirstNameEng,'')) WHEN E.Lng='Fra' THEN ISNULL(P.psnFirstNameFra,'') ELSE ISNULL(P.psnFirstNameEng,'') END AS psnFirstName,
CASE WHEN E.Lng='Spa' THEN ISNULL(P.psnMiddleNameSpa,ISNULL(P.psnMiddleNameEng,'')) WHEN E.Lng='Fra' THEN ISNULL(P.psnMiddleNameFra,'') ELSE ISNULL(P.psnMiddleNameEng,'') END AS psnMiddleName,
CASE WHEN E.Lng='Spa' THEN ISNULL(P.psnLastNameSpa,ISNULL(P.psnLastNameEng,'')) WHEN E.Lng='Fra' THEN ISNULL(P.psnLastNameFra,ISNULL(P.psnLastNameEng,'')) ELSE ISNULL(P.psnLastNameEng,'') END AS psnLastName,
E.id_Expert, 
E.Email, 
dbo.udf_ExpertEmailAll(E.id_Expert) EmailAll,
E.Phone, 
E.KgCVFile, 
P.psnBirthDate, 
E.expCreateDate, 
NULL expEmailBad,
E.expLastUpdate, 
dbo.udf_ExpertExperienceLastDate(E.id_Expert) wkeEndDate,
CAST(expComments AS varchar(8000)) expComments,
expIbfOnly, 
expHidden, 
expIncompleteCV, 
expApproved, 
expApprovedDate, 
expRemoved, 
expRemovedDate,
expRemovedComments,
expDeleted, 
expDeletedDate,
expDeletedComments, 
expToCompleteCVEmailSent, 
expToCompleteCVEmailDate, 
expToConfirmCvEmailSent, 
expToConfirmCvEmailDate
FROM tbl_Experts E 
INNER JOIN tbl_Persons P ON E.id_Expert=P.id_Expert
LEFT OUTER JOIN tbl_PersonTitles PT ON P.id_psnTitle=PT.id_psnTitle
WHERE 
0 = CASE
	WHEN @bShowHiddenExperts=1 THEN 0
	WHEN @bShowHiddenExperts=0 AND expHidden=0 THEN 0
	ELSE 1
	END 
AND
0 = CASE 
	WHEN @bShowRemovedExperts=1 THEN 0 
	ELSE CAST(expRemoved AS tinyint) + CAST(expDeleted AS tinyint) 
	END
AND id_ExpertOriginal=0
AND 0 <= CASE WHEN (@dLastExperienceFromDate IS NULL) THEN 0 ELSE DATEDIFF(m, @dLastExperienceFromDate, dbo.udf_ExpertExperienceLastDate(E.id_Expert)) END
AND 0 <= CASE WHEN (@dLastExperienceToDate IS NULL) THEN 0 ELSE DATEDIFF(m, dbo.udf_ExpertExperienceLastDate(E.id_Expert), @dLastExperienceToDate) END
)


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE FUNCTION udf_AdmSelectFromList (@sList varchar(7998), @sDelemiter char(1)) RETURNS TABLE AS
RETURN(SELECT SUBSTRING(@sDelemiter + @sList + @sDelemiter, Number + 1, CHARINDEX(@sDelemiter, @sDelemiter + @sList + @sDelemiter, Number + 1) - Number - 1) AS Value
FROM tmp_Numbers WHERE Number <= LEN(@sDelemiter + @sList + @sDelemiter) - 1 AND SUBSTRING(@sDelemiter + @sList + @sDelemiter, Number, 1) = @sDelemiter)

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE FUNCTION udf_ExpertNameFull (
@iExpertID int
) 
RETURNS varchar(500)
AS
BEGIN
DECLARE @sResult varchar(500)
SET @sResult=''

SELECT @sResult=LTRIM(RTRIM((ISNULL(PT.ptlNameEng, '') + ' ' + ISNULL(psnLastNameEng, '') + ', ' + ISNULL(psnFirstNameEng, ''))))
FROM tbl_Persons P
LEFT OUTER JOIN tbl_PersonTitles PT ON P.id_psnTitle=PT.id_psnTitle
WHERE p.id_Expert=@iExpertID

RETURN @sResult
END


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO


SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO

CREATE FUNCTION udf_ProjectListSelect (
@sProjectStatusList varchar(100),	-- List of statuses splitted by commas
@sKeyword varchar(250),			-- Specific reference / Keywords 
@sOrderBy varchar(50)
)
RETURNS @t_Result TABLE (id_Project int,
			id_ProjectStatus smallint)
AS
BEGIN

-- 1. Transform the list of statuses into a table
DECLARE @t_Status TABLE(id_Status varchar(3) COLLATE Latin1_General_CI_AI)
INSERT INTO @t_Status
SELECT VALUE FROM dbo.udf_AdmSelectFromList(@sProjectStatusList, ',')
WHERE VALUE IS NOT NULL 
AND ISNUMERIC(VALUE)=1

INSERT INTO @t_Result (
id_Project, 
id_ProjectStatus
)
SELECT DISTINCT P.id_Project, P.id_ProjectStatus
FROM tbl_Project P
WHERE 1 = CASE 
	WHEN LEN(ISNULL(@sProjectStatusList, ''))=0 THEN 1
	WHEN P.id_ProjectStatus IN (SELECT id_Status FROM @t_Status) THEN 1 
	ELSE 0
	END
AND 1 = CASE
	WHEN LEN(ISNULL(@sKeyword, ''))<=1 THEN 1
	WHEN P.prjTitle LIKE '%' + @sKeyword + '%' THEN 1
	ELSE 0
	END
ORDER BY P.id_ProjectStatus

RETURN
END


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO


CREATE FUNCTION udf_ExpertLanguage (
@iExpertID int
)
RETURNS varchar(5)
AS
BEGIN
RETURN (
	SELECT TOP 1 Lng
	FROM tbl_Experts
	WHERE id_Expert=@iExpertID
)
END
GO


CREATE PROCEDURE usp_ExpertIdDetailsSelect (
@iExpertID int
)
AS
SELECT E.id_Expert,
E.Lng,
E.id_ExpertOriginal,
dbo.udf_ExpertLanguage(E.id_ExpertOriginal) LngOriginal,
ISNULL(EL.id_Expert2, EL2.id_Expert) id_Expert2,
dbo.udf_ExpertLanguage(ISNULL(EL.id_Expert2, EL2.id_Expert)) Lng2
FROM tbl_Experts E
LEFT OUTER JOIN tbl_ExpertsLanguage EL ON (E.id_Expert=EL.id_Expert OR E.id_ExpertOriginal=EL.id_Expert)
LEFT OUTER JOIN tbl_ExpertsLanguage EL2 ON (E.id_Expert=EL2.id_Expert2 OR E.id_ExpertOriginal=EL2.id_Expert2)
WHERE E.id_Expert=@iExpertID
GO


CREATE PROCEDURE usp_ExpertCopyCvLanguage
@iExpertID int,
@iNewExpertID int

AS
--SET NOCOUNT ON
BEGIN TRAN

-- 1. No need to cerate a new User

-- 2. Expert details
UPDATE E 
SET 
id_ProfessionalStatus=E2.id_ProfessionalStatus,
expProfYears=E2.expProfYears,
expReferencesEng=E2.expReferencesEng,
expReferencesFra=E2.expReferencesFra,
expReferencesSpa=E2.expReferencesSpa,
expAvailabilityEng=E2.expAvailabilityEng,
expAvailabilityFra=E2.expAvailabilityFra,
expAvailabilitySpa=E2.expAvailabilitySpa,
expShortterm=E2.expShortterm,
expLongterm=E2.expLongterm,
Email=E2.Email,
Phone=E2.Phone,
BlackList=E2.BlackList,
KgEncoded=E2.KgEncoded,
KgCVFile=E2.KgCVFile,
expCreateDate=GETDATE(),
expLastUpdate=GETDATE(),
expRanking=E2.expRanking,
expComments=E2.expComments,
BlackListMailSent=E2.BlackListMailSent,
BlackListMailDate=E2.BlackListMailDate,
BlackListMe=E2.BlackListMe,
Subscribe=E2.Subscribe,
UnsubscribeDate=E2.UnsubscribeDate,
expAccountEmailSent=E2.expAccountEmailSent,
expRegNumber=E2.expRegNumber,
expPreferences=E2.expPreferences
FROM tbl_Experts E,
tbl_Experts E2
WHERE E2.id_Expert=@iExpertID
AND E.id_Expert=@iNewExpertID


-- 3. Person details
UPDATE P
SET
id_psnTitle=P2.id_psnTitle,

psnFirstNameEng=P2.psnFirstNameEng,
psnFirstNameFra=P2.psnFirstNameFra,
psnFirstNameSpa=P2.psnFirstNameSpa,
psnMiddleNameEng=P2.psnMiddleNameEng,
psnMiddleNameFra=P2.psnMiddleNameFra,
psnMiddleNameSpa=P2.psnMiddleNameSpa,
psnLastNameEng=P2.psnLastNameEng,
psnLastNameFra=P2.psnLastNameFra,
psnLastNameSpa=P2.psnLastNameSpa,

psnGender=P2.psnGender,
psnBirthPlaceEng=P2.psnBirthPlaceEng,
psnBirthPlaceFra=P2.psnBirthPlaceFra,
psnBirthPlaceSpa=P2.psnBirthPlaceSpa,
psnBirthDate=P2.psnBirthDate,
id_MaritalStatus=P2.id_MaritalStatus,
psnComments=P2.psnComments
FROM tbl_Persons P,
tbl_Persons P2
WHERE P2.id_Expert=@iExpertID
AND P.id_Expert=@iNewExpertID

-- 4.1 Native language
INSERT INTO tbl_Native_Lng (
id_Expert,
id_Language
)
SELECT 
@iNewExpertID,
id_Language
FROM tbl_Native_Lng
WHERE id_Expert = @iExpertID


-- 4.2 Other languages
INSERT INTO lnkExp_Lan (
id_Expert,
id_Language,
exlSpeaking,
exlWriting,
exlReading,
Language1
)
SELECT 
@iNewExpertID,
id_Language,
exlSpeaking,
exlWriting,
exlReading,
Language1
FROM lnkExp_Lan
WHERE id_Expert = @iExpertID


-- 5. Nationalities
INSERT INTO lnk_Exp_Nationality (
id_Expert,
id_Nationality,
exnCreateDate
)
SELECT 
@iNewExpertID,
id_Nationality,
exnCreateDate
FROM lnk_Exp_Nationality
WHERE id_Expert = @iExpertID

COMMIT TRAN
GO


CREATE PROCEDURE usp_ExpertExperienceSectorSelect
@iExpertID int, 
@iExpertExperienceID int = NULL,
@sOrderBy varchar(80) = NULL
AS

SELECT id_MainSector,
mnsDescriptionEng,
mnsDescriptionFra,
mnsDescriptionSpa,
id_Sector,
sctDescriptionEng,
sctDescriptionFra,
sctDescriptionSpa
FROM
(
	SELECT DISTINCT MS.id_MainSector, 
	MS.mnsDescriptionEng, 
	MS.mnsDescriptionFra, 
	MS.mnsDescriptionSpa, 
	S.id_Sector, 
	S.sctDescriptionEng,
	S.sctDescriptionFra,
	S.sctDescriptionSpa
	FROM lnkWke_Sct WS
	INNER JOIN lnkExp_Wke EW ON WS.id_ExpWke=EW.id_ExpWke 
	INNER JOIN tbl_Sectors S ON WS.id_Sector=S.id_Sector 
	INNER JOIN tbl_MainSectors MS ON S.id_MainSector=MS.id_MainSector
	WHERE EW.id_Expert=@iExpertID 
	AND (EW.id_ExpWke=@iExpertExperienceID OR @iExpertExperienceID IS NULL)
) T
ORDER BY 
id_MainSector,
CASE WHEN @sOrderBy = 'sctDescriptionEng' THEN sctDescriptionEng ELSE NULL END,
CASE WHEN @sOrderBy = 'sctDescriptionFra' THEN sctDescriptionFra ELSE NULL END,
CASE WHEN @sOrderBy = 'sctDescriptionSpa' THEN sctDescriptionSpa ELSE NULL END
GO


CREATE PROCEDURE usp_ExpertExperienceCountrySelect
@iExpertID int, 
@iExpertExperienceID int = NULL,
@sOrderBy varchar(80) = NULL
AS

SELECT id_GeoZone,
Geo_ZoneEng,
Geo_ZoneFra,
Geo_ZoneSpa,
id_Country,
couNameEng,
couNameFra,
couNameSpa
FROM
(
	SELECT DISTINCT R.id_GeoZone, 
	R.Geo_ZoneEng, 
	R.Geo_ZoneFra, 
	R.Geo_ZoneSpa, 
	C.id_Country, 
	C.couNameEng,
	C.couNameFra,
	C.couNameSpa
	FROM lnkWke_Cou WC 
	INNER JOIN lnkExp_Wke EW ON WC.id_ExpWke=EW.id_ExpWke 
	INNER JOIN tbl_Country C ON WC.id_Country=C.id_Country 
	INNER JOIN tbl_GeoZone R ON C.id_GeoZone=R.id_GeoZone
	WHERE EW.id_Expert=@iExpertID 
	AND (EW.id_ExpWke=@iExpertExperienceID OR @iExpertExperienceID IS NULL)
) T
ORDER BY 
CASE WHEN @sOrderBy LIKE '%only' THEN NULL ELSE id_GeoZone END,
CASE WHEN @sOrderBy LIKE 'couNameEng%' THEN couNameEng ELSE NULL END,
CASE WHEN @sOrderBy LIKE 'couNameFra%' THEN couNameFra ELSE NULL END,
CASE WHEN @sOrderBy LIKE 'couNameSpa%' THEN couNameSpa ELSE NULL END
GO


CREATE PROCEDURE usp_ExpertEcCountrySelect
@iExpertID int,
@sLanguage varchar(80) = 'Eng'
AS    
SET NOCOUNT ON    
    
DECLARE @sExpCou nvarchar(255), @sExpStartDate varchar(10), @sExpEndDate varchar(12), @sExpPrjTitle nvarchar(255)  
DECLARE @sExpStartDate1 varchar(200), @sExpEndDate1 varchar(200), @sExpPrjTitle1 nvarchar(3500)  
DECLARE @sExpCouFlag nvarchar(255)  
    
DECLARE ExpertExperienceCursor CURSOR FAST_FORWARD FOR    
SELECT CASE 
	WHEN @sLanguage = 'Fra' THEN C.couNameFra
	WHEN @sLanguage = 'Spa' THEN C.couNameSpa
	ELSE C.couNameEng
END couName, 
ISNULL(dbo.CONVERTDATE(EW.wkeStartDate), ' ') AS wkeStartYear, 
ISNULL(dbo.CONVERTDATE(EW.wkeEndDate), ' ') AS wkeEndYear, 
ISNULL(EW.wkePrjTitleEng, ISNULL(EW.wkeOrgNameEng, ISNULL(EW.wkePositionEng, ' '))) AS wkePrjTitleEng
/*,
C.couNameEng,
C.couNameFra,
C.couNameSpa
*/
FROM lnkExp_Wke EW INNER join lnkWke_Cou WC ON EW.id_ExpWke = WC.id_ExpWke INNER JOIN tbl_Country C ON WC.id_Country = C.id_Country  
INNER JOIN tbl_CountryDev CG ON C.id_Country=CG.id_Country  
WHERE CG.couDev=1 AND ((EW.wkeStartDate IS NOT NULL) OR (EW.wkeEndDate IS NOT NULL)) AND EW.id_Expert=@iExpertID
ORDER BY C.couNameEng, 
wkeStartYear DESC  
    
CREATE TABLE #tmp_ExpECCou    
(    
tmpExpCou nvarchar(255),    
tmpExpStartDate varchar(200),  
tmpExpEndDate varchar(200),  
tmpExpPrjTitle nvarchar(3500)  
)    
    
OPEN ExpertExperienceCursor    
FETCH NEXT FROM ExpertExperienceCursor    
INTO @sExpCou, @sExpStartDate, @sExpEndDate, @sExpPrjTitle  
    
SET @sExpCouFlag=@sExpCou    
SET @sExpStartDate1 = ''  
SET @sExpEndDate1 = ''  
SET @sExpPrjTitle1 = ''  
    
WHILE @@FETCH_STATUS = 0    
BEGIN    
 IF @sExpCouFlag=@sExpCou    
 BEGIN    
   -- Add #-# symbols to split data grouped by country name  
   SET @sExpStartDate1 = @sExpStartDate1 + IsNull(@sExpStartDate,'') + '#-#'  
   SET @sExpEndDate1 = @sExpEndDate1 + IsNull(@sExpEndDate,'') + '#-#'  
   SET @sExpPrjTitle1 = @sExpPrjTitle1 + IsNull(@sExpPrjTitle,'') + '#-#'  
 END    
 ELSE    
 BEGIN    
  
  IF RIGHT(@sExpPrjTitle1,3)='#-#'     
 BEGIN  
 SET @sExpStartDate1=LEFT(@sExpStartDate1, LEN(@sExpStartDate1)-3)    
 SET @sExpEndDate1=LEFT(@sExpEndDate1, LEN(@sExpEndDate1)-3)  
 SET @sExpPrjTitle1=LEFT(@sExpPrjTitle1, LEN(@sExpPrjTitle1)-3)  
 END  
  INSERT INTO #tmp_ExpECCou VALUES (@sExpCouFlag, @sExpStartDate1, @sExpEndDate1, @sExpPrjTitle1)  
  
  SET @sExpCouFlag=@sExpCou    
  SET @sExpStartDate1 = @sExpStartDate + '#-#'  
  SET @sExpEndDate1 = @sExpEndDate + '#-#'  
  SET @sExpPrjTitle1 = @sExpPrjTitle + '#-#'  
 END    
      
 FETCH NEXT FROM ExpertExperienceCursor    
 INTO @sExpCou, @sExpStartDate, @sExpEndDate, @sExpPrjTitle  
END    
  
  
IF RIGHT(@sExpPrjTitle1,3)='#-#'     
 BEGIN  
 SET @sExpStartDate1=LEFT(@sExpStartDate1, LEN(@sExpStartDate1)-3)    
 SET @sExpEndDate1=LEFT(@sExpEndDate1, LEN(@sExpEndDate1)-3)    
 SET @sExpPrjTitle1=LEFT(@sExpPrjTitle1, LEN(@sExpPrjTitle1)-3)    
 END  
INSERT INTO #tmp_ExpECCou VALUES (@sExpCouFlag, @sExpStartDate1, @sExpEndDate1, @sExpPrjTitle1)  
  
    
CLOSE ExpertExperienceCursor    
DEALLOCATE ExpertExperienceCursor    
    
SET NOCOUNT OFF    
    
SELECT tmpExpCou, tmpExpStartDate, tmpExpEndDate, tmpExpPrjTitle FROM #tmp_ExpECCou ORDER BY tmpExpStartDate DESC  
    
SET NOCOUNT ON    
DROP TABLE #tmp_ExpECCou    
SET NOCOUNT OFF
GO


SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE TRIGGER trg_ExpLastUpdateEdu ON [dbo].[lnkExp_Edu] 
FOR UPDATE
AS
DECLARE @iExpertID int
SELECT @iExpertID=id_Expert FROM inserted
UPDATE tbl_Experts
SET expLastUpdate=GETDATE()
WHERE id_Expert=@iExpertID

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE TRIGGER trg_ExpLastUpdateLng ON dbo.lnkExp_Lan 
FOR UPDATE
AS
DECLARE @iExpertID int
SELECT @iExpertID=id_Expert FROM inserted
UPDATE tbl_Experts
SET expLastUpdate=GETDATE()
WHERE id_Expert=@iExpertID

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE TRIGGER trg_ExpLastUpdateWke ON dbo.lnkExp_Wke 
FOR UPDATE
AS
DECLARE @iExpertID int
SELECT @iExpertID=id_Expert FROM inserted
UPDATE tbl_Experts
SET expLastUpdate=GETDATE()
WHERE id_Expert=@iExpertID

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE TRIGGER trg_ExpLastUpdateNat ON [dbo].[lnk_Exp_Nationality] 
FOR UPDATE
AS
DECLARE @iExpertID int
SELECT @iExpertID=id_Expert FROM inserted
UPDATE tbl_Experts
SET expLastUpdate=GETDATE()
WHERE id_Expert=@iExpertID

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE TRIGGER trg_ExpLastUpdateAdr ON dbo.tbl_Exp_Address 
FOR UPDATE
AS
DECLARE @iExpertID int
SELECT @iExpertID=id_Expert FROM inserted
UPDATE tbl_Experts
SET expLastUpdate=GETDATE()
WHERE id_Expert=@iExpertID

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE TRIGGER trg_ExpAddress ON dbo.tbl_Exp_Address 
FOR UPDATE
AS
DECLARE @iAddressOldID int, @iCountryOldID int, @sPhoneOld varchar(150), @sFaxOld varchar(150), @sEmailOld nvarchar(150), @sMobileOld varchar(150), @sStreetOld nvarchar(255), @sPostCodeOld varchar(150), @sCityOld nvarchar(150),  @sWebOld varchar(150)
DECLARE @iAddressNewID int, @iCountryNewID int, @sPhoneNew varchar(150), @sFaxNew varchar(150), @sEmailNew nvarchar(150), @sMobileNew varchar(150), @sStreetNew nvarchar(255), @sPostCodeNew varchar(150), @sCityNew nvarchar(150),  @sWebNew varchar(150)
SELECT @iAddressOldID=id_Address, @iCountryOldID=id_Country, @sPhoneOld=LTRIM(RTRIM(adrPhone)), @sFaxOld=LTRIM(RTRIM(adrFax)), @sEmailOld=LTRIM(RTRIM(adrEmail)), @sMobileOld=LTRIM(RTRIM(adrMobile)), @sStreetOld=LTRIM(RTRIM(adrStreetEng)), @sPostCodeOld=LTRIM(RTRIM(adrPostCode)), @sCityOld=LTRIM(RTRIM(adrCityEng)),  @sWebOld=LTRIM(RTRIM(adrWeb)) FROM DELETED
SELECT @iAddressNewID=id_Address, @iCountryNewID=id_Country, @sPhoneNew=LTRIM(RTRIM(adrPhone)), @sFaxNew=LTRIM(RTRIM(adrFax)), @sEmailNew=LTRIM(RTRIM(adrEmail)), @sMobileNew=LTRIM(RTRIM(adrMobile)), @sStreetNew=LTRIM(RTRIM(adrStreetEng)), @sPostCodeNew=LTRIM(RTRIM(adrPostCode)), @sCityNew=LTRIM(RTRIM(adrCityEng)),  @sWebNew=LTRIM(RTRIM(adrWeb)) FROM INSERTED
IF NOT (ISNULL(@iCountryOldID, 0)=ISNULL(@iCountryNewID, 0) AND @sPhoneOld=@sPhoneNew AND @sFaxOld=@sFaxNew AND @sEmailOld=@sEmailNew AND @sMobileOld=@sMobileNew AND @sStreetOld=@sStreetNew AND @sCityOld=@sCityNew AND @sPostCodeOld=@sPostCodeNew AND @sWebOld=@sWebNew)
	UPDATE tbl_Exp_Address SET adrModified=GETDATE() WHERE id_Address=@iAddressNewID

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE TRIGGER trg_ExpLastUpdateLnn ON [dbo].[tbl_Native_Lng] 
FOR UPDATE
AS
DECLARE @iExpertID int
SELECT @iExpertID=id_Expert FROM inserted
UPDATE tbl_Experts
SET expLastUpdate=GETDATE()
WHERE id_Expert=@iExpertID

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE TRIGGER trg_ExpLastUpdatePsn ON dbo.tbl_Persons 
FOR UPDATE
AS
DECLARE @iExpertID int
SELECT @iExpertID=id_Expert FROM inserted
UPDATE tbl_Experts
SET expLastUpdate=GETDATE()
WHERE id_Expert=@iExpertID

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO



GRANT EXEC ON usp_LogSessionUser TO rms
GRANT EXEC ON usp_SrvRemoveNonTextSymbols TO rms
GRANT EXEC ON usp_AdmCreateList1FromRecordset TO rms
GRANT EXEC ON usp_AdmCreateList2FromRecordset TO rms
GRANT EXEC ON usp_AdmCreateList3FromRecordset TO rms
GRANT EXEC ON usp_AdmCreateRecordset TO rms
GRANT EXEC ON usp_AdmExpAllListExtraModifiedSelect TO rms
GRANT EXEC ON usp_AdmExpAllListExtraModifiedWithPasswordSelect TO rms
GRANT EXEC ON usp_AdmExpAllListExtraSelect TO rms
GRANT EXEC ON usp_AdmExpAllListSelect TO rms
GRANT EXEC ON usp_AdmExpDuplicateHide TO rms
GRANT EXEC ON usp_AdmExpExistsCheck TO rms
GRANT EXEC ON usp_AdmExpRemove TO rms
GRANT EXEC ON usp_AdmExpRestore TO rms
GRANT EXEC ON usp_CurrencyListSelect TO rms
GRANT EXEC ON usp_DatContinentSelect TO rms
GRANT EXEC ON usp_ExpAccountDetailsSelect TO rms
GRANT EXEC ON usp_ExpCvvADBCouSelect TO rms
GRANT EXEC ON usp_ExpCvvAddressInsert TO rms
GRANT EXEC ON usp_ExpCvvAddressSelect TO rms
GRANT EXEC ON usp_ExpCvvAddressUpdate TO rms
GRANT EXEC ON usp_ExpCvvAvailabilityUpdate TO rms
GRANT EXEC ON usp_ExpCvvBring2Assortis TO rms
GRANT EXEC ON usp_ExpCvvCommentsUpdate TO rms
GRANT EXEC ON usp_ExpCvvECCouSelect TO rms
GRANT EXEC ON usp_ExpCvvEducationDelete TO rms
GRANT EXEC ON usp_ExpCvvEducationInfoSelect TO rms
GRANT EXEC ON usp_ExpCvvEducationInsert TO rms
GRANT EXEC ON usp_ExpCvvEducationInsertNew TO rms
GRANT EXEC ON usp_ExpCvvEducationSelect TO rms
GRANT EXEC ON usp_ExpCvvEducationUpdate TO rms
GRANT EXEC ON usp_ExpCvvExpInfoSelect TO rms
GRANT EXEC ON usp_ExpCvvExpInfoUpdate TO rms
GRANT EXEC ON usp_ExpCvvExperienceBriefInsert TO rms
GRANT EXEC ON usp_ExpCvvExperienceCouDelete TO rms
GRANT EXEC ON usp_ExpCvvExperienceCouInsert TO rms
GRANT EXEC ON usp_ExpCvvExperienceCouSelect TO rms
GRANT EXEC ON usp_ExpCvvExperienceDelete TO rms
GRANT EXEC ON usp_ExpCvvExperienceDonDelete TO rms
GRANT EXEC ON usp_ExpCvvExperienceDonInsert TO rms
GRANT EXEC ON usp_ExpCvvExperienceDonSelect TO rms
GRANT EXEC ON usp_ExpCvvExperienceInfoSelect TO rms
GRANT EXEC ON usp_ExpCvvExperienceInsert TO rms
GRANT EXEC ON usp_ExpCvvExperienceSctDelete TO rms
GRANT EXEC ON usp_ExpCvvExperienceSctInsert TO rms
GRANT EXEC ON usp_ExpCvvExperienceSctSelect TO rms
GRANT EXEC ON usp_ExpCvvExperienceSelect TO rms
GRANT EXEC ON usp_ExpCvvExperienceUpdate TO rms
GRANT EXEC ON usp_ExpCvvLanguageInfoSelect TO rms
GRANT EXEC ON usp_ExpCvvLanguageInsert TO rms
GRANT EXEC ON usp_ExpCvvLanguageNativeDelete TO rms
GRANT EXEC ON usp_ExpCvvLanguageNativeInsert TO rms  
GRANT EXEC ON usp_ExpCvvLanguageOtherDelete TO rms
GRANT EXEC ON usp_ExpCvvLanguageOtherInsert TO rms
GRANT EXEC ON usp_ExpCvvLanguageOtherUpdate TO rms
GRANT EXEC ON usp_ExpCvvLanguageSelect TO rms
GRANT EXEC ON usp_ExpCvvNationalityDelete TO rms
GRANT EXEC ON usp_ExpCvvNationalityInsert TO rms
GRANT EXEC ON usp_ExpCvvNationalitySelect TO rms
GRANT EXEC ON usp_ExpCvvOriginalSelect TO rms
GRANT EXEC ON usp_ExpCvvPsnInfoInsert TO rms
GRANT EXEC ON usp_ExpCvvPsnInfoSelect TO rms
GRANT EXEC ON usp_ExpCvvPsnInfoUpdate TO rms
GRANT EXEC ON usp_ExpCvvTrainingInsert TO rms
GRANT EXEC ON usp_ExpCvvTrainingUpdate TO rms
GRANT EXEC ON usp_ExpertAccountEmailSentUpdate TO rms
GRANT EXEC ON usp_ExpertNationalityUpdate TO rms
GRANT EXEC ON usp_ExpertProfileFullUpdate TO rms
GRANT EXEC ON usp_ExpertProfileShortUpdate TO rms
GRANT EXEC ON usp_ExpertProjectDelete TO rms
GRANT EXEC ON usp_ExpertProjectListSelect TO rms
GRANT EXEC ON usp_ExpertProjectSelect TO rms
GRANT EXEC ON usp_ExpertProjectUpdate TO rms
GRANT EXEC ON usp_ExpertStatusCVSelect TO rms
GRANT EXEC ON usp_ExpertStatusCVUpdate TO rms
GRANT EXEC ON usp_ExpertStatusListSelect TO rms
GRANT EXEC ON usp_GetExpertProfDetails TO rms
GRANT EXEC ON usp_LogErrorAdd TO rms
GRANT EXEC ON usp_LogSessionCreate TO rms
GRANT EXEC ON usp_LogSessionEvent TO rms
GRANT EXEC ON usp_LogSessionEvent TO rms
GRANT EXEC ON usp_LogSessionUserDataSelect TO rms
GRANT EXEC ON usp_LogSessionValidate TO rms
GRANT EXEC ON usp_MmbExpDownloadedCleanup TO rms
GRANT EXEC ON usp_MmbExpSearchFirstSelect TO rms
GRANT EXEC ON usp_MmbExpSearchRepeatSelect TO rms
GRANT EXEC ON usp_ProfessionalStatusListSelect TO rms
GRANT EXEC ON usp_ProjectDelete TO rms
GRANT EXEC ON usp_ProjectExpertListSelect TO rms
GRANT EXEC ON usp_ProjectListSelect TO rms
GRANT EXEC ON usp_ProjectSelect TO rms
GRANT EXEC ON usp_ProjectStatusListSelect TO rms
GRANT EXEC ON usp_ProjectUpdate TO rms
GRANT EXEC ON usp_StatusCVSelect TO rms
GRANT EXEC ON usp_UsrChangeEmail TO rms
GRANT EXEC ON usp_UsrChangePassword TO rms
GRANT EXEC ON usp_UsrLogin TO rms
GRANT EXEC ON usp_UsrPasswordSelect TO rms
GRANT EXEC ON usp_UsrSecuritySelect TO rms
GRANT EXEC ON usp_DocumentByUidDelete TO rms
GRANT EXEC ON usp_DocumentBlobByUidSelect TO rms
GRANT EXEC ON usp_DocumentByUidSelect TO rms
GRANT EXEC ON usp_ExpertDocumentListSelect TO rms
GRANT EXEC ON usp_ExpertDocumentUpdate TO rms
GRANT EXEC ON usp_ExpertProfileCustomUpdate TO rms
GRANT EXEC ON usp_ExpertProfileLanguageUpdate TO rms
GRANT EXEC ON usp_MmbExpSearchQueryUpdate TO rms
GRANT EXEC ON usp_MmbExpListQuerySelect TO rms
GRANT EXEC ON usp_MmbExpQueryUpdate TO rms
GRANT EXEC ON usp_MmbExpQuerySelect TO rms
GRANT EXEC ON usp_CountryListSelect TO rms
GRANT EXEC ON usp_LanguageListSelect TO rms
GRANT EXEC ON usp_EduSubjectListSelect TO rms
GRANT EXEC ON usp_EduTypeListSelect TO rms
GRANT EXEC ON usp_PersonTitleListSelect TO rms
GRANT EXEC ON usp_ExpertIdDetailsSelect TO rms
GRANT EXEC ON usp_ExpertLanguageLinkUpdate TO rms
GRANT EXEC ON usp_ExpertCopyCvLanguage TO rms
GRANT EXEC ON usp_ExpertExperienceSectorSelect TO rms
GRANT EXEC ON usp_ExpertExperienceCountrySelect TO rms
GRANT EXEC ON usp_ExpertEcCountrySelect TO rms

GRANT SELECT ON tbl_Country TO rms
GRANT SELECT ON tbl_EduSubjects TO rms
GRANT SELECT ON tbl_EducationType TO rms

GRANT SELECT ON tbl_Persons TO rms
GRANT SELECT ON tbl_Experts TO rms
GRANT SELECT ON lnkExp_Wke TO rms
GRANT SELECT ON lnkWke_Sct TO rms
GRANT SELECT ON lnkWke_Cou TO rms
GRANT SELECT ON lnkWke_Don TO rms
GRANT SELECT ON lnk_Exp_Nationality TO rms
GRANT SELECT ON lnkExp_RankSct TO rms
GRANT SELECT ON lnkExp_RankCou TO rms
GRANT SELECT ON tbl_Native_Lng TO rms
GRANT SELECT ON lnkExp_Edu TO rms
GRANT SELECT ON lnkExp_Lan TO rms
GRANT SELECT ON tbl_Exp_Address TO rms
GRANT SELECT ON lnkExp_StatusCV TO rms


INSERT INTO tbl_UserType
SELECT * FROM rms__default.dbo.tbl_UserType

INSERT INTO tbl_Country
SELECT * FROM rms__default.dbo.tbl_Country

INSERT INTO tbl_Sectors
SELECT * FROM rms__default.dbo.tbl_Sectors

INSERT INTO tbl_MainSectors
SELECT * FROM rms__default.dbo.tbl_MainSectors

INSERT INTO tmp_Numbers
SELECT * FROM rms__default.dbo.tmp_Numbers

INSERT INTO tbl_Continent
SELECT * FROM rms__default.dbo.tbl_Continent

INSERT INTO tbl_Currency
SELECT * FROM rms__default.dbo.tbl_Currency

INSERT INTO tbl_Donors
SELECT * FROM rms__default.dbo.tbl_Donors

INSERT INTO tbl_EducationType
SELECT * FROM rms__default.dbo.tbl_EducationType

INSERT INTO tbl_EduSubjects
SELECT * FROM rms__default.dbo.tbl_EduSubjects

INSERT INTO tbl_Gender
SELECT * FROM rms__default.dbo.tbl_Gender

INSERT INTO tbl_GeoZone
SELECT * FROM rms__default.dbo.tbl_GeoZone

INSERT INTO tbl_LangLevel
SELECT * FROM rms__default.dbo.tbl_LangLevel

INSERT INTO tbl_Languages
SELECT * FROM rms__default.dbo.tbl_Languages

INSERT INTO tbl_LegalStatus
SELECT * FROM rms__default.dbo.tbl_LegalStatus

INSERT INTO tbl_MaritalStatus
SELECT * FROM rms__default.dbo.tbl_MaritalStatus

INSERT INTO tbl_PersonTitles
SELECT * FROM rms__default.dbo.tbl_PersonTitles


IF NOT EXISTS (
	SELECT id_Country
	FROM tbl_CountryDev
)
BEGIN
	INSERT INTO tbl_CountryDev
	SELECT * FROM rms__default.dbo.tbl_CountryDev
END

IF NOT EXISTS (
	SELECT id_Country
	FROM tbl_CountryGDP
)
BEGIN
	INSERT INTO tbl_CountryGDP
	SELECT * FROM rms__default.dbo.tbl_CountryGDP
END
