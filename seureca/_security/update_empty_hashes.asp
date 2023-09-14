<!--#include virtual = "/_template/asp.header.notimeout.asp"-->

<!--#include file = "../dbc.asp"-->
<!--#include file = "../_fnc.asp"-->
<!--#include file = "../_encryption.asp"-->
<%

Dim sUserPassword, sUserPasswordHash

' Get the list of empty hashes
Set objTempRs=GetDataRecordsetSP("usp_UserPasswordHashEmptySelect", Array( _
	))
While Not objTempRs.Eof
	iUserID = CheckIntegerAndZero(objTempRs("id_User"))
	sUserLogin = objTempRs("usrLogin")
	sUserPassword = objTempRs("PassWord")

	' Generate a new hash
	sUserPasswordHash = Encryption.Sha1DoubleEncodeWithSalt(sUserLogin, sUserPassword)

	Response.Write "UserID: " & iUserID & " / Login: " & sUserLogin & " / Password: " & sUserPassword & " / Hash: " & sUserPasswordHash & "<br/>"

	' Update hash in DB
	objTempRs2=UpdateRecordSP("usp_UserPasswordHashUpdate", Array( _
		Array(, adInteger, , iUserID), _
		Array(, adVarChar, 2048, sUserPasswordHash)))

	Response.Flush
	objTempRs.MoveNext
WEnd

%>