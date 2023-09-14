<% 
'--------------------------------------------------------------------
'
' Data Access Layer functions.
'
' Functions for accessing & executing stored procedures
'
'--------------------------------------------------------------------

Dim objTempOutParams(50)


'--------------------------------------------------------------------
' Function for reading data through stored procedure to recordset
'--------------------------------------------------------------------
Function GetDataRecordsetSP(sStoredProcedureName, arrInParams())
    Dim objCmd1
    Dim objTempRs

    Set objCmd1=Server.CreateObject("ADODB.Command")

    ' Init the stored procedure sStoredProcedureName parameters
    objCmd1.ActiveConnection = objconn
    objCmd1.CommandText = sStoredProcedureName
    objCmd1.CommandType = adCmdStoredProc
    collectInParams objCmd1, arrInParams 

    Set objTempRs=Server.CreateObject("ADODB.Recordset")
    objTempRs.CursorType = adOpenStatic 
    objTempRs.CursorLocation = adUseClient
    objTempRs.LockType = adLockOptimistic

    objTempRs.Open objCmd1

    Set GetDataRecordsetSP=objTempRs
    
    Set objCmd1.ActiveConnection = Nothing
    Set objCmd1 = Nothing
    Set objTempRs = Nothing
End Function


'--------------------------------------------------------------------
' Function for reading data through stored procedure to recordset with specified Connection
'--------------------------------------------------------------------
Function GetDataRecordsetSPWithConn(objConn, sStoredProcedureName, arrInParams())
    Dim objCmd1
    Dim objTempRs

    Set objCmd1=Server.CreateObject("ADODB.Command")

    ' Init the stored procedure sStoredProcedureName parameters
    objCmd1.ActiveConnection = objconn
    objCmd1.CommandText = sStoredProcedureName
    objCmd1.CommandType = adCmdStoredProc
    collectInParams objCmd1, arrInParams 

    Set objTempRs=Server.CreateObject("ADODB.Recordset")
    objTempRs.CursorType = adOpenStatic 
    objTempRs.CursorLocation = adUseClient
    objTempRs.LockType = adLockOptimistic

    objTempRs.Open objCmd1

    Set GetDataRecordsetSPWithConn=objTempRs
    
    Set objCmd1.ActiveConnection = Nothing
    Set objCmd1 = Nothing
    Set objTempRs = Nothing
End Function


'--------------------------------------------------------------------
' Function for reading data through stored procedure to recordset 
' 						& output params
'--------------------------------------------------------------------
Function GetDataRecordsetPlusOutSP(sStoredProcedureName, arrInParams(), arrOutParams())
Dim objCmd1
Dim objTempRs

    Set objCmd1=Server.CreateObject("ADODB.Command")

    ' Init the stored procedure sStoredProcedureName parameters
    objCmd1.ActiveConnection = objconn
    objCmd1.CommandText = sStoredProcedureName
    objCmd1.CommandType = adCmdStoredProc
    collectInParams objCmd1, arrInParams 
    collectOutParams objCmd1, arrOutParams

    Set objTempRs=Server.CreateObject("ADODB.Recordset")
    objTempRs.CursorType = adOpenStatic 
    objTempRs.CursorLocation = adUseClient
    objTempRs.LockType = adLockOptimistic

    objTempRs.Open objCmd1

    For i=LBound(arrOutParams) To UBound(arrOutParams)
    objTempOutParams(i) = objCmd1.Parameters(i+UBound(arrInParams)+1).Value
    Next

    Set GetDataRecordsetPlusOutSP=objTempRs
    
    Set objCmd1.ActiveConnection = Nothing
    Set objCmd1 = Nothing
    Set objTempRs = Nothing
End Function


'--------------------------------------------------------------------
' Function for reading data through stored procedure to recordset 
' 						& output params
'--------------------------------------------------------------------
Function GetDataRecordsetPlusOutSPWithConn(objConn, sStoredProcedureName, arrInParams(), arrOutParams())
Dim objCmd1
Dim objTempRs

    Set objCmd1=Server.CreateObject("ADODB.Command")

    ' Init the stored procedure sStoredProcedureName parameters
    objCmd1.ActiveConnection = objConn
    objCmd1.CommandText = sStoredProcedureName
    objCmd1.CommandType = adCmdStoredProc
    collectInParams objCmd1, arrInParams 
    collectOutParams objCmd1, arrOutParams

    Set objTempRs=Server.CreateObject("ADODB.Recordset")
    objTempRs.CursorType = adOpenStatic 
    objTempRs.CursorLocation = adUseClient
    objTempRs.LockType = adLockOptimistic

    objTempRs.Open objCmd1

    For i=LBound(arrOutParams) To UBound(arrOutParams)
    objTempOutParams(i) = objCmd1.Parameters(i+UBound(arrInParams)+1).Value
    Next

    Set GetDataRecordsetPlusOutSPWithConn=objTempRs
    
    Set objCmd1.ActiveConnection = Nothing
    Set objCmd1 = Nothing
    Set objTempRs = Nothing
End Function


'--------------------------------------------------------------------
' Function for reading data through stored procedure to output params
'--------------------------------------------------------------------
Function GetDataOutParamsSP(sStoredProcedureName, arrInParams(), arrOutParams())
Dim objTempRs(50)
Dim objCmd1

    Set objCmd1=Server.CreateObject("ADODB.Command")

    ' Init the stored procedure sStoredProcedureName parameters
    objCmd1.ActiveConnection = objconn
    objCmd1.CommandText = sStoredProcedureName
    objCmd1.CommandType = adCmdStoredProc
    collectInParams objCmd1, arrInParams
    collectOutParams objCmd1, arrOutParams

    objCmd1.Execute ,,adExecuteNoRecords
    For i=LBound(arrOutParams) To UBound(arrOutParams)
    objTempRs(i) = objCmd1.Parameters(i+UBound(arrInParams)+1).Value
    Next
    
    Set objCmd1.ActiveConnection = Nothing
    Set objCmd1 = Nothing
    GetDataOutParamsSP=objTempRs
End Function

'--------------------------------------------------------------------
' Function for reading data through stored procedure to output params with Connection
'--------------------------------------------------------------------
Function GetDataOutParamsSPWithConn(objConn, sStoredProcedureName, arrInParams(), arrOutParams())
Dim objTempRs(50)
Dim objCmd1

    Set objCmd1=Server.CreateObject("ADODB.Command")

    ' Init the stored procedure sStoredProcedureName parameters
    objCmd1.ActiveConnection = objconn
    objCmd1.CommandText = sStoredProcedureName
    objCmd1.CommandType = adCmdStoredProc
    collectInParams objCmd1, arrInParams
    collectOutParams objCmd1, arrOutParams

    objCmd1.Execute ,,adExecuteNoRecords
    For i=LBound(arrOutParams) To UBound(arrOutParams)
    objTempRs(i) = objCmd1.Parameters(i+UBound(arrInParams)+1).Value
    Next
    
    Set objCmd1.ActiveConnection = Nothing
    Set objCmd1 = Nothing
    GetDataOutParamsSPWithConn=objTempRs
End Function


'--------------------------------------------------------------------
' Function inserting record through stored procedure sStoredProcedureName
' and returning record id
'--------------------------------------------------------------------
Function InsertRecordSP(sStoredProcedureName, arrParams(), strRV)
Dim rs
Dim objCmd1

    Set objCmd1=Server.CreateObject("ADODB.Command")

    ' Init the stored procedure sStoredProcedureName parameters
    objCmd1.ActiveConnection = objconn
    objCmd1.CommandText = sStoredProcedureName
    objCmd1.CommandType = adCmdStoredProc
    collectInParams objCmd1, arrParams

    ' On insert return identity value
    If strRV="ID" Then
	objCmd1.Parameters.Append objCmd1.CreateParameter("@ID", adInteger, adParamOutput)
    End If

    objCmd1.Execute ,,adExecuteNoRecords
    If strRV="ID" Then
	InsertRecordSP = objCmd1.Parameters("@ID").Value
    End If
    
    Set objCmd1.ActiveConnection = Nothing
    Set objCmd1 = Nothing
End Function


'--------------------------------------------------------------------
' Function updating record(s) through stored procedure sStoredProcedureName
'--------------------------------------------------------------------
Function UpdateRecordSP(sStoredProcedureName, arrInParams())
Dim objCmd1
Dim objTempRs, Result

    Set objCmd1=Server.CreateObject("ADODB.Command")

    ' Init the ADO objects & the stored proc parameters
    objCmd1.ActiveConnection = objconn
    objCmd1.CommandText = sStoredProcedureName
    objCmd1.CommandType = adCmdStoredProc
'	objCmd1.Parameters.Append objCmd1.CreateParameter ("RETURN_VALUE", adInteger, adParamReturnValue)
    collectInParams objCmd1, arrInParams
    objCmd1.Execute ,,adExecuteNoRecords
	
'	Result = objCmd1.Parameters("RETURN_VALUE")
    Set objCmd1.ActiveConnection = Nothing
    
UpdateRecordSP = Result
End Function

'--------------------------------------------------------------------
' Function updating record(s) through stored procedure sStoredProcedureName
'--------------------------------------------------------------------
Function UpdateRecordSPWithConn(objConn, sStoredProcedureName, sParams())
Dim objCmd1

    Set objCmd1=Server.CreateObject("ADODB.Command")

    ' Init the ADO objects & the stored proc parameters
    objCmd1.ActiveConnection = objconn
    objCmd1.CommandText = sStoredProcedureName
    objCmd1.CommandType = adCmdStoredProc
    collectInParams objCmd1, sParams
    objCmd1.Execute,,adExecuteNoRecords
    
    Set objCmd1.ActiveConnection = Nothing
    Exit Function

End Function


'--------------------------------------------------------------------
' Function deleting record(s) through stored procedure sStoredProcedureName
'--------------------------------------------------------------------
Function DeleteRecordSP(sStoredProcedureName, sParams())
Dim objCmd1

    Set objCmd1=Server.CreateObject("ADODB.Command")

    ' Init the ADO objects & the stored proc parameters
    objCmd1.ActiveConnection = objconn
    objCmd1.CommandText = sStoredProcedureName
    objCmd1.CommandType = adCmdStoredProc
    collectInParams objCmd1, sParams
    objCmd1.Execute ,,adExecuteNoRecords
    
    Set objCmd1.ActiveConnection = Nothing
    Exit Function

End Function


'--------------------------------------------------------------------
' Function deleting record(s) through stored procedure sStoredProcedureName
'--------------------------------------------------------------------
Function DeleteRecordSPWithConn(objConn, sStoredProcedureName, sParams())
Dim objCmd1

    Set objCmd1=Server.CreateObject("ADODB.Command")

    ' Init the ADO objects & the stored proc parameters
    objCmd1.ActiveConnection = objConn
    objCmd1.CommandText = sStoredProcedureName
    objCmd1.CommandType = adCmdStoredProc
    collectInParams objCmd1, sParams
    objCmd1.Execute ,,adExecuteNoRecords
    
    Set objCmd1.ActiveConnection = Nothing
    Exit Function

End Function

'--------------------------------------------------------------------
' Procedure that create input params in sp
'--------------------------------------------------------------------
Sub collectInParams(objCmd, arrParams())
Const adFldLong = 128   ' (&H80)
Const adTypeBinary = 1
Dim iErrorNumber, sErrorSource, sErrorDescription, sErrorParams

    Dim i, v
    For i = LBound(arrParams) To UBound(arrParams)
        If UBound(arrParams(i)) = 3 Then
            ' Check for nulls.
	    If arrParams(i)(1)=adVarWChar Or arrParams(i)(1)=adVarChar Then
		If arrParams(i)(3) > "" Then
			v=arrParams(i)(3)
		Else
			v=Null
		End If
            ElseIf arrParams(i)(1)=adInteger Or arrParams(i)(1)=adTinyInt Or arrParams(i)(1)=adSmallInt Or arrParams(i)(1)=adBigInt Then
		If IsNumeric(arrParams(i)(3)) And arrParams(i)(3)>"" Then
			v=arrParams(i)(3)
		Else
			v=Null
		End If
            Else
                v = arrParams(i)(3)
            End If
	    If (arrParams(i)(1) <> adLongVarBinary) Then
		    On Error Resume Next
	            objCmd.Parameters.Append objCmd.CreateParameter(arrParams(i)(0), _
        	      arrParams(i)(1), adParamInput, arrParams(i)(2), v)
		    iErrorNumber=Err.Number
		    sErrorSource=Err.Source
		    sErrorDescription=Err.Description

		    On Error GoTo 0
		    If iErrorNumber>0 Then 
			Response.Flush
			Response.Write i & "<br>"
			Response.Write iErrorNumber & "<br>"
			Response.Write sErrorSource & "<br>"
			Response.Write sErrorDescription & "<br>"
			Response.End
		    End If
		
	    Else  
		If IsNull(v) Then
	            objCmd.Parameters.Append objCmd.CreateParameter(arrParams(i)(0), _
        	      arrParams(i)(1), adParamInput, 4096, Null)
		Else
		    Dim nChunkSize
		    Dim nRem
		    Dim objStream

		    Set objStream=Server.CreateObject("ADODB.Stream")
		    objStream.Type = adTypeBinary
		    objStream.Open
		    objStream.LoadFromFile v

                    nChunkSize = objStream.Size / 4096
                    nRem = objStream.Size Mod 4096
                    objStream.Position = 0

	            objCmd.Parameters.Append objCmd.CreateParameter(arrParams(i)(0), _
        	      arrParams(i)(1), adParamInput, arrParams(i)(2))
                     
                    Do Until objStream.Position >= objStream.Size - nRem
                        objCmd.Parameters(i).AppendChunk objStream.Read(4096)
                    Loop
                             
                    If nRem > 0 Then
                        objCmd.Parameters(i).AppendChunk objStream.Read(nRem)
                    End If
                    objStream.Close
		End If

	    End If

        Else
            objCmd.Parameters.Append objCmd.CreateParameter(arrParams(i)(0), _
              arrParams(i)(1), adParamInput, arrParams(i)(2), arrParams(i)(3))
	End If


    Next
End Sub


'--------------------------------------------------------------------
' Procedure that create output params in sp
'--------------------------------------------------------------------
Sub collectOutParams(objCmd, arrParams())
    Dim i, v
    For i = LBound(arrParams) To UBound(arrParams)
	If (arrParams(i)(1)=adInteger Or arrParams(i)(1)=adTinyInt Or arrParams(i)(1)=adSmallInt Or arrParams(i)(1)=adBigInt Or arrParams(i)(1)=adDouble) Then
            objCmd.Parameters.Append objCmd.CreateParameter(arrParams(i)(0), _
              arrParams(i)(1), adParamOutput)
	Else
            objCmd.Parameters.Append objCmd.CreateParameter(arrParams(i)(0), _
              arrParams(i)(1), adParamOutput, arrParams(i)(2))
	End If
    Next
End Sub

%>