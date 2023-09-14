<%
'--------------------------------------------------------------------
' Function for checking & correcting text input strings 
'--------------------------------------------------------------------
Function CheckString(sTextInput)
'Dim sTextTemp
'	If Len(sTextInput)>0 Then
'		sTextTemp=Replace(Trim(sTextInput),"'","''")
'	Else
'		sTextTemp=""
'	End If
'	CheckString=sTextTemp
	CheckString=sTextInput
End Function

Function CheckSpaces(sTextInput, iSpacePosition)
Dim sResult, sTextTemp, i, k1, k2
sTextTemp=sTextInput
sResult=""

	If Len(sTextTemp)>iSpacePosition Then
		k1=0

		While Len(sTextTemp)>iSpacePosition
			k2=InStr(sTextTemp, " ")
			
			If k2>0 And k2-k1<iSpacePosition Then
				sResult=sResult & Left(sTextTemp, k2)
				k1=k2+1
			Else
				sResult=sResult & Left(sTextTemp, iSpacePosition) & " "
				k1=iSpacePosition+1
			End If

			sTextTemp=Mid(sTextTemp, k1, Len(sTextTemp)-k1+1)
		WEnd
	End If
	sResult=sResult & sTextTemp

	CheckSpaces=sResult
End Function


Function CheckInteger(sTextInput)
Dim sTextTemp
	If sTextInput>"" And IsNumeric(sTextInput) Then
		sTextTemp=sTextInput
	Else
		sTextTemp=Null
	End If
	CheckInteger=sTextTemp
End Function

Function CheckInt(sTextInput)
Dim sTextTemp
	If sTextInput>"" And IsNumeric(sTextInput) Then
		sTextTemp=CLng(sTextInput)
	Else
		sTextTemp=Null
	End If
	CheckInt=sTextTemp
End Function

Function CheckIntegerAndZero(sTextInput)
Dim sTextTemp
	If sTextInput>"" And IsNumeric(sTextInput) Then
		sTextTemp=CLng(sTextInput)
	Else
		sTextTemp=0
	End If
	CheckIntegerAndZero=sTextTemp
End Function

Function CheckIntegerAndNull(sTextInput)
Dim sTextTemp
	If sTextInput>"" And IsNumeric(sTextInput) Then
		sTextTemp=CLng(sTextInput)
	Else
		sTextTemp=Null
	End If
	CheckIntegerAndNull=sTextTemp
End Function

Function CheckSingleAndNull(sTextInput)
Dim sTextTemp
	If sTextInput>"" And IsNumeric(sTextInput) Then
		sTextTemp=CSng(sTextInput)
	Else
		sTextTemp=Null
	End If
	CheckSingleAndNull=sTextTemp
End Function

%>