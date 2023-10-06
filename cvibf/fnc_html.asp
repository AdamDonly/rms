<%
Dim colCellBG, colTableBG, colGridBG, sAlingment
Dim classTable

If InStr(sScriptFileName, "adb")>0 Then
	sCVFormat="ADB"
ElseIf InStr(sScriptFileName, "afb")>0 Then
	sCVFormat="AFB"
ElseIf InStr(sScriptFileName, "ec")>0 Then
	sCVFormat="EC"
ElseIf InStr(sScriptFileName, "ep")>0 Then
	sCVFormat="EP"
ElseIf InStr(sScriptFileName, "wb")>0 Then
	sCVFormat="WB"
Else
	sCVFormat=""
End If

If sCVFormat="" Then
	If iMemberID>0 And bCvValidForMemberOrExpert=1 Then
		colCellBG="#FFF8F4"
		colTableBG="#CC0000"
		colGridBG="#CC0000"
		classTable="cv"
	Else
		colCellBG=""
		colTableBG="#0066CC"
		colGridBG="#0066CC"
		classTable="cv"
	End If
Else
	colCellBG="#FFFFFF"
	colTableBG="#FFFFFF"
	colGridBG="#0066CC"
	classTable="cv"
End If


Sub WriteTableHeader
	Response.Write "" & _
	"<table width=""98%"" class=""" & classTable & """ border=""0"" cellpadding=""1"" cellspacing=""0"" align=""center"">" & vbCrLf
	WriteSpaceRow
End Sub

Sub WriteTableFooter
	Response.Write "" & _
	"</table><br/>" & vbCrLf
End Sub

Sub WriteGridTableFooter
	Response.Write "" & _
	"</table></div><br/>" & vbCrLf
End Sub

Sub WriteDataTitle(sTitle)
%>
		<div class="box results blue">
		<h3><span class="left">&nbsp;</span><span class="right">&nbsp;</span><% =sTitle %></h3>
<%
End Sub	

Sub WriteDataRow(sField, sValue)
	Response.Write "" & _
	"  <tr><td bgcolor=""" & colCellBG & """ width=""25%"" valign=""top""><p class=""txt"">" & sField & "</b></td><td width=1><img src=""image/x.gif"" width=1 height=1><br/></td><td bgcolor=""#FFFFFF"" align=""left""><p class=""txt"">" & sValue & "</td></tr>" & vbCrLf
End Sub

Sub WriteDataRow1Column(sField, sValue)
	Response.Write "" & _
	"  <tr><td bgcolor=""" & colCellBG & """ width=""100%"" valign=""top""><p class=""txt"">" & sField & "</b></td></tr>" & vbCrLf
End Sub

Sub WriteDataRowWithFormat(sField, sValue, sCvFormat)
	Dim sAlignField, sAlignValue
	sAlignField="left"
	sAlignValue="left"
	If sCvFormat="EP" Then sAlignField="right"
	
	Response.Write "" & _
	"  <tr><td bgcolor=""" & colCellBG & """ width=""25%"" valign=""top""><p class=""txt"" align=""" & sAlignField & """>" & sField & "</b></td><td width=1><img src=""image/x.gif"" width=1 height=1><br></td><td bgcolor=""#FFFFFF"" valign=""top"" align=""left""><p class=""txt""  align=""" & sAlignValue & """>" & sValue & "</td></tr>" & vbCrLf
End Sub

Sub WriteSpaceRow
	Response.Write "" & _
	"  <tr>" & vbCrLf & _
	"    <td width=""250"" bgcolor=""" & colCellBG & """></td>" & vbCrLf & _
	"    <td width=1><img src=""image/x.gif"" width=1 height=3><br/></td>" & vbCrLf & _
	"    <td width=""80%"" bgcolor=""#FFFFFF"" align=""left""></td>" & vbCrLf & _
	"  </tr>"
End Sub


Sub WriteGridTableHeader
	Response.Write "" & _
	"<div class=""grid blue""><table width=""98%"" class=""" & classTable & """ border=""0"" cellpadding=""0"" cellspacing=""0"" align=""center"">" & vbCrLf
End Sub

Sub WriteGridDataRow(iColumnsNum, sColumnsWidths, sColumnsValues)
Dim arrColumnsWidths
Dim arrColumnsValues

	arrColumnsWidths=Split(sColumnsWidths, "#|#")
	arrColumnsValues=Split(sColumnsValues, "#|#")

	Response.Write "  <tr>" & vbCrLf
	If UBound(arrColumnsValues)=UBound(arrColumnsWidths) Or UBound(arrColumnsWidths)<0 Then
	For i=0 To UBound(arrColumnsValues)
		sAlingment=""
		If Left(arrColumnsValues(i),3)="\qc" Then
			sAlingment="align=""center"""
			arrColumnsValues(i)=Right(arrColumnsValues(i),Len(arrColumnsValues(i))-3)
		End If
		If UBound(arrColumnsWidths)>=0 Then
			Response.Write "    <td width=""" & arrColumnsWidths(i) & """ bgcolor=""" & colCellBG & """ valign=""top""><p class=""txt"" " & sAlingment & ">" & arrColumnsValues(i) & "&nbsp;</td>" & vbCrLf
		Else
			Response.Write "    <td bgcolor=""" & colCellBG & """ valign=""top""><p class=""txt"" " & sAlingment & ">" & arrColumnsValues(i) & "&nbsp;</td>" & vbCrLf
		End If
	Next
	Else
		Response.Write "<td>Error</td>"
	End If
	Response.Write "  </tr>"
End Sub

'-----------------------------------------------------------------------------
' WriteGridDataMultiRow
' In:		iColumnsNum	int	Number of colums
'		iRowsNum	int	
' Out:		-
' Descr:	Procedure writes the first row of the rows batch with common first cell
'-----------------------------------------------------------------------------
Sub WriteGridDataMultiRow(iColumnsNum, iRowsNum, sColumnsWidths, sColumnsValues)
Dim arrColumnsWidths
Dim arrColumnsValues
Dim sAlingment, sRowCols

	arrColumnsWidths=Split(sColumnsWidths, "#|#")
	arrColumnsValues=Split(sColumnsValues, "#|#")

	Response.Write "  <tr>" & vbCrLf
	If UBound(arrColumnsValues)=UBound(arrColumnsWidths) Or UBound(arrColumnsWidths)<0 Then
	For i=0 To UBound(arrColumnsValues)
		sAlingment=""
		If Left(arrColumnsValues(i),3)="\qc" Then
			sAlingment="align=""center"""
			arrColumnsValues(i)=Right(arrColumnsValues(i),Len(arrColumnsValues(i))-3)
		End If
		sRowCols=""
		If i=0 And iRowsNum>1 Then
			sRowCols="rowspan=" & iRowsNum
		End If

		If UBound(arrColumnsWidths)>=0 Then
			Response.Write "    <td width=""" & arrColumnsWidths(i) & """ bgcolor=""" & colCellBG & """ valign=""top"" " & sRowCols & "><p class=""txt"" " & sAlingment & ">" & arrColumnsValues(i) & "</td>" & vbCrLf
		Else
			Response.Write "    <td bgcolor=""" & colCellBG & """ valign=""top"" " & sRowCols & "><p class=""txt"" " & sAlingment & ">" & arrColumnsValues(i) & "</td>" & vbCrLf
		End If
	Next
	Else
		Response.Write "<td>Error</td>"
	End If
	Response.Write "  </tr>"
End Sub

%>
