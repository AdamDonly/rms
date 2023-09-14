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
		colCellBG="#FFFFFF"
		colTableBG="#CC0000"
		colGridBG="#CC0000"
		classTable="cv"
	Else
		colCellBG="#FFFFFF"
		colTableBG="#FFFFFF"
		colGridBG="#FFFFFF"
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
	"<table width=""98%"" border=0 bgcolor=""" & colTableBG & """ cellpadding=1 cellspacing=0 align=center><tr><td>" & vbCrLf & _
	"  <table cellspacing=0 cellpadding=0 align=""center"" width=""100%"">" & vbCrLf
End Sub

Sub WriteTableFooter
	Response.Write "" & _
	" </table>" & vbCrLf & _
	"</td></tr></table><br>" & vbCrLf
End Sub

Sub WriteTableFooterWithoutSpace
	Response.Write "" & _
	" </table>" & vbCrLf & _
	"</td></tr></table>" & vbCrLf
End Sub

Sub WriteGridTableFooter
	Response.Write "" & _
	" </table>" & vbCrLf & _
	"</td></tr></table><br>" & vbCrLf
End Sub

Sub WriteDataTitle(sTitle)
	Response.Write "" & _
	"<table cellspacing=0 cellpadding=0 align=""center"" width=""99%"">" & vbCrLf & _
	"  <tr><td width=""100%"" colspan=3 bgcolor=""#FFFFFF"" align=""left""><p class=""txt""><b>" & sTitle & "</b></p></td></tr>" & vbCrLf & _
	"</table><img src=""image/x.gif"" width=1 height=5><br>"
End Sub	

Sub WriteDataRow(sField, sValue)
	Response.Write "" & _
	"  <tr><td bgcolor=""" & colCellBG & """ width=""25%"" valign=""top""><p class=""txt""><font><b>" & sField & "</b></font></td><td width=1><img src=""image/x.gif"" width=1 height=1><br></td><td bgcolor=""#FFFFFF"" align=""left""><p class=""txt"">" & sValue & "</td></tr>" & vbCrLf
End Sub

Sub WriteDataRow1Column(sField, sValue)
	Response.Write "" & _
	"  <tr><td bgcolor=""" & colCellBG & """ width=""100%"" valign=""top""><p class=""txt"">" & sField & "</b></td></tr>" & vbCrLf
End Sub

Sub WriteDataRowHeader(sTitle)
	Response.Write "" & _
	"  <tr><td bgcolor=""" & colCellBG & """ width=""100%"" valign=""top""><p class=""txt""><b>" & sTitle & "</b></td></tr>" & vbCrLf
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
	"    <td width=1><img src=""image/x.gif"" width=1 height=3><br></td>" & vbCrLf & _
	"    <td width=""80%"" bgcolor=""#FFFFFF"" align=""left""></td>" & vbCrLf & _
	"  </tr>"
End Sub


Sub WriteGridTableHeader
	Response.Write "" & _
	"<table width=""96%"" border=0 bgcolor=""" & colGridBG & """ cellpadding=0 cellspacing=0 align=center><tr><td>" & vbCrLf & _
	"  <table cellspacing=1 cellpadding=2 align=""center"" width=""100%"">" & vbCrLf
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
			Response.Write "    <td width=""" & arrColumnsWidths(i) & """ bgcolor=""" & colCellBG & """ valign=""top""><p class=""txt"" " & sAlingment & ">" & arrColumnsValues(i) & "</td>" & vbCrLf
		Else
			Response.Write "    <td bgcolor=""" & colCellBG & """ valign=""top""><p class=""txt"" " & sAlingment & ">" & arrColumnsValues(i) & "</td>" & vbCrLf
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

Function GetListStart()
	GetListStart="<ul style=""margin: 0 25px; padding:0"">"
End Function

Function GetListEnd()
	GetListEnd="</ul>"
End Function

Function GetListItemStart()
	GetListItemStart="<li><p class=""txt"">"
End Function

Function GetListItemEnd()
	GetListItemEnd="</p></li>"
End Function

Function GetListItemEndLast()
	GetListItemEndLast="</li>"
End Function

Sub WriteDataRow2(sTitle1, sValue1, sTitle2, sValue2)
	Response.Write "" & _
	"  <tr><td bgcolor=""" & colCellBG & """ width=""25%"" valign=""top""><p class=""txt""><font><b>" & sTitle1 & "</b></font></td><td width=1><img src=""image/x.gif"" width=1 height=1><br></td><td bgcolor=""#FFFFFF"" align=""left"">" & _
	"<table width=""100%"" cellpadding=0 cellspcing=0><tr valign=""top""><td width=""33%""><p class=""txt"">" & sValue1 & "</td><td width=""10%""><p class=""txt""><font><b>" & sTitle2 & "</b></font></td><td><p class=""txt"">" & sValue2 & "</td></tr></table></td></tr>"
End Sub

Sub WriteSimpleRow(sValue1)
	Response.Write "<tr><td colspan=3 height=1 ><img src=""image/x.gif"" width=1 height=1><br></td></tr>" & _
	"<tr><td colspan=3 bgcolor=""#FFFFFF"" align=""left""><p class=""txt"">" & sValue1 & "</p></td></tr>" & _
	"<tr><td colspan=3 height=1 ><img src=""image/x.gif"" width=1 height=1><br></td></tr>"
End Sub

Sub WriteSimpleRowNoBorder(sValue1)
	Response.Write "<tr><td colspan=3 bgcolor=""#FFFFFF"" align=""left""><p class=""txt"">" & sValue1 & "</p></td></tr>"
End Sub

Sub WriteRowBorder()
	Response.Write "<tr><td colspan=3 height=1 ><img src=""image/x.gif"" width=1 height=1><br></td></tr>"
End Sub
%>
