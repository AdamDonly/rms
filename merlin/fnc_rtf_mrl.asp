<%
Dim iPageWidth, iWidthFlag, iWidth
Dim objFso, fInTemplate

If InStr(sScriptFileName, "adb")>0 Then
	sCVFormat="ADB"
ElseIf InStr(sScriptFileName, "afb")>0 Then
	sCVFormat="AFB"
ElseIf InStr(sScriptFileName, "ec")>0 Then
	sCVFormat="EC"
ElseIf InStr(sScriptFileName, "wb")>0 Then
	sCVFormat="WB"
Else
	sCVFormat=""
End If


Sub WriteDataRow(sTitle, sValue)
	Response.Write "\trowd\irow1\trgaph0\trrh0\trql\trleft102 "
	Response.Write "\clvertalt\clbrdrt\brdrs\brdrcf2\brdrw12\clbrdrl\brdrs\brdrcf2\brdrw12\clbrdrr\brdrs\brdrcf2\brdrw12\clbrdrb\brdrs\brdrcf2\brdrw12\cltxlrtb\cellx1657 "
	Response.Write "\clvertalt\clbrdrt\brdrs\brdrcf2\brdrw12\clbrdrl\brdrs\brdrcf2\brdrw12\clbrdrr\brdrs\brdrcf2\brdrw12\clbrdrb\brdrs\brdrcf2\brdrw12\cltxlrtb\cellx9106 "
	Response.Write "\pard\plain\intbl\li102\sl-264{\f6\fs22\b1\up1\cf2 " & sTitle & "\b0\up0}\cell "
	Response.Write "\pard\plain\intbl\li100\sl-264{\f6\fs22 " & sValue & "}\cell\row " & vbCrLf
End Sub

Sub WriteDataRowHeader(sTitle)
	Response.Write "\pard \ltrpar\qj \li0\ri0\sa120\widctlpar\brdrb\brdrs\brdrw30\brsp20\brdrcf9 \wrapdefault\aspalpha\aspnum\faauto\outlinelevel0\adjustright\rin0\lin0\itap0 "
	Response.Write "{\rtlch\fcs1 \af0 \ltrch\fcs0 \b\embo\fs22\cf9\insrsid3360922 " & sTitle & " \b0\par }"
End Sub

Sub WriteDataRow2(sTitle1, sValue1, sTitle2, sValue2)
	Response.Write "\trowd\irow4\trgaph0\trrh0\trql\trleft102 "
	Response.Write "\clvertalt\clbrdrt\brdrs\brdrcf2\brdrw12\clbrdrl\brdrs\brdrcf2\brdrw12\clbrdrr\brdrs\brdrcf2\brdrw12\clbrdrb\brdrs\brdrcf2\brdrw12\cltxlrtb\cellx1657 "
	Response.Write "\clvertalt\clbrdrt\brdrs\brdrcf2\brdrw12\clbrdrl\brdrs\brdrcf2\brdrw12\clbrdrr\brdrs\brdrcf2\brdrw12\clbrdrb\brdrs\brdrcf2\brdrw12\cltxlrtb\cellx4598 "
	Response.Write "\clvertalt\clbrdrt\brdrs\brdrcf2\brdrw12\clbrdrl\brdrs\brdrcf2\brdrw12\clbrdrr\brdrs\brdrcf2\brdrw12\clbrdrb\brdrs\brdrcf2\brdrw12\cltxlrtb\cellx6194 "
	Response.Write "\clvertalt\clbrdrt\brdrs\brdrcf2\brdrw12\clbrdrl\brdrs\brdrcf2\brdrw12\clbrdrr\brdrs\brdrcf2\brdrw12\clbrdrb\brdrs\brdrcf2\brdrw12\cltxlrtb\cellx9106 "
	Response.Write "\pard\plain\intbl\li102\sl-264{\f6\fs22\b1\up1\cf2 " & sTitle1 & "\b0\up0}\cell "
	Response.Write "\pard\plain\intbl\li100\sl-264{\f6\fs22 " & sValue1 & "}\cell "
	Response.Write "\pard\plain\intbl\li102\sl-264{\f6\fs22\b1\up1\cf2 " & sTitle2 & "\b0\up0}\cell "
	Response.Write "\pard\plain\intbl\li100\sl-264{\f6\fs22 " & sValue2 & "}\cell\row " & vbCrLf
End Sub


Function GetListStart()
	GetListStart="\s18\ql \fi-360\li480\ri0\sa0\sl276\slmult1\widctlpar\wrapdefault\aspalpha\aspnum\faauto\ls2\adjustright\rin0\lin480\itap0\pararsid9253268\contextualspace \rtlch\fcs1 \af31507\afs22\alang1025 \ltrch\fcs0 " & _
		"\f31506\fs22\lang2057\langfe1033\cgrid\langnp2057\langfenp1033 "
End Function

Function GetListEnd()
	GetListEnd=""
End Function

Function GetListItemStart()
	GetListItemStart="{\rtlch\fcs1 \af31507 \ltrch\fcs0 \insrsid16399331 "
End Function

Function GetListItemEnd()
	GetListItemEnd="}\par "
End Function

Function GetListItemEndLast()
	GetListItemEndLast="}\par "
End Function



Sub WriteDataRowWithFormat(sField, sValue, sCvFormat)
	Dim sAlignField, sAlignValue
	sAlignField="\ql "
	sAlignValue="\ql "
	Dim iWidthField, iWidthValue
	iWidthField=2375
	iWidthValue=8300	
	
	If sCvFormat="EP" Then 
		sAlignField="\qr "
		iWidthField=2895
		iWidthValue=10108
	End If
	
	Response.Write("\trowd\trgaph140\trleft0\trrh0\cellx" & iWidthField & "\clbrdrl\brdrs\brdrw10 \cellx" & iWidthValue & "\pard\intbl\ql\sb20 " & sAlignField & sField & "\cell\pard\intbl\ql\sb20 " & sAlignValue & sValue & "\cell\pard\intbl\row" & vbCrLf)
End Sub


Sub WriteSimpleRow(sTitle)
	Response.Write "\trowd\irow1\trgaph0\trrh0\trql\trleft102 "
	Response.Write "\clvertalt\clbrdrt\brdrs\brdrcf2\brdrw12\clbrdrl\brdrs\brdrcf2\brdrw12\clbrdrr\brdrs\brdrcf2\brdrw12\clbrdrb\brdrs\brdrcf2\brdrw12\cltxlrtb\cellx9106 "
	Response.Write "\pard\plain\intbl\li102\sl-341{\f6\fs22\b1\up1\cf2 " & sTitle & "}\cell\row " & vbCrLf
End Sub

Sub WriteSpaceRow
End Sub

Sub WriteTableHeader
	Response.Write("{")
End Sub

Sub WriteGridTableFooter
	Response.Write("}\par" & vbCrLf)
End Sub

Sub WriteTableFooter
	Response.Write("}\par" & vbCrLf)
End Sub

Sub WriteTableFooterNoPar
	Response.Write("}" & vbCrLf)
End Sub

Sub WriteTableFooterWithoutSpace
	Response.Write("}" & vbCrLf)
End Sub

Sub WriteDataTitle(sTitle)
	Response.Write("\par\ql\f1\fs18\cf2 \b " & sTitle & "\b0\par\par" & vbCrLf)
End Sub	

Function ConvertText2RTF(sText)
Dim iPosStart, iPosEnd
Dim sRTF
sRTF=sText
If Not IsNull(sRTF) Then
	sRTF=Trim(sRTF)
	sRTF=Replace(sRTF,"&nbsp;"," ")		
	sRTF=Replace(sRTF,"<b>","\b ")
	
	sRTF=Replace(sRTF,CHR(13)+CHR(10),"\line ")
	sRTF=Replace(sRTF,"\line \line ","\line ")
	
	sRTF=Replace(sRTF,"<br>","\line ")
	sRTF=Replace(sRTF,"<br />","\line ")
	sRTF=Replace(sRTF,"</b>","\b0 ")
	sRTF=Replace(sRTF,"<font color=""#C0C0C0"">","\f1\fs18\cf16 ")
	sRTF=Replace(sRTF,"</font>","\f1\fs18\cf2 ")
	sRTF=Replace(sRTF,"<p class=""welcome"">","")
	sRTF=Replace(sRTF,"<p class=""txt"">","")
	sRTF=Replace(sRTF,"<span class=""sml"">","")
	sRTF=Replace(sRTF,"<p style=""intext"">","")
	sRTF=Replace(sRTF,"</span>","")
	sRTF=Replace(sRTF,"</p>","\par ")

	sRTF=Replace(sRTF,"&#61623;","\line * ")
	sRTF=Replace(sRTF,"&#61608;","\line * ")
	sRTF=Replace(sRTF,"&#8217;","'")
	sRTF=Replace(sRTF,"&#8220;","""")
	sRTF=Replace(sRTF,"&#8221;","""")
	sRTF=ReplaceHtmlSpecialCodes(sRTF, "\line • ")

	sRTF=Replace(sRTF,"{","(")
	sRTF=Replace(sRTF,"}",")")
	
	If Left(sRTF,6)="\line " Then
		sRTF=Right(sRTF,Len(sRTF)-6)
	End If
	If Right(sRTF,6)="\line " Then
		sRTF=Left(sRTF,Len(sRTF)-6)
	End If
	If Right(sRTF,6)="\line " Then
		sRTF=Left(sRTF,Len(sRTF)-6)
	End If

	sRTF=Replace(sRTF,"</a>","")
	iPosStart=InStr(sRTF, "<a ")
	While iPosStart>0
		
		iPosEnd=InStr(iPosStart, sRTF, ">")
		
		If iPosEnd>0 Then sRTF=Left(sRTF, iPosStart-1) & Mid(sRTF, iPosEnd+1, Len(sRTF)-iPosEnd)

		iPosStart=InStr(iPosStart, sRTF, "<a ")
	WEnd
End If
ConvertText2RTF=sRTF
End Function

Sub WriteGridTableHeader
	Response.Write("{")
End Sub

Sub WriteGridDataRow(iColumnsNum, sColumnsWidths, sColumnsValues, iTableBorder, sCellsBorders)
Dim arrColumnsWidths
Dim arrColumnsValues
Dim arrCellsBorders

	arrColumnsWidths=Split(sColumnsWidths, "#|#")
	arrColumnsValues=Split(sColumnsValues, "#|#")
	arrCellsBorders=Split(sCellsBorders, "#|#")

	Response.Write("\trowd\trgaph90\trleft0\trrh262")
	If iTableBorder=1 Then
		Response.Write("\trbrdrt\brdrs\brdrw10 \trbrdrl\brdrs\brdrw10 \trbrdrb\brdrs\brdrw10 \trbrdrr\brdrs\brdrw10 ")
	End If
	iPageWidth=8300
	iWidthFlag=0
	If (UBound(arrColumnsValues)=UBound(arrColumnsWidths) And UBound(arrColumnsWidths)=UBound(arrCellsBorders)) Then
											'Or UBound(arrColumnsWidths)<0 
	For i=0 To UBound(arrColumnsValues)
		If UBound(arrColumnsWidths)>=0 Then
			iWidth=iWidthFlag+Round(CInt(Replace(arrColumnsWidths(i),"%",""))*iPageWidth/100)
			If i=UBound(arrColumnsValues) Then
				iWidth=8300
			End If
			iWidthFlag=iWidth
			If arrCellsBorders(i)=1 Then
				Response.Write("\clbrdrl\brdrs\brdrw10 \clbrdrr\brdrs\brdrw10 \clbrdrt\brdrs\brdrw10 \clbrdrb\brdrs\brdrw10 ")
			End If
			Response.Write("\cellx" & iWidth)
		End If
	Next
	For i=0 To UBound(arrColumnsValues)
		Response.Write("\pard\intbl\ql\sb20\fs18 " & ConvertText2RTF(arrColumnsValues(i)) & "\cell")
	Next
	Response.Write("\pard\intbl\row" & vbCrLf)
	End If
End Sub

Sub WriteGridDataRowStart(iColumnsNum, sColumnsWidths, iTableBorder, sCellsBorders)
Dim arrColumnsWidths
Dim arrCellsBorders

	arrColumnsWidths=Split(sColumnsWidths, "#|#")
	arrCellsBorders=Split(sCellsBorders, "#|#")

	Response.Write("\trowd\trgaph90\trleft0\trrh262")
	If iTableBorder=1 Then
		Response.Write("\trbrdrt\brdrs\brdrw10 \trbrdrl\brdrs\brdrw10 \trbrdrb\brdrs\brdrw10 \trbrdrr\brdrs\brdrw10 ")
	End If
	iPageWidth=8300
	iWidthFlag=0
	If (UBound(arrColumnsWidths)=UBound(arrCellsBorders)) Then
											'Or UBound(arrColumnsWidths)<0 
	For i=0 To UBound(arrColumnsWidths)
		If UBound(arrColumnsWidths)>=0 Then
			iWidth=iWidthFlag+Round(CInt(Replace(arrColumnsWidths(i),"%",""))*iPageWidth/100)
			If i=UBound(arrColumnsWidths) Then
				iWidth=8300
			End If
			iWidthFlag=iWidth
			If arrCellsBorders(i)=1 Then
				Response.Write("\clbrdrl\brdrs\brdrw10 \clbrdrr\brdrs\brdrw10 \clbrdrt\brdrs\brdrw10 \clbrdrb\brdrs\brdrw10 ")
			End If
			Response.Write("\cellx" & iWidth)
		End If
	Next
	End If
End Sub

Sub WriteGridDataRowCellStart
	Response.Write("\pard\intbl\ql\sb20 ")
End Sub

Sub WriteGridDataRowCellEnd
	Response.Write("\cell")
End Sub

Sub WriteGridDataRowCellValue(sCellValue)
	Response.Write(ConvertText2RTF(sCellValue))
End Sub

Sub WriteGridDataRowEnd
	Response.Write("\pard\intbl\row" & vbCrLf)
End Sub


Sub WriteGridDataRowLandscape(iColumnsNum, sColumnsWidths, sColumnsValues, iTableBorder, sCellsBorders)
Dim arrColumnsWidths
Dim arrColumnsValues
Dim arrCellsBorders

	arrColumnsWidths=Split(sColumnsWidths, "#|#")
	arrColumnsValues=Split(sColumnsValues, "#|#")
	arrCellsBorders=Split(sCellsBorders, "#|#")

	Response.Write("\trowd\trgaph90\trleft0\trrh262")
	If iTableBorder=1 Then
		Response.Write("\trbrdrt\brdrs\brdrw10 \trbrdrl\brdrs\brdrw10 \trbrdrb\brdrs\brdrw10 \trbrdrr\brdrs\brdrw10 ")
	End If
	iPageWidth=13400
	iWidthFlag=0
	If (UBound(arrColumnsValues)=UBound(arrColumnsWidths) And UBound(arrColumnsWidths)=UBound(arrCellsBorders)) Then
											'Or UBound(arrColumnsWidths)<0 
	For i=0 To UBound(arrColumnsValues)
		If UBound(arrColumnsWidths)>=0 Then
			iWidth=iWidthFlag+Round(CInt(Replace(arrColumnsWidths(i),"%",""))*iPageWidth/100)
			If i=UBound(arrColumnsValues) Then
				iWidth=iPageWidth
			End If
			iWidthFlag=iWidth
			If arrCellsBorders(i)=1 Then
				Response.Write("\clbrdrl\brdrs\brdrw10 \clbrdrr\brdrs\brdrw10 \clbrdrt\brdrs\brdrw10 \clbrdrb\brdrs\brdrw10 ")
			End If
			Response.Write("\cellx" & iWidth)
		End If
	Next
	For i=0 To UBound(arrColumnsValues)
		Response.Write("\pard\intbl\ql\sb20 " & ConvertText2RTF(arrColumnsValues(i)) & "\cell")
	Next
	Response.Write("\pard\intbl\row" & vbCrLf)
	End If

End Sub


Sub WriteGridDataMultiRow(iColumnsNum, sMergeParams, sColumnsWidths, sColumnsValues, iTableBorder, sCellsBorders)
Dim arrColumnsWidths
Dim arrColumnsValues
Dim arrCellsBorders

	arrColumnsWidths=Split(sColumnsWidths, "#|#")
	arrColumnsValues=Split(sColumnsValues, "#|#")
	arrCellsBorders=Split(sCellsBorders, "#|#")

	Response.Write("\trowd\trgaph90\trleft0\trrh262")
	If iTableBorder=1 Then
		Response.Write("\trbrdrt\brdrs\brdrw10 \trbrdrl\brdrs\brdrw10 \trbrdrb\brdrs\brdrw10 \trbrdrr\brdrs\brdrw10 ")
	End If
	iPageWidth=8300
	iWidthFlag=0
	If (UBound(arrColumnsValues)=UBound(arrColumnsWidths) And UBound(arrColumnsWidths)=UBound(arrCellsBorders)) Then
											'Or UBound(arrColumnsWidths)<0 
	For i=0 To UBound(arrColumnsValues)
		If UBound(arrColumnsWidths)>=0 Then
			iWidth=iWidthFlag+Round(CInt(Replace(arrColumnsWidths(i),"%",""))*iPageWidth/100)
			If i=UBound(arrColumnsValues) Then
				iWidth=8300
			End If
			iWidthFlag=iWidth
			If arrCellsBorders(i)=1 Then
				Response.Write("\clbrdrl\brdrs\brdrw10 \clbrdrr\brdrs\brdrw10 \clbrdrt\brdrs\brdrw10 \clbrdrb\brdrs\brdrw10 ")
			End If
			If i=0 And sMergeParams="1s" Then
				Response.Write("\clvmgf")
			ElseIf i=0 And sMergeParams="1c" Then
				Response.Write("\clvmrg")
			End If
			Response.Write("\cellx" & iWidth)
		End If
	Next
	For i=0 To UBound(arrColumnsValues)
		Response.Write("\pard\intbl\ql\sb20 " & ConvertText2RTF(arrColumnsValues(i)) & "\cell")
	Next
	Response.Write("\pard\intbl\row" & vbCrLf)
	End If

End Sub


Function ReplaceHtmlSpecialCodes(AText, AReplaceWith)
Dim sResult

	Dim iPositionCodeStart, iPositionCodeEnd, iCodeLength, sCode, iPositionCodePreviousEnd
	iPositionCodeStart=InStr(1, AText, "&#", 1)
	iPositionCodePreviousEnd=0
	
	If iPositionCodeStart=0 Then
		sResult=AText
	Else
		While iPositionCodeStart>0
			iPositionCodeEnd=iPositionCodeStart + FindHtmlCodeEnd(Mid(AText, iPositionCodeStart, 10) & " ")
			If iPositionCodeStart>0 And iPositionCodeEnd>iPositionCodeStart Then
				iCodeLength=iPositionCodeEnd-iPositionCodeStart
				sCode=Mid(AText, iPositionCodeStart, iCodeLength)
				
				sResult=sResult & Mid(AText, iPositionCodePreviousEnd+1, iPositionCodeStart-iPositionCodePreviousEnd-1) & AReplaceWith
				iPositionCodePreviousEnd=iPositionCodeEnd-1
			End If
			
			iPositionCodeStart=InStr(iPositionCodeEnd+1, AText, "&#", 1)
		WEnd
		If iPositionCodePreviousEnd>0 And Len(AText)>iPositionCodePreviousEnd+1 Then
			sResult=sResult & Mid(AText, iPositionCodePreviousEnd, Len(AText)-iPositionCodePreviousEnd+1)
		End if
	End If

ReplaceHtmlSpecialCodes=sResult
End Function

Function FindHtmlCodeEnd(AText)
Dim iResult
iResult=0

Dim iLoop, iCharCode
	For iLoop=1 To Len(AText)
		iCharCode=Asc(Mid(AText, iLoop, 1))
		
		If Not (iCharCode=35 Or iCharCode=38 Or iCharCode=59 Or _
			(iCharCode>=48 And iCharCode<=57)) Then
			iResult=iLoop
			Exit For
		End If
	Next
	
FindHtmlCodeEnd=iResult
End Function

%>