<%
'--------------------------------------------------------------------
'
' Date / time functions
'
'--------------------------------------------------------------------

Function ConvertDateTimeForFilename(sDate)
Dim dTempDate, sYear, sMonth, sDay, sHour, sMinute, sSecond
	If IsDate(sDate) Then
		sYear=Year(sDate)
		sMonth=Month(sDate)
		sDay=Day(sDate)
		sHour=Hour(sDate)
		sMinute=Minute(sDate)
		sSecond=Second(sDate)

		dTempDate=sYear & Left("00", 2-Len(sMonth)) & sMonth & Left("00", 2-Len(sDay)) & sDay & "_" & Left("00", 2-Len(sHour)) & sHour & Left("00", 2-Len(sMinute)) & sMinute & Left("00", 2-Len(sSecond)) & sSecond
	Else
		dTempDate=Null
	End If
ConvertDateTimeForFilename=dTempDate
End Function

' Function for converting date to format YYYY/MM/DD from year, month and day values.
' It is used everywhere to transfer data to sql db
'
Function ConvertDMYForSql(sYear, sMonth, sDay)
Dim dTempDate
dTempDate=Null

If IsNumeric(sYear) And IsNumeric(sMonth) And IsNumeric(sDay) Then
	If Not (sDay>=1 And sDay<=31) Then sDay=1
	If Not (sMonth>=1 And sMonth<=12) Then sMonth=1
	If (sYear>=1800 And sYear<=2100) Then
		dTempDate=sYear & "/" & Left("00", 2-Len(sMonth)) & sMonth & "/" & Left("00", 2-Len(sDay)) & sDay
	End If
End If
ConvertDMYForSql=dTempDate
End Function


' Function for converting date to format YYYY/MM/DD from date value.
' It is used everywhere to transfer data to sql db
'
Function ConvertDateForSql(sDate)
Dim dTempDate, sYear, sMonth, sDay
	If IsDate(sDate) Then
		sYear=Year(sDate)
		sMonth=Month(sDate)
		sDay=Day(sDate)
		dTempDate=sYear & "/" & Left("00", 2-Len(sMonth)) & sMonth & "/" & Left("00", 2-Len(sDay)) & sDay
	Else
		dTempDate=Null
	End If
ConvertDateForSql=dTempDate
End Function


' Function for showing date in text format
' MM - first 3 characters of the month name
'
Function ConvertDateForText(sDate, sDelimeter, sFormat)
Dim dDateTemp, sDay, sMonth, sYear, sTime, sHour, sMinute
	If IsDate(sDate) Then
		If sFormat="MMYYYY" Then
			dDateTemp=Left(arrMonthName(Month(sDate)),3) & sDelimeter & Year(sDate)
		ElseIf sFormat="DayMonthYear" Then
			dDateTemp=Day(sDate) & sDelimeter & arrMonthName(Month(sDate)) & sDelimeter & Year(sDate)
		ElseIf sFormat="MonthYear" Then
			dDateTemp=arrMonthName(Month(sDate)) & sDelimeter & Year(sDate)
		ElseIf sFormat="DMY" Then
			dDateTemp=Day(sDate) & sDelimeter & Month(sDate) & sDelimeter & Year(sDate)
		ElseIf sFormat="DDMMYYYY HHMM" Then
			sDay=Day(sDate)
			sDay=Left("00", 2-Len(sDay)) & sDay
			sMonth=Month(sDate)
			sMonth=Left("00", 2-Len(sMonth)) & sMonth

			sHour=Hour(sDate)
			sTime=""
			If sHour<>"0" And sHour<>"" Then
				sHour=Left("00", 2-Len(sHour)) & sHour
				sMinute=Minute(sDate)
				sMinute=Left("00", 2-Len(sMinute)) & sMinute
				sTime="&nbsp;" & sHour & ":" & sMinute
			End If

			dDateTemp=sDay & sDelimeter & sMonth & sDelimeter & Year(sDate) & sTime
		ElseIf sFormat="DDMonYYYY HHMM" Then
			sDay=Day(sDate)
			sDay=Left("00", 2-Len(sDay)) & sDay
			sMonth=arrMonthName(Month(sDate))
			sMonth=Left(sMonth,3)

			sHour=Hour(sDate)
			sTime=""
			If sHour<>"0" And sHour<>"" Then
				sHour=Left("00", 2-Len(sHour)) & sHour
				sMinute=Minute(sDate)
				sMinute=Left("00", 2-Len(sMinute)) & sMinute
				sTime="&nbsp;" & sHour & ":" & sMinute
			End If

			dDateTemp=sDay & sDelimeter & sMonth & sDelimeter & Year(sDate) & sTime
		ElseIf sFormat="DayMonYYYY HHMM" Then
			sDay=Day(sDate)
			sMonth=arrMonthName(Month(sDate))
			sMonth=Left(sMonth,3)

			sHour=Hour(sDate)
			sTime=""
			If sHour<>"0" And sHour<>"" Then
				sHour=Left("00", 2-Len(sHour)) & sHour
				sMinute=Minute(sDate)
				sMinute=Left("00", 2-Len(sMinute)) & sMinute
				sTime="&nbsp;" & sHour & ":" & sMinute
			End If

			dDateTemp=sDay & sDelimeter & sMonth & sDelimeter & Year(sDate) & sTime
	        Else ' DDMMYYYY
			sDay=Day(sDate)
			sDay=Left("00", 2-Len(sDay)) & sDay

			sMonth=Month(sDate)
			sMonth=Left("00", 2-Len(sMonth)) & sMonth
			
			dDateTemp=sDay & sDelimeter & sMonth & sDelimeter & Year(sDate)
		End If
	Else
		dDateTemp=""
	End If 
ConvertDateForText=dDateTemp
End Function

Function ConvertSQLDateToText(sSqlDate, sDelimeter, sFormat)
Dim sYearTemp, sMonthTemp, sDayTemp, sDateTemp
sDateTemp=""
	If Len(sSqlDate)=8 Then
		sYearTemp=Left(sSqlDate,4)
		sMonthTemp=Mid(sSqlDate,5,2)
		sDayTemp=Right(sSqlDate,2)

		sDateTemp=sDayTemp & "/" & sMonthTemp & "/" & sYearTemp
		If IsDate(sDateTemp) Then
			sDateTemp=ConvertDateForText(sDateTemp, sDelimeter, sFormat)
		Else
			sDateTemp=""
		End If
	End If

ConvertSQLDateToText=sDateTemp
End Function

Function ShowInputDayMonthYear(AElementName, ADateSelect, AYearStart, AYearEnd)
	Dim iDaySelect, iMonthSelect, iYearSelect
	If IsDate(ADateSelect) Then
		iDaySelect = Day(ADateSelect)
		iMonthSelect = Month(ADateSelect)
		iYearSelect = Year(ADateSelect)		
	End If
	
	Dim sSelected
	Dim iDayLoop, iMonthLoop, iYearLoop
	%>
	<select name="<% =AElementName%>_d" size="1">
	<option value="0">Day</option>
	<% For iDayLoop = 1 To 31
		If iDayLoop = iDaySelect Then
			sSelected=" selected"
		Else
			sSelected=""
		End If		
		Response.Write("<option value=""" & iDayLoop & """" & sSelected & ">" & iDayLoop & "</option>")
	Next %>
	</select>
	<select name="<% =AElementName%>_m" size=1>
	<option value="0">Month</option>
	<% For iMonthLoop = 1 To 12
		If iMonthLoop = iMonthSelect Then
			sSelected=" selected"
		Else
			sSelected=""
		End If		
		Response.Write("<option value=""" & iMonthLoop & """" & sSelected & ">" & Left(MonthName(iMonthLoop), 3) & "</option>")
	Next %>
	</select>
	<select name="<% =AElementName%>_y" size="1">
	<option value="0">Year</option>
	<% For iYearLoop = AYearStart To AYearEnd
		If iYearLoop = iYearSelect Then
			sSelected=" selected"
		Else
			sSelected=""
		End If		
		Response.Write("<option value=""" & iYearLoop & """" & sSelected & ">" & iYearLoop & "</option>")
	Next %>
	</select>
<%
End Function
%>

