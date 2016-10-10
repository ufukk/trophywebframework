<%
Public Function dayBox(name,selected,blankValue,extra)
       dim x,str,output,strval
	   Set objField = new Trophy_FormField
	   objField.fType = "selectbox"
	   objField.Name = name
	   output = "<SELECT name=""" & name & """" & extra & ">"
	   if Not IsNull(blankValue) then
	   	if 0 = CInt(selected) then
		      str = " selected "
		   else
		      str = ""
		   end if
		output = output & "<OPTION value='" & "" & "'" & str & ">" & blankValue & "</OPTION>" & vbCRLF
	   end if
	   For x = 1 to 31
	       if x = CInt(selected) then
		      str = " selected "
		   else
		      str = ""
		   end if
		   if x < 10 then
		      strval = "0" & CStr(x)
		   else
		      strval = CStr(x)
		   end if
		   output = output & "<OPTION value='" & strval & "'" & str & ">" & x & "</OPTION>" & vbCRLF
	   Next
	   output = output & "</SELECT>"
	   dayBox = output
End Function

Public Function monthBox(name,selected,blankValue,extra)
       dim x,str,output,strval
	   output = "<SELECT name='" & name & "' " & extra & ">"
	   if Not IsNull(blankValue) then
	   	if 0 = CInt(selected) then
		      str = " selected "
		   else
		      str = ""
		   end if
		output = output & "<OPTION value='" & "" & "'" & str & ">" & blankValue & "</OPTION>" & vbCRLF
	   end if
	   For x = 1 to 12
	       if x = CInt(selected) then
		      str = " selected "
		   else
		      str = ""
		   end if
		   if x < 10 then
		      strval = "0" & CStr(x)
		   else
		      strval = CStr(x)
		   end if
		   output = output & "<OPTION value='" & strval & "'" & str & ">" & MonthName(x) & "</OPTION>"
	   Next
	   output = output & "</SELECT>"
	   monthBox = output
End Function

Public Function yearBox(name,selected,blankValue,startyear,endyear,extra)
       dim x,str,output,strval
	   output = "<SELECT name='" & name & "' " & extra & ">"
	   if Not IsNull(blankValue) then
	   	if 0 = CInt(selected) then
		      str = " selected "
		   else
		      str = ""
		   end if
		output = output & "<OPTION value='" & "" & "'" & str & ">" & blankValue & "</OPTION>" & vbCRLF
	   end if
	   startyear = CInt(startyear)
	   endyear = CInt(endyear)
	   For x = startyear to endyear
	       if Trim(selected) = "" then
		      selected = 0
		   else 
		      selected = CInt(selected)	  
		   end if
		   if x = selected then
		      str = " selected "
		   else
		      str = ""
		   end if
		   output = output & "<OPTION value='" & x & "'" & str & ">" & x & "</OPTION>" & vbCRLF
	   Next
	   output = output & "</SELECT>"
	   yearBox = output
End Function

Public Function hourBox(name,selected,extra)
       dim x,str,output,strval
	   output = "<SELECT name='" & name & "' " & extra & ">"
	   For x = 0 to 23
	       if x = CInt(selected) then
		      str = " selected "
		   else
		      str = ""
		   end if
		   if x < 10 then
		      strval = "0" & CStr(x)
		   else
		      strval = CStr(x)
		   end if
		   output = output & "<OPTION value='" & strval & "'" & str & ">" & strval & "</OPTION>"
	   Next
	   output = output & "</SELECT>"
	   hourBox = output
End Function

Public Function minuteBox(name,selected,extra)
       dim x,str,output,strval
	   output = "<SELECT name='" & name & "' " & extra & ">"
	   For x = 0 to 59
	       if x = CInt(selected) then
		      str = " selected "
		   else
		      str = ""
		   end if
		   if x < 10 then
		      strval = "0" & CStr(x)
		   else
		      strval = CStr(x)
		   end if
		   output = output & "<OPTION value='" & strval & "'" & str & ">" & strval & "</OPTION>"
	   Next
	   output = output & "</SELECT>"
	   minuteBox = output
End Function

Public Function weekDayBox(name,selected,extra)
	dim output,str
	output = "<SELECT name='" & name & "'" & extra & ">"
	For i = 0 to 6
		if CInt(selected) = i then
			str = " selected "
		else
			str = ""
		end if
		output = output & "<option" & str & " value=" & i & ">" & getDayName(i) & "</option>"
	Next
	output = output & "</select>"
	weekDayBox = output
End Function

Public Function date2sqldate(dateStr)
       dim output
	   output = Day(dateStr) & "." & Month(dateStr) & "." & Year(dateStr) & " " & Hour(dateStr) & ":" & Minute(dateStr) & ":" & Second(dateStr)
       date2sqldate = output
End Function

Public Function input2sqldate(d,m,y,h,g,s)
       dim output
	   if Trim(h) = "" then
	      h = "00" 
	   end if
	   if Trim(g) = "" then
	      g = "00" 
	   end if
	   if Trim(g) = "" then
	      s = "00"
	   end if
	   output = m & "." & d & "." & y & " " & h & ":" & g & ":" & s
	   input2sqldate = output
End Function

Public Function parseDate(dateStr,datetype)
       dim tmp1,tmp2,tmp3,timestr
	   tmp1 = split(dateStr,".")
	   tmp2 = split(tmp1(2)," ")
	   if Ubound(tmp2) > 0 then
	      timestr = tmp2(1)
		  tmp3 = split(timestr,":")
	      h = tmp3(0)
	      g = tmp3(1)
	      s = tmp3(2)
	   end if
	   m = tmp1(1)
	   d = tmp1(0)
	   y = tmp1(2)
	   parseDate = eval(datetype) 
End Function

Public Function setDtime(ByVal dayt,ByVal montht,ByVal yeart,ByVal hourt,ByVal minutet)
       setDtime = dayt+"-"+montht+"-"+yeart+" "+hourt+":"+minutet+":00"
End Function 

Public Function getTimeString(ByVal s) 
	   if Len(s) < 4 then
	      s = "0" & s
	   end if
	   r = Mid(s,1,2) & "." & Mid(s,3)
	   diff = 5-Len(r)
	   For z = 0 to diff-1
		r = r & "0"
	   Next
	   getTimeString = r
End Function

'4guysfromrolla
Function getDateString(strDate, strFormat)
	Dim intPosItem
	Dim intHourPart
	Dim strHourPart
	Dim strMinutePart
	Dim strSecondPart
	Dim strAMPM

	If not IsDate(strDate) Then
		DanDate = strDate
		Exit Function
	End If
	
	intPosItem = Instr(strFormat, "%m")
	Do While intPosItem > 0
                strFormat = Left(strFormat, intPosItem-1) & _
                        DatePart("m",strDate) & _
                        Right(strFormat, Len(strFormat) - (intPosItem + 1))
		intPosItem = Instr(strFormat, "%m")
	Loop

	intPosItem = Instr(strFormat, "%b")
	Do While intPosItem > 0
                strFormat = Left(strFormat, intPosItem-1) & _
                        MonthName(DatePart("m",strDate),True) & _
                        Right(strFormat, Len(strFormat) - (intPosItem + 1))
		intPosItem = Instr(strFormat, "%b")
	Loop
	
	intPosItem = Instr(strFormat, "%B")
	Do While intPosItem > 0
                strFormat = Left(strFormat, intPosItem-1) & _
                        MonthName(DatePart("m",strDate),False) & _
                        Right(strFormat, Len(strFormat) - (intPosItem + 1))
		intPosItem = Instr(strFormat, "%B")
	Loop
	
	intPosItem = Instr(strFormat, "%d")
	Do While intPosItem > 0
                strFormat = Left(strFormat, intPosItem-1) & _
                        DatePart("d",strDate) & _
                        Right(strFormat, Len(strFormat) - (intPosItem + 1))
		intPosItem = Instr(strFormat, "%d")
	Loop

	intPosItem = Instr(strFormat, "%j")
	Do While intPosItem > 0
                strFormat = Left(strFormat, intPosItem-1) & _
                        DatePart("y",strDate) & _
                        Right(strFormat, Len(strFormat) - (intPosItem + 1))
		intPosItem = Instr(strFormat, "%j")
	Loop

	intPosItem = Instr(strFormat, "%y")
	Do While intPosItem > 0
                strFormat = Left(strFormat, intPosItem-1) & _
                        Right(DatePart("yyyy",strDate),2) & _
                        Right(strFormat, Len(strFormat) - (intPosItem + 1))
		intPosItem = Instr(strFormat, "%y")
	Loop

	intPosItem = Instr(strFormat, "%Y")
	Do While intPosItem > 0
                strFormat = Left(strFormat, intPosItem-1) & _
                        DatePart("yyyy",strDate) & _
                        Right(strFormat, Len(strFormat) - (intPosItem + 1))
		intPosItem = Instr(strFormat, "%Y")
	Loop

	intPosItem = Instr(strFormat, "%w")
	Do While intPosItem > 0
                strFormat = Left(strFormat, intPosItem-1) & _
                        DatePart("w",strDate,1) & _
                        Right(strFormat, Len(strFormat) - (intPosItem + 1))
		intPosItem = Instr(strFormat, "%w")
	Loop

	intPosItem = Instr(strFormat, "%a")
	Do While intPosItem > 0
                strFormat = Left(strFormat, intPosItem-1) & _
                        WeekDayName(DatePart("w",strDate,1),True) & _
                        Right(strFormat, Len(strFormat) - (intPosItem + 1))
		intPosItem = Instr(strFormat, "%a")
	Loop

	intPosItem = Instr(strFormat, "%A")
	Do While intPosItem > 0
                strFormat = Left(strFormat, intPosItem-1) & _
                        WeekDayName(DatePart("w",strDate,1),False) & _
                        Right(strFormat, Len(strFormat) - (intPosItem + 1))
		intPosItem = Instr(strFormat, "%A")
	Loop

	intPosItem = Instr(strFormat, "%I")
	Do While intPosItem > 0
		intHourPart = DatePart("h",strDate) mod 12
		if intHourPart = 0 then intHourPart = 12
                strFormat = Left(strFormat, intPosItem-1) & _
                        intHourPart & _
                        Right(strFormat, Len(strFormat) - (intPosItem + 1))
		intPosItem = Instr(strFormat, "%I")
	Loop

	intPosItem = Instr(strFormat, "%H")
	Do While intPosItem > 0
		strHourPart = DatePart("h",strDate)
		if strHourPart < 10 Then strHourPart = "0" & strHourPart
                strFormat = Left(strFormat, intPosItem-1) & _
                        strHourPart & _
                        Right(strFormat, Len(strFormat) - (intPosItem + 1))
		intPosItem = Instr(strFormat, "%H")
	Loop

	intPosItem = Instr(strFormat, "%M")
	Do While intPosItem > 0
		strMinutePart = DatePart("n",strDate)
		if strMinutePart < 10 then strMinutePart = "0" & strMinutePart
                strFormat = Left(strFormat, intPosItem-1) & _
                        strMinutePart & _
                        Right(strFormat, Len(strFormat) - (intPosItem + 1))
		intPosItem = Instr(strFormat, "%M")
	Loop

	intPosItem = Instr(strFormat, "%S")
	Do While intPosItem > 0
		strSecondPart = DatePart("s",strDate)
		if strSecondPart < 10 then strSecondPart = "0" & strSecondPart
                strFormat = Left(strFormat, intPosItem-1) & _
                        strSecondPart & _
                        Right(strFormat, Len(strFormat) - (intPosItem + 1))
		intPosItem = Instr(strFormat, "%S")
	Loop

	intPosItem = Instr(strFormat, "%P")
	Do While intPosItem > 0
		if DatePart("h",strDate) >= 12 then
			strAMPM = "PM"
		Else
			strAMPM = "AM"
		End If
                strFormat = Left(strFormat, intPosItem-1) & _
                        strAMPM & _
                        Right(strFormat, Len(strFormat) - (intPosItem + 1))
		intPosItem = Instr(strFormat, "%P")
	Loop

	intPosItem = Instr(strFormat, "%%")
	Do While intPosItem > 0
                strFormat = Left(strFormat, intPosItem-1) & "%" & _
                        Right(strFormat, Len(strFormat) - (intPosItem + 1))
		intPosItem = Instr(strFormat, "%%")
	Loop

	getDateString = strFormat
End Function

Public Function getTimeStamp(ddate)
	getTimeStamp = Mid(Year(ddate),3,2) & Month(ddate) & Day(ddate) & Hour(ddate) & Minute(ddate) & Second(ddate)
End Function
%>
