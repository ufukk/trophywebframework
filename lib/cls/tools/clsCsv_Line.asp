<%



class CSV_Line

	Private content
	Private columnArray
	Private separator
	Public autoTrim
	
	Public Sub Class_Initialize()
		separator = ";"
		autoTrim = True
	End Sub
	
	Public Property Let LineString(l)
		content = l
		columnArray = split(content,separator)
	End Property
	
	Public Default Property Get Columns(i)
		Dim v
		if ColumnCount >= i then
			v = columnArray(i)
		else
			v = Null
		end if
		if autoTrim then
			v = Trim(v)
		end if
		Columns = v
	End Property
	
	Public Property Get ColumnCount
		ColumnCount = UBound(columnArray)
	End Property
	
	Public Property Get IsEmpty
		if Len(Trim(content)) < 1 then
			r = True
		else
			r = False
		end if
		IsEmpty = r
	End Property
	
end class





%>