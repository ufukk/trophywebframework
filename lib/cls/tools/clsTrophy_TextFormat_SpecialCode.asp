<%



class Trophy_TextFormat_SpecialCode

	Private newLine 
	Private htmlNewLine
	Private tags
	
	Public Sub Class_Initialize()
		Set tags = newDictionary()
		newLine = vbNewLine
		htmlNewLine = "<br/>"
		Call addDefaultTags()
	End Sub
	
	Private Function addDefaultTags()
		Dim keys,values
		keys = array("B","I","P","U","LINK")
		values = array("B","I","P","U","A")
		For i=0 to UBound(keys)
			Call tags.Add(keys(i),values(i))
		Next
	End Function
	
	Public Function addTag(key,value)
		Call tags.Add(key,value)
	End Function
	
	Private Function replaceNewLine(ByVal txt)
		replaceNewLine = replace(txt,newLine,htmlNewLine & newLine)
	End Function
	
	Public Function parse(ByVal txt)
		Set objRxp = new Regexp
		objRxp.Global = True
		objRxp.IgnoreCase = True
		For Each key in tags
			objRxp.pattern = "\[" & key & "([^]]*)" & "\]"
			replacePattern = "<" & tags.Item(key) & "$1>"
			txt = objRxp.replace(txt,replacePattern)
			objRxp.pattern = "\[\/" & key & "([^]]*)" & "\]"
			replacePattern = "</" & tags.Item(key) & "$1>"
			txt = objRxp.replace(txt,replacePattern)
		Next
		txt = replaceNewLine(txt)
		parse = txt
	End Function
	

end class


%>