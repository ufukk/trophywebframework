<%
Public Function toUpper(s)
	s = replace(s,"ý","I")
	s = replace(s,"i","Ý")
	toUpper = UCase(s)
End Function

Public Function wordSplit(str,count)
	if InStr(str," ") = 0 then
		Dim tmp(0)
		tmp(0) = str
	else
		tmp = split(Trim(str)," ")
	end if
	if count = "" then
		count = 10
	end if
	if CInt(count) > UBound(tmp) then
		estr = ""
		count = UBound(tmp)
	else
		estr = "..."
	end if
	For i = 0 to count
		output = output & tmp(i) & " "
	Next
	wordSplit = output & estr
End Function

Public Function getSearchSQL(words,fields,keyword)
	Dim output 
	output = " 1=0 "
	Set rp = newRegexp()
	rp.Pattern = """" & "[^" & """" & "]*" & """"
	For a=0 to UBound(words)
		if a = 0 then
			kw = " OR "
		else
			kw = keyword
		end if
		word = replace(words(a),"+"," ")
		output = output & " " & kw & " (1=0 "
		if Len(word) > 0 then
			For i=0 to UBound(fields)
				if Len(fields(i)) > 0 then
					output = output & " OR " & fields(i) & " LIKE '%" & word & "%'"
				end if
			Next
		end if
		output = output & " ) "
	Next
	getSearchSQL = output
End Function 

Public Function getFileExtension(ByVal s)
       tmpArray = split(s,".")
	   getFileExtension = tmpArray(Ubound(tmpArray))
End Function

Public Function getFilePrefix(ByVal s)
       getFilePrefix = Mid(s,1,InStrRev(s,"."))
End Function

Public Function getFileName(ByVal s)
	   Dim separator
	   separator = "\"
	   if InStr(s,separator) = 0 And InStr(s,"/") > 0 then
		  separator = "/"
	   end if
       tmpArray = split(s,"\")
	   getFileName = tmpArray(Ubound(tmpArray))
End Function

Public Function getFileDirectory(ByVal s)
       getFileDirectory = Mid(s,1,InStrRev(s,"\"))
End Function

Public Function strip_slashes(s) 
	s = replace(s,"""""""","""")
	s = replace(s,"''","'")
	strip_slashes = s
End Function

Public Function nl2br(s)
	s = replace(s,vbNewLine,vbNewLine & "<br />")
	nl2br = s
End Function

Public Function Trim(s)
	Trim = Rtrim(LTrim(s))
End Function

Public Function toWords(s) 
	toWords = split(Trim(s)," ")
End Function

Public Function str_unpush(s)
	str_unpush = Mid(s,1,Len(s)-1)
End Function

Public Function str_unshift(s)
	str_unshift = Mid(s,2,Len(s))
End Function

Public Function addSlashes(s)
	addSlashes = replace(s,"'","\'")
End Function

Public Function noSlashes(s)
	noSlashes = replace(s,"""","")
End Function

Public Function htmlEncode(s)
	if Not IsNull(s) then
		htmlEncode = Server.HtmlEncode(s)
	end if
End Function

Public Function urlEncode(s)
	if Not IsNull(s) then
		urlEncode = Server.urlEncode(s)
	end if
End Function

Public Function stripHTML(s)
	Set rxp = new Regexp
	rxp.Global = True
	rxp.MultiLine = True
	rxp.IgnoreCase = True
	rxp.Pattern = "\<[^>]*\>"
	stripHTML = rxp.replace(s,"")
End Function

Public Function stripTags(s,allowedtags)
	Dim results,result,rxp
	Set rxp = newRegexp()
	rxp.Pattern = "\<([^>]*)\>"
	Set results = rxp.Execute(s)
	For Each resultTags in results
		Set result = resultTags.subMatches(0)
		if InStr(result," ") > 0 then
			tagName = Trim(Mid(result,1,InStr(result," ")))
		else
			tagName = trim(result)
		end if
		if Not inArray(allowedtags,tagName) then
			s = replace(s,result,"")
		end if
	Next
	stripTags = s
End Function

Public Function addQuotes(s)
	addQuotes = """" & s & """"
End Function

Public Function addSingleQuotes(s)
	addSingleQuotes = "'" & s & "'"
End Function

Public Function stripQuotes(s)
	if Not IsNull(s) then
		if Len(s) > 1 then
			startIndex = 1
			endIndex = Len(s)
			if InStr(s,"""") = 1 then
				startIndex = 2
			end if
			if Mid(s,Len(s),1) = """" then
				endIndex = Len(s)-2
			end if
			stripQuotes = Mid(s,startIndex,endIndex)
		else
			stripQuotes = s
		end if
	end if
End Function

Public Function fixSlashes(s)
		s = replace(s,"""","""" & " & " & """""""""" & " & " & """") 
		s = replace(s,vbCRLF,"""" & " & vbCRLF & " & """")
		fixSlashes = s
End Function

Public Function fixSlashesAndQuote(s)
	fixSlashesAndQuote = addQuotes(fixSlashes(s))
End Function

Public Function getParameterList(str)
	Set r = newRegexp()
	str = str & ","
	r.Pattern =  ",(?=(?:[^""]*""[^""]*"")*(?![^""]*""))"
	Set results = r.Execute(str)
	startIndex = 1
	t = 0
	Dim resultDict
	Set resultDict = newDictionary()
	For each result in results
		if t <> 0 then
			startIndex = results(t-1).FirstIndex
		end if
		if t > 1 then
			startIndex = results(t-1).FirstIndex+2
		end if
		if t=1 then
			startIndex = startIndex+2
		end if
		substrLength = result.FirstIndex-startIndex+1
		Call resultDict.Add(t,Mid(str,startIndex,substrLength))
		t = t+1
	Next
	Set getParameterList = resultDict
End Function

Public Function urlDecode(s)
	urlDecode = s
End Function
%>